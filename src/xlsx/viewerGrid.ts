/**
 * 보기 화면(viewer.html)용 그리드 파일(JSONL: 헤더 1줄 + 행마다 1줄, spreadsheet.ts serializeGrid와 같은 형식) 만들기.
 *
 * 보통 크기는 예전처럼 워크북을 통째로 읽는다(parseXlsxToGrid — 저장된 계산값이 없는 수식도 계산해 보여 준다).
 * 시트가 아주 큰 파일은 통째로 읽으면 메모리가 폭증한다(열 40개 10만 행에서 2.2GB, 30만 행이면 서버 한도 초과) —
 * 그런 파일은 편집기와 같은 스트리밍 읽기(ingestXlsxLazily, 행을 임시 파일로 흘려보냄)나 편집 방이 이미 가진 데이터로
 * 만든다. 이 경로에서는 계산값이 저장되지 않은 수식 칸이 수식 글자 그대로 보일 수 있다(엑셀이 저장한 파일은 계산값이 있다).
 */
import { randomUUID } from 'node:crypto';
import type { DocumentStore } from '../storage.js';
import {
  computeLineOffsets,
  GRID_CACHE_FORMAT_VERSION,
  parseXlsxToGrid,
  serializeGrid,
  yieldToEventLoop,
  YIELD_EVERY_N_ROWS,
  type GridHeader,
} from '../spreadsheet.js';
import { declaredEntrySizes } from '../zipGuard.js';
import { closeLazyGridStore, getRowCells, ingestXlsxLazily, type LazyGridStore } from './lazyGridStore.js';

/** 시트 XML(풀린 크기) 합이 이보다 크면 스트리밍 경로 — 이 크기에서 통째 읽기가 약 1GB를 쓴다 */
export const LARGE_SHEET_XML_BYTES = 48 * 1024 * 1024;

/** 풀린 시트 XML 크기 합(압축을 풀지 않고 zip 목차에서 읽는다) */
export function worksheetXmlBytes(xlsx: Uint8Array): number {
  let total = 0;
  for (const [name, size] of declaredEntrySizes(xlsx) ?? []) if (/^xl\/worksheets\/sheet\d+\.xml$/.test(name)) total += size;
  return total;
}

export function isLargeXlsx(xlsx: Uint8Array): boolean {
  return worksheetXmlBytes(xlsx) > LARGE_SHEET_XML_BYTES;
}

/** 편집기용 지연 그리드(현재 편집 상태 포함)에서 그리드 파일을 만든다 — 행을 하나씩 읽어 붙이므로 메모리가 작다 */
export async function serializeGridFromLazyStore(store: LazyGridStore): Promise<{ buffer: Buffer; sheetCount: number }> {
  const header: GridHeader = {
    formatVersion: GRID_CACHE_FORMAT_VERSION,
    styles: store.styles,
    sheets: store.sheetMeta.map((m) => ({
      name: m.name,
      rowCount: m.rowCount,
      colCount: m.colCount,
      colWidths: m.colWidths,
      defaultColWidth: m.defaultColWidth,
      defaultRowHeight: m.defaultRowHeight,
      rowHeights: m.rowHeights,
      showGridLines: m.showGridLines,
      merges: m.merges,
      freezeRows: m.freezeRows,
      freezeCols: m.freezeCols,
    })),
  };
  const parts: Buffer[] = [Buffer.from(JSON.stringify(header), 'utf8')];
  const newline = Buffer.from('\n');
  for (let s = 0; s < store.sheetMeta.length; s += 1) {
    const { rowCount, colCount } = store.sheetMeta[s];
    // 행을 묶음으로 동시에 요청한다 — 저장소는 같은 순간의 요청을 한 번의 파일 열기로 읽는다(하나씩 기다리면 행마다
    // 파일을 열고 닫아 10만 행에 100초가 걸렸다)
    const BATCH = 500;
    for (let start = 0; start < rowCount; start += BATCH) {
      const end = Math.min(start + BATCH, rowCount);
      const rows = await Promise.all(Array.from({ length: end - start }, (_, i) => getRowCells(store, s, start + i, colCount)));
      for (const row of rows) {
        const cells = [...row];
        while (cells.length > 0 && cells[cells.length - 1] === 0) cells.pop();
        parts.push(newline, Buffer.from(JSON.stringify(cells), 'utf8'));
      }
      if (end % YIELD_EVERY_N_ROWS < BATCH) await yieldToEventLoop();
    }
  }
  return { buffer: Buffer.concat(parts), sheetCount: store.sheetMeta.length };
}

/** xlsx 바이트 → 그리드 파일. 큰 파일은 스트리밍 읽기로 */
export async function buildViewerGridFile(xlsxBytes: Buffer): Promise<{ buffer: Buffer; sheetCount: number }> {
  if (!isLargeXlsx(xlsxBytes)) {
    const grid = await parseXlsxToGrid(xlsxBytes);
    return { buffer: serializeGrid(grid), sheetCount: grid.sheets.length };
  }
  const store = await ingestXlsxLazily(`viewer-${randomUUID()}`, xlsxBytes);
  try {
    return await serializeGridFromLazyStore(store);
  } finally {
    await closeLazyGridStore(store);
  }
}

/** 그리드 파일과 줄 위치 색인을 저장한다(saveGridDerived와 같은 두 파일) */
export async function saveGridFile(store: Pick<DocumentStore, 'saveDerived'>, docId: string, buffer: Buffer): Promise<void> {
  await store.saveDerived(docId, 'grid', buffer);
  await store.saveDerived(docId, 'gridindex', Buffer.from(JSON.stringify({ lineOffsets: computeLineOffsets(buffer) })));
}
