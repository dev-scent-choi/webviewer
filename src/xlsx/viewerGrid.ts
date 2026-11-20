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
  gridCellValue,
  parseXlsxToGrid,
  serializeGrid,
  yieldToEventLoop,
  YIELD_EVERY_N_ROWS,
  type GridCell,
  type GridDocument,
  type GridHeader,
} from '../spreadsheet.js';
import { numFmtCodeOf, formatValueText } from '../collab/numberFormat.js';
import { bakeCondFormats, bakeRow, CF_BAKE_WORK_CAP, localDayKey, styleRegistry } from './viewerCondFormat.js';
import { createCfFormulaFallback } from './cfFormulaEngine.js';
import { cfIsVolatile, cfRowIndex, computeStats, needsStats, type CfStats } from '../collab/condFormatShared.js';
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

/** 편집으로 바뀐 칸은 값 글자만 들고 있다 — 서식 코드가 있는 칸은 값에서 표시 글자(와 서식 코드 색 [Red])를 다시 만든다
 * (파일에서 읽은 칸은 결과가 같다). */
function displayCellOf(cell: GridCell | 0, registry: ReturnType<typeof styleRegistry>): GridCell | 0 {
  if (cell === 0 || cell.s === undefined) return cell;
  const style = registry.styleOf(cell.s);
  if (!numFmtCodeOf(style)) return cell;
  const value = gridCellValue(cell);
  const shown = formatValueText(value, style);
  const s = shown.color && shown.color !== style.fc ? registry.register({ ...style, fc: shown.color }) : cell.s;
  return shown.text === value ? { v: shown.text, s } : { v: shown.text, n: value, s };
}

/**
 * 편집기용 지연 그리드(현재 편집 상태 포함)에서 그리드 파일을 만든다 — 행을 묶음으로 읽어 붙이므로 메모리가 작다.
 * 조건부 서식(sheetMeta.condFormats — 지금 편집 상태의 규칙)도 굽는다: 규칙이 있는 시트는 행을 한 번 더 읽어 범위 통계를 내고,
 * 붙일 때 행마다 그 행의 규칙만 본다(viewerCondFormat.ts와 같은 판정·상한). 이 경로에서 수식 규칙이 다른 행을 가리키면
 * 지금 읽은 행 묶음(500행) 안의 값만 보인다 — 같은 행을 보는 흔한 규칙($C2>=1)은 정확하다.
 */
export async function serializeGridFromLazyStore(store: LazyGridStore): Promise<{ buffer: Buffer; sheetCount: number }> {
  const styles = [...store.styles];
  const registry = styleRegistry(styles);
  const work = { left: CF_BAKE_WORK_CAP };
  let volatile = false;
  const parts: Buffer[] = [];
  const newline = Buffer.from('\n');
  const BATCH = 500;
  // 행을 묶음으로 동시에 요청한다 — 저장소는 같은 순간의 요청을 한 번의 파일 열기로 읽는다(하나씩 기다리면 행마다
  // 파일을 열고 닫아 10만 행에 100초가 걸렸다)
  const readBatch = (s: number, start: number, end: number, colCount: number) =>
    Promise.all(Array.from({ length: end - start }, (_, i) => getRowCells(store, s, start + i, colCount)));
  for (let s = 0; s < store.sheetMeta.length; s += 1) {
    const { rowCount, colCount } = store.sheetMeta[s];
    const entries = (store.sheetMeta[s].condFormats ?? []).filter((e) => e.rule.t !== 'other');
    let stats = new Map<string, CfStats>();
    const byRow = entries.length > 0 ? cfRowIndex(entries, rowCount) : [];
    if (entries.length > 0) {
      volatile ||= cfIsVolatile(entries);
      // 1차: 통계가 필요한 규칙의 범위 값만 모은다(행 묶음으로 읽고 바로 버린다)
      const statEntries = entries.filter((e) => needsStats(e.rule));
      if (statEntries.length > 0) {
        const values = new Map<string, string[]>(statEntries.map((e) => [e.id, []]));
        for (let start = 0; start < rowCount; start += BATCH) {
          const end = Math.min(start + BATCH, rowCount);
          if (!byRow.slice(start, end).some((list) => list?.some((e) => needsStats(e.rule)))) continue;
          const rows = await readBatch(s, start, end, colCount);
          rows.forEach((line, i) => {
            const r = start + i;
            for (const e of byRow[r] ?? []) {
              if (!needsStats(e.rule)) continue;
              for (const rect of e.ranges) {
                if (r < rect.r0 || r > rect.r1) continue;
                for (let c = rect.c0; c <= Math.min(rect.c1, line.length - 1); c += 1) {
                  const cell = line[c];
                  const v = cell ? gridCellValue(cell) : '';
                  if (v !== '') values.get(e.id)?.push(v);
                }
              }
            }
          });
          await yieldToEventLoop();
        }
        stats = new Map(statEntries.map((e) => [e.id, computeStats(e.rule, values.get(e.id) ?? [])]));
      }
    }
    for (let start = 0; start < rowCount; start += BATCH) {
      const end = Math.min(start + BATCH, rowCount);
      const rows = await readBatch(s, start, end, colCount);
      const valueAt = (r: number, c: number): string => {
        const cell = r >= start && r < end ? rows[r - start][c] : undefined;
        return cell ? gridCellValue(cell) : '';
      };
      const fallback = entries.length > 0 ? createCfFormulaFallback(valueAt, () => ({ rows: rowCount, cols: colCount })) : undefined;
      rows.forEach((row, i) => {
        let cells = row.map((cell) => displayCellOf(cell, registry));
        if (entries.length > 0) cells = [...bakeRow(start + i, cells, byRow[start + i], stats, valueAt, fallback, registry, work)];
        while (cells.length > 0 && cells[cells.length - 1] === 0) cells.pop();
        parts.push(newline, Buffer.from(JSON.stringify(cells), 'utf8'));
      });
      if (end % YIELD_EVERY_N_ROWS < BATCH) await yieldToEventLoop();
    }
  }
  if (work.left <= 0) console.warn(`조건부 서식 굽기 상한(${CF_BAKE_WORK_CAP})에 닿아 큰 파일 보기 화면에 일부만 칠했다`);
  // 머리(서식 목록)는 행을 다 붙인 뒤에 만든다 — 굽기와 서식 코드 색이 서식을 더하므로
  const header: GridHeader = {
    formatVersion: GRID_CACHE_FORMAT_VERSION,
    ...(volatile ? { cfDay: localDayKey() } : {}),
    styles,
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
  return { buffer: Buffer.concat([Buffer.from(JSON.stringify(header), 'utf8'), ...parts]), sheetCount: store.sheetMeta.length };
}

/** 보기 화면용 그리드 — 조건부 서식까지 칸 서식에 구워 둔다(큰 파일은 serializeGridFromLazyStore가 같은 판정으로 굽는다) */
export async function parseViewerGrid(xlsxBytes: Uint8Array): Promise<GridDocument> {
  return bakeCondFormats(await parseXlsxToGrid(Buffer.from(xlsxBytes)), xlsxBytes);
}

/** xlsx 바이트 → 그리드 파일. 큰 파일은 스트리밍 읽기로 */
export async function buildViewerGridFile(xlsxBytes: Buffer): Promise<{ buffer: Buffer; sheetCount: number }> {
  if (!isLargeXlsx(xlsxBytes)) {
    const grid = await parseViewerGrid(xlsxBytes);
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
