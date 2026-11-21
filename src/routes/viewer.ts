import fs from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { installedLineRatios } from '../docx/fontMetrics.js';
import { basePdfForEditor, sha1Of } from '../pdf/pdfModel.js';
import { addPdfPart, compressRoomPdf, currentPlainBytes, editorSourceOf, signRoomPdf, extractPdfGroups, PdfRoomError, pdfInfoOf, protectPdf, readPdfPart, tokenKey, unlockPdf, type ProtectRequest } from '../pdf/pdfRoomOps.js';
import { isCompressLevel } from '../pdf/pdfCompress.js';
import { convertPdf, isConvertTarget } from '../pdf/pdfConvert.js';
import { fileToPdf, FileToPdfError } from '../pdf/fileToPdf.js';
import { ConversionError } from '../converters/types.js';
import os from 'node:os';
import path from 'node:path';
import type { FastifyInstance, FastifyReply, FastifyRequest } from 'fastify';
import { recordAudit } from '../audit.js';
import { SizedLruCache } from '../cache.js';
import { config } from '../config.js';
import { findConverter } from '../converters/index.js';
import { IMAGE_MIME_BY_EXT, readSlideSizeEmu } from '../pptx/slideContent.js';
import { unzipSync } from 'fflate';
import { assertZipNotTooLarge, guardedUnzipSync, ZipTooLargeError } from '../zipGuard.js';
import { convertToPdf, editorUrlFor, viewPathFor } from './documents.js';
import {
  PageOutOfRangeError,
  buildPrintPdf,
  burnMarkupIntoPdf,
  getPageDimensions,
  invalidateRenderCache,
  renderPage,
  type PageSize,
  type PersistedPageSizes,
} from '../rendering.js';
import { InvalidTokenError, verifyViewToken } from '../security.js';
import { extractPdfText, searchPdfText, type PageText } from '../search.js';
import { detectNearbyTextStyle } from '../textStyle.js';
import {
  getDocxDefaultFontFamily,
  getDocxDefaultParagraphSpacing,
  getDocxDefaultFontSizePt,
  getDocxPageSetup,
  getDocxRenderExtras,
} from '../docx/import.js';
import { cachedDocxDrawing, DrawingRateLimited, renderDocxDrawingPng } from '../docx/drawingImage.js';
import { MarkupValidationError, addMarkup, deleteMarkup, listMarkup, updateMarkup } from '../markup.js';
import { touchPresence } from '../presence.js';
import { docxFragmentsOf, flushNow, getCfRemoteAnswer, getCurrentBytes, getChangeLog, getFileCharts, getFileExtras, getFileImages } from '../collab/room.js';
import { CfBusyError, cfTokenBudget, cfTokenSpend, parseCfRemoteRequest } from '../xlsx/cfRangeValues.js';

/** 편집 화면 조건부 서식 서버 답 — 편집할 때마다(0.3초 묶음) 다시 묻는다 */
// 토큰(편집 화면)마다 분당 40회 — verify:xlsx:browser(편집·스크롤·1,500·2만 행 COUNTIF 등 39항목)에서 한 화면이 1분에 부른 최대가 18회라
// 그 2배 남짓. 와일드카드 작업량은 따로 토큰마다 1분 예산(cfTokenBudget)으로 묶는다
const CF_REMOTE_RATE_LIMIT = { max: 40, timeWindow: '1 minute', keyGenerator: (request: { query?: unknown; ip: string }) => `cf:${(request.query as { token?: string } | undefined)?.token ?? request.ip}` } as const;
import {
  buildGridFromOriginal,
  computeLineOffsets,
  GRID_CACHE_FORMAT_VERSION,
  saveGridDerived,
  type GridCell,
  type GridSheet,
  type GridHeader,
  gridCellValue,
} from '../spreadsheet.js';
import { store } from '../storage.js';
import { LEGACY_EDIT_TARGET, legacyDownload } from '../legacyOffice.js';
import { isViewOnly, type DocumentMeta, type ViewClaims } from '../types.js';
import { isOpenable, waitForInitialConversion } from '../conversionGate.js';
import { buildViewerGridFile, saveGridFile } from '../xlsx/viewerGrid.js';
import { localDayKey } from '../xlsx/viewerCondFormat.js';
import { singleFlightWithBackoff } from '../xlsx/singleFlight.js';
import type { GridChart } from '../xlsx/chartModel.js';
import { chartDataOf, renderChartSvg } from '../xlsx/chartSvg.js';
import { rowTopPx } from '../xlsx/sheetGeometry.js';
import { cellBoxCssText, sanitizeCellStyle } from '../collab/cellBoxCss.js';
import { FILL_MARK } from '../collab/numberFormat.js';

export interface ViewContext {
  readonly claims: ViewClaims;
  readonly meta: DocumentMeta;
}

/**
 * 토큰 검증 + 문서 조회. 실패 시 reply를 직접 전송하고 null을 반환한다.
 * 토큰은 경로가 아닌 쿼리 파라미터로 받는다(JWT의 `.`이 경로 라우팅을 깨뜨리고
 * 액세스 로그에 토큰이 남는 문제를 피하기 위함).
 *
 * `src/routes/edit.ts`(동시편집 WebSocket 라우트)도 동일한 검증이 필요해 이 함수를 재사용한다.
 */
/** PDF 편집기용 PDF(다루는 주석을 뺀 것) — 문서 id·길이·내용 해시로 최근 몇 개만 기억한다 */
const editorBaseCache = new Map<string, Uint8Array>();
async function editorBaseOf(docId: string, bytes: Uint8Array): Promise<Uint8Array> {
  const key = `${docId}:${bytes.length}:${sha1Of(bytes)}`;
  const hit = editorBaseCache.get(key);
  if (hit) return hit;
  const base = await basePdfForEditor(bytes);
  editorBaseCache.set(key, base);
  while (editorBaseCache.size > 8) editorBaseCache.delete(editorBaseCache.keys().next().value as string);
  return base;
}

/** 편집기만 쓰는 데이터(원본 그림·hwp 원본·차트·링크·변경 내용) — 열람만 토큰(types.ts isViewOnly)은 거부 */
async function resolveEditorDataView(token: string | undefined, reply: FastifyReply): Promise<ViewContext | null> {
  const view = await resolveView(token, reply);
  if (view && isViewOnly(view.claims.perms)) {
    await reply.code(403).send({ success: false, error: '열람만 허용된 문서입니다.' });
    return null;
  }
  return view;
}

/** 문서에서 꺼낸 SVG 등을 그 주소로 직접 열었을 때 스크립트가 이 사이트 출처로 돌지 않게 — <img>·innerHTML 표시에는 영향 없음 */
const SANDBOX_CSP = "sandbox; default-src 'none'; img-src data:; style-src 'unsafe-inline'";

/** 큰 본문을 받는 라우트에 onRequest 로 — 본문(최대 maxUploadBytes)을 메모리에 다 받기 전에 토큰부터 본다.
 * 권한·문서 확인은 핸들러의 resolveView 가 그대로 한다. */
async function tokenBeforeBody(request: FastifyRequest, reply: FastifyReply): Promise<void> {
  const token = (request.query as { token?: unknown }).token;
  try {
    verifyViewToken(typeof token === 'string' ? token : '');
  } catch {
    await reply.code(401).send({ success: false, error: '열람 토큰이 없거나 유효하지 않습니다.' });
  }
}

export async function resolveView(token: string | undefined, reply: FastifyReply): Promise<ViewContext | null> {
  if (!token) {
    await reply.code(401).send({ success: false, error: '열람 토큰이 없습니다.' });
    return null;
  }
  let claims: ViewClaims;
  try {
    claims = verifyViewToken(token);
  } catch (error) {
    if (error instanceof InvalidTokenError) {
      await reply.code(401).send({ success: false, error: error.message });
      return null;
    }
    throw error;
  }
  const meta = await store.readMeta(claims.docId);
  if (!isOpenable(meta)) {
    await reply.code(404).send({ success: false, error: '문서를 찾을 수 없습니다.' });
    return null;
  }
  return { claims, meta };
}

/**
 * JSONL 그리드 인덱스: 헤더만 파싱하고 행 데이터는 줄 오프셋으로 색인해
 * 요청된 범위의 줄만 그때그때 파싱한다 (수십만 행 문서 대응).
 */
interface LoadedGrid {
  readonly header: GridHeader;
  readonly buffer: Buffer;
  readonly lineOffsets: readonly number[];
  readonly sheetStartLine: readonly number[];
}

/** 그리드·PDF텍스트 캐시는 렌더링 문서 캐시와 나머지 예산을 나눠 쓴다 */
const gridCache = new SizedLruCache<LoadedGrid>({
  maxBytes: Math.floor(config.cacheBudgetBytes * 0.25),
  maxEntries: 4,
  sizeOf: (grid) => grid.buffer.length + grid.lineOffsets.length * 8,
});

const pdfTextCache = new SizedLruCache<PageText[]>({
  maxBytes: Math.floor(config.cacheBudgetBytes * 0.1),
  maxEntries: 6,
  sizeOf: (pages) =>
    pages.reduce(
      (sum, page) => sum + page.items.reduce((acc, item) => acc + item.str.length * 2 + 48, 0),
      0,
    ),
});

/**
 * 동시편집 저장(src/collab/room.ts)이 문서를 다시 생성한 뒤 호출한다. docx/pptx는 새 PDF로
 * 재변환된 뒤 이전 텍스트 레이어가 검색 결과에 남지 않도록, xlsx는 그리드(+gridindex)가 다시
 * 저장된 뒤 이전 LoadedGrid가 캐시에 남아 낡은 내용을 서빙하지 않도록 각각 무효화한다.
 */
export function invalidateViewCaches(docId: string): void {
  pdfTextCache.invalidate(docId);
  gridCache.invalidate(docId);
}

function buildLoadedGridFromOffsets(buffer: Buffer, lineOffsets: readonly number[]): LoadedGrid {
  const headerEnd = lineOffsets.length > 1 ? lineOffsets[1] - 1 : buffer.length;
  const header = JSON.parse(buffer.toString('utf8', 0, headerEnd)) as GridHeader;
  const sheetStartLine: number[] = [];
  let line = 1;
  for (const sheet of header.sheets) {
    sheetStartLine.push(line);
    line += sheet.rowCount;
  }
  return { header, buffer, lineOffsets, sheetStartLine };
}

function indexGrid(buffer: Buffer): LoadedGrid {
  return buildLoadedGridFromOffsets(buffer, computeLineOffsets(buffer));
}

/**
 * 영구 저장된 gridindex가 구조적으로 말이 되는지 값싸게(오프셋 배열만 훑고, 버퍼 바이트는
 * 건드리지 않는다) 검사한다 — 실제 스캔 결과와 바이트 단위로 똑같은지까지는 확인하지 않는다
 * (그러면 인덱스를 신뢰하는 의미가 없어진다). 이 검사의 목적은 딱 하나: Task 6에서 실제로
 * 겪은 버그 클래스 — 마지막 오프셋이 버퍼 길이를 넘어서 Buffer.toString이 end를 조용히
 * clamp해 여러 줄이 뒤섞인 문자열을 반환하고 JSON.parse가 깨지는 것 — 을 막는다.
 */
function offsetsLookValid(buffer: Buffer, lineOffsets: readonly number[]): boolean {
  if (lineOffsets.length === 0 || lineOffsets[0] !== 0) return false;
  for (let i = 1; i < lineOffsets.length; i += 1) {
    if (lineOffsets[i] <= lineOffsets[i - 1]) return false;
  }
  return lineOffsets[lineOffsets.length - 1] <= buffer.length;
}

/** grid.json 헤더 한 줄만 값싸게 들여다본다(버퍼 전체를 다시 스캔하지 않는다) — formatVersion
 * 확인 전용. 파싱이 실패하면(손상 등) undefined를 돌려줘 "버전을 알 수 없음"으로 취급한다 —
 * 어차피 아래 loadGridUncached의 자가치유(gridindex 재스캔) 경로가 이런 손상을 이미 다룬다. */
function peekGridHeader(buffer: Buffer): Pick<GridHeader, 'formatVersion' | 'cfDay'> {
  try {
    const newline = buffer.indexOf(10);
    const headerEnd = newline === -1 ? buffer.length : newline;
    const header = JSON.parse(buffer.toString('utf8', 0, headerEnd)) as GridHeader;
    return { formatVersion: header.formatVersion, cfDay: header.cfDay };
  } catch {
    return {};
  }
}

/** 낡은 그리드 — 계산 방식 버전이 다르거나, TODAY()·NOW() 조건부 서식을 구운 날이 오늘이 아니다 */
function gridIsStale(header: Pick<GridHeader, 'formatVersion' | 'cfDay'>): boolean {
  return header.formatVersion !== GRID_CACHE_FORMAT_VERSION || (header.cfDay !== undefined && header.cfDay !== localDayKey());
}

/**
 * 업로드 시점(routes/documents.ts의 runGridExtraction)에 딱 한 번 만들어진 뒤로, 문서를 열어보는
 * 것만으로는 절대 재생성되지 않던 파생 그리드 캐시(derived kind 'grid')가 있다 — xlsx/formulas.ts의
 * 시트 간 참조(cross-sheet reference) 계산 버그가 고쳐지기 전에 만들어진 캐시는, 계산에 실패해
 * 수식 원문을 그대로 담고 있거나(옛날 동작) 그 버그로 인해 조용히 틀린 값을 담고 있을 수 있다 —
 * 어느 쪽이든 파일이 다시 저장(동시편집기를 통한 저장)되기 전까지는 디스크에 무기한 남는다.
 *
 * 이 함수는 grid.json 헤더의 formatVersion(GRID_CACHE_FORMAT_VERSION, spreadsheet.ts 문서 주석
 * 참고)이 낡았거나 아예 없으면(그 필드 자체가 이 수정으로 새로 생겼으므로, 이전에 저장된 캐시는
 * 전부 없음) 원본 파일에서 그리드를 다시 만들어 저장한다 — "다음 조회 시점에 감지해 lazy하게
 * 재생성"하는 전략으로, 저장된 모든 문서를 한꺼번에 마이그레이션하지 않는다(그 자체가 디스크에
 * 저장된 문서 수를 알 수 없는 상태에서 새로운 리스크가 된다). pagesizes/gridindex가 이미 쓰는
 * "없거나 낡으면 그 자리에서 다시 계산해 자가치유" 패턴과 같은 성질이다.
 *
 * 실패(원본이 없거나 — 예: 메모리 저장소에서 만료됨 — 재파싱이 실패함)하면 조용히 포기하고 false를
 * 돌려준다 — 그러면 호출부는 기존의(낡았을 수 있는) 캐시를 그대로 서빙한다. 문서 열람 경로가 이
 * 자가치유 때문에 깨지는 일은 없어야 한다: 재생성은 "되면 좋은 것"이지 열람의 전제조건이 아니다.
 *
 * **알려진 쓰기 경합(리뷰 지적)**: 이 함수가 부르는 saveGridDerived(spreadsheet.ts)에는 이 경로가
 * 같은 문서의 동시편집 자동저장(collab/room.ts의 flushToStore)과 락 없이 동시에 돌 때의 좁고
 * 자가치유되는 grid/gridindex 쓰기 경합이 문서화돼 있다 — saveGridDerived의 문서 주석 참고. 이
 * 뷰어 재생성 경로는 그 경합의 세 번째 writer일 뿐이라 그 문서 주석을 그대로 적용받는다.
 */
async function regenerateStaleGridCache(docId: string): Promise<boolean> {
  const meta = await store.readMeta(docId);
  if (!meta) return false;
  try {
    if (meta.ext === 'xlsx') {
      // 큰 파일은 스트리밍 읽기(xlsx/viewerGrid.ts)
      const built = await buildViewerGridFile(await store.readOriginal(docId, meta.ext));
      await saveGridFile(store, docId, built.buffer);
    } else {
      const grid = await buildGridFromOriginal(store, meta);
      await saveGridDerived(store, docId, grid);
    }
    return true;
  } catch (error) {
    console.error(`낡은 grid 캐시 재생성 실패(기존 캐시를 그대로 서빙): 문서 ${docId}:`, error);
    return false;
  }
}

async function loadGridUncached(docId: string, allowRegenerate = true): Promise<LoadedGrid> {
  let buffer: Buffer;
  try {
    buffer = await store.readDerived(docId, 'grid');
  } catch (error) {
    // 아직 만들지 않았다 — 메모리 저장 모드(향기랩)는 편집기로만 열어서 보기용 그리드를 미리 만들지 않는다
    // (documents.ts). 누가 보기 화면을 열면 그때 만든다.
    // 그리드가 아예 없으면 보기 화면이 오류다 — 일시 실패에서 빨리 회복하게 쉬는 시간을 짧게(45초)
    if (!allowRegenerate || !(await gridRegen.run(docId, () => regenerateStaleGridCache(docId), 45_000))) throw error;
    return loadGridUncached(docId, false);
  }
  if (allowRegenerate && gridIsStale(peekGridHeader(buffer))) {
    // 재생성이 성공하면 방금 막 저장한 최신 버퍼를 처음부터 다시 읽어 이어간다(아래 재귀 호출,
    // allowRegenerate=false로 무한 재귀를 막는다) — 실패하면(위 문서 주석) 그냥 아래로 흘러가
    // 지금 손에 있는(낡았을 수 있는) buffer를 기존 경로 그대로 서빙한다.
    if (await gridRegen.run(docId, () => regenerateStaleGridCache(docId))) {
      return loadGridUncached(docId, false);
    }
  }
  try {
    const raw = await store.readDerived(docId, 'gridindex');
    const { lineOffsets } = JSON.parse(raw.toString('utf8')) as { lineOffsets: number[] };
    if (offsetsLookValid(buffer, lineOffsets)) {
      return buildLoadedGridFromOffsets(buffer, lineOffsets);
    }
  } catch {
    // gridindex가 없거나(정상 — 최초 조회), 손상됐거나, 구조적으로 말이 안 되면(offsetsLookValid
    // 실패) 여기로 떨어진다 — 아래에서 다시 스캔해 올바른 값으로 덮어써 자가치유한다.
  }
  const grid = indexGrid(buffer);
  await store.saveDerived(
    docId,
    'gridindex',
    Buffer.from(JSON.stringify({ lineOffsets: grid.lineOffsets })),
  );
  return grid;
}

/** 보기 그리드 재생성은 문서마다 한 번에 하나, 실패하면 10분 쉰다(요청마다 전체를 다시 만드는 폭주 방지) */
const gridRegen = singleFlightWithBackoff(10 * 60_000);
/** 날짜가 바뀐 그리드를 다시 불러오는 중인 것(문서마다 하나) */
const staleReloads = new Map<string, Promise<LoadedGrid>>();

export async function loadGrid(docId: string): Promise<LoadedGrid> {
  const grid = await gridCache.getOrLoad(docId, () => loadGridUncached(docId));
  // 메모리에 남은 그리드도 날짜가 바뀌었으면(TODAY() 조건부 서식) 다시 만든다 — 최근에 재생성이 실패했으면 쉬는 동안은 있는 그대로
  if (!gridIsStale(grid.header) || gridRegen.backingOff(docId)) return grid;
  let reload = staleReloads.get(docId);
  if (!reload) {
    invalidateViewCaches(docId);
    reload = gridCache.getOrLoad(docId, () => loadGridUncached(docId)).finally(() => staleReloads.delete(docId));
    staleReloads.set(docId, reload);
  }
  return reload;
}

/**
 * PDF 페이지 크기를 영구 저장소에서 읽고/쓰는 콜백 쌍 — rendering.ts의 getPageSizes에 주입한다.
 * 변환 시점(convertToPdf)에 이미 저장돼 있으면 그 파일만 읽고 PDF를 다시 파싱하지 않는다.
 * 없으면(이 기능 배포 전에 업로드된 문서) 기존처럼 그 자리에서 계산한 뒤 이 save로 저장해,
 * 다음부터는 그 문서도 빠른 경로를 타게 한다(자가치유).
 */
export function pageSizePersistence(docId: string): PersistedPageSizes {
  return {
    load: async () => {
      try {
        const raw = await store.readDerived(docId, 'pagesizes');
        const sizes = JSON.parse(raw.toString('utf8')) as unknown;
        return Array.isArray(sizes) ? (sizes as readonly PageSize[]) : null;
      } catch {
        return null; // 없음/손상 모두 캐시 미스로 취급 — getPageSizes가 다시 계산해 덮어써 자가치유한다.
      }
    },
    save: async (sizes) => {
      await store.saveDerived(docId, 'pagesizes', Buffer.from(JSON.stringify(sizes)));
    },
  };
}

/** PDF 텍스트 추출 결과 캐시 (문서당 1회 추출) */
function loadPdfText(docId: string): Promise<PageText[]> {
  return pdfTextCache.getOrLoad(docId, () => store.readDerived(docId, 'pdf').then(extractPdfText));
}

interface GridSearchHit {
  readonly sheet: number;
  readonly r: number;
  readonly c: number;
  readonly snippet: string;
}

/** JSONL 그리드에서 부분 문자열을 검색한다. 줄 단위 1차 필터 후 매칭 줄만 파싱한다. */
function searchGrid(grid: LoadedGrid, query: string, limit: number, skipHidden = false): { hits: GridSearchHit[]; total: number } {
  const needle = query.toLowerCase();
  const hits: GridSearchHit[] = [];
  let total = 0;
  for (let sheetIdx = 0; sheetIdx < grid.header.sheets.length; sheetIdx += 1) {
    const sheet = grid.header.sheets[sheetIdx];
    const startLine = grid.sheetStartLine[sheetIdx];
    const hidden = skipHidden ? hiddenLookup(sheet) : null;
    for (let r = 1; r <= sheet.rowCount; r += 1) {
      if (hidden?.row(r - 1)) continue;
      const line = startLine + r - 1;
      const begin = grid.lineOffsets[line];
      const finish =
        line + 1 < grid.lineOffsets.length ? grid.lineOffsets[line + 1] - 1 : grid.buffer.length;
      const raw = grid.buffer.toString('utf8', begin, finish);
      if (!raw.toLowerCase().includes(needle)) continue;
      const cells = JSON.parse(raw) as Array<GridCell | 0>;
      for (let c = 0; c < cells.length; c += 1) {
        const cell = cells[c];
        // 보이는 글자(채움 표시는 빼고)와 값(서식 전 — ####로 가린 칸, (4,321) → -4321) 둘 다로 찾는다
        if (cell === 0 || hidden?.col(c) || !(cell.v.replace(FILL_MARK, '').toLowerCase().includes(needle) || (cell.n ?? '').toLowerCase().includes(needle))) continue;
        total += 1;
        if (hits.length < limit) {
          hits.push({ sheet: sheetIdx, r, c: c + 1, snippet: cell.v.replace(FILL_MARK, ' ').slice(0, 120) });
        }
      }
    }
  }
  return { hits, total };
}

/** 숨긴 행(높이 0)·열(폭 0)인지 — 열람만 토큰에는 그 칸 값을 보내지 않는다 */
function hiddenLookup(sheet: Pick<GridSheet, 'rowHeights' | 'colWidths'>): { row: (r0: number) => boolean; col: (c0: number) => boolean } {
  return {
    row: (r0) => sheet.rowHeights.some(([start, count, px]) => px === 0 && r0 >= start && r0 < start + count),
    col: (c0) => sheet.colWidths[c0] === 0,
  };
}

function maskHiddenCells(sheet: Pick<GridSheet, 'rowHeights' | 'colWidths'>, from: number, rows: Array<Array<GridCell | 0>>): Array<Array<GridCell | 0>> {
  const hidden = hiddenLookup(sheet);
  return rows.map((cells, i) => {
    if (hidden.row(from - 1 + i)) return cells.map(() => 0);
    return cells.map((cell, c) => (cell !== 0 && hidden.col(c) ? 0 : cell));
  });
}

function sliceRows(
  grid: LoadedGrid,
  sheetIdx: number,
  from: number,
  count: number,
): Array<Array<GridCell | 0>> {
  const sheet = grid.header.sheets[sheetIdx];
  const start = grid.sheetStartLine[sheetIdx] + from - 1;
  const end = Math.min(start + count, grid.sheetStartLine[sheetIdx] + sheet.rowCount);
  const rows: Array<Array<GridCell | 0>> = [];
  for (let line = start; line < end; line += 1) {
    const begin = grid.lineOffsets[line];
    const finish = line + 1 < grid.lineOffsets.length ? grid.lineOffsets[line + 1] - 1 : grid.buffer.length;
    rows.push(JSON.parse(grid.buffer.toString('utf8', begin, finish)));
  }
  return rows;
}

/** 차트를 보기 화면용 그림(SVG data URL)으로 — 크기는 그 시트의 열 폭·행 높이로 앵커에서 계산한다 */
function chartImagesOf(charts: readonly GridChart[], grid: LoadedGrid) {
  const rowCache = new Map<string, Array<GridCell | 0>>();
  const valueText = (sheetName: string, r: number, c: number): string => {
    const s = grid.header.sheets.findIndex((sh) => sh.name === sheetName);
    if (s < 0 || r >= grid.header.sheets[s].rowCount) return '';
    const key = `${s}:${r}`;
    let row = rowCache.get(key);
    if (!row) rowCache.set(key, (row = sliceRows(grid, s, r + 1, 1)[0] ?? []));
    const cell = row[c];
    return cell ? gridCellValue(cell) : '';
  };
  return charts.flatMap((chart) => {
    const sheet = grid.header.sheets.find((sh) => sh.name === chart.sheetName);
    if (!sheet) return [];
    const colLeft = (col: number) => {
      let x = 0;
      for (let i = 0; i < col; i += 1) x += sheet.colWidths[i] ?? sheet.defaultColWidth;
      return x;
    };
    const rowTop = (row: number) => rowTopPx(sheet.rowHeights, sheet.defaultRowHeight, row);
    const to = chart.to ?? { row: chart.from.row + 15, col: chart.from.col + 7, dx: 0, dy: 0 };
    const width = Math.round(chart.size?.width ?? colLeft(to.col) + to.dx - colLeft(chart.from.col) - chart.from.dx);
    const height = Math.round(chart.size?.height ?? rowTop(to.row) + to.dy - rowTop(chart.from.row) - chart.from.dy);
    if (!(width > 0 && height > 0)) return [];
    const svg = renderChartSvg(chart.kind, chartDataOf(chart, valueText), width, height);
    return [{
      id: chart.id,
      sheetName: chart.sheetName,
      name: chart.title,
      from: chart.from,
      to: chart.to,
      size: { width, height },
      src: `data:image/svg+xml;base64,${Buffer.from(svg, 'utf8').toString('base64')}`,
    }];
  });
}

/**
 * 그리드 문서의 PDF는 업로드 시 만들지 않고 출력 요청 시 생성한다.
 * (대형 리스트의 수 분짜리 LibreOffice 변환을 열람 경로에서 제거)
 */
async function ensurePdf(meta: DocumentMeta): Promise<Buffer> {
  if (await store.hasDerived(meta.id, 'pdf')) {
    return store.readDerived(meta.id, 'pdf');
  }
  const converter = findConverter(meta.ext);
  if (!converter) {
    throw new Error(`PDF 변환기를 찾을 수 없습니다: ${meta.ext}`);
  }
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-print-'));
  try {
    const sourcePath = path.join(workDir, `source.${meta.ext}`);
    await fs.writeFile(sourcePath, await store.readOriginal(meta.id, meta.ext));
    const targetPdfPath = path.join(workDir, 'converted.pdf');
    await converter.convert(sourcePath, targetPdfPath, meta.name);
    const pdfBytes = await fs.readFile(targetPdfPath);
    await store.saveDerived(meta.id, 'pdf', pdfBytes);
    return pdfBytes;
  } finally {
    // Windows에서 백신 등이 방금 만든 임시 파일을 잠깐 스캔하는 동안 삭제가 EBUSY로 실패하는
    // 경우가 있어, 짧게 재시도한다.
    await fs.rm(workDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 });
  }
}

// token 누락/무효는 resolveView가 401로 일관 처리하도록 스키마에서 required로 강제하지 않는다.
const tokenQuerySchema = {
  type: 'object',
  properties: { token: { type: 'string' } },
} as const;

/** /api/view/info — 워드 편집기만 editor=1을 붙여 편집기 화면 정보(docxRender)를 받는다 */
const infoQuerySchema = {
  type: 'object',
  properties: { token: { type: 'string' }, editor: { type: 'string' } },
} as const;

/** 워드 편집기 화면 정보 — 문서 판(원본 바이트 sha1)마다 한 번만 만든다(FINAL-REVIEW I1). 크기는 JSON 길이로 어림한다 */
const docxRenderCache = new SizedLruCache<ReturnType<typeof getDocxRenderExtras>>({
  maxBytes: 32 * 1024 * 1024,
  maxEntries: 64,
  sizeOf: (value) => JSON.stringify(value).length * 2,
});
function docxRenderOf(docId: string, docxBytes: Uint8Array): Promise<ReturnType<typeof getDocxRenderExtras>> {
  const version = `${docId}:${createHash('sha1').update(docxBytes).digest('base64url')}`;
  return docxRenderCache.getOrLoad(version, async () => getDocxRenderExtras(docxBytes));
}

const presenceBodySchema = {
  type: 'object',
  required: ['sessionId'],
  additionalProperties: false,
  properties: { sessionId: { type: 'string', minLength: 1, maxLength: 100 } },
} as const;

interface PresenceBody {
  readonly sessionId: string;
}

const pageRequestSchema = {
  params: {
    type: 'object',
    required: ['page'],
    properties: { page: { type: 'integer', minimum: 1 } },
  },
  querystring: {
    type: 'object',
    properties: {
      token: { type: 'string' },
      scale: { type: 'integer', minimum: 1, maximum: 16, default: 2 },
      /** 슬라이드 창 미리보기 — 0.5배로 작게(본문 2배 그림은 장당 0.3~1.7MB였다) */
      thumb: { type: 'boolean' },
    },
  },
} as const;

/** 미리보기 배율 — 슬라이드 창 썸네일 폭(약 190px)을 고해상도 화면에서도 선명하게 */
const THUMB_SCALE = 0.5;

const gridRowsSchema = {
  type: 'object',
  properties: {
    token: { type: 'string' },
    sheet: { type: 'integer', minimum: 0, default: 0 },
    from: { type: 'integer', minimum: 1, default: 1 },
    count: { type: 'integer', minimum: 1, maximum: 500, default: 200 },
  },
} as const;

interface TokenQuery {
  readonly token?: string;
}

const tokenWithMarkupQuerySchema = {
  type: 'object',
  properties: { token: { type: 'string' }, includeMarkup: { type: 'string' } },
} as const;

interface TokenWithMarkupQuery extends TokenQuery {
  readonly includeMarkup?: string;
}

const addMarkupBodySchema = {
  type: 'object',
  required: ['page', 'x', 'y', 'text'],
  additionalProperties: false,
  properties: {
    kind: { type: 'string', enum: ['comment', 'text'] },
    page: { type: 'integer', minimum: 1 },
    x: { type: 'number', minimum: 0, maximum: 1 },
    y: { type: 'number', minimum: 0, maximum: 1 },
    w: { type: 'number', minimum: 0, maximum: 1 },
    h: { type: 'number', minimum: 0, maximum: 1 },
    fontSizePt: { type: 'number', minimum: 1, maximum: 400 },
    bold: { type: 'boolean' },
    fontFamily: { type: 'string', enum: ['sans-serif', 'serif', 'monospace'] },
    text: { type: 'string', minLength: 1, maxLength: 2000 },
  },
  if: { properties: { kind: { const: 'text' } }, required: ['kind'] },
  then: { required: ['fontSizePt'] },
  else: { required: ['w', 'h'] },
} as const;

interface AddMarkupBody {
  readonly kind?: 'comment' | 'text';
  readonly page: number;
  readonly x: number;
  readonly y: number;
  readonly w?: number;
  readonly h?: number;
  readonly fontSizePt?: number;
  readonly bold?: boolean;
  readonly fontFamily?: 'sans-serif' | 'serif' | 'monospace';
  readonly text: string;
}

const updateMarkupBodySchema = {
  type: 'object',
  required: ['x', 'y'],
  additionalProperties: false,
  properties: {
    x: { type: 'number', minimum: 0, maximum: 1 },
    y: { type: 'number', minimum: 0, maximum: 1 },
    fontSizePt: { type: 'number', minimum: 1, maximum: 400 },
  },
} as const;

interface UpdateMarkupBody {
  readonly x: number;
  readonly y: number;
  readonly fontSizePt?: number;
}

const markupParamsSchema = {
  type: 'object',
  required: ['markupId'],
  properties: { markupId: { type: 'string' } },
} as const;

const textStyleSchema = {
  params: {
    type: 'object',
    required: ['page'],
    properties: { page: { type: 'integer', minimum: 1 } },
  },
  querystring: {
    type: 'object',
    required: ['x', 'y'],
    properties: {
      token: { type: 'string' },
      x: { type: 'number', minimum: 0, maximum: 1 },
      y: { type: 'number', minimum: 0, maximum: 1 },
    },
  },
} as const;

interface TextStyleQuery extends TokenQuery {
  readonly x: number;
  readonly y: number;
}

interface PageQuery extends TokenQuery {
  readonly thumb?: boolean;
  readonly scale: number;
}

interface GridRowsQuery extends TokenQuery {
  readonly sheet: number;
  readonly from: number;
  readonly count: number;
}

const searchQuerySchema = {
  type: 'object',
  required: ['q'],
  properties: {
    token: { type: 'string' },
    q: { type: 'string', minLength: 1, maxLength: 200 },
    limit: { type: 'integer', minimum: 1, maximum: 500, default: 200 },
  },
} as const;

interface SearchQuery extends TokenQuery {
  readonly q: string;
  readonly limit: number;
}

const slideMediaQuerySchema = {
  type: 'object',
  required: ['path'],
  properties: {
    token: { type: 'string' },
    // ppt/media/·word/media/ 바로 아래(하위 폴더 없음) 파일만 — 경로 순회나 zip의 다른 파트(슬라이드 XML 등)
    // 접근을 막는다. parseSlideVisuals·docx 가져오기(mediaPathOf)가 만드는 mediaPath 형태와 정확히 일치해야 한다.
    path: { type: 'string', pattern: '^(ppt|word)/media/[^/]+$', maxLength: 300 },
  },
} as const;

interface SlideMediaQuery extends TokenQuery {
  readonly path: string;
}

/** 무거운 작업(LibreOffice 변환·암호 풀기·서명) — 접속 IP마다 분당 상한(비밀번호 맞히기·변환 대기열 채우기 방지) */
const HEAVY_RATE_LIMIT = { max: 20, timeWindow: '1 minute' } as const;

export async function viewerRoutes(app: FastifyInstance): Promise<void> {
  const legacyConvertLimit = app.createRateLimit(HEAVY_RATE_LIMIT);
  app.get<{ Querystring: TokenQuery & { editor?: string } }>(
    '/api/view/info',
    { schema: { querystring: infoQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      await recordAudit({ type: 'opened', docId: view.meta.id, viewer: view.claims.viewer });
      const viewKind = view.meta.viewKind ?? 'pages';
      // 페이지 크기를 미리 내려 뷰어가 로딩 전 프레임 크기를 고정한다 (스크롤 점프 정확도)
      const pages =
        viewKind === 'pages'
          ? await getPageDimensions(
              view.meta.id,
              () => store.readDerived(view.meta.id, 'pdf'),
              pageSizePersistence(view.meta.id),
            ).catch(() => undefined)
          : undefined;
      // 아래 네 값 모두 docx일 때만 필요하고 전부 같은 원본 파일을 본다 — 예전에는 각자
      // store.readOriginal을 따로 불러 같은(수 MB일 수 있는) 파일을 최대 4번 읽고 4번 다시
      // zip/XML 파싱했다. 한 번만 읽어 네 파서에 그대로 넘기고, 서로 독립적으로 실패할 수 있게
      // (한 파서가 손상된 문서 때문에 던져도 나머지는 자기 값을 그대로 쓰게) 각각 개별
      // catch로 감싼 채 병렬로 돌린다.
      const docxBytes =
        view.meta.ext === 'docx' ? await store.readOriginal(view.meta.id, view.meta.ext).catch(() => null) : null;
      const [pageSetup, defaultFontSizePt, defaultFontFamily, defaultParagraphSpacing, docxRender] = await Promise.all([
        // docx 에디터가 실제 Word 페이지 크기·여백 비율로 화면을 그리기 위한 값(src/client/editor.ts).
        docxBytes ? Promise.resolve().then(() => getDocxPageSetup(docxBytes)).catch(() => null) : Promise.resolve(null),
        // 직접 서식도 스타일도 없는 실행문이 상속받는 문서 기본 폰트 크기(w:docDefaults) —
        // 없으면 브라우저 기본값(16px)이 적용되어 Word가 실제로 보여주는 크기와 달라진다.
        docxBytes
          ? Promise.resolve().then(() => getDocxDefaultFontSizePt(docxBytes)).catch(() => null)
          : Promise.resolve(null),
        // 문서 기본 서체(word/theme/theme1.xml로 해석) — 없으면 null(에디터가 기존 폰트를 그대로 씀).
        docxBytes
          ? Promise.resolve().then(() => getDocxDefaultFontFamily(docxBytes)).catch(() => null)
          : Promise.resolve(null),
        // 직접 간격이 없는 문단의 줄 간격·앞뒤 간격(문서 기본값 + 표준 스타일) — 없으면 편집기 기존 값
        docxBytes
          ? Promise.resolve().then(() => getDocxDefaultParagraphSpacing(docxBytes)).catch(() => null)
          : Promise.resolve(null),
        // 번호 매기기 정의·머리글/바닥글·각주·스타일 — 편집기가 Word 모양으로 덧그린다(src/client/docxLayout.ts)
        // 편집기 요청(editor=1)에만 — 읽기 전용 보기(view.html·viewer.js)도 이 경로를 불러, 예전엔 docx를 열 때마다 원본 전체를 풀고 파싱했다
        docxBytes && request.query.editor === '1'
          ? docxRenderOf(view.meta.id, docxBytes).then((extras) => ({ ...extras, lineRatios: installedLineRatios() })).catch(() => null)
          : Promise.resolve(null),
      ]);
      // pptx 편집기가 슬라이드를 실제 크기(4:3 등) 좌표계로 그리기 위한 값. presentation.xml 하나만
      // 압축 해제한다(filter) — 큰 덱이라도 미디어 파트는 풀지 않는다.
      const slideSize =
        view.meta.ext === 'pptx'
          ? await store
              .readOriginal(view.meta.id, view.meta.ext)
              .then((bytes) => readSlideSizeEmu(unzipSync(bytes, { filter: (f) => f.name === 'ppt/presentation.xml' })))
              .catch(() => null)
          : null;
      return {
        success: true,
        data: {
          // 편집기(editor.js)가 협업 WebSocket 룸 이름으로 사용한다 — 토큰이 이미 이 문서로
          // 범위가 고정돼 있으므로 docId 자체를 별도로 노출해도 새로운 권한 노출은 없다.
          docId: view.meta.id,
          name: view.meta.name,
          ext: view.meta.ext,
          pageCount: view.meta.pageCount,
          viewKind,
          pages,
          permissions: view.claims.perms,
          viewer: view.claims.viewer,
          pageSetup,
          defaultFontSizePt,
          defaultFontFamily,
          defaultParagraphSpacing,
          docxRender,
          slideSize,
          // 통합 보기 껍데기(view.html)의 편집 단추와 편집기의 "보기로"가 쓴다 — 토큰 하나로 오가므로 새 토큰을 만들지 않는다
          editorUrl: editorUrlFor(
            view.meta.ext,
            { effectiveCanEdit: view.claims.perms.canEdit, canDownload: view.claims.perms.canDownload, pdfEditor: view.claims.perms.pdfEditor === true },
            request.query.token!,
          ),
          viewerUrl: `${viewPathFor(view.meta.ext, view.meta.viewKind)}?token=${encodeURIComponent(request.query.token!)}`,
          processing: view.meta.status === 'processing',
          // 향기랩 옛 형식(.doc·.xls·.ppt) — 편집기가 형식 이름·내려받기 문구·매크로 경고를 보인다(설계 §2.5)
          // .doc 이름의 OOXML(ZIP)은 실제로 OOXML이라 옛 형식 문구를 보이지 않는다(내려받기는 그래도 안 고쳤으면 올린 바이트 그대로)
          legacy:
            view.meta.legacy && LEGACY_EDIT_TARGET[view.meta.legacy.ext] === view.meta.ext
              ? { ext: view.meta.legacy.ext, hasMacros: view.meta.legacy.hasMacros }
              : undefined,
        },
      };
    },
  );

  /**
   * 변환 상태 — 보기 화면이 첫 변환(editReady)을 기다리는 동안 부른다("쪽을 준비하는 중", 설계 V8). 감사 기록을 남기지 않고,
   * 변환에 실패한 문서(열 수 없는 상태)도 그 문구를 돌려준다.
   */
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/status',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      let claims: ViewClaims;
      try {
        claims = verifyViewToken(request.query.token ?? '');
      } catch {
        return reply.code(401).send({ success: false, error: '열람 토큰이 없거나 유효하지 않습니다.' });
      }
      const meta = await store.readMeta(claims.docId);
      if (!meta) return reply.code(404).send({ success: false, error: '문서를 찾을 수 없습니다.' });
      return {
        success: true,
        data: {
          status: meta.status,
          progress: meta.progress ?? 0,
          pageCount: meta.pageCount,
          ...(meta.status === 'failed' ? { error: meta.error ?? '문서를 변환하지 못했습니다.' } : {}),
        },
      };
    },
  );

  /**
   * 보기용 PDF(설계 V5) — 통합 보기 껍데기가 내려받기 권한이 있을 때 pdf.js로 그린다. 권한이 없으면 403(PDF가 브라우저로 가지
   * 않는다). 원본이 PDF면 지금 상태(pdf-current와 같은 규칙: 편집 방이 열려 있으면 저장 전 상태, 암호 파일은 비밀번호를 맞힌 토큰만),
   * 그 밖에는 변환한 PDF. 첫 변환 중이면 202. 내려받기와 달리 보기라서 파일 이름·감사 기록을 붙이지 않는다.
   */
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/derived-pdf',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canDownload) {
        return reply.code(403).send({ success: false, error: '저장(다운로드) 권한이 없습니다.' });
      }
      const meta = view.meta;
      try {
        let bytes: Uint8Array;
        if (meta.ext === 'pdf') {
          bytes = await currentPlainBytes(meta.id, () => store.readOriginal(meta.id, meta.ext), tokenKey(request.query.token!));
        } else if (meta.status === 'processing') {
          return reply.code(202).send({ success: true, data: { processing: true } });
        } else if ((meta.viewKind ?? 'pages') !== 'pages') {
          return reply.code(400).send({ success: false, error: '쪽으로 보는 문서가 아닙니다.' });
        } else {
          bytes = await store.readDerived(meta.id, 'pdf');
        }
        return await reply
          .type('application/pdf')
          .header('cache-control', 'no-store')
          .send(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength));
      } catch (error) {
        if (error instanceof PdfRoomError) return reply.code(error.status).send({ success: false, error: error.message, locked: error.status === 423 });
        request.log.error(error, 'derived pdf failed');
        return reply.code(500).send({ success: false, error: '보기용 PDF를 불러오지 못했습니다.' });
      }
    },
  );

  /**
   * "지금 이 문서를 보고 있는 사람" 표시용 하트비트 — pages/grid 등 viewKind와 무관하게 viewer.html을
   * 여는 모든 문서에 동일하게 적용된다(src/presence.ts). 클라이언트가 주기적으로 호출하며, 응답에
   * 갱신된 현재 접속자 목록을 실어 보내 별도 조회 요청이 필요 없게 한다.
   */
  app.post<{ Querystring: TokenQuery; Body: PresenceBody }>(
    '/api/view/presence',
    { schema: { querystring: tokenQuerySchema, body: presenceBodySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      const viewers = touchPresence(view.meta.id, request.body.sessionId, view.claims.viewer);
      return { success: true, data: { viewers } };
    },
  );

  // Ctrl+S — 8초 유휴 자동저장을 기다리지 않고 지금 바로 저장한다(Office 365의 수동 저장과 동일한
  // 역할이지만, 이 앱은 이미 계속 자동저장하므로 사실상 "저장 큐를 지금 비워라"에 가깝다).
  // 엑셀 편집기가 그릴 파일 속 차트(종류·제목·참조 범위·위치)
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/xlsx-charts',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveEditorDataView(request.query.token, reply);
      if (!view) return reply;
      if (view.meta.ext !== 'xlsx') return { success: true, data: { charts: [] } };
      try {
        return { success: true, data: { charts: await getFileCharts(view.meta.id) } };
      } catch (error) {
        request.log.warn(error, 'xlsx chart read failed');
        return { success: true, data: { charts: [] } };
      }
    },
  );

  // 파일에 있는 그림(로고·사진) — 보기·편집 화면이 칸 위치에 맞춰 그대로 얹는다.
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/xlsx-images',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (view.meta.ext !== 'xlsx') return { success: true, data: { images: [] } };
      try {
        return { success: true, data: { images: await getFileImages(view.meta.id) } };
      } catch (error) {
        request.log.warn(error, 'xlsx image read failed');
        return { success: true, data: { images: [] } };
      }
    },
  );

  // 파일에 있는 차트 — 보기 화면은 편집 화면과 같은 그리기(xlsx/chartSvg.ts)로 서버가 만든 SVG를 그림처럼 칸 위치에 얹는다.
  // 값은 지금 그리드의 칸 값(편집 후 저장된 값)으로 다시 계산하고, 칸이 비면 파일에 저장된 값을 쓴다.
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/xlsx-chart-images',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (view.meta.ext !== 'xlsx') return { success: true, data: { images: [] } };
      try {
        const [charts, grid] = await Promise.all([getFileCharts(view.meta.id), loadGrid(view.meta.id)]);
        return { success: true, data: { images: chartImagesOf(charts, grid) } };
      } catch (error) {
        request.log.warn(error, 'xlsx chart image failed');
        return { success: true, data: { images: [] } };
      }
    },
  );

  // 보기 화면(viewer.js)의 하이퍼링크 — 열람만 토큰도 받는다(엑셀에서도 보이는 바로가기다). 목록 상자 값·표 정보는 주지 않고,
  // 열람만 토큰에는 숨긴 행·열 칸에 걸린 링크와, 주소를 칸 값으로 만드는 수식 링크(expr — 수식 글이다)를 뺀다
  // (열람만에는 숨긴 칸 값·수식을 보내지 않는 것과 같다 — maskHiddenCells)
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/xlsx-links',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (view.meta.ext !== 'xlsx' || (view.meta.viewKind ?? 'pages') !== 'grid') return { success: true, data: { links: [] } };
      try {
        let { links } = await getFileExtras(view.meta.id);
        if (isViewOnly(view.claims.perms)) {
          const sheets = new Map((await loadGrid(view.meta.id)).header.sheets.map((s) => [s.name, hiddenLookup(s)]));
          links = links.filter((l) => {
            const hidden = sheets.get(l.sheetName);
            return hidden && !l.expr && !hidden.row(l.r1) && !hidden.col(l.c1);
          });
        }
        return { success: true, data: { links } };
      } catch (error) {
        request.log.warn(error, 'xlsx links read failed');
        return { success: true, data: { links: [] } };
      }
    },
  );

  // 파일의 하이퍼링크(바로가기)·목록 상자(데이터 유효성 "목록") — 편집 화면이 링크 이동·목록 고르기에 쓴다.
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/xlsx-extras',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveEditorDataView(request.query.token, reply);
      if (!view) return reply;
      if (view.meta.ext !== 'xlsx') return { success: true, data: { links: [], lists: [] } };
      try {
        return { success: true, data: await getFileExtras(view.meta.id) };
      } catch (error) {
        request.log.warn(error, 'xlsx extras read failed');
        return { success: true, data: { links: [], lists: [] } };
      }
    },
  );

  // 편집 화면 조건부 서식 — 범위 전체를 보아야 하는 규칙(COUNTIF($A:$A,$A1)·중복·상위·색조 …)용 범위 값 개수 표.
  // 편집 화면은 보고 있는 페이지만 받으므로 서버(편집 방의 지금 상태)가 만들어 준다(src/xlsx/cfRangeValues.ts).
  app.post<{ Querystring: TokenQuery; Body: unknown }>(
    '/api/view/xlsx-cf-ranges',
    { schema: { querystring: tokenQuerySchema }, bodyLimit: 131_072, config: { rateLimit: CF_REMOTE_RATE_LIMIT } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      // 편집 방 전용 — 편집 권한 토큰만(내려받기만 되는 토큰은 열린 편집 방에 일을 시킬 수 없다)
      if (!view.claims.perms.canEdit) return reply.code(403).send({ success: false, error: '편집 권한이 없습니다.' });
      const req = parseCfRemoteRequest(request.body);
      if (view.meta.ext !== 'xlsx' || !req) {
        return reply.code(400).send({ success: false, error: '잘못된 요청입니다.' });
      }
      const token = request.query.token!;
      const pass = cfTokenBudget(token);
      const startLeft = pass.left;
      let data;
      try {
        data = await getCfRemoteAnswer(view.meta.id, req, pass);
      } catch (error) {
        if (!(error instanceof CfBusyError)) throw error;
        return reply.code(409).send({ success: false, error: '조건부 서식 계산이 이미 진행 중입니다. 잠시 뒤 다시 시도하세요.' });
      } finally {
        cfTokenSpend(token, pass, startLeft);
      }
      if (!data) return reply.code(409).send({ success: false, error: '편집 방이 열려 있지 않습니다.' });
      return { success: true, data };
    },
  );

  // 검토 탭 "변경 내용" — 이 세션(룸)이 열린 뒤 누가 어느 문단·칸·슬라이드를 바꿨는지. 열람 토큰만 있으면
  // 볼 수 있다(문서 내용을 이미 보고 있는 사람에게 같은 내용의 변경 이력을 보여주는 것이다).
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/changes',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveEditorDataView(request.query.token, reply);
      if (!view) return reply;
      return { success: true, data: { changes: getChangeLog(view.meta.id) } };
    },
  );

  app.post<{ Querystring: TokenQuery }>(
    '/api/view/save',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canEdit) {
        return reply.code(403).send({ success: false, error: '편집 권한이 없습니다.' });
      }
      const saved = await flushNow(view.meta.id);
      return { success: true, data: { saved } };
    },
  );

  app.get<{ Params: { page: number }; Querystring: PageQuery }>(
    '/api/view/pages/:page',
    { schema: pageRequestSchema },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      // 변환 전에 연 편집기(editReady — pptx 등)의 쪽 그림 요청은 첫 변환이 끝날 때까지 기다렸다 준다
      await waitForInitialConversion(view.meta.id);
      try {
        const png = await renderPage(
          view.meta.id,
          request.params.page,
          request.query.thumb ? THUMB_SCALE : request.query.scale,
          () => store.readDerived(view.meta.id, 'pdf'),
          pageSizePersistence(view.meta.id),
          request.query.thumb === true,
        );
        return await reply.type('image/png').header('cache-control', 'private, no-store').send(png);
      } catch (error) {
        if (error instanceof PageOutOfRangeError) {
          return reply.code(404).send({ success: false, error: error.message });
        }
        request.log.error(error, 'page rendering failed');
        return reply.code(500).send({ success: false, error: '페이지 렌더링에 실패했습니다.' });
      }
    },
  );

  /** 도면(dxf/dwg) 벡터 뷰 — LibreOffice가 만든 SVG를 그대로 내려준다(래스터 재변환 없음). */
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/svg',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if ((view.meta.viewKind ?? 'pages') !== 'svg') {
        return reply.code(400).send({ success: false, error: '도면 벡터 뷰 문서가 아닙니다.' });
      }
      try {
        const svg = await store.readDerived(view.meta.id, 'svg');
        return await reply.type('image/svg+xml').header('cache-control', 'private, no-store').header('content-security-policy', SANDBOX_CSP).send(svg);
      } catch (error) {
        request.log.error(error, 'svg view failed');
        return reply.code(500).send({ success: false, error: 'SVG를 불러오지 못했습니다.' });
      }
    },
  );

  /** pptx 슬라이드 안 이미지(p:pic)와 docx 본문·머리글 그림(word/media/) 원본 바이트를 그대로 내려준다 — editor-slides.ts·docx 편집기(schema.ts ooxml 그림)가
   * slideVisuals의 mediaPath로 이 라우트를 <img src>에 직접 건다. 수 MB짜리 이미지를 base64로
   * Yjs 문서에 인라인하던 예전 방식은 그 크기 그대로 WebSocket 업데이트에 실려 텍스트박스보다
   * 몇 초씩 늦게 나타나는 원인이었다 — 평범한 HTTP 이미지 요청으로 바꿔 브라우저가 캐시·병렬
   * 로딩을 알아서 하게 한다. */
  app.get<{ Querystring: SlideMediaQuery }>(
    '/api/view/slide-media',
    { schema: { querystring: slideMediaQuerySchema } },
    async (request, reply) => {
      const view = await resolveEditorDataView(request.query.token, reply);
      if (!view) return reply;
      const mediaPath = request.query.path;
      // pptx는 ppt/media/, docx는 word/media/ — 문서 종류와 경로가 맞아야 한다
      const kind = view.meta.ext === 'pptx' ? 'ppt' : view.meta.ext === 'docx' ? 'word' : null;
      if (!kind || !mediaPath.startsWith(`${kind}/media/`)) {
        return reply.code(400).send({ success: false, error: 'pptx·docx 문서의 그림이 아닙니다.' });
      }
      const ext = mediaPath.split('.').pop()?.toLowerCase() ?? '';
      const mime = IMAGE_MIME_BY_EXT[ext];
      if (!mime) {
        return reply.code(400).send({ success: false, error: '지원하지 않는 이미지 형식입니다.' });
      }
      try {
        const bytes = await store.readOriginal(view.meta.id, view.meta.ext);
        // 요청한 그림 하나만 푼다 — 예전엔 그림 하나마다 파일 전체(실측 21MB pptx)를 풀어, 그동안 서버가
        // 다른 요청을 못 받았다(그림당 약 0.1초, 동기 처리)
        const zipEntries = guardedUnzipSync(bytes, 'slide-media', { filter: (file) => file.name === mediaPath });
        const mediaBytes = zipEntries[mediaPath];
        if (!mediaBytes) {
          return reply.code(404).send({ success: false, error: '이미지를 찾을 수 없습니다.' });
        }
        // 편집기는 그림을 바꾸지 않는다 — 슬라이드를 오갈 때마다 수 MB를 다시 받지 않게 브라우저가 1시간 보관한다
        // (주소에 열람 토큰이 있어 그 사람의 브라우저에만 남는다: private)
        return await reply.type(mime).header('cache-control', 'private, max-age=3600').header('content-security-policy', SANDBOX_CSP).send(Buffer.from(mediaBytes));
      } catch (error) {
        request.log.error(error, 'slide media fetch failed');
        return reply.code(500).send({ success: false, error: '이미지를 불러오지 못했습니다.' });
      }
    },
  );

  /**
   * docx 편집기의 브라우저가 못 그리는 그림 조각(EMF/WMF·차트·도형) — 이 문서(열린 방이면 그 저장소까지)에 있는 조각만 LibreOffice로 그려 PNG로.
   * 문서 판마다 그런 조각을 한 번에 그려 두므로, 그린 그림·묶음이 있으면 바로 돌려주고 무거운 첫 그리기에만 분당 상한을 센다.
   */
  const drawingLimit = app.createRateLimit(HEAVY_RATE_LIMIT);
  app.get<{ Querystring: { token?: string; ref: string } }>(
    '/api/view/docx-drawing',
    { schema: { querystring: { type: 'object', required: ['ref'], properties: { token: { type: 'string' }, ref: { type: 'string', pattern: '^[A-Za-z0-9:._-]{1,64}@[A-Za-z0-9_-]{20}(#[0-9]+)?$' } } } } },
    async (request, reply) => {
      const view = await resolveEditorDataView(request.query.token, reply);
      if (!view) return reply;
      if (view.meta.ext !== 'docx') return reply.code(400).send({ success: false, error: 'docx 문서가 아닙니다.' });
      const docId = view.meta.id;
      // LibreOffice로 새 묶음을 그릴 때만 센다(저장으로 문서 판이 바뀐 뒤의 다시 그리기 포함) — 그려 둔 쪽·그림은 세지 않는다
      // isAllowed는 허용 목록일 때만 true — 넘었는지는 isExceeded
      const allowNewBatch = async () => {
        const limit = await drawingLimit(request);
        return limit.isAllowed || !limit.isExceeded;
      };
      try {
        const png = cachedDocxDrawing(docId, request.query.ref)
          ? await cachedDocxDrawing(docId, request.query.ref)!
          : await renderDocxDrawingPng(() => store.readOriginal(docId, 'docx'), docId, request.query.ref, docxFragmentsOf(docId), allowNewBatch);
        if (!png) return reply.code(404).send({ success: false, error: '그림을 찾을 수 없습니다.' });
        return await reply.type('image/png').header('cache-control', 'private, max-age=3600').header('content-security-policy', SANDBOX_CSP).send(png);
      } catch (error) {
        if (error instanceof DrawingRateLimited) return reply.code(429).send({ success: false, error: '잠시 뒤 다시 시도하세요.' });
        request.log.error(error, 'docx drawing render failed');
        return reply.code(500).send({ success: false, error: '그림을 그리지 못했습니다.' });
      }
    },
  );

  /** 그리드 메타: 시트 목록·크기·병합·스타일 팔레트 (셀 데이터 제외) */
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/grid/meta',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if ((view.meta.viewKind ?? 'pages') !== 'grid') {
        return reply.code(400).send({ success: false, error: '그리드 뷰 문서가 아닙니다.' });
      }
      const grid = await loadGrid(view.meta.id);
      // 테두리 선 종류·색, 무늬 채우기, 글자 회전·들여쓰기 CSS는 편집 화면과 같은 함수(collab/cellBoxCss.ts)로 만들어 함께 보낸다
      // 서식 값은 한 번 더 검사한다(큰 파일은 편집 방의 서식이 그대로 그리드 파일에 들어간다) — 보기 화면 <style>로 들어가는 값이다
      const styles = grid.header.styles.map((style) => {
        const safe = sanitizeCellStyle(style);
        return { ...safe, ...cellBoxCssText(safe) };
      });
      return { success: true, data: { ...grid.header, styles } };
    },
  );

  /** 그리드 행 데이터를 범위 단위로 반환한다 (가상 스크롤용) */
  app.get<{ Querystring: GridRowsQuery }>(
    '/api/view/grid/rows',
    { schema: { querystring: gridRowsSchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if ((view.meta.viewKind ?? 'pages') !== 'grid') {
        return reply.code(400).send({ success: false, error: '그리드 뷰 문서가 아닙니다.' });
      }
      const grid = await loadGrid(view.meta.id);
      if (!grid.header.sheets[request.query.sheet]) {
        return reply.code(404).send({ success: false, error: '시트를 찾을 수 없습니다.' });
      }
      const from = request.query.from;
      let rows = sliceRows(grid, request.query.sheet, from, request.query.count);
      if (isViewOnly(view.claims.perms)) rows = maskHiddenCells(grid.header.sheets[request.query.sheet], from, rows);
      return { success: true, data: { from, rows } };
    },
  );

  /**
   * 문서 내 문자 검색.
   * grid 문서: 셀 값 검색 → {kind:'grid', hits:[{sheet,r,c,snippet}]}
   * pages 문서: PDF 텍스트 레이어 검색 → {kind:'pages', hits:[{page,rects,snippet}]}
   * rects는 페이지 크기 대비 비율(0~1)이라 클라이언트가 배율과 무관하게 하이라이트를 그린다.
   */
  app.get<{ Querystring: SearchQuery }>(
    '/api/view/search',
    { schema: { querystring: searchQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      const { q, limit } = request.query;
      try {
        if ((view.meta.viewKind ?? 'pages') === 'grid') {
          const grid = await loadGrid(view.meta.id);
          const { hits, total } = searchGrid(grid, q, limit, isViewOnly(view.claims.perms));
          return { success: true, data: { kind: 'grid', hits, total } };
        }
        const pages = await loadPdfText(view.meta.id);
        const { hits, total } = searchPdfText(pages, q, limit);
        // 복사 금지 토큰에는 글 조각을 주지 않는다 — 검색어를 바꿔 가며 조각을 모으면 본문을 되살릴 수 있었다(화면은 위치만 쓴다)
        const shown = view.claims.perms.canCopy ? hits : hits.map((h) => ({ ...h, snippet: '' }));
        return { success: true, data: { kind: 'pages', hits: shown, total } };
      } catch (error) {
        request.log.error(error, 'search failed');
        return reply.code(500).send({ success: false, error: '검색에 실패했습니다.' });
      }
    },
  );

  app.get<{ Querystring: TokenQuery }>(
    '/api/view/markup',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      const items = await listMarkup(view.meta.id);
      return { success: true, data: items };
    },
  );

  app.post<{ Querystring: TokenQuery; Body: AddMarkupBody }>(
    '/api/view/markup',
    { schema: { querystring: tokenQuerySchema, body: addMarkupBodySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canEdit) {
        await recordAudit({ type: 'markup_add_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '마크업 추가 권한이 없습니다.' });
      }
      if ((view.meta.viewKind ?? 'pages') !== 'pages') {
        return reply.code(400).send({ success: false, error: '이 문서 형식은 마크업을 지원하지 않습니다.' });
      }
      try {
        const body = request.body;
        const item = await addMarkup(
          view.meta.id,
          view.meta,
          body.kind === 'text'
            ? {
                kind: 'text',
                page: body.page,
                x: body.x,
                y: body.y,
                fontSizePt: body.fontSizePt ?? 12,
                bold: body.bold ?? false,
                fontFamily: body.fontFamily,
                text: body.text,
                author: view.claims.viewer,
              }
            : {
                kind: 'comment',
                page: body.page,
                x: body.x,
                y: body.y,
                w: body.w ?? 0,
                h: body.h ?? 0,
                text: body.text,
                author: view.claims.viewer,
              },
        );
        await recordAudit({ type: 'markup_added', docId: view.meta.id, viewer: view.claims.viewer, detail: item.id });
        return await reply.code(201).send({ success: true, data: item });
      } catch (error) {
        if (error instanceof MarkupValidationError) {
          return reply.code(400).send({ success: false, error: error.message });
        }
        request.log.error(error, 'markup add failed');
        return reply.code(500).send({ success: false, error: '마크업 추가에 실패했습니다.' });
      }
    },
  );

  app.patch<{ Querystring: TokenQuery; Params: { markupId: string }; Body: UpdateMarkupBody }>(
    '/api/view/markup/:markupId',
    { schema: { querystring: tokenQuerySchema, params: markupParamsSchema, body: updateMarkupBodySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canEdit) {
        await recordAudit({ type: 'markup_move_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '마크업 이동 권한이 없습니다.' });
      }
      try {
        const updated = await updateMarkup(view.meta.id, request.params.markupId, request.body);
        if (!updated) {
          return reply.code(404).send({ success: false, error: '마크업을 찾을 수 없습니다.' });
        }
        await recordAudit({
          type: 'markup_moved',
          docId: view.meta.id,
          viewer: view.claims.viewer,
          detail: updated.id,
        });
        return { success: true, data: updated };
      } catch (error) {
        if (error instanceof MarkupValidationError) {
          return reply.code(400).send({ success: false, error: error.message });
        }
        request.log.error(error, 'markup move failed');
        return reply.code(500).send({ success: false, error: '마크업 이동에 실패했습니다.' });
      }
    },
  );

  /**
   * 클릭 지점 근처의 기존 텍스트에서 글자 크기·굵기·베이스라인을 추정한다 — "텍스트 추가" 도구가
   * 새 텍스트를 주변 폰트와 어울리게 자동으로 놓기 위해 쓴다(src/textStyle.ts).
   */
  app.get<{ Params: { page: number }; Querystring: TextStyleQuery }>(
    '/api/view/pages/:page/text-style',
    { schema: textStyleSchema },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if ((view.meta.viewKind ?? 'pages') !== 'pages') {
        return reply.code(400).send({ success: false, error: '이 문서 형식은 지원하지 않습니다.' });
      }
      try {
        const pages = await loadPdfText(view.meta.id);
        const pageText = pages[request.params.page - 1];
        if (!pageText) {
          return reply.code(404).send({ success: false, error: '페이지를 찾을 수 없습니다.' });
        }
        const style = detectNearbyTextStyle(pageText, request.query.x, request.query.y);
        return { success: true, data: style };
      } catch (error) {
        request.log.error(error, 'text style detection failed');
        return reply.code(500).send({ success: false, error: '폰트 스타일 감지에 실패했습니다.' });
      }
    },
  );

  app.delete<{ Querystring: TokenQuery; Params: { markupId: string } }>(
    '/api/view/markup/:markupId',
    { schema: { querystring: tokenQuerySchema, params: markupParamsSchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canEdit) {
        await recordAudit({ type: 'markup_delete_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '마크업 삭제 권한이 없습니다.' });
      }
      const deleted = await deleteMarkup(view.meta.id, request.params.markupId);
      if (!deleted) {
        return reply.code(404).send({ success: false, error: '마크업을 찾을 수 없습니다.' });
      }
      await recordAudit({
        type: 'markup_deleted',
        docId: view.meta.id,
        viewer: view.claims.viewer,
        detail: request.params.markupId,
      });
      return { success: true, data: { id: request.params.markupId } };
    },
  );

  app.get<{ Querystring: TokenWithMarkupQuery }>(
    '/api/view/print',
    { schema: { querystring: tokenWithMarkupQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canPrint) {
        await recordAudit({ type: 'print_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '출력 권한이 없습니다.' });
      }
      let pdfSource: Buffer;
      try {
        pdfSource = await ensurePdf(view.meta);
      } catch (error) {
        request.log.error(error, 'print pdf generation failed');
        return reply.code(500).send({ success: false, error: '출력용 PDF 생성에 실패했습니다.' });
      }
      let pdfBytes: Buffer | Uint8Array = pdfSource;
      if (request.query.includeMarkup === '1') {
        const items = await listMarkup(view.meta.id);
        if (items.length > 0) {
          pdfBytes = await burnMarkupIntoPdf(Buffer.from(pdfBytes), items);
        }
      }
      const watermark = `${view.claims.viewer}  ${new Date().toISOString().slice(0, 10)}`;
      const finalBytes = await buildPrintPdf(Buffer.from(pdfBytes), watermark);
      await recordAudit({ type: 'printed', docId: view.meta.id, viewer: view.claims.viewer });
      return reply
        .type('application/pdf')
        .header('content-disposition', 'inline; filename="print.pdf"')
        .send(Buffer.from(finalBytes));
    },
  );

  // PDF 편집기(editor-pdf.html)가 pdf.js로 그릴 원본. 바이트를 브라우저로 보내는 것이라 PDF 편집기용 토큰(pdfEditor —
  // 내려받기 권한이 있을 때만 붙는다, documents.ts editorUrlFor)이어야 한다. 방이 열려 있으면 아직 저장 전인 현재 상태를 준다.
  app.get<{ Querystring: { token?: string } }>('/api/view/pdf-source', async (request, reply) => {
    const view = await resolveView(request.query.token, reply);
    if (!view) return reply;
    if (view.meta.ext !== 'pdf') return reply.code(400).send({ success: false, error: 'PDF 문서가 아닙니다.' });
    if (view.claims.perms.pdfEditor !== true || !view.claims.perms.canDownload) {
      return reply.code(403).send({ success: false, error: '이 문서는 PDF 편집기로 열 수 없습니다.' });
    }
    // 편집기가 다루는 주석은 뺀 PDF — 그 주석은 편집 방의 맵에서 그린다(src/pdf/pdfModel.ts). 주석만 바뀌는 동안에는 저장된
    // 원본에서 빼도 결과가 같으므로 저장본 기준으로 만들고, 같은 바이트면 다시 계산하지 않는다.
    // 편집 방이 열려 있으면 방을 열 때의 원본(쪽 목록 열쇠의 기준) — 지문(x-pdf-base)으로 편집기가 방과 맞는지 확인한다.
    // 암호 파일은 비밀번호를 맞힌 토큰에게만(423이면 편집기가 비밀번호 창을 띄운다).
    const meta = view.meta;
    try {
      const { bytes, base } = await editorSourceOf(meta.id, () => store.readOriginal(meta.id, meta.ext), tokenKey(request.query.token!));
      return reply
        .type('application/pdf')
        .header('cache-control', 'no-store')
        .header('x-pdf-base', base)
        .send(await editorBaseOf(meta.id, bytes).then((b) => Buffer.from(b.buffer, b.byteOffset, b.byteLength))); // 복사 없이
    } catch (error) {
      if (error instanceof PdfRoomError) return reply.code(error.status).send({ success: false, error: error.message, locked: error.status === 423 });
      throw error;
    }
  });

  // PDF 편집기 6단계 — 암호 파일 풀기·문서 보호(암호 걸기·풀기)·압축·문서 정보
  const pdfRoute = (perm: (p: ViewClaims['perms']) => boolean, deny: string) => async (token: string | undefined, reply: FastifyReply) => {
    const view = await resolveView(token, reply);
    if (!view) return null;
    if (view.meta.ext !== 'pdf' || view.claims.perms.pdfEditor !== true || !perm(view.claims.perms)) {
      await reply.code(403).send({ success: false, error: deny });
      return null;
    }
    return view;
  };
  const pdfFail = (reply: FastifyReply, error: unknown) => {
    if (error instanceof PdfRoomError) return reply.code(error.status).send({ success: false, error: error.message, locked: error.status === 423 });
    throw error;
  };
  const viewOnly = pdfRoute(() => true, '이 문서는 PDF 편집기로 열 수 없습니다.');
  const editOnly = pdfRoute((p) => p.canEdit, '이 문서를 고칠 권한이 없습니다.');

  app.post<{ Querystring: { token?: string }; Body: { password?: unknown } }>('/api/view/pdf-unlock', { config: { rateLimit: HEAVY_RATE_LIMIT } }, async (request, reply) => {
    const view = await viewOnly(request.query.token, reply);
    if (!view) return reply;
    const password = typeof request.body?.password === 'string' ? request.body.password.slice(0, 256) : '';
    try {
      await unlockPdf(view.meta.id, password, tokenKey(request.query.token!));
      await recordAudit({ type: 'pdf_unlocked', docId: view.meta.id, viewer: view.claims.viewer });
      return reply.send({ success: true });
    } catch (error) {
      return pdfFail(reply, error);
    }
  });

  app.post<{ Querystring: { token?: string }; Body: ProtectRequest }>('/api/view/pdf-protect', async (request, reply) => {
    const view = await editOnly(request.query.token, reply);
    if (!view) return reply;
    const body = request.body ?? ({} as ProtectRequest);
    if (body.mode !== 'set' && body.mode !== 'remove') return reply.code(400).send({ success: false, error: '요청이 올바르지 않습니다.' });
    try {
      await protectPdf(view.meta.id, body, tokenKey(request.query.token!), view.claims.viewer);
      return reply.send({ success: true });
    } catch (error) {
      return pdfFail(reply, error);
    }
  });

  app.post<{ Querystring: { token?: string }; Body: { level?: unknown; dryRun?: unknown } }>('/api/view/pdf-compress', async (request, reply) => {
    const dryRun = request.body?.dryRun === true;
    const view = await (dryRun ? viewOnly : editOnly)(request.query.token, reply);
    if (!view) return reply;
    const level = request.body?.level ?? null;
    if (level !== null && !isCompressLevel(level)) return reply.code(400).send({ success: false, error: '압축 단계가 올바르지 않습니다.' });
    try {
      const data = await compressRoomPdf(view.meta.id, level, dryRun, tokenKey(request.query.token!), view.claims.viewer);
      return reply.send({ success: true, data });
    } catch (error) {
      return pdfFail(reply, error);
    }
  });

  // 7단계 — 변환(PDF → 워드·PPT·엑셀·텍스트), 그림 내보내기용 지금 상태의 PDF, 다른 파일 → PDF
  const downloadOnly = pdfRoute((p) => p.canDownload, '저장(다운로드) 권한이 없습니다.');
  const attachment = (reply: FastifyReply, bytes: Uint8Array, fileName: string, mime: string) =>
    reply.type(mime).header('cache-control', 'no-store').header('content-disposition', `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`).send(Buffer.from(bytes));

  app.post<{ Querystring: { token?: string }; Body: { to?: unknown } }>('/api/view/pdf-convert', { config: { rateLimit: HEAVY_RATE_LIMIT } }, async (request, reply) => {
    const view = await downloadOnly(request.query.token, reply);
    if (!view) return reply;
    const to = request.body?.to;
    if (!isConvertTarget(to)) return reply.code(400).send({ success: false, error: '바꿀 형식이 올바르지 않습니다.' });
    const meta = view.meta;
    try {
      const plain = await currentPlainBytes(meta.id, () => store.readOriginal(meta.id, meta.ext), tokenKey(request.query.token!));
      const out = await convertPdf(plain, to, meta.name);
      await recordAudit({ type: 'downloaded', docId: meta.id, viewer: view.claims.viewer });
      return attachment(reply, out.bytes, out.fileName, out.mime);
    } catch (error) {
      if (error instanceof ConversionError) return reply.code(422).send({ success: false, error: error.message });
      return pdfFail(reply, error);
    }
  });

  app.get<{ Querystring: { token?: string } }>('/api/view/pdf-current', async (request, reply) => {
    const view = await downloadOnly(request.query.token, reply);
    if (!view) return reply;
    const meta = view.meta;
    try {
      const plain = await currentPlainBytes(meta.id, () => store.readOriginal(meta.id, meta.ext), tokenKey(request.query.token!));
      return reply.type('application/pdf').header('cache-control', 'no-store').send(Buffer.from(plain));
    } catch (error) {
      return pdfFail(reply, error);
    }
  });

  app.post<{ Querystring: { token?: string; name?: string } }>('/api/view/file-to-pdf', { bodyLimit: config.maxUploadBytes, onRequest: tokenBeforeBody, config: { rateLimit: HEAVY_RATE_LIMIT } }, async (request, reply) => {
    const view = await downloadOnly(request.query.token, reply);
    if (!view) return reply;
    if (!Buffer.isBuffer(request.body) || request.body.length === 0) return reply.code(415).send({ success: false, error: '파일을 보내 주세요.' });
    const name = (request.query.name ?? 'file').slice(0, 200);
    try {
      const pdf = await fileToPdf(new Uint8Array(request.body), name);
      return attachment(reply, pdf, `${name.replace(/\.[^.]+$/, '') || 'document'}.pdf`, 'application/pdf');
    } catch (error) {
      if (error instanceof FileToPdfError) return reply.code(error.status).send({ success: false, error: error.message });
      throw error;
    }
  });

  // 9단계 — 인증서 전자서명. 본문: { p12(base64), password, reason?, location?, pageKey?, rect? } — 인증서·비밀번호는 기록하지 않는다
  const signOnly = pdfRoute((p) => p.canEdit && p.canDownload, '이 문서에 서명할 권한이 없습니다.');
  app.post<{ Querystring: { token?: string }; Body: Record<string, unknown> }>('/api/view/pdf-sign', { config: { rateLimit: HEAVY_RATE_LIMIT } }, async (request, reply) => {
    const view = await signOnly(request.query.token, reply);
    if (!view) return reply;
    const body = request.body ?? {};
    const p12b64 = typeof body.p12 === 'string' ? body.p12 : '';
    const password = typeof body.password === 'string' ? body.password.slice(0, 256) : '';
    if (!p12b64 || p12b64.length > 200_000 || !/^[A-Za-z0-9+/=]+$/.test(p12b64)) return reply.code(400).send({ success: false, error: '인증서 파일(.pfx·.p12)을 보내 주세요.' });
    const rect = Array.isArray(body.rect) && body.rect.length === 4 && body.rect.every((v) => typeof v === 'number' && Number.isFinite(v)) ? (body.rect as [number, number, number, number]) : undefined;
    const meta = view.meta;
    try {
      const out = await signRoomPdf(meta.id, () => store.readOriginal(meta.id, meta.ext), tokenKey(request.query.token!), {
        p12: new Uint8Array(Buffer.from(p12b64, 'base64')),
        password,
        reason: typeof body.reason === 'string' ? body.reason : undefined,
        location: typeof body.location === 'string' ? body.location : undefined,
        pageKey: typeof body.pageKey === 'string' ? body.pageKey : undefined,
        rect,
      });
      await recordAudit({ type: 'pdf_signed', docId: meta.id, viewer: view.claims.viewer });
      return attachment(reply, out.bytes, `${meta.name.replace(/\.pdf$/i, '') || 'document'}-서명.pdf`, 'application/pdf');
    } catch (error) {
      return pdfFail(reply, error);
    }
  });

  app.get<{ Querystring: { token?: string } }>('/api/view/pdf-info', async (request, reply) => {
    const view = await viewOnly(request.query.token, reply);
    if (!view) return reply;
    const meta = view.meta;
    try {
      const data = await pdfInfoOf(meta.id, () => store.readOriginal(meta.id, meta.ext), tokenKey(request.query.token!));
      return reply.send({ success: true, data });
    } catch (error) {
      return pdfFail(reply, error);
    }
  });

  // PDF 편집기 — 다른 PDF 넣기(파일에서 넣기·파일 합치기). 본문은 PDF 바이트(application/octet-stream).
  app.post<{ Querystring: { token?: string; name?: string } }>('/api/view/pdf-part', { bodyLimit: config.maxUploadBytes, onRequest: tokenBeforeBody }, async (request, reply) => {
    const view = await resolveView(request.query.token, reply);
    if (!view) return reply;
    const perms = view.claims.perms;
    if (view.meta.ext !== 'pdf' || perms.pdfEditor !== true || !perms.canEdit || !perms.canDownload) {
      return reply.code(403).send({ success: false, error: '이 문서에는 파일을 넣을 수 없습니다.' });
    }
    if (!Buffer.isBuffer(request.body)) return reply.code(415).send({ success: false, error: 'PDF 파일을 보내 주세요.' });
    try {
      const data = await addPdfPart(view.meta.id, new Uint8Array(request.body), config.maxUploadBytes, tokenKey(request.query.token!), request.query.name ?? 'file.pdf');
      return reply.send({ success: true, data });
    } catch (error) {
      if (error instanceof PdfRoomError) return reply.code(error.status).send({ success: false, error: error.message });
      throw error;
    }
  });

  app.get<{ Querystring: { token?: string }; Params: { partId: string } }>('/api/view/pdf-part/:partId', async (request, reply) => {
    const view = await resolveView(request.query.token, reply);
    if (!view) return reply;
    if (view.meta.ext !== 'pdf' || view.claims.perms.pdfEditor !== true || !view.claims.perms.canDownload) {
      return reply.code(403).send({ success: false, error: '이 문서는 PDF 편집기로 열 수 없습니다.' });
    }
    let bytes: Uint8Array | null = null;
    try {
      bytes = /^[A-Za-z0-9_-]{1,40}$/.test(request.params.partId) ? await readPdfPart(view.meta.id, request.params.partId, tokenKey(request.query.token!)) : null;
    } catch (error) {
      if (error instanceof PdfRoomError) return reply.code(error.status).send({ success: false, error: error.message });
      throw error;
    }
    if (!bytes) return reply.code(404).send({ success: false, error: '넣은 파일을 찾을 수 없습니다.' });
    return reply.type('application/pdf').header('cache-control', 'no-store').send(Buffer.from(bytes));
  });

  // PDF 편집기 — 고른 쪽 추출(한 묶음 → pdf) / 나누기(여러 묶음 → zip). 문서는 그대로.
  app.post<{ Querystring: { token?: string }; Body: { groups?: unknown } }>('/api/view/pdf-extract', async (request, reply) => {
    const view = await resolveView(request.query.token, reply);
    if (!view) return reply;
    if (view.meta.ext !== 'pdf' || view.claims.perms.pdfEditor !== true || !view.claims.perms.canDownload) {
      return reply.code(403).send({ success: false, error: '저장(다운로드) 권한이 없습니다.' });
    }
    const groups = request.body?.groups;
    const valid =
      Array.isArray(groups) &&
      groups.length > 0 &&
      groups.length <= 2000 &&
      groups.every((g) => Array.isArray(g) && g.length <= 10000 && g.every((k) => typeof k === 'string' && k.length <= 64));
    if (!valid) return reply.code(400).send({ success: false, error: '고른 쪽이 올바르지 않습니다.' });
    const meta = view.meta;
    try {
      const out = await extractPdfGroups(meta.id, groups as string[][], meta.name, () => store.readOriginal(meta.id, meta.ext), tokenKey(request.query.token!));
      await recordAudit({ type: 'downloaded', docId: meta.id, viewer: view.claims.viewer });
      return reply
        .type(out.type === 'pdf' ? 'application/pdf' : 'application/zip')
        .header('content-disposition', `attachment; filename*=UTF-8''${encodeURIComponent(out.fileName)}`)
        .send(Buffer.from(out.bytes));
    } catch (error) {
      if (error instanceof PdfRoomError) return reply.code(error.status).send({ success: false, error: error.message });
      throw error;
    }
  });

  app.get<{ Querystring: TokenWithMarkupQuery }>(
    '/api/view/download',
    { schema: { querystring: tokenWithMarkupQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canDownload) {
        await recordAudit({ type: 'download_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '저장(다운로드) 권한이 없습니다.' });
      }

      if (request.query.includeMarkup === '1') {
        const items = await listMarkup(view.meta.id);
        if (items.length > 0) {
          let pdfSource: Buffer;
          try {
            pdfSource = await ensurePdf(view.meta);
          } catch (error) {
            request.log.error(error, 'markup download pdf generation failed');
            return reply.code(500).send({ success: false, error: '마크업 포함 PDF 생성에 실패했습니다.' });
          }
          const merged = await burnMarkupIntoPdf(pdfSource, items);
          await recordAudit({
            type: 'downloaded',
            docId: view.meta.id,
            viewer: view.claims.viewer,
            detail: 'with markup',
          });
          const baseName = view.meta.name.replace(/\.[^./]+$/, '');
          const encodedName = encodeURIComponent(`${baseName}-markup.pdf`);
          return reply
            .type('application/pdf')
            .header('content-disposition', `attachment; filename*=UTF-8''${encodedName}`)
            .send(Buffer.from(merged));
        }
      }

      // 향기랩 옛 형식(.doc·.xls·.ppt): 원래 형식으로(설계 2026-11-12-hyanggilab-always-edit §2.4)
      if (view.meta.legacy) {
        // 고친 내용을 다시 바꿀 때만(캐시에 없을 때) LibreOffice를 띄우므로 그때만 무거운 요청 제한을 센다
        let waitSeconds = 0;
        const out = await legacyDownload(view.meta, undefined, async () => {
          const limit = await legacyConvertLimit(request);
          // isAllowed는 허용 목록일 때만 true — 센 요청은 늘 false이고, 넘었는지는 isExceeded
          if (!limit.isAllowed && limit.isExceeded) {
            waitSeconds = limit.ttlInSeconds;
            throw new Error('rate limited');
          }
        }).catch((error: unknown) => {
          if (waitSeconds) return null;
          throw error;
        });
        if (!out) return reply.code(429).send({ success: false, error: `요청이 너무 많습니다. ${waitSeconds}초 후 다시 시도하세요.` });
        await recordAudit({ type: 'downloaded', docId: view.meta.id, viewer: view.claims.viewer, detail: `legacy ${out.how}` });
        return reply
          .type('application/octet-stream')
          .header('content-disposition', `attachment; filename*=UTF-8''${encodeURIComponent(out.name)}`)
          .send(out.bytes);
      }

      // 지금 이 문서를 편집 중인 동시편집 세션이 떠 있으면(=유휴 자동저장 8초 debounce가 아직 안
      // 지났을 수 있음) 디스크의 원본이 최신이 아닐 수 있다 — 그 경우 메모리의 Y.Doc에서 바로 최신
      // 바이트를 만든다. 세션이 없으면(이미 디스크에 반영 완료) 원본을 그대로 쓴다.
      const data = (await getCurrentBytes(view.meta.id)) ?? (await store.readOriginal(view.meta.id, view.meta.ext));
      await recordAudit({ type: 'downloaded', docId: view.meta.id, viewer: view.claims.viewer });
      const encodedName = encodeURIComponent(view.meta.name);
      return reply
        .type('application/octet-stream')
        .header('content-disposition', `attachment; filename*=UTF-8''${encodedName}`)
        .send(data);
    },
  );

  // 한글(hwp/hwpx) 전용 편집기(public/editor-hwp.html, @rhwp/editor iframe)가 문서를 로드할 때
  // 쓴다. canDownload가 아니라 유효한 열람 토큰만 요구한다 — docx 에디터가 canEdit=false 토큰도
  // "읽기 전용 모드"로 받아 원본을 로드하는 것과 같은 성격(759214c). 실제 쓰기 차단은
  // /api/view/hwp-save의 canEdit 검사가 담당한다.
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/hwp-source',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveEditorDataView(request.query.token, reply);
      if (!view) return reply;
      if (view.meta.ext !== 'hwp' && view.meta.ext !== 'hwpx') {
        return reply.code(400).send({ success: false, error: 'hwp/hwpx 문서가 아닙니다.' });
      }
      const data = await store.readOriginal(view.meta.id, view.meta.ext);
      return reply.type('application/octet-stream').send(data);
    },
  );

  // rhwp 에디터(iframe)가 브라우저에서 편집 후 exportHwp()/exportHwpx()로 내보낸 바이트를
  // 그대로 받아 원본을 덮어쓴다. docx/xlsx/pptx의 Yjs 자동저장과 달리 여기는 단일 사용자
  // 편집 모델이라(rhwp가 실시간 동시편집 CRDT를 지원하지 않음) 클라이언트가 명시적으로
  // 호출하는 "저장" 한 번이 곧 전체 문서 교체다.
  app.post<{ Querystring: TokenQuery }>(
    '/api/view/hwp-save',
    { schema: { querystring: tokenQuerySchema }, bodyLimit: config.maxUploadBytes, onRequest: tokenBeforeBody },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (view.meta.ext !== 'hwp' && view.meta.ext !== 'hwpx') {
        return reply.code(400).send({ success: false, error: 'hwp/hwpx 문서가 아닙니다.' });
      }
      if (!view.claims.perms.canEdit) {
        return reply.code(403).send({ success: false, error: '편집 권한이 없습니다.' });
      }
      const data = request.body as Buffer;
      if (!data || data.length === 0) {
        return reply.code(400).send({ success: false, error: '저장할 문서 내용이 없습니다.' });
      }
      try {
        assertZipNotTooLarge(data, 'hwp-save');
      } catch (error) {
        if (error instanceof ZipTooLargeError) return reply.code(413).send({ success: false, error: '압축을 풀면 너무 큰 파일입니다.' });
        throw error;
      }
      await store.saveOriginal(view.meta.id, view.meta.ext, data);

      let pageCount = view.meta.pageCount;
      const converter = findConverter(view.meta.ext);
      if (converter) {
        try {
          const converted = await convertToPdf({ ...view.meta }, converter);
          pageCount = converted.pageCount;
          await store.writeMeta({ ...view.meta, pageCount, status: 'ready', viewKind: 'pages', progress: 100 });
        } catch (error) {
          request.log.error(error, 'hwp 저장 후 PDF 재변환 실패');
        }
      }
      invalidateRenderCache(view.meta.id);
      invalidateViewCaches(view.meta.id);
      await recordAudit({ type: 'edit_saved', docId: view.meta.id, viewer: view.claims.viewer });
      return { success: true, data: { pageCount } };
    },
  );
}
