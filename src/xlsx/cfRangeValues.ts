/**
 * 편집 화면 조건부 서식용 서버 답(/api/view/xlsx-cf-ranges). 편집 화면은 보고 있는 페이지(500행 단위)의 칸만 받으므로
 * COUNTIF($A:$A,$A1)>1·중복 값·상위 10개·색조처럼 범위 전체를 보아야 하는 규칙을 스스로 계산할 수 없다.
 *
 * - 범위는 서버가 방의 조건부 서식 규칙에서 뽑는다(cfRuleRanges + 통계 규칙의 범위). 편집 화면이 묻는 범위가 그 안에 없으면
 *   답하지 않는다(skipped) — 아무 범위나 세게 할 수 없다.
 * - 범위마다 "값 글자 → 개수" 표를 방에 들고 있고, 칸 편집마다 이전 값 −1·새 값 +1 한다(증분). 행/열 삽입·삭제, 규칙·시트 크기
 *   변경 때만 다시 센다. 수식 칸을 고치면(계산값을 아직 모른다) 그 범위만 다시 센다.
 * - 답은 판정에 쓰는 것만: 범위 집계(개수·합·최소·최대)와 편집 화면이 물은 COUNTIF 조건의 개수, 통계 규칙은 문턱값·평균·백분위,
 *   중복 규칙은 2개 이상인 값만. 범위 값 표 전체를 보내지 않는다.
 * - 상한: 방마다 세는 칸 CF_ROOM_CELL_CAP, 서로 다른 값 CF_ROOM_DISTINCT_CAP(넘는 범위는 skipped → 편집 화면 "일부 생략"),
 *   요청 하나의 조건 수·통계 수·중복 값 수. 세기는 방마다 한 번에 하나(같은 요청은 같은 계산을 기다린다), 서버 전체에서 동시에
 *   CF_GLOBAL_BUILDS개까지. 세는 동안 CF_YIELD_MS마다 양보한다.
 */
import {
  cfRangeAnswer,
  cfRangeTableAdjust,
  cfRangeTableFromRaw,
  cfRuleRanges,
  compactStats,
  computeStatsFromCounts,
  needsStats,
  type CfRangeAgg,
  type CfRangeTable,
  type CfRect,
  type CfStats,
  type CfVal,
  type CondFormatEntry,
} from '../collab/condFormatShared.js';
import { gridCellValue, yieldToEventLoop } from '../spreadsheet.js';
import { getRowCells, type LazyGridStore } from './lazyGridStore.js';

/** 방 하나가 세는 칸 수 상한(모든 범위 합) */
export const CF_ROOM_CELL_CAP = 2_000_000;
/** 방 하나가 들고 있는 서로 다른 값 수 상한(모든 범위 합) */
export const CF_ROOM_DISTINCT_CAP = 200_000;
/** 요청 하나 — 범위 수, COUNTIF 조건 수(모든 범위 합), 통계 규칙 수, 중복 값 수(모든 통계 합) */
export const CF_MAX_RANGES = 64;
export const CF_MAX_CRITS = 2_000;
export const CF_MAX_STATS = 64;
export const CF_MAX_DUP_KEYS = 50_000;
/** 서버 전체에서 동시에 세는 방 수 */
export const CF_GLOBAL_BUILDS = 2;
const CF_YIELD_MS = 15;
const ROW_BATCH = 500;

interface RectTally {
  readonly rect: CfRect;
  readonly cells: number;
  /** 값 글자 → 개수(빈 칸 제외). null = 상한을 넘어 세지 않음 */
  raw: Map<string, number> | null;
  /** 수식 칸이 바뀌어 다시 세야 함 */
  dirty: boolean;
  table?: CfRangeTable;
}
interface SheetTally {
  readonly entries: readonly CondFormatEntry[];
  readonly rows: number;
  readonly cols: number;
  readonly rects: Map<string, RectTally>;
  readonly stats: Map<string, CfStats | null>;
}
/** 방(Room.cfRanges)마다 */
export interface CfRoomTally {
  rev: number;
  readonly sheets: Map<number, SheetTally>;
  /** 방마다 세기는 한 번에 하나 — 이 체인 끝에 붙는다 */
  chain: Promise<unknown>;
  /** 반영 중인 칸 편집(웹소켓으로 막 들어온 것) — 답하기 전에 기다린다 */
  applying: Promise<void>;
}
export const newCfRoomTally = (): CfRoomTally => ({ rev: 0, sheets: new Map(), chain: Promise.resolve(), applying: Promise.resolve() });

const keyOf = (r: CfRect) => `${r.r0}:${r.c0}:${r.r1}:${r.c1}`;

/** 칸 하나가 바뀌었다(값 글자 before → after) — 그 칸이 든 범위의 개수 표를 −1/+1 */
export function cfTallyApply(t: CfRoomTally, sheet: number, row: number, col: number, before: string, after: string): void {
  t.rev += 1;
  const s = t.sheets.get(sheet);
  if (!s || before === after) return;
  let changed = false;
  for (const r of s.rects.values()) {
    const { rect } = r;
    if (row < rect.r0 || row > rect.r1 || col < rect.c0 || col > rect.c1 || !r.raw) continue;
    changed = true;
    if (before.startsWith('=') || after.startsWith('=')) {
      r.dirty = true; // 수식 칸의 계산값은 아직 모른다 — 이 범위만 다시 센다
      r.table = undefined;
      continue;
    }
    if (before !== '') {
      const n = (r.raw.get(before) ?? 0) - 1;
      if (n > 0) r.raw.set(before, n);
      else r.raw.delete(before);
      if (r.table) cfRangeTableAdjust(r.table, before, -1);
    }
    if (after !== '') {
      r.raw.set(after, (r.raw.get(after) ?? 0) + 1);
      if (r.table) cfRangeTableAdjust(r.table, after, 1);
    }
  }
  if (changed) s.stats.clear();
  if (changed && [...s.rects.values()].reduce((a, r) => a + (r.raw?.size ?? 0), 0) > CF_ROOM_DISTINCT_CAP) t.sheets.delete(sheet);
}

/** 행/열 삽입·삭제 — 그 시트는 다음에 물을 때 다시 센다 */
export function cfTallyDropSheet(t: CfRoomTally, sheet: number): void {
  t.rev += 1;
  t.sheets.delete(sheet);
}

let globalActive = 0;
const globalWaiters: Array<() => void> = [];
async function withGlobalSlot<T>(work: () => Promise<T>): Promise<T> {
  if (globalActive >= CF_GLOBAL_BUILDS) await new Promise<void>((resolve) => globalWaiters.push(resolve));
  globalActive += 1;
  try {
    return await work();
  } finally {
    globalActive -= 1;
    globalWaiters.shift()?.();
  }
}

/** 범위들을 센다(행을 묶음으로 읽고, 범위 합 칸 수·서로 다른 값 상한을 지킨다) */
async function countRects(
  store: LazyGridStore,
  sheet: number,
  rects: RectTally[],
  computedAt: (row: number, col: number) => string | undefined,
  distinctLeft: number,
): Promise<void> {
  if (rects.length === 0) return;
  for (const r of rects) r.raw = new Map();
  const rowFrom = Math.min(...rects.map((r) => r.rect.r0));
  const rowTo = Math.max(...rects.map((r) => r.rect.r1));
  const colTo = Math.max(...rects.map((r) => r.rect.c1)) + 1;
  let distinct = 0;
  let lastYield = performance.now();
  for (let start = rowFrom; start <= rowTo; start += ROW_BATCH) {
    const end = Math.min(rowTo, start + ROW_BATCH - 1);
    const lines = await Promise.all(Array.from({ length: end - start + 1 }, (_, i) => getRowCells(store, sheet, start + i, colTo)));
    for (let i = 0; i < lines.length; i += 1) {
      const row = start + i;
      const line = lines[i];
      for (const r of rects) {
        const raw = r.raw;
        if (!raw || row < r.rect.r0 || row > r.rect.r1) continue;
        for (let c = r.rect.c0; c <= r.rect.c1; c += 1) {
          const cell = line[c];
          if (!cell) continue;
          let v = gridCellValue(cell);
          if (v.startsWith('=')) v = computedAt(row, c) ?? '';
          if (v === '') continue;
          const n = raw.get(v);
          if (n === undefined) {
            distinct += 1;
            if (distinct > distinctLeft) {
              distinct -= raw.size;
              r.raw = null; // 서로 다른 값이 너무 많다 — 이 범위는 답하지 않는다
              break;
            }
          }
          raw.set(v, (n ?? 0) + 1);
        }
      }
      if (performance.now() - lastYield > CF_YIELD_MS) {
        await yieldToEventLoop();
        lastYield = performance.now();
      }
    }
  }
  for (const r of rects) r.dirty = false;
}

/** 시트의 범위 표를 (필요하면 세어) 돌려준다 — 방마다 한 번에 하나, 서버 전체 동시 상한 */
function ensureSheetTally(
  t: CfRoomTally,
  store: LazyGridStore,
  sheet: number,
  entries: readonly CondFormatEntry[],
  computedAt: (row: number, col: number) => string | undefined,
): Promise<SheetTally> {
  const run = t.chain.then(async () => {
    const meta = store.sheetMeta[sheet];
    const rows = meta?.rowCount ?? 0;
    const cols = meta?.colCount ?? 0;
    let s = t.sheets.get(sheet);
    if (!s || s.entries !== entries || s.rows !== rows || s.cols !== cols) {
      // 규칙이 쓰는 범위 + 통계 규칙(중복·상위·색조 …)의 범위
      const rects = new Map<string, RectTally>();
      const add = (rect: CfRect) => {
        const clamped = { r0: rect.r0, c0: rect.c0, r1: Math.min(rect.r1, rows - 1), c1: Math.min(rect.c1, cols - 1) };
        if (clamped.r1 < clamped.r0 || clamped.c1 < clamped.c0 || rects.size >= 256) return;
        const cells = (clamped.r1 - clamped.r0 + 1) * (clamped.c1 - clamped.c0 + 1);
        rects.set(keyOf(clamped), { rect: clamped, cells, raw: null, dirty: true });
      };
      for (const rect of cfRuleRanges(entries, rows, cols)) add(rect);
      for (const e of entries) if (needsStats(e.rule)) e.ranges.forEach(add);
      s = { entries, rows, cols, rects, stats: new Map() };
      t.sheets.set(sheet, s);
    }
    const todo = [...s.rects.values()].filter((r) => r.dirty);
    if (todo.length > 0) {
      // 칸 수 상한 — 넘는 범위는 세지 않는다(작은 범위부터)
      let cellsLeft = CF_ROOM_CELL_CAP - [...s.rects.values()].filter((r) => !r.dirty && r.raw).reduce((a, r) => a + r.cells, 0);
      const counted: RectTally[] = [];
      for (const r of todo.sort((a, b) => a.cells - b.cells)) {
        if (r.cells <= cellsLeft) {
          counted.push(r);
          cellsLeft -= r.cells;
        } else {
          r.raw = null;
          r.dirty = false;
        }
      }
      const distinctLeft = CF_ROOM_DISTINCT_CAP - [...s.rects.values()].filter((r) => !r.dirty && r.raw).reduce((a, r) => a + (r.raw?.size ?? 0), 0);
      await withGlobalSlot(() => countRects(store, sheet, counted, computedAt, distinctLeft));
      s.stats.clear();
    }
    return s;
  });
  t.chain = run.catch(() => undefined);
  return run;
}

export interface CfRemoteRequest {
  readonly sheet: number;
  readonly ranges: ReadonlyArray<CfRect & { readonly crit?: readonly CfVal[] }>;
  readonly stats?: readonly string[];
}
export type CfRemoteRangeAnswer = CfRect & ({ readonly cells: number; readonly agg: CfRangeAgg; readonly crit: Array<[string, number]> } | { readonly skipped: true });
export interface CfRemoteAnswer {
  readonly rev: number;
  readonly ranges: CfRemoteRangeAnswer[];
  /** 규칙 id → 압축 통계(중복은 [값, 개수][]로 — JSON) 또는 skipped */
  readonly stats: Record<string, (Omit<CfStats, 'counts'> & { counts?: Array<[string, number]> }) | 'skipped'>;
}

const isCrit = (v: unknown): v is CfVal => v === null || typeof v === 'boolean' || (typeof v === 'number' && Number.isFinite(v)) || (typeof v === 'string' && v.length <= 300);
const isIndex = (v: unknown): v is number => Number.isInteger(v) && (v as number) >= 0 && (v as number) <= 1_048_576;

/** 신뢰할 수 없는 요청 본문 검증 */
export function parseCfRemoteRequest(body: unknown): CfRemoteRequest | null {
  if (!body || typeof body !== 'object') return null;
  const { sheet, ranges, stats } = body as Record<string, unknown>;
  if (!isIndex(sheet) || !Array.isArray(ranges) || ranges.length > CF_MAX_RANGES) return null;
  let crits = 0;
  const out: Array<CfRect & { crit?: CfVal[] }> = [];
  for (const r of ranges) {
    if (!r || typeof r !== 'object') return null;
    const { r0, c0, r1, c1, crit } = r as Record<string, unknown>;
    if (![r0, c0, r1, c1].every(isIndex) || (r1 as number) < (r0 as number) || (c1 as number) < (c0 as number)) return null;
    if (crit !== undefined && (!Array.isArray(crit) || !crit.every(isCrit))) return null;
    crits += Array.isArray(crit) ? crit.length : 0;
    out.push({ r0: r0 as number, c0: c0 as number, r1: r1 as number, c1: c1 as number, ...(Array.isArray(crit) ? { crit: crit as CfVal[] } : {}) });
  }
  if (crits > CF_MAX_CRITS) return null;
  if (stats !== undefined && (!Array.isArray(stats) || stats.length > CF_MAX_STATS || !stats.every((id) => typeof id === 'string' && id.length <= 100))) return null;
  return { sheet: sheet as number, ranges: out, ...(Array.isArray(stats) ? { stats: stats as string[] } : {}) };
}

/** 편집 화면의 물음에 답한다(세어 둔 표가 없거나 낡았으면 먼저 센다) */
export async function answerCfRemote(
  t: CfRoomTally,
  store: LazyGridStore,
  req: CfRemoteRequest,
  entries: readonly CondFormatEntry[],
  computedAt: (row: number, col: number) => string | undefined,
): Promise<CfRemoteAnswer> {
  const s = await ensureSheetTally(t, store, req.sheet, entries, computedAt);
  const ranges: CfRemoteRangeAnswer[] = req.ranges.map((q) => {
    const r = s.rects.get(keyOf(q));
    const rect = { r0: q.r0, c0: q.c0, r1: q.r1, c1: q.c1 };
    if (!r?.raw) return { ...rect, skipped: true as const };
    r.table ??= cfRangeTableFromRaw(r.cells, r.raw);
    return { ...rect, cells: r.cells, ...cfRangeAnswer(r.table, q.crit ?? []) };
  });
  const stats: CfRemoteAnswer['stats'] = {};
  let dupKeys = 0;
  for (const id of req.stats ?? []) {
    const e = s.entries.find((x) => x.id === id);
    if (!e || !needsStats(e.rule)) {
      stats[id] = 'skipped';
      continue;
    }
    let st = s.stats.get(id);
    if (st === undefined) {
      const parts = e.ranges.map((rect) => s.rects.get(keyOf({ r0: rect.r0, c0: rect.c0, r1: Math.min(rect.r1, s.rows - 1), c1: Math.min(rect.c1, s.cols - 1) })));
      if (parts.some((p) => !p?.raw)) st = null;
      else {
        const raw = new Map<string, number>();
        for (const p of parts) for (const [v, n] of p!.raw!) raw.set(v, (raw.get(v) ?? 0) + n);
        st = compactStats(e.rule, computeStatsFromCounts(e.rule, raw));
      }
      s.stats.set(id, st);
    }
    dupKeys += st?.counts?.size ?? 0;
    if (!st || dupKeys > CF_MAX_DUP_KEYS) {
      stats[id] = 'skipped';
      continue;
    }
    const { counts, ...rest } = st;
    stats[id] = { ...rest, ...(counts ? { counts: [...counts] } : {}) };
  }
  return { rev: t.rev, ranges, stats };
}
