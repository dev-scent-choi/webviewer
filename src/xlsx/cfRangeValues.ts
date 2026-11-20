/**
 * 편집 화면 조건부 서식용 "범위 값 개수 표" — 편집 화면은 보고 있는 페이지(500행 단위)의 칸만 받으므로
 * COUNTIF($A:$A,$A1)>1·중복 값·상위 10개·색조처럼 범위 전체를 보아야 하는 규칙을 스스로 계산할 수 없다.
 * 서버(편집 방의 lazyGrid — 편집 내용이 이미 반영된 지금 상태)가 범위마다 "값 → 개수" 표를 만들어 준다.
 * 값이 같은 칸은 한 줄이라 중복이 많은 열일수록 작다. 편집 방의 리비전이 같으면 앞서 만든 표를 그대로 쓴다.
 */
import { gridCellValue, yieldToEventLoop } from '../spreadsheet.js';
import { getRowCells, type LazyGridStore } from './lazyGridStore.js';

export interface CfRangeRect {
  readonly r0: number;
  readonly c0: number;
  readonly r1: number;
  readonly c1: number;
}

/** 범위 하나의 값 개수 표 — values: [값 글자, 개수][](빈 칸 제외), cells: 범위 칸 수(빈 칸 포함) */
export type CfRangeValues = CfRangeRect & ({ readonly cells: number; readonly values: ReadonlyArray<readonly [string, number]> } | { readonly skipped: true });

/** 한 범위에서 읽는 칸 수 상한 — 넘으면 그 범위는 건너뛴다(편집 화면이 "일부 생략"을 알린다) */
export const CF_RANGE_CELL_CAP = 2_000_000;
/** 한 번에 물을 수 있는 범위 수 */
export const CF_RANGE_MAX_RANGES = 64;
const ROW_BATCH = 500;

export function isCfRangeRequest(body: unknown): body is { sheet: number; ranges: CfRangeRect[] } {
  if (!body || typeof body !== 'object') return false;
  const { sheet, ranges } = body as { sheet?: unknown; ranges?: unknown };
  if (!Number.isInteger(sheet) || (sheet as number) < 0) return false;
  if (!Array.isArray(ranges) || ranges.length === 0 || ranges.length > CF_RANGE_MAX_RANGES) return false;
  return ranges.every(
    (r) =>
      r &&
      typeof r === 'object' &&
      ['r0', 'c0', 'r1', 'c1'].every((k) => Number.isInteger((r as Record<string, unknown>)[k]) && ((r as Record<string, number>)[k] ?? -1) >= 0) &&
      (r as CfRangeRect).r1 >= (r as CfRangeRect).r0 &&
      (r as CfRangeRect).c1 >= (r as CfRangeRect).c0,
  );
}

/** 리비전마다 만든 표 캐시(편집 방 하나에 하나) */
export interface CfRangeCache {
  rev: number;
  readonly byKey: Map<string, CfRangeValues>;
  /** 반영 중인 칸 편집(웹소켓으로 막 들어온 것) — 세기 전에 기다린다 */
  applying: Promise<void>;
}

/**
 * 범위들의 값 개수 표를 만든다. computedAt: 수식 칸의 지금 계산값(편집 방의 계산값 맵, 없으면 undefined).
 * 범위는 시트 크기로 자른다. 겹치는 행은 한 번만 읽는다.
 */
export async function cfRangeValues(
  store: LazyGridStore,
  sheet: number,
  ranges: readonly CfRangeRect[],
  computedAt: (row: number, col: number) => string | undefined,
  cache: CfRangeCache,
): Promise<CfRangeValues[]> {
  const meta = store.sheetMeta[sheet];
  const rows = meta?.rowCount ?? 0;
  const cols = meta?.colCount ?? 0;
  const keyOf = (r: CfRangeRect) => `${sheet}:${r.r0}:${r.c0}:${r.r1}:${r.c1}`;
  if (cache.byKey.size > 256) cache.byKey.clear(); // 리비전 하나에 범위를 마구 물어도 메모리가 늘지 않게
  const todo: Array<{ rect: CfRangeRect; r1: number; c1: number; counts: Map<string, number>; nonEmpty: number }> = [];
  for (const rect of ranges) {
    if (cache.byKey.has(keyOf(rect))) continue;
    const r1 = Math.min(rect.r1, rows - 1);
    const c1 = Math.min(rect.c1, cols - 1);
    const cells = Math.max(0, r1 - rect.r0 + 1) * Math.max(0, c1 - rect.c0 + 1);
    if (cells > CF_RANGE_CELL_CAP) cache.byKey.set(keyOf(rect), { ...pick(rect), skipped: true });
    else todo.push({ rect, r1, c1, counts: new Map(), nonEmpty: 0 });
  }
  if (todo.length > 0) {
    const rowFrom = Math.min(...todo.map((t) => t.rect.r0));
    const rowTo = Math.max(...todo.map((t) => t.r1));
    const colTo = Math.max(...todo.map((t) => t.c1)) + 1;
    // 행을 묶음으로 한꺼번에 읽는다(같은 디스크 조각을 한 번에 읽게 — viewerGrid.ts와 같은 방식). 묶음마다 이벤트 루프에 양보한다
    for (let start = rowFrom; start <= rowTo; start += ROW_BATCH) {
      const end = Math.min(rowTo, start + ROW_BATCH - 1);
      const lines = await Promise.all(Array.from({ length: end - start + 1 }, (_, i) => getRowCells(store, sheet, start + i, colTo)));
      lines.forEach((line, i) => {
        const r = start + i;
        for (const t of todo) {
          if (r < t.rect.r0 || r > t.r1) continue;
          for (let c = t.rect.c0; c <= t.c1; c += 1) {
            const cell = line[c];
            if (!cell) continue;
            let v = gridCellValue(cell);
            if (v.startsWith('=')) v = computedAt(r, c) ?? '';
            if (v === '') continue;
            t.counts.set(v, (t.counts.get(v) ?? 0) + 1);
          }
        }
      });
      await yieldToEventLoop();
    }
    for (const t of todo) {
      const cells = Math.max(0, t.r1 - t.rect.r0 + 1) * Math.max(0, t.c1 - t.rect.c0 + 1);
      cache.byKey.set(keyOf(t.rect), { ...pick(t.rect), cells, values: [...t.counts] });
    }
  }
  return ranges.map((r) => cache.byKey.get(keyOf(r))!);
}

const pick = (r: CfRangeRect): CfRangeRect => ({ r0: r.r0, c0: r.c0, r1: r.r1, c1: r.c1 });
