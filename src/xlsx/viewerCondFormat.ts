/**
 * 보기 화면(읽기 전용 그리드)에 조건부 서식을 미리 입힌다 — 보기 화면은 칸 값이 바뀌지 않으므로, 편집 화면과 같은 판정
 * (collab/condFormatShared.ts paintCell)을 그리드 파일을 만들 때 한 번 돌려 결과를 칸 서식(GridStyle)에 굽는다.
 * 편집 화면은 이 결과를 쓰지 않는다(편집 방은 원본을 따로 읽고, 규칙은 따로 받아 그릴 때 판정한다).
 *
 * 비용 다스리기(서버의 모든 사용자가 같은 프로세스를 쓴다 — 저장할 때마다, TODAY 규칙은 날마다 다시 돈다):
 * - 시간 예산(CF_BAKE_BUDGET_MS) — 그리드 하나에 이만큼만 굽고 나머지는 굽지 않는다. 그리드에 "일부 생략"(cfSkipped)을 적어
 *   보기 화면이 알린다.
 * - 시간 기준 양보 — 굽는 동안 CF_YIELD_MS마다 이벤트 루프에 양보한다(행 수 기준이 아니다).
 * - 규칙 색인은 행 구간 훑기(cfRowSweeper) — 행마다 배열을 만들지 않는다.
 * - 범위 함수(COUNTIF($A:$A,$A2) 등)는 범위 값 표를 한 번 만들어 쓴다(condFormatShared.ts CfEnv.tables).
 * - 범위 통계(색조·막대·상위·중복 …)는 범위가 같은 규칙끼리 값을 한 번만 모은다.
 * - 수식 엔진이 칸마다 큰 범위를 읽어야 하는 규칙은 미리 어림해 예산을 넘으면 건너뛴다.
 * TODAY()·NOW()가 든 규칙이 있으면 구운 날짜(cfDay)를 그리드에 적어, 날짜가 바뀌면 보기 화면이 그리드를 다시 만든다.
 */
import { strFromU8, unzip } from 'fflate';
import {
  cfIsVolatile,
  cfRuleTooCostly,
  cfRowSweeper,
  computeStats,
  needsStats,
  paintCell,
  type CfEnv,
  type CfPaint,
  type CfRangeTable,
  type CfStats,
  type CondFormatEntry,
} from '../collab/condFormatShared.js';
import { gridCellValue, parseThemeColors, yieldToEventLoop, type GridCell, type GridDocument, type GridStyle } from '../spreadsheet.js';
import { createCfFormulaFallback } from './cfFormulaEngine.js';
import { parseCondFormats, parseDxfs } from './condFormatParse.js';
import { sheetPathsByName } from './drawingPassthrough.js';
import { WILDCARD_WORK_BUDGET, runWildcardStep, withWildcardBudget, type WildcardBudget } from '../collab/wildcard.js';

/** 그리드 하나에 조건부 서식을 굽는 시간 예산(ms) */
export const CF_BAKE_BUDGET_MS = 2_000;
/** 이 간격(ms)마다 이벤트 루프에 양보한다 — 한 번에 이보다 오래 막지 않는다 */
export const CF_YIELD_MS = 15;

/** 현지(서버) 날짜 YYYY-MM-DD — TODAY() 규칙을 구운 날. 운영 PC는 KST라 한국 날짜와 같다 */
export function localDayKey(date = new Date()): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

/** 굽기 시간 예산 — 경과 시간으로 양보하고, 다 쓰면 skipped */
export class CfBakeBudget {
  private readonly deadline: number;
  private lastYield = performance.now();
  skipped = false;
  /** 와일드카드 작업 예산 — 굽기 하나에 하나(재검토 C1: 긴 칸 × 거의 맞는 조건이 칸 하나 칠하는 데 몇 초씩 막았다) */
  readonly wildcard: WildcardBudget = { left: WILDCARD_WORK_BUDGET };
  constructor(budgetMs = CF_BAKE_BUDGET_MS) {
    this.deadline = performance.now() + budgetMs;
  }
  get expired(): boolean {
    if (performance.now() > this.deadline) this.skipped = true;
    return this.skipped;
  }
  /** 마지막 양보에서 CF_YIELD_MS가 지났으면 양보 */
  async tick(): Promise<void> {
    if (performance.now() - this.lastYield < CF_YIELD_MS) return;
    await yieldToEventLoop();
    this.lastYield = performance.now();
  }
}

/** 색조 색을 채널마다 4단계로 묶는다 — 칸마다 다른 색이 수천 개의 서식으로 불어나지 않게(눈으로는 구별이 안 된다) */
function quantize(hex: string): string {
  return `#${[1, 3, 5].map((i) => Math.min(255, Math.round(parseInt(hex.slice(i, i + 2), 16) / 4) * 4).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

/** 시트 이름 → 그 시트의 조건부 서식(우선순위 순) */
async function condFormatsBySheet(xlsx: Uint8Array): Promise<Map<string, CondFormatEntry[]>> {
  // 시트 XML(큰 파일은 수십 MB)은 다른 스레드에서 푼다 — 그동안 이 서버의 다른 요청이 멈추지 않게
  const entries = await new Promise<Record<string, Uint8Array>>((resolve, reject) =>
    unzip(
      xlsx,
      {
        filter: (f) => f.name === 'xl/workbook.xml' || f.name === 'xl/_rels/workbook.xml.rels' || f.name === 'xl/styles.xml' || f.name === 'xl/theme/theme1.xml' || f.name.startsWith('xl/worksheets/sheet'),
      },
      (err, data) => (err ? reject(err) : resolve(data)),
    ),
  );
  const themeColors = parseThemeColors(entries['xl/theme/theme1.xml'] ? strFromU8(entries['xl/theme/theme1.xml']) : undefined);
  const dxfs = parseDxfs(entries['xl/styles.xml'] ? strFromU8(entries['xl/styles.xml']) : undefined, themeColors);
  const out = new Map<string, CondFormatEntry[]>();
  for (const [name, path] of sheetPathsByName(entries)) {
    await yieldToEventLoop();
    // 조건부 서식은 시트 데이터 뒤에 있다 — 큰 시트를 통째로 글자로 바꾸지 않고 뒤쪽만 본다
    const bytes = entries[path];
    if (!bytes) continue;
    const tail = bytes.subarray(Math.max(0, lastSheetDataEnd(bytes)));
    const xml = strFromU8(tail);
    if (!xml.includes('onditionalFormatting')) continue;
    const list = parseCondFormats(xml, dxfs, themeColors).filter((e) => e.rule.t !== 'other');
    if (list.length > 0) out.set(name, list);
  }
  return out;
}

/** 시트 XML 바이트에서 </sheetData>가 끝나는 곳(없으면 0) — 그 뒤(조건부 서식·확장)만 글자로 바꾼다 */
function lastSheetDataEnd(bytes: Uint8Array): number {
  const needle = [60, 47, 115, 104, 101, 101, 116, 68, 97, 116, 97, 62]; // "</sheetData>"
  for (let i = bytes.length - needle.length; i >= 0; i -= 1) {
    let ok = true;
    for (let k = 0; k < needle.length; k += 1) {
      if (bytes[i + k] !== needle[k]) {
        ok = false;
        break;
      }
    }
    if (ok) return i + needle.length;
  }
  return 0;
}

/** 구운 결과를 칸 서식에 얹은 서식 */
function paintedStyle(base: GridStyle, paint: CfPaint): GridStyle {
  const next: { -readonly [K in keyof GridStyle]?: GridStyle[K] } = { ...base };
  const bg = paint.style?.bg ?? (paint.scale ? quantize(paint.scale) : undefined);
  if (bg) next.bg = bg;
  if (paint.style?.col) next.col = paint.style.col;
  if (paint.style?.b) next.b = 1;
  if (paint.bar) next.cb = [paint.bar.color, Math.round(paint.bar.ratio * 50) * 2];
  if (paint.icon) next.ci = paint.icon.hide ? [paint.icon.ch, paint.icon.color, 1] : [paint.icon.ch, paint.icon.color];
  return next;
}

/** 서식 목록에 새 서식을 더하는 도구(같은 서식은 한 번만) */
export function styleRegistry(styles: GridStyle[]): { styleOf(s: number | undefined): GridStyle; register(style: GridStyle): number } {
  const index = new Map(styles.map((s, i) => [JSON.stringify(s), i]));
  return {
    styleOf: (s) => (s === undefined ? {} : styles[s]),
    register: (style) => {
      const key = JSON.stringify(style);
      let i = index.get(key);
      if (i === undefined) {
        i = styles.push(style) - 1;
        index.set(key, i);
      }
      return i;
    },
  };
}

/** 굽기에 쓰는 행 읽기 — 작은 파일은 메모리의 행, 큰 파일(스트리밍)은 행 묶음을 읽어 넘긴다 */
export type CfRowSource = (onRow: (r: number, line: ReadonlyArray<GridCell | 0>) => void, budget: CfBakeBudget) => Promise<void>;

/**
 * 시트 하나의 조건부 서식을 굽는다. readRows로 행을 두 번 읽는다(1차 통계, 2차 칠하기) — onPainted로 칠한 행을 돌려준다.
 * valueAt은 칸 값 글자(작은 파일은 시트 전체, 큰 파일은 지금 읽은 행 묶음 안).
 */
export async function bakeSheetCondFormats(
  entries: readonly CondFormatEntry[],
  extent: { rows: number; cols: number },
  readRows: CfRowSource,
  valueAt: (row: number, col: number) => string,
  registry: ReturnType<typeof styleRegistry>,
  budget: CfBakeBudget,
  onPainted: (r: number, line: ReadonlyArray<GridCell | 0>) => void,
): Promise<void> {
  // 수식 엔진이 칸마다 큰 범위를 읽어야 하는 규칙은 어림해 너무 무거우면 건너뛴다
  const usable = entries.filter((e) => {
    const heavy = cfRuleTooCostly(e, extent.rows, extent.cols);
    if (heavy) budget.skipped = true;
    return !heavy;
  });
  if (usable.length === 0) {
    await readRows((r, line) => onPainted(r, line), budget);
    return;
  }
  // 1차: 통계 — 범위가 같은 규칙은 값을 한 번만 모은다
  const statEntries = usable.filter((e) => needsStats(e.rule));
  const groups = new Map<string, { entries: CondFormatEntry[]; values: string[] }>();
  for (const e of statEntries) {
    const key = JSON.stringify(e.ranges);
    const g = groups.get(key) ?? { entries: [], values: [] };
    g.entries.push(e);
    groups.set(key, g);
  }
  if (groups.size > 0) {
    const sweep = cfRowSweeper(statEntries);
    await readRows((r, line) => {
      const seen = new Set<string>();
      for (const e of sweep(r)) {
        const key = JSON.stringify(e.ranges);
        if (seen.has(key)) continue;
        seen.add(key);
        const values = groups.get(key)!.values;
        for (const rect of e.ranges) {
          if (r < rect.r0 || r > rect.r1) continue;
          for (let c = rect.c0; c <= Math.min(rect.c1, line.length - 1); c += 1) {
            const cell = line[c];
            const v = cell ? gridCellValue(cell) : '';
            if (v !== '') values.push(v);
          }
        }
      }
    }, budget);
  }
  const stats = new Map<string, CfStats>();
  for (const g of groups.values()) {
    const byKind = new Map<string, CfStats>();
    for (const e of g.entries) {
      const kindKey = e.rule.t === 'top' ? JSON.stringify(e.rule) : e.rule.t === 'avg' || e.rule.t === 'dup' ? e.rule.t : 'num';
      let st = byKind.get(kindKey);
      if (!st) byKind.set(kindKey, (st = computeStats(e.rule, g.values)));
      stats.set(e.id, st);
    }
  }
  // 2차: 칠하기
  const env: CfEnv = { fallback: createCfFormulaFallback(valueAt, () => extent), tables: new Map<string, CfRangeTable>(), extent: () => extent, unsupported: () => (budget.skipped = true) };
  const sweep = cfRowSweeper(usable);
  // 칸(열)마다 걸리는 규칙 — 규칙 목록이 같은 행끼리는 다시 고르지 않는다(범위가 하나인 규칙만이면 열로만 정해진다)
  let lastRowEntries: readonly CondFormatEntry[] | null = null;
  let byCol: Array<readonly CondFormatEntry[]> = [];
  let colsFixed = false;
  await readRows((r, line) => {
    if (budget.expired) {
      onPainted(r, line);
      return;
    }
    const rowEntries = sweep(r);
    if (rowEntries !== lastRowEntries || !colsFixed) {
      lastRowEntries = rowEntries;
      colsFixed = rowEntries.every((e) => e.ranges.length === 1);
      byCol = [];
    }
    let out: Array<GridCell | 0> | null = null;
    if (rowEntries.length > 0) {
      for (let c = 0; c < line.length; c += 1) {
        const cell = line[c];
        if (!cell) continue;
        const here = (byCol[c] ??= rowEntries.filter((e) => e.ranges.some((rect) => r >= rect.r0 && r <= rect.r1 && c >= rect.c0 && c <= rect.c1)));
        if (here.length === 0) continue;
        if (budget.expired) break; // 칸마다도 본다 — 한 행에 무거운 칸이 많아도 시간 예산을 단계 하나 넘게 넘기지 않게
        const paint = runWildcardStep(budget.wildcard, (step) => withWildcardBudget(step, () => paintCell(here, r, c, gridCellValue(cell), stats, valueAt, env)));
        if (!paint) continue;
        out ??= line.slice();
        out[c] = { ...cell, s: registry.register(paintedStyle(registry.styleOf(cell.s), paint)) };
      }
    }
    onPainted(r, out ?? line);
  }, budget);
}

/** 그리드에 파일의 조건부 서식을 입힌 사본(규칙이 없으면 그대로) */
export async function bakeCondFormats(grid: GridDocument, xlsx: Uint8Array): Promise<GridDocument> {
  // 시트 XML 풀기·규칙 읽기가 실패해도 보기 그리드는 만든다 — 조건부 서식만 건너뛰고 알린다
  let bySheet: Map<string, CondFormatEntry[]>;
  try {
    bySheet = await condFormatsBySheet(xlsx);
  } catch (error) {
    console.warn('조건부 서식 읽기 실패 — 보기 그리드는 조건부 서식 없이 만든다:', error instanceof Error ? error.message : error);
    return { ...grid, cfSkipped: 1 };
  }
  if (bySheet.size === 0) return grid;
  const styles = [...grid.styles];
  const registry = styleRegistry(styles);
  const budget = new CfBakeBudget();
  let volatile = false;
  const sheets = [];
  for (const sheet of grid.sheets) {
    const entries = bySheet.get(sheet.name);
    if (!entries) {
      sheets.push(sheet);
      continue;
    }
    volatile ||= cfIsVolatile(entries);
    const valueAt = (r: number, c: number): string => {
      const cell = sheet.rows[r]?.[c];
      return cell ? gridCellValue(cell) : '';
    };
    const rows: Array<ReadonlyArray<GridCell | 0>> = sheet.rows.slice();
    const readRows: CfRowSource = async (onRow, b) => {
      for (let r = 0; r < sheet.rows.length; r += 1) {
        onRow(r, sheet.rows[r]);
        await b.tick();
      }
    };
    await bakeSheetCondFormats(entries, { rows: sheet.rowCount, cols: sheet.colCount }, readRows, valueAt, registry, budget, (r, line) => {
      rows[r] = line;
    });
    sheets.push({ ...sheet, rows });
  }
  if (budget.skipped) console.warn(`조건부 서식 굽기: 시간 예산(${CF_BAKE_BUDGET_MS}ms)·비용 상한 때문에 일부 생략(보기 화면에 알림)`);
  return { sheets, styles, ...(volatile ? { cfDay: localDayKey() } : {}), ...(budget.skipped ? { cfSkipped: 1 as const } : {}) };
}
