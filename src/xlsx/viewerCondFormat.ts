/**
 * 보기 화면(읽기 전용 그리드)에 조건부 서식을 미리 입힌다 — 보기 화면은 칸 값이 바뀌지 않으므로, 편집 화면과 같은 판정
 * (collab/condFormatShared.ts paintCell)을 그리드 파일을 만들 때 한 번 돌려 결과를 칸 서식(GridStyle)에 굽는다.
 * 편집 화면은 이 결과를 쓰지 않는다(편집기 시딩은 parseXlsxToGrid 결과를 그대로 쓰고, 규칙은 따로 받아 칸마다 다시 판정한다).
 *
 * 비용: 규칙을 행마다 색인해(cfRowIndex) 칸마다 그 행에 걸린 규칙만 보고, 행 묶음마다 이벤트 루프에 양보하고,
 * 일의 양(칸×규칙)에 상한을 둔다 — 넘으면 그 시트의 나머지는 굽지 않고 경고만 남긴다(보기 화면이 조금 덜 칠해질 뿐 서버는 멈추지 않는다).
 * TODAY()·NOW()가 든 규칙이 있으면 구운 날짜(cfDay)를 그리드에 적어, 날짜가 바뀌면 보기 화면이 그리드를 다시 만든다.
 */
import { strFromU8, unzipSync } from 'fflate';
import {
  cfIsVolatile,
  cfRowIndex,
  computeStats,
  needsStats,
  paintCell,
  type CfFormulaFallback,
  type CfPaint,
  type CfStats,
  type CondFormatEntry,
} from '../collab/condFormatShared.js';
import {
  gridCellValue,
  parseThemeColors,
  yieldToEventLoop,
  YIELD_EVERY_N_ROWS,
  type GridCell,
  type GridDocument,
  type GridStyle,
} from '../spreadsheet.js';
import { createCfFormulaFallback } from './cfFormulaEngine.js';
import { parseCondFormats, parseDxfs } from './condFormatParse.js';
import { sheetPathsByName } from './drawingPassthrough.js';

/** 한 그리드 파일에서 조건부 서식 판정(칸 × 그 행의 규칙)을 이만큼까지만 한다 — 넘으면 나머지는 굽지 않는다 */
export const CF_BAKE_WORK_CAP = 20_000_000;

/** 현지 날짜(YYYY-MM-DD) — TODAY() 규칙을 구운 날 */
export function localDayKey(date = new Date()): string {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

/** 색조 색을 채널마다 4단계로 묶는다 — 칸마다 다른 색이 수천 개의 서식으로 불어나지 않게(눈으로는 구별이 안 된다) */
function quantize(hex: string): string {
  return `#${[1, 3, 5].map((i) => Math.min(255, Math.round(parseInt(hex.slice(i, i + 2), 16) / 4) * 4).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

/** 시트 이름 → 그 시트의 조건부 서식(우선순위 순) */
function condFormatsBySheet(xlsx: Uint8Array): Map<string, CondFormatEntry[]> {
  const entries = unzipSync(xlsx, {
    filter: (f) => f.name === 'xl/workbook.xml' || f.name === 'xl/_rels/workbook.xml.rels' || f.name === 'xl/styles.xml' || f.name === 'xl/theme/theme1.xml' || f.name.startsWith('xl/worksheets/sheet'),
  });
  const themeColors = parseThemeColors(entries['xl/theme/theme1.xml'] ? strFromU8(entries['xl/theme/theme1.xml']) : undefined);
  const dxfs = parseDxfs(entries['xl/styles.xml'] ? strFromU8(entries['xl/styles.xml']) : undefined, themeColors);
  const out = new Map<string, CondFormatEntry[]>();
  for (const [name, path] of sheetPathsByName(entries)) {
    const xml = entries[path] ? strFromU8(entries[path]) : '';
    if (!xml.includes('onditionalFormatting')) continue;
    const list = parseCondFormats(xml, dxfs, themeColors).filter((e) => e.rule.t !== 'other');
    if (list.length > 0) out.set(name, list);
  }
  return out;
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

/** 굽기에 쓰는 시트 접근 — 작은 파일은 메모리의 행, 큰 파일(스트리밍)은 지금 읽고 있는 행 묶음 */
export interface CfBakeRows {
  readonly rowCount: number;
  readonly colCount: number;
  row(r: number): ReadonlyArray<GridCell | 0> | undefined;
}

/** 규칙마다 범위 안 값의 통계(중복·상위·평균·색조·막대·아이콘). rowsOf로 필요한 행만 읽는다 */
export async function cfStatsFor(entries: readonly CondFormatEntry[], rows: CfBakeRows): Promise<Map<string, CfStats>> {
  const stats = new Map<string, CfStats>();
  for (const e of entries) {
    if (!needsStats(e.rule)) continue;
    const values: string[] = [];
    let seen = 0;
    for (const rect of e.ranges) {
      for (let r = rect.r0; r <= Math.min(rect.r1, rows.rowCount - 1); r += 1) {
        const line = rows.row(r);
        if (line) for (let c = rect.c0; c <= Math.min(rect.c1, rows.colCount - 1, line.length - 1); c += 1) {
          const cell = line[c];
          if (cell) {
            const v = gridCellValue(cell);
            if (v !== '') values.push(v);
          }
        }
        seen += 1;
        if (seen % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
      }
    }
    stats.set(e.id, computeStats(e.rule, values));
  }
  return stats;
}

/** 한 행의 칸들에 조건부 서식을 입힌 새 행(바뀐 칸이 없으면 같은 배열). work: 남은 일의 양(줄여 간다) */
export function bakeRow(
  r: number,
  line: ReadonlyArray<GridCell | 0>,
  rowEntries: readonly CondFormatEntry[] | undefined,
  stats: ReadonlyMap<string, CfStats>,
  valueAt: (row: number, col: number) => string,
  fallback: CfFormulaFallback | undefined,
  registry: ReturnType<typeof styleRegistry>,
  work: { left: number },
): ReadonlyArray<GridCell | 0> {
  if (!rowEntries || rowEntries.length === 0 || work.left <= 0) return line;
  let out: Array<GridCell | 0> | null = null;
  for (let c = 0; c < line.length; c += 1) {
    const cell = line[c];
    if (!cell) continue;
    const entries = rowEntries.filter((e) => e.ranges.some((rect) => r >= rect.r0 && r <= rect.r1 && c >= rect.c0 && c <= rect.c1));
    if (entries.length === 0) continue;
    work.left -= entries.length;
    const paint = paintCell(entries, r, c, gridCellValue(cell), stats, valueAt, fallback);
    if (!paint) continue;
    out ??= line.slice();
    out[c] = { ...cell, s: registry.register(paintedStyle(registry.styleOf(cell.s), paint)) };
  }
  return out ?? line;
}

/** 그리드에 파일의 조건부 서식을 입힌 사본(규칙이 없으면 그대로) */
export async function bakeCondFormats(grid: GridDocument, xlsx: Uint8Array): Promise<GridDocument> {
  const bySheet = condFormatsBySheet(xlsx);
  if (bySheet.size === 0) return grid;
  const styles = [...grid.styles];
  const registry = styleRegistry(styles);
  const work = { left: CF_BAKE_WORK_CAP };
  let volatile = false;
  const sheets = [];
  for (const sheet of grid.sheets) {
    const entries = bySheet.get(sheet.name);
    if (!entries) {
      sheets.push(sheet);
      continue;
    }
    volatile ||= cfIsVolatile(entries);
    const src: CfBakeRows = { rowCount: sheet.rowCount, colCount: sheet.colCount, row: (r) => sheet.rows[r] };
    const valueAt = (r: number, c: number): string => {
      const cell = sheet.rows[r]?.[c];
      return cell ? gridCellValue(cell) : '';
    };
    const fallback = createCfFormulaFallback(valueAt, () => ({ rows: sheet.rowCount, cols: sheet.colCount }));
    const stats = await cfStatsFor(entries, src);
    const byRow = cfRowIndex(entries, sheet.rowCount);
    const rows: Array<ReadonlyArray<GridCell | 0>> = [];
    for (let r = 0; r < sheet.rows.length; r += 1) {
      rows.push(bakeRow(r, sheet.rows[r], byRow[r], stats, valueAt, fallback, registry, work));
      if ((r + 1) % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
    }
    if (work.left <= 0) console.warn(`조건부 서식 굽기 상한(${CF_BAKE_WORK_CAP})에 닿아 시트 "${sheet.name}" 일부만 칠했다`);
    sheets.push({ ...sheet, rows });
  }
  return { sheets, styles, ...(volatile ? { cfDay: localDayKey() } : {}) };
}
