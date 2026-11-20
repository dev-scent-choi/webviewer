/**
 * 보기 화면(읽기 전용 그리드)에 조건부 서식을 미리 입힌다 — 보기 화면은 칸 값이 바뀌지 않으므로, 편집 화면과 같은 판정
 * (collab/condFormatShared.ts paintCell)을 그리드 파일을 만들 때 한 번 돌려 결과를 칸 서식(GridStyle)에 굽는다.
 * 편집 화면은 이 결과를 쓰지 않는다(편집기 시딩은 parseXlsxToGrid 결과를 그대로 쓰고, 규칙은 따로 받아 칸마다 다시 판정한다).
 */
import { strFromU8, unzipSync } from 'fflate';
import { computeStats, needsStats, paintCell, type CfStats, type CondFormatEntry } from '../collab/condFormatShared.js';
import { gridCellValue, parseThemeColors, type GridCell, type GridDocument, type GridSheet, type GridStyle } from '../spreadsheet.js';
import { parseCondFormats, parseDxfs } from './condFormatParse.js';
import { sheetPathsByName } from './drawingPassthrough.js';

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
    if (!xml.includes('<conditionalFormatting')) continue;
    const list = parseCondFormats(xml, dxfs, themeColors).filter((e) => e.rule.t !== 'other');
    if (list.length > 0) out.set(name, list);
  }
  return out;
}

function bakeSheet(sheet: GridSheet, entries: readonly CondFormatEntry[], styleOf: (s: number | undefined) => GridStyle, register: (style: GridStyle) => number): GridSheet {
  const rows = sheet.rows.map((row) => row.slice()) as Array<Array<GridCell | 0>>;
  const valueAt = (r: number, c: number): string => {
    const cell = rows[r]?.[c];
    return cell ? gridCellValue(cell) : '';
  };
  const clip = (r: { r0: number; r1: number; c0: number; c1: number }) => ({
    r0: r.r0,
    r1: Math.min(r.r1, sheet.rowCount - 1),
    c0: r.c0,
    c1: Math.min(r.c1, sheet.colCount - 1),
  });
  const stats = new Map<string, CfStats>();
  for (const e of entries) {
    if (!needsStats(e.rule)) continue;
    const values: string[] = [];
    for (const rect of e.ranges.map(clip)) for (let r = rect.r0; r <= rect.r1; r += 1) for (let c = rect.c0; c <= rect.c1; c += 1) {
      const v = valueAt(r, c);
      if (v !== '') values.push(v);
    }
    stats.set(e.id, computeStats(e.rule, values));
  }
  const seen = new Set<number>();
  for (const rect of entries.flatMap((e) => e.ranges.map(clip))) {
    for (let r = rect.r0; r <= rect.r1; r += 1) {
      for (let c = rect.c0; c <= rect.c1; c += 1) {
        const key = r * 16384 + c;
        if (seen.has(key)) continue;
        seen.add(key);
        const cell = rows[r]?.[c];
        if (!cell) continue;
        const paint = paintCell(entries, r, c, gridCellValue(cell), stats, valueAt);
        if (!paint) continue;
        const base = styleOf(cell.s);
        const next: { -readonly [K in keyof GridStyle]?: GridStyle[K] } = { ...base };
        const bg = paint.style?.bg ?? (paint.scale ? quantize(paint.scale) : undefined);
        if (bg) next.bg = bg;
        if (paint.style?.col) next.col = paint.style.col;
        if (paint.style?.b) next.b = 1;
        if (paint.bar) next.cb = [paint.bar.color, Math.round(paint.bar.ratio * 50) * 2];
        if (paint.icon) next.ci = paint.icon.hide ? [paint.icon.ch, paint.icon.color, 1] : [paint.icon.ch, paint.icon.color];
        rows[r][c] = { ...cell, s: register(next) };
      }
    }
  }
  return { ...sheet, rows };
}

/** 그리드에 파일의 조건부 서식을 입힌 사본(규칙이 없으면 그대로) */
export function bakeCondFormats(grid: GridDocument, xlsx: Uint8Array): GridDocument {
  const bySheet = condFormatsBySheet(xlsx);
  if (bySheet.size === 0) return grid;
  const styles = [...grid.styles];
  const index = new Map(styles.map((s, i) => [JSON.stringify(s), i]));
  const register = (style: GridStyle): number => {
    const key = JSON.stringify(style);
    let i = index.get(key);
    if (i === undefined) {
      i = styles.push(style) - 1;
      index.set(key, i);
    }
    return i;
  };
  const styleOf = (s: number | undefined): GridStyle => (s === undefined ? {} : styles[s]);
  const sheets = grid.sheets.map((sheet) => {
    const entries = bySheet.get(sheet.name);
    return entries ? bakeSheet(sheet, entries, styleOf, register) : sheet;
  });
  return { sheets, styles };
}
