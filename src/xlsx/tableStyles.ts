/**
 * 엑셀 표(Excel 표, xl/tables/tableN.xml)의 표 스타일 — 머리 줄 색·줄무늬·요약 행·첫/끝 열 굵게를 칸 서식으로 입힌다.
 * 보기·편집 화면 모두 그리드를 만들 때(spreadsheet.ts parseXlsxToGrid, lazyGridStore.ts 스트리밍 읽기) 같은 함수로 입힌다.
 * 칸에 직접 준 서식(채우기·글자 색·굵게)이 있으면 그것이 이긴다(엑셀과 같다).
 *
 * 기본 표 스타일 60가지(Light1~21, Medium1~28, Dark1~11)는 "어느 색(진한 글꼴색 dk1 또는 강조 1~6)으로, 머리 줄을 채우는지,
 * 줄무늬를 얼마나 옅게 칠하는지, 선을 어디에 긋는지"의 조합이다 — 그 규칙을 계열별로 옮겼다(색은 문서 테마에서).
 * 사용자 정의 표 스타일(styles.xml의 tableStyles)은 읽지 않고 같은 번호의 기본 스타일 모양 대신 Medium2로 본다.
 */
import { strFromU8 } from 'fflate';
import { parseRangeFormula } from './chartModel.js';
import { relationshipsOf, relsPathOf, resolvePartPath } from './drawingPassthrough.js';

export interface SheetTable {
  /** 0부터 센 범위(양 끝 포함) */
  readonly r0: number;
  readonly r1: number;
  readonly c0: number;
  readonly c1: number;
  readonly headerRows: number;
  readonly totalsRows: number;
  readonly style: string;
  readonly rowStripes: boolean;
  readonly colStripes: boolean;
  readonly firstCol: boolean;
  readonly lastCol: boolean;
}

/** 표가 칸에 입히는 모양(칸 서식 GridStyle의 일부) */
export interface TableLook {
  bg?: string;
  col?: string;
  b?: 1;
  bt?: 1 | 2;
  bb?: 1 | 2;
  bl?: 1 | 2;
  br?: 1 | 2;
  bd?: { t?: readonly [string, string?]; b?: readonly [string, string?]; l?: readonly [string, string?]; r?: readonly [string, string?] };
}

const attr = (xml: string, name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(xml)?.[1];

export function parseTableXml(xml: string): SheetTable | null {
  const tag = /<table\b[^>]*>/.exec(xml)?.[0] ?? '';
  const ref = attr(tag, 'ref');
  const range = ref ? parseRangeFormula(ref) : null;
  if (!range) return null;
  const info = /<tableStyleInfo\b[^>]*>/.exec(xml)?.[0] ?? '';
  const on = (name: string, def: boolean) => {
    const v = attr(info, name);
    return v === undefined ? def : v === '1' || v === 'true';
  };
  return {
    r0: range.r1,
    r1: range.r2,
    c0: range.c1,
    c1: range.c2,
    headerRows: Number(attr(tag, 'headerRowCount') ?? 1),
    totalsRows: Number(attr(tag, 'totalsRowCount') ?? 0),
    style: attr(info, 'name') ?? (info ? '' : 'TableStyleMedium2'),
    rowStripes: on('showRowStripes', true),
    colStripes: on('showColumnStripes', false),
    firstCol: on('showFirstColumn', false),
    lastCol: on('showLastColumn', false),
  };
}

/** 시트 파트에 연결된 표들(<tableParts>) */
export function sheetTablesOf(entries: Readonly<Record<string, Uint8Array>>, sheetPath: string, sheetXml?: string): SheetTable[] {
  const xml = sheetXml ?? (entries[sheetPath] ? strFromU8(entries[sheetPath]) : '');
  if (!xml.includes('<tablePart')) return [];
  const relsBytes = entries[relsPathOf(sheetPath)];
  if (!relsBytes) return [];
  const rels = new Map(relationshipsOf(strFromU8(relsBytes)).map((r) => [r.id, r]));
  const out: SheetTable[] = [];
  for (const m of xml.matchAll(/<tablePart\b[^>]*\br:id="([^"]+)"/g)) {
    const rel = rels.get(m[1]);
    const bytes = rel ? entries[resolvePartPath(sheetPath, rel.target)] : undefined;
    const table = bytes ? parseTableXml(strFromU8(bytes)) : null;
    if (table) out.push(table);
  }
  return out;
}

/** '#RRGGBB'을 HSL 명도로 옅게(t>0, 흰색 쪽)·짙게(t<0) — 엑셀 tint와 같은 계산 */
function tint(hex: string, t: number): string {
  const rgb = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(...rgb);
  const min = Math.min(...rgb);
  const l = (max + min) / 2;
  const d = max - min;
  const s = d === 0 ? 0 : l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h = 0;
  if (d !== 0) {
    const [r, g, b] = rgb;
    h = max === r ? ((g - b) / d + (g < b ? 6 : 0)) / 6 : max === g ? ((b - r) / d + 2) / 6 : ((r - g) / d + 4) / 6;
  }
  const nl = t < 0 ? l * (1 + t) : l * (1 - t) + t;
  const hue = (p: number, q: number, x: number) => {
    const tt = x < 0 ? x + 1 : x > 1 ? x - 1 : x;
    return tt < 1 / 6 ? p + (q - p) * 6 * tt : tt < 1 / 2 ? q : tt < 2 / 3 ? p + (q - p) * (2 / 3 - tt) * 6 : p;
  };
  const q = nl < 0.5 ? nl * (1 + s) : nl + s - nl * s;
  const p = 2 * nl - q;
  const out = s === 0 ? [nl, nl, nl] : [hue(p, q, h + 1 / 3), hue(p, q, h), hue(p, q, h - 1 / 3)];
  return `#${out.map((v) => Math.round(Math.min(1, Math.max(0, v)) * 255).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

const WHITE = '#FFFFFF';

/**
 * 표 안 칸 (r, c)의 모양. themeColors: parseThemeColors 결과(0 lt1, 1 dk1, 4~9 강조 1~6).
 * 계열별 규칙(엑셀 기본 표 스타일을 본뜸):
 * - Light 1~7: 머리 줄 굵게·아래 선, 줄무늬 아주 옅은 색, 표 위아래 선
 * - Light 8~14: 머리 줄 색 채우기·흰 굵은 글자, 표 테두리
 * - Light 15~21: 모든 칸 선, 머리 줄 굵게, 줄무늬
 * - Medium 1~7: 머리 줄 색·흰 굵은 글자, 줄무늬 옅은 색, 행 사이 옅은 선
 * - Medium 8~14: 머리 줄 진한 색, 몸통 두 단계 옅은 색 번갈아
 * - Medium 15~21: 머리 줄 검정, 줄무늬, 행 사이 선
 * - Medium 22~28: 몸통 옅은 색 채우기, 줄무늬는 조금 진하게, 칸 사이 흰 선
 * - Dark 1~7·8~11: 진한 몸통, 흰 글자
 */
export function tableCellLook(table: SheetTable, r: number, c: number, themeColors: readonly string[]): TableLook | undefined {
  if (r < table.r0 || r > table.r1 || c < table.c0 || c > table.c1) return undefined;
  const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(table.style) ?? (table.style === '' ? null : ['', 'Medium', '2']);
  if (!m) return undefined;
  const family = m[1];
  const n = Number(m[2]);
  const k = (n - 1) % 7; // 0 = dk1(검정·회색), 1~6 = 강조 1~6
  const group = Math.floor((n - 1) / 7); // 계열 안의 묶음(0,1,2,3)
  const theme = (i: number) => `#${themeColors[i] ?? (i === 1 ? '000000' : '4472C4')}`;
  const A = k === 0 ? '#000000' : theme(3 + k);
  const isHeader = r < table.r0 + table.headerRows;
  const isTotal = table.totalsRows > 0 && r > table.r1 - table.totalsRows;
  const dataIndex = r - table.r0 - table.headerRows;
  const stripe = table.rowStripes && !isHeader && !isTotal && dataIndex % 2 === 0;
  const colStripe = table.colStripes && !isHeader && !isTotal && (c - table.c0) % 2 === 0;
  const bold = isHeader || isTotal || (table.firstCol && c === table.c0) || (table.lastCol && c === table.c1);
  const look: TableLook = bold ? { b: 1 } : {};
  const line = (color: string, weight: 1 | 2 = 1, style = weight === 2 ? 'medium' : 'thin') => [style, color] as const;
  const top = r === table.r0;
  const bottom = r === table.r1;
  if (family === 'Light') {
    if (group === 0) {
      if (stripe || colStripe) look.bg = tint(A, k === 0 ? 0.85 : 0.8);
      if (k !== 0) look.col = tint(A, -0.25);
      if (isHeader) Object.assign(look, { bb: 1, bd: { b: line(A) } });
      if (top) Object.assign(look, { bt: 1, bd: { ...look.bd, t: line(A) } });
      if (bottom) Object.assign(look, { bb: 1, bd: { ...look.bd, b: line(A) } });
    } else if (group === 1) {
      if (isHeader) Object.assign(look, { bg: A, col: WHITE });
      Object.assign(look, { bt: 1, bb: 1, bd: { t: line(A), b: line(A) } });
      if (c === table.c0) Object.assign(look, { bl: 1, bd: { ...look.bd, l: line(A) } });
      if (c === table.c1) Object.assign(look, { br: 1, bd: { ...look.bd, r: line(A) } });
    } else {
      if (stripe || colStripe) look.bg = tint(A, 0.8);
      Object.assign(look, { bt: 1, bb: 1, bl: 1, br: 1, bd: { t: line(A), b: line(A), l: line(A), r: line(A) } });
    }
  } else if (family === 'Medium') {
    if (group === 0) {
      if (isHeader) Object.assign(look, { bg: A, col: WHITE });
      else if (stripe || colStripe) look.bg = tint(A, 0.8);
      if (!isHeader) Object.assign(look, { bt: 1, bb: 1, bd: { t: line(tint(A, 0.4)), b: line(tint(A, 0.4)) } });
    } else if (group === 1) {
      if (isHeader) Object.assign(look, { bg: A, col: WHITE });
      else look.bg = tint(A, stripe || colStripe ? 0.6 : 0.8);
      Object.assign(look, { bb: 1, bd: { b: line(WHITE) } });
    } else if (group === 2) {
      if (isHeader) Object.assign(look, { bg: '#000000', col: WHITE });
      else if (stripe || colStripe) look.bg = tint(A, 0.8);
      Object.assign(look, { bt: 1, bb: 1, bd: { t: line(A), b: line(A) } });
    } else {
      look.bg = isHeader ? tint(A, 0.6) : tint(A, stripe || colStripe ? 0.6 : 0.8);
      Object.assign(look, { bt: 1, bb: 1, bl: 1, br: 1, bd: { t: line(WHITE), b: line(WHITE), l: line(WHITE), r: line(WHITE) } });
    }
  } else {
    // Dark
    if (group === 0) {
      look.bg = isHeader ? '#000000' : tint(A, stripe || colStripe ? -0.5 : -0.25);
      look.col = WHITE;
    } else {
      if (isHeader) Object.assign(look, { bg: tint(A, -0.5), col: WHITE });
      else look.bg = tint(A, stripe || colStripe ? 0.6 : 0.8);
    }
  }
  if (isTotal) Object.assign(look, { bt: 2, bd: { ...look.bd, t: ['double', family === 'Dark' ? WHITE : A] } });
  return look;
}

/** 칸 서식 위에 표 모양을 얹는다 — 칸에 직접 준 채우기·글자 색·굵게·테두리가 이긴다 */
export function withTableLook<T extends object>(base: T, look: TableLook): T {
  const b = base as Record<string, unknown>;
  const out: Record<string, unknown> = { ...b };
  if (look.bg && !b.bg) out.bg = look.bg;
  if (look.col && !b.col) out.col = look.col;
  if (look.b && !b.b) out.b = 1;
  const bd: Record<string, unknown> = { ...((b.bd as object) ?? {}) };
  for (const side of ['t', 'r', 'b', 'l'] as const) {
    const key = `b${side}`;
    const w = look[key as 'bt'];
    if (!w || b[key]) continue;
    out[key] = w;
    if (look.bd?.[side]) bd[side] = look.bd[side];
  }
  if (Object.keys(bd).length > 0) out.bd = bd;
  return out as T;
}
