/**
 * 엑셀 표(Excel 표, xl/tables/tableN.xml)의 표 스타일 — 머리 줄 색·줄무늬·요약 행·첫/끝 열 굵게를 칸 서식으로 입힌다.
 * 보기 화면은 그리드 파일을 만들 때(spreadsheet.ts parseXlsxToGrid — 보기 전용 그리드) 칸 서식에 굽고, 편집 화면은 그릴 때만
 * 덧칠한다(grid-editor.ts — 저장되는 칸 서식에는 절대 넣지 않는다. 값 하나를 고쳐 저장해도 줄무늬 색이 그 칸의 직접 서식이 되지 않게,
 * 정렬해도 줄무늬가 행 위치를 따르게). 칸에 직접 준 서식(채우기·글자 색·굵게)이 있으면 그것이 이긴다(엑셀과 같다).
 * DOM·Node를 쓰지 않는다(브라우저 번들에도 들어간다).
 *
 * 기본 표 스타일 60가지(Light1~21, Medium1~28, Dark1~11)는 "어느 색(진한 글꼴색 dk1 또는 강조 1~6)으로, 머리 줄을 채우는지,
 * 줄무늬를 얼마나 옅게 칠하는지, 선을 어디에 긋는지"의 조합이다 — 그 규칙을 계열별로 옮겼다(색은 문서 테마에서).
 * 사용자 정의 표 스타일(styles.xml의 tableStyles)은 읽지 않고 같은 번호의 기본 스타일 모양 대신 Medium2로 본다.
 */
import { shiftSpan } from '../collab/condFormatShared.js';
import { parseRangeFormula } from './chartModel.js';

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
  /** 표 이름·열 이름(파일 기준 — 구조적 참조를 고칠 때) */
  readonly name?: string;
  readonly columns?: readonly string[];
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

/** 행/열 삽입·삭제(편집 방 structuralOps — 0부터 센 at)를 순서대로 표 범위에 반영한다. 다 지워진 표는 뺀다 */
export function shiftTables<T extends SheetTable & { readonly sheetName: string }>(
  tables: readonly T[],
  ops: ReadonlyArray<{ readonly sheet: number; readonly kind: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols'; readonly at: number; readonly count: number }>,
  sheetNameOf: (sheet: number) => string | undefined,
): T[] {
  let out = [...tables];
  for (const op of ops) {
    const name = sheetNameOf(op.sheet);
    const rows = op.kind === 'insertRows' || op.kind === 'deleteRows';
    const insert = op.kind === 'insertRows' || op.kind === 'insertCols';
    out = out.flatMap((t) => {
      if (t.sheetName !== name) return [t];
      if (rows) {
        const box = shiftTableRows(t, op.at, op.count, insert);
        return box ? [{ ...t, ...box }] : [];
      }
      const span = shiftSpan(t.c0, t.c1, op.at, op.count, insert);
      if (!span) return [];
      return [{ ...t, c0: span[0], c1: span[1], ...(t.columns ? { columns: shiftTableColumns(t.columns, t.c0, t.c1, op.at, op.count, insert) } : {}) }];
    });
  }
  return out;
}

/** 열 삽입·삭제 뒤의 표 열 이름(표 안에 넣으면 "열N", 지우면 뺀다 — 저장 collab.ts shiftWorksheetTables와 같다) */
export function shiftTableColumns(columns: readonly string[], c0: number, c1: number, at: number, count: number, insert: boolean): string[] {
  const out = [...columns];
  if (insert) {
    if (at <= c0 || at > c1) return out;
    const used = new Set(out.map((c) => c.toLowerCase()));
    let k = 1;
    const added = Array.from({ length: count }, () => {
      while (used.has(`열${k}`)) k += 1;
      used.add(`열${k}`);
      return `열${k}`;
    });
    out.splice(at - c0, 0, ...added);
    return out;
  }
  const from = Math.max(at, c0) - c0;
  const to = Math.min(at + count - 1, c1) - c0;
  if (to >= from) out.splice(from, to - from + 1);
  return out;
}

/** 표 위치와 머리 줄·요약 행 수 */
export interface TableRows {
  readonly r0: number;
  readonly r1: number;
  readonly headerRows: number;
  readonly totalsRows: number;
}

/**
 * 행 삽입·삭제를 표 하나에 반영한다(저장 collab.ts와 화면 shiftTables가 같은 규칙). null = 표를 뺀다(칸 값은 남는다).
 * - 표 전체를 지우면 표를 뺀다.
 * - 머리 줄만 지우면(엑셀은 이 삭제를 막는다 — 편집 화면도 막는다) 표를 일반 범위로 바꾼다(다른 참여자·옛 화면이 보낸 연산 대비).
 * - 요약 행을 지우면 요약 행을 끈다(마지막 데이터 행이 요약 행으로 바뀌지 않게).
 * - 데이터 행을 모두 지우는 삭제는 편집 방(room.ts applyGridStructureOp)이 엑셀처럼 "첫 데이터 행은 비우고 나머지만 지우기"로 바꿔
 *   받는다(빈 데이터 행 하나가 남는다 — 저장과 화면이 같다). 그래도 데이터 행이 0개가 되면(옛 기록) 일반 범위로 바꾼다.
 */
export function shiftTableRows(t: TableRows, at: number, count: number, insert: boolean): TableRows | null {
  const span = shiftSpan(t.r0, t.r1, at, count, insert);
  if (!span) return null;
  if (insert) return { ...t, r0: span[0], r1: span[1] };
  const gone = (row: number) => row >= at && row < at + count;
  if (t.headerRows > 0 && gone(t.r0)) return null;
  const totalsRows = t.totalsRows > 0 && gone(t.r1) ? 0 : t.totalsRows;
  // 데이터 행이 하나도 안 남으면 일반 범위로 바꾼다(편집 방은 이런 삭제를 "첫 데이터 행 비우기"로 바꿔 받으므로 여기 오지 않는다)
  if (span[1] - span[0] + 1 - t.headerRows - totalsRows <= 0) return null;
  return { ...t, r0: span[0], r1: span[1], totalsRows };
}

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
    name: xmlText(attr(tag, 'name') ?? ''),
    columns: [...xml.matchAll(/<tableColumn\b[^>]*?\sname="([^"]*)"/g)].map((m) => xmlText(m[1])),
  };
}
const xmlText = (s: string) => s.replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');

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

/** 표가 일반 범위가 됐을 때 남은 자리 — 구조적 참조를 칸 주소로 바꾼다(엑셀 "범위로 변환"과 같다) */
export interface TableAsRange {
  readonly c0: number;
  readonly columns: readonly string[];
  /** 데이터 행(0부터, 양 끝 포함). dataR1 < dataR0이면 데이터 행 없음 */
  readonly dataR0: number;
  readonly dataR1: number;
  readonly totalsRow?: number;
}
/** 표 하나의 열 이름 바뀜(소문자 옛 이름 → 새 이름)과 지운 열(소문자 이름). range: 표가 일반 범위가 됨(null = 표 자리가 다 지워짐 → #REF!) */
export interface TableColumnChanges {
  readonly renames: ReadonlyMap<string, string>;
  readonly deleted: ReadonlySet<string>;
  readonly range?: TableAsRange | null;
}

/** 0부터 센 열 → $B 같은 절대 열 */
function absCol(c: number): string {
  let n = c + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return `$${s}`;
}
/** 일반 범위가 된 표를 가리키는 구조적 참조 하나 → 칸 주소(없는 부분이면 #REF!) */
function structuredRefToAddress(items: readonly string[], range: TableAsRange | null, ownRow: number | undefined): string {
  if (!range) return '#REF!';
  const specials = items.map((x) => x.trim().toLowerCase()).filter((x) => x.startsWith('#'));
  const thisRow = items.some((x) => x.startsWith('@')) || specials.includes('#this row');
  const cols = items.map((x) => x.replace(/^@/, '').trim()).filter((x) => x !== '' && !x.startsWith('#'));
  let r0 = range.dataR0;
  let r1 = range.dataR1;
  if (specials.includes('#headers')) return '#REF!'; // 머리 줄은 없어졌다
  if (specials.includes('#totals')) {
    if (range.totalsRow === undefined) return '#REF!';
    r0 = r1 = range.totalsRow;
  } else if (specials.includes('#all') && range.totalsRow !== undefined) r1 = range.totalsRow;
  const idx = cols.map((c) => range.columns.findIndex((x) => x.toLowerCase() === c.toLowerCase()));
  if (idx.some((i) => i < 0)) return '#REF!';
  const ca = range.c0 + (idx.length > 0 ? Math.min(...idx) : 0);
  const cb = range.c0 + (idx.length > 0 ? Math.max(...idx) : range.columns.length - 1);
  if (thisRow) {
    if (ownRow === undefined) return '#REF!';
    return ca === cb ? `${absCol(ca)}${ownRow + 1}` : `${absCol(ca)}${ownRow + 1}:${absCol(cb)}${ownRow + 1}`;
  }
  if (r1 < r0) return '#REF!';
  const a = `${absCol(ca)}$${r0 + 1}`;
  const b = `${absCol(cb)}$${r1 + 1}`;
  return a === b ? a : `${a}:${b}`;
}

/**
 * 수식의 구조적 참조(표1[값], 표1[@값], 표1[[#Headers],[값]], 표1[[값]:[비고]], 표 안 칸의 [@값])를 표 열 변경에 맞춘다 — 엑셀처럼
 * 이름이 바뀐 열은 새 이름으로, 지운 열을 가리키는 참조는 #REF!로. tables: 소문자 표 이름 → 변경. ownTable: 이 칸이 든 표(이름 없는
 * [@값] 참조용). 따옴표 안 글자는 건드리지 않는다.
 */
export function rewriteStructuredRefs(formula: string, tables: ReadonlyMap<string, TableColumnChanges>, ownTable?: string, ownRow?: number): string {
  let out = '';
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch === '"') {
      const end = formula.indexOf('"', i + 1);
      const stop = end < 0 ? formula.length : end + 1;
      out += formula.slice(i, stop);
      i = stop;
      continue;
    }
    if (ch !== '[') {
      out += ch;
      i += 1;
      continue;
    }
    // 이 '[' 앞의 이름(표 이름) — 없으면 이 칸이 든 표
    const nameMatch = /[A-Za-z_\\À-￿][\w.À-￿]*$/.exec(out);
    const name = nameMatch ? nameMatch[0] : ownTable;
    // 짝 맞는 ']'
    let depth = 0;
    let j = i;
    for (; j < formula.length; j += 1) {
      if (formula[j] === '[') depth += 1;
      else if (formula[j] === ']') {
        depth -= 1;
        if (depth === 0) break;
      }
    }
    const body = formula.slice(i, j + 1);
    const changes = name ? tables.get(name.toLowerCase()) : undefined;
    if (!changes || j >= formula.length) {
      out += body;
      i = j + 1;
      continue;
    }
    const inner = body.slice(1, -1);
    const items = inner.includes('[') ? [...inner.matchAll(/\[([^\]]*)\]/g)].map((m) => m[1]) : [inner];
    const columns = items.map((x) => x.replace(/^@/, '').trim()).filter((x) => x !== '' && !x.startsWith('#'));
    if (changes.range !== undefined) {
      // 표가 일반 범위가 됐다(또는 다 지워졌다) — 칸 주소로(엑셀 "범위로 변환")
      if (nameMatch) out = out.slice(0, out.length - nameMatch[0].length);
      out += structuredRefToAddress(items, changes.range, ownRow);
    } else if (columns.some((c) => changes.deleted.has(c.toLowerCase()))) {
      if (nameMatch) out = out.slice(0, out.length - nameMatch[0].length);
      out += '#REF!';
    } else {
      const renameOne = (x: string) => {
        const at = x.startsWith('@') ? '@' : '';
        const col = x.slice(at.length);
        return at + (changes.renames.get(col.trim().toLowerCase()) ?? col);
      };
      out += inner.includes('[') ? `[${inner.replace(/\[([^\]]*)\]/g, (_m, x: string) => `[${x.startsWith('#') ? x : renameOne(x)}]`)}]` : `[${inner.startsWith('#') ? inner : renameOne(inner)}]`;
    }
    i = j + 1;
  }
  return out;
}
