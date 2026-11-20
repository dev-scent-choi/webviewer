/**
 * 엑셀 조건부 서식 — 편집기(화면 표시·규칙 만들기)와 서버(파일 읽기·행/열 이동·저장)가 함께 쓰는 형식과 판정.
 *
 * 규칙 목록은 시트마다 GRID_COND_FORMATS_MAP(Y.Map, 키 = 시트 번호)에 통째로 들어 있다(병합처럼 윈도잉하지 않는다).
 * 목록 순서가 곧 우선순위다(앞이 높다) — 같은 칸에 여러 규칙이 맞으면 앞 규칙의 색이 이긴다(엑셀과 같다).
 *
 * 파일에서 읽은 규칙은 `p`(파일의 priority)를 갖는다 — 저장할 때 ExcelJS가 읽어 둔 원래 규칙 객체를 이 값으로
 * 다시 찾아 그대로 쓴다(그래서 화면에 못 그리는 종류(`t: 'other'`, 수식·아이콘 집합 등)도 잃지 않는다).
 */

export const GRID_COND_FORMATS_MAP = 'condFormats';

/** 0부터 센 사각형(양 끝 포함) */
export interface CfRect {
  readonly r0: number;
  readonly r1: number;
  readonly c0: number;
  readonly c1: number;
}

/** 강조 서식 — 색은 '#RRGGBB' */
export interface CfStyle {
  readonly bg?: string;
  readonly col?: string;
  readonly b?: 1;
}

export type CfCellOp = 'gt' | 'lt' | 'ge' | 'le' | 'eq' | 'ne' | 'between' | 'notBetween';
export type CfTextOp = 'contains' | 'notContains' | 'begins' | 'ends';

export type CfRule =
  | { readonly t: 'cell'; readonly op: CfCellOp; readonly v1: string; readonly v2?: string; readonly s: CfStyle }
  | { readonly t: 'text'; readonly op: CfTextOp; readonly text: string; readonly s: CfStyle }
  | { readonly t: 'dup'; readonly unique?: 1; readonly s: CfStyle }
  | { readonly t: 'top'; readonly n: number; readonly bottom?: 1; readonly percent?: 1; readonly s: CfStyle }
  | { readonly t: 'avg'; readonly below?: 1; readonly s: CfStyle }
  | { readonly t: 'scale'; readonly colors: readonly string[] }
  | { readonly t: 'bar'; readonly color: string }
  /** 아이콘 집합 — th: 둘째 아이콘부터의 문턱값(첫 cfvo는 엑셀도 보지 않는다), gt: 이상(>=)이 아니라 초과(>) */
  | { readonly t: 'icon'; readonly set: string; readonly th: readonly CfIconThreshold[]; readonly reverse?: 1; readonly hide?: 1 }
  /** 수식 규칙 — 범위 왼쪽 위 칸 기준으로 쓴 수식(상대 참조는 칸마다 옮긴다) */
  | { readonly t: 'expr'; readonly f: string; readonly s: CfStyle }
  | { readonly t: 'other' };

export interface CfIconThreshold {
  readonly type: 'percent' | 'percentile' | 'num';
  readonly v: number;
  readonly gt?: 1;
}

/** 아이콘 집합 모양(낮은 값 → 높은 값) — 엑셀 아이콘을 글자 기호와 색으로 흉내 낸다 */
const RED = '#E5322D';
const YELLOW = '#F2B800';
const GREEN = '#2E9D44';
const GRAY = '#7F7F7F';
const BLUE = '#2F6EBA';
function arrows(n: 3 | 4 | 5, gray = false): Array<{ ch: string; color: string }> {
  const c = (color: string) => (gray ? GRAY : color);
  const marks = n === 3 ? ['↓', '→', '↑'] : n === 4 ? ['↓', '↘', '↗', '↑'] : ['↓', '↘', '→', '↗', '↑'];
  return marks.map((ch, i) => ({ ch, color: c(i === 0 ? RED : i === marks.length - 1 ? GREEN : YELLOW) }));
}
const dots = (colors: readonly string[]) => colors.map((color) => ({ ch: '●', color }));
export const CF_ICON_SETS: Readonly<Record<string, ReadonlyArray<{ readonly ch: string; readonly color: string }>>> = {
  '3Arrows': arrows(3),
  '3ArrowsGray': arrows(3, true),
  '4Arrows': arrows(4),
  '4ArrowsGray': arrows(4, true),
  '5Arrows': arrows(5),
  '5ArrowsGray': arrows(5, true),
  '3TrafficLights1': dots([RED, YELLOW, GREEN]),
  '3TrafficLights2': dots([RED, YELLOW, GREEN]),
  '4TrafficLights': dots(['#3F3F3F', RED, YELLOW, GREEN]),
  '4RedToBlack': dots(['#3F3F3F', GRAY, '#F4A6A6', RED]),
  '3Signs': [{ ch: '◆', color: RED }, { ch: '▲', color: YELLOW }, { ch: '●', color: GREEN }],
  '3Symbols': [{ ch: '✖', color: RED }, { ch: '!', color: YELLOW }, { ch: '✔', color: GREEN }],
  '3Symbols2': [{ ch: '✖', color: RED }, { ch: '!', color: YELLOW }, { ch: '✔', color: GREEN }],
  '3Flags': [{ ch: '⚑', color: RED }, { ch: '⚑', color: YELLOW }, { ch: '⚑', color: GREEN }],
  '3Stars': [{ ch: '☆', color: YELLOW }, { ch: '★', color: '#F7D57A' }, { ch: '★', color: YELLOW }],
  '3Triangles': [{ ch: '▼', color: RED }, { ch: '▬', color: YELLOW }, { ch: '▲', color: GREEN }],
  '4Rating': ['▂', '▄', '▆', '█'].map((ch) => ({ ch, color: BLUE })),
  '5Rating': ['▁', '▂', '▄', '▆', '█'].map((ch) => ({ ch, color: BLUE })),
  '5Quarters': ['○', '◔', '◑', '◕', '●'].map((ch) => ({ ch, color: '#404040' })),
  '5Boxes': ['□', '▫', '◧', '◩', '■'].map((ch) => ({ ch, color: BLUE })),
};

export interface CondFormatEntry {
  readonly id: string;
  readonly ranges: readonly CfRect[];
  readonly rule: CfRule;
  /** 파일에서 읽은 규칙의 priority — 저장 때 원래 규칙을 찾는 열쇠 */
  readonly p?: number;
}

/** 엑셀 "셀 강조 규칙" 대화 상자의 서식 고르기 6가지(엑셀과 같은 색) */
export const CF_PRESETS: ReadonlyArray<{ readonly key: string; readonly label: string; readonly s: CfStyle }> = [
  { key: 'red', label: '진한 빨강 텍스트가 있는 연한 빨강 채우기', s: { bg: '#FFC7CE', col: '#9C0006' } },
  { key: 'yellow', label: '진한 노랑 텍스트가 있는 노랑 채우기', s: { bg: '#FFEB9C', col: '#9C5700' } },
  { key: 'green', label: '진한 녹색 텍스트가 있는 녹색 채우기', s: { bg: '#C6EFCE', col: '#006100' } },
  { key: 'lightRed', label: '연한 빨강 채우기', s: { bg: '#FFC7CE' } },
  { key: 'redText', label: '빨강 텍스트', s: { col: '#9C0006' } },
  { key: 'redBorder', label: '빨강 굵은 텍스트', s: { col: '#FF0000', b: 1 } },
];

export function rectContains(rect: CfRect, row: number, col: number): boolean {
  return row >= rect.r0 && row <= rect.r1 && col >= rect.c0 && col <= rect.c1;
}

export function rectsIntersect(a: CfRect, b: CfRect): boolean {
  return !(a.r0 > b.r1 || a.r1 < b.r0 || a.c0 > b.c1 || a.c1 < b.c0);
}

/** a에서 b를 뺀 나머지(최대 4조각) — "선택한 셀에서 규칙 지우기" */
export function subtractRect(a: CfRect, b: CfRect): CfRect[] {
  if (!rectsIntersect(a, b)) return [a];
  const out: CfRect[] = [];
  if (a.r0 < b.r0) out.push({ r0: a.r0, r1: b.r0 - 1, c0: a.c0, c1: a.c1 });
  if (a.r1 > b.r1) out.push({ r0: b.r1 + 1, r1: a.r1, c0: a.c0, c1: a.c1 });
  const r0 = Math.max(a.r0, b.r0);
  const r1 = Math.min(a.r1, b.r1);
  if (a.c0 < b.c0) out.push({ r0, r1, c0: a.c0, c1: b.c0 - 1 });
  if (a.c1 > b.c1) out.push({ r0, r1, c0: b.c1 + 1, c1: a.c1 });
  return out;
}

/** 한 축의 [start, end](0부터, 양 끝 포함)를 삽입·삭제에 맞춘다 — gridShift.ts의 병합 규칙과 같다. 없어지면 null */
function shiftSpan(start: number, end: number, at: number, count: number, insert: boolean): [number, number] | null {
  if (insert) {
    if (start >= at) return [start + count, end + count];
    if (at <= end) return [start, end + count];
    return [start, end];
  }
  const delEnd = at + count; // 지우는 범위 [at, delEnd)
  const overlap = Math.max(0, Math.min(end + 1, delEnd) - Math.max(start, at));
  const removedBefore = Math.max(0, Math.min(delEnd, start) - at);
  const span = end - start + 1 - overlap;
  if (span <= 0) return null;
  const newStart = start - removedBefore;
  return [newStart, newStart + span - 1];
}

/** 행/열 삽입·삭제에 맞춰 규칙 범위를 옮긴다. 범위가 모두 사라진 규칙은 뺀다. */
export function shiftCondFormats(
  entries: readonly CondFormatEntry[],
  axis: 'row' | 'col',
  at: number,
  count: number,
  insert: boolean,
): CondFormatEntry[] {
  const out: CondFormatEntry[] = [];
  for (const entry of entries) {
    const ranges: CfRect[] = [];
    for (const r of entry.ranges) {
      const span = axis === 'row' ? shiftSpan(r.r0, r.r1, at, count, insert) : shiftSpan(r.c0, r.c1, at, count, insert);
      if (!span) continue;
      ranges.push(axis === 'row' ? { ...r, r0: span[0], r1: span[1] } : { ...r, c0: span[0], c1: span[1] });
    }
    if (ranges.length > 0) out.push({ ...entry, ranges });
  }
  return out;
}

// ---------- A1 표기 ----------
function colLetters(col0: number): string {
  let n = col0 + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
function colIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
export function cellA1(row0: number, col0: number, abs = false): string {
  return abs ? `$${colLetters(col0)}$${row0 + 1}` : `${colLetters(col0)}${row0 + 1}`;
}
export function rectA1(rect: CfRect, abs = false): string {
  const a = cellA1(rect.r0, rect.c0, abs);
  return rect.r0 === rect.r1 && rect.c0 === rect.c1 ? a : `${a}:${cellA1(rect.r1, rect.c1, abs)}`;
}
/** "A1:B3 D5" → 사각형들(엑셀 sqref). 시트 끝까지(A:A, 1:1)도 받는다. */
export function parseSqref(sqref: string, maxRow = 1_048_575, maxCol = 16_383): CfRect[] {
  const out: CfRect[] = [];
  for (const part of sqref.trim().split(/\s+/)) {
    const [a, b = a] = part.replace(/\$/g, '').split(':');
    const pa = /^([A-Z]*)(\d*)$/i.exec(a);
    const pb = /^([A-Z]*)(\d*)$/i.exec(b);
    if (!pa || !pb || (!pa[1] && !pa[2])) continue;
    const r0 = pa[2] ? Number(pa[2]) - 1 : 0;
    const r1 = pb[2] ? Number(pb[2]) - 1 : maxRow;
    const c0 = pa[1] ? colIndex(pa[1]) : 0;
    const c1 = pb[1] ? colIndex(pb[1]) : maxCol;
    out.push({ r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1) });
  }
  return out;
}

// ---------- 판정 ----------
const NUMERIC = /^-?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$/;
export function cfNumber(value: string): number | null {
  const t = value.trim().replace(/,/g, '');
  return t !== '' && NUMERIC.test(t) ? Number(t) : null;
}

/** 규칙의 값 자리("10", "=10", "=\"사과\"", "=$B$1") → 비교할 값. 다른 칸을 가리키면 그 칸 값을 읽는다. */
export function resolveOperand(operand: string, valueAt: (row: number, col: number) => string): string | null {
  let s = operand.trim();
  if (s.startsWith('=')) s = s.slice(1).trim();
  if (s.length >= 2 && s.startsWith('"') && s.endsWith('"')) return s.slice(1, -1).replace(/""/g, '"');
  if (cfNumber(s) !== null) return s;
  const ref = /^\$?([A-Z]{1,3})\$?(\d+)$/i.exec(s);
  if (ref) return valueAt(Number(ref[2]) - 1, colIndex(ref[1]));
  if (s === '' || /[(),:!+*/^&<>-]/.test(s)) return s === '' ? '' : null; // 수식은 계산하지 않는다
  return s; // 따옴표 없는 글자(편집기에서 그냥 적은 값)
}

/** 범위 전체를 보아야 하는 규칙(중복·상위·평균·색조·막대)의 통계 */
export interface CfStats {
  readonly counts?: ReadonlyMap<string, number>;
  readonly threshold?: number;
  readonly mean?: number;
  readonly min?: number;
  readonly max?: number;
  readonly mid?: number;
  /** 아이콘 집합의 백분위 문턱값용 — 오름차순 값 */
  readonly sorted?: readonly number[];
}

export function needsStats(rule: CfRule): boolean {
  return rule.t === 'dup' || rule.t === 'top' || rule.t === 'avg' || rule.t === 'scale' || rule.t === 'bar' || rule.t === 'icon';
}

/** values: 규칙 범위 안의 (불러와 둔) 값들 */
export function computeStats(rule: CfRule, values: readonly string[]): CfStats {
  if (rule.t === 'dup') {
    const counts = new Map<string, number>();
    for (const v of values) {
      const k = v.trim().toLowerCase();
      if (k !== '') counts.set(k, (counts.get(k) ?? 0) + 1);
    }
    return { counts };
  }
  const nums = values.map(cfNumber).filter((n): n is number => n !== null);
  if (nums.length === 0) return {};
  if (rule.t === 'top') {
    const sorted = [...nums].sort((a, b) => (rule.bottom ? a - b : b - a));
    const n = rule.percent ? Math.max(1, Math.floor((sorted.length * rule.n) / 100)) : Math.max(1, rule.n);
    return { threshold: sorted[Math.min(n, sorted.length) - 1] };
  }
  if (rule.t === 'avg') return { mean: nums.reduce((a, b) => a + b, 0) / nums.length };
  const min = Math.min(...nums);
  const max = Math.max(...nums);
  if (rule.t === 'scale') {
    const sorted = [...nums].sort((a, b) => a - b);
    return { min, max, mid: percentileOf(sorted, 50) }; // 50번째 백분위수(엑셀 3색 색조 기본)
  }
  if (rule.t === 'icon') return { min, max, sorted: [...nums].sort((a, b) => a - b) };
  return { min: Math.min(0, min), max: Math.max(0, max) }; // 데이터 막대: 엑셀 "자동" 최소·최대는 0을 포함한다
}

/** 오름차순 값의 p번째 백분위수(엑셀 PERCENTILE.INC와 같은 보간) */
function percentileOf(sorted: readonly number[], p: number): number {
  const pos = ((sorted.length - 1) * Math.max(0, Math.min(100, p))) / 100;
  const lo = Math.floor(pos);
  return sorted[lo] + (sorted[Math.min(lo + 1, sorted.length - 1)] - sorted[lo]) * (pos - lo);
}

function cellRuleMatches(op: CfCellOp, value: string, a: string | null, b: string | null): boolean {
  if (a === null) return false;
  const v = value.trim() === '' ? '0' : value; // 엑셀은 빈 칸을 0으로 본다
  const nv = cfNumber(v);
  const na = cfNumber(a);
  if (op === 'between' || op === 'notBetween') {
    const nb = b === null ? null : cfNumber(b);
    if (nv === null || na === null || nb === null) return op === 'notBetween' && nv === null;
    const inside = nv >= Math.min(na, nb) && nv <= Math.max(na, nb);
    return op === 'between' ? inside : !inside;
  }
  if (nv !== null && na !== null) {
    switch (op) {
      case 'gt': return nv > na;
      case 'lt': return nv < na;
      case 'ge': return nv >= na;
      case 'le': return nv <= na;
      case 'eq': return nv === na;
      case 'ne': return nv !== na;
    }
  }
  const sv = value.trim().toLowerCase();
  const sa = a.trim().toLowerCase();
  if (op === 'eq') return sv === sa;
  if (op === 'ne') return sv !== sa;
  return false;
}

export interface CfPaint {
  style?: CfStyle;
  /** 데이터 막대: 색과 길이(0~1) */
  bar?: { color: string; ratio: number };
  /** 색조: 칸 배경 */
  scale?: string;
  /** 아이콘 집합: 칸 왼쪽에 그릴 기호와 색(hide면 값은 숨긴다) */
  icon?: { ch: string; color: string; hide?: true };
}

function mixHex(a: string, b: string, t: number): string {
  const pa = [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16));
  const pb = [1, 3, 5].map((i) => parseInt(b.slice(i, i + 2), 16));
  return `#${pa.map((x, i) => Math.round(x + (pb[i] - x) * t).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

/**
 * 칸 하나에 걸리는 강조. entries는 우선순위 순(앞이 높다), stats는 entry.id별 통계(needsStats인 것만).
 * 우선순위가 낮은 규칙부터 덮어 써서 높은 규칙의 속성이 이기게 한다.
 */
export function paintCell(
  entries: readonly CondFormatEntry[],
  row: number,
  col: number,
  value: string,
  stats: ReadonlyMap<string, CfStats>,
  valueAt: (row: number, col: number) => string,
): CfPaint | null {
  let paint: CfPaint | null = null;
  for (let i = entries.length - 1; i >= 0; i -= 1) {
    const entry = entries[i];
    if (!entry.ranges.some((r) => rectContains(r, row, col))) continue;
    const rule = entry.rule;
    const st = stats.get(entry.id) ?? {};
    let hit: CfStyle | undefined;
    switch (rule.t) {
      case 'cell': {
        const a = resolveOperand(rule.v1, valueAt);
        const b = rule.v2 === undefined ? null : resolveOperand(rule.v2, valueAt);
        if (cellRuleMatches(rule.op, value, a, b)) hit = rule.s;
        break;
      }
      case 'text': {
        const v = value.toLowerCase();
        const t = rule.text.toLowerCase();
        const ok =
          rule.op === 'contains' ? v.includes(t) : rule.op === 'notContains' ? !v.includes(t) : rule.op === 'begins' ? v.startsWith(t) : v.endsWith(t);
        if (ok) hit = rule.s;
        break;
      }
      case 'dup': {
        const k = value.trim().toLowerCase();
        if (k !== '') {
          const n = st.counts?.get(k) ?? 0;
          if (rule.unique ? n === 1 : n > 1) hit = rule.s;
        }
        break;
      }
      case 'top': {
        const n = cfNumber(value);
        if (n !== null && st.threshold !== undefined && (rule.bottom ? n <= st.threshold : n >= st.threshold)) hit = rule.s;
        break;
      }
      case 'avg': {
        const n = cfNumber(value);
        if (n !== null && st.mean !== undefined && (rule.below ? n < st.mean : n > st.mean)) hit = rule.s;
        break;
      }
      case 'scale': {
        const n = cfNumber(value);
        if (n === null || st.min === undefined || st.max === undefined) break;
        const [c0, c1, c2] = rule.colors;
        let color: string;
        if (st.max === st.min) color = c2 && c1 ? c1 : c0;
        else if (c2) {
          const mid = st.mid ?? (st.min + st.max) / 2;
          color = n <= mid ? mixHex(c0, c1, mid === st.min ? 1 : (n - st.min) / (mid - st.min)) : mixHex(c1, c2, st.max === mid ? 1 : (n - mid) / (st.max - mid));
        } else color = mixHex(c0, c1, (n - st.min) / (st.max - st.min));
        paint = { ...(paint ?? {}), scale: color };
        break;
      }
      case 'bar': {
        const n = cfNumber(value);
        if (n === null || st.min === undefined || st.max === undefined || st.max === st.min) break;
        paint = { ...(paint ?? {}), bar: { color: rule.color, ratio: Math.max(0, Math.min(1, (n - st.min) / (st.max - st.min))) } };
        break;
      }
      case 'icon': {
        const n = cfNumber(value);
        const icons = CF_ICON_SETS[rule.set] ?? CF_ICON_SETS['3TrafficLights1'];
        if (n === null || st.min === undefined || st.max === undefined || !st.sorted) break;
        let level = 0;
        for (let k = Math.min(rule.th.length, icons.length - 1); k >= 1; k -= 1) {
          const th = rule.th[k - 1];
          const at = th.type === 'num' ? th.v : th.type === 'percentile' ? percentileOf(st.sorted, th.v) : st.min + ((st.max - st.min) * th.v) / 100;
          if (th.gt ? n > at : n >= at) {
            level = k;
            break;
          }
        }
        const icon = icons[rule.reverse ? icons.length - 1 - level : level];
        paint = { ...(paint ?? {}), icon: { ...icon, ...(rule.hide ? { hide: true as const } : {}) } };
        break;
      }
      case 'expr': {
        const anchor = entry.ranges[0];
        if (anchor && evalCfFormula(rule.f, row - anchor.r0, col - anchor.c0, row, col, valueAt)) hit = rule.s;
        break;
      }
      default:
        break;
    }
    if (hit) paint = { ...(paint ?? {}), style: { ...(paint?.style ?? {}), ...hit } };
  }
  return paint;
}

// ---------- 수식 규칙(type="expression") 계산 ----------
// 조건부 서식 수식은 짧다($C2>=1, MOD(ROW(),2)=0, AND($D2="완료",$E2>0) …). 편집기 번들에 수식 엔진 전체를 싣지 않고
// 비교·산술·연결·자주 쓰는 함수만 여기서 계산한다. 모르는 함수나 범위·다른 시트 참조가 나오면 규칙이 안 맞는 것으로 본다.
type CfVal = number | string | boolean | null; // null = 빈 칸

function truthy(v: CfVal): boolean {
  return typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : false;
}
function num(v: CfVal): number {
  if (v === null || v === '') return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (typeof v === 'number') return v;
  const n = cfNumber(v);
  if (n === null) throw new Error('#VALUE!');
  return n;
}
function text(v: CfVal): string {
  return v === null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
}
/** 비교 — 빈 칸은 상대가 숫자면 0, 글자면 ""(엑셀과 같다). 종류가 다르면 숫자 < 글자 < 논리값 */
function compare(a: CfVal, b: CfVal): number {
  const x = a === null ? (typeof b === 'string' ? '' : 0) : a;
  const y = b === null ? (typeof a === 'string' ? '' : 0) : b;
  if (typeof x === 'number' && typeof y === 'number') return x - y;
  if (typeof x === typeof y) {
    const sx = text(x).toLowerCase();
    const sy = text(y).toLowerCase();
    return sx < sy ? -1 : sx > sy ? 1 : 0;
  }
  const rank = (v: CfVal) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  return rank(x) - rank(y);
}

const CF_FUNCS: Readonly<Record<string, (args: CfVal[], at: { row: number; col: number }) => CfVal>> = {
  AND: (a) => a.every((v) => truthy(v)),
  OR: (a) => a.some((v) => truthy(v)),
  NOT: (a) => !truthy(a[0] ?? null),
  IF: (a) => (truthy(a[0] ?? null) ? (a[1] ?? true) : (a[2] ?? false)),
  ISBLANK: (a) => a[0] === null,
  ISNUMBER: (a) => typeof a[0] === 'number',
  ISTEXT: (a) => typeof a[0] === 'string',
  LEN: (a) => text(a[0] ?? null).length,
  ABS: (a) => Math.abs(num(a[0] ?? null)),
  INT: (a) => Math.floor(num(a[0] ?? null)),
  ROUND: (a) => {
    const f = 10 ** num(a[1] ?? 0);
    return Math.round(num(a[0] ?? null) * f) / f;
  },
  MOD: (a) => {
    const n = num(a[0] ?? null);
    const d = num(a[1] ?? null);
    if (d === 0) throw new Error('#DIV/0!');
    return n - d * Math.floor(n / d);
  },
  ROW: (_a, at) => at.row + 1,
  COLUMN: (_a, at) => at.col + 1,
  LEFT: (a) => text(a[0] ?? null).slice(0, a.length > 1 ? num(a[1]) : 1),
  RIGHT: (a) => {
    const t = text(a[0] ?? null);
    const n = a.length > 1 ? num(a[1]) : 1;
    return n <= 0 ? '' : t.slice(-n);
  },
  UPPER: (a) => text(a[0] ?? null).toUpperCase(),
  LOWER: (a) => text(a[0] ?? null).toLowerCase(),
  TRIM: (a) => text(a[0] ?? null).trim().replace(/\s+/g, ' '),
  TODAY: () => Math.floor((Date.now() - Date.UTC(1899, 11, 30)) / 86_400_000),
};

const NUM_TOKEN = /^(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?/;
const REF_TOKEN = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(:!])/;
const FUNC_TOKEN = /^([A-Za-z][A-Za-z0-9.]*)\s*\(/;
const BOOL_TOKEN = /^(TRUE|FALSE)(?![\w(])/i;
const CMP_TOKEN = /^(<>|<=|>=|=|<|>)/;

/**
 * 조건부 서식 수식 하나를 칸 (row, col)에 대해 계산해 참/거짓을 낸다. dr/dc: 범위 왼쪽 위에서 이 칸까지 — $ 없는 참조를 이만큼 옮긴다.
 * 계산할 수 없으면(모르는 함수, 범위·다른 시트 참조, 문법 오류) false.
 */
export function evalCfFormula(
  formula: string,
  dr: number,
  dc: number,
  row: number,
  col: number,
  valueAt: (row: number, col: number) => string,
): boolean {
  const src = formula.trim().replace(/^=/, '');
  let i = 0;
  const rest = () => src.slice(i);
  const peek = () => {
    while (src[i] === ' ') i += 1;
    return src[i] ?? '';
  };
  const fail = (): never => {
    throw new Error('cf formula');
  };
  const expect = (ch: string) => {
    if (peek() !== ch) fail();
    i += 1;
  };
  const cellValue = (r: number, c: number): CfVal => {
    if (r < 0 || c < 0) fail();
    const v = valueAt(r, c);
    if (v === '') return null;
    const n = cfNumber(v);
    if (n !== null) return n;
    if (v === 'TRUE' || v === 'FALSE') return v === 'TRUE';
    return v;
  };
  const primary = (): CfVal => {
    const ch = peek();
    if (ch === '(') {
      i += 1;
      const v = comparison();
      expect(')');
      return v;
    }
    if (ch === '"') {
      let out = '';
      i += 1;
      for (;;) {
        if (i >= src.length) fail();
        if (src[i] === '"') {
          if (src[i + 1] !== '"') {
            i += 1;
            return out;
          }
          i += 1;
        }
        out += src[i];
        i += 1;
      }
    }
    const numMatch = NUM_TOKEN.exec(rest());
    if (numMatch) {
      i += numMatch[0].length;
      return Number(numMatch[0]);
    }
    const bool = BOOL_TOKEN.exec(rest());
    if (bool) {
      i += bool[0].length;
      return bool[1].toUpperCase() === 'TRUE';
    }
    const ref = REF_TOKEN.exec(rest());
    if (ref) {
      i += ref[0].length;
      let c = 0;
      for (const l of ref[2].toUpperCase()) c = c * 26 + (l.charCodeAt(0) - 64);
      return cellValue(Number(ref[4]) - 1 + (ref[3] ? 0 : dr), c - 1 + (ref[1] ? 0 : dc));
    }
    const name = FUNC_TOKEN.exec(rest());
    if (name) {
      i += name[0].length;
      const fn = CF_FUNCS[name[1].toUpperCase()] ?? fail();
      const args: CfVal[] = [];
      if (peek() !== ')') {
        for (;;) {
          args.push(comparison());
          if (peek() !== ',') break;
          i += 1;
        }
      }
      expect(')');
      return fn(args, { row, col });
    }
    return fail();
  };
  const unary = (): CfVal => {
    if (peek() === '-') {
      i += 1;
      return -num(unary());
    }
    if (peek() === '+') {
      i += 1;
      return unary();
    }
    let v = primary();
    while (peek() === '%') {
      i += 1;
      v = num(v) / 100;
    }
    return v;
  };
  const power = (): CfVal => {
    let v = unary();
    while (peek() === '^') {
      i += 1;
      v = num(v) ** num(unary());
    }
    return v;
  };
  const product = (): CfVal => {
    let v = power();
    for (let op = peek(); op === '*' || op === '/'; op = peek()) {
      i += 1;
      const r = num(power());
      if (op === '/' && r === 0) fail();
      v = op === '*' ? num(v) * r : num(v) / r;
    }
    return v;
  };
  const sum = (): CfVal => {
    let v = product();
    for (let op = peek(); op === '+' || op === '-'; op = peek()) {
      i += 1;
      const r = num(product());
      v = op === '+' ? num(v) + r : num(v) - r;
    }
    return v;
  };
  const concat = (): CfVal => {
    let v = sum();
    while (peek() === '&') {
      i += 1;
      v = text(v) + text(sum());
    }
    return v;
  };
  function comparison(): CfVal {
    let v = concat();
    for (;;) {
      peek();
      const op = CMP_TOKEN.exec(rest())?.[1];
      if (!op) return v;
      i += op.length;
      const d = compare(v, concat());
      v = op === '=' ? d === 0 : op === '<>' ? d !== 0 : op === '<' ? d < 0 : op === '>' ? d > 0 : op === '<=' ? d <= 0 : d >= 0;
    }
  }
  try {
    const v = comparison();
    return peek() === '' && truthy(v);
  } catch {
    return false;
  }
}
