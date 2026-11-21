/**
 * 엑셀 조건부 서식 — 편집기(화면 표시·규칙 만들기)와 서버(파일 읽기·행/열 이동·저장)가 함께 쓰는 형식과 판정.
 *
 * 규칙 목록은 시트마다 GRID_COND_FORMATS_MAP(Y.Map, 키 = 시트 번호)에 통째로 들어 있다(병합처럼 윈도잉하지 않는다).
 * 목록 순서가 곧 우선순위다(앞이 높다) — 같은 칸에 여러 규칙이 맞으면 앞 규칙의 색이 이긴다(엑셀과 같다).
 *
 * 파일에서 읽은 규칙은 `p`(파일의 priority)를 갖는다 — 저장할 때 ExcelJS가 읽어 둔 원래 규칙 객체를 이 값으로
 * 다시 찾아 그대로 쓴다(그래서 화면에 못 그리는 종류(`t: 'other'`, 수식·아이콘 집합 등)도 잃지 않는다).
 */

import { dateTextToSerial } from './cellValueShared.js';

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
  /** 색조 — th: 색마다의 문턱값(cfvo), 없으면 최소·(50 백분위)·최대 */
  | { readonly t: 'scale'; readonly colors: readonly string[]; readonly th?: readonly CfThreshold[] }
  /** 데이터 막대 — th: [최소, 최대] 문턱값, len: 칸 폭 대비 [최소, 최대] %(기본 10·90), neg: 음수 막대 색(x14) */
  | { readonly t: 'bar'; readonly color: string; readonly th?: readonly [CfThreshold, CfThreshold]; readonly len?: readonly [number, number]; readonly neg?: string }
  /** 아이콘 집합 — th: 둘째 아이콘부터의 문턱값(첫 cfvo는 엑셀도 보지 않는다), gt: 이상(>=)이 아니라 초과(>) */
  | { readonly t: 'icon'; readonly set: string; readonly th: readonly CfThreshold[]; readonly reverse?: 1; readonly hide?: 1 }
  /** 수식 규칙 — 범위 왼쪽 위 칸 기준으로 쓴 수식(상대 참조는 칸마다 옮긴다) */
  | { readonly t: 'expr'; readonly f: string; readonly s: CfStyle }
  | { readonly t: 'other' };

/** 문턱값(엑셀 cfvo) — autoMin/autoMax는 x14 막대의 "자동"(0을 포함하는 최소·최대) */
export interface CfThreshold {
  readonly type: 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula' | 'autoMin' | 'autoMax';
  readonly v?: number;
  /** type이 formula일 때(또는 값 자리에 수식이 있을 때) 그 수식 */
  readonly f?: string;
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
  /** "true이면 중지"(stopIfTrue) — 이 규칙이 맞으면 우선순위가 낮은 규칙은 보지 않는다 */
  readonly stop?: 1;
  /** x14 확장(extLst) 규칙의 priority — 정렬에만 쓴다(저장은 파일의 확장 부분을 그대로 둔다) */
  readonly xp?: number;
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
/** 행/열 삽입·삭제 뒤의 구간(양 끝 포함) — 다 지워지면 null. 조건부 서식 범위·엑셀 표 범위가 함께 쓴다 */
export function shiftSpan(start: number, end: number, at: number, count: number, insert: boolean): [number, number] | null {
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
/** 칸 값 → 수. 날짜 글자("2026-11-20")는 엑셀처럼 날짜 일련번호로 본다(날짜 칸과 TODAY() 비교, 날짜 색조 등). */
export function cfNumber(value: string): number | null {
  const t = value.trim().replace(/,/g, '');
  if (t === '') return null;
  if (NUMERIC.test(t)) return Number(t);
  return dateTextToSerial(t);
}

/** 범위 전체를 보아야 하는 규칙(중복·상위·평균·색조·막대·아이콘)의 통계 */
export interface CfStats {
  readonly counts?: ReadonlyMap<string, number>;
  readonly threshold?: number;
  readonly mean?: number;
  readonly min?: number;
  readonly max?: number;
  /** 백분위 문턱값용 — 오름차순 값 */
  readonly sorted?: readonly number[];
  /** 서버가 보낸 압축형 중복 통계 — counts에는 2개 이상인 값만 있고, 없는 값은 1개(빈 칸 아님) */
  readonly dupOnly?: boolean;
  /** 서버가 보낸 압축형 백분위 — 규칙이 쓰는 백분위(문턱값 v)마다 값 하나(sorted 대신) */
  readonly pct?: Readonly<Record<string, number>>;
}

export function needsStats(rule: CfRule): boolean {
  return rule.t === 'dup' || rule.t === 'top' || rule.t === 'avg' || rule.t === 'scale' || rule.t === 'bar' || rule.t === 'icon';
}

/** values: 규칙 범위 안의 (불러와 둔) 값들 */
export function computeStats(rule: CfRule, values: readonly string[]): CfStats {
  const counts = new Map<string, number>();
  for (const v of values) counts.set(v, (counts.get(v) ?? 0) + 1);
  return computeStatsFromCounts(rule, counts);
}

/** 값 글자 → 개수 표로 통계(서버는 이 표를 칸 편집마다 −1/+1 한다). 색조·막대·아이콘은 sorted를 펼친다 */
export function computeStatsFromCounts(rule: CfRule, raw: ReadonlyMap<string, number>): CfStats {
  if (rule.t === 'dup') {
    const counts = new Map<string, number>();
    for (const [v, n] of raw) {
      const k = v.trim().toLowerCase();
      if (k !== '') counts.set(k, (counts.get(k) ?? 0) + n);
    }
    return { counts };
  }
  const pairs: Array<[number, number]> = [];
  let total = 0;
  for (const [v, n] of raw) {
    const x = cfNumber(v);
    if (x === null || n <= 0) continue;
    pairs.push([x, n]);
    total += n;
  }
  if (total === 0) return {};
  if (rule.t === 'top') {
    pairs.sort((a, b) => (rule.bottom ? a[0] - b[0] : b[0] - a[0]));
    const want = Math.min(total, rule.percent ? Math.max(1, Math.floor((total * rule.n) / 100)) : Math.max(1, rule.n));
    let seen = 0;
    for (const [x, n] of pairs) {
      seen += n;
      if (seen >= want) return { threshold: x };
    }
    return {};
  }
  if (rule.t === 'avg') return { mean: pairs.reduce((a, [x, n]) => a + x * n, 0) / total };
  pairs.sort((a, b) => a[0] - b[0]);
  const sorted: number[] = [];
  for (const [x, n] of pairs) for (let i = 0; i < n; i += 1) sorted.push(x);
  return { min: sorted[0], max: sorted[sorted.length - 1], sorted };
}

/** 서버 → 편집 화면으로 보낼 작은 통계(값 표·정렬 배열 대신 판정에 쓰는 것만). 중복 값이 maxKeys보다 많으면 null */
export function compactStats(rule: CfRule, st: CfStats, maxKeys = 20_000): CfStats | null {
  if (rule.t === 'dup') {
    const counts = new Map<string, number>();
    for (const [k, n] of st.counts ?? []) {
      if (n < 2) continue;
      if (counts.size >= maxKeys) return null;
      counts.set(k, n);
    }
    return { counts, dupOnly: true };
  }
  if (!st.sorted) return st;
  const pct: Record<string, number> = {};
  const ths = rule.t === 'scale' || rule.t === 'bar' || rule.t === 'icon' ? (rule.th ?? (rule.t === 'scale' && rule.colors.length === 3 ? DEFAULT_SCALE3 : [])) : [];
  for (const th of ths) if (th.type === 'percentile') pct[String(th.v ?? 0)] = percentileOf(st.sorted, th.v ?? 0);
  return { min: st.min, max: st.max, pct };
}

/** 오름차순 값의 p번째 백분위수(엑셀 PERCENTILE.INC와 같은 보간) */
function percentileOf(sorted: readonly number[], p: number): number {
  const pos = ((sorted.length - 1) * Math.max(0, Math.min(100, p))) / 100;
  const lo = Math.floor(pos);
  return sorted[lo] + (sorted[Math.min(lo + 1, sorted.length - 1)] - sorted[lo]) * (pos - lo);
}

export interface CfPaint {
  style?: CfStyle;
  /** 데이터 막대: 색과 길이(0~1, 칸 폭 대비) */
  bar?: { color: string; ratio: number };
  /** 색조: 칸 배경 */
  scale?: string;
  /** 아이콘 집합: 칸 왼쪽에 그릴 기호와 색(hide면 값은 숨긴다) */
  icon?: { ch: string; color: string; hide?: true };
}

function mixHex(a: string, b: string, t: number): string {
  const pa = [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16));
  const pb = [1, 3, 5].map((i) => parseInt(b.slice(i, i + 2), 16));
  return `#${pa.map((x, i) => Math.round(x + (pb[i] - x) * Math.max(0, Math.min(1, t))).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

/** 수식 엔진(fast-formula-parser)으로 계산하는 대체 경로 — 작은 계산기가 못 하는 범위·함수(COUNTIF($A:$A,$A2) 등)용.
 * formula는 이미 이 칸 기준으로 참조를 옮긴 수식, 결과가 없으면(오류·지원 안 함) undefined. */
export type CfFormulaFallback = (formula: string, row: number, col: number) => CfVal | undefined | typeof CF_CANNOT;

/** 문턱값 하나(cfvo)를 수로 — 최소·최대·수·백분율·백분위·수식, x14의 자동 최소·최대(0을 포함) */
function thresholdOf(
  th: CfThreshold,
  st: CfStats,
  anchor: CfRect | undefined,
  valueAt: (row: number, col: number) => string,
  env: CfEnv | undefined,
): number | null {
  if (st.min === undefined || st.max === undefined) return null;
  switch (th.type) {
    case 'min':
      return st.min;
    case 'max':
      return st.max;
    case 'autoMin':
      return Math.min(0, st.min);
    case 'autoMax':
      return Math.max(0, st.max);
    case 'num':
      return th.v ?? null;
    case 'percent':
      return st.min + ((st.max - st.min) * (th.v ?? 0)) / 100;
    case 'percentile':
      return st.sorted ? percentileOf(st.sorted, th.v ?? 0) : (st.pct?.[String(th.v ?? 0)] ?? null);
    case 'formula': {
      const v = th.f && anchor ? evalCfValue(th.f, 0, 0, anchor.r0, anchor.c0, valueAt, env) : undefined;
      return typeof v === 'number' ? v : typeof v === 'string' ? cfNumber(v) : null;
    }
    default:
      return null;
  }
}

/**
 * 칸 하나에 걸리는 강조. entries는 우선순위 순(앞이 높다) — 높은 규칙부터 보며 아직 정해지지 않은 속성만 채우고,
 * "true이면 중지"(stopIfTrue)인 규칙이 맞으면 거기서 멈춘다(엑셀과 같다). stats는 entry.id별 통계(needsStats인 것만).
 * 큰 시트에서는 부르는 쪽이 이 칸이 든 행에 걸리는 규칙만 골라 넘긴다(cfRowSweeper·편집 화면 행 기억).
 */
export function paintCell(
  entries: readonly CondFormatEntry[],
  row: number,
  col: number,
  value: string,
  stats: ReadonlyMap<string, CfStats>,
  valueAt: (row: number, col: number) => string,
  env?: CfEnv,
): CfPaint | null {
  const paint: CfPaint = {};
  let any = false;
  for (const entry of entries) {
    if (!entry.ranges.some((r) => rectContains(r, row, col))) continue;
    const rule = entry.rule;
    const st = stats.get(entry.id) ?? {};
    const anchor = entry.ranges[0];
    let hit: CfStyle | undefined;
    let matched = false;
    switch (rule.t) {
      case 'cell': {
        const dr = row - anchor.r0;
        const dc = col - anchor.c0;
        const a = evalCfValue(rule.v1, dr, dc, row, col, valueAt, env);
        const b = rule.v2 === undefined ? undefined : evalCfValue(rule.v2, dr, dc, row, col, valueAt, env);
        matched = cellRuleMatches(rule.op, cellVal(value), a, b);
        break;
      }
      case 'text': {
        const v = value.toLowerCase();
        const t = rule.text.toLowerCase();
        matched = rule.op === 'contains' ? v.includes(t) : rule.op === 'notContains' ? !v.includes(t) : rule.op === 'begins' ? v.startsWith(t) : v.endsWith(t);
        break;
      }
      case 'dup': {
        const k = value.trim().toLowerCase();
        if (k !== '') {
          const n = st.counts?.get(k) ?? (st.dupOnly ? 1 : 0);
          matched = rule.unique ? n === 1 : n > 1;
        }
        break;
      }
      case 'top': {
        const n = cfNumber(value);
        matched = n !== null && st.threshold !== undefined && (rule.bottom ? n <= st.threshold : n >= st.threshold);
        break;
      }
      case 'avg': {
        const n = cfNumber(value);
        matched = n !== null && st.mean !== undefined && (rule.below ? n < st.mean : n > st.mean);
        break;
      }
      case 'scale': {
        const n = cfNumber(value);
        if (n === null || paint.scale) break;
        const stops = (rule.th ?? (rule.colors.length === 3 ? DEFAULT_SCALE3 : DEFAULT_SCALE2)).map((th) => thresholdOf(th, st, anchor, valueAt, env));
        if (stops.some((s) => s === null)) break;
        const at = stops as number[];
        const colors = rule.colors;
        let color = colors[0];
        if (n >= at[at.length - 1]) color = colors[colors.length - 1];
        else if (n > at[0]) {
          for (let k = 1; k < at.length; k += 1) {
            if (n <= at[k]) {
              color = mixHex(colors[k - 1], colors[k], at[k] === at[k - 1] ? 1 : (n - at[k - 1]) / (at[k] - at[k - 1]));
              break;
            }
          }
        }
        paint.scale = color;
        matched = true;
        break;
      }
      case 'bar': {
        const n = cfNumber(value);
        if (n === null || paint.bar) break;
        const lo = thresholdOf(rule.th?.[0] ?? { type: 'min' }, st, anchor, valueAt, env);
        const hi = thresholdOf(rule.th?.[1] ?? { type: 'max' }, st, anchor, valueAt, env);
        if (lo === null || hi === null) break;
        const ratio = hi === lo ? 1 : Math.max(0, Math.min(1, (n - lo) / (hi - lo)));
        // 막대 길이는 칸 폭의 minLength~maxLength %(엑셀 기본 10~90, x14 막대는 보통 0~100)
        const minL = (rule.len?.[0] ?? 10) / 100;
        const maxL = (rule.len?.[1] ?? 90) / 100;
        paint.bar = { color: n < 0 && rule.neg ? rule.neg : rule.color, ratio: minL + (maxL - minL) * ratio };
        matched = true;
        break;
      }
      case 'icon': {
        const n = cfNumber(value);
        const icons = CF_ICON_SETS[rule.set] ?? CF_ICON_SETS['3TrafficLights1'];
        if (n === null || paint.icon) break;
        let level = 0;
        for (let k = Math.min(rule.th.length, icons.length - 1); k >= 1; k -= 1) {
          const th = rule.th[k - 1];
          const at = thresholdOf(th, st, anchor, valueAt, env);
          if (at !== null && (th.gt ? n > at : n >= at)) {
            level = k;
            break;
          }
        }
        const icon = icons[rule.reverse ? icons.length - 1 - level : level];
        paint.icon = { ...icon, ...(rule.hide ? { hide: true as const } : {}) };
        matched = true;
        break;
      }
      case 'expr':
        matched = truthy(evalCfValue(rule.f, row - anchor.r0, col - anchor.c0, row, col, valueAt, env) ?? false);
        break;
      default:
        break;
    }
    if (matched && 's' in rule) hit = rule.s;
    if (hit) {
      const merged: { bg?: string; col?: string; b?: 1 } = { ...paint.style };
      if (hit.bg && merged.bg === undefined) merged.bg = hit.bg;
      if (hit.col && merged.col === undefined) merged.col = hit.col;
      if (hit.b && merged.b === undefined) merged.b = 1;
      paint.style = merged;
    }
    if (matched) any = true;
    if (matched && entry.stop) break;
  }
  return any ? paint : null;
}

const DEFAULT_SCALE2: readonly CfThreshold[] = [{ type: 'min' }, { type: 'max' }];
const DEFAULT_SCALE3: readonly CfThreshold[] = [{ type: 'min' }, { type: 'percentile', v: 50 }, { type: 'max' }];

/** 칸 값 글자 → 비교용 값(빈 칸 null, 수·날짜는 수, TRUE/FALSE는 논리값, 나머지는 글자) */
function cellVal(v: string): CfVal {
  if (v.trim() === '') return null;
  const n = cfNumber(v);
  if (n !== null) return n;
  if (v === 'TRUE' || v === 'FALSE') return v === 'TRUE';
  return v;
}

/** "셀 값 > …" 비교 — 엑셀은 빈 칸을 0으로 보고, 수끼리는 수로, 글자끼리는 대소문자 없이 비교한다 */
function cellRuleMatches(op: CfCellOp, value: CfVal, a: CfVal | undefined, b: CfVal | undefined): boolean {
  if (a === undefined) return false;
  const v = value === null ? 0 : value;
  if (op === 'between' || op === 'notBetween') {
    if (b === undefined) return false;
    const lo = compare(a, b) <= 0 ? a : b;
    const hi = lo === a ? b : a;
    const inside = compare(v, lo) >= 0 && compare(v, hi) <= 0;
    return op === 'between' ? inside : !inside;
  }
  const d = compare(v, a);
  switch (op) {
    case 'gt': return d > 0;
    case 'lt': return d < 0;
    case 'ge': return d >= 0;
    case 'le': return d <= 0;
    case 'eq': return d === 0;
    case 'ne': return d !== 0;
    default: return false;
  }
}

/** 수식 안에 TODAY()·NOW()가 있는지 — 날마다 결과가 바뀌는 규칙(보기 화면 캐시를 날짜별로 다시 만든다) */
export function cfIsVolatile(entries: readonly CondFormatEntry[]): boolean {
  const vol = (f: string | undefined) => Boolean(f && /\b(TODAY|NOW)\s*\(/i.test(f));
  return entries.some((e) => {
    const r = e.rule;
    if (r.t === 'expr') return vol(r.f);
    if (r.t === 'cell') return vol(r.v1) || vol(r.v2);
    if (r.t === 'scale' || r.t === 'bar' || r.t === 'icon') return (r.th ?? []).some((th) => vol(th.f));
    return false;
  });
}

// ---------- 수식(규칙 수식·피연산자) 계산 ----------
// 조건부 서식 수식은 대개 짧다($C2>=1, MOD(ROW(),2)=0, $A2<TODAY() …). 그런 것은 여기 작은 계산기가 바로 하고,
// 범위(A1:B5, $A:$A)·다른 시트·모르는 함수가 나오면 수식 엔진(fast-formula-parser, CfFormulaFallback)에 넘긴다.
export type CfVal = number | string | boolean | null; // null = 빈 칸

const UNSUPPORTED = new Error('cf: unsupported');

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

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
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
  // 오늘(서버·브라우저 현지 날짜) — 보기 화면 캐시는 날짜가 바뀌면 다시 만든다(cfIsVolatile)
  TODAY: () => {
    const d = new Date();
    return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EXCEL_EPOCH) / 86_400_000;
  },
  NOW: () => {
    const d = new Date();
    return (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - EXCEL_EPOCH) / 86_400_000;
  },
  DATE: (a) => (Date.UTC(num(a[0] ?? null), num(a[1] ?? null) - 1, num(a[2] ?? null)) - EXCEL_EPOCH) / 86_400_000,
};

const NUM_TOKEN = /^(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?/;
const REF_TOKEN = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(!])/;
const FUNC_TOKEN = /^([A-Za-z][A-Za-z0-9.]*)\s*\(/;
const BOOL_TOKEN = /^(TRUE|FALSE)(?![\w(])/i;
const CMP_TOKEN = /^(<>|<=|>=|=|<|>)/;

/** 상대 참조(A1, A$1, A:A)를 dRow·dCol만큼 옮긴다 — $가 붙은 쪽, 따옴표 안 글자, 함수 이름은 그대로(정렬·채우기·조건부 서식 공용) */
export function shiftFormulaRefs(formula: string, dRow: number, dCol: number): string {
  if (dRow === 0 && dCol === 0) return formula;
  const colName = (c: number) => {
    let n = c;
    let s = '';
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };
  const colNum = (letters: string) => [...letters.toUpperCase()].reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
  const moveCol = (abs: string, letters: string) => (abs ? `$${letters}` : colName(Math.max(1, colNum(letters) + dCol)));
  let out = '';
  let i = 0;
  const cellRef = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/y;
  const colRange = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(])/y;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch === '"' || ch === "'") {
      const end = formula.indexOf(ch, i + 1);
      const stop = end < 0 ? formula.length : end + 1;
      out += formula.slice(i, stop);
      i = stop;
      continue;
    }
    const boundary = i === 0 || !/[A-Za-z0-9_.$]/.test(formula[i - 1]);
    if (boundary) {
      colRange.lastIndex = i;
      const cr = colRange.exec(formula);
      if (cr) {
        out += `${moveCol(cr[1], cr[2])}:${moveCol(cr[3], cr[4])}`;
        i += cr[0].length;
        continue;
      }
      cellRef.lastIndex = i;
      const m = cellRef.exec(formula);
      if (m) {
        const [, colAbs, col, rowAbs, row] = m;
        out += `${moveCol(colAbs, col)}${rowAbs}${rowAbs ? Number(row) : Math.max(1, Number(row) + dRow)}`;
        i += m[0].length;
        continue;
      }
    }
    out += ch;
    i += 1;
  }
  return out;
}

/**
 * 조건부 서식 수식·피연산자 하나를 칸 (row, col)에 대해 계산한다. dr/dc: 범위 왼쪽 위에서 이 칸까지 — $ 없는 참조를 이만큼 옮긴다.
 * 작은 계산기가 못 하는 것(범위·다른 시트·모르는 함수)은 fallback(수식 엔진)에 옮긴 수식으로 넘긴다. 계산할 수 없으면 undefined.
 */
export function evalCfValue(
  formula: string,
  dr: number,
  dc: number,
  row: number,
  col: number,
  valueAt: (row: number, col: number) => string,
  env?: CfEnv,
): CfVal | undefined {
  // 상수(숫자·글자·논리값)는 수식마다 한 번만 읽는다 — 칸마다 같은 "10"을 다시 해석하지 않게
  const constant = constantOperand(formula);
  if (constant !== NOT_CONSTANT) return constant;
  const original = formula.trim().replace(/^=/, '');
  const src = rewriteSumproduct(original);
  try {
    return evalSmall(src, dr, dc, row, col, valueAt, env);
  } catch (error) {
    if (error === CF_PENDING) return undefined; // 서버 범위 값 답을 기다리는 중 — 오면 다시 칠한다
    // 작은 계산기가 모르는 것(다른 시트·모르는 함수·못 읽는 문법)만 수식 엔진으로 — 계산 오류(#VALUE! 등)는 안 맞음
    const unsupported = error === UNSUPPORTED || (error instanceof Error && error.message === 'cf formula');
    if (!unsupported || !env?.fallback) return undefined;
    const v = env.fallback(shiftFormulaRefs(`=${original}`, dr, dc).slice(1), row, col);
    if (v === CF_CANNOT) {
      env.unsupported?.(); // 엔진도 계산하지 못하는 규칙 — 칠하지 않고 "일부 생략"을 알린다
      return undefined;
    }
    return v;
  }
}

/** 수식 엔진이 계산하지 못한다는 표시(모르는 함수·배열 비교 등) — 오류값(#N/A 등, 안 맞음)과 다르다 */
export const CF_CANNOT: unique symbol = Symbol('cf-cannot');
/** 범위 값 표를 서버에 물었고 아직 답이 없다 */
const CF_PENDING = new Error('cf: pending');

const sumproductMemo = new Map<string, string>();
const SP_RANGE = String.raw`\$?[A-Za-z]{1,3}\$?\d*:\$?[A-Za-z]{1,3}\$?\d*`;
const SP_ARG = String.raw`\$?[A-Za-z]{1,3}\$?\d+|"[^"]*"|-?\d+(?:\.\d+)?`;
const SP_PATTERN = new RegExp(
  String.raw`SUMPRODUCT\(\s*(?:--|1\s*\*\s*)?\(\s*(?:(${SP_RANGE})\s*=\s*(${SP_ARG})|(${SP_ARG})\s*=\s*(${SP_RANGE}))\s*\)\s*(?:\*\s*1\s*)?\)`,
  'gi',
);
/** 바꿔 쓴 SUMPRODUCT가 쓰는 내부 함수 — 범위에서 값과 **정확히 같은** 칸 수(조건·와일드카드로 읽지 않는다) */
const CF_EXACT_FN = 'CFEXACTCOUNT';
/**
 * SUMPRODUCT((범위=값)*1)·SUMPRODUCT(--(범위=값))는 "범위에서 값과 정확히 같은 칸 수"다 — 내부 함수 CFEXACTCOUNT(범위,값)으로 바꿔
 * 범위 값 개수 표(서버 답 포함)의 정확 키로 센다(수식 엔진은 이 배열 비교를 못 한다). COUNTIF로 바꾸면 "a*"·"<5"·"문의?"가 조건으로
 * 읽혀 다른 값과 맞았다(검토 r5 F1). 엑셀의 = 비교처럼 글자는 대소문자를 가리지 않고, 수와 글자는 다르다.
 * shortcut: 칸 값은 글자로만 들고 있어(그리드 모델에 "숫자처럼 보이는 글자" 표시가 없다) 글자 "1"과 숫자 1은 같은 값으로 센다 — 칸 모델에 글자 표시를 더할 때 고친다.
 */
export function rewriteSumproduct(src: string): string {
  if (!/SUMPRODUCT/i.test(src)) return src;
  let out = sumproductMemo.get(src);
  if (out === undefined) {
    out = src.replace(SP_PATTERN, (_m, r1: string | undefined, a1: string | undefined, a2: string | undefined, r2: string | undefined) => `${CF_EXACT_FN}(${r1 ?? r2},${a1 ?? a2})`);
    if (sumproductMemo.size > 1_000) sumproductMemo.clear();
    sumproductMemo.set(src, out);
  }
  return out;
}

const NOT_CONSTANT = Symbol('not-constant');
const constantMemo = new Map<string, CfVal | typeof NOT_CONSTANT>();
function constantOperand(formula: string): CfVal | typeof NOT_CONSTANT {
  let v = constantMemo.get(formula);
  if (v !== undefined) return v;
  const t = formula.trim().replace(/^=/, '').trim();
  const str = /^"((?:[^"]|"")*)"$/.exec(t);
  v = str ? str[1].replace(/""/g, '"') : /^-?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$/.test(t) ? Number(t) : /^(TRUE|FALSE)$/i.test(t) ? t.toUpperCase() === 'TRUE' : NOT_CONSTANT;
  if (constantMemo.size > 10_000) constantMemo.clear();
  constantMemo.set(formula, v);
  return v;
}

/** 조건부 서식 수식 규칙이 이 칸에서 참인지(계산할 수 없으면 거짓 — 엑셀도 오류는 안 맞음으로 본다) */
export function evalCfFormula(
  formula: string,
  dr: number,
  dc: number,
  row: number,
  col: number,
  valueAt: (row: number, col: number) => string,
  env?: CfEnv,
): boolean {
  return truthy(evalCfValue(formula, dr, dc, row, col, valueAt, env) ?? false);
}

function evalSmall(src: string, dr: number, dc: number, row: number, col: number, valueAt: (row: number, col: number) => string, env?: CfEnv): CfVal {
  // 범위(A1:B5, $A:$A)는 범위 함수의 인자로만 — 범위 값 표(rangeTableOf)를 계산 한 번(굽기·화면 한 번)에 한 번만 만든다
  const rangeArg = (): CfRange | null => {
    peek();
    const col = COL_RANGE_TOKEN.exec(rest());
    if (col) {
      i += col[0].length;
      const a = colNumber(col[2]) - 1 + (col[1] ? 0 : dc);
      const b = colNumber(col[4]) - 1 + (col[3] ? 0 : dc);
      return { r0: 0, r1: 1_048_575, c0: Math.min(a, b), c1: Math.max(a, b) };
    }
    const m = CELL_RANGE_TOKEN.exec(rest());
    if (!m) return null;
    i += m[0].length;
    const r0 = Number(m[4]) - 1 + (m[3] ? 0 : dr);
    const r1 = Number(m[8]) - 1 + (m[7] ? 0 : dr);
    const c0 = colNumber(m[2]) - 1 + (m[1] ? 0 : dc);
    const c1 = colNumber(m[6]) - 1 + (m[5] ? 0 : dc);
    if (r0 < 0 || c0 < 0 || r1 < 0 || c1 < 0) fail();
    return { r0: Math.min(r0, r1), r1: Math.max(r0, r1), c0: Math.min(c0, c1), c1: Math.max(c0, c1) };
  };
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
    if (ch === "'" || ch === '!' || ch === ':') throw UNSUPPORTED; // 다른 시트·범위 — 수식 엔진으로
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
      if (src[i] === ':') throw UNSUPPORTED;
      let c = 0;
      for (const l of ref[2].toUpperCase()) c = c * 26 + (l.charCodeAt(0) - 64);
      const r0 = Number(ref[4]) - 1 + (ref[3] ? 0 : dr);
      const c0 = c - 1 + (ref[1] ? 0 : dc);
      if (r0 < 0 || c0 < 0) fail();
      return cellVal(valueAt(r0, c0));
    }
    const name = FUNC_TOKEN.exec(rest());
    if (name) {
      i += name[0].length;
      const upper = name[1].toUpperCase();
      const rangeFn = CF_RANGE_FUNCS[upper];
      if (rangeFn) {
        // 첫 인자가 범위인 함수(COUNTIF·SUM·AVERAGE …) — 범위 값 표를 한 번 만들어 다시 쓴다
        const range = rangeArg();
        if (!range || !env?.extent) throw UNSUPPORTED;
        const rest2: CfVal[] = [];
        while (peek() === ',') {
          i += 1;
          rest2.push(comparison());
        }
        expect(')');
        return rangeFn(rangeTableOf(range, valueAt, env), rest2);
      }
      const fn = CF_FUNCS[upper];
      if (!fn) throw UNSUPPORTED;
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
    if (/^[A-Za-z_]/.test(rest())) throw UNSUPPORTED; // 이름 정의·열 범위(A:A) 등
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
  const v = comparison();
  if (peek() !== '') {
    if (peek() === ':' || peek() === '!') throw UNSUPPORTED;
    fail();
  }
  return v;
}


// ---------- 범위 함수(COUNTIF·SUM·AVERAGE·MIN·MAX·COUNT·COUNTA) — 범위마다 값 표 한 번 ----------
/** 굽기(또는 편집 화면 한 번 칠하기)에 함께 쓰는 것: 수식 엔진(대체 계산), 범위 값 표 캐시, 시트 크기 */
export interface CfEnv {
  readonly fallback?: CfFormulaFallback;
  /** 범위 키 → 값 표. 칸 값이 바뀌면 새 Map으로 */
  readonly tables?: Map<string, CfRangeTable>;
  readonly extent?: () => { rows: number; cols: number };
  /** 범위 값 표를 밖(편집 화면 — 서버가 답하는 범위)에서 받는다. undefined면 valueAt으로 만든다 */
  readonly remote?: (r0: number, c0: number, r1: number, c1: number) => CfRangeTable | undefined;
  /** 수식 엔진도 계산하지 못한 규칙이 있었다 — "일부 생략" 알림 */
  readonly unsupported?: () => void;
}

interface CfRange {
  readonly r0: number;
  readonly r1: number;
  readonly c0: number;
  readonly c1: number;
}

export interface CfRangeTable {
  readonly cells: number;
  /** 칸 값(이 화면에서 직접 센 표). 서버 개수 표로 만든 표는 counts·sorted만 있다 */
  readonly values?: readonly CfVal[];
  counts?: Map<string, number>;
  sorted?: number[];
  /** 편집 화면 — 서버가 답한 집계와 물어본 조건별 개수(범위 값 표 전체를 받지 않는다) */
  readonly remote?: CfRemoteRange;
}

/** 범위 집계(COUNT·COUNTA·SUM·MIN·MAX — AVERAGE는 SUM/COUNT) */
export interface CfRangeAgg {
  readonly count: number;
  readonly counta: number;
  readonly sum: number;
  readonly min: number;
  readonly max: number;
}
/** 서버 답 — 아직 없는 조건은 ask로 물어 두고 CF_PENDING(나중에 다시 칠한다) */
export interface CfRemoteRange {
  readonly agg?: CfRangeAgg;
  /** cfCritKey(조건) / cfExactKey(값) → 개수 */
  readonly crit: ReadonlyMap<string, number>;
  ask(crit: CfVal, exact?: boolean): void;
}
/** COUNTIF 조건 하나의 키, 정확히 같은 값 개수의 키(서버·편집 화면 공용) */
export const cfCritKey = (crit: CfVal): string => JSON.stringify(crit);
export const cfExactKey = (value: CfVal): string => `=${JSON.stringify(value)}`;

/** 서버 — 범위 값 표에서 집계와 물어본 조건들의 개수만 답한다 */
export function cfRangeAnswer(t: CfRangeTable, crits: readonly CfVal[], exact: readonly CfVal[] = []): { agg: CfRangeAgg; crit: Array<[string, number]> } {
  const s = sortedOf(t);
  const agg = { count: s.length, counta: [...countsOf(t).values()].reduce((a, b) => a + b, 0), sum: s.reduce((x, y) => x + y, 0), min: s[0] ?? 0, max: s[s.length - 1] ?? 0 };
  return { agg, crit: [...crits.map((c): [string, number] => [cfCritKey(c), countIf(t, c)]), ...exact.map((v): [string, number] => [cfExactKey(v), exactCount(t, v)])] };
}

/** 범위에서 값과 정확히 같은 칸 수 — 빈 값은 빈 칸 수, 글자는 대소문자 무시, 수와 글자는 다르다(엑셀 = 비교) */
function exactCount(t: CfRangeTable, v: CfVal): number {
  if (v === null || v === '') return t.cells - [...countsOf(t).values()].reduce((a, b) => a + b, 0);
  return countsOf(t).get(countKey(v)) ?? 0;
}

/** cfRangeTableFromRaw를 chunk 항목마다 tick()으로 양보하며 만든다(서버 — 큰 표를 만드는 동안 다른 요청이 멈추지 않게) */
export async function cfRangeTableFromRawChunked(cells: number, raw: ReadonlyMap<string, number>, tick: () => Promise<void>, chunk = 20_000): Promise<CfRangeTable> {
  const counts = new Map<string, number>();
  const sorted: number[] = [];
  let i = 0;
  for (const [v, n] of raw) {
    const cv = cellVal(v);
    if (cv !== null && n > 0) {
      counts.set(countKey(cv), (counts.get(countKey(cv)) ?? 0) + n);
      if (typeof cv === 'number') for (let k = 0; k < n; k += 1) sorted.push(cv);
    }
    i += 1;
    if (i % chunk === 0) await tick();
  }
  sorted.sort((a, b) => a - b);
  return { cells, counts, sorted };
}

/** 값 글자 → 개수 표(빈 칸 제외)와 칸 수 → 범위 값 표(서버) */
export function cfRangeTableFromRaw(cells: number, raw: ReadonlyMap<string, number>): CfRangeTable {
  const counts = new Map<string, number>();
  const sorted: number[] = [];
  for (const [v, n] of raw) {
    const cv = cellVal(v);
    if (cv === null || n <= 0) continue;
    counts.set(countKey(cv), (counts.get(countKey(cv)) ?? 0) + n);
    if (typeof cv === 'number') for (let i = 0; i < n; i += 1) sorted.push(cv);
  }
  sorted.sort((a, b) => a - b);
  return { cells, counts, sorted };
}

const COL_RANGE_TOKEN = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![\w(])/;
const CELL_RANGE_TOKEN = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+):(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(])/;
function colNumber(letters: string): number {
  let n = 0;
  for (const l of letters.toUpperCase()) n = n * 26 + (l.charCodeAt(0) - 64);
  return n;
}

function rangeTableOf(range: CfRange, valueAt: (row: number, col: number) => string, env: CfEnv): CfRangeTable {
  const { rows, cols } = env.extent ? env.extent() : { rows: 0, cols: 0 };
  const r1 = Math.min(range.r1, rows - 1);
  const c1 = Math.min(range.c1, cols - 1);
  const key = `${range.r0}:${range.c0}:${r1}:${c1}`;
  const hit = env.tables?.get(key) ?? env.remote?.(range.r0, range.c0, r1, c1);
  if (hit) {
    env.tables?.set(key, hit);
    return hit;
  }
  const values: CfVal[] = [];
  for (let r = range.r0; r <= r1; r += 1) for (let c = range.c0; c <= c1; c += 1) values.push(cellVal(valueAt(r, c)));
  const table: CfRangeTable = { cells: Math.max(0, (r1 - range.r0 + 1) * (c1 - range.c0 + 1)), values };
  env.tables?.set(key, table);
  return table;
}

/** 서버 — 칸 편집으로 값 하나가 늘거나 줄었다. 표를 다시 만들지 않고 고친다(개수 표 −1/+1, 정렬된 수는 이진 탐색으로 넣고 뺀다) */
export function cfRangeTableAdjust(t: CfRangeTable, text: string, delta: 1 | -1): void {
  const cv = cellVal(text);
  if (cv === null) return;
  const counts = countsOf(t);
  const k = countKey(cv);
  const n = (counts.get(k) ?? 0) + delta;
  if (n > 0) counts.set(k, n);
  else counts.delete(k);
  if (typeof cv !== 'number') return;
  const sorted = sortedOf(t);
  const i = countBelow(sorted, cv, false);
  if (delta > 0) sorted.splice(i, 0, cv);
  else if (sorted[i] === cv) sorted.splice(i, 1);
}

const countKey = (v: CfVal): string => (typeof v === 'number' ? `n${v}` : typeof v === 'boolean' ? `b${v}` : `s${String(v ?? '').toLowerCase()}`);
function countsOf(t: CfRangeTable): Map<string, number> {
  if (!t.counts) {
    t.counts = new Map();
    for (const v of t.values ?? []) if (v !== null) t.counts.set(countKey(v), (t.counts.get(countKey(v)) ?? 0) + 1);
  }
  return t.counts;
}
function sortedOf(t: CfRangeTable): number[] {
  if (!t.sorted) t.sorted = (t.values ?? []).filter((v): v is number => typeof v === 'number').sort((a, b) => a - b);
  return t.sorted;
}
/** 오름차순 배열에서 x보다 작은(strict) 또는 작거나 같은 값의 개수 */
function countBelow(sorted: readonly number[], x: number, inclusive: boolean): number {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (inclusive ? sorted[mid] <= x : sorted[mid] < x) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

/** COUNTIF 조건 — "사과"(같음), 10, ">=10", "<>완료", "사*"(와일드카드) */
function countIf(t: CfRangeTable, crit: CfVal): number {
  if (crit === null || crit === '') return t.cells - [...countsOf(t).values()].reduce((a, b) => a + b, 0);
  if (typeof crit !== 'string') return countsOf(t).get(countKey(crit)) ?? 0;
  const m = /^(<=|>=|<>|<|>|=)(.*)$/.exec(crit);
  const op = m ? m[1] : '=';
  const operand = m ? m[2] : crit;
  const n = cfNumber(operand);
  if (op === '=' || op === '<>') {
    let eq: number;
    if (operand === '') eq = t.cells - [...countsOf(t).values()].reduce((a, b) => a + b, 0);
    else if (/[*?]/.test(operand)) {
      const re = new RegExp(`^${operand.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.')}$`, 'i');
      eq = 0;
      for (const [k, n] of countsOf(t)) if (k.startsWith('s') && re.test(k.slice(1))) eq += n;
    } else eq = (countsOf(t).get(`s${operand.toLowerCase()}`) ?? 0) + (n !== null ? (countsOf(t).get(`n${n}`) ?? 0) : 0);
    return op === '=' ? eq : t.cells - eq;
  }
  if (n === null) return 0;
  const s = sortedOf(t);
  if (op === '<') return countBelow(s, n, false);
  if (op === '<=') return countBelow(s, n, true);
  if (op === '>') return s.length - countBelow(s, n, true);
  return s.length - countBelow(s, n, false);
}

/** 서버 답에서 — 없으면 물어 두고 기다린다 */
function remoteAgg(t: CfRangeTable): CfRangeAgg | null {
  if (!t.remote) return null;
  if (!t.remote.agg) {
    t.remote.ask(null);
    throw CF_PENDING;
  }
  return t.remote.agg;
}
const CF_RANGE_FUNCS: Readonly<Record<string, (t: CfRangeTable, args: CfVal[]) => CfVal>> = {
  [CF_EXACT_FN]: (t, a) => {
    if (!t.remote) return exactCount(t, a[0] ?? null);
    const n = t.remote.crit.get(cfExactKey(a[0] ?? null));
    if (n !== undefined) return n;
    t.remote.ask(a[0] ?? null, true);
    throw CF_PENDING;
  },
  COUNTIF: (t, a) => {
    if (!t.remote) return countIf(t, a[0] ?? null);
    const n = t.remote.crit.get(cfCritKey(a[0] ?? null));
    if (n !== undefined) return n;
    t.remote.ask(a[0] ?? null);
    throw CF_PENDING;
  },
  COUNT: (t) => remoteAgg(t)?.count ?? sortedOf(t).length,
  COUNTA: (t) => remoteAgg(t)?.counta ?? [...countsOf(t).values()].reduce((a, b) => a + b, 0),
  SUM: (t) => remoteAgg(t)?.sum ?? sortedOf(t).reduce((x, y) => x + y, 0),
  AVERAGE: (t) => {
    const a = remoteAgg(t);
    const s = a ? null : sortedOf(t);
    const count = a ? a.count : s!.length;
    if (count === 0) throw new Error('#DIV/0!');
    return (a ? a.sum : s!.reduce((x, y) => x + y, 0)) / count;
  },
  MIN: (t) => remoteAgg(t)?.min ?? sortedOf(t)[0] ?? 0,
  MAX: (t) => remoteAgg(t)?.max ?? sortedOf(t)[sortedOf(t).length - 1] ?? 0,
};

/** 행 순서대로 그 행에 걸린 규칙을 내 준다(행마다 배열을 미리 만들지 않는다 — 규칙 수에 비례하는 메모리). rows는 오름차순으로만 물어야 한다 */
export function cfRowSweeper(entries: readonly CondFormatEntry[]): (row: number) => readonly CondFormatEntry[] {
  const order = new Map(entries.map((e, i) => [e, i]));
  const spans = entries.flatMap((e) => e.ranges.map((r) => ({ e, r0: r.r0, r1: r.r1 }))).sort((a, b) => a.r0 - b.r0);
  let next = 0;
  let active: Array<{ e: CondFormatEntry; r1: number }> = [];
  let cache: readonly CondFormatEntry[] = [];
  let dirty = true;
  return (row) => {
    const before = active.length;
    active = active.filter((a) => a.r1 >= row);
    if (active.length !== before) dirty = true;
    while (next < spans.length && spans[next].r0 <= row) {
      if (spans[next].r1 >= row) {
        active.push({ e: spans[next].e, r1: spans[next].r1 });
        dirty = true;
      }
      next += 1;
    }
    if (dirty) {
      const set = [...new Set(active.map((a) => a.e))];
      set.sort((a, b) => (order.get(a) ?? 0) - (order.get(b) ?? 0));
      cache = set;
      dirty = false;
    }
    return cache;
  };
}


// ---------- 무거운 규칙 어림(보기 화면 굽기·편집 화면이 같은 기준으로 건너뛴다) ----------
/** 수식 엔진으로 계산해야 하는 규칙의 어림 비용(규칙 칸 수 × 읽는 범위 칸 수) 상한 — 넘으면 그 규칙은 칠하지 않는다 */
export const CF_FALLBACK_COST_CAP = 20_000_000;
/** 칸 하나를 판정할 때 수식 엔진이 읽는 범위 칸 수 상한 — 넘으면 칸 하나에도 오래 걸리므로 건너뛴다 */
export const CF_FALLBACK_RANGE_CAP = 100_000;

const cfCellsIn = (e: CondFormatEntry, rows: number, cols: number) =>
  e.ranges.reduce((sum, r) => sum + Math.max(0, Math.min(r.r1, rows - 1) - r.r0 + 1) * Math.max(0, Math.min(r.c1, cols - 1) - r.c0 + 1), 0);

/** 수식 엔진이 칸마다 읽을 범위(행 구간·열 수). 작은 계산기가 바로 하는 범위 함수(COUNTIF 등 — 범위 값 표를 한 번만 만든다)만 쓰면 [] */
export function cfEngineRanges(e: CondFormatEntry, rows: number, cols: number): Array<{ r0: number; r1: number; cols: number }> {
  const formulas = e.rule.t === 'expr' ? [e.rule.f] : e.rule.t === 'cell' ? [e.rule.v1, e.rule.v2 ?? ''] : [];
  const out: Array<{ r0: number; r1: number; cols: number }> = [];
  for (const f0 of formulas) {
    const f = rewriteSumproduct(f0);
    const engineOnly = /\b(?!(?:COUNTIF|SUM|AVERAGE|MIN|MAX|COUNT|COUNTA|CFEXACTCOUNT)\s*\()[A-Z][A-Z0-9.]*\s*\(/i.test(f.replace(/\b(AND|OR|NOT|IF|ISBLANK|ISNUMBER|ISTEXT|LEN|ABS|INT|ROUND|MOD|ROW|COLUMN|LEFT|RIGHT|UPPER|LOWER|TRIM|TODAY|NOW|DATE)\s*\(/gi, '')) || f.includes('!');
    if (!engineOnly) continue;
    for (const m of f.matchAll(/\$?([A-Z]{1,3})\$?(\d*):\$?([A-Z]{1,3})\$?(\d*)/gi)) {
      const full = !(m[2] && m[4]);
      out.push({
        r0: full ? 0 : Math.min(Number(m[2]), Number(m[4])) - 1,
        r1: full ? rows - 1 : Math.max(Number(m[2]), Number(m[4])) - 1,
        cols: Math.min(cols, Math.abs(m[3].toUpperCase().charCodeAt(0) - m[1].toUpperCase().charCodeAt(0)) + 1),
      });
    }
  }
  return out;
}

/** 이 규칙이 너무 무거워 건너뛰어야 하는지(규칙 칸 수 × 읽는 범위 칸 수, 칸 하나가 읽는 범위 칸 수) */
export function cfRuleTooCostly(e: CondFormatEntry, rows: number, cols: number): boolean {
  const perCell = cfEngineRanges(e, rows, cols).reduce((sum, r) => sum + (r.r1 - r.r0 + 1) * r.cols, 0);
  return perCell > CF_FALLBACK_RANGE_CAP || cfCellsIn(e, rows, cols) * perCell > CF_FALLBACK_COST_CAP;
}

/**
 * 규칙 수식이 범위 함수에 쓰는 범위 중 행이 고정된 것($A$1:$A$500)·열 전체($A:$A) — 서버가 편집 화면에 답해도 되는 범위다
 * (편집 화면 rangeTableOf가 만드는 키와 같은 모양: 시트 크기로 자른 r0·c0·r1·c1). 열이 상대 참조면 규칙이 걸친 열마다 옮긴다.
 * 행이 상대 참조인 범위($A$1:$A1 — 칸마다 다름)는 넣지 않는다(편집 화면은 불러온 행으로 세거나 건너뛴다).
 */
export function cfRuleRanges(entries: readonly CondFormatEntry[], rows: number, cols: number, maxRects = 256): CfRect[] {
  const out = new Map<string, CfRect>();
  for (const e of entries) {
    const r = e.rule;
    const formulas = r.t === 'expr' ? [r.f] : r.t === 'cell' ? [r.v1, r.v2 ?? ''] : r.t === 'scale' || r.t === 'bar' || r.t === 'icon' ? (r.th ?? []).map((th) => th.f ?? '') : [];
    const anchor = e.ranges[0];
    if (!anchor) continue;
    const colSet = new Set<number>();
    for (const rect of e.ranges) for (let c = rect.c0; c <= Math.min(rect.c1, cols - 1) && colSet.size < 256; c += 1) colSet.add(c);
    for (const f0 of formulas) {
      const f = rewriteSumproduct(f0.replace(/^=/, ''));
      if (!f.includes(':')) continue;
      for (const col of colSet) {
        const dc = col - anchor.c0;
        for (let i = 0; i < f.length; i += 1) {
          if (i > 0 && /[A-Za-z0-9_.$]/.test(f[i - 1])) continue;
          const rest = f.slice(i);
          const cr = COL_RANGE_TOKEN.exec(rest);
          let rect: CfRect | null = null;
          if (cr) {
            const a = colNumber(cr[2]) - 1 + (cr[1] ? 0 : dc);
            const b = colNumber(cr[4]) - 1 + (cr[3] ? 0 : dc);
            rect = { r0: 0, r1: rows - 1, c0: Math.min(a, b), c1: Math.min(Math.max(a, b), cols - 1) };
          } else {
            const m = CELL_RANGE_TOKEN.exec(rest);
            if (m && m[3] && m[7]) {
              const a = colNumber(m[2]) - 1 + (m[1] ? 0 : dc);
              const b = colNumber(m[6]) - 1 + (m[5] ? 0 : dc);
              const r0 = Math.min(Number(m[4]), Number(m[8])) - 1;
              rect = { r0, r1: Math.min(Math.max(Number(m[4]), Number(m[8])) - 1, rows - 1), c0: Math.min(a, b), c1: Math.min(Math.max(a, b), cols - 1) };
            }
          }
          if (!rect || rect.c0 < 0 || rect.r1 < rect.r0 || rect.c1 < rect.c0) continue;
          out.set(`${rect.r0}:${rect.c0}:${rect.r1}:${rect.c1}`, rect);
          if (out.size >= maxRects) return [...out.values()];
        }
      }
    }
  }
  return [...out.values()];
}
