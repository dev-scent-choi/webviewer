/**
 * 엑셀 와일드카드(`*` 아무 글자들, `?` 한 글자, `~*` `~?` `~~` 글자 그대로) — 정규식 없이 맞춘다.
 * 예전엔 `*`를 `.*`로 바꾼 정규식을 썼는데, `*a*a*a*a*b` 같은 조건을 긴 칸에 대면 되돌아가기가 지수로 늘어 서버가 수십 초 멈췄다
 * (FINAL-REVIEW C1). 여기서는 `*`로 나눈 조각을 차례로 찾는다 — 조각 하나를 찾는 데 칸 길이 × 조각 길이라 전체도 그 정도다.
 * COUNTIF(조건부 서식·서버 답)와 수식 엔진(fast-formula-parser의 WildCard.toRegex를 덮어씀: COUNTIF·SUMIF·AVERAGEIF·MATCH·
 * VLOOKUP·HLOOKUP·SEARCH)이 함께 쓴다.
 */

/** null = 아무 한 글자(`?`) */
type Seg = (string | null)[];

function compile(pattern: string, ignoreCase: boolean): Seg[] {
  const p = ignoreCase ? pattern.toLowerCase() : pattern;
  const segs: Seg[] = [[]];
  for (let i = 0; i < p.length; i += 1) {
    const ch = p[i];
    if (ch === '~' && (p[i + 1] === '*' || p[i + 1] === '?' || p[i + 1] === '~')) {
      segs[segs.length - 1].push(p[i + 1]);
      i += 1;
    } else if (ch === '*') segs.push([]);
    else segs[segs.length - 1].push(ch === '?' ? null : ch);
  }
  return segs;
}

const segAt = (s: string, at: number, seg: Seg): boolean => {
  if (at < 0 || at + seg.length > s.length) return false;
  for (let j = 0; j < seg.length; j += 1) if (seg[j] !== null && seg[j] !== s[at + j]) return false;
  return true;
};
/** from 이상에서 seg가 맞는 가장 왼쪽 자리(못 찾으면 -1) */
function findFrom(s: string, from: number, seg: Seg): number {
  for (let i = from; i + seg.length <= s.length; i += 1) if (segAt(s, i, seg)) return i;
  return -1;
}
/** 끝이 end 이하인 자리 가운데 seg가 맞는 가장 오른쪽 시작(못 찾으면 -1) */
function findBackFrom(s: string, end: number, seg: Seg): number {
  for (let i = end - seg.length; i >= 0; i -= 1) if (segAt(s, i, seg)) return i;
  return -1;
}

/** 가운데 조각들(segs[a..b])을 from 뒤에 차례로(가장 왼쪽부터) 놓을 수 있으면 마지막 끝 자리, 아니면 -1 */
function placeMiddle(s: string, from: number, segs: Seg[], a: number, b: number): number {
  let at = from;
  for (let k = a; k <= b; k += 1) {
    const i = findFrom(s, at, segs[k]);
    if (i < 0) return -1;
    at = i + segs[k].length;
  }
  return at;
}

export interface WildcardMatcher {
  /** 칸 전체가 맞는가(엑셀 COUNTIF·MATCH와 같다) */
  test(value: unknown): boolean;
  /** 처음 맞는 자리(SEARCH, String.prototype.search가 부른다), 없으면 -1 */
  [Symbol.search](value: string): number;
}

/**
 * 와일드카드 작업량 상한(재검토 C1) — 맞추기 한 번은 칸 길이 × 조건 길이라 지수는 아니지만, 32k자 칸 수백 개 × 긴 조건 수천 개면
 * 요청 하나가 몇 시간 걸린다. 일한 만큼(칸 길이 + 1) × (조건 길이 + 1)을 예산에서 빼고, 모자라면 WildcardBudgetError를 던진다.
 * 예산은 범위(withWildcardBudget — 요청 하나·굽기 하나·업로드 계산 하나)마다 하나, 범위 밖이면 매처 하나(엑셀 함수 호출 하나)마다 하나다.
 */
export const WILDCARD_WORK_BUDGET = 5e7;
export class WildcardBudgetError extends Error {
  constructor() {
    super('wildcard work budget exceeded');
  }
}
export interface WildcardBudget {
  left: number;
}
let scoped: WildcardBudget | null = null;
/** fn(동기) 안의 모든 와일드카드 맞추기가 budget 하나를 나눠 쓴다 */
export function withWildcardBudget<T>(budget: WildcardBudget, fn: () => T): T {
  const outer = scoped;
  scoped = budget;
  try {
    return fn();
  } finally {
    scoped = outer;
  }
}
/** 남은 예산이 cost보다 적으면 던진다 — 일을 하기 전에 본다 */
export function chargeWildcardWork(cost: number, own: WildcardBudget): void {
  const budget = scoped ?? own;
  if (budget.left < cost) {
    budget.left = 0;
    throw new WildcardBudgetError();
  }
  budget.left -= cost;
}

export function wildcardMatcher(pattern: string, ignoreCase = true): WildcardMatcher {
  const segs = compile(pattern, ignoreCase);
  const norm = (v: unknown) => (ignoreCase ? String(v).toLowerCase() : String(v));
  const last = segs.length - 1;
  const own: WildcardBudget = { left: WILDCARD_WORK_BUDGET };
  // 실제로 하는 일만큼: 앞·뒤 조각은 한 번 비교, 가운데 조각은 칸 위를 훑는다(칸 길이 × 가운데 조각 길이)
  const middle = segs.slice(1, -1).reduce((a, seg) => a + seg.length, 0);
  const charge = (s: string, all = false) => chargeWildcardWork(all ? (s.length + 1) * (pattern.length + 1) : last >= 2 ? (s.length + 1) * (middle + 1) : pattern.length + 1, own);
  return {
    test(value: unknown): boolean {
      const s = norm(value);
      charge(s);
      if (last === 0) return s.length === segs[0].length && segAt(s, 0, segs[0]);
      if (!segAt(s, 0, segs[0])) return false;
      const tail = s.length - segs[last].length;
      if (tail < segs[0].length || !segAt(s, tail, segs[last])) return false;
      const end = placeMiddle(s, segs[0].length, segs, 1, last - 1);
      return end >= 0 && end <= tail;
    },
    [Symbol.search](value: string): number {
      const s = norm(value);
      charge(s, true); // 찾기는 모든 조각이 칸 위를 훑는다
      // 첫 조각 뒤에 나머지 조각을 놓을 수 있는 가장 늦은 자리(오른쪽부터 가장 오른쪽에 놓아 본다)
      let latest = s.length;
      for (let k = last; k >= 1; k -= 1) {
        latest = findBackFrom(s, latest, segs[k]);
        if (latest < 0) return -1;
      }
      for (let i = 0; i + segs[0].length <= latest; i += 1) if (segAt(s, i, segs[0])) return i;
      return -1;
    },
  };
}

/** fast-formula-parser의 와일드카드를 이 매처로 바꾼다 — 라이브러리 안 COUNTIF·SUMIF·AVERAGEIF·VLOOKUP·SEARCH가 `.test`/`search`만 쓴다 */
export function patchFormulaWildcards(api: { WildCard: { toRegex(value: unknown, flags?: string): unknown } }): void {
  api.WildCard.toRegex = (value: unknown, flags?: string) => wildcardMatcher(String(value), (flags ?? '').includes('i'));
}
