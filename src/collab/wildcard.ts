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
 * 와일드카드 작업량 상한(재검토 C1·재검토 2) — 맞추기 한 번은 칸 길이 × 조건 길이라 지수는 아니지만, 32k자 칸 수백 개 × 긴 조건 수천 개면
 * 요청 하나가 몇 시간 걸린다. 일한 만큼을 예산에서 빼고(값싼 조건도 칸마다 최소 2 × (칸 길이 + 1) — 잘라 내기·소문자로 바꾸기만으로도 칸을 두 번 훑는다),
 * 모자라면 WildcardBudgetError를 던진다.
 * - WILDCARD_WORK_BUDGET(1e9): 묶음 하나(cf-ranges 요청 하나·굽기 하나·업로드 수식 계산 하나·편집 방 쪽 하나 재계산)의 합.
 *   재검토 2 측정으로 5e7 ≈ 114ms(이 PC) → 1e9 ≈ 2.3초. 정상 파일(1,000행 × 칸 20자 × "*"&B1&"*")은 1,000칸 전체가 약 1.5e8이라 넉넉하다.
 * - WILDCARD_STEP_BUDGET(5e7, ≈ 114ms): 양보 없이 한 번에 하는 일 하나(칸 하나·조건 묶음 하나)의 상한 — 묶음 예산에서 떼어 쓴다(runWildcardStep).
 *   그래서 묶음 전체가 길어도 이벤트 루프를 한 번에 이보다 오래 잡지 않는다(묶음 쪽이 단계 사이에 양보한다).
 * 범위 밖(브라우저 등)에서는 매처 하나(엑셀 함수 호출 하나)마다 WILDCARD_STEP_BUDGET 하나다.
 */
export const WILDCARD_WORK_BUDGET = 1e9;
export const WILDCARD_STEP_BUDGET = 5e7;
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
/**
 * 묶음 예산 pass에서 단계 하나(최대 WILDCARD_STEP_BUDGET)를 떼어 fn을 그 예산으로 돌리고, 쓴 만큼 pass에서 뺀다.
 * fn은 받은 예산을 withWildcardBudget으로 쓴다(동기 구간마다) — 비동기 경계는 이 예산 객체를 인자로 넘겨 건넌다.
 */
export function runWildcardStep<T>(pass: WildcardBudget, fn: (step: WildcardBudget) => T): T {
  const step: WildcardBudget = { left: Math.min(WILDCARD_STEP_BUDGET, pass.left) };
  const start = step.left;
  const settle = () => {
    pass.left -= start - step.left;
  };
  let out: T;
  try {
    out = fn(step);
  } catch (error) {
    settle();
    throw error;
  }
  if (out instanceof Promise) return out.finally(settle) as T;
  settle();
  return out;
}
/** 예산 오류인가 — fast-formula-parser는 함수 안의 오류를 FormulaError('#ERROR!', 글, details=원래 오류)로 감싼다 */
export function isWildcardBudgetError(error: unknown): boolean {
  for (let e: unknown = error, depth = 0; e && depth < 5; depth += 1) {
    if (e instanceof WildcardBudgetError) return true;
    e = (e as { details?: unknown; cause?: unknown }).details ?? (e as { cause?: unknown }).cause;
  }
  return false;
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
  const own: WildcardBudget = { left: WILDCARD_STEP_BUDGET };
  // 실제로 하는 일만큼: 칸을 소문자로 바꾸며 한 번 훑고(칸 길이 + 1), 가운데 조각은 칸 위를 훑는다(칸 길이 × 가운데 조각 길이).
  // 칸을 바꾸기 전에 뺀다(값싼 조건도 칸마다 최소 칸 길이 + 1 — 재검토 2: 예전엔 조건 길이만 빼서 2,000조건 요청이 11초)
  const middle = segs.slice(1, -1).reduce((a, seg) => a + seg.length, 0);
  const charge = (v: unknown, all = false) => {
    const n = typeof v === 'string' ? v.length : String(v).length;
    // +2: 부르는 쪽이 칸을 한 번 잘라 내고(countIf의 slice) 여기서 소문자로 한 번 바꾼다 — 조각이 없어도 칸을 두 번 훑는다
    chargeWildcardWork((n + 1) * ((all ? pattern.length : middle) + 2), own);
  };
  return {
    test(value: unknown): boolean {
      charge(value);
      const s = norm(value);
      if (last === 0) return s.length === segs[0].length && segAt(s, 0, segs[0]);
      if (!segAt(s, 0, segs[0])) return false;
      const tail = s.length - segs[last].length;
      if (tail < segs[0].length || !segAt(s, tail, segs[last])) return false;
      const end = placeMiddle(s, segs[0].length, segs, 1, last - 1);
      return end >= 0 && end <= tail;
    },
    [Symbol.search](value: string): number {
      charge(value, true); // 찾기는 모든 조각이 칸 위를 훑는다
      const s = norm(value);
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
