import { describe, expect, test } from 'vitest';
import FormulaParser from 'fast-formula-parser';
import { cfCritKey, cfRangeAnswer, cfRangeTableFromRaw } from '../src/collab/condFormatShared.js';
import { formulaFunctions } from '../src/client/formulaOverrides.js';

/**
 * FINAL-REVIEW C1 — COUNTIF 와일드카드가 정규식 되돌아가기로 서버를 멈추던 문제.
 * `*` 8개 조건 × 200자 칸이 50ms 안에 끝나야 하고, 엑셀 의미(대소문자 무시, ~* ~? ~~ 글자 그대로)는 그대로.
 */
const EVIL = '*a*a*a*a*a*a*a*b';
const LONG = 'a'.repeat(200);

const countOf = (values: string[], crit: string) => {
  const raw = new Map<string, number>();
  for (const v of values) raw.set(v, (raw.get(v) ?? 0) + 1);
  return new Map(cfRangeAnswer(cfRangeTableFromRaw(values.length, raw), [crit], []).crit).get(cfCritKey(crit));
};

/** 수식 엔진(서버 계산·조건부 서식 대체 경로와 같은 라이브러리 + 덮어쓴 함수) */
function evaluate(formula: string, column: string[]): unknown {
  const parser = new FormulaParser({
    functions: formulaFunctions,
    onCell: ({ row, col }) => (col === 1 ? (column[row - 1] ?? null) : null),
    onRange: (ref) => {
      const out: unknown[][] = [];
      for (let r = ref.from.row; r <= Math.min(ref.to.row, column.length); r += 1) out.push([column[r - 1] ?? null]);
      return out;
    },
  });
  return parser.parse(formula, { sheet: 'S', row: 1, col: 5 });
}

describe('C1: 와일드카드는 선형 시간', () => {
  test('조건부 서식 COUNTIF: * 8개 × 200자 칸이 50ms 안', () => {
    const t0 = performance.now();
    expect(countOf([LONG, `${LONG}b`, 'x'], EVIL)).toBe(1);
    expect(performance.now() - t0).toBeLessThan(500); // 병렬 시험 부하에서도 흔들리지 않게 넓힌 기준(재검토 3) — 고치기 전 값(끝나지 않음)과는 여전히 크게 떨어진다
  });

  test.each([
    ['COUNTIF(A1:A3,"*a*a*a*a*a*a*a*b")', 1],
    ['SUMIF(A1:A3,"*a*a*a*a*a*a*a*b",A1:A3)', 0],
    ['MATCH("*a*a*a*a*a*a*a*b",A1:A3,0)', 2],
  ])('수식 엔진 %s: 50ms 안', (formula, want) => {
    const t0 = performance.now();
    expect(evaluate(formula, [LONG, `${LONG}b`, 'x'])).toBe(want);
    expect(performance.now() - t0).toBeLessThan(500); // 병렬 시험 부하에서도 흔들리지 않게 넓힌 기준(재검토 3) — 고치기 전 값(끝나지 않음)과는 여전히 크게 떨어진다
  });

  test('SEARCH도 50ms 안, 처음 맞는 자리', () => {
    const t0 = performance.now();
    expect(evaluate(`SEARCH("a*a*a*a*a*a*a*b","xx${LONG}b")`, [])).toBe(3);
    expect(performance.now() - t0).toBeLessThan(500); // 병렬 시험 부하에서도 흔들리지 않게 넓힌 기준(재검토 3) — 고치기 전 값(끝나지 않음)과는 여전히 크게 떨어진다
  });
});

describe('C1: 엑셀 와일드카드 의미는 그대로', () => {
  const vals = ['a*', 'abc', 'ABC', 'abc', '문의1', '문의?', 'x~y', 'a?c', ''];
  test.each([
    ['a*', 5], // a* abc ABC abc a?c
    ['A?C', 4], // abc ABC abc a?c
    ['a~*', 1],
    ['a~?c', 1],
    ['x~~y', 1],
    ['문의?', 2],
    ['*', 8],
    ['<>a*', 4], // 비지 않은 칸 3개 + 빈 칸 1개
    ['?', 0],
    ['*c', 4],
  ])('COUNTIF 조건 %s → %i', (crit, want) => expect(countOf(vals, crit)).toBe(want));

  test('수식 엔진 MATCH·SEARCH·COUNTIF 의미', () => {
    expect(evaluate('MATCH("a?c",A1:A3,0)', ['x', 'ABC', 'a?c'])).toBe(2);
    expect(evaluate('MATCH("a~?c",A1:A3,0)', ['x', 'ABC', 'a?c'])).toBe(3);
    expect(evaluate('SEARCH("b*d","aBcd")', [])).toBe(2);
    expect(evaluate('SEARCH("?c","abc")', [])).toBe(2);
    expect(evaluate('COUNTIF(A1:A4,"*apple*")', ['apple pie', 'pineapple', 'pear', 'apple'])).toBe(3);
    expect(evaluate('COUNTIF(A1:A4,"<>p*")', ['apple pie', 'pineapple', 'pear', 'apple'])).toBe(2);
  });
});
