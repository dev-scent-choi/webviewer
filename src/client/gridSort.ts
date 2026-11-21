import { shiftFormulaRefs } from '../collab/condFormatShared.js';
/**
 * 엑셀 "정렬"·"필터"의 순수 계산부 — 화면·Yjs와 무관해 단위 테스트로 검증한다(tests/gridSort.test.ts).
 */

export type SortValue = string;

/** Excel 정렬 순서: 숫자 < 글자(한글·영문 사전순) < (오름차순이든 내림차순이든) 빈 칸은 항상 맨 뒤 */
export function compareCells(a: SortValue, b: SortValue, descending: boolean): number {
  const aEmpty = a.trim() === '';
  const bEmpty = b.trim() === '';
  if (aEmpty || bEmpty) return aEmpty === bEmpty ? 0 : aEmpty ? 1 : -1;
  const an = Number(a);
  const bn = Number(b);
  const aNum = Number.isFinite(an);
  const bNum = Number.isFinite(bn);
  let result: number;
  if (aNum && bNum) result = an - bn;
  else if (aNum !== bNum) result = aNum ? -1 : 1;
  else result = a.localeCompare(b, 'ko', { numeric: true, sensitivity: 'base' });
  return descending ? -result : result;
}

/** 정렬 뒤 각 자리(0..n-1)에 올 원래 행 번호 — 같은 값은 원래 순서를 지킨다(안정 정렬, Excel과 같다). */
export function sortOrder(keys: readonly SortValue[], descending: boolean): number[] {
  return keys
    .map((key, index) => ({ key, index }))
    .sort((x, y) => compareCells(x.key, y.key, descending) || x.index - y.index)
    .map((item) => item.index);
}

/**
 * 수식의 상대 행 참조를 delta만큼 옮긴다(정렬로 행이 움직이면 Excel도 이렇게 한다) — $가 붙은 절대
 * 행(A$1)과 따옴표 안 글자, 함수 이름(LOG10( 등)은 건드리지 않는다.
 */
export function shiftFormulaRows(formula: string, delta: number): string {
  if (!formula.startsWith('=') || delta === 0) return formula;
  return shiftFormulaRefs(formula, delta, 0);
}

export { filteredRows, GRID_FILTERS_MAP, type FilterCriteria, type SheetFilter } from '../xlsx/filterModel.js';
