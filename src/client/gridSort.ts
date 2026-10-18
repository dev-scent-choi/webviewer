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
  let out = '';
  let i = 0;
  const pattern = /(?<![A-Za-z0-9_.$])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/y;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch === '"') {
      const end = formula.indexOf('"', i + 1);
      const stop = end < 0 ? formula.length : end + 1;
      out += formula.slice(i, stop);
      i = stop;
      continue;
    }
    if (ch === "'") {
      // 따옴표로 감싼 시트 이름('실적 2026'!A1) — 이름 안은 그대로 둔다
      const end = formula.indexOf("'", i + 1);
      const stop = end < 0 ? formula.length : end + 1;
      out += formula.slice(i, stop);
      i = stop;
      continue;
    }
    pattern.lastIndex = i;
    const m = pattern.exec(formula);
    if (m && (i === 0 || !/[A-Za-z0-9_.$]/.test(formula[i - 1]))) {
      const [, colAbs, col, rowAbs, row] = m;
      const nextRow = rowAbs ? Number(row) : Math.max(1, Number(row) + delta);
      out += `${colAbs}${col}${rowAbs}${nextRow}`;
      i += m[0].length;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}

export { filteredRows, GRID_FILTERS_MAP, type FilterCriteria, type SheetFilter } from '../xlsx/filterModel.js';
