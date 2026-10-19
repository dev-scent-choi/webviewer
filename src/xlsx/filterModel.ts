/**
 * 엑셀 자동 필터 모델 — 편집기(숨길 행 계산)와 서버(파일의 필터 읽기·저장)가 함께 쓴다.
 */

/** 필터 조건: 열 번호(시트 기준) → 보이게 둘 값 목록. 목록에 없는 값의 행은 숨긴다. */
export type FilterCriteria = Readonly<Record<string, readonly string[]>>;

export interface SheetFilter {
  /** 머리글 행 포함 범위(0부터) */
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
  readonly criteria: FilterCriteria;
  /** 시트가 아니라 엑셀 표(표 서식)의 필터면 그 표 파트 경로 — 저장 때 시트 자동 필터로 다시 쓰지 않는다 */
  readonly table?: string;
}

export const GRID_FILTERS_MAP = 'gridFilters';

/** 필터로 숨길 행(머리글 아래 데이터 행 중 조건에 안 맞는 것) */
export function filteredRows(filter: SheetFilter, valueAt: (row: number, col: number) => string): number[] {
  const columns = Object.entries(filter.criteria).map(([col, allowed]) => ({ col: Number(col), allowed: new Set(allowed) }));
  if (columns.length === 0) return [];
  const hidden: number[] = [];
  for (let r = filter.r1 + 1; r <= filter.r2; r += 1) {
    if (columns.some(({ col, allowed }) => !allowed.has(valueAt(r, col)))) hidden.push(r);
  }
  return hidden;
}
