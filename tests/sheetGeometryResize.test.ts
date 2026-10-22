import { describe, expect, test } from 'vitest';
import { excelColWidthToPx, pxToExcelColWidth, setRowHeightRange } from '../src/xlsx/sheetGeometry.js';

describe('pxToExcelColWidth', () => {
  test('다시 px로 바꾸면 같은 값', () => {
    for (const mdw of [7, 8, 9]) {
      for (let px = 1; px <= 600; px += 1) expect(excelColWidthToPx(pxToExcelColWidth(px, mdw), mdw)).toBe(px);
    }
    expect(pxToExcelColWidth(0)).toBe(0);
  });
});

describe('setRowHeightRange', () => {
  test('기존 묶음 가운데를 바꾸면 앞뒤로 나뉜다', () => {
    expect(setRowHeightRange([[0, 10, 30]], 3, 2, 50, 20)).toEqual([[0, 3, 30], [3, 2, 50], [5, 5, 30]]);
  });
  test('기본 높이로 되돌리면 묶음에서 빠지고, 이웃과 같은 높이면 합친다', () => {
    expect(setRowHeightRange([[0, 10, 30]], 3, 2, 20, 20)).toEqual([[0, 3, 30], [5, 5, 30]]);
    expect(setRowHeightRange([[0, 3, 30], [5, 5, 30]], 3, 2, 30, 20)).toEqual([[0, 10, 30]]);
  });
  test('빈 묶음·묶음 사이·끝', () => {
    expect(setRowHeightRange([], 4, 1, 0, 20)).toEqual([[4, 1, 0]]);
    expect(setRowHeightRange([[0, 1, 30], [10, 1, 30]], 5, 1, 40, 20)).toEqual([[0, 1, 30], [5, 1, 40], [10, 1, 30]]);
    expect(setRowHeightRange([[0, 1, 30]], 5, 2, 40, 20)).toEqual([[0, 1, 30], [5, 2, 40]]);
  });
});
