import { randomUUID } from 'node:crypto';
import fs from 'node:fs';
import { describe, expect, test } from 'vitest';
import { parseXlsxToGrid } from '../src/spreadsheet.js';
import { closeLazyGridStore, getRowCells, ingestXlsxLazily } from '../src/xlsx/lazyGridStore.js';
import { parseTableXml, tableCellLook, withTableLook } from '../src/xlsx/tableStyles.js';

const THEME = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646'];

describe('엑셀 표 스타일(검토 12)', () => {
  const table = parseTableXml('<table ref="A1:C6" totalsRowCount="1"><tableStyleInfo name="TableStyleMedium2" showRowStripes="1" showFirstColumn="1"/></table>')!;
  test('표 파트를 읽는다', () => {
    expect(table).toMatchObject({ r0: 0, r1: 5, c0: 0, c1: 2, headerRows: 1, totalsRows: 1, style: 'TableStyleMedium2', rowStripes: true, firstCol: true });
  });
  test('Medium2: 머리 줄 강조1·흰 굵은 글자, 줄무늬 옅은 강조1, 요약 행 굵게·이중선, 첫 열 굵게', () => {
    expect(tableCellLook(table, 0, 1, THEME)).toMatchObject({ bg: '#4F81BD', col: '#FFFFFF', b: 1 });
    expect(tableCellLook(table, 1, 1, THEME)?.bg).toBe('#DCE6F2'); // 엑셀 tint 0.8(반올림 차이 1)
    expect(tableCellLook(table, 2, 1, THEME)?.bg).toBeUndefined();
    expect(tableCellLook(table, 5, 1, THEME)).toMatchObject({ b: 1, bt: 2, bd: { t: ['double', '#4F81BD'] } });
    expect(tableCellLook(table, 3, 0, THEME)?.b).toBe(1);
    expect(tableCellLook(table, 9, 0, THEME)).toBeUndefined();
  });
  test('칸에 직접 준 서식이 이긴다', () => {
    expect(withTableLook({ bg: '#FF0000' }, { bg: '#4F81BD', col: '#FFFFFF', b: 1 })).toEqual({ bg: '#FF0000', col: '#FFFFFF', b: 1 });
  });
  test('a3 둘째 시트의 표(Medium2): 보기(통째 읽기)·편집(스트리밍 읽기) 둘 다 머리 줄이 칠해진다', async () => {
    const bytes = fs.readFileSync('audit/office/corpus/xlsx/a3-fidelity-grid.xlsx');
    const grid = await parseXlsxToGrid(bytes);
    const sheet = grid.sheets.findIndex((s) => s.name === '둘째');
    const head = grid.sheets[sheet].rows[2][0];
    expect(head !== 0 && head.s !== undefined && grid.styles[head.s].bg).toBe('#4F81BD');
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    try {
      const row = await getRowCells(store, sheet, 2, 2);
      const cell = row[0];
      expect(cell !== 0 && cell.s !== undefined && store.styles[cell.s].bg).toBe('#4F81BD');
    } finally {
      await closeLazyGridStore(store);
    }
  });
});
