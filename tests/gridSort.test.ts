import { describe, expect, test } from 'vitest';
import { compareCells, filteredRows, shiftFormulaRows, sortOrder } from '../src/client/gridSort.js';

describe('Excel 정렬·필터 계산', () => {
  test('숫자 < 글자, 빈 칸은 오름·내림차순 모두 맨 뒤, 같은 값은 원래 순서', () => {
    expect(sortOrder(['b', '10', '', '2', 'a', '2'], false)).toEqual([3, 5, 1, 4, 0, 2]);
    expect(sortOrder(['b', '10', '', '2', 'a'], true)).toEqual([0, 4, 1, 3, 2]);
    expect(compareCells('가나', '다라', false)).toBeLessThan(0);
  });

  test('상대 행 참조만 옮기고 절대 행·따옴표 안·함수 이름은 그대로', () => {
    expect(shiftFormulaRows('=B2*C2', 3)).toBe('=B5*C5');
    expect(shiftFormulaRows('=SUM($B$2:B4)+A$1', 2)).toBe('=SUM($B$2:B6)+A$1');
    expect(shiftFormulaRows('=LOG10(A2)&"B2"', 1)).toBe('=LOG10(A3)&"B2"');
    expect(shiftFormulaRows("='실적 2026'!B2", -1)).toBe("='실적 2026'!B1");
    expect(shiftFormulaRows('그냥 B2', 5)).toBe('그냥 B2');
  });

  test('필터: 조건에 없는 값의 데이터 행만 숨긴다(머리글 제외)', () => {
    const data = [
      ['지역', '매출'],
      ['서울', '10'],
      ['부산', '20'],
      ['서울', '30'],
    ];
    const filter = { r1: 0, c1: 0, r2: 3, c2: 1, criteria: { '0': ['서울'] } };
    expect(filteredRows(filter, (r, c) => data[r][c])).toEqual([2]);
    expect(filteredRows({ ...filter, criteria: {} }, (r, c) => data[r][c])).toEqual([]);
  });
});

describe('파일의 자동 필터 읽기·조건 쓰기', () => {
  test('범위와 값 목록 조건(빈 값 포함)을 읽는다', async () => {
    const { parseAutoFilter } = await import('../src/xlsx/autoFilter.js');
    const xml =
      '<worksheet><sheetData/><autoFilter ref="B2:D9"><filterColumn colId="1"><filters blank="1"><filter val="서울"/><filter val="A&amp;B"/></filters></filterColumn></autoFilter></worksheet>';
    expect(parseAutoFilter(xml)).toEqual({ r1: 1, c1: 1, r2: 8, c2: 3, criteria: { '2': ['서울', 'A&B', ''] } });
    expect(parseAutoFilter('<worksheet><autoFilter ref="A1:B4"/></worksheet>')).toEqual({ r1: 0, c1: 0, r2: 3, c2: 1, criteria: {} });
  });

  test('ExcelJS가 쓴 범위만 있는 <autoFilter/>에 조건을 붙인다', async () => {
    const ExcelJS = (await import('exceljs')).default;
    const { strFromU8, unzipSync } = await import('fflate');
    const { writeFilterCriteria, parseAutoFilter } = await import('../src/xlsx/autoFilter.js');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('목록');
    ws.addRows([['지역'], ['서울'], ['부산']]);
    ws.autoFilter = 'A1:A3';
    const bytes = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
    const filter = { r1: 0, c1: 0, r2: 2, c2: 0, criteria: { '0': ['서울'] } };
    const out = unzipSync(writeFilterCriteria(bytes, new Map([['목록', filter]])));
    const sheetXml = strFromU8(out['xl/worksheets/sheet1.xml']);
    expect(parseAutoFilter(sheetXml)).toEqual(filter);
  });
});
