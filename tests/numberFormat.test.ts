import { describe, expect, test } from 'vitest';
import ExcelJS from 'exceljs';
import { formatNumberCode, formatValueText, hashIfTooWide, numFmtCodeOf } from '../src/collab/numberFormat.js';
import { cellTexts, parseXlsxToGrid } from '../src/spreadsheet.js';

describe('numberFormat — 엑셀 표시 형식(보기·편집 공용)', () => {
  test.each([
    [-4321, '#,##0_);[Red](#,##0)', '(4,321)'],
    [123456789, '0.00E+00', '1.23E+08'],
    [0.75, '# ?/?', ' 3/4'],
    [42, '00000', '00042'],
    [12345678, '#,##0,"천원"', '12,346천원'],
    [1500, '[>=1000]"많음 "#,##0;"적음 "#,##0', '많음 1,500'],
    [9108000, '"일금 "#,##0"원정"', '일금 9,108,000원정'],
    [46346, 'yyyy"년" m"월" d"일"', '2026년 11월 20일'],
    [46346, 'yyyy-mm-dd (aaa)', '2026-11-20 (금)'],
    [46346, '[$-412]yyyy-mm-dd dddd', '2026-11-20 금요일'],
    [0.586909722, 'hh:mm:ss', '14:05:09'],
    [0.586805556, 'AM/PM h:mm', 'PM 2:05'],
    [0.30000000000000004, 'General', '0.3'],
  ])('%s · %s → %s', (value, code, expected) => {
    expect(formatNumberCode(value, code).text).toBe(expected);
  });

  test('색 지정은 값이 고른 구역의 색', () => {
    expect(formatNumberCode(-1, '#,##0;[Red]-#,##0').color).toBe('#FF0000');
    expect(formatNumberCode(1, '#,##0;[Red]-#,##0').color).toBeUndefined();
    expect(formatNumberCode(5, '[Blue][>=1000]#,##0;[Green]#,##0').color).toBe('#00FF00');
  });

  test('깨진 서식 코드는 일반 서식으로', () => {
    expect(formatNumberCode(1.5, '[[[').text).toBe('1.5');
  });

  test('값 글자 + 칸 서식: 날짜 글자·툴바 범주·글자 서식', () => {
    expect(formatValueText('1899-12-30 14:05:09', { fmt: 'hh:mm:ss' }).text).toBe('14:05:09');
    expect(formatValueText('2026-11-20', { df: 1 }).text).toBe('2026-11-20');
    expect(formatValueText('46346', { df: 1 }).text).toBe('2026-11-20');
    expect(formatValueText('1234.5', { nf: 'currency', nd: 2 }).text).toBe('₩1,234.50');
    expect(formatValueText('0.4', { nf: 'percent' }).text).toBe('40%');
    expect(formatValueText('00123', { fmt: '@' }).text).toBe('00123');
    expect(formatValueText('abc', { fmt: '#,##0' }).text).toBe('abc');
    expect(formatValueText('1234.5', undefined).text).toBe('1234.5');
    expect(numFmtCodeOf({ fmt: '0.0', nf: 'number' })).toBe('0.0');
  });
});

describe('cellTexts — 칸의 표시·값 글자', () => {
  test('오류값은 [object Object]가 아니라 #DIV/0!', () => {
    const cell = { value: { formula: '1/0', result: { error: '#DIV/0!' } }, numFmt: 'General', formula: '1/0' } as unknown as ExcelJS.Cell;
    expect(cellTexts(cell)).toEqual({ text: '#DIV/0!', value: '#DIV/0!' });
  });

  test('공유 수식 복제 칸({ sharedFormula, result })도 캐시 결과를 보인다', () => {
    const cell = { value: { sharedFormula: 'F9', result: 900000 }, numFmt: '#,##0', formula: 'D10*E10' } as unknown as ExcelJS.Cell;
    expect(cellTexts(cell)).toEqual({ text: '900,000', value: '900000' });
  });

  test('그리드: 값이 표시와 다르면 n, 글자 숫자는 왼쪽 정렬, 서식 색은 fc', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('s');
    ws.getCell('A1').value = -4321;
    ws.getCell('A1').numFmt = '#,##0;[Red](#,##0)';
    ws.getCell('B1').value = '0.00';
    ws.getCell('C1').value = new Date(Date.UTC(1899, 11, 30, 14, 5, 9));
    ws.getCell('C1').numFmt = 'hh:mm:ss';
    const grid = await parseXlsxToGrid(Buffer.from(await wb.xlsx.writeBuffer()));
    const [a1, b1, c1] = grid.sheets[0].rows[0] as Array<{ v: string; n?: string; s?: number }>;
    expect(a1).toMatchObject({ v: '(4,321)', n: '-4321' });
    expect(grid.styles[a1.s!].fc).toBe('#FF0000');
    expect(grid.styles[b1.s!].a).toBe('l');
    expect(c1).toMatchObject({ v: '14:05:09', n: '1899-12-30 14:05:09' });
  });

  test('ExcelJS가 영어 코드로 바꾼 기본 서식 38번은 한국어 엑셀처럼 -4,321 빨강', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('s');
    ws.getCell('A1').value = -4321;
    ws.getCell('A1').numFmt = '#,##0 ;[Red](#,##0)';
    ws.getColumn(2).width = 3;
    ws.getCell('B1').value = 123456789;
    ws.getCell('B1').numFmt = '#,##0';
    const grid = await parseXlsxToGrid(Buffer.from(await wb.xlsx.writeBuffer()));
    const [a1, b1] = grid.sheets[0].rows[0] as Array<{ v: string; n?: string; s?: number }>;
    expect(a1.v).toBe('-4,321');
    expect(grid.styles[a1.s!].fc).toBe('#FF0000');
    // 좁은 칸의 서식 숫자는 ####, 값은 그대로
    expect(b1).toMatchObject({ n: '123456789' });
    expect(b1.v).toMatch(/^#+$/);
  });
});

describe('hashIfTooWide', () => {
  test('들어가면 그대로, 넘치면 #', () => {
    expect(hashIfTooWide('1,234', 64, 15)).toBe('1,234');
    expect(hashIfTooWide('123,456,789,012', 40, 15)).toBe('####');
  });
});
