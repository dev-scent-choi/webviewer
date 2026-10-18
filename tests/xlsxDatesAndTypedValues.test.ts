import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { computeGridCellDisplay } from '../src/xlsx/formulas.js';
import { parseXlsxToGrid, type GridDocument } from '../src/spreadsheet.js';
import { coerceCellValue, writeExportCellValue } from '../src/xlsx/collab.js';

function grid(rows: Array<Array<{ v: string } | 0>>): GridDocument {
  return {
    styles: [],
    sheets: [{ name: 'S', rowCount: rows.length, colCount: 4, colWidths: [], defaultColWidth: 64, defaultRowHeight: 20, rowHeights: [], merges: [], freezeRows: 0, freezeCols: 0, rows }],
  };
}

describe('수식: 날짜 텍스트 산술(엑셀과 같은 자동 변환)', () => {
  const g = grid([[{ v: '2026-07-05' }, { v: '2026-07-15' }, { v: '=B1-A1' }, { v: '=A1+7' }]]);

  test('날짜 차이 =B1-A1 → 10일 (예전엔 #VALUE!)', () => {
    expect(computeGridCellDisplay(g, 0, 0, 2)).toBe(10);
  });
  test('날짜 + 일수 =A1+7 → 날짜 일련번호 46215', () => {
    expect(computeGridCellDisplay(g, 0, 0, 3)).toBe(46215);
  });
  test('빈 칸은 산술에서 0 (=A1+B1에서 B1이 비어도 #VALUE!가 아님), 문자열 비교는 그대로', () => {
    const g3 = grid([[{ v: '5' }, 0, { v: '=A1+B1' }, { v: '=IF(B1="","빈칸","값")' }]]);
    expect(computeGridCellDisplay(g3, 0, 0, 2)).toBe(5);
    expect(computeGridCellDisplay(g3, 0, 0, 3)).toBe('빈칸');
  });
  test('날짜가 아닌 텍스트 산술은 여전히 #VALUE! (기존 동작 보존)', () => {
    const g2 = grid([[{ v: '본사' }, { v: '=A1+1' }]]);
    expect(String(computeGridCellDisplay(g2, 0, 0, 1))).toBe('#VALUE!'); // FormulaError 객체 — 표시 시 문자열화
  });
});

describe('뷰어 경로: 날짜 형식 칸의 수식 결과는 날짜로 표시', () => {
  test('=A1+7 (yyyy-mm-dd 형식) → "2026-07-12", 날짜 형식 칸은 df 서식 표시를 가진다', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = new Date(Date.UTC(2026, 6, 5));
    ws.getCell('A1').numFmt = 'yyyy-mm-dd';
    ws.getCell('B1').value = { formula: 'A1+7' };
    ws.getCell('B1').numFmt = 'yyyy-mm-dd';
    ws.getCell('C1').value = { formula: 'B1-A1' };
    const parsed = await parseXlsxToGrid(Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer));
    const row = parsed.sheets[0].rows[0];
    const b1 = row[1] as { v: string; s?: number };
    expect(b1.v).toBe('2026-07-12');
    expect(parsed.styles[b1.s!].df).toBe(1);
    expect((row[2] as { v: string }).v).toBe('7');
  });
});

describe('내보내기: 편집한 칸의 값이 엑셀 타입으로 저장된다', () => {
  test('coerceCellValue: 날짜/쉼표/백분율/원화 → 날짜·숫자, 그 밖은 그대로', () => {
    expect(coerceCellValue('2026-07-05')).toEqual(new Date(Date.UTC(2026, 6, 5)));
    expect(coerceCellValue('1,000')).toBe(1000);
    expect(coerceCellValue('50%')).toBe(0.5);
    expect(coerceCellValue('₩2,500')).toBe(2500);
    expect(coerceCellValue('010-1234-5678')).toBe('010-1234-5678');
    expect(coerceCellValue('2026-02-31')).toBe('2026-02-31');
  });

  test('"일반" 칸에 쓴 날짜·백분율은 엑셀이 자동으로 붙이는 형식까지 붙는다', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    writeExportCellValue(ws.getCell('A1'), '2026-07-05', undefined);
    expect(ws.getCell('A1').value).toEqual(new Date(Date.UTC(2026, 6, 5)));
    expect(ws.getCell('A1').numFmt).toBe('yyyy-mm-dd');
    writeExportCellValue(ws.getCell('A2'), '12.5%', undefined);
    expect(ws.getCell('A2').value).toBe(0.125);
    expect(ws.getCell('A2').numFmt).toBe('0.0%');
    writeExportCellValue(ws.getCell('A3'), '1,000', undefined);
    expect(ws.getCell('A3').numFmt).toBe('#,##0');
  });

  test('이미 서식이 있는 칸은 그 서식을 존중하고, df 칸은 날짜 형식으로 저장된다', () => {
    const ws = new ExcelJS.Workbook().addWorksheet('S');
    writeExportCellValue(ws.getCell('A1'), '1,000', { nf: 'currency' });
    expect(ws.getCell('A1').numFmt).toBe('"₩"#,##0');
    writeExportCellValue(ws.getCell('A2'), '=A1+7', { df: 1 });
    expect(ws.getCell('A2').numFmt).toBe('yyyy-mm-dd');
  });
});
