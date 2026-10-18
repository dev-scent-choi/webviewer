import ExcelJS from 'exceljs';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { buildChartXml, parseChartXml, parseRangeFormula, rangeFormula, seriesFromRange, type GridChart } from '../src/xlsx/chartModel.js';
import { addNewCharts, readSheetCharts, sheetPathsByName } from '../src/xlsx/drawingPassthrough.js';

describe('차트 모델', () => {
  test('범위 수식 읽기/쓰기(시트 이름 따옴표, 단일 칸)', () => {
    expect(parseRangeFormula("'실적 2026'!$B$2:$C$4")).toEqual({ sheetName: '실적 2026', r1: 1, c1: 1, r2: 3, c2: 2 });
    expect(parseRangeFormula('Sheet1!A1')).toEqual({ sheetName: 'Sheet1', r1: 0, c1: 0, r2: 0, c2: 0 });
    expect(rangeFormula("O'Brien", { r1: 0, c1: 0, r2: 2, c2: 0 })).toBe("'O''Brien'!$A$1:$A$3");
  });

  test('선택 범위 → 머리글은 계열 이름, 첫 열 글자는 항목, 계열은 열마다', () => {
    const data = [
      ['분기', '매출', '비용'],
      ['1분기', '12', '5'],
      ['2분기', '17', '8'],
    ];
    const { categories, series } = seriesFromRange({ r1: 0, c1: 0, r2: 2, c2: 2 }, (r, c) => data[r][c]);
    expect(categories).toEqual({ ref: { r1: 1, c1: 0, r2: 2, c2: 0 }, cache: ['1분기', '2분기'] });
    expect(series.map((s) => [s.name, s.values.cache])).toEqual([
      ['매출', [12, 17]],
      ['비용', [5, 8]],
    ]);
  });

  test('만든 차트 XML을 다시 읽으면 종류·제목·참조가 같다', () => {
    const chart: GridChart = {
      id: 'n1',
      sheetName: '실적',
      kind: 'bar',
      title: '분기 매출',
      from: { row: 1, col: 4, dx: 0, dy: 0 },
      to: null,
      size: { width: 480, height: 288 },
      source: 'new',
      categories: { ref: { r1: 1, c1: 0, r2: 2, c2: 0 }, cache: ['1분기', '2분기'] },
      series: [{ name: '매출', nameRef: { r1: 0, c1: 1, r2: 0, c2: 1 }, values: { ref: { r1: 1, c1: 1, r2: 2, c2: 1 }, cache: [12, 17] } }],
    };
    const parsed = parseChartXml(buildChartXml(chart, { categories: ['1분기', '2분기'], series: [{ name: '매출', values: [12, 17] }] }));
    expect(parsed?.kind).toBe('bar');
    expect(parsed?.title).toBe('분기 매출');
    expect(parsed?.series[0].values).toEqual({ ref: { sheetName: '실적', r1: 1, c1: 1, r2: 2, c2: 1 }, cache: [12, 17] });
    expect(parsed?.categories.cache).toEqual(['1분기', '2분기']);
  });

  test('새 차트를 그리기 없는 시트에 붙이고, 파일 차트로 다시 읽힌다', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('실적').addRows([
      ['분기', '매출'],
      ['1분기', 12],
    ]);
    const base = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
    const chart: GridChart = {
      id: 'n2',
      sheetName: '실적',
      kind: 'pie',
      title: '',
      from: { row: 0, col: 3, dx: 0, dy: 0 },
      to: null,
      size: { width: 400, height: 300 },
      source: 'new',
      categories: { ref: { r1: 1, c1: 0, r2: 1, c2: 0 }, cache: ['1분기'] },
      series: [{ name: '매출', nameRef: null, values: { ref: { r1: 1, c1: 1, r2: 1, c2: 1 }, cache: [12] } }],
    };
    const out = addNewCharts(base, [{ chart, cache: { categories: ['1분기'], series: [{ name: '매출', values: [12] }] } }]);
    const entries = unzipSync(out);
    expect(Object.keys(entries)).toEqual(expect.arrayContaining(['xl/drawings/drawing1.xml', 'xl/charts/chart1.xml']));
    expect(strFromU8(entries[sheetPathsByName(entries).get('실적') as string])).toMatch(/<drawing r:id="rIdDv\d+"\/>/);
    const read = readSheetCharts(out);
    expect(read).toHaveLength(1);
    expect(read[0]).toMatchObject({ sheetName: '실적', kind: 'pie', from: { row: 0, col: 3 }, source: 'file' });
    expect(read[0].size?.width).toBeCloseTo(400);
    // ExcelJS로 다시 열어도 깨지지 않는다(값 그대로)
    const check = new ExcelJS.Workbook();
    const { stripDrawings } = await import('../src/xlsx/drawingPassthrough.js');
    await check.xlsx.load(stripDrawings(out).bytes as unknown as ArrayBuffer);
    expect(check.getWorksheet('실적')!.getCell('B2').value).toBe(12);
  });
});
