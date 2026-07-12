import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { parseThemeColors, parseXlsxToGrid } from '../src/spreadsheet.js';

describe('parseThemeColors', () => {
  test('extracts the palette from theme XML with the Excel lt1/dk1 swap', () => {
    const themeXml = `
      <a:clrScheme name="Office">
        <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
        <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
        <a:dk2><a:srgbClr val="44546A"/></a:dk2>
        <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
        <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
        <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
        <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
        <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
        <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
        <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
        <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
        <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
      </a:clrScheme>`;
    const colors = parseThemeColors(themeXml);
    expect(colors[0]).toBe('FFFFFF'); // 인덱스 0 = lt1 (Excel 스왑)
    expect(colors[1]).toBe('000000'); // 인덱스 1 = dk1
    expect(colors[4]).toBe('4472C4'); // 인덱스 4 = accent1
    expect(colors[9]).toBe('70AD47'); // 인덱스 9 = accent6
  });

  test('returns an empty palette when theme XML is missing', () => {
    expect(parseThemeColors(undefined)).toHaveLength(0);
  });
});

describe('xlsx style extraction', () => {
  async function makeStyledXlsx(): Promise<Buffer> {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('S');
    const themed = sheet.getCell('A1');
    themed.value = '테마색';
    themed.font = { color: { theme: 4 } }; // accent1
    themed.fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: 9 } }; // accent6
    const fancy = sheet.getCell('B1');
    fancy.value = '이탤릭';
    fancy.font = { italic: true, underline: true, size: 16 };
    return Buffer.from(await workbook.xlsx.writeBuffer());
  }

  test('resolves theme colors and extended font styles', async () => {
    const grid = await parseXlsxToGrid(await makeStyledXlsx());
    const row = grid.sheets[0].rows[0];
    const themedCell = row[0];
    const fancyCell = row[1];
    expect(themedCell).not.toBe(0);
    expect(fancyCell).not.toBe(0);
    if (themedCell === 0 || fancyCell === 0) return;

    const themedStyle = grid.styles[themedCell.s ?? -1];
    expect(themedStyle.col).toBe('#4F81BD'); // accent1 (exceljs 기본 Office 2007 테마)
    expect(themedStyle.bg).toBe('#F79646'); // accent6

    const fancyStyle = grid.styles[fancyCell.s ?? -1];
    expect(fancyStyle.i).toBe(1);
    expect(fancyStyle.u).toBe(1);
    expect(fancyStyle.fs).toBe(Math.round(16 * (96 / 72)));
  });
});
