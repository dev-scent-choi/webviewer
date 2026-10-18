import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { insertDrawingElement, restoreDrawings, sheetPathsByName, stripDrawings } from '../src/xlsx/drawingPassthrough.js';
import { parseXlsxToGrid } from '../src/spreadsheet.js';

/** ExcelJS로 만든 통합 문서에 openpyxl/Excel처럼 그리기+차트 파트를 직접 붙인다(ExcelJS는 차트를 못 만든다). */
async function workbookWithChart(): Promise<Uint8Array> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('실적');
  ws.addRows([
    ['분기', '매출'],
    ['1분기', 12],
    ['2분기', 17],
  ]);
  wb.addWorksheet('메모');
  const entries: Record<string, Uint8Array> = unzipSync(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));
  const sheetPath = sheetPathsByName(entries).get('실적') as string;
  const relsPath = sheetPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
  entries[sheetPath] = strToU8(insertDrawingElement(strFromU8(entries[sheetPath]), 'rIdD1'));
  const rels = entries[relsPath] ? strFromU8(entries[relsPath]) : '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  entries[relsPath] = strToU8(
    rels.replace('</Relationships>', '<Relationship Id="rIdD1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>'),
  );
  entries['xl/drawings/drawing1.xml'] = strToU8(
    '<?xml version="1.0" encoding="UTF-8"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<xdr:oneCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="4572000" cy="2743200"/>' +
      '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="1" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>' +
      '</xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>',
  );
  entries['xl/drawings/_rels/drawing1.xml.rels'] = strToU8(
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>',
  );
  entries['xl/charts/chart1.xml'] = strToU8('<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>');
  const types = strFromU8(entries['[Content_Types].xml']).replace(
    '</Types>',
    '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
      '<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>',
  );
  entries['[Content_Types].xml'] = strToU8(types);
  return zipSync(entries);
}

describe('xlsx 그리기·차트 보존(ExcelJS를 거치지 않고 되붙이기)', () => {
  test('떼어 낸 파일은 ExcelJS로 읽히고, 다시 쓴 뒤 차트가 원본 그대로 같은 시트에 붙는다', async () => {
    const original = await workbookWithChart();
    const { bytes, stash } = stripDrawings(original);
    expect(Object.keys(unzipSync(bytes)).some((k) => k.startsWith('xl/charts/'))).toBe(false);
    expect(stash.sheets).toEqual([{ sheetName: '실적', drawingPath: 'xl/drawings/drawing1.xml' }]);

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(bytes as unknown as ArrayBuffer);
    wb.getWorksheet('실적')!.getCell('B2').value = 99;
    const written = new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer);
    const restored = unzipSync(restoreDrawings(written, stash));

    expect(strFromU8(restored['xl/charts/chart1.xml'])).toBe(strFromU8(unzipSync(original)['xl/charts/chart1.xml']));
    const sheetPath = sheetPathsByName(restored).get('실적') as string;
    const sheetXml = strFromU8(restored[sheetPath]);
    const rid = /<drawing r:id="([^"]+)"\/>/.exec(sheetXml)?.[1];
    expect(rid).toBeTruthy();
    const rels = strFromU8(restored[sheetPath.replace('worksheets/', 'worksheets/_rels/') + '.rels']);
    expect(rels).toContain(`Id="${rid}"`);
    expect(rels).toContain('Target="../drawings/drawing1.xml"');
    expect(strFromU8(restored['[Content_Types].xml'])).toContain('/xl/charts/chart1.xml');
    expect(strFromU8(restored[sheetPathsByName(restored).get('메모') as string])).not.toContain('<drawing');
    // 편집한 값도 살아 있다
    const check = new ExcelJS.Workbook();
    await check.xlsx.load(stripDrawings(restoreDrawings(written, stash)).bytes as unknown as ArrayBuffer);
    expect(check.getWorksheet('실적')!.getCell('B2').value).toBe(99);
  });

  test('차트가 있는 파일도 그리드로 읽힌다(뷰어·편집기 열기)', async () => {
    const grid = await parseXlsxToGrid(Buffer.from(await workbookWithChart()));
    expect(grid.sheets.map((s) => s.name)).toContain('실적');
  });

  test('<drawing>은 tableParts·extLst보다 앞에 들어간다', () => {
    const xml = '<worksheet xmlns="x" xmlns:r="y"><sheetData/><pageMargins/><tableParts count="1"/><extLst/></worksheet>';
    expect(insertDrawingElement(xml, 'rId9')).toBe('<worksheet xmlns="x" xmlns:r="y"><sheetData/><pageMargins/><drawing r:id="rId9"/><tableParts count="1"/><extLst/></worksheet>');
  });
});

test('숫자 문자 참조로 쓴 시트 이름(&#49892;&#51201;)도 찾는다', () => {
  const entries = {
    'xl/workbook.xml': strToU8('<workbook><sheets><sheet name="&#49892;&#51201;" sheetId="1" r:id="rId1"/></sheets></workbook>'),
    'xl/_rels/workbook.xml.rels': strToU8('<Relationships><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>'),
  };
  expect(sheetPathsByName(entries).get('실적')).toBe('xl/worksheets/sheet1.xml');
});

test('관계 대상이 절대 경로(/xl/worksheets/…)인 파일(openpyxl)도 편집기가 원래 시트 이름으로 읽는다', async () => {
  const { ingestXlsxLazily, closeLazyGridStore } = await import('../src/xlsx/lazyGridStore.js');
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('실적').addRow(['a', 1]);
  const entries: Record<string, Uint8Array> = unzipSync(new Uint8Array((await wb.xlsx.writeBuffer()) as ArrayBuffer));
  entries['xl/_rels/workbook.xml.rels'] = strToU8(strFromU8(entries['xl/_rels/workbook.xml.rels']).replace(/Target="worksheets\//g, 'Target="/xl/worksheets/'));
  const store = await ingestXlsxLazily('abs-target-test', Buffer.from(zipSync(entries)));
  try {
    expect(store.sheetMeta.map((s) => s.name)).toEqual(['실적']);
  } finally {
    await closeLazyGridStore(store);
  }
});
