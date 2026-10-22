import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { readSheetExtras, readSheetExtrasFromEntries } from '../src/xlsx/sheetExtras.js';
import { readXlsxSkeleton, SKELETON_SHEET_BYTES, skeletonizeSheetXml } from '../src/xlsx/zipSkeleton.js';

async function* pieces(text: string, size: number): AsyncGenerator<Buffer> {
  const buf = Buffer.from(text);
  for (let i = 0; i < buf.length; i += size) yield buf.subarray(i, i + size);
}

const HEAD = '<?xml version="1.0"?><worksheet><sheetViews><sheetView/></sheetViews><cols><col min="1" max="1" width="20"/></cols>';
const TAIL = '<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells><hyperlinks><hyperlink ref="C1" location="표!A1"/></hyperlinks><drawing r:id="rId1"/></worksheet>';
const linkCell = '<c r="D7" t="str"><f>HYPERLINK("#표!A1","처음으로")</f><v>처음으로</v></c>';

function sheet(rows: number, withLinkAt: number): string {
  let data = '<sheetData>';
  for (let r = 1; r <= rows; r += 1) data += `<row r="${r}"><c r="A${r}"><v>${r}</v></c>${r === withLinkAt ? linkCell : ''}</row>`;
  return `${HEAD}${data}</sheetData>${TAIL}`;
}

describe('skeletonizeSheetXml', () => {
  test('데이터는 빼고 앞뒤와 HYPERLINK 칸은 남긴다 — 조각 경계가 어디서 잘려도 같다', async () => {
    const xml = sheet(2000, 1234);
    for (const size of [7, 64, 1000, 4096, 1 << 20]) {
      const out = (await skeletonizeSheetXml(pieces(xml, size))).toString();
      expect(out).toBe(`${HEAD}<sheetData>${linkCell}</sheetData>${TAIL}`);
    }
  });

  test('빈 시트(<sheetData/>)와 sheetData 없는 시트는 그대로', async () => {
    const empty = `${HEAD}<sheetData/>${TAIL}`;
    expect((await skeletonizeSheetXml(pieces(empty, 5))).toString()).toBe(empty);
    const none = `${HEAD}${TAIL}`;
    expect((await skeletonizeSheetXml(pieces(none, 5))).toString()).toBe(none);
  });
});

describe('readXlsxSkeleton', () => {
  const workbook = '<workbook><sheets><sheet name="표" sheetId="1" r:id="rId1"/></sheets></workbook>';
  const rels = '<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';

  test('작은 파일은 unzipSync와 같은 파트를 준다', async () => {
    const zip = zipSync({ 'xl/workbook.xml': strToU8(workbook), 'xl/_rels/workbook.xml.rels': strToU8(rels), 'xl/worksheets/sheet1.xml': strToU8(sheet(50, 7)) });
    const skeleton = await readXlsxSkeleton(zip);
    const full = unzipSync(zip);
    expect(Object.keys(skeleton).sort()).toEqual(Object.keys(full).sort());
    for (const name of Object.keys(full)) expect(strFromU8(skeleton[name])).toBe(strFromU8(full[name]));
  });

  test('큰 시트도 링크·목록 결과는 통째로 읽은 것과 같다', async () => {
    let rows = 1000;
    while (sheet(rows, 0).length < SKELETON_SHEET_BYTES + 1024) rows *= 2;
    const big = sheet(rows, rows - 3);
    const zip = zipSync({ 'xl/workbook.xml': strToU8(workbook), 'xl/_rels/workbook.xml.rels': strToU8(rels), 'xl/worksheets/sheet1.xml': strToU8(big) });
    const skeleton = await readXlsxSkeleton(zip);
    expect(skeleton['xl/worksheets/sheet1.xml'].length).toBeLessThan(2048);
    expect(readSheetExtrasFromEntries(skeleton)).toEqual(readSheetExtras(zip));
    expect(readSheetExtrasFromEntries(skeleton).links.length).toBe(2);
  });
});
