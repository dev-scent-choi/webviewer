import { strToU8 } from 'fflate';
import { describe, expect, test } from 'vitest';
import { parseSheetExtras, tableFilterOf } from '../src/xlsx/sheetExtras.js';
import { parseLocation } from '../src/client/gridLinks.js';

const RELS = (items: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items}</Relationships>`;

describe('하이퍼링크(바로가기) 읽기', () => {
  // 엑셀이 실제로 쓰는 모양: 문서 안 링크는 location만(# 없음), 바깥 링크는 r:id → 관계(TargetMode="External")
  const sheet =
    '<worksheet><sheetData>' +
    '<row r="5"><c r="A5"/><c r="B5"><f>HYPERLINK("#자료!A10","10행으로")</f><v>10행으로</v></c></row>' +
    '<row r="6"><c r="A6" t="str"><f>HYPERLINK("javascript:alert(1)","x")</f><v>x</v></c></row>' +
    '</sheetData><hyperlinks>' +
    '<hyperlink ref="A2" location="\'자료 목록\'!A1" display="자료" tooltip="자료로 가기"/>' +
    '<hyperlink ref="A3:A4" r:id="rId1"/>' +
    '<hyperlink ref="A7" r:id="rId2"/>' +
    '</hyperlinks></worksheet>';
  const rels = RELS(
    '<Relationship Id="rId1" Type="…/hyperlink" Target="https://hyanggilab.com/?a=1&amp;b=2" TargetMode="External"/>' +
      '<Relationship Id="rId2" Type="…/hyperlink" Target="javascript:alert(1)" TargetMode="External"/>',
  );
  const { links } = parseSheetExtras('목차', sheet, rels);

  test('문서 안 위치·바깥 주소·범위', () => {
    expect(links).toContainEqual({ sheetName: '목차', r1: 1, c1: 0, r2: 1, c2: 0, location: "'자료 목록'!A1", tooltip: '자료로 가기' });
    expect(links).toContainEqual({ sheetName: '목차', r1: 2, c1: 0, r2: 3, c2: 0, url: 'https://hyanggilab.com/?a=1&b=2' });
  });

  test('HYPERLINK 수식 — 바로 앞 빈 칸(<c r="A5"/>)에 잘못 붙지 않는다', () => {
    expect(links).toContainEqual({ sheetName: '목차', r1: 4, c1: 1, r2: 4, c2: 1, location: '자료!A10' });
    expect(links.some((l) => l.r1 === 4 && l.c1 === 0)).toBe(false);
  });

  test('스크립트 주소는 버린다', () => {
    expect(links.some((l) => l.r1 === 5 || l.r1 === 6)).toBe(false);
  });

  test('편집 화면의 위치 해석', () => {
    expect(parseLocation("'자료 목록'!$B$3")).toEqual({ sheetName: '자료 목록', row: 2, col: 1 });
    expect(parseLocation('#요약!C5')).toEqual({ sheetName: '요약', row: 4, col: 2 });
    expect(parseLocation('AA10')).toEqual({ sheetName: null, row: 9, col: 26 });
    expect(parseLocation('이름있는범위')).toBeNull();
  });
});

describe('목록 상자(데이터 유효성 "목록") 읽기', () => {
  const sheet =
    '<worksheet><dataValidations count="3">' +
    '<dataValidation type="list" allowBlank="1" sqref="E2:E31 G2"><formula1>"김철수,이영희, 박민수"</formula1></dataValidation>' +
    '<dataValidation type="list" sqref="F2:F9"><formula1>목록!$A$1:$A$3</formula1></dataValidation>' +
    '<dataValidation type="list" showDropDown="1" sqref="H2"><formula1>"숨김"</formula1></dataValidation>' +
    '<dataValidation type="whole" sqref="I2"><formula1>1</formula1></dataValidation>' +
    '</dataValidations></worksheet>';
  const { lists } = parseSheetExtras('자료', sheet, undefined);

  test('직접 적은 항목·칸 범위 목록, 드롭다운 숨김과 숫자 제한은 제외', () => {
    expect(lists).toEqual([
      { sheetName: '자료', ranges: [{ r1: 1, c1: 4, r2: 30, c2: 4 }, { r1: 1, c1: 6, r2: 1, c2: 6 }], values: ['김철수', '이영희', '박민수'] },
      { sheetName: '자료', ranges: [{ r1: 1, c1: 5, r2: 8, c2: 5 }], source: { sheetName: '목록', r1: 0, c1: 0, r2: 2, c2: 0 } },
    ]);
  });
});

describe('표(표 서식)의 자동 필터', () => {
  test('시트에 연결된 표의 필터 범위를 표 파트 경로와 함께', () => {
    const entries = {
      'xl/worksheets/sheet4.xml': strToU8('<worksheet><tableParts count="1"><tablePart r:id="rId3"/></tableParts></worksheet>'),
      'xl/worksheets/_rels/sheet4.xml.rels': strToU8(RELS('<Relationship Id="rId3" Type="…/table" Target="../tables/table1.xml"/>')),
      'xl/tables/table1.xml': strToU8('<table ref="B2:D14"><autoFilter ref="B2:D14"/></table>'),
    };
    expect(tableFilterOf(entries, 'xl/worksheets/sheet4.xml')).toEqual({ r1: 1, c1: 1, r2: 13, c2: 3, criteria: {}, table: 'xl/tables/table1.xml' });
  });

  test('필터 단추를 끈 표(autoFilter 없음)는 없음', () => {
    const entries = {
      'xl/worksheets/sheet1.xml': strToU8('<worksheet><tablePart r:id="rId1"/></worksheet>'),
      'xl/worksheets/_rels/sheet1.xml.rels': strToU8(RELS('<Relationship Id="rId1" Type="…/table" Target="../tables/table1.xml"/>')),
      'xl/tables/table1.xml': strToU8('<table ref="A1:B3"/>'),
    };
    expect(tableFilterOf(entries, 'xl/worksheets/sheet1.xml')).toBeNull();
  });
});

describe('자동 필터 범위가 머리글 한 줄뿐일 때', () => {
  test('엑셀처럼 머리글 아래로 빈 줄 없이 이어진 데이터까지 늘린다(서식만 있는 빈 칸·빈 줄 다음은 제외)', async () => {
    const { parseAutoFilter } = await import('../src/xlsx/autoFilter.js');
    const sheet =
      '<worksheet><sheetData>' +
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>' +
      '<row r="2"><c r="A2" t="s"><v>2</v></c></row>' +
      '<row r="3"><c r="B3"><v>5</v></c></row>' +
      '<row r="4"><c r="A4" s="3"/><c r="B4" s="3"/></row>' + // 서식만 있는 빈 줄 — 여기서 끊긴다
      '<row r="5"><c r="A5"><v>9</v></c></row>' +
      '</sheetData><autoFilter ref="A1:B1"/></worksheet>';
    expect(parseAutoFilter(sheet)).toMatchObject({ r1: 0, c1: 0, r2: 2, c2: 1 });
    // 범위가 이미 데이터를 덮고 있으면 그대로
    expect(parseAutoFilter(sheet.replace('ref="A1:B1"', 'ref="A1:B5"'))).toMatchObject({ r2: 4 });
  });
});
