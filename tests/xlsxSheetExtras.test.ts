import { strToU8 } from 'fflate';
import { describe, expect, test } from 'vitest';
import { parseDefinedNames, parseSheetExtras, readSheetExtrasFromEntries, resolveDefinedNameLocation, tableFilterOf } from '../src/xlsx/sheetExtras.js';
import { evalLinkExpr, parseLocation } from '../src/client/gridLinks.js';

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
    expect(parseLocation("'자료 목록'!$B$3")).toEqual({ sheetName: '자료 목록', row: 2, col: 1, row2: 2, col2: 1 });
    expect(parseLocation('#요약!C5')).toEqual({ sheetName: '요약', row: 4, col: 2, row2: 4, col2: 2 });
    expect(parseLocation('AA10')).toEqual({ sheetName: null, row: 9, col: 26, row2: 9, col2: 26 });
    expect(parseLocation('이름있는범위')).toBeNull();
    // 칸 주소로 시작하는 이름(정의된 이름 AB12x)은 칸이 아니다
    expect(parseLocation('AB12x')).toBeNull();
  });

  test('범위 위치 — 범위 전체(뒤집혀 적혀도 왼쪽 위부터)', () => {
    expect(parseLocation('CS!B3:H3')).toEqual({ sheetName: 'CS', row: 2, col: 1, row2: 2, col2: 7 });
    expect(parseLocation("'통신 규약'!$D$9:$B$2")).toEqual({ sheetName: '통신 규약', row: 1, col: 1, row2: 8, col2: 3 });
  });
});

describe('정의된 이름으로 가는 링크', () => {
  const wb =
    '<workbook><sheets><sheet name="요약" sheetId="1" r:id="rId1"/><sheet name="자료" sheetId="2" r:id="rId2"/></sheets>' +
    '<definedNames><definedName name="목표">&apos;자료&apos;!$E$3000</definedName>' +
    '<definedName name="목표" localSheetId="1">자료!$A$1</definedName>' +
    '<definedName name="여럿">자료!$A$1:$B$2,자료!$D$4</definedName>' +
    '<definedName name="세율">0.1</definedName></definedNames></workbook>';
  const names = parseDefinedNames(wb);

  test('이름을 칸 주소로 — 시트 범위 이름이 먼저, 여러 영역이면 첫 영역', () => {
    expect(resolveDefinedNameLocation('목표', '요약', names)).toBe("'자료'!$E$3000");
    expect(resolveDefinedNameLocation('목표', '자료', names)).toBe('자료!$A$1');
    expect(resolveDefinedNameLocation('여럿', '요약', names)).toBe('자료!$A$1:$B$2');
  });

  test('칸 주소·없는 이름은 그대로(화면이 "참조가 올바르지 않습니다"로 알린다)', () => {
    expect(resolveDefinedNameLocation("'통신 규약'!B3", '요약', names)).toBe("'통신 규약'!B3");
    expect(resolveDefinedNameLocation('없는이름', '요약', names)).toBe('없는이름');
    expect(parseLocation(resolveDefinedNameLocation('세율', '요약', names))).toBeNull();
  });

  // 재검토 4 Critical: 링크마다 workbook.xml을 다시 훑어 8,000 링크 × 8,000 이름에 23.6초 걸렸다
  test('8,000 링크 × 8,000 이름 — 200ms 안에, 결과는 같다', () => {
    const N = 8000;
    const defined = Array.from({ length: N }, (_, i) => `<definedName name="이름${i}">자료!$A$${i + 1}</definedName>`).join('');
    const links = Array.from({ length: N }, (_, i) => `<hyperlink ref="A${i + 1}" location="이름${N - 1 - i}"/>`).join('');
    const entries = {
      'xl/workbook.xml': strToU8(`<workbook><sheets><sheet name="요약" sheetId="1" r:id="rId1"/><sheet name="자료" sheetId="2" r:id="rId2"/></sheets><definedNames>${defined}</definedNames></workbook>`),
      'xl/_rels/workbook.xml.rels': strToU8(RELS('<Relationship Id="rId1" Type="…/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="…/worksheet" Target="worksheets/sheet2.xml"/>')),
      'xl/worksheets/sheet1.xml': strToU8(`<worksheet><sheetData/><hyperlinks>${links}</hyperlinks></worksheet>`),
      'xl/worksheets/sheet2.xml': strToU8('<worksheet><sheetData/></worksheet>'),
    };
    const started = performance.now();
    const out = readSheetExtrasFromEntries(entries);
    const ms = performance.now() - started;
    expect(out.links).toHaveLength(N);
    expect(out.links[0].location).toBe(`자료!$A$${N}`);
    expect(out.links[N - 1].location).toBe('자료!$A$1');
    expect(ms).toBeLessThan(200);
  });
});

describe('주소를 칸 값으로 만드는 HYPERLINK 수식 — =IFERROR(HYPERLINK("#\'"&$A2&"\'!a1",">> Go <<"), "")', () => {
  // 엑셀이 아래로 복사한 공유 수식: 첫 칸만 글이 있고 ref가 범위, 나머지 칸은 <f t="shared" si="0"/>
  const formula = 'IFERROR(HYPERLINK("#\'"&amp;$A2&amp;"\'!a1","&gt;&gt; Go &lt;&lt;"), "")';
  const sheet =
    '<worksheet><sheetData>' +
    `<row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2" t="str"><f t="shared" ref="B2:B4" si="0">${formula}</f><v>&gt;&gt; Go &lt;&lt;</v></c></row>` +
    '<row r="3"><c r="A3" t="s"><v>1</v></c><c r="B3" t="str"><f t="shared" si="0"/><v>&gt;&gt; Go &lt;&lt;</v></c></row>' +
    `<row r="6"><c r="C6" t="str"><f>${formula}</f><v>x</v></c></row>` +
    '</sheetData></worksheet>';
  const { links } = parseSheetExtras('목차', sheet, undefined);

  test('IFERROR로 감싸도 찾고, 공유 수식은 복사된 범위 전체를 링크로', () => {
    expect(links).toContainEqual({ sheetName: '목차', r1: 1, c1: 1, r2: 3, c2: 1, expr: `"#'"&$A2&"'!a1"`, anchor: { row: 1, col: 1 } });
    expect(links).toContainEqual({ sheetName: '목차', r1: 5, c1: 2, r2: 5, c2: 2, expr: `"#'"&$A2&"'!a1"`, anchor: { row: 5, col: 2 } });
  });

  test('누른 칸 기준으로 주소를 푼다 — $A 는 열 고정, 행은 따라 내려간다', () => {
    const values: Record<string, string> = { '1:0': '1월 실적', '2:0': "김's 시트" };
    const valueAt = (r: number, c: number) => values[`${r}:${c}`] ?? '';
    const expr = `"#'"&$A2&"'!a1"`;
    expect(evalLinkExpr(expr, { row: 1, col: 1 }, { row: 1, col: 1 }, valueAt)).toBe("#'1월 실적'!a1");
    expect(evalLinkExpr(expr, { row: 1, col: 1 }, { row: 2, col: 1 }, valueAt)).toBe("#'김's 시트'!a1");
    expect(parseLocation("'1월 실적'!a1")).toEqual({ sheetName: '1월 실적', row: 0, col: 0, row2: 0, col2: 0 });
    // 칸 참조·글자·& 말고 다른 것이 섞이면 풀지 않는다
    expect(evalLinkExpr('"#"&TEXT(A1,"0")', { row: 0, col: 0 }, { row: 0, col: 0 }, valueAt)).toBeNull();
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

describe('열 수 없는 바로가기(검토 r3 경미)', () => {
  test('파일·네트워크 경로는 열지 않고 안내용으로 남기고, 스크립트 주소는 버린다', () => {
    const sheet = '<worksheet><sheetData/><hyperlinks><hyperlink ref="A1" r:id="rId1"/><hyperlink ref="A2" r:id="rId2"/><hyperlink ref="A3" r:id="rId3"/></hyperlinks></worksheet>';
    const rels =
      '<Relationships><Relationship Id="rId1" Type="…/hyperlink" Target="file:///C:/자료/계약.pdf" TargetMode="External"/>' +
      '<Relationship Id="rId2" Type="…/hyperlink" Target="\\server\share\a.xlsx" TargetMode="External"/>' +
      '<Relationship Id="rId3" Type="…/hyperlink" Target="vbscript:msgbox(1)" TargetMode="External"/></Relationships>';
    const { links } = parseSheetExtras('s', sheet, rels);
    expect(links).toContainEqual({ sheetName: 's', r1: 0, c1: 0, r2: 0, c2: 0, blocked: 'file:///C:/자료/계약.pdf' });
    expect(links.find((l) => l.r1 === 1)?.blocked).toContain('server');
    expect(links.some((l) => l.r1 === 2)).toBe(false);
    expect(links.every((l) => l.url === undefined)).toBe(true);
  });
});
