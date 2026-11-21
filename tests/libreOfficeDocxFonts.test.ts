import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { docxViewLayoutFixes, literalizeDocxThemeFonts } from '../src/converters/libreOffice.js';
import { makeSampleDocx } from './helpers.js';

/**
 * LibreOffice가 이 프로젝트에서 다루는 문서들의 OOXML 테마 폰트 참조(w:eastAsiaTheme 등)를
 * 안정적으로 해석하지 못해, 실제 설치된 서체(예: 맑은 고딕)를 두고도 자체 기본 서체(굴림 등)로
 * 대체해버리는 문제가 있다 — 변환 전에 테마 참조를 word/theme/theme1.xml이 실제로 가리키는
 * 리터럴 서체 이름으로 바꿔 넣어야 한다. 이 파일은 그 순수 XML 변환 로직만 검증한다
 * (실제 soffice 변환은 LibreOffice 설치가 필요해 여기서 다루지 않는다).
 */
describe('literalizeDocxThemeFonts', () => {
  const THEME1_XML =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>' +
    '<a:fontScheme name="Office">' +
    '<a:majorFont><a:latin typeface="Calibri Light"/><a:font script="Hang" typeface="맑은 고딕 제목"/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:font script="Hang" typeface="맑은 고딕"/></a:minorFont>' +
    '</a:fontScheme></a:themeElements></a:theme>';

  function docxWith(stylesXml: string, documentPatch?: (xml: string) => string, includeTheme = true): Buffer {
    const base = makeSampleDocx(['x']);
    const entries = unzipSync(base);
    entries['word/styles.xml'] = strToU8(stylesXml);
    if (includeTheme) entries['word/theme/theme1.xml'] = strToU8(THEME1_XML);
    if (documentPatch) {
      entries['word/document.xml'] = strToU8(documentPatch(strFromU8(entries['word/document.xml'])));
    }
    return Buffer.from(zipSync(entries));
  }

  test('replaces w:asciiTheme/w:eastAsiaTheme in styles.xml docDefaults with the theme minorFont typefaces', () => {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi"/>' +
      '</w:rPr></w:rPrDefault></w:docDefaults></w:styles>';
    const result = literalizeDocxThemeFonts(docxWith(stylesXml));
    const out = strFromU8(unzipSync(result)['word/styles.xml']);
    expect(out).toContain('w:ascii="Calibri"');
    expect(out).toContain('w:eastAsia="맑은 고딕"');
    expect(out).toContain('w:hAnsi="Calibri"');
    expect(out).not.toContain('Theme=');
  });

  test('resolves a "major*" theme reference to the theme majorFont typefaces', () => {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"/>' +
      '</w:rPr></w:rPrDefault></w:docDefaults></w:styles>';
    const out = strFromU8(unzipSync(literalizeDocxThemeFonts(docxWith(stylesXml)))['word/styles.xml']);
    expect(out).toContain('w:ascii="Calibri Light"');
    expect(out).toContain('w:eastAsia="맑은 고딕 제목"');
  });

  test('leaves a literal w:eastAsia untouched but still strips a redundant theme reference', () => {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr>' +
      '<w:rFonts w:eastAsia="바탕체" w:eastAsiaTheme="minorEastAsia"/>' +
      '</w:rPr></w:rPrDefault></w:docDefaults></w:styles>';
    const out = strFromU8(unzipSync(literalizeDocxThemeFonts(docxWith(stylesXml)))['word/styles.xml']);
    expect(out).toContain('w:eastAsia="바탕체"');
    expect(out).not.toContain('eastAsiaTheme');
    expect(out).not.toContain('맑은 고딕"'); // 리터럴이 이미 있으므로 테마 값으로 덮어쓰지 않는다
  });

  test('also literalizes theme font references directly in word/document.xml (per-run rFonts)', () => {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>';
    const docx = docxWith(stylesXml, (xml) =>
      xml.replace('<w:r>', '<w:r><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/></w:rPr>'),
    );
    const out = strFromU8(unzipSync(literalizeDocxThemeFonts(docx))['word/document.xml']);
    expect(out).toContain('w:eastAsia="맑은 고딕"');
  });

  test('returns the original bytes unchanged when the docx has no word/theme/theme1.xml', () => {
    const stylesXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/></w:rPr></w:rPrDefault></w:docDefaults>' +
      '</w:styles>';
    const original = docxWith(stylesXml, undefined, false);
    const result = literalizeDocxThemeFonts(original);
    expect(strFromU8(unzipSync(result)['word/styles.xml'])).toBe(strFromU8(unzipSync(original)['word/styles.xml']));
  });
  test('보기용 사본: 본문 가장자리에 붙은 떠 있는 개체만 "양쪽으로 글 흐름"을 "넓은 쪽에만"으로(보고서 D-v2, 검토 I5)', () => {
    const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/></w:sectPr>';
    const anchor = (posH: string) =>
      `<w:p><w:r><w:drawing><wp:anchor><wp:positionH relativeFrom="column">${posH}</wp:positionH><wp:extent cx="2286000" cy="1143000"/>` +
      '<wp:wrapSquare wrapText="bothSides"/></wp:anchor></w:drawing></w:r></w:p>';
    // 오른쪽 끝 상자(d07: 여백 기준 3200400 + 2286000 ≒ 본문 폭 5760720 - 274320)
    const edge = docxViewLayoutFixes(`<w:body>${anchor('<wp:posOffset>3200400</wp:posOffset>')}${anchor('<wp:align>right</wp:align>')}${sect}</w:body>`);
    expect((edge.match(/wrapText="largest"/g) ?? []).length).toBe(2);
    // 가운데쯤 놓인 그림은 Word도 양쪽으로 흘린다 — 그대로
    const centred = docxViewLayoutFixes(`<w:body>${anchor('<wp:posOffset>1700000</wp:posOffset>')}${anchor('<wp:align>center</wp:align>')}${sect}</w:body>`);
    expect(centred).not.toContain('largest');
  });

  test('보기용 사본: 병합 칸·표 안 표가 없는 표만 tblGrid를 첫 행 칸 너비로(검토 Minor 1)', () => {
    const row = (a: number, b: number) => `<w:tr><w:tc><w:tcPr><w:tcW w:w="${a}" w:type="dxa"/></w:tcPr><w:p/></w:tc><w:tc><w:tcPr><w:tcW w:w="${b}" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr>`;
    const grid = '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>';
    const simple = docxViewLayoutFixes(`<w:tbl>${grid}${row(1134, 7866)}</w:tbl>`);
    expect(simple).toContain('<w:gridCol w:w="1134"/><w:gridCol w:w="7866"/>');
    const merged = `<w:tbl>${grid}<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    expect(docxViewLayoutFixes(merged)).toBe(merged);
    const nested = `<w:tbl>${grid}<w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:tbl>${grid}${row(2000, 2500)}</w:tbl><w:p/></w:tc><w:tc><w:tcPr><w:tcW w:w="8000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>`;
    const fixedNested = docxViewLayoutFixes(nested);
    expect(fixedNested.startsWith(`<w:tbl>${grid}`)).toBe(true); // 바깥 표(안에 표가 있음)는 그대로
    expect(fixedNested).toContain('<w:gridCol w:w="2000"/><w:gridCol w:w="2500"/>'); // 안쪽 단순한 표는 맞춤
  });
});
