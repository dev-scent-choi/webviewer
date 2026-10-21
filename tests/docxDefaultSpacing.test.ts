import { strToU8, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { getDocxDefaultParagraphSpacing } from '../src/docx/import.js';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const docx = (styles?: string) =>
  zipSync({
    'word/document.xml': strToU8(`<w:document ${W}><w:body><w:p/></w:body></w:document>`),
    ...(styles ? { 'word/styles.xml': strToU8(`<w:styles ${W}>${styles}</w:styles>`) } : {}),
  });

describe('워드 기본 문단 간격(getDocxDefaultParagraphSpacing)', () => {
  test('워드 기본 서식 — 문서 기본값의 뒤 8pt·1.08줄', () => {
    const d = docx('<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>');
    expect(getDocxDefaultParagraphSpacing(d)).toEqual({ line: 259, lineRule: 'auto', before: 0, after: 160 });
  });

  test('표준 스타일이 문서 기본값보다 우선(속성마다)', () => {
    const d = docx('<w:docDefaults><w:pPrDefault><w:pPr><w:spacing w:before="100" w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="a"><w:pPr><w:spacing w:after="0" w:line="360" w:lineRule="exact"/></w:pPr></w:style>');
    expect(getDocxDefaultParagraphSpacing(d)).toEqual({ line: 360, lineRule: 'exact', before: 100, after: 0 });
  });

  test('styles.xml이 없거나 간격 지정이 없으면 워드처럼 앞 0·뒤 0·1줄', () => {
    const none = { line: 240, lineRule: 'auto', before: 0, after: 0 };
    expect(getDocxDefaultParagraphSpacing(docx())).toEqual(none);
    expect(getDocxDefaultParagraphSpacing(docx('<w:docDefaults/>'))).toEqual(none);
  });
});
