import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { exportDocx } from '../src/docx/export.js';
import { importDocx } from '../src/docx/import.js';
import { docxSchema } from '../src/docx/schema.js';
import { makeSampleDocx } from './helpers.js';

/** 본문(w:body 안쪽)과 styles.xml 끝부분을 바꿔 끼운 docx */
function docxWith(bodyXml: string, extraStylesXml = ''): Buffer {
  const entries = unzipSync(makeSampleDocx(['x']));
  const documentXml = strFromU8(entries['word/document.xml']);
  entries['word/document.xml'] = strToU8(documentXml.replace(/<w:body>[\s\S]*<w:sectPr/, `<w:body>${bodyXml}<w:sectPr`));
  if (extraStylesXml) {
    entries['word/styles.xml'] = strToU8(strFromU8(entries['word/styles.xml']).replace('</w:styles>', `${extraStylesXml}</w:styles>`));
  }
  return Buffer.from(zipSync(entries));
}

function roundTrip(bytes: Buffer): { documentXml: string; stylesXml: string } {
  const { doc } = importDocx(bytes);
  const out = unzipSync(exportDocx(new Uint8Array(bytes), doc));
  return { documentXml: strFromU8(out['word/document.xml']), stylesXml: strFromU8(out['word/styles.xml']) };
}

describe('docx 서식 왕복(가져오기 → 저장)에서 잃지 않는다', () => {
  const runXml =
    '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="굴림"/><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>a</w:t></w:r>' +
    '<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="FFC000"/><w:vertAlign w:val="subscript"/></w:rPr><w:t>b</w:t></w:r></w:p>';

  test('글꼴(한글·영문 따로), 강조색(이름·임의 색), 위·아래 첨자를 마크로 읽는다', () => {
    const { doc } = importDocx(docxWith(runXml));
    const [a, b] = [doc.child(0).child(0), doc.child(0).child(1)];
    const markOf = (node: typeof a, name: string) => node.marks.find((m) => m.type.name === name);
    expect(markOf(a, 'fontFamily')?.attrs).toEqual({ ascii: 'Arial', eastAsia: '굴림' });
    expect(markOf(a, 'highlight')?.attrs.color).toBe('#ffff00');
    expect(markOf(a, 'superscript')).toBeDefined();
    expect(markOf(b, 'highlight')?.attrs.color).toBe('#ffc000');
    expect(markOf(b, 'subscript')).toBeDefined();
  });

  test('다시 저장하면 rFonts·highlight·shd·vertAlign이 스키마 순서대로 나간다', () => {
    const { documentXml } = roundTrip(docxWith(runXml));
    expect(documentXml).toContain(
      '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" w:eastAsia="굴림"/><w:highlight w:val="yellow"/><w:vertAlign w:val="superscript"/></w:rPr>',
    );
    expect(documentXml).toContain('<w:rPr><w:shd w:val="clear" w:color="auto" w:fill="FFC000"/><w:vertAlign w:val="subscript"/></w:rPr>');
  });

  test('문단 들여쓰기·줄 간격·음영·양쪽 맞춤과 원본 문단 스타일(pStyle)을 지킨다', () => {
    const pXml =
      '<w:p><w:pPr><w:pStyle w:val="MyBody"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' +
      '<w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="800" w:hanging="400"/><w:jc w:val="both"/></w:pPr>' +
      '<w:r><w:t>본문</w:t></w:r></w:p>';
    const { doc } = importDocx(docxWith(pXml, '<w:style w:type="paragraph" w:styleId="MyBody"><w:name w:val="My Body"/></w:style>'));
    expect(doc.child(0).attrs).toMatchObject({
      align: 'justify',
      styleId: 'MyBody',
      shading: '#d9d9d9',
      indent: { left: 800, firstLine: 0, hanging: 400 },
      spacing: { line: 360, lineRule: 'auto', before: 120, after: 240 },
    });
    const { documentXml } = roundTrip(docxWith(pXml, '<w:style w:type="paragraph" w:styleId="MyBody"><w:name w:val="My Body"/></w:style>'));
    expect(documentXml).toContain(
      '<w:pPr><w:pStyle w:val="MyBody"/><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>' +
        '<w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="800" w:hanging="400"/><w:jc w:val="both"/></w:pPr>',
    );
  });
});

describe('스타일 갤러리 문단 스타일', () => {
  const titleDoc = () =>
    docxSchema.node('doc', null, [docxSchema.node('paragraph', { styleId: 'Title' }, [docxSchema.text('제목')])]);

  test('문서에 그 스타일이 없으면 styles.xml에 정의를 추가하고 그 id로 참조한다', () => {
    const bytes = makeSampleDocx(['x']);
    const out = unzipSync(exportDocx(new Uint8Array(bytes), titleDoc()));
    expect(strFromU8(out['word/document.xml'])).toContain('<w:pStyle w:val="Title"/>');
    const styles = strFromU8(out['word/styles.xml']);
    expect(styles).toContain('<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/>');
    expect(styles.trim().endsWith('</w:styles>')).toBe(true);
  });

  test('같은 이름의 스타일이 이미 있으면(한국어 Word의 "a5") 그 id를 쓰고, 읽을 때는 표준 id로 본다', () => {
    const bytes = docxWith(
      '<w:p><w:pPr><w:pStyle w:val="a5"/></w:pPr><w:r><w:t>제목</w:t></w:r></w:p>',
      '<w:style w:type="paragraph" w:styleId="a5"><w:name w:val="Title"/></w:style>',
    );
    const { doc } = importDocx(bytes);
    expect(doc.child(0).attrs.styleId).toBe('Title');
    const out = unzipSync(exportDocx(new Uint8Array(bytes), titleDoc()));
    expect(strFromU8(out['word/document.xml'])).toContain('<w:pStyle w:val="a5"/>');
    expect(strFromU8(out['word/styles.xml'])).not.toContain('w:styleId="Title"');
  });
});

describe('styles.xml이 없는 문서에 갤러리 스타일', () => {
  test('styles.xml을 새로 만들고 콘텐츠 형식·관계까지 등록한다', () => {
    const entries = unzipSync(makeSampleDocx(['x']));
    delete entries['word/styles.xml'];
    entries['word/_rels/document.xml.rels'] = strToU8(
      strFromU8(entries['word/_rels/document.xml.rels']).replace(/<Relationship [^>]*styles[^>]*\/>/, ''),
    );
    const bytes = new Uint8Array(zipSync(entries));
    const doc = docxSchema.node('doc', null, [docxSchema.node('paragraph', { styleId: 'Quote' }, [docxSchema.text('인용')])]);
    const out = unzipSync(exportDocx(bytes, doc));
    expect(strFromU8(out['word/styles.xml'])).toContain('w:styleId="Quote"');
    expect(strFromU8(out['[Content_Types].xml'])).toContain('PartName="/word/styles.xml"');
    expect(strFromU8(out['word/_rels/document.xml.rels'])).toMatch(/relationships\/styles" Target="styles.xml"/);
  });
});
