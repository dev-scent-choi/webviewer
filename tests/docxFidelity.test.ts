import fs from 'node:fs';
import path from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { exportDocx } from '../src/docx/export.js';
import { getDocxRenderExtras, importDocx } from '../src/docx/import.js';
import { NumberingCounter, formatNumber, parseNumberingXml } from '../src/docx/numbering.js';
import { docxSchema, fontStack, safeFontName } from '../src/docx/schema.js';
import { refHashOf } from '../src/docx/ooxml.js';
import { makeSampleDocx } from './helpers.js';

/** 품질 점검용 합성 문서(audit/office/corpus) — 고객 문서가 아니다 */
const corpus = (name: string) => new Uint8Array(fs.readFileSync(path.join('audit/office/corpus/docx', name)));

function nodesOf(doc: PMNode, predicate: (node: PMNode) => boolean): PMNode[] {
  const out: PMNode[] = [];
  doc.descendants((node) => {
    if (predicate(node)) out.push(node);
  });
  return out;
}

function withBody(body: string, extra: Record<string, string> = {}): Uint8Array {
  const entries = unzipSync(makeSampleDocx([]));
  entries['word/document.xml'] = strToU8(strFromU8(entries['word/document.xml']).replace('<w:sectPr/>', `${body}<w:sectPr/>`));
  for (const [name, xml] of Object.entries(extra)) entries[name] = strToU8(xml);
  return zipSync(entries);
}

describe('numbering (word/numbering.xml)', () => {
  test('formats Korean and Latin numbering styles', () => {
    expect(formatNumber(1, 'ganada')).toBe('가');
    expect(formatNumber(2, 'chosung')).toBe('ㄴ');
    // 14개 다음은 처음으로(Word 출력과는 미검증 — numbering.ts cycle 설명)
    expect([14, 15, 29].map((n) => formatNumber(n, 'ganada'))).toEqual(['하', '가', '가']);
    expect(formatNumber(15, 'chosung')).toBe('ㄱ');
    expect(formatNumber(4, 'upperRoman')).toBe('IV');
    expect(formatNumber(28, 'lowerLetter')).toBe('bb');
    expect(formatNumber(3, 'decimalEnclosedCircle')).toBe('③');
  });

  test('d02: multi-level labels 1. / 1.1. / 1.1.1. / 가) / (ㄱ) restart under their parent', () => {
    const extras = getDocxRenderExtras(corpus('d02-numbering.docx'));
    const { doc } = importDocx(corpus('d02-numbering.docx'));
    const counter = new NumberingCounter(extras.numbering!);
    const labels = nodesOf(doc, (n) => n.isTextblock && !!n.attrs.num).map((n) => counter.next(n.attrs.num.numId, n.attrs.num.ilvl)?.label);
    expect(labels.slice(0, 14)).toEqual(['1.', '1.1.', '1.1.1.', '1.1.2.', '가)', '나)', '(ㄱ)', '(ㄴ)', '1.2.', '2.', '2.1.', '2.2.', '3.', '●']);
    // 스타일(List Number)로 매긴 번호도 번호가 된다
    expect(labels).toContain('3.');
    expect(labels.filter((l) => l === '1.').length).toBeGreaterThanOrEqual(2);
  });

  test('startOverride and lvlRestart=0 are honoured', () => {
    const defs = parseNumberingXml(
      '<w:numbering><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>' +
        '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:lvlRestart w:val="0"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%2)"/></w:lvl></w:abstractNum>' +
        '<w:num w:numId="5"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride></w:num></w:numbering>',
    );
    const c = new NumberingCounter(defs);
    expect([c.next('5', 0)?.label, c.next('5', 1)?.label, c.next('5', 0)?.label, c.next('5', 1)?.label]).toEqual(['7.', 'a)', '8.', 'b)']);
  });
});

describe('docx import: content that used to be dropped', () => {
  test('d04: inline and floating pictures become image nodes with media path, size and float', () => {
    const { doc } = importDocx(corpus('d04-images.docx'));
    const images = nodesOf(doc, (n) => n.type.name === 'ooxml' && n.attrs.kind === 'image');
    expect(images).toHaveLength(3);
    expect(images.every((n) => /^word\/media\/image\d\.png$/.test(n.attrs.media))).toBe(true);
    expect(images.map((n) => n.attrs.float)).toEqual([null, 'right', null]);
    expect(images[0].attrs.widthPx).toBeCloseTo(4320000 / 9525, 0);
  });

  test('d07: text box keeps its text (rendered from PM JSON) and tabs become tab nodes with the paragraph tab stops', () => {
    const { doc } = importDocx(corpus('d07-tabs-textbox-breaks.docx'));
    const box = nodesOf(doc, (n) => n.type.name === 'ooxml' && n.attrs.kind === 'textbox')[0];
    expect(box.attrs.text).toContain('글상자: 주의 사항');
    expect(box.attrs.box.fill).toBe('#fff2cc');
    expect(docxSchema.nodeFromJSON(box.attrs.content[0]).textContent).toBe('글상자: 주의 사항');
    const toc = nodesOf(doc, (n) => n.type.name === 'paragraph' && n.textContent.startsWith('제1장'))[0];
    expect(toc.textContent).toBe('제1장 사업 개요\t1');
    expect(toc.attrs.tabs).toEqual([{ pos: 9071, val: 'right', leader: 'dot' }]);
  });

  test('d06: footnote references get labels, the 2-column section is wrapped in a columns node', () => {
    const { doc } = importDocx(corpus('d06-columns-footnotes.docx'));
    const refs = nodesOf(doc, (n) => n.type.name === 'ooxml' && n.attrs.kind === 'note');
    expect(refs.map((n) => n.attrs.text)).toEqual(['1', '2', '3']);
    const columns = nodesOf(doc, (n) => n.type.name === 'columns');
    expect(columns).toHaveLength(1);
    expect(columns[0].attrs.count).toBe(2);
    expect(columns[0].textContent).toContain('단락 1');
  });

  test('complex fields collapse to one preserved node showing the result; PAGE is recognised', () => {
    const body =
      '<w:p><w:r><w:t xml:space="preserve">- </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>4</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkStart w:id="0" w:name="a"/></w:p>';
    const { doc } = importDocx(withBody(body));
    const field = nodesOf(doc, (n) => n.attrs.kind === 'field')[0];
    expect(field.attrs.field).toBe('PAGE');
    expect(field.attrs.text).toBe('4');
    expect(doc.textContent).toBe('- 4');
  });

  test('table style (Light Grid Accent 1): borders, bold first row/column, tcW column widths', () => {
    const { doc, fragments } = importDocx(corpus('d03-tables.docx'));
    const tables = nodesOf(doc, (n) => n.type.name === 'table');
    const first = tables[0];
    // python-docx 표: tblGrid는 균등이지만 칸 너비(tcW)가 실제 값 — 구분 열 1134dxa ≈ 76px
    expect(first.child(0).child(0).attrs.colwidth).toEqual([76]);
    expect(first.child(0).child(0).attrs.borders.top).toMatchObject({ style: 'solid' }); // Table Grid 스타일 테두리
    const grid = tables[2];
    const header = grid.child(0).child(0);
    expect(header.attrs.borders.bottom.widthPx).toBe(3); // firstRow 아래 굵은 선(sz 18)
    expect(header.firstChild!.firstChild!.marks.map((m) => m.type.name)).toContain('strong');
    expect(grid.attrs.tblPrExtra.some((ref: string) => fragments.get(refHashOf(ref))?.includes('LightGrid-Accent1'))).toBe(true);
    expect(grid.child(0).attrs.trPrExtra.map((ref: string) => fragments.get(refHashOf(ref)))).toEqual(['<w:tblHeader/>']);
  });

  test('Korean heading style ids ("1" named "heading 1") become headings and keep their id', () => {
    const styles =
      '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="1"><w:name w:val="heading 1"/>' +
      '<w:rPr><w:color w:val="2F5496"/><w:sz w:val="32"/></w:rPr></w:style></w:styles>';
    const bytes = withBody('<w:p><w:pPr><w:pStyle w:val="1"/></w:pPr><w:r><w:t>제목</w:t></w:r></w:p>', { 'word/styles.xml': styles });
    const heading = importDocx(bytes).doc.child(0);
    expect(heading.type.name).toBe('heading');
    expect(heading.attrs.styleId).toBe('1');
    const rule = getDocxRenderExtras(bytes).styles.find((r) => r.kind === 'h' && r.key === '1');
    expect(rule).toMatchObject({ color: '#2f5496', sizePt: 16, bold: false });
  });
});

describe('docx export: what is displayed survives saving', () => {
  test('every corpus and review-repro document saves without losing preserved parts, numbering, sections, tracked changes or tab stops', () => {
    const files = ['audit/office/corpus/docx', 'tests/fixtures/docx-review'].flatMap((d) => fs.readdirSync(d).filter((f) => f.endsWith('.docx')).map((f) => path.join(d, f)));
    for (const file of files) {
      const name = path.basename(file);
      const bytes = new Uint8Array(fs.readFileSync(file));
      const { doc, comments } = importDocx(bytes);
      const saved = exportDocx(bytes, doc, null, comments);
      const before = strFromU8(unzipSync(bytes)['word/document.xml']);
      const after = strFromU8(unzipSync(saved)['word/document.xml']);
      for (const tag of ['w:numPr', 'w:sectPr', 'w:drawing', 'w:footnoteReference', 'w:tblStyle', 'w:tabs', 'w:fldChar', 'w:tblHeader', 'w:ins', 'w:del', 'w:hyperlink']) {
        const count = (s: string) => (s.match(new RegExp(`<${tag}[\\s/>]`, 'g')) ?? []).length;
        expect(`${name} ${tag} ${count(after)}`).toBe(`${name} ${tag} ${count(before)}`);
      }
      expect(importDocx(saved).doc.textContent).toBe(doc.textContent);
    }
  });

  test('only fragments from the room store are written — references a collaborator invents (or raw XML) are dropped', () => {
    const bytes = withBody('<w:p><w:r><w:t>본문</w:t></w:r></w:p>');
    const { doc, fragments } = importDocx(bytes);
    const injected = docxSchema.nodes.ooxml.create({ ref: 'w:fldSimple@AAAAAAAAAAAAAAAAAAAA', kind: 'hidden' });
    const para = docxSchema.nodes.paragraph.create({ pPrExtra: ['<w:framePr w:w="1"/>', 'w:framePr@BBBBBBBBBBBBBBBBBBBB'] }, [docxSchema.text('x'), injected]);
    const tampered = doc.copy(doc.content.append(docxSchema.nodes.doc.create(null, [para]).content));
    const xml = strFromU8(unzipSync(exportDocx(bytes, tampered, null, [], fragments))['word/document.xml']);
    expect(xml).not.toContain('fldSimple');
    expect(xml).not.toContain('framePr');
    expect(xml).toContain('본문');
  });

  test('C1: three sections with identical w:sectPr stay three sections after saving without edits', () => {
    const bytes = new Uint8Array(fs.readFileSync('tests/fixtures/docx-review/A-same-sections.docx'));
    const count = (b: Uint8Array) => (strFromU8(unzipSync(b)['word/document.xml']).match(/<w:sectPr[\s>]/g) ?? []).length;
    const { doc, comments } = importDocx(bytes);
    expect(count(exportDocx(bytes, doc, null, comments))).toBe(count(bytes));
    expect(count(bytes)).toBe(3);
  });

  test('C3: a picture deleted, saved, then restored by undo is saved again — the room keeps the fragment store from when it opened', () => {
    const bytes = corpus('d04-images.docx');
    const { doc, fragments } = importDocx(bytes);
    const drawings = (b: Uint8Array) => (strFromU8(unzipSync(b)['word/document.xml']).match(/<w:drawing>/g) ?? []).length;
    // 그림이 든 첫 문단을 지우고 저장(자동저장) → 디스크의 원본엔 그 그림 조각이 없다
    let imagePara = -1;
    doc.forEach((n, _, i) => {
      if (imagePara < 0 && nodesOf(n, (c) => c.attrs.kind === 'image').length > 0) imagePara = i;
    });
    const without = docxSchema.nodes.doc.create(null, Array.from({ length: doc.childCount }, (_, i) => doc.child(i)).filter((_, i) => i !== imagePara));
    const savedOnce = exportDocx(bytes, without, null, [], fragments);
    expect(drawings(savedOnce)).toBe(2);
    // 되돌리기로 문서가 원래대로 → 디스크의 직전 저장본을 바탕으로 저장해도 방의 저장소 덕분에 그림 3개
    expect(drawings(exportDocx(savedOnce, doc, null, [], fragments))).toBe(3);
  });

  test('C2: copying a picture/bookmark/footnote reference inside the same document — drawing ids renumbered, bookmarks and note references written once', () => {
    const bytes = corpus('d06-columns-footnotes.docx');
    const { doc, fragments } = importDocx(bytes);
    const note = nodesOf(doc, (n) => n.attrs.kind === 'note')[0];
    const imgBytes = corpus('d04-images.docx');
    const imgDoc = importDocx(imgBytes);
    const img = nodesOf(imgDoc.doc, (n) => n.attrs.kind === 'image')[0];
    const twiceNote = docxSchema.nodes.paragraph.create(null, [docxSchema.text('복사 '), note, docxSchema.text(' '), note]);
    const withCopy = doc.copy(doc.content.append(docxSchema.nodes.doc.create(null, [twiceNote]).content));
    const xml = strFromU8(unzipSync(exportDocx(bytes, withCopy, null, [], fragments))['word/document.xml']);
    expect((xml.match(/<w:footnoteReference /g) ?? []).length).toBe(5); // 원래 3개 + 복사 2개 — 복사본은 각주를 복제해 새 번호로(N3)
    const twiceImg = docxSchema.nodes.paragraph.create(null, [img, img]);
    const imgSaved = strFromU8(unzipSync(exportDocx(imgBytes, imgDoc.doc.copy(imgDoc.doc.content.append(docxSchema.nodes.doc.create(null, [twiceImg]).content)), null, [], imgDoc.fragments))['word/document.xml']);
    const ids = [...imgSaved.matchAll(/<wp:docPr\b[^>]*?\bid="(\d+)"/g)].map((m) => m[1]);
    expect(ids).toHaveLength(5);
    expect(new Set(ids).size).toBe(5);
  });

  test('I2: tracked deletion/insertion/move fixture keeps every w:ins·w:del on save', () => {
    const bytes = new Uint8Array(fs.readFileSync('tests/fixtures/docx-review/B-tracked.docx'));
    const { doc, comments } = importDocx(bytes);
    const before = strFromU8(unzipSync(bytes)['word/document.xml']);
    const after = strFromU8(unzipSync(exportDocx(bytes, doc, null, comments))['word/document.xml']);
    for (const tag of ['w:ins', 'w:del', 'w:moveFrom', 'w:moveTo']) {
      const count = (s: string) => (s.match(new RegExp(`<${tag}[\\s>]`, 'g')) ?? []).length;
      expect(`${tag} ${count(after)}`).toBe(`${tag} ${count(before)}`);
    }
  });

  test('I3: a number the editor turned on (docview:decimal) is saved as a real Word list — numbering.xml definition + w:numPr', () => {
    const bytes = withBody('<w:p><w:r><w:t>하나</w:t></w:r></w:p>');
    const { doc } = importDocx(bytes);
    const numbered = docxSchema.nodes.doc.create(null, [docxSchema.nodes.paragraph.create({ num: { numId: 'docview:decimal', ilvl: 0, fromStyle: false } }, docxSchema.text('하나'))]);
    void doc;
    const out = unzipSync(exportDocx(bytes, numbered));
    const numbering = strFromU8(out['word/numbering.xml']);
    const numId = /<w:num w:numId="(\d+)">/.exec(numbering)?.[1];
    expect(numId).toBeDefined();
    expect(strFromU8(out['word/document.xml'])).toContain(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr>`);
    expect(strFromU8(out['word/document.xml'])).not.toContain('1. ');
    expect(strFromU8(out['[Content_Types].xml'])).toContain('/word/numbering.xml');
    expect(strFromU8(out['word/_rels/document.xml.rels'])).toContain('Target="numbering.xml"');
  });

  test('a split section-break paragraph writes its w:sectPr only once (on the last half)', () => {
    const { doc } = importDocx(corpus('d05-header-footer-sections.docx'));
    let holder: PMNode | null = null;
    doc.forEach((n) => {
      if (n.attrs.sect !== null && n.attrs.sect !== undefined) holder ??= n;
    });
    const copy = docxSchema.nodes.paragraph.create(holder!.attrs, docxSchema.text('나뉜 앞쪽'));
    const children: PMNode[] = [];
    doc.forEach((n) => {
      if (n === holder) children.push(copy);
      children.push(n);
    });
    const split = docxSchema.nodes.doc.create(null, children);
    const bytes = corpus('d05-header-footer-sections.docx');
    const xml = strFromU8(unzipSync(exportDocx(bytes, split))['word/document.xml']);
    expect((xml.match(/<w:sectPr[\s>]/g) ?? []).length).toBe(2);
  });
});

describe('numbering counter: Word rules (REVIEW I6, Minor 9)', () => {
  test('numIds sharing an abstractNum continue counting; a startOverride restarts on first use', () => {
    const lvl = '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>';
    const defs = parseNumberingXml(
      `<w:numbering><w:abstractNum w:abstractNumId="0">${lvl}</w:abstractNum>` +
        '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>' +
        '<w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num></w:numbering>',
    );
    const c = new NumberingCounter(defs);
    expect(['1', '1', '2', '2', '3', '3'].map((id) => c.next(id, 0)?.label)).toEqual(['1.', '2.', '3.', '4.', '1.', '2.']);
  });

  test('koreanLegal is Sino-Korean hangul numerals', () => {
    expect([1, 10, 11, 20, 105].map((n) => formatNumber(n, 'koreanLegal'))).toEqual(['일', '십', '십일', '이십', '백오']);
  });
});

describe('style sheet safety (REVIEW I1)', () => {
  test('a font name from styles.xml cannot break out of the CSS string (newline, braces, quotes, url)', () => {
    const extras = getDocxRenderExtras(new Uint8Array(fs.readFileSync('tests/fixtures/docx-review/E-css.docx')));
    const fonts = extras.styles.flatMap((r) => (r.font ? [fontStack([r.font.eastAsia, r.font.ascii])] : []));
    expect(fonts.length).toBeGreaterThan(0);
    for (const f of fonts) expect(f).not.toMatch(/[{}\n@;<>\\]|url\(/);
    expect(safeFontName('X\n} p { color: red } x{')).toBe('X p color red x');
    expect(safeFontName('맑은 고딕')).toBe('맑은 고딕');
  });
});

describe('render extras: headers, footers, footnotes', () => {
  test('d05: sections carry first-page and default headers and a footer with PAGE / NUMPAGES fields', () => {
    const extras = getDocxRenderExtras(corpus('d05-header-footer-sections.docx'));
    expect(extras.sections).toHaveLength(2);
    const [first, last] = extras.sections;
    expect(first.titlePg).toBe(true);
    const text = (blocks: unknown[] | null) => (blocks ?? []).map((b) => docxSchema.nodeFromJSON(b).textContent).join('');
    expect(text(first.header.first)).toContain('표지 머리글');
    expect(text(first.header.default)).toContain('대외비');
    // 마지막(가로) 구역은 머리글 참조가 없어 앞 구역 것을 물려받는다
    expect(text(last.footer.default)).toBe(text(first.footer.default));
    const fields = nodesOf(docxSchema.nodeFromJSON(first.footer.default![0]), (n) => n.attrs.kind === 'field').map((n) => n.attrs.field);
    expect(fields).toEqual(['PAGE', 'NUMPAGES']);
  });

  test('d06: footnote bodies are keyed by id with their visible label', () => {
    const extras = getDocxRenderExtras(corpus('d06-columns-footnotes.docx'));
    const note = docxSchema.nodeFromJSON(extras.notes.footnote['1'][0]);
    expect(note.textContent).toBe('1 각주 1: 관련 법령 제1조 참조. 설비 점검 기준 고시.');
    expect(Object.keys(extras.notes.footnote)).toEqual(['1', '2', '3']);
  });
});

describe('N3: 같은 문서 안에서 복사한 각주', () => {
  test('복사한 각주 참조는 새 각주(본문 복제)를 가리키고, 저장을 거듭해도 footnotes.xml이 늘지 않으며, 복사본을 지우면 복제 각주도 지운다', () => {
    const bytes = corpus('d06-columns-footnotes.docx');
    const { doc, fragments } = importDocx(bytes);
    const note = nodesOf(doc, (n) => n.attrs.kind === 'note')[0];
    const withCopies = doc.copy(doc.content.append(docxSchema.nodes.doc.create(null, [docxSchema.nodes.paragraph.create(null, [docxSchema.text('복사 '), note, docxSchema.text(' '), note])]).content));
    const save = (from: Uint8Array, d: PMNode) => exportDocx(from, d, null, [], fragments);
    const first = save(bytes, withCopies);
    const refs = (b: Uint8Array) => [...strFromU8(unzipSync(b)['word/document.xml']).matchAll(/<w:footnoteReference w:id="(\d+)"/g)].map((m) => m[1]);
    expect(refs(first)).toHaveLength(5);
    expect(new Set(refs(first)).size).toBe(5);
    const notes = strFromU8(unzipSync(first)['word/footnotes.xml']);
    for (const id of refs(first)) expect(notes).toContain(`<w:footnote w:id="${id}">`);
    expect((notes.match(/각주 1: 관련 법령/g) ?? []).length).toBe(3);
    const third = save(save(first, withCopies), withCopies);
    expect(strFromU8(unzipSync(third)['word/footnotes.xml'])).toBe(notes);
    expect(refs(third)).toEqual(refs(first));
    // 복사본을 지우고 저장 → 복제 각주도 사라진다
    const cleaned = strFromU8(unzipSync(save(third, doc))['word/footnotes.xml']);
    expect((cleaned.match(/각주 1: 관련 법령/g) ?? []).length).toBe(1);
  });
});

describe('3차 검토(m4·Minor 5)', () => {
  const tracked = '<w:p><w:ins w:id="7" w:author="원래 작성자" w:date="2026-01-01T00:00:00Z"><w:r><w:t>추적 삽입</w:t></w:r></w:ins></w:p>';

  test('m4: 원본에 있던 변경 추적 삽입 감싸개만 저장 — 지어낸 작성자는 빼고 글자만', () => {
    const original = withBody(tracked);
    const { doc, fragments } = importDocx(original);
    const text = nodesOf(doc, (n) => n.isText)[0];
    const wrap = text.marks.find((m) => m.type.name === 'wrap')!;
    // 그대로 저장하면 감싸개가 남는다
    const kept = strFromU8(unzipSync(exportDocx(original, doc, null, [], fragments))['word/document.xml']);
    expect(kept).toContain('w:author="원래 작성자"');
    // 다른 참여자가 작성자를 바꾼 감싸개(Yjs로 온 값)는 저장하지 않는다 — 글자는 남는다
    const forgedMark = wrap.type.create({ ...wrap.attrs, attrs: { ...(wrap.attrs.attrs as object), 'w:author': '가짜 작성자' } });
    const forged = doc.copy(doc.content.replaceChild(0, doc.firstChild!.copy(doc.firstChild!.content.replaceChild(0, text.mark([forgedMark])))));
    const xml = strFromU8(unzipSync(exportDocx(original, forged, null, [], fragments))['word/document.xml']);
    expect(xml).not.toContain('가짜 작성자');
    expect(xml).not.toContain('<w:ins');
    expect(xml).toContain('추적 삽입');
  });

  test('Minor 5: 글꼴별 1줄 — 표에 있는 글꼴 런은 자기 1줄 × 문단 배수, 고정 줄 간격은 px 그대로', () => {
    const span = docxSchema.marks.fontFamily.create({ eastAsia: '바탕', ascii: null });
    const spec = docxSchema.marks.fontFamily.spec.toDOM!(span, true) as [string, Record<string, string>];
    expect(spec[1].class).toBe('docx-font-lh');
    expect(spec[1].style).toContain('--doc-single-line: 1.3');
    const unknown = docxSchema.marks.fontFamily.spec.toDOM!(docxSchema.marks.fontFamily.create({ eastAsia: '나눔고딕', ascii: null }), true) as [string, Record<string, string>];
    expect(unknown[1].class).toBeUndefined();
    const para = (spacing: object) => docxSchema.nodes.paragraph.spec.toDOM!(docxSchema.nodes.paragraph.create({ spacing: { before: null, after: null, ...spacing } }, docxSchema.text('가', [span]))) as [string, Record<string, string>];
    expect(para({ line: 360, lineRule: 'auto' })[1].style).toContain('--line-mult: 1.5');
    expect(para({ line: 360, lineRule: 'auto' })[1].style).toContain('--doc-single-line: 1.3');
    expect(para({ line: 300, lineRule: 'exact' })[1].style).toContain('--run-line-height: 20px');
  });
});

describe('Word 1줄 = 1.3 × win 높이(REVIEW-r3 I-r3-1)', () => {
  test('깔린 글꼴에서 잰 값: 맑은 고딕 1.729, 바탕·굴림·돋움·궁서 1.3, 라틴 글꼴(Arial)은 표에 없음', async () => {
    const { lineRatiosOfFont } = await import('../src/docx/fontMetrics.js');
    const read = (name: string) => (fs.existsSync(`C:/Windows/Fonts/${name}`) ? lineRatiosOfFont(fs.readFileSync(`C:/Windows/Fonts/${name}`)) : null);
    const malgun = read('malgun.ttf');
    if (malgun) expect(malgun['맑은 고딕']).toBe(1.729);
    const batang = read('batang.ttc');
    if (batang) for (const name of ['바탕', '바탕체', '궁서', '궁서체']) expect(batang[name]).toBe(1.3);
    const gulim = read('gulim.ttc');
    if (gulim) for (const name of ['굴림', '굴림체', '돋움', '돋움체']) expect(gulim[name]).toBe(1.3);
    const arial = read('arial.ttf');
    if (arial) expect(arial.Arial).toBeUndefined();
  });

  test('d01 Office 정답 재현: 바탕 11pt·1.15줄·뒤 10pt 문단의 줄 간격 26.5pt', () => {
    const d01 = corpus('d01-styles-fonts.docx');
    const styles = strFromU8(unzipSync(d01)['word/styles.xml']);
    // 문서 기본 문단: 1.15줄(276)·뒤 간격 200(10pt), 11pt(22 half-points)
    expect(styles).toContain('<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>');
    expect(styles).toContain('w:sz w:val="22"');
    const ratio = 1.3; // 기본 표 = fontMetrics 식(batang.ttc (879+145)/1024 × 1.3)
    const pitchPt = 11 * ratio * 1.15 + 10;
    expect(Math.abs(pitchPt - 26.5)).toBeLessThan(0.2);
    const span = docxSchema.marks.fontFamily.spec.toDOM!(docxSchema.marks.fontFamily.create({ eastAsia: '바탕', ascii: null }), true) as [string, Record<string, string>];
    expect(span[1].style).toContain(`--doc-single-line: ${ratio}`);
  });
});

describe('그림 경로 상한(REVIEW-r3 Minor)', () => {
  test('새 묶음을 그려야 할 때만 상한을 묻고, 넘었으면 LibreOffice를 부르지 않고 DrawingRateLimited', async () => {
    const { renderDocxDrawingPng, DrawingRateLimited } = await import('../src/docx/drawingImage.js');
    const bytes = new Uint8Array(fs.readFileSync('tests/fixtures/docx-review/H-emf-chart.docx'));
    const { serverDrawn } = importDocx(bytes);
    const hash = [...serverDrawn][0];
    let asked = 0;
    await expect(renderDocxDrawingPng(async () => bytes, 'rate-test', `w:drawing@${hash}`, null, async () => ((asked += 1), false))).rejects.toBeInstanceOf(DrawingRateLimited);
    expect(asked).toBe(1);
    // 이 문서에 없는 조각은 상한을 묻지도 않고 null
    expect(await renderDocxDrawingPng(async () => bytes, 'rate-test', 'w:drawing@AAAAAAAAAAAAAAAAAAAA', null, async () => ((asked += 1), false))).toBeNull();
    expect(asked).toBe(1);
  });
});
