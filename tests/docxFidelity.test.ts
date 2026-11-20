import fs from 'node:fs';
import path from 'node:path';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { exportDocx } from '../src/docx/export.js';
import { getDocxRenderExtras, importDocx } from '../src/docx/import.js';
import { NumberingCounter, formatNumber, parseNumberingXml } from '../src/docx/numbering.js';
import { docxSchema } from '../src/docx/schema.js';
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
    const { doc } = importDocx(corpus('d03-tables.docx'));
    const tables = nodesOf(doc, (n) => n.type.name === 'table');
    const first = tables[0];
    // python-docx 표: tblGrid는 균등이지만 칸 너비(tcW)가 실제 값 — 구분 열 1134dxa ≈ 76px
    expect(first.child(0).child(0).attrs.colwidth).toEqual([76]);
    expect(first.child(0).child(0).attrs.borders.top).toMatchObject({ style: 'solid' }); // Table Grid 스타일 테두리
    const grid = tables[2];
    const header = grid.child(0).child(0);
    expect(header.attrs.borders.bottom.widthPx).toBe(3); // firstRow 아래 굵은 선(sz 18)
    expect(header.firstChild!.firstChild!.marks.map((m) => m.type.name)).toContain('strong');
    expect(grid.attrs.tblPrExtra.some((x: string) => x.includes('LightGrid-Accent1'))).toBe(true);
    expect(grid.child(0).attrs.trPrExtra).toEqual(['<w:tblHeader/>']);
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
  test('every corpus document saves without losing preserved parts, numbering, sections or tab stops', () => {
    for (const name of fs.readdirSync('audit/office/corpus/docx').filter((f) => f.endsWith('.docx'))) {
      const bytes = corpus(name);
      const { doc, comments } = importDocx(bytes);
      const saved = exportDocx(bytes, doc, null, comments);
      const before = strFromU8(unzipSync(bytes)['word/document.xml']);
      const after = strFromU8(unzipSync(saved)['word/document.xml']);
      for (const tag of ['w:numPr', 'w:sectPr', 'w:drawing', 'w:footnoteReference', 'w:tblStyle', 'w:tabs', 'w:fldChar', 'w:tblHeader']) {
        const count = (s: string) => (s.match(new RegExp(`<${tag}[\\s/>]`, 'g')) ?? []).length;
        expect(`${name} ${tag} ${count(after)}`).toBe(`${name} ${tag} ${count(before)}`);
      }
      expect(importDocx(saved).doc.textContent).toBe(doc.textContent);
    }
  });

  test('raw XML that was not in the original document (e.g. injected by another collaborator) is not written', () => {
    const bytes = withBody('<w:p><w:r><w:t>본문</w:t></w:r></w:p>');
    const { doc } = importDocx(bytes);
    const injected = docxSchema.nodes.ooxml.create({ xml: '<w:r><w:fldSimple w:instr="INCLUDEPICTURE \\\\evil\\x"/></w:r>', kind: 'hidden' });
    const para = docxSchema.nodes.paragraph.create({ pPrExtra: ['<w:framePr w:w="1"/>'] }, [docxSchema.text('x'), injected]);
    const tampered = doc.copy(doc.content.append(docxSchema.nodes.doc.create(null, [para]).content));
    const xml = strFromU8(unzipSync(exportDocx(bytes, tampered))['word/document.xml']);
    expect(xml).not.toContain('evil');
    expect(xml).not.toContain('framePr');
    expect(xml).toContain('본문');
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
