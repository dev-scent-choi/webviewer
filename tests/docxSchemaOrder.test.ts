import fs from 'node:fs';
import path from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { exportDocx } from '../src/docx/export.js';
import { importDocx } from '../src/docx/import.js';
import { PR_ORDER, childrenOf, parser, tagOf, type XmlNode } from '../src/docx/ooxml.js';
import { docxSchema } from '../src/docx/schema.js';

/**
 * 저장본의 요소 순서 점검 — Word는 파트 안 자식 순서(ECMA-376 스키마의 sequence)에 엄격해, 순서가 틀리면 "읽을 수 없는 내용"으로 고치거나
 * 열지 못한다. 우리가 쓰는 속성 묶음(w:pPr·w:rPr·w:tcPr·w:trPr·w:tblPr)과 numbering.xml의 순서를 순서표로 확인한다.
 */
const NUMBERING_ORDER = ['numPicBullet', 'abstractNum', 'num', 'numIdMacAtCleanup'];
const ORDERED: Record<string, readonly string[]> = { ...PR_ORDER, 'w:numbering': NUMBERING_ORDER };

function orderViolations(list: readonly XmlNode[], where: string, out: string[]): void {
  for (const node of list) {
    const tag = tagOf(node);
    if (!tag || tag === '#text') continue;
    const order = ORDERED[tag];
    if (order) {
      let last = -1;
      for (const child of childrenOf(node)) {
        const name = (tagOf(child) ?? '').replace(/^w:/, '');
        if (!name || name === '#text') continue;
        const rank = order.indexOf(name);
        if (rank < 0) continue; // 표에 없는(확장) 요소는 판단하지 않는다
        if (rank < last) out.push(`${where}: <${tag}> 안 ${name}이(가) ${order[last]} 뒤에 옴`);
        last = Math.max(last, rank);
      }
    }
    orderViolations(childrenOf(node), where, out);
  }
}

const files = ['audit/office/corpus/docx', 'tests/fixtures/docx-review'].flatMap((d) =>
  fs.readdirSync(d).filter((f) => f.endsWith('.docx')).map((f) => path.join(d, f)),
);

/** 편집기에서 할 만한 편집: 첫 문단 번호 켜기(편집기 기본 번호), 둘째 문단에 갤러리 스타일·정렬·간격·굵게 */
function edited(doc: PMNode): PMNode {
  const children: PMNode[] = [];
  let n = 0;
  doc.forEach((node) => {
    if (node.type === docxSchema.nodes.paragraph && n < 2) {
      n += 1;
      children.push(
        node.type.create(
          n === 1
            ? { ...node.attrs, num: { numId: 'docview:decimal', ilvl: 0, fromStyle: false } }
            : { ...node.attrs, styleId: 'Quote', align: 'center', spacing: { line: 360, lineRule: 'auto', before: 120, after: 0 } },
          node.content,
          node.marks,
        ),
      );
    } else children.push(node);
  });
  return docxSchema.nodes.doc.create(null, children);
}

describe('저장본 요소 순서(REVIEW-r2 N2)', () => {
  test.each(files)('%s — 편집 뒤 저장한 document.xml·numbering.xml의 속성 묶음 순서가 스키마대로', (file) => {
    const bytes = new Uint8Array(fs.readFileSync(file));
    const { doc, comments, fragments } = importDocx(bytes);
    const out = unzipSync(exportDocx(bytes, edited(doc), null, comments, fragments));
    const problems: string[] = [];
    orderViolations(parser.parse(strFromU8(out['word/document.xml'])) as XmlNode[], 'document.xml', problems);
    orderViolations(parser.parse(strFromU8(out['word/numbering.xml'])) as XmlNode[], 'numbering.xml', problems);
    expect(problems).toEqual([]);
  });
});

describe('편집기가 켠 번호 정의(REVIEW-r2 N1)', () => {
  test('다섯 번 이어 저장해도 numbering.xml은 첫 저장 뒤 그대로, numId도 그대로', () => {
    const bytes = new Uint8Array(fs.readFileSync('tests/fixtures/docx-review/I-num-cleanup.docx'));
    const { doc, comments, fragments } = importDocx(bytes);
    const changed = edited(doc);
    let current = exportDocx(bytes, changed, null, comments, fragments);
    const first = strFromU8(unzipSync(current)['word/numbering.xml']);
    for (let i = 0; i < 4; i += 1) current = exportDocx(current, changed, null, comments, fragments);
    const fifth = unzipSync(current);
    expect(strFromU8(fifth['word/numbering.xml'])).toBe(first);
    // numIdMacAtCleanup은 맨 끝에 남는다
    expect(first.trimEnd().endsWith('<w:numIdMacAtCleanup w:val="3"/></w:numbering>')).toBe(true);
    const numId = /<w:num w:numId="(\d+)"><w:abstractNumId w:val="(\d+)"\/><\/w:num>(?![\s\S]*<w:num )/.exec(first);
    expect(strFromU8(fifth['word/document.xml'])).toContain(`<w:numId w:val="${numId?.[1]}"/>`);
  });
});
