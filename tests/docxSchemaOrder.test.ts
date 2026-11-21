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
/** CT_SectPr — 머리글·바닥글 참조는 서로 섞여도 된다(같은 순번으로 본다) */
const SECTPR_ORDER = ['headerReference', 'footnotePr', 'endnotePr', 'type', 'pgSz', 'pgMar', 'paperSrc', 'pgBorders', 'lnNumType', 'pgNumType', 'cols',
  'formProt', 'vAlign', 'noEndnote', 'titlePg', 'textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];
const ORDERED: Record<string, readonly string[]> = { ...PR_ORDER, 'w:numbering': NUMBERING_ORDER, 'w:sectPr': SECTPR_ORDER };

function orderViolations(list: readonly XmlNode[], where: string, out: string[]): void {
  for (const node of list) {
    const tag = tagOf(node);
    if (!tag || tag === '#text') continue;
    const order = ORDERED[tag];
    if (order) {
      let last = -1;
      for (const child of childrenOf(node)) {
        const name = (tagOf(child) ?? '').replace(/^w:/, '').replace(/^footerReference$/, 'headerReference');
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

describe('저장본 요소 순서(REVIEW-r2 N2, CT_SectPr는 r3)', () => {
  test.each(files)('%s — 편집 뒤 저장한 document.xml·numbering.xml의 속성 묶음·구역 속성 순서가 스키마대로', (file) => {
    const bytes = new Uint8Array(fs.readFileSync(file));
    const { doc, comments, fragments } = importDocx(bytes);
    // 쪽 설정도 바꿔 저장해 마지막 w:sectPr의 pgSz·pgMar를 내보내기가 다시 쓰게 한다
    const page = { width: 16838, height: 11906, orientation: 'landscape', top: 1000, bottom: 1000, left: 1200, right: 1200 } as const;
    const out = unzipSync(exportDocx(bytes, edited(doc), page, comments, fragments));
    const problems: string[] = [];
    orderViolations(parser.parse(strFromU8(out['word/document.xml'])) as XmlNode[], 'document.xml', problems);
    orderViolations(parser.parse(strFromU8(out['word/numbering.xml'])) as XmlNode[], 'numbering.xml', problems);
    // 원본에 이미 있던 순서 문제(합성 문서 생성기가 w:cols를 docGrid 뒤에 덧붙인 a2·d06)는 우리가 만든 것이 아니다 — 원본보다 늘면 안 된다
    const before: string[] = [];
    orderViolations(parser.parse(strFromU8(unzipSync(bytes)['word/document.xml'])) as XmlNode[], 'document.xml', before);
    expect(problems.filter((p) => !before.includes(p))).toEqual([]);
    expect(problems.length).toBeLessThanOrEqual(before.length);
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

/** 같은 방에서 저장 → 방을 닫았다 저장본에서 다시 열어(가져오기) 편집 없이 저장 — 복제·번호 다시 매기기가 다시 연 방에서도 맞나(REVIEW-r3 C-r3-1) */
describe('다시 열어 저장(REVIEW-r3 C-r3-1)', () => {
  const reopenSave = (saved: Uint8Array) => {
    const again = importDocx(saved);
    return exportDocx(saved, again.doc, null, again.comments, again.fragments);
  };
  const withCopyOf = (doc: PMNode, pick: (n: PMNode) => boolean) => {
    let copy: PMNode | null = null;
    doc.forEach((n) => {
      if (!copy && pick(n)) copy = n;
    });
    const kids: PMNode[] = [];
    doc.forEach((n) => kids.push(n));
    return docxSchema.nodes.doc.create(null, [...kids, copy!]);
  };
  const has = (type: string, attr?: [string, unknown]) => (n: PMNode) => {
    let found = false;
    n.descendants((c) => {
      if (c.type.name === 'ooxml' && (!attr || c.attrs[attr[0]] === attr[1]) && (type === 'any' || c.attrs.kind === type)) found = true;
    });
    return found;
  };

  test('복사한 각주: 다시 연 방에서 저장해도 복제 각주가 남는다(끊긴 참조 0)', () => {
    const bytes = new Uint8Array(fs.readFileSync('audit/office/corpus/docx/d06-columns-footnotes.docx'));
    const { doc, comments, fragments } = importDocx(bytes);
    const first = exportDocx(bytes, withCopyOf(doc, has('note')), null, comments, fragments);
    for (const saved of [first, reopenSave(first), reopenSave(reopenSave(first))]) {
      const z = unzipSync(saved);
      const refs = [...strFromU8(z['word/document.xml']).matchAll(/<w:footnoteReference\b[^>]*w:id="(-?\d+)"/g)].map((m) => m[1]);
      const notes = new Set([...strFromU8(z['word/footnotes.xml']).matchAll(/<w:footnote\b[^>]*w:id="(-?\d+)"/g)].map((m) => m[1]));
      expect(refs.length).toBe(6); // 각주 3개가 든 다단 구역을 통째로 복사
      expect(new Set(refs).size).toBe(6);
      expect(refs.filter((id) => !notes.has(id))).toEqual([]);
    }
  });

  test('복사한 그림: 다시 연 방에서 저장해도 그림 수·그림 번호(겹침 없음) 그대로', () => {
    const bytes = new Uint8Array(fs.readFileSync('audit/office/corpus/docx/d04-images.docx'));
    const { doc, comments, fragments } = importDocx(bytes);
    const first = exportDocx(bytes, withCopyOf(doc, has('image')), null, comments, fragments);
    const ids = (b: Uint8Array) => [...strFromU8(unzipSync(b)['word/document.xml']).matchAll(/<wp:docPr\b[^>]*?\bid="(\d+)"/g)].map((m) => m[1]);
    const again = reopenSave(first);
    expect(ids(first).length).toBe(4);
    expect(ids(again)).toEqual(ids(first));
    expect(new Set(ids(again)).size).toBe(4);
  });

  test('편집기가 켠 번호: 다시 연 방에서 저장해도 numbering.xml 그대로', () => {
    const bytes = new Uint8Array(fs.readFileSync('tests/fixtures/docx-review/I-num-cleanup.docx'));
    const { doc, comments, fragments } = importDocx(bytes);
    const first = exportDocx(bytes, edited(doc), null, comments, fragments);
    const again = reopenSave(reopenSave(first));
    expect(strFromU8(unzipSync(again)['word/numbering.xml'])).toBe(strFromU8(unzipSync(first)['word/numbering.xml']));
    expect(strFromU8(unzipSync(again)['word/document.xml'])).toBe(strFromU8(unzipSync(first)['word/document.xml']));
  });
});
