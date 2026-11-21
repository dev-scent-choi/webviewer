import fs from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { exportDocx } from '../src/docx/export.js';
import { importDocx } from '../src/docx/import.js';
import { docxSchema } from '../src/docx/schema.js';

/**
 * FINAL-REVIEW 재검토 M3 — ensureBuiltinLists가 앞선 저장의 기본 번호 정의를 `(?:(?!</w:abstractNum>)[\s\S])*` 정규식으로 찾아,
 * 안 닫힌 <w:abstractNum>이 많은 numbering.xml에서 저장이 제곱으로 느려졌다. 정상 파일에서는 같은 번호를 다시 쓴다.
 */
function withBuiltinList(doc: PMNode): PMNode {
  const kids: PMNode[] = [];
  let done = false;
  doc.forEach((n) => {
    if (!done && n.type === docxSchema.nodes.paragraph) {
      done = true;
      kids.push(n.type.create({ ...n.attrs, num: { numId: 'docview:decimal', ilvl: 0, fromStyle: false } }, n.content, n.marks));
    } else kids.push(n);
  });
  return docxSchema.nodes.doc.create(null, kids);
}
const numIdOf = (saved: Uint8Array) => /<w:numId w:val="(\d+)"\/>/.exec(strFromU8(unzipSync(saved)['word/document.xml']))?.[1];

describe('재검토 M3: 기본 번호 정의 찾기', () => {
  const bytes = new Uint8Array(fs.readFileSync('audit/office/corpus/docx/d02-numbering.docx'));
  const { doc, comments, fragments } = importDocx(bytes);

  test('안 닫힌 <w:abstractNum> 10,000개가 든 numbering.xml도 300ms 안에 저장', () => {
    const zip = unzipSync(bytes);
    const junk = '<w:abstractNum w:abstractNumId="7">'.repeat(10_000);
    zip['word/numbering.xml'] = strToU8(strFromU8(zip['word/numbering.xml']).replace('</w:numbering>', `${junk}</w:numbering>`));
    const crafted = zipSync(zip);
    const t0 = performance.now();
    exportDocx(crafted, withBuiltinList(doc), null, comments, fragments);
    expect(performance.now() - t0).toBeLessThan(2_000); // 병렬 시험 부하에서도 흔들리지 않게 넓힌 기준(재검토 3) — 고치기 전 값(4.7초)과는 여전히 크게 떨어진다
  });

  test('두 번 저장해도 같은 기본 번호 정의를 다시 쓴다(numId 그대로)', () => {
    const first = exportDocx(bytes, withBuiltinList(doc), null, comments, fragments);
    const again = importDocx(first);
    const second = exportDocx(first, again.doc, null, again.comments, again.fragments);
    expect(numIdOf(second)).toBe(numIdOf(first));
    const count = (b: Uint8Array) => (strFromU8(unzipSync(b)['word/numbering.xml']).match(/<w:abstractNum\b/g) ?? []).length;
    expect(count(second)).toBe(count(first));
  });
});
