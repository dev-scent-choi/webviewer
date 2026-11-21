import fs from 'node:fs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { exportDocx } from '../src/docx/export.js';
import { importDocx } from '../src/docx/import.js';
import { docxSchema } from '../src/docx/schema.js';

/**
 * FINAL-REVIEW M3 — 각주 복제(복사한 각주 참조 저장)가 Math.max(...전체 id)로 큰 각주 파트에서 RangeError를 내고,
 * 복제마다 파트 전체를 정규식으로 다시 훑던 문제. 각주 20만 개 파트에서도 저장이 되고, 새 id는 가장 큰 id + 1이다.
 */
const withCopiedNotes = (doc: PMNode) => {
  let copy: PMNode | null = null;
  doc.forEach((n) => {
    n.descendants((c) => {
      if (!copy && c.type.name === 'ooxml' && c.attrs.kind === 'note') copy = n;
    });
  });
  const kids: PMNode[] = [];
  doc.forEach((n) => kids.push(n));
  return docxSchema.nodes.doc.create(null, [...kids, copy!]);
};

describe('M3: 큰 각주 파트에서 각주 복제', () => {
  test('각주 20만 개 — RangeError 없이 저장, 복제 id는 가장 큰 id + 1, 끊긴 참조 없음', () => {
    const zip = unzipSync(new Uint8Array(fs.readFileSync('audit/office/corpus/docx/d06-columns-footnotes.docx')));
    const extra = Array.from({ length: 200_000 }, (_, i) => `<w:footnote w:id="${1000 + i}"><w:p/></w:footnote>`).join('');
    zip['word/footnotes.xml'] = strToU8(strFromU8(zip['word/footnotes.xml']).replace('</w:footnotes>', `${extra}</w:footnotes>`));
    const bytes = zipSync(zip);
    const { doc, comments, fragments } = importDocx(bytes);
    const t0 = performance.now();
    const saved = unzipSync(exportDocx(bytes, withCopiedNotes(doc), null, comments, fragments));
    expect(performance.now() - t0).toBeLessThan(10_000);
    const refs = [...strFromU8(saved['word/document.xml']).matchAll(/<w:footnoteReference\b[^>]*w:id="(-?\d+)"/g)].map((m) => Number(m[1]));
    const notes = new Set([...strFromU8(saved['word/footnotes.xml']).matchAll(/<w:footnote\b[^>]*w:id="(-?\d+)"/g)].map((m) => Number(m[1])));
    expect(new Set(refs).size).toBe(refs.length);
    expect(refs.filter((id) => !notes.has(id))).toEqual([]);
    expect(Math.max(...refs)).toBe(1000 + 200_000 - 1 + refs.length / 2); // 복제본은 201000, 201001, …
  }, 60_000);
});
