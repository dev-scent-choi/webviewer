import { PDFDocument, StandardFonts } from 'pdf-lib';
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { exportPdfFromRoom, PdfChangeTracker, PDF_ANNOTS_MAP, PDF_PAGES_ARRAY, PDF_REDACT_MAP, PDF_TEXT_EDITS_MAP, redactedBase, seedPdfModel } from '../src/pdf/pdfModel.js';
import type { PdfAnnot, PdfPageEntry, PdfRedaction, PdfTextEdit } from '../src/pdf/annotTypes.js';

async function texts(bytes: Uint8Array): Promise<string[]> {
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, standardFontDataUrl: 'node_modules/pdfjs-dist/standard_fonts/' }).promise;
  const out: string[] = [];
  for (let p = 1; p <= pdf.numPages; p += 1) out.push((await (await pdf.getPage(p)).getTextContent()).items.map((i) => ('str' in i ? i.str : '')).join('').trim());
  await pdf.destroy();
  return out;
}

async function sample(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  doc.addPage([400, 200]).drawText('Account 9999 secret', { x: 20, y: 100, size: 14, font });
  doc.addPage([400, 200]).drawText('Page two total 50', { x: 20, y: 100, size: 14, font });
  return doc.save();
}
const at = '2026-10-20T01:00:00.000Z';
const W = (font: { widthOfTextAtSize(t: string, s: number): number }, t: string) => [...t].reduce((w, c) => w + font.widthOfTextAtSize(c, 14), 0);

describe('방 모델: 글자 고치기·가리기', async () => {
  const helv = await (await PDFDocument.create()).embedFont(StandardFonts.Helvetica);
  const x = (t: string) => 20 + W(helv, t);

  test('적용한 가리기만 지우고, 고친 글은 쪽 정리 뒤에도 그 쪽에, 가린 곳의 주석은 빠진다', async () => {
    const original = await sample();
    const ydoc = new Y.Doc();
    await seedPdfModel(original, ydoc);
    const redactions = ydoc.getMap<PdfRedaction>(PDF_REDACT_MAP);
    redactions.set('r1', { id: 'r1', pageKey: 'o:0', rect: [x('Account ') + 0.5, 95, x('Account 9999') - 0.5, 115], applied: true, author: '가', createdAt: at });
    redactions.set('r2', { id: 'r2', pageKey: 'o:0', rect: [x('Account 9999 ') + 0.5, 95, x('Account 9999 secret') - 0.5, 115], applied: false, author: '가', createdAt: at });
    ydoc.getMap<PdfTextEdit>(PDF_TEXT_EDITS_MAP).set('e1', { id: 'e1', pageKey: 'o:1', rect: [x('Page two total ') + 0.5, 98, x('Page two total 50') - 0.5, 112], oldText: '50', newText: '05', fontSize: 14, author: '가', createdAt: at });
    const over: PdfAnnot = { id: 'a1', pageKey: 'o:0', type: 'rect', rect: [x('Account ') + 2, 96, x('Account ') + 10, 110], color: '#D93025', opacity: 1, width: 1, author: '가', createdAt: at, modifiedAt: at };
    ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP).set('a1', over);
    // 쪽 순서를 바꿔도 고친 글은 그 쪽을 따라간다
    const pages = ydoc.getArray<PdfPageEntry>(PDF_PAGES_ARRAY);
    ydoc.transact(() => {
      pages.delete(0, 2);
      pages.push([{ key: 'o:1', rotate: 0 }, { key: 'o:0', rotate: 0 }]);
    });
    const out = await exportPdfFromRoom(original, ydoc, true);
    const [first, second] = await texts(out);
    expect(first.replace(/\s+/g, ' ')).toBe('Page two total 05');
    expect(second).toContain('Account');
    expect(second).toContain('secret');
    expect(second).not.toContain('9999');
    const pdf = await getDocument({ data: new Uint8Array(out), isEvalSupported: false }).promise;
    expect((await (await pdf.getPage(2)).getAnnotations()).length).toBe(0);
    await pdf.destroy();
  });

  test('편집기에 줄 원본 — 적용한 가리기만 반영(쪽 순서는 그대로)', async () => {
    const original = await sample();
    const ydoc = new Y.Doc();
    await seedPdfModel(original, ydoc);
    expect(await redactedBase(original, ydoc)).toBe(original);
    ydoc.getMap<PdfRedaction>(PDF_REDACT_MAP).set('r1', { id: 'r1', pageKey: 'o:0', rect: [10, 90, 390, 120], applied: true, author: '가', createdAt: at });
    const base = await texts(await redactedBase(original, ydoc));
    expect(base).toEqual(['', 'Page two total 50']);
  });

  test('변경 기록 — 고친 글(되돌리면 빠짐)·적용한 가리기', async () => {
    const ydoc = new Y.Doc();
    await seedPdfModel(await sample(), ydoc);
    const tracker = new PdfChangeTracker(ydoc, () => '가');
    const edits = ydoc.getMap<PdfTextEdit>(PDF_TEXT_EDITS_MAP);
    const redactions = ydoc.getMap<PdfRedaction>(PDF_REDACT_MAP);
    ydoc.transact(() => edits.set('e1', { id: 'e1', pageKey: 'o:1', rect: [1, 1, 2, 2], oldText: '50', newText: '05', fontSize: 14, author: '가', createdAt: at }), 'u');
    ydoc.transact(() => edits.set('e2', { id: 'e2', pageKey: 'o:1', rect: [1, 1, 2, 2], oldText: 'a', newText: 'b', fontSize: 14, author: '가', createdAt: at }), 'u');
    ydoc.transact(() => edits.delete('e2'), 'u');
    ydoc.transact(() => redactions.set('r1', { id: 'r1', pageKey: 'o:0', rect: [1, 1, 2, 2], applied: false, author: '가', createdAt: at }), 'u');
    expect(tracker.list().map((e) => e.kind)).toEqual(['text_edited']);
    ydoc.transact(() => redactions.set('r1', { id: 'r1', pageKey: 'o:0', rect: [1, 1, 2, 2], applied: true, author: '가', createdAt: at }), 'u');
    expect(tracker.list().map((e) => `${e.kind}:${e.page}:${e.text ?? ''}`)).toEqual(['text_edited:2:50 → 05', 'redaction_applied:1:']);
  });
});
