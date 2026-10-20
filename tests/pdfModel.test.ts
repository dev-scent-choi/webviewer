import { PDFArray, PDFDocument, PDFName, PDFString } from 'pdf-lib';
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import {
  basePdfForEditor,
  exportPdfFromRoom,
  PdfChangeTracker,
  pdfModelIsEmpty,
  PDF_ANNOTS_MAP,
  seedPdfModel,
} from '../src/pdf/pdfModel.js';
import { readManagedAnnots } from '../src/pdf/annotRead.js';
import type { PdfAnnot } from '../src/pdf/annotTypes.js';

/** 쪽 2개 — 1쪽에 사각형 주석(다루는 종류)과 링크(다루지 않음) */
async function samplePdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const p1 = doc.addPage([300, 400]);
  doc.addPage([300, 400]);
  const ctx = doc.context;
  p1.node.addAnnot(ctx.register(ctx.obj({ Type: 'Annot', Subtype: 'Square', Rect: [10, 10, 100, 100], C: [1, 0, 0], NM: PDFString.of('sq') })));
  p1.node.addAnnot(ctx.register(ctx.obj({ Type: 'Annot', Subtype: 'Link', Rect: [200, 200, 250, 220] })));
  return doc.save();
}

const at = '2026-10-20T01:00:00.000Z';
const ink: PdfAnnot = { id: 'u1', pageKey: 'o:1', type: 'ink', rect: [0, 0, 50, 50], paths: [[0, 0, 50, 50]], color: '#0000FF', opacity: 1, width: 2, author: '시험', createdAt: at, modifiedAt: at };

describe('PDF 방 모델', () => {
  test('방을 열면 파일 주석이 맵에 들어가고, 편집이 없으면 저장 결과는 원본 그대로', async () => {
    const original = await samplePdf();
    const ydoc = new Y.Doc();
    expect(pdfModelIsEmpty(ydoc)).toBe(true);
    await seedPdfModel(original, ydoc);
    expect([...ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP).keys()]).toEqual(['sq']);
    expect(await exportPdfFromRoom(original, ydoc, false)).toBe(original);
  });

  test('편집기에 주는 PDF에서는 다루는 주석이 빠지고 링크는 남는다', async () => {
    const base = await PDFDocument.load(await basePdfForEditor(await samplePdf()));
    expect(readManagedAnnots(base)).toEqual([]);
    expect(base.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)?.size()).toBe(1);
  });

  test('주석을 더하고 원래 주석을 지우면 저장 결과에 그대로 — 링크는 손대지 않는다', async () => {
    const original = await samplePdf();
    const ydoc = new Y.Doc();
    await seedPdfModel(original, ydoc);
    const map = ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP);
    map.set(ink.id, ink);
    map.delete('sq');
    const saved = await PDFDocument.load(await exportPdfFromRoom(original, ydoc, true));
    expect(readManagedAnnots(saved).map((a) => `${a.pageKey}:${a.type}:${a.id}`)).toEqual(['o:1:ink:u1']);
    expect(saved.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)?.size()).toBe(1);
  });

  test('변경 기록 — 시딩은 빼고, 추가·변경·삭제를 열람자와 함께(만들었다 지운 것은 남기지 않음)', async () => {
    const ydoc = new Y.Doc();
    await seedPdfModel(await samplePdf(), ydoc);
    const tracker = new PdfChangeTracker(ydoc, (origin) => (origin === 'conn-a' ? '가' : null));
    const map = ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP);
    ydoc.transact(() => map.set(ink.id, ink), 'conn-a');
    ydoc.transact(() => map.set('tmp', { ...ink, id: 'tmp' }), 'conn-a');
    ydoc.transact(() => map.delete('tmp'), 'conn-a');
    ydoc.transact(() => map.set('sq', { ...map.get('sq')!, color: '#00FF00' }), 'conn-a');
    const list = tracker.list();
    expect(list.map((e) => `${e.kind}:${e.annotType}:${e.page}:${e.viewer}`)).toEqual(['annot_added:ink:2:가', 'annot_changed:rect:1:가']);
    ydoc.transact(() => map.delete('sq'), 'conn-a');
    expect(tracker.list().map((e) => e.kind)).toEqual(['annot_added', 'annot_removed']);
  });
});
