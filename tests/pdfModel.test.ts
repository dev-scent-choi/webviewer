import { PDFDocument } from 'pdf-lib';
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { exportPdfFromRoom, pdfChangeLogOf, pdfModelIsEmpty, PDF_ANNOTS_MAP } from '../src/pdf/pdfModel.js';

async function samplePdf(pages = 2): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i += 1) doc.addPage([300, 400]);
  return doc.save();
}

describe('PDF 방 모델(1단계)', () => {
  test('편집이 없으면 저장 결과는 원본 바이트 그대로', async () => {
    const original = await samplePdf();
    const ydoc = new Y.Doc();
    expect(pdfModelIsEmpty(ydoc)).toBe(true);
    const out = await exportPdfFromRoom(original, ydoc);
    expect(out).toBe(original);
    expect(pdfChangeLogOf(ydoc)).toEqual([]);
  });
  test('모델에 무언가 있으면 비어 있지 않다(다음 단계가 채운다)', () => {
    const ydoc = new Y.Doc();
    ydoc.getMap(PDF_ANNOTS_MAP).set('a', { type: 'note' });
    expect(pdfModelIsEmpty(ydoc)).toBe(false);
  });
});
