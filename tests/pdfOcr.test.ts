import sharp from 'sharp';
import { PDFDocument, degrees } from 'pdf-lib';
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { exportPdfFromRoom, PdfChangeTracker, PDF_OCR_MAP, PDF_REDACT_MAP, seedPdfModel } from '../src/pdf/pdfModel.js';
import type { OcrWord, PdfOcrPage, PdfRedaction } from '../src/pdf/annotTypes.js';
import { streamBytes } from '../src/pdf/textRewrite.js';

/** 글자가 그림으로만 있는 쪽(스캔본 흉내) */
async function scannedPdf(rotate = 0): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const page = doc.addPage([400, 300]);
  const png = await sharp({ create: { width: 800, height: 600, channels: 3, background: '#f4f2ec' } }).png().toBuffer();
  page.drawImage(await doc.embedPng(png), { x: 0, y: 0, width: 400, height: 300 });
  if (rotate) page.setRotation(degrees(rotate));
  return doc.save();
}

async function items(bytes: Uint8Array): Promise<Array<{ str: string; x: number; y: number; w: number; a: number; b: number }>> {
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise;
  const content = await (await pdf.getPage(1)).getTextContent();
  await pdf.destroy();
  return content.items
    .filter((i) => 'str' in i && i.str.trim())
    .map((i) => {
      const it = i as { str: string; transform: number[]; width: number };
      return { str: it.str, x: it.transform[4], y: it.transform[5], w: it.width, a: it.transform[0], b: it.transform[1] };
    });
}

const at = '2026-10-20T01:00:00.000Z';
const word = (t: string, x: number, y: number, len: number, h = 14): OcrWord => ({ t, x, y, ex: x + len, ey: y, h });

describe('OCR 글자 층(8단계)', () => {
  test('인식한 낱말이 보이지 않는 글자로 들어가고, 위치·폭이 낱말 상자에 맞는다', async () => {
    const original = await scannedPdf();
    const ydoc = new Y.Doc();
    await seedPdfModel(original, ydoc);
    const ocr: PdfOcrPage = { words: [word('분기', 40, 240, 40), word('보고서', 90, 240, 60), word('Revenue', 40, 200, 70)], author: '가', at };
    ydoc.getMap<PdfOcrPage>(PDF_OCR_MAP).set('o:0', ocr);
    const out = await exportPdfFromRoom(original, ydoc, true);
    const found = await items(out);
    expect(found.map((i) => i.str)).toEqual(['분기', '보고서', 'Revenue']);
    const rev = found[2];
    expect(Math.abs(rev.x - 40)).toBeLessThan(0.5);
    expect(Math.abs(rev.y - 200)).toBeLessThan(0.5);
    expect(Math.abs(rev.w - 70)).toBeLessThan(1.5);
    // 렌더 모드 3(보이지 않음)
    const doc = await PDFDocument.load(out);
    const contents = doc.getPage(0).node.Contents()!;
    const all = (contents as unknown as { asArray(): unknown[] }).asArray().map((r) => Buffer.from(streamBytes(doc.context.lookup(r as never))!).toString('latin1')).join('\n');
    expect(all).toMatch(/BT 3 Tr \/DvOcr 14 Tf/);
  });

  test('가린 영역의 낱말은 쓰지 않는다', async () => {
    const original = await scannedPdf();
    const ydoc = new Y.Doc();
    await seedPdfModel(original, ydoc);
    ydoc.getMap<PdfOcrPage>(PDF_OCR_MAP).set('o:0', { words: [word('공개', 40, 240, 40), word('비밀번호', 40, 200, 70)], author: '가', at });
    ydoc.getMap<PdfRedaction>(PDF_REDACT_MAP).set('r', { id: 'r', pageKey: 'o:0', rect: [30, 190, 130, 225], applied: true, author: '가', createdAt: at });
    expect((await items(await exportPdfFromRoom(original, ydoc, true))).map((i) => i.str)).toEqual(['공개']);
  });

  test('회전한 쪽 — 낱말의 기준선 방향대로', async () => {
    const original = await scannedPdf(90);
    const ydoc = new Y.Doc();
    await seedPdfModel(original, ydoc);
    // 화면에서 가로로 읽히는 낱말은 90° 쪽의 사용자 공간에서 세로(위쪽) 방향
    ydoc.getMap<PdfOcrPage>(PDF_OCR_MAP).set('o:0', { words: [{ t: '세로', x: 100, y: 50, ex: 100, ey: 110, h: 14 }], author: '가', at });
    const [it] = await items(await exportPdfFromRoom(original, ydoc, true));
    expect(it.str).toBe('세로');
    expect(Math.abs(it.a)).toBeLessThan(0.01);
    expect(it.b).toBeGreaterThan(1);
  });

  test('변경 기록 — 쪽마다 낱말 수', async () => {
    const ydoc = new Y.Doc();
    await seedPdfModel(await scannedPdf(), ydoc);
    const tracker = new PdfChangeTracker(ydoc, () => '가');
    ydoc.transact(() => ydoc.getMap<PdfOcrPage>(PDF_OCR_MAP).set('o:0', { words: [word('a', 1, 1, 5)], author: '가', at }), 'u');
    expect(tracker.list().map((e) => `${e.kind}:${e.page}:${e.text}`)).toEqual(['ocr_added:1:낱말 1개']);
  });
});
