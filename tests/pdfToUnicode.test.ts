import { PDFDocument, StandardFonts } from 'pdf-lib';
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { exportPdfFromRoom, PDF_OCR_MAP, PDF_TEXT_EDITS_MAP, seedPdfModel } from '../src/pdf/pdfModel.js';
import type { PdfOcrPage, PdfTextEdit } from '../src/pdf/annotTypes.js';

async function texts(bytes: Uint8Array): Promise<string[]> {
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, standardFontDataUrl: 'node_modules/pdfjs-dist/standard_fonts/' }).promise;
  const out = (await (await pdf.getPage(1)).getTextContent()).items.map((i) => ('str' in i ? i.str : '')).filter((s) => s.trim());
  await pdf.destroy();
  return out;
}

describe('한글 글꼴 글자 대응표(ToUnicode)', () => {
  test('저장한 글자를 꺼내면 같은 글자 — "-"·"—"가 호환 문자(㏄·㏆)로 바뀌지 않는다', async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    doc.addPage([400, 300]).drawText('Name: Kim', { x: 20, y: 100, size: 14, font });
    const original = await doc.save();
    const ydoc = new Y.Doc();
    await seedPdfModel(original, ydoc);
    const at = '2026-10-20T01:00:00.000Z';
    const ocr: PdfOcrPage = {
      words: [
        { t: '010-1234-5678', x: 20, y: 250, ex: 150, ey: 250, h: 12 },
        { t: '가동률—정상', x: 20, y: 220, ex: 120, ey: 220, h: 12 },
      ],
      author: '가',
      at,
    };
    ydoc.getMap<PdfOcrPage>(PDF_OCR_MAP).set('o:0', ocr);
    const x0 = 20 + ['N', 'a', 'm', 'e', ':', ' '].reduce((w, c) => w + font.widthOfTextAtSize(c, 14), 0);
    const edit: PdfTextEdit = { id: 'e', pageKey: 'o:0', rect: [x0 - 0.5, 98, 200, 112], oldText: 'Kim', newText: '김-철수', fontSize: 14, author: '가', createdAt: at };
    ydoc.getMap<PdfTextEdit>(PDF_TEXT_EDITS_MAP).set('e', edit);
    const found = await texts(await exportPdfFromRoom(original, ydoc, true));
    expect(found).toContain('010-1234-5678');
    expect(found).toContain('가동률—정상');
    expect(found.join('|')).toContain('김-철수');
    expect(found.join('')).not.toMatch(/[㏄㏆]/);
  });
});
