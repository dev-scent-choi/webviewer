import fs from 'node:fs';
import fontkit from '@pdf-lib/fontkit';
import ExcelJS from 'exceljs';
import sharp from 'sharp';
import { strFromU8, unzipSync } from 'fflate';
import { PDFDocument, rgb } from 'pdf-lib';
import { describe, expect, test } from 'vitest';
import { convertPdf, pdfToDocxText, pdfToText, pdfToXlsx } from '../src/pdf/pdfConvert.js';
import { fileToPdf, FileToPdfError } from '../src/pdf/fileToPdf.js';
import { sofficePath } from '../src/converters/libreOffice.js';

/** 제목·문단 두 줄·표(항목/값/비고) 한 쪽 + 둘째 쪽 */
async function reportPdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const font = await doc.embedFont(fs.readFileSync('assets/fonts/Pretendard-Regular.ttf'), { subset: false });
  const p = doc.addPage([595, 842]);
  p.drawText('분기 보고서', { x: 60, y: 770, size: 24, font, color: rgb(0.1, 0.2, 0.4) });
  p.drawText('이번 분기 가동률은 목표를 넘었습니다.', { x: 60, y: 720, size: 11, font });
  p.drawText('다음 분기 계획을 검토합니다.', { x: 60, y: 704, size: 11, font });
  const rows = [
    ['항목', '값', '비고'],
    ['가동률', '97.2%', '정상'],
    ['생산량', '12,500', '증가'],
  ];
  rows.forEach((r, i) => r.forEach((cell, c) => p.drawText(cell, { x: 60 + c * 120, y: 640 - i * 20, size: 11, font })));
  doc.addPage([595, 842]).drawText('둘째 쪽 내용', { x: 60, y: 770, size: 11, font });
  return doc.save();
}

describe('PDF 변환(7단계)', async () => {
  const pdf = await reportPdf();

  test('텍스트 — 줄 단위, 쪽 사이 빈 줄, BOM', async () => {
    const text = await pdfToText(pdf);
    expect(text.startsWith('﻿')).toBe(true);
    const lines = text.slice(1).split('\r\n');
    expect(lines.slice(0, 3)).toEqual(['분기 보고서', '이번 분기 가동률은 목표를 넘었습니다.', '다음 분기 계획을 검토합니다.']);
    expect(lines).toContain('가동률 97.2% 정상');
    expect(text).toContain('\r\n\r\n둘째 쪽 내용');
  });

  test('워드(글만) — 제목은 굵게, 두 줄 문단은 하나로, 쪽 나눔', async () => {
    const z = unzipSync(await pdfToDocxText(pdf));
    const xml = strFromU8(z['word/document.xml']);
    const paras = [...xml.matchAll(/<w:p>(.*?)<\/w:p>/g)].map((m) => [...m[1].matchAll(/<w:t[^>]*>([^<]*)</g)].map((t) => t[1]).join(''));
    expect(paras[0]).toBe('분기 보고서');
    expect(xml).toMatch(/<w:b\/><w:sz w:val="48"\/>/);
    expect(paras[1]).toBe('이번 분기 가동률은 목표를 넘었습니다. 다음 분기 계획을 검토합니다.');
    expect(xml).toContain('<w:br w:type="page"/>');
    expect(paras.at(-1)).toBe('둘째 쪽 내용');
  });

  test('엑셀 — 표의 행·열, 숫자·백분율은 숫자로, 쪽마다 시트', async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(Buffer.from(await pdfToXlsx(pdf)) as unknown as ArrayBuffer);
    expect(wb.worksheets.map((w) => w.name)).toEqual(['1쪽', '2쪽']);
    const s = wb.worksheets[0];
    const rowOf = (label: string) => {
      let found = 0;
      s.eachRow((row, n) => {
        if (row.getCell(1).value === label) found = n;
      });
      return s.getRow(found);
    };
    expect(rowOf('항목').values).toEqual([undefined, '항목', '값', '비고']);
    expect(rowOf('가동률').getCell(2).value).toBeCloseTo(0.972);
    expect(rowOf('가동률').getCell(2).numFmt).toBe('0.0%');
    expect(rowOf('생산량').getCell(2).value).toBe(12500);
  });

  test.skipIf(!sofficePath)('워드·파워포인트(모양 그대로) — LibreOffice', async () => {
    const docx = await convertPdf(pdf, 'docx', '분기보고서.pdf');
    expect(docx.fileName).toBe('분기보고서.docx');
    expect(strFromU8(unzipSync(docx.bytes)['word/document.xml'])).toContain('분기 보고서');
    const pptx = await convertPdf(pdf, 'pptx', '분기보고서.pdf');
    const z = unzipSync(pptx.bytes);
    expect(Object.keys(z).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))).toHaveLength(2);
    expect(strFromU8(z['ppt/slides/slide1.xml'])).toContain('가동률');
  }, 120_000);

  test('다른 파일 → PDF — 그림, 모르는 형식은 415', async () => {
    const png = await sharp({ create: { width: 200, height: 100, channels: 3, background: '#3a7bd5' } }).png().toBuffer();
    const out = await PDFDocument.load(await fileToPdf(png, '도장.png'));
    expect(out.getPageCount()).toBe(1);
    await expect(fileToPdf(new Uint8Array([1, 2, 3]), 'a.xyz')).rejects.toBeInstanceOf(FileToPdfError);
    expect(await fileToPdf(pdf, 'x.pdf')).toBe(pdf);
  });
});
