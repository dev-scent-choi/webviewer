/**
 * PDF 변환(7단계, 설계 §3 "변환") — 지금 편집 상태의 PDF를
 *  - 워드: 모양 그대로(LibreOffice writer_pdf_import, 쪽마다 글상자) / 글만(문단으로 흐르는 docx — 고쳐 쓰기 좋게)
 *  - 파워포인트: LibreOffice impress_pdf_import(쪽마다 슬라이드, 글상자·도형)
 *  - 엑셀: 글자 위치로 행·열을 맞춰 표로(쪽마다 시트, 숫자는 숫자로)
 *  - 텍스트: 줄 단위
 * 로 바꾼다. 그림(PNG/JPG)은 브라우저가 pdf.js로 그린다(src/client/pdf/convertUi.ts).
 */
import path from 'node:path';
import { createRequire } from 'node:module';
import ExcelJS from 'exceljs';
import { strToU8, zipSync } from 'fflate';
import { convertPdfToOffice } from '../converters/libreOffice.js';

export type ConvertTarget = 'docx' | 'docx-text' | 'pptx' | 'xlsx' | 'txt';
export const CONVERT_TARGETS: readonly ConvertTarget[] = ['docx', 'docx-text', 'pptx', 'xlsx', 'txt'];
export const isConvertTarget = (v: unknown): v is ConvertTarget => CONVERT_TARGETS.includes(v as ConvertTarget);

const MIME: Record<ConvertTarget, string> = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'docx-text': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  txt: 'text/plain; charset=utf-8',
};
const EXT: Record<ConvertTarget, string> = { docx: 'docx', 'docx-text': 'docx', pptx: 'pptx', xlsx: 'xlsx', txt: 'txt' };

/** 글 조각(PDF 좌표, 왼쪽 아래 기준) */
interface Item {
  readonly str: string;
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly size: number;
}
interface Line {
  readonly items: Item[];
  readonly y: number;
  readonly size: number;
  readonly x: number;
  text: string;
}
interface PageText {
  readonly lines: Line[];
  readonly items: Item[];
  readonly width: number;
}

const require = createRequire(import.meta.url);
function pdfjsAssets(): { cMapUrl: string; standardFontDataUrl: string } {
  const root = path.dirname(require.resolve('pdfjs-dist/package.json'));
  return { cMapUrl: `${path.join(root, 'cmaps')}${path.sep}`, standardFontDataUrl: `${path.join(root, 'standard_fonts')}${path.sep}` };
}

/** 쪽마다 글 조각과 줄 */
async function readPages(bytes: Uint8Array): Promise<PageText[]> {
  const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, verbosity: 0, cMapPacked: true, ...pdfjsAssets() }).promise;
  const pages: PageText[] = [];
  try {
    for (let p = 1; p <= pdf.numPages; p += 1) {
      const page = await pdf.getPage(p);
      const view = page.getViewport({ scale: 1 });
      const content = await page.getTextContent();
      const items: Item[] = [];
      for (const it of content.items) {
        if (!('str' in it) || !it.str.trim()) continue;
        const [a, b, , d, e, f] = it.transform as number[];
        items.push({ str: it.str, x: e, y: f, w: it.width, size: Math.max(1, Math.hypot(a, b) || Math.abs(d) || it.height || 10) });
      }
      pages.push({ items, lines: toLines(items), width: view.width });
      page.cleanup();
    }
  } finally {
    await pdf.destroy();
  }
  return pages;
}

/** 기준선이 비슷한 조각끼리 한 줄로(위에서 아래로), 조각 사이가 벌어졌으면 띄어 쓴다 */
function toLines(items: readonly Item[]): Line[] {
  const sorted = [...items].sort((a, b) => b.y - a.y || a.x - b.x);
  const lines: Line[] = [];
  for (const it of sorted) {
    const line = lines.find((l) => Math.abs(l.y - it.y) <= Math.min(l.size, it.size) * 0.45);
    if (line) line.items.push(it);
    else lines.push({ items: [it], y: it.y, size: it.size, x: it.x, text: '' });
  }
  for (const line of lines) {
    line.items.sort((a, b) => a.x - b.x);
    let text = '';
    let end = -Infinity;
    for (const it of line.items) {
      if (text && it.x - end > it.size * 0.2 && !text.endsWith(' ') && !it.str.startsWith(' ')) text += ' ';
      text += it.str;
      end = it.x + it.w;
    }
    line.text = text.replace(/\s+$/, '');
    (line as { x: number }).x = line.items[0].x;
    (line as { size: number }).size = Math.max(...line.items.map((i) => i.size));
  }
  return lines.sort((a, b) => b.y - a.y);
}

/** 줄 → 문단(줄 간격이 벌어지거나 글자 크기가 바뀌면 새 문단) */
function toParagraphs(lines: readonly Line[]): Array<{ text: string; size: number }> {
  const out: Array<{ text: string; size: number }> = [];
  let prev: Line | null = null;
  for (const line of lines) {
    const gap = prev ? prev.y - line.y : 0;
    const sameBlock = prev && gap < Math.max(prev.size, line.size) * 1.7 && Math.abs(prev.size - line.size) < 1.5 && Math.abs(prev.x - line.x) < prev.size * 2;
    if (sameBlock && out.length) {
      const last = out[out.length - 1];
      // 한글은 줄이 바뀌어도 띄어쓰기가 이미 있거나 붙여 쓰는 경우가 많다 — 영문 낱말 사이만 띄운다
      last.text += /[A-Za-z0-9,.;:]$/.test(last.text) && /^[A-Za-z0-9]/.test(line.text) ? ` ${line.text}` : /\s$/.test(last.text) ? line.text : ` ${line.text}`;
    } else out.push({ text: line.text, size: line.size });
    prev = line;
  }
  return out;
}

// ---------- 텍스트 ----------
export async function pdfToText(bytes: Uint8Array): Promise<string> {
  const pages = await readPages(bytes);
  // 메모장에서도 한글이 바로 보이게 BOM을 붙인다
  return `﻿${pages.map((p) => p.lines.map((l) => l.text).join('\r\n')).join('\r\n\r\n')}\r\n`;
}

// ---------- 워드(글만) ----------
const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');

export async function pdfToDocxText(bytes: Uint8Array): Promise<Uint8Array> {
  const pages = await readPages(bytes);
  const allSizes = pages.flatMap((p) => p.lines.map((l) => l.size)).sort((a, b) => a - b);
  const body = allSizes.length ? allSizes[Math.floor(allSizes.length / 2)] : 11;
  const paras: string[] = [];
  pages.forEach((page, i) => {
    for (const p of toParagraphs(page.lines)) {
      const heading = p.size >= body * 1.3;
      const sz = Math.round(Math.min(72, Math.max(8, heading ? p.size : body)) * 2);
      paras.push(
        `<w:p>${heading ? '<w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>' : ''}<w:r><w:rPr>${heading ? '<w:b/>' : ''}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/></w:rPr><w:t xml:space="preserve">${esc(p.text)}</w:t></w:r></w:p>`,
      );
    }
    if (i < pages.length - 1) paras.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
  });
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const files: Record<string, Uint8Array> = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    'word/_rels/document.xml.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    ),
    'word/styles.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="맑은 고딕" w:hAnsi="맑은 고딕" w:eastAsia="맑은 고딕" w:cs="맑은 고딕"/><w:sz w:val="22"/><w:lang w:val="ko-KR" w:eastAsia="ko-KR"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`,
    ),
    'word/document.xml': strToU8(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>${paras.join('')}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>`,
    ),
  };
  return zipSync(files);
}

// ---------- 엑셀(표로) ----------
const NUM = /^[-+]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/;
const PCT = /^[-+]?(\d{1,3}(,\d{3})+|\d+)(\.\d+)?%$/;

/** 조각의 왼쪽 끝을 모아 열 경계를 정한다(가까운 것끼리 한 열) */
function columnStarts(items: readonly Item[]): number[] {
  const xs = [...items.map((i) => i.x)].sort((a, b) => a - b);
  const starts: number[] = [];
  for (const x of xs) if (!starts.length || x - starts[starts.length - 1] > 12) starts.push(x);
  return starts;
}

export async function pdfToXlsx(bytes: Uint8Array): Promise<Uint8Array> {
  const pages = await readPages(bytes);
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Docview';
  pages.forEach((page, pi) => {
    const sheet = wb.addWorksheet(`${pi + 1}쪽`);
    const starts = columnStarts(page.items);
    const colOf = (x: number) => {
      let c = 0;
      for (let i = 0; i < starts.length; i += 1) if (starts[i] <= x + 2) c = i;
      return c;
    };
    const widths = new Array<number>(starts.length).fill(8);
    page.lines.forEach((line, r) => {
      const cells = new Map<number, string>();
      for (const it of line.items) {
        const c = colOf(it.x);
        cells.set(c, cells.has(c) ? `${cells.get(c)} ${it.str.trim()}` : it.str.trim());
      }
      const row = sheet.getRow(r + 1);
      for (const [c, raw] of cells) {
        const text = raw.trim();
        const cell = row.getCell(c + 1);
        if (NUM.test(text)) cell.value = Number(text.replace(/[,+]/g, ''));
        else if (PCT.test(text)) {
          cell.value = Number(text.replace(/[,%+]/g, '')) / 100;
          cell.numFmt = text.includes('.') ? '0.0%' : '0%';
        } else cell.value = text;
        if (line.size >= 14) cell.font = { bold: true, size: Math.min(20, Math.round(line.size)) };
        widths[c] = Math.max(widths[c], Math.min(60, [...text].reduce((w, ch) => w + (ch.charCodeAt(0) > 0x2e80 ? 2 : 1), 0) + 2));
      }
      row.commit();
    });
    widths.forEach((w, i) => (sheet.getColumn(i + 1).width = w));
  });
  if (pages.length === 0) wb.addWorksheet('1쪽');
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

/** 대상 형식으로 — 파일 이름(확장자 포함)과 MIME도 */
export async function convertPdf(bytes: Uint8Array, to: ConvertTarget, baseName: string): Promise<{ bytes: Uint8Array; fileName: string; mime: string }> {
  const stem = baseName.replace(/\.pdf$/i, '') || 'document';
  let out: Uint8Array;
  if (to === 'docx' || to === 'pptx') out = await convertPdfToOffice(bytes, to);
  else if (to === 'docx-text') out = await pdfToDocxText(bytes);
  else if (to === 'xlsx') out = await pdfToXlsx(bytes);
  else out = Buffer.from(await pdfToText(bytes), 'utf8');
  const suffix = to === 'docx-text' ? '-글만' : '';
  return { bytes: out, fileName: `${stem}${suffix}.${EXT[to]}`, mime: MIME[to] };
}

