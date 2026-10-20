import fs from 'node:fs';
import fontkit from '@pdf-lib/fontkit';
import sharp from 'sharp';
import { PDFArray, PDFDocument, PDFName, PDFRawStream, PDFString, StandardFonts, type PDFPage } from 'pdf-lib';
import { describe, expect, test } from 'vitest';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { parseContent } from '../src/pdf/contentStream.js';
import { fontsFor } from '../src/pdf/annotWrite.js';
import { rewritePage, streamBytes, type RewriteTarget } from '../src/pdf/textRewrite.js';

const PRETENDARD = fs.readFileSync('assets/fonts/Pretendard-Regular.ttf');
/** 글자별 너비 합 — drawText는 커닝 없이 그린다(widthOfTextAtSize는 커닝 포함) */
const wid = (font: { widthOfTextAtSize(t: string, s: number): number }, text: string, size: number) => [...text].reduce((w, ch) => w + font.widthOfTextAtSize(ch, size), 0);

/** 쪽마다 글 조각(글자·x·y) */
async function items(bytes: Uint8Array, pageNo = 1): Promise<Array<{ str: string; x: number; y: number }>> {
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, standardFontDataUrl: 'node_modules/pdfjs-dist/standard_fonts/' }).promise;
  const content = await (await pdf.getPage(pageNo)).getTextContent();
  const out = content.items
    .map((it) => ('str' in it ? { str: it.str, x: Math.round(it.transform[4] * 10) / 10, y: Math.round(it.transform[5] * 10) / 10 } : null))
    .filter((x): x is { str: string; x: number; y: number } => !!x && x.str.trim() !== '');
  await pdf.destroy();
  return out;
}
const textOf = async (bytes: Uint8Array, pageNo = 1) => (await items(bytes, pageNo)).map((i) => i.str).join('');

async function run(built: PDFDocument, pageObj: PDFPage, targets: RewriteTarget[]) {
  // 실제 저장 경로처럼 저장했다 불러온 문서에서
  const index = built.getPages().indexOf(pageObj);
  const doc = await PDFDocument.load(await built.save());
  const page = doc.getPage(index);
  const result = await rewritePage(doc, page, targets, fontsFor(doc));
  return { result, bytes: await doc.save() };
}

describe('내용 스트림 해석', () => {
  test('문자열 이스케이프·배열·사전·인라인 그림·리터럴 null', () => {
    const src = Buffer.from('q 1 0 0 1 5 5 cm BT /F1 12 Tf (a\\(b\\)\\101) Tj [(x) -250 <4142>] TJ ET /Tag <</MCID 3>> BDC EMC null pop BI /W 2 /H 1 /BPC 8 /CS /G ID \x00EI\x01 EI Q', 'latin1');
    const ops = parseContent(src);
    expect(ops.map((o) => o.op)).toEqual(['q', 'cm', 'BT', 'Tf', 'Tj', 'TJ', 'ET', 'BDC', 'EMC', 'pop', 'BI', 'Q']);
    const tj = ops[4].operands[0] as { bytes: Uint8Array };
    expect(Buffer.from(tj.bytes).toString('latin1')).toBe('a(b)A');
    expect(ops[9].operands).toEqual([null]);
    // 원래 바이트 위치 — 그대로 옮겨 쓸 수 있다
    expect(src.subarray(ops[3].start, ops[3].end).toString('latin1')).toBe('/F1 12 Tf');
  });
});

describe('가리기', () => {
  test('Helvetica — 걸친 글자만 지우고 남은 글자 위치는 그대로, 검은 상자', async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const page = doc.addPage([400, 200]);
    page.drawText('Hello World Again', { x: 20, y: 100, size: 20, font });
    const before = await items(await doc.save());
    const x0 = 20 + wid(font, 'Hello ', 20);
    const x1 = 20 + wid(font, 'Hello World', 20);
    const { result, bytes } = await run(doc, page, [{ id: 'r', mode: 'redact', rect: [x0 + 0.5, 95, x1 - 0.5, 120] }]);
    expect(result.removedGlyphs).toBe(5);
    expect(await textOf(bytes)).not.toContain('World');
    const after = await items(bytes);
    // "Again"은 같은 자리
    const againBefore = before.find((i) => i.str.includes('Again'));
    const again = after.find((i) => i.str.includes('Again'));
    expect(again).toBeTruthy();
    const xAgain = 20 + wid(font, 'Hello World ', 20);
    expect(Math.abs((again!.x + (again!.str.indexOf('Again') ? font.widthOfTextAtSize(again!.str.slice(0, again!.str.indexOf('Again')), 20) : 0)) - xAgain)).toBeLessThan(0.5);
    expect(againBefore).toBeTruthy();
    const saved = await PDFDocument.load(bytes);
    const content = Buffer.from(streamBytes(saved.context.lookup(saved.getPage(0).node.get(PDFName.of('Contents'))!)!)!).toString('latin1');
    expect(content).toMatch(/0 g [\d.]+ 95 [\d.]+ 25 re f/);
  });

  test('한글(Type0 Identity-H) — 숫자 부분만 지운다', async () => {
    const doc = await PDFDocument.create();
    doc.registerFontkit(fontkit);
    const font = await doc.embedFont(PRETENDARD, { subset: false });
    const page = doc.addPage([400, 200]);
    page.drawText('가동률은 97.2%입니다', { x: 20, y: 100, size: 16, font });
    const x0 = 20 + wid(font, '가동률은 ', 16);
    const x1 = 20 + wid(font, '가동률은 97.2%', 16);
    const { result, bytes } = await run(doc, page, [{ id: 'r', mode: 'redact', rect: [x0 + 0.5, 95, x1 - 0.5, 115] }]);
    expect(result.removedGlyphs).toBe(5);
    const text = await textOf(bytes);
    expect(text).toContain('가동률은');
    expect(text).toContain('입니다');
    expect(text).not.toMatch(/97|2%/);
  });

  test('Form XObject 안의 글자 — 이 쪽은 사본을 고치고 다른 쪽(원본)은 그대로', async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const src = doc.addPage([300, 200]);
    src.drawText('SECRET 1234', { x: 20, y: 100, size: 20, font });
    const embedded = await doc.embedPage(src);
    const target = doc.addPage([300, 200]);
    target.drawPage(embedded, { x: 0, y: 0 });
    const { result, bytes } = await run(doc, target, [{ id: 'r', mode: 'redact', rect: [10, 90, 300, 130] }]);
    expect(result.removedGlyphs).toBe(11);
    expect(await textOf(bytes, 2)).toBe('');
    expect(await textOf(bytes, 1)).toBe('SECRET 1234');
  });

  test('그림 — PNG(Flate)·JPEG 걸친 부분을 검게, 링크는 뺀다', async () => {
    const doc = await PDFDocument.create();
    const page = doc.addPage([300, 300]);
    const white = await sharp({ create: { width: 100, height: 100, channels: 3, background: '#ffffff' } }).png().toBuffer();
    const whiteJpg = await sharp({ create: { width: 100, height: 100, channels: 3, background: '#ffffff' } }).jpeg().toBuffer();
    const png = await doc.embedPng(white);
    const jpg = await doc.embedJpg(whiteJpg);
    page.drawImage(png, { x: 0, y: 0, width: 100, height: 100 });
    page.drawImage(jpg, { x: 150, y: 0, width: 100, height: 100 });
    const ctx = doc.context;
    page.node.addAnnot(ctx.register(ctx.obj({ Type: 'Annot', Subtype: 'Link', Rect: [10, 10, 40, 40], A: ctx.obj({ S: 'URI', URI: PDFString.of('https://example.com') }) })));
    page.node.addAnnot(ctx.register(ctx.obj({ Type: 'Annot', Subtype: 'Link', Rect: [200, 200, 240, 240] })));
    // 두 그림의 왼쪽 아래 1/4씩
    const { result, bytes } = await run(doc, page, [
      { id: 'a', mode: 'redact', rect: [0, 0, 50, 50] },
      { id: 'b', mode: 'redact', rect: [150, 0, 200, 50] },
    ]);
    expect(result.blackedImages).toBe(2);
    expect(result.removedAnnots).toBe(1);
    const saved = await PDFDocument.load(bytes);
    expect((saved.getPage(0).node.Annots() as PDFArray).size()).toBe(1);
    // 바뀐 그림들의 화소 — 왼쪽 아래(그림 좌표로 아래쪽 = 행 큰 쪽)는 검정, 오른쪽 위는 흰색
    const images = saved.context
      .enumerateIndirectObjects()
      .map(([, o]) => o)
      .filter((o): o is PDFRawStream => o instanceof PDFRawStream && o.dict.get(PDFName.of('Subtype')) === PDFName.of('Image'))
      .filter((o) => o.dict.get(PDFName.of('Filter')) === PDFName.of('FlateDecode') && !o.dict.has(PDFName.of('DecodeParms')));
    const blacked = images.map((img) => streamBytes(img)!).filter((raw) => raw.length === 100 * 100 * 3 && raw[(90 * 100 + 10) * 3] === 0);
    expect(blacked.length).toBe(2);
    for (const raw of blacked) expect(raw[(10 * 100 + 90) * 3]).toBeGreaterThan(200);
  });
});

describe('글자 고치기', () => {
  test('새 글이 원래 글자들로만 되어 있으면 원래 글꼴로, 같은 자리에', async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const page = doc.addPage([400, 200]);
    page.drawText('Total 1234 won', { x: 20, y: 100, size: 14, font });
    const x0 = 20 + wid(font, 'Total ', 14);
    const x1 = 20 + wid(font, 'Total 1234', 14);
    const { result, bytes } = await run(doc, page, [{ id: 'e', mode: 'edit', rect: [x0 - 0.5, 98, x1 - 0.5, 112], oldText: '1234', newText: '4321' }]);
    expect(result.removedGlyphs).toBe(4);
    const text = await textOf(bytes);
    expect(text.replace(/\s+/g, ' ')).toContain('4321');
    expect(text).not.toContain('1234');
    const saved = await PDFDocument.load(bytes);
    const fonts = saved.getPage(0).node.Resources()!.lookup(PDFName.of('Font')) as unknown as { keys(): PDFName[] };
    expect(fonts.keys().map((k) => k.decodeText()).some((k) => k.startsWith('DvF'))).toBe(false);
    const neu = (await items(bytes)).find((i) => i.str.includes('4321'))!;
    expect(Math.abs(neu.x - x0)).toBeLessThan(0.5);
    expect(Math.abs(neu.y - 100)).toBeLessThan(0.5);
  });

  test('다른 글자(한글)면 Pretendard로 같은 크기·자리에', async () => {
    const doc = await PDFDocument.create();
    const font = await doc.embedFont(StandardFonts.Helvetica);
    const page = doc.addPage([400, 200]);
    page.drawText('Name: Kim', { x: 20, y: 100, size: 14, font });
    const x0 = 20 + wid(font, 'Name: ', 14);
    const { bytes } = await run(doc, page, [{ id: 'e', mode: 'edit', rect: [x0 - 0.5, 98, 200, 112], oldText: 'Kim', newText: '김철수' }]);
    const text = await textOf(bytes);
    expect(text).toContain('김철수');
    expect(text).not.toContain('Kim');
    const neu = (await items(bytes)).find((i) => i.str.includes('김철수'))!;
    expect(Math.abs(neu.x - x0)).toBeLessThan(0.5);
  });
});
