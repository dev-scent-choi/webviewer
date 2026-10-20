import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNumber, PDFRef, PDFString } from 'pdf-lib';
import { describe, expect, test } from 'vitest';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { readManagedAnnots, stripManagedAnnots } from '../src/pdf/annotRead.js';
import { writeAnnots } from '../src/pdf/annotWrite.js';
import type { PdfAnnot } from '../src/pdf/annotTypes.js';

const at = '2026-10-20T01:30:00.000Z';
const base = { author: '향기랩 방문자', createdAt: at, modifiedAt: at, opacity: 1, width: 2 } as const;
const SAMPLE: PdfAnnot[] = [
  { ...base, id: 'n1', pageKey: 'o:0', type: 'note', rect: [50, 700, 74, 724], color: '#FFD84D', text: '확인 부탁드립니다' },
  { ...base, id: 'h1', pageKey: 'o:0', type: 'highlight', rect: [100, 600, 300, 620], quads: [100, 620, 300, 620, 100, 600, 300, 600], color: '#FFE600', opacity: 0.5 },
  { ...base, id: 'u1', pageKey: 'o:0', type: 'underline', rect: [100, 560, 300, 580], quads: [100, 580, 300, 580, 100, 560, 300, 560], color: '#1A73E8' },
  { ...base, id: 's1', pageKey: 'o:0', type: 'strike', rect: [100, 520, 300, 540], quads: [100, 540, 300, 540, 100, 520, 300, 520], color: '#D93025' },
  { ...base, id: 'i1', pageKey: 'o:1', type: 'ink', rect: [10, 10, 100, 100], paths: [[10, 10, 50, 60, 100, 100]], color: '#0000FF', width: 3 },
  { ...base, id: 'f1', pageKey: 'o:1', type: 'freetext', rect: [100, 500, 300, 560], text: '한글 글상자 English', fontSize: 14, color: '#0000FF', width: 1 },
  { ...base, id: 'f2', pageKey: 'o:1', type: 'freetext', rect: [100, 400, 300, 440], text: 'ASCII only', fontSize: 12, color: '#000000', width: 1 },
  { ...base, id: 'r1', pageKey: 'o:1', type: 'rect', rect: [300, 300, 400, 380], color: '#FF0000', fill: '#E6E6E6' },
  { ...base, id: 'e1', pageKey: 'o:1', type: 'ellipse', rect: [300, 100, 400, 180], color: '#008000', fill: null },
  { ...base, id: 'l1', pageKey: 'o:1', type: 'line', rect: [100, 100, 200, 150], line: [100, 100, 200, 150], color: '#000000' },
  { ...base, id: 'a1', pageKey: 'o:1', type: 'arrow', rect: [100, 200, 200, 250], line: [100, 200, 200, 250], color: '#D93025' },
];

async function blank(pages = 2): Promise<PDFDocument> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i += 1) doc.addPage([600, 800]);
  return doc;
}

/** 쪽 2개: 1쪽에 형광펜·메모(+Popup)·링크·답글, 2쪽에 펜·사각형·화살표·글상자 */
async function fixture(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const p1 = doc.addPage([600, 800]);
  const p2 = doc.addPage([600, 800]);
  const ctx = doc.context;
  const add = (page: typeof p1, dict: Record<string, unknown>) => {
    const ref = ctx.register(ctx.obj({ Type: 'Annot', ...dict }));
    page.node.addAnnot(ref);
    return ref;
  };
  add(p1, {
    Subtype: 'Highlight',
    Rect: [100, 600, 300, 620],
    QuadPoints: [100, 620, 300, 620, 100, 600, 300, 600],
    C: [1, 1, 0],
    CA: 0.5,
    T: PDFHexString.fromText('홍길동'),
    NM: PDFString.of('hl-1'),
    M: PDFString.of('D:20261020103000+09\'00\''),
  });
  const note = add(p1, { Subtype: 'Text', Rect: [50, 700, 74, 724], Contents: PDFHexString.fromText('검토 필요'), C: [1, 0.85, 0.3] });
  const popup = add(p1, { Subtype: 'Popup', Rect: [80, 600, 280, 700], Parent: note });
  (ctx.lookup(note) as PDFDict).set(PDFName.of('Popup'), popup);
  add(p1, { Subtype: 'Link', Rect: [10, 10, 60, 30] });
  add(p1, { Subtype: 'Text', Rect: [50, 650, 74, 674], Contents: PDFString.of('reply'), IRT: note });
  add(p2, { Subtype: 'Ink', Rect: [10, 10, 200, 200], InkList: [[10, 10, 50, 60, 100, 100]], C: [0, 0, 1], BS: ctx.obj({ W: 3 }) });
  add(p2, { Subtype: 'Square', Rect: [300, 300, 400, 380], C: [1, 0, 0], IC: [0.9, 0.9, 0.9], Border: [0, 0, 2] });
  add(p2, { Subtype: 'Line', Rect: [0, 0, 600, 800], L: [100, 100, 200, 150], LE: [PDFName.of('None'), PDFName.of('OpenArrow')], C: [0, 0.5, 0] });
  add(p2, { Subtype: 'FreeText', Rect: [100, 500, 300, 540], Contents: PDFHexString.fromText('안녕 PDF'), DA: PDFString.of('/Helv 14 Tf 0 0 1 rg') });
  return doc.save();
}

describe('PDF 주석 읽기', () => {
  test('다루는 종류만 우리 형식으로 — 답글·Popup·링크는 읽지 않는다', async () => {
    const doc = await PDFDocument.load(await fixture());
    const annots = readManagedAnnots(doc);
    expect(annots.map((a) => `${a.pageKey}:${a.type}`)).toEqual(['o:0:highlight', 'o:0:note', 'o:1:ink', 'o:1:rect', 'o:1:arrow', 'o:1:freetext']);
    const hl = annots[0];
    expect(hl).toMatchObject({ id: 'hl-1', color: '#FFFF00', opacity: 0.5, author: '홍길동', quads: [100, 620, 300, 620, 100, 600, 300, 600] });
    expect(hl.modifiedAt.startsWith('2026-10-20T01:30:00')).toBe(true);
    expect(annots[1]).toMatchObject({ type: 'note', text: '검토 필요', rect: [50, 700, 74, 724] });
    expect(annots[2]).toMatchObject({ paths: [[10, 10, 50, 60, 100, 100]], color: '#0000FF', width: 3 });
    expect(annots[3]).toMatchObject({ color: '#FF0000', fill: '#E6E6E6', width: 2 });
    expect(annots[4]).toMatchObject({ line: [100, 100, 200, 150], color: '#008000' });
    expect(annots[5]).toMatchObject({ text: '안녕 PDF', fontSize: 14, color: '#0000FF' });
    // NM이 없으면 쪽·순번으로 같은 id를 다시 만든다(다시 열어도 같은 id)
    expect(annots[1].id).toBe(readManagedAnnots(await PDFDocument.load(await fixture()))[1].id);
  });

  test('빼기 — 다루는 주석과 그 Popup만 빠지고 링크·답글은 남는다', async () => {
    const doc = await PDFDocument.load(await fixture());
    stripManagedAnnots(doc);
    const subtypes = (pageIndex: number) => {
      const arr = doc.getPage(pageIndex).node.lookupMaybe(PDFName.of('Annots'), PDFArray);
      if (!arr) return [];
      return arr.asArray().map((ref) => (doc.context.lookup(ref as PDFRef) as PDFDict).get(PDFName.of('Subtype'))?.toString());
    };
    expect(subtypes(0)).toEqual(['/Link', '/Text']); // 링크 + 답글
    expect(subtypes(1)).toEqual([]);
    expect(readManagedAnnots(doc)).toEqual([]);
  });

  test('주석이 없거나 이상한 값이어도 죽지 않는다', async () => {
    const doc = await PDFDocument.create();
    const page = doc.addPage([100, 100]);
    page.node.addAnnot(doc.context.register(doc.context.obj({ Type: 'Annot', Subtype: 'Square', Rect: [PDFNumber.of(5)] })));
    expect(() => readManagedAnnots(doc)).not.toThrow();
    expect(readManagedAnnots(doc)).toEqual([]);
  });
});

describe('PDF 주석 쓰기', () => {
  test('쓴 것을 다시 읽으면 같은 값(왕복)', async () => {
    const doc = await blank();
    await writeAnnots(doc, SAMPLE);
    const back = readManagedAnnots(await PDFDocument.load(await doc.save()));
    const norm = (a: PdfAnnot) => ({ ...a, rect: undefined });
    expect(back.map(norm)).toEqual(SAMPLE.map(norm));
  });

  test('다른 뷰어(pdf.js)가 표준 주석과 모양으로 읽는다', async () => {
    const doc = await blank();
    await writeAnnots(doc, SAMPLE);
    const bytes = await doc.save();
    const pdf = await getDocument({ data: bytes.slice(), isEvalSupported: false }).promise;
    const p1 = (await (await pdf.getPage(1)).getAnnotations()).map((a: { subtype: string; hasAppearance: boolean }) => `${a.subtype}:${a.hasAppearance}`);
    const p2 = (await (await pdf.getPage(2)).getAnnotations()).map((a: { subtype: string; hasAppearance: boolean }) => `${a.subtype}:${a.hasAppearance}`);
    expect(p1).toEqual(['Text:true', 'Highlight:true', 'Underline:true', 'StrikeOut:true']);
    expect(p2).toEqual(['Ink:true', 'FreeText:true', 'FreeText:true', 'Square:true', 'Circle:true', 'Line:true', 'Line:true']);
    const note = (await (await pdf.getPage(1)).getAnnotations())[0] as { contentsObj: { str: string }; titleObj: { str: string } };
    expect(note.contentsObj.str).toBe('확인 부탁드립니다');
    expect(note.titleObj.str).toBe('향기랩 방문자');
    await pdf.destroy();
  });

  test('한글 글상자는 모양에 한글 글꼴을 넣고, 영어만이면 표준 글꼴(파일이 커지지 않게)', async () => {
    const korean = await blank(1);
    await writeAnnots(korean, [{ ...SAMPLE[5], pageKey: 'o:0' }]);
    const ascii = await blank(1);
    await writeAnnots(ascii, [{ ...SAMPLE[6], pageKey: 'o:0' }]);
    const kBytes = await korean.save();
    const aBytes = await ascii.save();
    expect(kBytes.length).toBeGreaterThan(500_000); // Pretendard 통째로(subset은 글자가 빠진다)
    expect(aBytes.length).toBeLessThan(20_000);
    // pdf.js로 그 글상자 모양을 그려 글자가 실제로 찍히는지는 브라우저 확인(verify:pdf)에서 본다
    const pdf = await getDocument({ data: kBytes.slice(), isEvalSupported: false }).promise;
    const [ft] = await (await pdf.getPage(1)).getAnnotations();
    expect((ft as { contentsObj: { str: string } }).contentsObj.str).toBe('한글 글상자 English');
    await pdf.destroy();
  });

  test('쪽이 없는 pageKey는 건너뛴다', async () => {
    const doc = await blank(1);
    await writeAnnots(doc, [{ ...SAMPLE[0], pageKey: 'o:9' }]);
    expect(readManagedAnnots(doc)).toEqual([]);
  });
});
