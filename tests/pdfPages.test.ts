import { PDFArray, PDFDocument, PDFName, StandardFonts } from 'pdf-lib';
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { exportPdfFromRoom, extractPdf, pdfBaseFingerprint, PDF_ANNOTS_MAP, PDF_META_MAP, PDF_PAGES_ARRAY, PDF_SEED_ORIGIN, seedPdfModel } from '../src/pdf/pdfModel.js';
import { cleanPageList, parsePageKey, type PdfAnnot, type PdfPageEntry } from '../src/pdf/annotTypes.js';
import { readManagedAnnots } from '../src/pdf/annotRead.js';

/** 쪽마다 "쪽 N" 글자와 서로 다른 크기 — 순서를 글자·크기로 확인한다 */
async function numberedPdf(count: number, label = '쪽'): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let i = 0; i < count; i += 1) {
    const page = doc.addPage([300 + i * 10, 400]);
    page.drawText(`${label === '쪽' ? 'page' : label} ${i + 1}`, { x: 20, y: 350, size: 20, font });
  }
  return doc.save();
}

/** 쪽마다 첫 글자 줄 */
async function pageTexts(bytes: Uint8Array): Promise<string[]> {
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise;
  const out: string[] = [];
  for (let p = 1; p <= pdf.numPages; p += 1) {
    const content = await (await pdf.getPage(p)).getTextContent();
    out.push(content.items.map((it) => ('str' in it ? it.str : '')).join('').trim());
  }
  await pdf.destroy();
  return out;
}

const at = '2026-10-20T01:00:00.000Z';
const rectOn = (id: string, pageKey: string): PdfAnnot => ({ id, pageKey, type: 'rect', rect: [10, 10, 60, 60], color: '#D93025', opacity: 1, width: 2, author: '시험', createdAt: at, modifiedAt: at });

async function seeded(original: Uint8Array): Promise<Y.Doc> {
  const ydoc = new Y.Doc();
  await seedPdfModel(original, ydoc);
  return ydoc;
}
const setPages = (ydoc: Y.Doc, entries: PdfPageEntry[]) => {
  const arr = ydoc.getArray<PdfPageEntry>(PDF_PAGES_ARRAY);
  ydoc.transact(() => {
    arr.delete(0, arr.length);
    arr.push(entries);
  });
};

describe('PDF 쪽 정리 저장', () => {
  test('방을 열면 쪽 목록은 원본 순서, 원본 지문이 들어간다(시딩)', async () => {
    const original = await numberedPdf(3);
    const ydoc = await seeded(original);
    expect(ydoc.getArray<PdfPageEntry>(PDF_PAGES_ARRAY).toArray()).toEqual([
      { key: 'o:0', rotate: 0 },
      { key: 'o:1', rotate: 0 },
      { key: 'o:2', rotate: 0 },
    ]);
    expect(ydoc.getMap(PDF_META_MAP).get('base')).toBe(pdfBaseFingerprint(original));
  });

  test('순서 바꾸기·지우기·회전 — 지운 쪽의 주석은 빠지고, 남은 쪽 주석은 그 쪽을 따라간다', async () => {
    const original = await numberedPdf(4);
    const ydoc = await seeded(original);
    const map = ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP);
    map.set('a3', rectOn('a3', 'o:2'));
    map.set('a2', rectOn('a2', 'o:1'));
    setPages(ydoc, [
      { key: 'o:2', rotate: 90 },
      { key: 'o:0', rotate: 0 },
      { key: 'o:3', rotate: 0 },
    ]);
    const out = await exportPdfFromRoom(original, ydoc, true);
    expect(await pageTexts(out)).toEqual(['page 3', 'page 1', 'page 4']);
    const doc = await PDFDocument.load(out);
    expect(doc.getPages().map((p) => p.getRotation().angle)).toEqual([90, 0, 0]);
    expect(doc.getPages().map((p) => Math.round(p.getWidth()))).toEqual([320, 300, 330]);
    expect(readManagedAnnots(doc).map((a) => `${a.pageKey}:${a.id}`)).toEqual(['o:0:a3']);
  });

  test('지운 쪽의 내용은 파일에 남지 않는다', async () => {
    const original = await numberedPdf(3);
    const ydoc = await seeded(original);
    setPages(ydoc, [{ key: 'o:0', rotate: 0 }]);
    const out = await exportPdfFromRoom(original, ydoc, true);
    const doc = await PDFDocument.load(out);
    const pageObjects = doc.context.enumerateIndirectObjects().filter(([, o]) => (o as { get?: (n: PDFName) => unknown }).get?.(PDFName.of('Type')) === PDFName.of('Page'));
    expect(pageObjects).toHaveLength(1);
  });

  test('원래 쪽의 회전에 더한다', async () => {
    const src = await PDFDocument.load(await numberedPdf(1));
    src.getPage(0).setRotation({ type: 'degrees' as never, angle: 90 } as never);
    const original = await src.save();
    const ydoc = await seeded(original);
    setPages(ydoc, [{ key: 'o:0', rotate: 270 }]);
    const doc = await PDFDocument.load(await exportPdfFromRoom(original, ydoc, true));
    expect(doc.getPage(0).getRotation().angle).toBe(0);
  });

  test('빈 쪽과 넣은 PDF(부속)의 쪽, 그 쪽 주석', async () => {
    const original = await numberedPdf(2);
    const part = await numberedPdf(2, 'part');
    const ydoc = await seeded(original);
    ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP).set('pa', rectOn('pa', 'p:abc:1'));
    setPages(ydoc, [
      { key: 'o:0', rotate: 0 },
      { key: 'b:x1', rotate: 0, blank: { w: 200, h: 100 } },
      { key: 'p:abc:1', rotate: 0 },
      { key: 'p:zzz:0', rotate: 0 }, // 없는 부속 — 빠진다
      { key: 'o:1', rotate: 0 },
    ]);
    const out = await exportPdfFromRoom(original, ydoc, true, (id) => (id === 'abc' ? part : undefined));
    expect(await pageTexts(out)).toEqual(['page 1', '', 'part 2', 'page 2']);
    const doc = await PDFDocument.load(out);
    expect([Math.round(doc.getPage(1).getWidth()), Math.round(doc.getPage(1).getHeight())]).toEqual([200, 100]);
    expect(readManagedAnnots(doc).map((a) => `${a.pageKey}:${a.id}`)).toEqual(['o:2:pa']);
  });

  test('여러 번 저장해도(방을 열 때의 원본에서) 같은 결과', async () => {
    const original = await numberedPdf(3);
    const ydoc = await seeded(original);
    setPages(ydoc, [
      { key: 'o:2', rotate: 0 },
      { key: 'o:1', rotate: 0 },
    ]);
    const first = await exportPdfFromRoom(original, ydoc, true);
    const second = await exportPdfFromRoom(original, ydoc, true);
    expect(await pageTexts(first)).toEqual(['page 3', 'page 2']);
    expect(await pageTexts(second)).toEqual(['page 3', 'page 2']);
  });

  test('같은 열쇠가 두 번이면 첫 번째만, 형식이 틀린 항목은 뺀다', () => {
    expect(
      cleanPageList([
        { key: 'o:1', rotate: 0 },
        { key: 'o:1', rotate: 90 },
        { key: 'x', rotate: 0 },
        { key: 'b:a', rotate: 0 },
        { key: 'b:b', rotate: 0, blank: { w: 10, h: 10 } },
      ]),
    ).toEqual([
      { key: 'o:1', rotate: 0 },
      { key: 'b:b', rotate: 0, blank: { w: 10, h: 10 } },
    ]);
    expect(parsePageKey('p:abc_1:3')).toEqual({ kind: 'part', partId: 'abc_1', index: 3 });
  });

  test('쪽 트리가 여러 단계여도 물려받던 크기가 유지된다', async () => {
    // 2단계 트리: 루트 → 중간 노드(MediaBox를 물려줌) → 쪽 2개
    const base = await PDFDocument.load(await numberedPdf(2));
    const ctx = base.context;
    const [p1, p2] = base.getPages();
    for (const p of [p1, p2]) p.node.delete(PDFName.of('MediaBox'));
    const rootRef = base.catalog.get(PDFName.of('Pages'));
    const midRef = ctx.nextRef();
    ctx.assign(midRef, ctx.obj({ Type: 'Pages', Parent: rootRef, Kids: [p1.ref, p2.ref], Count: 2, MediaBox: [0, 0, 420, 420] }));
    p1.node.set(PDFName.of('Parent'), midRef);
    p2.node.set(PDFName.of('Parent'), midRef);
    const root = base.catalog.lookup(PDFName.of('Pages')) as unknown as { set(n: PDFName, v: unknown): void };
    root.set(PDFName.of('Kids'), ctx.obj([midRef]));
    const original = await base.save();
    expect(await pageTexts(original)).toEqual(['page 1', 'page 2']);
    expect((await PDFDocument.load(original)).getPages().map((p) => Math.round(p.getWidth()))).toEqual([420, 420]);
    const ydoc = await seeded(original);
    setPages(ydoc, [
      { key: 'o:1', rotate: 0 },
      { key: 'o:0', rotate: 0 },
    ]);
    const out = await exportPdfFromRoom(original, ydoc, true);
    expect(await pageTexts(out)).toEqual(['page 2', 'page 1']);
    const doc = await PDFDocument.load(out);
    expect(doc.getPages().map((p) => Math.round(p.getWidth()))).toEqual([420, 420]);
  });

  test('추출 — 고른 쪽만 새 PDF로(주석 포함), 여러 묶음이면 여러 개', async () => {
    const original = await numberedPdf(4);
    const ydoc = await seeded(original);
    ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP).set('a', rectOn('a', 'o:2'));
    const full = await exportPdfFromRoom(original, ydoc, true);
    const [one, two] = await extractPdf(full, [[2, 0], [3]]);
    expect(await pageTexts(one)).toEqual(['page 3', 'page 1']);
    expect(await pageTexts(two)).toEqual(['page 4']);
    const doc = await PDFDocument.load(one);
    expect(doc.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)?.size()).toBe(1);
  });

  test('시딩은 편집이 아니다 — 시딩 트랜잭션 표시', async () => {
    const ydoc = new Y.Doc();
    const origins: unknown[] = [];
    ydoc.on('update', (_u: Uint8Array, origin: unknown) => origins.push(origin));
    await seedPdfModel(await numberedPdf(1), ydoc);
    expect(origins).toEqual([PDF_SEED_ORIGIN]);
  });
});
