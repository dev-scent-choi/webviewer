import { PDFDocument, PDFName } from 'pdf-lib';
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { exportPdfFromRoom, PdfChangeTracker, PDF_ANNOTS_MAP, PDF_FORM_MAP, PDF_PAGES_ARRAY, seedPdfModel } from '../src/pdf/pdfModel.js';
import { readManagedAnnots, stripManagedAnnots } from '../src/pdf/annotRead.js';
import type { PdfAnnot, PdfPageEntry } from '../src/pdf/annotTypes.js';

/** 1쪽: 글자 칸·체크·라디오·목록·여러 고르기 목록·읽기 전용 칸, 2쪽: 글자 칸 하나 */
async function formPdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const p1 = doc.addPage([595, 842]);
  const p2 = doc.addPage([595, 842]);
  const form = doc.getForm();
  form.createTextField('name').addToPage(p1, { x: 100, y: 760, width: 200, height: 24 });
  form.createCheckBox('agree').addToPage(p1, { x: 100, y: 720, width: 16, height: 16 });
  const radio = form.createRadioGroup('size');
  radio.addOptionToPage('S', p1, { x: 100, y: 680, width: 14, height: 14 });
  radio.addOptionToPage('L', p1, { x: 130, y: 680, width: 14, height: 14 });
  const dd = form.createDropdown('city');
  dd.addOptions(['Seoul', 'Busan']);
  dd.addToPage(p1, { x: 100, y: 640, width: 200, height: 22 });
  const ol = form.createOptionList('tags');
  ol.addOptions(['a', 'b', 'c']);
  ol.enableMultiselect();
  ol.addToPage(p1, { x: 100, y: 560, width: 200, height: 60 });
  const ro = form.createTextField('fixed');
  ro.setText('keep');
  ro.enableReadOnly();
  ro.addToPage(p1, { x: 100, y: 520, width: 200, height: 22 });
  form.createTextField('second').addToPage(p2, { x: 100, y: 760, width: 200, height: 24 });
  return doc.save();
}

async function seeded(original: Uint8Array): Promise<Y.Doc> {
  const ydoc = new Y.Doc();
  await seedPdfModel(original, ydoc);
  return ydoc;
}

describe('PDF 양식·서명 저장', () => {
  test('바꾼 칸만 채운다 — 한글 글자·체크·라디오·목록·여러 고르기, 읽기 전용·없는 칸·틀린 값은 건너뜀', async () => {
    const original = await formPdf();
    const ydoc = await seeded(original);
    const form = ydoc.getMap<unknown>(PDF_FORM_MAP);
    form.set('name', '홍길동');
    form.set('agree', true);
    form.set('size', 'L');
    form.set('city', 'Busan');
    form.set('tags', ['a', 'c', 'zz']);
    form.set('fixed', 'changed');
    form.set('nope', 'x');
    form.set('second', 42 as unknown);
    const out = await PDFDocument.load(await exportPdfFromRoom(original, ydoc, true));
    const f = out.getForm();
    expect(f.getTextField('name').getText()).toBe('홍길동');
    expect(f.getCheckBox('agree').isChecked()).toBe(true);
    expect(f.getRadioGroup('size').getSelected()).toBe('L');
    expect(f.getDropdown('city').getSelected()).toEqual(['Busan']);
    expect(f.getOptionList('tags').getSelected()).toEqual(['a', 'c']);
    expect(f.getTextField('fixed').getText()).toBe('keep');
    expect(f.getTextField('second').getText()).toBeUndefined();
  });

  test('한글 값은 pdf.js로 다시 읽어도 그 값(모양 포함)', async () => {
    const original = await formPdf();
    const ydoc = await seeded(original);
    ydoc.getMap<unknown>(PDF_FORM_MAP).set('name', '서명란 한글');
    const out = await exportPdfFromRoom(original, ydoc, true);
    const pdf = await getDocument({ data: new Uint8Array(out), isEvalSupported: false }).promise;
    const widget = (await (await pdf.getPage(1)).getAnnotations()).find((a) => a.fieldName === 'name');
    expect(widget?.fieldValue).toBe('서명란 한글');
    expect(widget?.hasAppearance).toBe(true);
    await pdf.destroy();
  });

  test('지운 쪽에만 있던 칸은 양식에서 빠지고, 그 쪽은 파일에 남지 않는다', async () => {
    const original = await formPdf();
    const ydoc = await seeded(original);
    const pages = ydoc.getArray<PdfPageEntry>(PDF_PAGES_ARRAY);
    ydoc.transact(() => pages.delete(1, 1));
    const out = await PDFDocument.load(await exportPdfFromRoom(original, ydoc, true));
    const names = out.getForm().getFields().map((x) => x.getName());
    expect(names).not.toContain('second');
    expect(names).toContain('name');
    const pageObjects = out.context.enumerateIndirectObjects().filter(([, o]) => (o as { get?: (n: PDFName) => unknown }).get?.(PDFName.of('Type')) === PDFName.of('Page'));
    expect(pageObjects).toHaveLength(1);
  });

  test('서명·도장(그림) — 표준 Stamp 주석(모양 포함)으로 쓰고, 다시 읽으면 같은 그림', async () => {
    const png =
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR4nGNgYGD4z8DAwMDAxMDAAAAMBAEBkT9gkAAAAABJRU5ErkJggg==';
    const original = await formPdf();
    const ydoc = await seeded(original);
    const at = '2026-10-20T01:00:00.000Z';
    const stamp: PdfAnnot = { id: 'sig1', pageKey: 'o:0', type: 'stamp', rect: [300, 100, 460, 160], color: '#000000', opacity: 1, width: 0, image: png, author: '가', createdAt: at, modifiedAt: at };
    ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP).set('sig1', stamp);
    const bytes = await exportPdfFromRoom(original, ydoc, true);
    const back = readManagedAnnots(await PDFDocument.load(bytes));
    expect(back).toHaveLength(1);
    expect(back[0]).toMatchObject({ id: 'sig1', type: 'stamp', image: png, rect: [300, 100, 460, 160] });
    const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise;
    const annots = (await (await pdf.getPage(1)).getAnnotations()).filter((a) => a.subtype === 'Stamp');
    expect(annots).toHaveLength(1);
    expect(annots[0].hasAppearance).toBe(true);
    await pdf.destroy();
    // 다른 프로그램의 Stamp(그림 사본 없음)는 다루지 않는다
    const other = await PDFDocument.load(bytes);
    const dict = other.context.lookup(other.getPage(0).node.Annots()!.asArray().find((r) => (other.context.lookup(r) as { get(n: PDFName): unknown }).get(PDFName.of('Subtype')) === PDFName.of('Stamp'))!) as unknown as { delete(n: PDFName): void };
    dict.delete(PDFName.of('DocviewImage'));
    expect(readManagedAnnots(other)).toHaveLength(0);
    stripManagedAnnots(other);
    expect(other.getPage(0).node.Annots()!.asArray().some((r) => (other.context.lookup(r) as { get(n: PDFName): unknown }).get(PDFName.of('Subtype')) === PDFName.of('Stamp'))).toBe(true);
  });

  test('변경 기록 — 양식 칸마다 마지막 값', async () => {
    const ydoc = await seeded(await formPdf());
    const tracker = new PdfChangeTracker(ydoc, () => '가');
    const form = ydoc.getMap<unknown>(PDF_FORM_MAP);
    ydoc.transact(() => form.set('name', '처음'), 'u');
    ydoc.transact(() => form.set('name', '나중'), 'u');
    ydoc.transact(() => form.set('agree', true), 'u');
    const rows = tracker.list().filter((e) => e.kind === 'form_changed').map((e) => e.text);
    expect(rows).toEqual(['name: 나중', 'agree: 체크']);
  });
});
