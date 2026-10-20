import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { unzipSync } from 'fflate';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import type { PdfAnnot, PdfPageEntry } from '../src/pdf/annotTypes.js';

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  send(_data: Uint8Array, cb?: (err?: Error) => void): void {
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

async function numberedPdf(count: number, label: string): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let i = 0; i < count; i += 1) doc.addPage([300 + i * 10, 400]).drawText(`${label} ${i + 1}`, { x: 20, y: 350, size: 20, font });
  return doc.save();
}
const widths = async (bytes: Uint8Array) => (await PDFDocument.load(bytes)).getPages().map((p) => Math.round(p.getWidth()));

describe('PDF 방: 쪽 정리 저장·부속·추출', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-pdf-pages-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  async function open(original: Uint8Array) {
    const { store } = await import('../src/storage.js');
    const room = await import('../src/collab/room.js');
    const ops = await import('../src/pdf/pdfRoomOps.js');
    const model = await import('../src/pdf/pdfModel.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'pdf', Buffer.from(original));
    await store.writeMeta({ id: docId, name: '보고서.pdf', ext: 'pdf', status: 'ready', pageCount: 3, createdAt: new Date().toISOString(), progress: 100 });
    await room.joinEditSession(new FakeSocket() as never, {} as never, docId, 'pdf', [], false);
    const ydoc = room.getRoomYdocForTest(docId)!;
    return { store, room, ops, model, docId, ydoc };
  }

  test('저장을 두 번 해도 쪽 정리가 한 번만 적용된다(방을 열 때의 원본 기준)', async () => {
    const { store, room, model, docId, ydoc } = await open(await numberedPdf(3, 'page'));
    const pages = ydoc.getArray<PdfPageEntry>(model.PDF_PAGES_ARRAY);
    ydoc.transact(() => {
      pages.delete(0, 1); // 1쪽 지움 → [o:1, o:2]
    }, 'user');
    await room.flushNow(docId);
    expect(await widths(await store.readOriginal(docId, 'pdf'))).toEqual([310, 320]);
    ydoc.transact(() => {
      const last = pages.get(1);
      pages.delete(1, 1);
      pages.insert(0, [last]); // → [o:2, o:1]
    }, 'user');
    await room.flushNow(docId);
    expect(await widths(await store.readOriginal(docId, 'pdf'))).toEqual([320, 310]);
    expect(await widths((await room.getCurrentBytes(docId))!)).toEqual([320, 310]);
  });

  test('부속 PDF — 넣고 받고, 저장하면 그 쪽과 주석이 들어간다. 편집기 원본은 방을 열 때의 것', async () => {
    const original = await numberedPdf(2, 'page');
    const { store, room, ops, model, docId, ydoc } = await open(original);
    const part = await ops.addPdfPart(docId, await numberedPdf(2, 'part'), 10_000_000);
    expect(part.pageCount).toBe(2);
    expect(part.sizes.map((s) => Math.round(s.w))).toEqual([300, 310]);
    expect(await ops.readPdfPart(docId, part.partId)).toBeInstanceOf(Uint8Array);
    expect(await ops.readPdfPart(docId, 'nope')).toBeNull();
    const at = '2026-10-20T01:00:00.000Z';
    const annot: PdfAnnot = { id: 'x', pageKey: `p:${part.partId}:1`, type: 'rect', rect: [10, 10, 50, 50], color: '#D93025', opacity: 1, width: 2, author: '가', createdAt: at, modifiedAt: at };
    ydoc.transact(() => {
      ydoc.getArray<PdfPageEntry>(model.PDF_PAGES_ARRAY).push([{ key: `p:${part.partId}:1`, rotate: 90 }]);
      ydoc.getMap<PdfAnnot>(model.PDF_ANNOTS_MAP).set('x', annot);
    }, 'user');
    await room.flushNow(docId);
    const saved = await PDFDocument.load(await store.readOriginal(docId, 'pdf'));
    expect(saved.getPages().map((p) => Math.round(p.getWidth()))).toEqual([300, 310, 310]);
    expect(saved.getPage(2).getRotation().angle).toBe(90);
    // 저장 뒤에도 편집기 원본은 방을 열 때의 것(열쇠 기준이 바뀌지 않게)
    const src = await ops.editorSourceOf(docId, () => store.readOriginal(docId, 'pdf'));
    expect(src.base).toBe(model.pdfBaseFingerprint(original));
    expect(ydoc.getMap(model.PDF_META_MAP).get('base')).toBe(src.base);
  });

  test('부속 — 방이 없으면 409, PDF가 아니면 400, 너무 크면 413', async () => {
    const { ops, docId } = await open(await numberedPdf(1, 'page'));
    await expect(ops.addPdfPart(randomUUID(), await numberedPdf(1, 'a'), 10_000_000)).rejects.toMatchObject({ status: 409 });
    await expect(ops.addPdfPart(docId, new TextEncoder().encode('hello'), 10_000_000)).rejects.toMatchObject({ status: 400 });
    await expect(ops.addPdfPart(docId, await numberedPdf(1, 'a'), 10)).rejects.toMatchObject({ status: 413 });
  });

  test('추출 — 한 묶음은 pdf, 여러 묶음은 zip(지금 편집 상태 기준)', async () => {
    const { store, ops, model, docId, ydoc } = await open(await numberedPdf(3, 'page'));
    const pages = ydoc.getArray<PdfPageEntry>(model.PDF_PAGES_ARRAY);
    ydoc.transact(() => pages.delete(0, 1), 'user');
    const read = () => store.readOriginal(docId, 'pdf');
    const one = await ops.extractPdfGroups(docId, [['o:2', 'o:1']], '보고서.pdf', read);
    expect(one.type).toBe('pdf');
    expect(one.fileName).toBe('보고서-추출.pdf');
    expect(await widths(one.bytes)).toEqual([320, 310]);
    const many = await ops.extractPdfGroups(docId, [['o:1'], ['o:2'], ['o:0']], '보고서.pdf', read);
    expect(many.type).toBe('zip');
    const files = unzipSync(many.bytes);
    expect(Object.keys(files)).toEqual(['보고서-1.pdf', '보고서-2.pdf']); // 지운 쪽(o:0) 묶음은 빠짐
    await expect(ops.extractPdfGroups(docId, [['o:0']], '보고서.pdf', read)).rejects.toMatchObject({ status: 400 });
  });

  test('변경 기록 — 쪽 정리 요약 한 줄', async () => {
    const { room, model, docId, ydoc } = await open(await numberedPdf(3, 'page'));
    const pages = ydoc.getArray<PdfPageEntry>(model.PDF_PAGES_ARRAY);
    ydoc.transact(() => {
      pages.delete(0, 1);
      const e = pages.get(0);
      pages.delete(0, 1);
      pages.insert(0, [{ ...e, rotate: 90 }]);
      pages.push([{ key: 'b:z', rotate: 0, blank: { w: 100, h: 100 } }]);
    }, 'user');
    const log = room.getChangeLog ? await room.getChangeLog(docId) : null;
    const entries = (log ?? []) as Array<{ kind: string; text?: string }>;
    expect(entries.find((e) => e.kind === 'pages_changed')?.text).toBe('1쪽 지움 · 1쪽 넣음 · 1쪽 돌림');
  });
});
