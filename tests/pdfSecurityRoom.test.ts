import { EventEmitter } from 'node:events';
import { randomFillSync, randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import sharp from 'sharp';
import { PDFDocument as Cantoo } from '@cantoo/pdf-lib';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  send(_d: Uint8Array, cb?: (err?: Error) => void): void {
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

async function plainPdf(withImage = false): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const page = doc.addPage([400, 400]);
  page.drawText('Confidential report', { x: 20, y: 360, size: 16, font });
  if (withImage) {
    // 잡음 그림(압축이 잘 안 되는 큰 PNG) — 쪽에는 작게 그린다
    const noise = Buffer.alloc(1200 * 900 * 3);
    randomFillSync(noise);
    const png = await sharp(noise, { raw: { width: 1200, height: 900, channels: 3 } }).png().toBuffer();
    page.drawImage(await doc.embedPng(png), { x: 20, y: 20, width: 240, height: 180 });
  }
  return doc.save();
}

async function encrypted(bytes: Uint8Array, password: string): Promise<Uint8Array> {
  const d = await Cantoo.load(bytes);
  d.encrypt({ userPassword: password, ownerPassword: `${password}-owner`, permissions: { printing: 'highResolution', copying: false } });
  return d.save();
}

async function textOf(bytes: Uint8Array, password?: string): Promise<string> {
  const pdf = await getDocument({ data: new Uint8Array(bytes), password, isEvalSupported: false, standardFontDataUrl: 'node_modules/pdfjs-dist/standard_fonts/' }).promise;
  const t = (await (await pdf.getPage(1)).getTextContent()).items.map((i) => ('str' in i ? i.str : '')).join('');
  await pdf.destroy();
  return t;
}

describe('PDF 방: 암호·압축·속성(6단계)', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-pdf-sec-'));
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
    await store.writeMeta({ id: docId, name: '보고서.pdf', ext: 'pdf', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100 });
    await room.joinEditSession(new FakeSocket() as never, {} as never, docId, 'pdf', [], false);
    const ydoc = room.getRoomYdocForTest(docId)!;
    const read = () => store.readOriginal(docId, 'pdf');
    return { store, room, ops, model, docId, ydoc, read };
  }

  test('암호 파일 — 잠김, 틀린 비밀번호 거부, 맞히면 그 토큰만 원본을 받고, 저장하면 같은 비밀번호로 다시 잠긴다', async () => {
    const { room, ops, model, docId, ydoc, read, store } = await open(await encrypted(await plainPdf(), 'pw1234'));
    expect(ydoc.getMap(model.PDF_META_MAP).get('locked')).toBe(true);
    await expect(ops.editorSourceOf(docId, read, 'a')).rejects.toMatchObject({ status: 423 });
    await expect(ops.unlockPdf(docId, 'nope', 'a')).rejects.toMatchObject({ status: 403 });
    await ops.unlockPdf(docId, 'pw1234', 'a');
    expect(ydoc.getMap(model.PDF_META_MAP).get('locked')).toBe(false);
    expect(ydoc.getArray(model.PDF_PAGES_ARRAY).length).toBe(1);
    const src = await ops.editorSourceOf(docId, read, 'a');
    expect(await textOf(src.bytes)).toBe('Confidential report');
    // 다른 토큰은 아직 잠김 — 같은 비밀번호를 넣으면 열린다
    await expect(ops.editorSourceOf(docId, read, 'b')).rejects.toMatchObject({ status: 423 });
    await ops.unlockPdf(docId, 'pw1234', 'b');
    await ops.editorSourceOf(docId, read, 'b');
    // 편집 후 저장 → 같은 비밀번호로 잠긴 파일
    ydoc.transact(() => ydoc.getMap<string>(model.PDF_PROPS_MAP).set('title', '분기 보고서'), 'u');
    await room.flushNow(docId);
    const saved = await store.readOriginal(docId, 'pdf');
    await expect(textOf(saved)).rejects.toMatchObject({ name: 'PasswordException' });
    expect(await textOf(saved, 'pw1234')).toBe('Confidential report');
    // 제목은 pdf.js로 확인(@cantoo/pdf-lib는 암호 파일의 정보 사전을 읽지 못한다)
    const pdf = await getDocument({ data: new Uint8Array(saved), password: 'pw1234', isEvalSupported: false }).promise;
    expect(((await pdf.getMetadata()).info as { Title?: string }).Title).toBe('분기 보고서');
    await pdf.destroy();
  });

  test('문서 보호 — 열기 비밀번호·복사 금지로 잠갔다가 풀기, 변경 기록', async () => {
    const { room, ops, model, docId, ydoc, store } = await open(await plainPdf());
    await ops.protectPdf(docId, { mode: 'set', userPassword: 'open1', ownerPassword: 'own1', permissions: { copying: false } }, 'a', '가');
    expect(ydoc.getMap(model.PDF_META_MAP).get('protection')).toMatchObject({ enabled: true, openPassword: true, permissions: { copying: false, printing: true } });
    await room.flushNow(docId);
    const locked = await store.readOriginal(docId, 'pdf');
    expect(await textOf(locked, 'open1')).toBe('Confidential report');
    const pdf = await getDocument({ data: new Uint8Array(locked), password: 'open1', isEvalSupported: false }).promise;
    const perms = (await pdf.getPermissions()) ?? [];
    await pdf.destroy();
    expect(perms).not.toContain(16); // 복사(비트 5)
    expect(perms).toContain(4); // 인쇄
    await ops.protectPdf(docId, { mode: 'remove' }, 'a', '가');
    await room.flushNow(docId);
    expect(await textOf(await store.readOriginal(docId, 'pdf'))).toBe('Confidential report');
    expect(room.getChangeLog(docId).map((e) => (e as { kind: string; text?: string }).text)).toContain('암호 풀기');
    await expect(ops.protectPdf(docId, { mode: 'set' }, 'a', '가')).rejects.toMatchObject({ status: 400 });
  });

  test('압축 — 미리 보기 크기, 적용하면 저장 파일이 작아지고 글자는 그대로', async () => {
    const { room, ops, docId, store } = await open(await plainPdf(true));
    const preview = await ops.compressRoomPdf(docId, 'medium', true, 'a', '가');
    expect(preview!.images).toBe(1);
    expect(preview!.after).toBeLessThan(preview!.before / 3);
    await ops.compressRoomPdf(docId, 'medium', false, 'a', '가');
    await room.flushNow(docId);
    const saved = await store.readOriginal(docId, 'pdf');
    expect(saved.length).toBeLessThan(preview!.before / 3);
    expect(await textOf(saved)).toBe('Confidential report');
  });

  test('문서 정보 — 버전·쪽 크기·글꼴(포함 여부)', async () => {
    const { ops, docId, read } = await open(await plainPdf());
    const info = await ops.pdfInfoOf(docId, read, 'a');
    expect(info.pageCount).toBe(1);
    expect(info.pageSizes).toEqual([{ w: 400, h: 400, count: 1 }]);
    expect(info.fonts).toEqual([{ name: 'Helvetica', type: 'Type1', embedded: false }]);
    expect(info.version).toMatch(/^1\.\d$/);
  });
});
