import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fsSync from 'node:fs';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/**
 * FINAL-REVIEW M1 — 유예 중 저장이 도는 동안 방을 강제로 닫으면 teardown 정리(조각 기록 삭제·ydoc.destroy)가 저장보다 먼저 돌아,
 * 저장이 조각 기록을 다시 만들고(누수) Y.Doc이 저장 도중 정리됐다. M2 — getCfRemoteAnswer가 기다리는 동안 방이 닫히면 닫힌 방으로 답했다.
 */
const MESSAGE_SYNC = 0;
const MESSAGE_PAGE_SUBSCRIBE = 2;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();
  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    if (decoding.readVarUint(decoder) === MESSAGE_SYNC) {
      const t = decoding.readVarUint(decoder);
      if (t === messageYjsSyncStep2 || t === messageYjsUpdate) readUpdate(decoder, this.mirror, this);
    }
    cb?.();
  }
  close(): void {
    this.readyState = 3;
  }
  ping(): void {}
  syncStep1(): void {
    const e = encoding.createEncoder();
    encoding.writeVarUint(e, MESSAGE_SYNC);
    writeSyncStep1(e, this.mirror);
    this.emit('message', encoding.toUint8Array(e));
  }
  edit(change: (doc: Y.Doc) => void): void {
    let captured: Uint8Array | null = null;
    const on = (u: Uint8Array) => (captured = u);
    this.mirror.on('update', on);
    change(this.mirror);
    this.mirror.off('update', on);
    const e = encoding.createEncoder();
    encoding.writeVarUint(e, MESSAGE_SYNC);
    writeUpdate(e, captured!);
    this.emit('message', encoding.toUint8Array(e));
  }
}
const request = { url: '/ws/edit/x?token=t' } as never;
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
async function flushEventLoop(turns = 50): Promise<void> {
  const { drainRoomBackgroundWorkForTest } = await import('../src/collab/room.js');
  for (let i = 0; i < turns; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
    await fs.stat(os.tmpdir()).catch(() => undefined);
    await drainRoomBackgroundWorkForTest();
  }
}

describe('M1·M2: teardown과 진행 중 작업', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-m1m2-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_ROOM_CLOSE_GRACE_MS', '5000');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });
  afterEach(async () => {
    vi.restoreAllMocks();
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('M1: 유예 저장 도중 forceCloseRoom — Y.Doc 정리는 저장 뒤, 조각 기록은 다시 생기지 않는다', async () => {
    const { joinEditSession, forceCloseRoom, getRoomYdocForTest, hasDocxFragmentRecordsForTest } = await import('../src/collab/room.js');
    const { store } = await import('../src/storage.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'docx', fsSync.readFileSync('audit/office/corpus/docx/d04-images.docx'));
    await store.writeMeta({ id: docId, name: 'x.docx', ext: 'docx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100 });
    const a = new FakeSocket();
    await joinEditSession(a as never, request, docId, 'docx');
    a.syncStep1();
    await flushEventLoop();
    a.edit((doc) => doc.getXmlFragment('prosemirror').delete(0, 1));
    await flushEventLoop();

    const order: string[] = [];
    const realSave = store.saveOriginal.bind(store);
    vi.spyOn(store, 'saveOriginal').mockImplementation(async (...args: Parameters<typeof store.saveOriginal>) => {
      await sleep(300); // 느린 디스크
      await realSave(...args);
      order.push('saved');
    });
    getRoomYdocForTest(docId)!.on('destroy', () => order.push('destroyed'));
    a.emit('close'); // 유예 — 저장 시작
    await sleep(20);
    forceCloseRoom(docId); // 저장 도중 문서 삭제
    await sleep(900);
    await flushEventLoop();
    expect(order).toEqual(['saved', 'destroyed']);
    expect(hasDocxFragmentRecordsForTest(docId)).toBe(false);
  }, 60_000);

  test('M2: 기다리는 동안 방이 닫히면 null(경로 409)', async () => {
    const { joinEditSession, getCfRemoteAnswer, forceCloseRoom } = await import('../src/collab/room.js');
    const { store } = await import('../src/storage.js');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    for (let r = 1; r <= 600; r += 1) ws.addRow([`K${r}`]);
    ws.addConditionalFormatting({ ref: 'A1:A600', rules: [{ type: 'expression', priority: 1, formulae: ['COUNTIF($A:$A,$A1)>1'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
    const docId = randomUUID();
    await store.saveOriginal(docId, 'xlsx', Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer));
    await store.writeMeta({ id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid' });
    const s = new FakeSocket();
    await joinEditSession(s as never, request, docId, 'xlsx', [], false, 'v');
    const e = encoding.createEncoder();
    encoding.writeVarUint(e, MESSAGE_PAGE_SUBSCRIBE);
    encoding.writeVarString(e, JSON.stringify({ sheet: 0, pages: [0] }));
    s.emit('message', encoding.toUint8Array(e));
    await flushEventLoop();
    const asking = getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 599, c1: 0, crit: ['K1'] }] });
    forceCloseRoom(docId); // 답을 준비하는 사이 방이 닫힌다
    expect(await asking).toBeNull();
  }, 60_000);
});
