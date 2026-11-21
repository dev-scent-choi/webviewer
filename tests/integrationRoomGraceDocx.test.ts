import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fsSync from 'node:fs';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1, writeUpdate } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { MESSAGE_ROOM_EPOCH, MESSAGE_ROOM_STALE } from '../src/collab/roomEpoch.js';

/**
 * [통합] pptx 갈래의 방 세대·닫기 유예(room.ts [RX-C1])와 docx 갈래의 방 수명 보존 조각 저장소(room.docxFragments, docxfrag.json).
 * 지운 그림 조각이 유예 동안 같은 방에 남고, 유예가 끝나 방이 닫힌 뒤 새로 만든 방에도 docxfrag.json으로 이어지는지.
 */
const MESSAGE_SYNC = 0;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();
  readonly types: number[] = [];
  epoch: string | null = null;
  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    const type = decoding.readVarUint(decoder);
    this.types.push(type);
    if (type === MESSAGE_ROOM_EPOCH) this.epoch = decoding.readVarString(decoder);
    if (type === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) readUpdate(decoder, this.mirror, this);
    }
    cb?.();
  }
  close(): void {
    this.readyState = 3;
  }
  ping(): void {}
  syncStep1(): void {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeSyncStep1(encoder, this.mirror);
    this.emit('message', encoding.toUint8Array(encoder));
  }
  /** mirror에서 고친 변경만 서버로 보낸다 */
  edit(change: (doc: Y.Doc) => void): void {
    let captured: Uint8Array | null = null;
    const on = (u: Uint8Array) => (captured = u);
    this.mirror.on('update', on);
    change(this.mirror);
    this.mirror.off('update', on);
    if (!captured) throw new Error('no update');
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeUpdate(encoder, captured);
    this.emit('message', encoding.toUint8Array(encoder));
  }
}

const request = (epoch?: string) => ({ url: epoch ? `/ws/edit/x?token=t&epoch=${epoch}` : '/ws/edit/x?token=t' }) as never;
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));
async function flushEventLoop(turns = 50): Promise<void> {
  const { drainRoomBackgroundWorkForTest } = await import('../src/collab/room.js');
  for (let i = 0; i < turns; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
    await fs.stat(os.tmpdir()).catch(() => undefined);
    await drainRoomBackgroundWorkForTest();
  }
}

/** 그림(ooxml kind=image)이 든 첫 최상위 블록의 번호 */
function imageBlockIndex(fragment: Y.XmlFragment): number {
  const hasImage = (el: Y.XmlElement | Y.XmlText): boolean =>
    el instanceof Y.XmlElement && ((el.nodeName === 'ooxml' && el.getAttribute('kind') === 'image') || el.toArray().some((c) => hasImage(c as Y.XmlElement)));
  return fragment.toArray().findIndex((c) => hasImage(c as Y.XmlElement));
}
const drawingHashes = (frags: ReadonlyMap<string, string> | null) => [...(frags ?? new Map()).entries()].filter(([, xml]) => xml.includes('<w:drawing')).map(([h]) => h);

describe('[통합] 방 닫기 유예·세대 × 워드 보존 조각 저장소', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-integ-docx-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_ROOM_CLOSE_GRACE_MS', '3000'); // 병렬 시험 부하에서도 "유예 중" 확인이 유예보다 먼저 끝나게(재검토 3 — 400ms는 부하 때 흔들렸다)
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

  test('지운 그림 조각: 유예 중엔 같은 방 저장소에, 방이 닫힌 뒤엔 docxfrag.json으로 새 방에 이어진다', async () => {
    const { joinEditSession, docxFragmentsOf, isRoomOpen, hasDocxFragmentRecordsForTest } = await import('../src/collab/room.js');
    const { store } = await import('../src/storage.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'docx', fsSync.readFileSync('audit/office/corpus/docx/d04-images.docx'));
    await store.writeMeta({ id: docId, name: 'x.docx', ext: 'docx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100 });

    const a = new FakeSocket();
    await joinEditSession(a as never, request(), docId, 'docx');
    a.syncStep1();
    await flushEventLoop();
    const frags = docxFragmentsOf(docId)!;
    const before = drawingHashes(frags);
    expect(before.length).toBeGreaterThan(0);
    const at = imageBlockIndex(a.mirror.getXmlFragment('prosemirror'));
    expect(at).toBeGreaterThanOrEqual(0);
    a.edit((doc) => doc.getXmlFragment('prosemirror').delete(at, 1)); // 그림 문단을 지운다
    await flushEventLoop();
    a.emit('close'); // 유예 — 저장은 바로 한다
    await flushEventLoop();
    expect(isRoomOpen(docId)).toBe(true);
    expect(docxFragmentsOf(docId)).toBe(frags); // 같은 방, 같은 저장소(지운 조각도 그대로)
    expect(drawingHashes(frags)).toEqual(before);
    const b = new FakeSocket();
    await joinEditSession(b as never, request(a.epoch!), docId, 'docx');
    expect(b.types).not.toContain(MESSAGE_ROOM_STALE);
    expect(docxFragmentsOf(docId)).toBe(frags);
    b.emit('close');
    await sleep(3_600);
    await flushEventLoop();
    expect(isRoomOpen(docId)).toBe(false);
    expect(hasDocxFragmentRecordsForTest(docId)).toBe(false); // 닫힌 방의 문서별 기록은 놓는다
    // 지운 그림 조각이 docxfrag.json에 남았다
    const saved = JSON.parse((await store.readDerived(docId, 'docxfrag')).toString('utf8')) as Record<string, string>;
    expect(Object.values(saved).some((xml) => xml.includes('<w:drawing'))).toBe(true);
    // 지난 세대 탭은 거절, 새로 연 방은 파일에서 조각을 이어받는다(되돌리기로 다시 넣어도 저장된다)
    const stale = new FakeSocket();
    await joinEditSession(stale as never, request(a.epoch!), docId, 'docx');
    expect(stale.types).toContain(MESSAGE_ROOM_STALE);
    const c = new FakeSocket();
    await joinEditSession(c as never, request(), docId, 'docx');
    c.syncStep1();
    await flushEventLoop();
    const fresh = docxFragmentsOf(docId)!;
    expect(fresh).not.toBe(frags);
    expect(hasDocxFragmentRecordsForTest(docId)).toBe(true);
    for (const h of Object.keys(saved)) expect(fresh.has(h)).toBe(true);
  }, 60_000);
});
