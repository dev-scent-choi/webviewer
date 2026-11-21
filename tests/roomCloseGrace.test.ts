import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { writeSyncStep1 } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSampleDocx } from './helpers.js';
import { MESSAGE_ROOM_EPOCH, MESSAGE_ROOM_STALE, ROOM_STALE_CLOSE_CODE } from '../src/collab/roomEpoch.js';

// 방 세대·닫기 유예·옛 상태 거절(src/collab/room.ts [RX-C1] 블록) — 3차 검토 R3-I1·R3-I2
const MESSAGE_SYNC = 0;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly types: number[] = [];
  epoch: string | null = null;
  closedWith: number | null = null;
  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    const type = decoding.readVarUint(decoder);
    this.types.push(type);
    if (type === MESSAGE_ROOM_EPOCH) this.epoch = decoding.readVarString(decoder);
    cb?.();
  }
  close(code?: number): void {
    if (this.closedWith === null) this.closedWith = code ?? 1005;
    this.readyState = 3;
  }
  ping(): void {}
  /** 연결 직후 첫 동기화(SyncStep1) — 주어진 문서의 상태 벡터를 보낸다 */
  sync(doc: Y.Doc): void {
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_SYNC);
    writeSyncStep1(encoder, doc);
    this.emit('message', encoding.toUint8Array(encoder));
  }
}

const request = (epoch?: string) => ({ url: epoch ? `/ws/edit/x?token=t&epoch=${epoch}` : '/ws/edit/x?token=t' }) as never;

describe('room.ts [RX-C1]: 방 세대·닫기 유예·옛 상태 거절', () => {
  let tmpDir: string;
  let docId: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-grace-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_ROOM_CLOSE_GRACE_MS', '400');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
    const { store } = await import('../src/storage.js');
    docId = randomUUID();
    await store.saveOriginal(docId, 'docx', makeSampleDocx(['Hello']));
    await store.writeMeta({ id: docId, name: 'x.docx', ext: 'docx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100 });
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('마지막 참여자가 나가도 유예 동안은 같은 세대로 다시 붙고, 유예가 끝나면 닫힌다', async () => {
    const { joinEditSession, isRoomOpen } = await import('../src/collab/room.js');
    const a = new FakeSocket();
    await joinEditSession(a as never, request(), docId, 'docx');
    expect(a.epoch).toBeTruthy();
    a.emit('close');
    expect(isRoomOpen(docId)).toBe(true); // 유예 중
    const b = new FakeSocket();
    await joinEditSession(b as never, request(a.epoch!), docId, 'docx');
    expect(b.closedWith).toBeNull();
    expect(b.epoch).toBe(a.epoch); // 같은 세대
    b.emit('close');
    await new Promise((r) => setTimeout(r, 700));
    expect(isRoomOpen(docId)).toBe(false);
  });

  test('세대가 다르면(방이 새로 만들어졌으면) 메시지 8을 보내고 보통 닫기, 방을 새로 만들지 않는다', async () => {
    const { joinEditSession, isRoomOpen } = await import('../src/collab/room.js');
    const s = new FakeSocket();
    await joinEditSession(s as never, request('old-epoch'), docId, 'docx');
    expect(s.types).toEqual([MESSAGE_ROOM_STALE]);
    expect(s.closedWith).toBe(1000);
    expect(isRoomOpen(docId)).toBe(false);
  });

  test('세대 값 없는 연결: 다른 세대의 상태 벡터면 4409로 거절, 빈 문서·같은 세대 상태면 통과', async () => {
    const { joinEditSession } = await import('../src/collab/room.js');
    const fresh = new FakeSocket();
    const join = joinEditSession(fresh as never, request(), docId, 'docx');
    fresh.sync(new Y.Doc());
    await join;
    expect(fresh.closedWith).toBeNull();

    const foreign = new Y.Doc();
    foreign.getXmlFragment('prosemirror').insert(0, [new Y.XmlElement('paragraph')]);
    const old = new FakeSocket();
    const join2 = joinEditSession(old as never, request(), docId, 'docx');
    old.sync(foreign);
    await join2;
    expect(old.closedWith).toBe(ROOM_STALE_CLOSE_CODE);
    expect(old.types).toContain(MESSAGE_ROOM_STALE);
    expect(old.types).not.toContain(MESSAGE_SYNC); // 문서를 받지도 주지도 않았다
  });

  test('문서 삭제(forceCloseRoom)는 유예 중인 방도 바로 닫는다', async () => {
    const { joinEditSession, isRoomOpen, forceCloseRoom } = await import('../src/collab/room.js');
    const a = new FakeSocket();
    await joinEditSession(a as never, request(), docId, 'docx');
    a.emit('close');
    expect(isRoomOpen(docId)).toBe(true);
    forceCloseRoom(docId);
    expect(isRoomOpen(docId)).toBe(false);
  });
});
