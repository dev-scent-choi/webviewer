import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSampleDocx, makeSamplePdf } from './helpers.js';

/**
 * FINAL-REVIEW I2 — 닫기 유예가 방 개수(20)로만 묶여 큰 PDF·덱·문서 방 20개가 45초 동안 메모리에 남을 수 있던 문제.
 * 방 하나가 상한(DOCVIEW_ROOM_LINGER_MAX_BYTES)을 넘으면 유예하지 않고, 유예 중인 방들의 합이 상한(DOCVIEW_ROOM_LINGER_TOTAL_BYTES)을
 * 넘으면 가장 오래된 방부터 닫는다.
 */
class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  send(_data: Uint8Array, cb?: (err?: Error) => void): void {
    cb?.();
  }
  close(): void {
    this.readyState = 3;
  }
  ping(): void {}
}
const request = { url: '/ws/edit/x?token=t' } as never;

describe('I2: 닫기 유예의 바이트 상한', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-linger-bytes-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_ROOM_CLOSE_GRACE_MS', '5000');
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

  async function openAndLeave(paragraphs: number): Promise<string> {
    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'docx', makeSampleDocx(Array.from({ length: paragraphs }, (_, i) => `문단 ${i} ${'가'.repeat(200)}`)));
    await store.writeMeta({ id: docId, name: 'x.docx', ext: 'docx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100 });
    const s = new FakeSocket();
    await joinEditSession(s as never, request, docId, 'docx');
    s.emit('close');
    return docId;
  }

  test('상한보다 큰 방은 유예하지 않고 바로 닫는다(작은 방은 유예)', async () => {
    vi.stubEnv('DOCVIEW_ROOM_LINGER_MAX_BYTES', String(200 * 1024));
    const { isRoomOpen } = await import('../src/collab/room.js');
    const small = await openAndLeave(2);
    const big = await openAndLeave(600); // 600 × 약 600바이트 — 상한 200KB를 넘는다
    expect(isRoomOpen(small)).toBe(true);
    expect(isRoomOpen(big)).toBe(false);
  });

  test('유예 중인 방들의 합이 상한을 넘으면 가장 오래된 방부터 닫는다', async () => {
    vi.stubEnv('DOCVIEW_ROOM_LINGER_TOTAL_BYTES', String(400 * 1024));
    const { isRoomOpen } = await import('../src/collab/room.js');
    const a = await openAndLeave(250);
    const b = await openAndLeave(250);
    const c = await openAndLeave(250);
    expect(isRoomOpen(a)).toBe(false); // 가장 오래된 방
    expect(isRoomOpen(b)).toBe(true);
    expect(isRoomOpen(c)).toBe(true);
  });
  test('재검토 I2: 큰 PDF를 불러오는 중(채우기 전) 마지막 사람이 나가면 유예 없이 닫힌다', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, isRoomOpen } = await import('../src/collab/room.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'pdf', await makeSamplePdf(2));
    await store.writeMeta({ id: docId, name: 'x.pdf', ext: 'pdf', status: 'ready', pageCount: 2, createdAt: new Date().toISOString(), progress: 100 });
    const realRead = store.readOriginal.bind(store);
    vi.spyOn(store, 'readOriginal').mockImplementation(async (...args: Parameters<typeof store.readOriginal>) => {
      await new Promise((r) => setTimeout(r, 300)); // 큰 파일 — 읽는 데 오래 걸린다
      return realRead(...args);
    });
    const s = new FakeSocket();
    const joining = joinEditSession(s as never, request, docId, 'pdf');
    await new Promise((r) => setTimeout(r, 20));
    s.emit('close'); // 아직 채우는 중
    expect(isRoomOpen(docId)).toBe(false);
    await joining;
    vi.restoreAllMocks();
  });
});
