import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import * as encoding from 'lib0/encoding';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';
import { makeSamplePptx } from './helpers.js';

const MESSAGE_SLIDE_CONTENT_SUBSCRIBE = 5;

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

async function flushEventLoop(turns = 30): Promise<void> {
  for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

function subscribe(socket: FakeSocket, slideEntryId: string): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SLIDE_CONTENT_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ slideEntryId }));
  socket.emit('message', encoding.toUint8Array(encoder));
}

// 슬라이드를 "열기"만 한 것(원본 XML → Y.Doc 시딩)은 편집이 아니다 — 예전엔 이것이 자동저장을
// 예약해, 아무도 고치지 않은 대용량 pptx를 8초 뒤 통째로 다시 내보내고 PDF 재변환까지 돌렸다.
describe('room.ts: pptx 슬라이드 시딩은 저장을 일으키지 않는다', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-pptx-seed-nosave-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  async function openAndSeed() {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getRoomYdocForTest } = await import('../src/collab/room.js');
    const { PPTX_SLIDES_ARRAY, PPTX_SLIDE_CONTENTS_MAP, PPTX_SLIDE_VISUALS_MAP } = await import(
      '../src/collab/pptxShared.js'
    );
    const docId = randomUUID();
    await store.saveOriginal(docId, 'pptx', makeSamplePptx(['Slide A', 'Slide B']));
    await store.writeMeta({
      id: docId, name: 'x.pptx', ext: 'pptx', status: 'ready', pageCount: 2,
      createdAt: new Date().toISOString(), progress: 100,
    });
    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'pptx', [], false);
    const ydoc = getRoomYdocForTest(docId)!;
    const ids = ydoc.getArray<{ id: string }>(PPTX_SLIDES_ARRAY).toArray().map((e) => e.id);
    const saveSpy = vi.spyOn(store, 'saveOriginal');
    for (const id of ids) subscribe(socket, id);
    await flushEventLoop();
    await vi.advanceTimersByTimeAsync(0);
    await flushEventLoop();
    expect(ids.every((id) => ydoc.getMap(PPTX_SLIDE_CONTENTS_MAP).has(id))).toBe(true);
    expect(ids.every((id) => ydoc.getMap(PPTX_SLIDE_VISUALS_MAP).has(id))).toBe(true);
    return { socket, ydoc, saveSpy };
  }

  test('슬라이드를 열기만 하면 유휴 자동저장이 돌지 않는다', async () => {
    const { saveSpy } = await openAndSeed();
    await vi.advanceTimersByTimeAsync(20_000);
    await flushEventLoop();
    expect(saveSpy).not.toHaveBeenCalled();
  });

  test('열기만 하고 나가면 종료 저장도 하지 않는다', async () => {
    const { socket, saveSpy } = await openAndSeed();
    socket.emit('close');
    await flushEventLoop(60);
    await vi.advanceTimersByTimeAsync(0);
    await flushEventLoop(60);
    expect(saveSpy).not.toHaveBeenCalled();
  });

  test('실제 편집이 있으면 나갈 때 저장한다', async () => {
    const { socket, ydoc, saveSpy } = await openAndSeed();
    ydoc.transact(() => ydoc.getMap('pptx-test-edit').set('x', 1));
    socket.emit('close');
    await flushEventLoop(60);
    await vi.advanceTimersByTimeAsync(0);
    await flushEventLoop(60);
    expect(saveSpy).toHaveBeenCalled();
  });
});
