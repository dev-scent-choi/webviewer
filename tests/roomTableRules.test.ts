import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1, writeUpdate } from 'y-protocols/sync';
import path from 'node:path';
import v8 from 'node:v8';
import vm from 'node:vm';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/**
 * 편집 방의 엑셀 표 규칙(검토 r5): 머리 줄만 지우는 행 삭제는 받지 않고, 데이터 행을 모두 지우면 첫 데이터 행을 비워 남기며,
 * 표 열 삭제·머리 칸 이름 바꿈·표 전체 삭제를 이 방 수식의 구조적 참조에 반영한다(편집 화면에 보이는 수식 글자도).
 * FakeSocket·편집 도우미는 roomCellChangeLog.test.ts와 같다.
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
    const messageType = decoding.readVarUint(decoder);
    if (messageType === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) {
        readUpdate(decoder, this.mirror, this);
      }
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

async function makeSampleXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  sheet.getCell('A1').value = 'Name';
  sheet.getCell('A2').value = 'Alice';
  sheet.getCell('B2').value = 10;
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

/** roomCanonicalGrid.test.ts의 sendCellEdit과 같은 clientID 고정 트릭을 쓰지만, 한 소켓이
 * **여러 번** 편집을 보내는 테스트가 있어(같은 칸을 두 번 고치는 경우 등) scratch Y.Doc을 호출마다
 * 새로 만들지 않고 소켓별로 하나씩만 만들어 재사용한다 — 매번 새 Y.Doc을 쓰면 clock이 매번 0으로
 * 리셋되는데 clientID는 그대로 고정이라, 두 번째 호출이 Yjs 입장에서는 "이미 적용된 것과 똑같은
 * (clientID, clock) 항목"이 되어 조용히 무시된다(관찰자가 아예 안 불림 — 실제로 이 버그로 처음엔
 * 테스트가 실패했다). 소켓마다 하나의 scratch를 계속 쓰면 실제 브라우저 클라이언트의 Y.Doc처럼
 * clock이 매 편집마다 자연스럽게 올라간다. */
const scratchDocs = new WeakMap<FakeSocket, Y.Doc>();
let nextScratchClientId = 0xffffffff;
function scratchDocFor(socket: FakeSocket): Y.Doc {
  let scratch = scratchDocs.get(socket);
  if (!scratch) {
    scratch = new Y.Doc();
    // 소켓마다 다른(하지만 여전히 room의 무작위 32비트 clientID보다 확실히 큰) 값을 준다 — 두
    // 소켓이 동시에 서로 다른 칸을 고치는 테스트에서, 둘 다 같은 clientID를 쓰면 clock이 각자
    // 0부터 다시 시작해 서로 다른 항목인데도 Yjs가 같은 (clientID, clock) 좌표로 오인할 수 있다.
    scratch.clientID = nextScratchClientId;
    nextScratchClientId -= 1;
    scratchDocs.set(socket, scratch);
  }
  return scratch;
}

function sendCellEdit(socket: FakeSocket, key: string, value: string): void {
  const scratch = scratchDocFor(socket);
  scratch.getMap<string>('cells').set(key, value);
  const update = Y.encodeStateAsUpdate(scratch);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, update);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** roomCanonicalGrid.test.ts의 동명 헬퍼와 동일 — Task 7(윈도잉 컷오버) 이후 Yjs cells 맵은
 * "구독 중인 페이지"만 담으므로, 원본에 이미 있던 값(예: "Alice")을 지우는 걸 재현하려면 먼저 그
 * 페이지를 구독해 room이 그 값을 cells 맵에 시딩하게 해야 한다 — 안 그러면 mirror도 room도
 * 그 키를 아예 모르는 상태라 delete가 아무것도 지우지 못한다. */
function sendSubscribe(socket: FakeSocket, sheet: number, pages: readonly number[]): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ sheet, pages }));
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** roomCanonicalGrid.test.ts의 동명 헬퍼와 동일 — 접속 직후 클라이언트가 자기 상태 벡터(비어
 * 있음)를 보내는 것을 흉내낸다. socket.mirror가 room.ydoc과 같은 아이템 식별자를 공유해야만
 * 그 뒤 mirror에서 만든 delete가 room의 실제 항목을 지우는 delete로 정확히 병합된다. */
function syncMirrorWithRoom(socket: FakeSocket): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeSyncStep1(encoder, socket.mirror);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** 서로 다른 소켓이 **같은 칸**을 순서대로 고치는 시나리오에서 sendCellEdit(독립 scratch,
 * clientID로 승부)를 쓰면 안 된다 — 두 스크래치가 서로의 존재를 몰라 Yjs가 "동시 편집"으로
 * 보고 clientID가 더 큰 쪽을 무조건 이기게 해서, 실제로 나중에 온 편집이 조용히 버려질 수
 * 있다(Y.Map의 동시-편집 충돌 해소 규칙 — 적용 순서가 아니라 clientID 크기로 승자를 정한다).
 * `syncMirrorWithRoom`을 먼저 불러 mirror가 room의 "지금까지" 상태(이전 편집 포함)를 안 뒤에
 * 편집해야, 이 편집이 그 이전 편집의 인과적 후속으로 인식되어 항상(clientID와 무관하게) 이긴다.
 * deleteCellViaClient와 같은 이유로 전체 상태가 아니라 델타만 캡처해 보낸다. */
function editCellViaClient(socket: FakeSocket, key: string, value: string): void {
  let captured: Uint8Array | null = null;
  const onUpdate = (update: Uint8Array) => {
    captured = update;
  };
  socket.mirror.on('update', onUpdate);
  socket.mirror.getMap<string>('cells').set(key, value);
  socket.mirror.off('update', onUpdate);
  if (!captured) {
    throw new Error('edit produced no update');
  }
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, captured);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** roomCanonicalGrid.test.ts의 동명 헬퍼와 동일 — 실제 클라이언트(grid-editor.ts의 commitCell)가
 * 칸을 지울 때 하는 것과 같은 동작(cells.delete(key), cells.set(key,'')가 아니다)을 mirror에서
 * 수행하고, 그 결과로 생긴 update만(전체 상태가 아니라 델타만) 캡처해 서버로 보낸다. */
function deleteCellViaClient(socket: FakeSocket, key: string): void {
  let captured: Uint8Array | null = null;
  const onUpdate = (update: Uint8Array) => {
    captured = update;
  };
  socket.mirror.on('update', onUpdate);
  socket.mirror.getMap<string>('cells').delete(key);
  socket.mirror.off('update', onUpdate);
  if (!captured) {
    throw new Error('delete produced no update — was syncMirrorWithRoom called first?');
  }
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, captured);
  socket.emit('message', encoding.toUint8Array(encoder));
}

async function drainPendingDiskIo(): Promise<void> {
  await fs.stat(os.tmpdir()).catch(() => undefined);
}

async function flushEventLoop(turns = 50): Promise<void> {
  const { drainRoomBackgroundWorkForTest } = await import('../src/collab/room.js');
  for (let i = 0; i < turns; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
    await drainPendingDiskIo();
    await drainRoomBackgroundWorkForTest();
  }
}


const MESSAGE_GRID_STRUCTURE_OP = 4;
function sendStructureOp(socket: FakeSocket, payload: unknown): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_GRID_STRUCTURE_OP);
  encoding.writeVarString(encoder, JSON.stringify(payload));
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** T1 = A1:C5(머리 줄 1행, 데이터 2~4행, 요약 5행) */
async function tableBook(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.addTable({ name: 'T1', ref: 'A1', headerRow: true, totalsRow: true, columns: [{ name: '이름', totalsRowLabel: '합계' }, { name: '값', totalsRowFunction: 'sum' }, { name: '비고' }], rows: [['a', 1, 'x'], ['b', 2, 'y'], ['c', 3, 'z']] });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function openRoom(bytes: Buffer) {
  const { store } = await import('../src/storage.js');
  const { joinEditSession } = await import('../src/collab/room.js');
  const docId = randomUUID();
  await store.saveOriginal(docId, 'xlsx', bytes);
  await store.writeMeta({ id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid' });
  const socket = new FakeSocket();
  await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false, 'viewer-a');
  sendSubscribe(socket, 0, [0]);
  await flushEventLoop();
  syncMirrorWithRoom(socket);
  await flushEventLoop();
  return { docId, socket };
}

describe('collab/room.ts: 엑셀 표 규칙(행/열 삭제·머리 칸 이름 바꿈)', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-tables-'));
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

  const tableXml = async (docId: string) => {
    const { getCurrentBytes } = await import('../src/collab/room.js');
    const { strFromU8, unzipSync } = await import('fflate');
    const files = unzipSync((await getCurrentBytes(docId))!);
    return Object.entries(files)
      .filter(([n]) => /^xl\/tables\/table\d+\.xml$/.test(n))
      .map(([, b]) => strFromU8(b));
  };

  test('옛 화면이 데이터 행을 모두 지우라고 보내도: 첫 데이터 행은 비워 남기고 요약 행은 그대로(저장·화면 같음)', async () => {
    const { getFileExtras } = await import('../src/collab/room.js');
    const { docId, socket } = await openRoom(await tableBook());
    sendStructureOp(socket, { sheet: 0, kind: 'deleteRows', at: 1, count: 3 });
    await flushEventLoop();
    const t = (await getFileExtras(docId)).tables[0];
    expect(t).toMatchObject({ r0: 0, r1: 2, totalsRows: 1 });
    expect(socket.mirror.getMap('cells').get('0:1:0')).toBeUndefined(); // 비운 데이터 행
    expect(socket.mirror.getMap('cells').get('0:2:0')).toBe('합계');
    const [xml] = await tableXml(docId);
    expect(xml).toMatch(/<table [^>]*ref="A1:C3"/);
    expect(xml).toMatch(/totalsRowCount="1"/);
  }, 60_000);

  test('머리 줄만 지우는 행 삭제는 받지 않는다(엑셀과 같다)', async () => {
    const { getFileExtras } = await import('../src/collab/room.js');
    const { docId, socket } = await openRoom(await tableBook());
    sendStructureOp(socket, { sheet: 0, kind: 'deleteRows', at: 0, count: 2 });
    await flushEventLoop();
    expect((await getFileExtras(docId)).tables[0]).toMatchObject({ r0: 0, r1: 4 });
    expect(socket.mirror.getMap('cells').get('0:0:0')).toBe('이름');
  }, 60_000);

  test('표 열 삭제·머리 칸 이름 바꿈·표 전체 삭제가 이 방 수식(편집 화면에 보이는 글자)에 반영된다', async () => {
    const { gridCellKey } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(await tableBook());
    const cells = socket.mirror.getMap<string>('cells');
    editCellViaClient(socket, gridCellKey(0, 6, 5), "=SUM(T1[값])");
    editCellViaClient(socket, gridCellKey(0, 7, 5), "=COUNTA(T1[비고])");
    await flushEventLoop();
    // 머리 칸 C1 "비고" → "메모"
    syncMirrorWithRoom(socket);
    editCellViaClient(socket, gridCellKey(0, 0, 2), '메모');
    await flushEventLoop();
    expect(cells.get(gridCellKey(0, 7, 5))).toBe('=COUNTA(T1[메모])');
    // 값 열(B) 삭제 → F열이 E열로, T1[값] → #REF!
    sendStructureOp(socket, { sheet: 0, kind: 'deleteCols', at: 1, count: 1 });
    await flushEventLoop();
    expect(cells.get(gridCellKey(0, 6, 4))).toBe('=SUM(#REF!)');
    expect(cells.get(gridCellKey(0, 7, 4))).toBe('=COUNTA(T1[메모])');
    // 표 전체(1~5행) 삭제 → 표가 없어지고 참조는 #REF!
    sendStructureOp(socket, { sheet: 0, kind: 'deleteRows', at: 0, count: 5 });
    await flushEventLoop();
    expect(await tableXml(docId)).toEqual([]);
    const all = [...cells.values()].filter((v) => String(v).startsWith('='));
    expect(all).toEqual(expect.arrayContaining(['=COUNTA(#REF!)']));
  }, 60_000);
});
