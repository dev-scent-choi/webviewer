import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as decoding from 'lib0/decoding';
import * as encoding from 'lib0/encoding';
import * as Y from 'yjs';
import { messageYjsSyncStep2, messageYjsUpdate, readUpdate, writeSyncStep1 } from 'y-protocols/sync';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

const MESSAGE_SYNC = 0;
const MESSAGE_PAGE_SUBSCRIBE = 2;
const MESSAGE_GRID_STRUCTURE_OP = 4;

class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  readonly mirror = new Y.Doc();

  send(data: Uint8Array, cb?: (err?: Error) => void): void {
    const decoder = decoding.createDecoder(data);
    if (decoding.readVarUint(decoder) === MESSAGE_SYNC) {
      const syncType = decoding.readVarUint(decoder);
      if (syncType === messageYjsSyncStep2 || syncType === messageYjsUpdate) readUpdate(decoder, this.mirror, this);
    }
    cb?.();
  }
  close(): void {}
  ping(): void {}
}

function sendSubscribe(socket: FakeSocket, payload: unknown): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify(payload));
  socket.emit('message', encoding.toUint8Array(encoder));
}

/** 클라이언트의 sendGridStructureOp와 완전히 같은 인코딩(varUint 메시지 타입 + varString JSON). */
function sendStructureOp(socket: FakeSocket, payload: unknown): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_GRID_STRUCTURE_OP);
  encoding.writeVarString(encoder, JSON.stringify(payload));
  socket.emit('message', encoding.toUint8Array(encoder));
}

async function flushEventLoop(turns = 80): Promise<void> {
  for (let i = 0; i < turns; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
    await drainPendingDiskIo();
    await drainRoomBackgroundWork();
  }
}

/**
 * 서버가 흘려보낸(fire-and-forget) 백그라운드 작업 — 페이지 창 반영, 수식 재계산, 셀/서식 편집의
 * canonical 되반영 — 이 실제로 끝날 때까지 기다린다. 턴 수 같은 **대리 지표가 아니라 그 작업의
 * 완료 자체**를 기다리는 것이 핵심이다(room.ts의 Room.backgroundWork 주석 참고).
 *
 * 룸 모듈을 아직(또는 전혀) 불러오지 않은 테스트에서도 안전하도록 조용히 넘어간다.
 */
async function drainRoomBackgroundWork(): Promise<void> {
  try {
    const { drainRoomBackgroundWorkForTest } = await import('../src/collab/room.js');
    await drainRoomBackgroundWorkForTest();
  } catch {
    /* 이 테스트는 룸을 쓰지 않는다 */
  }
}


/**
 * setImmediate 턴만 흘려보내는 것으로는 **진짜 디스크 I/O를 기다릴 수 없다**.
 *
 * 예전에는 룸을 열 때 `seedGridStylesFromXlsx`가 워크북 전체를 훑으며 `LazyGridStore.readCache`를
 * 통째로 데워 놨고 `joinEditSession`이 그 완료를 기다렸다 — 그래서 페이지를 구독하는 시점에는 그
 * 페이지의 모든 행이 이미 메모리에 있었고 시딩이 디스크를 건드리지 않았다. 그 조건에서만
 * "setImmediate 턴 N번"이 충분한 대기였다. 서식이 콘텐츠와 함께 페이지 단위로 윈도잉되면서 그
 * 선점적 전량 데우기가 사라졌고, 이제 페이지를 처음 구독하면 실제로 디스크를 읽는다.
 *
 * setImmediate 턴은 **벽시계 시간을 사주지 않는다**(실측: 500턴이 약 3ms — 턴마다 poll 타임아웃 0으로
 * 그냥 스핀한다). 그래서 턴 수를 늘리는 것은 해결책이 아니라 더 조용한 플래키의 원인이다 — 턴 수는
 * 예전 그대로 두고, 대신 턴마다 실제 디스크 작업을 한 번 흘려보낸다. `fs.stat`은 행 읽기와 같은
 * libuv 스레드풀의 FIFO 작업 큐를 쓰므로, 이 stat이 끝났다는 것은 그 앞에 쌓여 있던 읽기들이 이미
 * 큐에서 빠져나갔다는 뜻이다(임의의 sleep이 아니라 의미가 분명한 대기다). 타이머를 쓰지 않으므로
 * 가짜 타이머를 켠 테스트에서도 안전하다. 자세한 배경은 roomWindowedSeeding.test.ts의 같은 헬퍼 참고.
 */
async function drainPendingDiskIo(): Promise<void> {
  await fs.stat(os.tmpdir()).catch(() => undefined);
}


function syncMirrorWithRoom(socket: FakeSocket): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeSyncStep1(encoder, socket.mirror);
  socket.emit('message', encoding.toUint8Array(encoder));
}

/**
 * 행/열 연산은 canonical 갱신 체인에서 돌고(enqueueCanonicalGrowth) 그 안에서 페이지 재시딩·수식
 * 재계산이 다시 청크마다 양보하므로, 고정된 틱 수만 흘려보내는 것으로는 완료를 보장할 수 없다 —
 * 체인이 실제로 비는 것을 기다린 뒤 남은 후속 작업까지 한 번 더 흘려보낸다.
 */
async function settleStructureOp(docId: string): Promise<void> {
  const { awaitCanonicalGrowthForTest } = await import('../src/collab/room.js');
  await awaitCanonicalGrowthForTest(docId);
  await flushEventLoop();
  await awaitCanonicalGrowthForTest(docId);
}

async function seedDoc(docId: string, original: Buffer): Promise<void> {
  const { store } = await import('../src/storage.js');
  await store.saveOriginal(docId, 'xlsx', original);
  await store.writeMeta({
    id: docId,
    name: 'x.xlsx',
    ext: 'xlsx',
    status: 'ready',
    pageCount: 1,
    createdAt: new Date().toISOString(),
    progress: 100,
    viewKind: 'grid',
  });
}

async function downloadWorkbook(docId: string): Promise<ExcelJS.Workbook> {
  const { getCurrentBytes } = await import('../src/collab/room.js');
  const bytes = await getCurrentBytes(docId);
  expect(bytes).not.toBeNull();
  const reopened = new ExcelJS.Workbook();
  await reopened.xlsx.load(bytes! as unknown as ArrayBuffer);
  return reopened;
}


async function makeXlsx(): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= 20; r += 1) for (let c = 1; c <= 4; c += 1) sheet.getCell(r, c).value = `${r}-${c}`;
  sheet.getColumn(4).width = 30;
  sheet.getRow(5).height = 33;
  return Buffer.from((await workbook.xlsx.writeBuffer()) as ArrayBuffer);
}

/** 열 폭·행 높이 끌어 바꾸기(MESSAGE_GRID_STRUCTURE_OP의 resizeCols/resizeRows) — 화면(구조 맵)과 저장 파일 둘 다 */
describe('room.ts: 열 폭·행 높이 바꾸기', () => {
  let tmpDir: string;

  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-grid-resize-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.doMock('../src/converters/index.js', () => ({ findConverter: () => null }));
    vi.doMock('../src/audit.js', () => ({ recordAudit: async () => undefined }));
    vi.useFakeTimers({ toFake: ['setTimeout', 'clearTimeout', 'setInterval', 'clearInterval'] });
  });

  afterEach(async () => {
    vi.doUnmock('../src/converters/index.js');
    vi.doUnmock('../src/audit.js');
    vi.unstubAllEnvs();
    vi.useRealTimers();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  async function openRoom(readOnly = false): Promise<{ docId: string; socket: FakeSocket }> {
    const { joinEditSession } = await import('../src/collab/room.js');
    const docId = randomUUID();
    await seedDoc(docId, await makeXlsx());
    const socket = new FakeSocket();
    await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], readOnly);
    sendSubscribe(socket, { sheet: 0, pages: [0] });
    await flushEventLoop();
    syncMirrorWithRoom(socket);
    return { docId, socket };
  }

  const structureOf = (socket: FakeSocket) =>
    socket.mirror.getMap<{ colWidths?: number[]; rowHeights?: Array<[number, number, number]> }>('structure').get('0');

  test('바꾼 열·행만 파일에 쓰고, 화면용 구조 맵도 바뀐다', async () => {
    const { excelColWidthToPx } = await import('../src/xlsx/sheetGeometry.js');
    const { docId, socket } = await openRoom();
    const before = await downloadWorkbook(docId);
    const untouchedWidth = before.worksheets[0].getColumn(4).width;
    sendStructureOp(socket, { sheet: 0, kind: 'resizeCols', at: 1, count: 2, size: 150 });
    sendStructureOp(socket, { sheet: 0, kind: 'resizeRows', at: 2, count: 1, size: 40 });
    sendStructureOp(socket, { sheet: 0, kind: 'resizeRows', at: 7, count: 1, size: 0 });
    await settleStructureOp(docId);
    syncMirrorWithRoom(socket);
    const s = structureOf(socket);
    expect(s?.colWidths?.[1]).toBe(150);
    expect(s?.colWidths?.[2]).toBe(150);
    expect(s?.rowHeights).toEqual(expect.arrayContaining([[2, 1, 40], [7, 1, 0]]));

    const ws = (await downloadWorkbook(docId)).worksheets[0];
    expect(excelColWidthToPx(ws.getColumn(2).width!, 7)).toBe(150);
    expect(excelColWidthToPx(ws.getColumn(3).width!, 7)).toBe(150);
    expect(ws.getColumn(4).width).toBe(untouchedWidth); // 안 건드린 열은 원본 그대로
    expect(ws.getRow(3).height).toBe(30);
    expect(ws.getRow(8).hidden).toBe(true);
    expect(ws.getRow(5).height).toBe(33); // 안 건드린 행 높이도 그대로
    expect(ws.getCell('B3').value).toBe('3-2');
  });

  test('바꾼 뒤 앞에 열·행을 끼우면 기록도 같이 밀린다', async () => {
    const { excelColWidthToPx } = await import('../src/xlsx/sheetGeometry.js');
    const { docId, socket } = await openRoom();
    sendStructureOp(socket, { sheet: 0, kind: 'resizeCols', at: 1, count: 1, size: 200 });
    sendStructureOp(socket, { sheet: 0, kind: 'resizeRows', at: 3, count: 1, size: 50 });
    sendStructureOp(socket, { sheet: 0, kind: 'insertCols', at: 0, count: 1 });
    sendStructureOp(socket, { sheet: 0, kind: 'insertRows', at: 0, count: 2 });
    await settleStructureOp(docId);
    const ws = (await downloadWorkbook(docId)).worksheets[0];
    expect(ws.getCell('C3').value).toBe('1-2');
    expect(excelColWidthToPx(ws.getColumn(3).width!, 7)).toBe(200);
    expect(ws.getRow(6).height).toBe(37.5);
  });

  test('읽기 전용 연결·범위를 벗어난 값은 버린다', async () => {
    const ro = await openRoom(true);
    sendStructureOp(ro.socket, { sheet: 0, kind: 'resizeCols', at: 0, count: 1, size: 300 });
    await settleStructureOp(ro.docId);
    syncMirrorWithRoom(ro.socket);
    expect(structureOf(ro.socket)?.colWidths?.[0]).not.toBe(300);

    const rw = await openRoom();
    sendStructureOp(rw.socket, { sheet: 0, kind: 'resizeCols', at: 0, count: 1, size: 99999 });
    sendStructureOp(rw.socket, { sheet: 0, kind: 'resizeRows', at: 0, count: 1, size: 12.5 });
    sendStructureOp(rw.socket, { sheet: 0, kind: 'resizeCols', at: 100000, count: 1, size: 50 });
    await settleStructureOp(rw.docId);
    syncMirrorWithRoom(rw.socket);
    expect(structureOf(rw.socket)?.colWidths?.[0]).not.toBe(99999);
    expect(structureOf(rw.socket)?.colWidths?.length ?? 0).toBeLessThan(1000);
  });
});
