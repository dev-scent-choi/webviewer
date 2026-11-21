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
import { MESSAGE_ROOM_EPOCH, MESSAGE_ROOM_STALE } from '../src/collab/roomEpoch.js';

/**
 * [통합] pptx 갈래의 방 세대·닫기 유예(room.ts [RX-C1], teardownRoom)와 xlsx 갈래의 방 추가물(조건부 서식 값 개수 표 cfRanges,
 * 표 규칙, 전역 LRU 상한)이 함께 맞게 도는지. 두 갈래 시험은 유예를 끈 채(vitest.config.ts) 돌아서 이 조합은 여기서만 본다.
 */
const MESSAGE_SYNC = 0;
const MESSAGE_PAGE_SUBSCRIBE = 2;
const MESSAGE_GRID_STRUCTURE_OP = 4;

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
}

const request = (epoch?: string) => ({ url: epoch ? `/ws/edit/x?token=t&epoch=${epoch}` : '/ws/edit/x?token=t' }) as never;

function sendRaw(socket: FakeSocket, type: number, text: string): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, type);
  encoding.writeVarString(encoder, text);
  socket.emit('message', encoding.toUint8Array(encoder));
}
function syncMirror(socket: FakeSocket): void {
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeSyncStep1(encoder, socket.mirror);
  socket.emit('message', encoding.toUint8Array(encoder));
}
async function flushEventLoop(turns = 50): Promise<void> {
  const { drainRoomBackgroundWorkForTest } = await import('../src/collab/room.js');
  for (let i = 0; i < turns; i += 1) {
    await new Promise((resolve) => setImmediate(resolve));
    await fs.stat(os.tmpdir()).catch(() => undefined);
    await drainRoomBackgroundWorkForTest();
  }
}

/** A열 rows행(맨 끝 행만 첫 행과 같은 값) + COUNTIF 중복 규칙 */
async function countifBook(rows: number): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  for (let r = 1; r <= rows; r += 1) ws.addRow([r === rows ? 'K1' : `K${r}`]);
  ws.addConditionalFormatting({ ref: `A1:A${rows}`, rules: [{ type: 'expression', priority: 1, formulae: ['COUNTIF($A:$A,$A1)>1'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
  return Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}
async function tableBook(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  ws.addTable({ name: 'T1', ref: 'A1', headerRow: true, totalsRow: true, columns: [{ name: '이름', totalsRowLabel: '합계' }, { name: '값', totalsRowFunction: 'sum' }, { name: '비고' }], rows: [['a', 1, 'x'], ['b', 2, 'y'], ['c', 3, 'z']] });
  return Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
}

async function saveDoc(bytes: Buffer): Promise<string> {
  const { store } = await import('../src/storage.js');
  const docId = randomUUID();
  await store.saveOriginal(docId, 'xlsx', bytes);
  await store.writeMeta({ id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid' });
  return docId;
}
async function join(docId: string, epoch?: string): Promise<FakeSocket> {
  const { joinEditSession } = await import('../src/collab/room.js');
  const socket = new FakeSocket();
  await joinEditSession(socket as never, request(epoch), docId, 'xlsx', [], false, 'viewer-a');
  if (!epoch) syncMirror(socket);
  sendRaw(socket, MESSAGE_PAGE_SUBSCRIBE, JSON.stringify({ sheet: 0, pages: [0] }));
  await flushEventLoop();
  return socket;
}
const RANGE = (rows: number) => ({ r0: 0, c0: 0, r1: rows - 1, c1: 0, crit: ['K1'] });
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

describe('[통합] 방 닫기 유예·세대 × 엑셀 방 추가물', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-integ-grace-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_ROOM_CLOSE_GRACE_MS', '400');
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

  test('값 개수 표: 유예 중에도 답하고 같은 세대로 다시 붙으면 같은 표, 유예가 끝나면 놓는다, 새 방은 새 표', async () => {
    const { getCfRemoteAnswer, getRoomCfTallyForTest, isRoomOpen } = await import('../src/collab/room.js');
    const docId = await saveDoc(await countifBook(1_500));
    const a = await join(docId);
    expect((await getCfRemoteAnswer(docId, { sheet: 0, ranges: [RANGE(1_500)] }))!.ranges[0]).toMatchObject({ cells: 1_500, crit: [[expect.any(String), 2]] });
    const tally = getRoomCfTallyForTest(docId)!;
    const rev = tally.rev;
    expect(tally.sheets.size).toBe(1);
    a.emit('close');
    await flushEventLoop();
    expect(isRoomOpen(docId)).toBe(true); // 유예 중
    // 유예 중 편집 화면의 xlsx-cf-ranges 물음(재접속 직전)에도 답한다
    expect((await getCfRemoteAnswer(docId, { sheet: 0, ranges: [RANGE(1_500)] }))!.ranges[0]).toMatchObject({ cells: 1_500 });
    const b = await join(docId, a.epoch!);
    expect(b.types).not.toContain(MESSAGE_ROOM_STALE);
    expect(b.epoch).toBe(a.epoch);
    expect(getRoomCfTallyForTest(docId)).toBe(tally); // 같은 방·같은 표 → 리비전이 이어져 편집 화면 캐시가 맞다
    expect(tally.rev).toBe(rev);
    b.emit('close');
    await sleep(700);
    await flushEventLoop();
    expect(isRoomOpen(docId)).toBe(false);
    expect(tally.sheets.size).toBe(0); // teardownRoom → cfTallyRelease
    expect(await getCfRemoteAnswer(docId, { sheet: 0, ranges: [RANGE(1_500)] })).toBeNull(); // 경로는 409
    // 지난 세대를 들고 온 탭은 거절된다(새로 불러오기) — 방을 새로 만들지 않는다
    const stale = await join(docId, a.epoch!);
    expect(stale.types).toContain(MESSAGE_ROOM_STALE);
    expect(isRoomOpen(docId)).toBe(false);
    // 새로 불러온 탭(세대 없음, 빈 상태)은 새 방·새 표로 다시 답을 받는다
    await join(docId);
    expect(getRoomCfTallyForTest(docId)).not.toBe(tally);
    expect((await getCfRemoteAnswer(docId, { sheet: 0, ranges: [RANGE(1_500)] }))!.ranges[0]).toMatchObject({ cells: 1_500 });
  }, 60_000);

  test('유예 중인 방을 문서 삭제(forceCloseRoom)로 닫으면 값 개수 표도 바로 놓는다', async () => {
    const { getCfRemoteAnswer, getRoomCfTallyForTest, isRoomOpen, forceCloseRoom } = await import('../src/collab/room.js');
    const docId = await saveDoc(await countifBook(600));
    const a = await join(docId);
    await getCfRemoteAnswer(docId, { sheet: 0, ranges: [RANGE(600)] });
    const tally = getRoomCfTallyForTest(docId)!;
    a.emit('close');
    await flushEventLoop();
    expect(isRoomOpen(docId)).toBe(true);
    forceCloseRoom(docId);
    expect(isRoomOpen(docId)).toBe(false);
    expect(tally.sheets.size).toBe(0);
  }, 60_000);

  test('전역 상한(LRU): 유예 중인 방의 표가 먼저 버려지고, 그 방에 다시 물으면 다시 세어 맞게 답한다', async () => {
    const { getCfRemoteAnswer, getRoomCfTallyForTest } = await import('../src/collab/room.js');
    const { cfEnforceGlobalCap } = await import('../src/xlsx/cfRangeValues.js');
    const d1 = await saveDoc(await countifBook(600));
    const d2 = await saveDoc(await countifBook(700));
    const s1 = await join(d1);
    await getCfRemoteAnswer(d1, { sheet: 0, ranges: [RANGE(600)] });
    s1.emit('close'); // d1은 유예 중
    await join(d2);
    await getCfRemoteAnswer(d2, { sheet: 0, ranges: [RANGE(700)] });
    const t1 = getRoomCfTallyForTest(d1)!;
    const t2 = getRoomCfTallyForTest(d2)!;
    cfEnforceGlobalCap(t2, 1); // 상한을 아주 낮춰 강제
    expect(t1.sheets.size).toBe(0);
    expect(t2.sheets.size).toBe(1);
    expect((await getCfRemoteAnswer(d1, { sheet: 0, ranges: [RANGE(600)] }))!.ranges[0]).toMatchObject({ cells: 600, crit: [[expect.any(String), 2]] });
  }, 60_000);

  test('표 규칙: 유예 뒤 같은 세대로 다시 붙어도 머리 줄만 지우는 행 삭제는 거절, 데이터 행 전부 삭제는 한 행을 남긴다', async () => {
    const { getFileExtras } = await import('../src/collab/room.js');
    const docId = await saveDoc(await tableBook());
    const a = await join(docId);
    expect((await getFileExtras(docId)).tables[0]).toMatchObject({ r0: 0, r1: 4 });
    a.emit('close');
    await flushEventLoop();
    const b = await join(docId, a.epoch!);
    sendRaw(b, MESSAGE_GRID_STRUCTURE_OP, JSON.stringify({ sheet: 0, kind: 'deleteRows', at: 0, count: 2 }));
    await flushEventLoop();
    expect((await getFileExtras(docId)).tables[0]).toMatchObject({ r0: 0, r1: 4 });
    sendRaw(b, MESSAGE_GRID_STRUCTURE_OP, JSON.stringify({ sheet: 0, kind: 'deleteRows', at: 1, count: 3 }));
    await flushEventLoop();
    expect((await getFileExtras(docId)).tables[0]).toMatchObject({ r0: 0, r1: 2, totalsRows: 1 });
  }, 60_000);
});
