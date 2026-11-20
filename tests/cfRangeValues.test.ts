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
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/**
 * 편집 화면 조건부 서식용 범위 값 개수 표(getCfRangeValues — 검토 r3 E2). 편집 화면은 보고 있는 페이지만 받으므로
 * COUNTIF($A:$A,$A1)>1 같은 규칙은 서버가 센 표로 판정한다. 편집으로 중복이 생기고 없어지는 것이 반영돼야 한다.
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


async function makeColumnXlsx(rows: number): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  // A열: 모두 다른 값, 맨 끝 행만 첫 행과 같은 값(멀리 떨어진 중복 — 편집 화면이 그 페이지를 받지 않은 상태)
  for (let r = 1; r <= rows; r += 1) sheet.getCell(r, 1).value = r === rows ? 'K1' : `K${r}`;
  const buffer = await workbook.xlsx.writeBuffer();
  return Buffer.from(buffer as ArrayBuffer);
}

describe('collab/room.ts: getCfRangeValues (편집 화면 조건부 서식 범위 표)', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-cfranges-'));
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

  for (const rows of [1_500, 20_000]) {
    test(`${rows}행: 불러오지 않은 행의 중복까지 세고, 편집으로 중복이 생기고 없어지면 반영한다`, async () => {
      const { store } = await import('../src/storage.js');
      const { joinEditSession, getCfRangeValues } = await import('../src/collab/room.js');
      const { gridCellKey } = await import('../src/collab/gridShared.js');
      const { cfRangeTableFromCounts, paintCell } = await import('../src/collab/condFormatShared.js');
      const docId = randomUUID();
      await store.saveOriginal(docId, 'xlsx', await makeColumnXlsx(rows));
      await store.writeMeta({ id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid' });
      const socket = new FakeSocket();
      await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false, 'viewer-a');
      sendSubscribe(socket, 0, [0]); // 편집 화면처럼 첫 페이지(500행)만 받는다
      await flushEventLoop();

      const range = { r0: 0, c0: 0, r1: 1_048_575, c1: 0 }; // $A:$A — 시트 크기로 잘린다
      const first = await getCfRangeValues(docId, 0, [range], undefined);
      expect(first?.ranges?.[0]).toMatchObject({ cells: rows });
      // 편집 화면과 같은 판정: COUNTIF($A:$A,$A1)>1 — 1행은 맨 끝 행과 중복, 2행은 아니다
      const rule = { id: 'd', ranges: [{ r0: 0, c0: 0, r1: rows - 1, c1: 0 }], rule: { t: 'expr' as const, f: 'COUNTIF($A:$A,$A1)>1', s: { bg: '#FFC7CE' } } };
      const painted = (res: typeof first, row: number, value: string) => {
        const r = res!.ranges![0] as unknown as { cells: number; values: Array<[string, number]> };
        const table = cfRangeTableFromCounts(r.cells, r.values);
        const env = { tables: new Map(), extent: () => ({ rows, cols: 1 }), remote: () => table };
        return paintCell([rule], row, 0, value, new Map(), () => value, env)?.style?.bg;
      };
      expect(painted(first, 0, 'K1')).toBe('#FFC7CE');
      expect(painted(first, 1, 'K2')).toBeUndefined();

      // 리비전이 같으면 표를 다시 보내지 않는다
      expect(await getCfRangeValues(docId, 0, [range], first!.rev)).toEqual({ rev: first!.rev });

      // 편집으로 중복을 만든다: A2 = "K3" → K3이 둘
      editCellViaClient(socket, gridCellKey(0, 1, 0), 'K3');
      await flushEventLoop();
      const second = await getCfRangeValues(docId, 0, [range], first!.rev);
      expect(second!.rev).not.toBe(first!.rev);
      expect(painted(second, 2, 'K3')).toBe('#FFC7CE');
      // 편집으로 중복을 없앤다: 맨 끝 행(받지 않은 페이지)을 지운다 → K1은 하나뿐
      syncMirrorWithRoom(socket);
      editCellViaClient(socket, gridCellKey(0, rows - 1, 0), 'Z');
      await flushEventLoop();
      const third = await getCfRangeValues(docId, 0, [range], second!.rev);
      expect(painted(third, 0, 'K1')).toBeUndefined();
      expect(painted(third, 2, 'K3')).toBe('#FFC7CE');
    }, 60_000);
  }

  test('범위 칸이 너무 많으면 건너뛰고(skipped), 요청 모양이 틀리면 거른다', async () => {
    const { isCfRangeRequest, CF_RANGE_CELL_CAP } = await import('../src/xlsx/cfRangeValues.js');
    expect(isCfRangeRequest({ sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 5, c1: 0 }] })).toBe(true);
    expect(isCfRangeRequest({ sheet: 0, ranges: [{ r0: 5, c0: 0, r1: 1, c1: 0 }] })).toBe(false);
    expect(isCfRangeRequest({ sheet: -1, ranges: [{ r0: 0, c0: 0, r1: 1, c1: 0 }] })).toBe(false);
    expect(isCfRangeRequest({ sheet: 0, ranges: Array.from({ length: 65 }, () => ({ r0: 0, c0: 0, r1: 1, c1: 0 })) })).toBe(false);
    expect(isCfRangeRequest({ sheet: 0, ranges: [{ r0: '0', c0: 0, r1: 1, c1: 0 }] })).toBe(false);
    const { cfRangeValues } = await import('../src/xlsx/cfRangeValues.js');
    const fakeStore = { sheetMeta: [{ rowCount: 1_000_000, colCount: 10 }] } as never;
    const out = await cfRangeValues(fakeStore, 0, [{ r0: 0, c0: 0, r1: 999_999, c1: 9 }], () => undefined, { rev: 0, byKey: new Map(), applying: Promise.resolve() });
    expect(10_000_000 > CF_RANGE_CELL_CAP).toBe(true);
    expect(out[0]).toMatchObject({ skipped: true });
  });
});
