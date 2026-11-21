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


/** cols열 × rows행. A열은 모두 다른 값이고 맨 끝 행만 첫 행과 같다(편집 화면이 받지 않는 페이지의 중복). 나머지 열도 모두 다른 값 */
async function makeSheetXlsx(rows: number, cols: number, rules: Array<{ ref: string; formula?: string; dup?: boolean }>): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Sheet1');
  for (let r = 1; r <= rows; r += 1) {
    const line: string[] = [];
    for (let c = 1; c <= cols; c += 1) line.push(c === 1 ? (r === rows ? 'K1' : `K${r}`) : `v${r}_${c}`);
    sheet.addRow(line);
  }
  const fill = { type: 'pattern' as const, pattern: 'solid' as const, bgColor: { argb: 'FFFFC7CE' } };
  rules.forEach((rule, i) => {
    if (!rule.dup) sheet.addConditionalFormatting({ ref: rule.ref, rules: [{ type: 'expression', priority: i + 1, formulae: [rule.formula!], style: { fill } }] });
  });
  const zip = unzipSync(new Uint8Array(await workbook.xlsx.writeBuffer()));
  const dups = rules.filter((r) => r.dup);
  if (dups.length > 0) {
    // ExcelJS는 중복 값 규칙을 쓰지 못한다 — 시트 XML에 직접 넣는다
    const xml = dups.map((r) => `<conditionalFormatting sqref="${r.ref}"><cfRule type="duplicateValues" dxfId="0" priority="1"/></conditionalFormatting>`).join('');
    zip['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(zip['xl/worksheets/sheet1.xml']).replace('</sheetData>', `</sheetData>${xml}`));
    zip['xl/styles.xml'] = strToU8(strFromU8(zip['xl/styles.xml']).replace(/<dxfs count="0"\/>|(?=<tableStyles)/, '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'));
  }
  return Buffer.from(zipSync(zip));
}

async function openRoom(bytes: Buffer) {
  const { store } = await import('../src/storage.js');
  const { joinEditSession } = await import('../src/collab/room.js');
  const docId = randomUUID();
  await store.saveOriginal(docId, 'xlsx', bytes);
  await store.writeMeta({ id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid' });
  const socket = new FakeSocket();
  await joinEditSession(socket as never, {} as never, docId, 'xlsx', [], false, 'viewer-a');
  sendSubscribe(socket, 0, [0]); // 편집 화면처럼 첫 페이지(500행)만 받는다
  await flushEventLoop();
  return { docId, socket };
}

/** 힙 정리 — 실행 옵션 없이 이 프로세스에서 켠다 */
function forceGc(): () => void {
  v8.setFlagsFromString('--expose-gc');
  return vm.runInNewContext('gc') as () => void;
}

/** 이벤트 루프가 얼마나 오래 막혔는지 — 5ms 간격 타이머가 늦은 만큼 */
function stallMeter(): { stop(): number } {
  let last = performance.now();
  let worst = 0;
  const timer = setInterval(() => {
    const now = performance.now();
    worst = Math.max(worst, now - last - 5);
    last = now;
  }, 5);
  return {
    stop() {
      clearInterval(timer);
      return worst;
    },
  };
}

describe('collab/room.ts: getCfRemoteAnswer (편집 화면 조건부 서식 서버 답)', () => {
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
    test(`${rows}행 COUNTIF($A:$A,$A1)>1: 받지 않은 행의 중복을 세고, 편집으로 생기고 없어지면 −1/+1로 따라간다`, async () => {
      const { getCfRemoteAnswer } = await import('../src/collab/room.js');
      const { gridCellKey } = await import('../src/collab/gridShared.js');
      const { cfCritKey, paintCell } = await import('../src/collab/condFormatShared.js');
      const { docId, socket } = await openRoom(await makeSheetXlsx(rows, 1, [{ ref: `A1:A${rows}`, formula: 'COUNTIF($A:$A,$A1)>1' }]));
      const range = { r0: 0, c0: 0, r1: rows - 1, c1: 0 }; // 편집 화면 rangeTableOf 키(시트 크기로 자른 $A:$A)
      const ask = (crit: string[]) => getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ ...range, crit }] });
      const painted = async (row: number, value: string) => {
        const a = (await ask([value]))!.ranges[0];
        if ('skipped' in a) throw new Error('skipped');
        const crit = new Map(a.crit);
        const rule = { id: 'd', ranges: [{ r0: 0, c0: 0, r1: rows - 1, c1: 0 }], rule: { t: 'expr' as const, f: 'COUNTIF($A:$A,$A1)>1', s: { bg: '#FFC7CE' } } };
        const env = { tables: new Map(), extent: () => ({ rows, cols: 1 }), remote: () => ({ cells: a.cells, remote: { agg: a.agg, crit, ask: () => undefined } }) };
        return paintCell([rule], row, 0, value, new Map(), () => value, env)?.style?.bg;
      };
      const first = await ask(['K1', 'K2']);
      expect(first!.ranges[0]).toMatchObject({ cells: rows, crit: [[cfCritKey('K1'), 2], [cfCritKey('K2'), 1]] });
      expect(await painted(0, 'K1')).toBe('#FFC7CE');
      expect(await painted(1, 'K2')).toBeUndefined();
      // 편집으로 중복을 만든다(A2 = K3)
      editCellViaClient(socket, gridCellKey(0, 1, 0), 'K3');
      await flushEventLoop();
      expect(await painted(2, 'K3')).toBe('#FFC7CE');
      // 받지 않은 맨 끝 행을 고쳐 K1 중복을 없앤다
      syncMirrorWithRoom(socket);
      editCellViaClient(socket, gridCellKey(0, rows - 1, 0), 'Z');
      await flushEventLoop();
      expect(await painted(0, 'K1')).toBeUndefined();
      expect((await ask(['Z']))!.ranges[0]).toMatchObject({ crit: [[cfCritKey('Z'), 1]] });
    }, 120_000);
  }

  test('규칙에 없는 범위는 답하지 않는다(skipped) — 범위를 한 행씩 밀어도 세지 않는다', async () => {
    const { getCfRemoteAnswer } = await import('../src/collab/room.js');
    const { docId } = await openRoom(await makeSheetXlsx(1_500, 1, [{ ref: 'A1:A1500', formula: 'COUNTIF($A:$A,$A1)>1' }]));
    const res = await getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ r0: 1, c0: 0, r1: 1_499, c1: 0 }, { r0: 0, c0: 0, r1: 1_499, c1: 0 }] });
    expect(res!.ranges[0]).toMatchObject({ skipped: true });
    expect(res!.ranges[1]).toMatchObject({ cells: 1_500 });
  }, 60_000);

  test('중복 값 규칙: 2개 이상인 값만 보낸다(작은 통계)', async () => {
    const { getCfRemoteAnswer } = await import('../src/collab/room.js');
    const { paintCell } = await import('../src/collab/condFormatShared.js');
    const { docId } = await openRoom(await makeSheetXlsx(1_500, 1, [{ ref: 'A1:A1500', dup: true }]));
    const { GRID_COND_FORMATS_MAP } = await import('../src/collab/condFormatShared.js');
    const { getRoomYdocForTest } = await import('../src/collab/room.js');
    const entries = getRoomYdocForTest(docId)!.getMap<never[]>(GRID_COND_FORMATS_MAP).get('0') as unknown as Array<{ id: string }>;
    const id = entries[0].id;
    const res = await getCfRemoteAnswer(docId, { sheet: 0, ranges: [], stats: [id] });
    const st = res!.stats[id];
    expect(st).not.toBe('skipped');
    expect(st).toMatchObject({ dupOnly: true, counts: [['k1', 2]] });
    const stats = new Map([[id, { counts: new Map((st as { counts: Array<[string, number]> }).counts), dupOnly: true }]]);
    const rule = entries[0] as never;
    expect(paintCell([rule], 0, 0, 'K1', stats, () => '')?.style?.bg).toBeDefined();
    expect(paintCell([rule], 5, 0, 'K6', stats, () => '')).toBeNull();
  }, 60_000);

  test('C1: 2만 행 × 10열(모두 다른 값) — 64범위 × 동시 4요청이 상한 안(메모리·멈춤)', async () => {
    const { getCfRemoteAnswer } = await import('../src/collab/room.js');
    const { CF_ROOM_CELL_CAP, CF_ROOM_DISTINCT_CAP } = await import('../src/xlsx/cfRangeValues.js');
    const { docId } = await openRoom(await makeSheetXlsx(20_000, 10, [{ ref: 'A1:J20000', formula: 'COUNTIF($A:$J,A1)>1' }]));
    const gc = forceGc();
    gc();
    const heap0 = process.memoryUsage().heapUsed;
    let peak = heap0;
    const sampler = setInterval(() => (peak = Math.max(peak, process.memoryUsage().heapUsed)), 10);
    const meter = stallMeter();
    // 거의 시트 전체인 범위를 한 행씩 밀어 64개 — 규칙 범위($A:$J, 20만 칸) 하나만 답한다
    const ranges = Array.from({ length: 64 }, (_, i) => ({ r0: i, c0: 0, r1: 19_999, c1: 9, crit: [`K${i + 1}`] }));
    const t0 = performance.now();
    const answers = await Promise.all(Array.from({ length: 4 }, () => getCfRemoteAnswer(docId, { sheet: 0, ranges })));
    const ms = performance.now() - t0;
    const stall = meter.stop();
    clearInterval(sampler);
    for (const a of answers) {
      expect(a!.ranges[0]).toMatchObject({ cells: 200_000 });
      expect(a!.ranges.slice(1).every((r) => 'skipped' in r)).toBe(true);
      expect(JSON.stringify(a).length).toBeLessThan(20_000); // 범위 값 표를 보내지 않는다
    }
    const grew = (peak - heap0) / 1e6;
    console.log(`C1: 4요청 ${Math.round(ms)}ms, 힙 최고 +${grew.toFixed(1)}MB, 가장 긴 멈춤 ${Math.round(stall)}ms (상한 ${CF_ROOM_CELL_CAP}칸·${CF_ROOM_DISTINCT_CAP}값)`);
    expect(grew).toBeLessThan(150);
    expect(stall).toBeLessThan(100);
  }, 120_000);

  test('I1: 20만 칸 범위에서 편집 50번 — 다시 세지 않고(빠름) 메모리가 늘지 않는다', async () => {
    const { getCfRemoteAnswer, getRoomCfTallyForTest } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');
    const { docId, socket } = await openRoom(await makeSheetXlsx(20_000, 10, [{ ref: 'A1:J20000', formula: 'COUNTIF($A:$J,A1)>1' }]));
    const range = { r0: 0, c0: 0, r1: 19_999, c1: 9 };
    await getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ ...range, crit: ['K1'] }] }); // 처음 한 번 센다
    const tally = () => [...getRoomCfTallyForTest(docId)!.sheets.get(0)!.rects.values()][0].raw!;
    const distinct0 = tally().size;
    const gc = forceGc();
    gc();
    const heap0 = process.memoryUsage().heapUsed;
    let worst = 0;
    for (let i = 0; i < 50; i += 1) {
      syncMirrorWithRoom(socket);
      editCellViaClient(socket, gridCellKey(0, 1 + (i % 10), 1), i % 2 === 0 ? 'K1' : `v${2 + (i % 10)}_2`);
      await flushEventLoop(5);
      const t0 = performance.now();
      const a = await getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ ...range, crit: ['K1'] }] });
      worst = Math.max(worst, performance.now() - t0);
      expect(a!.ranges[0]).toMatchObject({ cells: 200_000 });
    }
    gc();
    const grew = (process.memoryUsage().heapUsed - heap0) / 1e6;
    console.log(`I1: 편집 50번, 답 가장 느린 것 ${Math.round(worst)}ms, 값 종류 ${distinct0}→${tally().size}, 힙 ${grew.toFixed(1)}MB`);
    expect(tally().size).toBeLessThanOrEqual(distinct0 + 1);
    expect(worst).toBeLessThan(200);
    expect(grew).toBeLessThan(20);
  }, 300_000);

  test('요청 검증: 모양이 틀리거나 너무 크면 거른다', async () => {
    const { parseCfRemoteRequest } = await import('../src/xlsx/cfRangeValues.js');
    expect(parseCfRemoteRequest({ sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 5, c1: 0, crit: ['a', 1, null, true] }], stats: ['x'] })).not.toBeNull();
    expect(parseCfRemoteRequest({ sheet: 0, ranges: [{ r0: 5, c0: 0, r1: 1, c1: 0 }] })).toBeNull();
    expect(parseCfRemoteRequest({ sheet: -1, ranges: [] })).toBeNull();
    expect(parseCfRemoteRequest({ sheet: 0, ranges: Array.from({ length: 65 }, () => ({ r0: 0, c0: 0, r1: 1, c1: 0 })) })).toBeNull();
    expect(parseCfRemoteRequest({ sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 1, c1: 0, crit: Array.from({ length: 2_001 }, () => 'a') }] })).toBeNull();
    expect(parseCfRemoteRequest({ sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 1, c1: 0, crit: [{}] }] })).toBeNull();
    expect(parseCfRemoteRequest({ sheet: 0, ranges: [{ r0: '0', c0: 0, r1: 1, c1: 0 }] })).toBeNull();
  });
});

/** 한 업데이트로 여러 칸(붙여넣기) */
function pasteViaClient(socket: FakeSocket, entries: Array<[string, string]>): void {
  let captured: Uint8Array | null = null;
  const onUpdate = (update: Uint8Array) => {
    captured = update;
  };
  socket.mirror.on('update', onUpdate);
  socket.mirror.transact(() => {
    const cells = socket.mirror.getMap<string>('cells');
    for (const [k, v] of entries) cells.set(k, v);
  });
  socket.mirror.off('update', onUpdate);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SYNC);
  writeUpdate(encoder, captured!);
  socket.emit('message', encoding.toUint8Array(encoder));
}

describe('검토 r5 — 큰 붙여넣기(F2)·전역 상한·낡은 수식 알림', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-room-cfranges5-'));
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

  test('F2: 20만 칸 범위에 2만 칸 붙여넣기 — 다른 요청이 100ms 넘게 기다리지 않고, 답이 맞다', async () => {
    const { getCfRemoteAnswer } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');
    const { cfCritKey } = await import('../src/collab/condFormatShared.js');
    const { docId, socket } = await openRoom(await makeSheetXlsx(20_000, 10, [{ ref: 'A1:J20000', formula: 'COUNTIF($A:$J,A1)>1' }]));
    const range = { r0: 0, c0: 0, r1: 19_999, c1: 9 };
    await getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ ...range, crit: ['K1'] }] });
    syncMirrorWithRoom(socket);
    const gc = forceGc();
    gc();
    // B열 2만 칸에 수(500000…) 붙여넣기 — 정렬된 수 배열을 칸마다 고치면 범위 크기에 비례해 오래 걸렸다.
    // 붙여넣기 반영(웹소켓 업데이트 풀기·칸 쓰기)은 조건부 서식과 무관하게 원래 100ms 안팎 걸리므로 따로 재고(기록만),
    // 조건부 서식 몫(묶음 반영 + 다음 물음 때 표 다시 만들기)을 상한과 비교한다
    const applyMeter = stallMeter();
    pasteViaClient(socket, Array.from({ length: 20_000 }, (_, r): [string, string] => [gridCellKey(0, r, 1), String(500_000 + r)]));
    await flushEventLoop(5);
    const applyStall = applyMeter.stop();
    const meter = stallMeter();
    const a = await getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ ...range, crit: ['500000', 'K1'] }] });
    const stall = meter.stop();
    const r = a!.ranges[0] as { agg: { count: number; min: number; max: number }; crit: Array<[string, number]> };
    console.log(`F2: 붙여넣기 반영 중 가장 긴 멈춤 ${Math.round(applyStall)}ms(조건부 서식 없이도 비슷), 다음 물음(표 다시 만들기) 중 ${Math.round(stall)}ms`);
    expect(r.agg).toMatchObject({ count: 20_000, min: 500_000, max: 519_999 });
    expect(new Map(r.crit).get(cfCritKey('K1'))).toBe(2);
    expect(stall).toBeLessThan(100);
  }, 300_000);

  test('전역 상한: 열린 방들의 값 수가 넘으면 가장 오래 안 쓴 방의 표부터 버린다', async () => {
    const { newCfRoomTally, cfEnforceGlobalCap, cfTallyRelease } = await import('../src/xlsx/cfRangeValues.js');
    const fake = (n: number, lastUsed: number) => {
      const t = newCfRoomTally();
      t.lastUsed = lastUsed;
      const raw = new Map(Array.from({ length: n }, (_, i) => [`v${i}`, 1] as [string, number]));
      t.sheets.set(0, { entries: [], rows: n, cols: 1, rects: new Map([['k', { rect: { r0: 0, c0: 0, r1: n - 1, c1: 0 }, cells: n, raw, dirty: false, formulaCells: 0 }]]), stats: new Map() } as never);
      return t;
    };
    const old = fake(400, 1);
    const mid = fake(400, 2);
    const now = fake(400, 3);
    cfEnforceGlobalCap(now, 900);
    expect(old.sheets.size).toBe(0); // 가장 오래 안 쓴 방
    expect(mid.sheets.size).toBe(1);
    expect(now.sheets.size).toBe(1);
    [old, mid, now].forEach(cfTallyRelease);
  });

  test('범위에 이 세션에서 넣은 수식 칸이 있으면(보지 않는 페이지는 다시 계산되지 않는다) 답에 stale(편집 화면 알림)', async () => {
    const { getCfRemoteAnswer } = await import('../src/collab/room.js');
    const { gridCellKey } = await import('../src/collab/gridShared.js');
    // 파일의 수식 칸은 저장된 계산값으로 들어온다(수식 아님) — 이 세션에서 넣은 수식만 낡을 수 있다
    const { docId, socket } = await openRoom(await makeSheetXlsx(1_500, 1, [{ ref: 'A1:A1500', formula: 'COUNTIF($A:$A,$A1)>1' }]));
    const ask = () => getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 1_499, c1: 0, crit: ['K1'] }] });
    expect((await ask())!.stale).toBeUndefined();
    editCellViaClient(socket, gridCellKey(0, 1_200, 0), '=A1'); // 받지 않은 페이지(3번째)에 수식
    await flushEventLoop(5);
    expect((await ask())!.stale).toBe(true);
  }, 120_000);
});
