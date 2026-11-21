import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import * as encoding from 'lib0/encoding';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/**
 * FINAL-REVIEW 재검토 2 — C1 예산의 남은 구멍 셋.
 * 1) 값싼 조건(`*` 0~1개, `~`)이 칸 길이를 내지 않아 2,000 × 32k자 × 200칸 요청이 11초.
 * 2) 편집 방 수식 재계산이 칸마다 따로 예산이라 무거운 열 하나가 쪽 하나에 1분 가까이 막았다.
 * 3) 업로드 예산이 빠듯해 정상 파일(1,000행 COUNTIF "*"&B1&"*")의 213칸이 '#ERROR!'.
 */
class FakeSocket extends EventEmitter {
  binaryType = '';
  readyState = 1;
  readonly CONNECTING = 0;
  readonly OPEN = 1;
  send(_d: Uint8Array, cb?: (err?: Error) => void): void {
    cb?.();
  }
  close(): void {}
  ping(): void {}
}
const LONG = (i: number) => `${'a'.repeat(32_000)}${i}`;

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

describe('재검토 2: C1 예산', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-c1-budget2-'));
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

  async function openXlsx(build: (ws: ExcelJS.Worksheet) => void): Promise<{ docId: string; socket: FakeSocket }> {
    const { store } = await import('../src/storage.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    const wb = new ExcelJS.Workbook();
    build(wb.addWorksheet('S'));
    const docId = randomUUID();
    await store.saveOriginal(docId, 'xlsx', Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer));
    await store.writeMeta({ id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid' });
    const socket = new FakeSocket();
    await joinEditSession(socket as never, { url: '/ws' } as never, docId, 'xlsx', [], false, 'v');
    return { docId, socket };
  }
  const subscribe = (socket: FakeSocket) => {
    const e = encoding.createEncoder();
    encoding.writeVarUint(e, 2);
    encoding.writeVarString(e, JSON.stringify({ sheet: 0, pages: [0] }));
    socket.emit('message', encoding.toUint8Array(e));
  };

  test('1) 값싼 조건 2,000개(q123*z·~) × 32k자 칸 200개 요청이 1초 안', async () => {
    const { getCfRemoteAnswer } = await import('../src/collab/room.js');
    const { docId, socket } = await openXlsx((ws) => {
      for (let r = 1; r <= 200; r += 1) ws.getCell(r, 1).value = LONG(r);
      ws.addConditionalFormatting({ ref: 'A1:A200', rules: [{ type: 'expression', priority: 1, formulae: ['COUNTIF($A$1:$A$200,"q*")>0'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
    });
    subscribe(socket);
    const range = { r0: 0, c0: 0, r1: 199, c1: 0 };
    await getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ ...range, crit: ['x'] }] }); // 표를 먼저 센다
    const crit = Array.from({ length: 2_000 }, (_, i) => (i % 2 ? `q${i}*z` : `~q${i}`));
    const t0 = performance.now();
    const answer = await getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ ...range, crit }] });
    expect(performance.now() - t0).toBeLessThan(1_000);
    expect(answer).not.toBeNull();
  }, 120_000);

  test('2) 편집 방 재계산: 캐시 없는 무거운 COUNTIF 열 500행 — 이벤트 루프 최대 지연 200ms 미만, 넘은 칸은 수식 원문', async () => {
    const { getRoomYdocForTest, drainRoomBackgroundWorkForTest } = await import('../src/collab/room.js');
    const { docId, socket } = await openXlsx((ws) => {
      for (let r = 1; r <= 500; r += 1) {
        if (r <= 200) ws.getCell(r, 1).value = LONG(r);
        ws.getCell(r, 2).value = `${'a'.repeat(290)}b${r}`;
        ws.getCell(r, 3).value = { formula: `COUNTIF($A$1:$A$200,"*"&B${r}&"*")` } as never;
      }
    });
    const meter = stallMeter();
    const t0 = performance.now();
    subscribe(socket);
    const computed = getRoomYdocForTest(docId)!.getMap<string>('computed');
    while (performance.now() - t0 < 60_000) {
      await drainRoomBackgroundWorkForTest();
      if (Array.from({ length: 500 }, (_, r) => computed.get(`0:${r}:2`)).every((v) => v !== undefined)) break;
      await new Promise((r) => setTimeout(r, 50));
    }
    const ms = performance.now() - t0;
    const stall = meter.stop();
    console.log(`재계산 500칸 ${Math.round(ms)}ms, 최대 지연 ${Math.round(stall)}ms`);
    expect(stall).toBeLessThan(200);
    expect(ms).toBeLessThan(30_000);
    const values = Array.from({ length: 500 }, (_, r) => computed.get(`0:${r}:2`));
    expect(values.every((v) => v !== undefined)).toBe(true);
    expect(values).not.toContain('#ERROR!');
    expect(values.some((v) => String(v).startsWith('=COUNTIF'))).toBe(true); // 예산을 넘은 칸은 수식 원문
  }, 120_000);

  test('3) 정상 파일: 1,000행 COUNTIF($A$1:$A$1000,"*"&B1&"*") — 1,000칸 모두 숫자', async () => {
    const { parseXlsxToGrid, gridCellValue } = await import('../src/spreadsheet.js');
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    for (let r = 1; r <= 1_000; r += 1) {
      ws.getCell(r, 1).value = `제품코드-${String(r).padStart(6, '0')}-가나다`;
      ws.getCell(r, 2).value = String(r).padStart(6, '0'); // A의 r번째 값에만 든다 → 모두 1
      ws.getCell(r, 3).value = { formula: `COUNTIF($A$1:$A$1000,"*"&B${r}&"*")` } as never;
    }
    const grid = await parseXlsxToGrid(Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer));
    const cells = grid.sheets[0].rows.map((row) => row[2]).slice(0, 1_000).map((c) => (c ? gridCellValue(c) : ''));
    expect(cells.filter((v) => /^\d+$/.test(v))).toHaveLength(1_000);
  }, 120_000);
});
