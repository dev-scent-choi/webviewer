import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import type { FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import { CF_TOKEN_WINDOW_MS, CF_TOKEN_WORK_PER_WINDOW, CfBusyError, cfAnswerSlot, cfTokenBudget, cfTokenSpend, newCfRoomTally } from '../src/xlsx/cfRangeValues.js';

/**
 * 재검토 2 판정 — cf-ranges 요청 하나가 최악 약 2초라 분당 240회면 한 토큰이 서버를 계속 잡을 수 있었다.
 * (1) 문서마다 답 만들기는 동시에 하나(하나만 뒤에서 기다리고 더 오면 409), (2) 토큰마다 1분 와일드카드 작업량 5e8,
 * (3) 요청 수는 토큰마다 분당 40회(편집 화면 정상 사용 최대 18회의 2배 남짓).
 */
describe('cf-ranges 남용 막기', () => {
  test('(1) 문서마다 동시에 하나 — 하나는 기다렸다 이어서, 그다음은 CfBusyError', async () => {
    const t = newCfRoomTally();
    const order: string[] = [];
    let release!: () => void;
    const a = cfAnswerSlot(t, async () => {
      order.push('a start');
      await new Promise<void>((r) => (release = r));
      order.push('a end');
      return 'a';
    });
    const b = cfAnswerSlot(t, async () => {
      order.push('b start');
      return 'b';
    });
    await expect(cfAnswerSlot(t, async () => 'c')).rejects.toBeInstanceOf(CfBusyError);
    release();
    expect(await Promise.all([a, b])).toEqual(['a', 'b']);
    expect(order).toEqual(['a start', 'a end', 'b start']); // 겹치지 않았다
    expect(await cfAnswerSlot(t, async () => 'd')).toBe('d'); // 끝나면 다시 받는다
  });

  test('(2) 토큰마다 1분 작업량 — 다 쓰면 다음 요청 예산 0, 창이 지나면 다시', () => {
    const now = Date.now();
    const tok = `t-${randomUUID()}`;
    const first = cfTokenBudget(tok, now);
    expect(first.left).toBe(CF_TOKEN_WORK_PER_WINDOW);
    const start = first.left;
    first.left = 0; // 이번 요청이 다 썼다
    cfTokenSpend(tok, first, start);
    expect(cfTokenBudget(tok, now + 1_000).left).toBe(0);
    expect(cfTokenBudget(`t-${randomUUID()}`, now + 1_000).left).toBe(CF_TOKEN_WORK_PER_WINDOW); // 다른 토큰은 그대로
    expect(cfTokenBudget(tok, now + CF_TOKEN_WINDOW_MS).left).toBe(CF_TOKEN_WORK_PER_WINDOW);
  });
});

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

describe('cf-ranges 경로: 토큰 예산·요청 수 제한', () => {
  let app: FastifyInstance;
  let tmpDir: string;
  let docId: string;
  let issueViewToken: (typeof import('../src/security.js'))['issueViewToken'];
  beforeAll(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-cf-abuse-'));
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    const { store, newDocumentId } = await import('../src/storage.js');
    ({ issueViewToken } = await import('../src/security.js'));
    const { buildApp } = await import('../src/app.js');
    const { joinEditSession } = await import('../src/collab/room.js');
    app = await buildApp();
    await app.ready();
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    for (let r = 1; r <= 300; r += 1) ws.getCell(r, 1).value = `K${r}`;
    ws.addConditionalFormatting({ ref: 'A1:A300', rules: [{ type: 'expression', priority: 1, formulae: ['COUNTIF($A$1:$A$300,"K1*")>0'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
    docId = newDocumentId();
    await store.saveOriginal(docId, 'xlsx', Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer));
    await store.writeMeta({ id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid' });
    await joinEditSession(new FakeSocket() as never, { url: '/ws' } as never, docId, 'xlsx', [], false, 'v');
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });
  const editToken = (viewer: string) => issueViewToken({ docId, viewer, perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: true } }, 600);
  const ask = (token: string, crit: string[] = ['K1*']) =>
    app.inject({ method: 'POST', url: `/api/view/xlsx-cf-ranges?token=${encodeURIComponent(token)}`, payload: { sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 299, c1: 0, crit }] } });

  test('(3) 정상 사용의 2배(36회/분)는 막히지 않고, 41번째는 429 — 다른 토큰은 따로 센다', async () => {
    const token = editToken('a');
    for (let i = 0; i < 36; i += 1) {
      const res = await ask(token);
      expect(res.statusCode).toBe(200);
      expect(res.json().data.ranges[0]).toMatchObject({ cells: 300, crit: [[expect.any(String), 111]] }); // K1·K10~K19·K100~K199
    }
    for (let i = 36; i < 40; i += 1) expect((await ask(token)).statusCode).toBe(200);
    expect((await ask(token)).statusCode).toBe(429);
    expect((await ask(editToken('b'))).statusCode).toBe(200);
  }, 60_000);

  test('(2) 토큰의 1분 작업량을 다 쓰면 다음 요청은 바로 overBudget', async () => {
    const token = editToken('c');
    const start = cfTokenBudget(token).left;
    cfTokenSpend(token, { left: 0 }, start); // 이 토큰이 1분 예산을 이미 다 썼다
    const t0 = performance.now();
    const res = await ask(token);
    expect(performance.now() - t0).toBeLessThan(200);
    expect(res.statusCode).toBe(200);
    expect(res.json().data).toMatchObject({ overBudget: true, ranges: [{ skipped: true }] });
  }, 60_000);
});
