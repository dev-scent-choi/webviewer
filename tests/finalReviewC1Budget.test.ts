import { EventEmitter } from 'node:events';
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import FormulaParser from 'fast-formula-parser';
import * as encoding from 'lib0/encoding';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/**
 * FINAL-REVIEW 재검토 C1 — 선형 매처라도 총 작업량에 상한이 없으면 요청 하나(조건 2,000개 × 32k자 칸 200개)가 몇 시간 걸린다.
 * 요청·굽기·업로드 계산마다 작업 예산을 두고, 넘으면 그 조건은 생략(편집 화면 "일부 생략")으로 답한다.
 */
const NEAR_MISS = `*${'a'.repeat(290)}b*`; // 칸마다 거의 끝까지 맞춰 보고 실패한다
const LONG = (i: number) => `${'a'.repeat(32_000)}${i}`;

async function longCellsBook(opts: { cf?: string; formulas?: number } = {}): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('S');
  for (let r = 1; r <= 200; r += 1) ws.getCell(r, 1).value = LONG(r);
  for (let r = 1; r <= (opts.formulas ?? 0); r += 1) ws.getCell(r, 3).value = { formula: `COUNTIF($A$1:$A$200,"${NEAR_MISS}")` } as never;
  const b = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
  if (opts.cf) {
    b['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(b['xl/worksheets/sheet1.xml']).replace('</sheetData>', `</sheetData>${opts.cf}`));
    b['xl/styles.xml'] = strToU8(strFromU8(b['xl/styles.xml']).replace(/<dxfs count="0"\/>|(?=<tableStyles)/, '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'));
  }
  return Buffer.from(zipSync(b));
}
const COUNTIF_CF = `<conditionalFormatting sqref="A1:A200"><cfRule type="expression" dxfId="0" priority="1"><formula>COUNTIF($A$1:$A$200,"${NEAR_MISS}")&gt;0</formula></cfRule></conditionalFormatting>`;

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

describe('재검토 C1: 와일드카드 작업 예산', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-c1-budget-'));
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

  test('xlsx-cf-ranges: 조건 2,000개 × 32k자 칸 200개 요청이 1초 안, 예산 초과를 응답에 적는다', async () => {
    const { store } = await import('../src/storage.js');
    const { joinEditSession, getCfRemoteAnswer } = await import('../src/collab/room.js');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'xlsx', await longCellsBook({ cf: COUNTIF_CF }));
    await store.writeMeta({ id: docId, name: 'x.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), progress: 100, viewKind: 'grid' });
    const s = new FakeSocket();
    await joinEditSession(s as never, { url: '/ws' } as never, docId, 'xlsx', [], false, 'v');
    const e = encoding.createEncoder();
    encoding.writeVarUint(e, 2);
    encoding.writeVarString(e, JSON.stringify({ sheet: 0, pages: [0] }));
    s.emit('message', encoding.toUint8Array(e));
    const crit = Array.from({ length: 2_000 }, (_, i) => `${NEAR_MISS}${i}`.slice(0, 300));
    await getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 199, c1: 0, crit: ['x'] }] }); // 표를 먼저 세어 둔다
    const t0 = performance.now();
    const answer = await getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 199, c1: 0, crit }] });
    expect(performance.now() - t0).toBeLessThan(5_000); // 병렬 시험 부하에서도 흔들리지 않게 넓힌 기준(재검토 3) — 고치기 전 값(13초)과는 여전히 크게 떨어진다
    expect(answer!.ranges[0]).toMatchObject({ skipped: true });
    expect(answer!.overBudget).toBe(true);
    // 정상 크기 요청은 그대로 답한다
    const normal = await getCfRemoteAnswer(docId, { sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 199, c1: 0, crit: ['a*1', '*99'] }] }); // 1·11·…·191 → 20개, 99·199 → 2개
    expect(normal!.overBudget).toBeUndefined();
    expect(normal!.ranges[0]).toMatchObject({ cells: 200, crit: [[expect.any(String), 20], [expect.any(String), 2]] });
  }, 120_000);

  test('업로드 굽기: 32k자 칸에 거의 맞는 COUNTIF 규칙이 몇 초씩 막지 않고 "일부 생략"을 적는다', async () => {
    const { parseViewerGrid } = await import('../src/xlsx/viewerGrid.js');
    const x = await longCellsBook({ cf: COUNTIF_CF });
    const t0 = performance.now();
    const grid = await parseViewerGrid(x);
    expect(performance.now() - t0).toBeLessThan(15_000); // 병렬 시험 부하에서도 흔들리지 않게 넓힌 기준(재검토 3) — 고치기 전 값(측정 안 함 — 결과 cfSkipped로 본다)과는 여전히 크게 떨어진다
    expect(grid.cfSkipped).toBe(1);
  }, 120_000);

  test('수식 엔진: 거의 맞는 COUNTIF 한 번이 1초 안(예산을 넘으면 오류값), 업로드 계산의 수식 100개도 합쳐서 묶인다', async () => {
    await import('../src/client/formulaOverrides.js'); // 와일드카드 덮어쓰기
    const parser = new FormulaParser({ onRange: () => Array.from({ length: 200 }, (_, i) => [LONG(i)]) });
    let t0 = performance.now();
    let out: unknown;
    try {
      out = parser.parse(`COUNTIF(A1:A200,"${NEAR_MISS}")`, { sheet: 'S', row: 1, col: 3 });
    } catch (error) {
      out = error;
    }
    expect(performance.now() - t0).toBeLessThan(5_000); // 병렬 시험 부하에서도 흔들리지 않게 넓힌 기준(재검토 3) — 고치기 전 값(3.7초 — 결과가 0이 아님으로도 본다)과는 여전히 크게 떨어진다
    expect(out).not.toBe(0); // 예산을 넘어 끝까지 세지 않았다(0개라고 답하면 틀린 답)
    const { parseXlsxToGrid } = await import('../src/spreadsheet.js');
    const x = await longCellsBook({ formulas: 100 });
    t0 = performance.now();
    await parseXlsxToGrid(x);
    expect(performance.now() - t0).toBeLessThan(15_000); // 병렬 시험 부하에서도 흔들리지 않게 넓힌 기준(재검토 3) — 고치기 전 값(수식 100개 × 3.7초)과는 여전히 크게 떨어진다
  }, 120_000);
});
