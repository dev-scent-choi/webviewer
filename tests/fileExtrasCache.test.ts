import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { afterEach, beforeEach, describe, expect, test, vi } from 'vitest';

/** 재검토 4 — 파일 링크·목록·표 해석(getFileExtras)은 보기 화면을 열 때마다 원본을 다시 풀지 않는다. 저장(무효화) 뒤엔 다시 읽는다 */
describe('getFileExtras 캐시', () => {
  let tmpDir: string;
  beforeEach(async () => {
    tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-extras-'));
    vi.resetModules();
    vi.stubEnv('DOCVIEW_DATA_DIR', tmpDir);
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
  });
  afterEach(async () => {
    vi.useRealTimers();
    vi.restoreAllMocks();
    vi.unstubAllEnvs();
    vi.resetModules();
    await fs.rm(tmpDir, { recursive: true, force: true });
  });

  test('적중하면 원본을 다시 읽지 않고, invalidateFileExtras 뒤엔 다시 읽는다', async () => {
    const { getFileExtras, invalidateFileExtras } = await import('../src/collab/room.js');
    const { store } = await import('../src/storage.js');
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('요약').getCell('A1').value = { text: '자료로', hyperlink: '#자료!A1' };
    wb.addWorksheet('자료');
    const docId = randomUUID();
    await store.saveOriginal(docId, 'xlsx', Buffer.from(await wb.xlsx.writeBuffer()));
    const read = vi.spyOn(store, 'readOriginal');

    vi.useFakeTimers();
    const first = await getFileExtras(docId);
    expect(first.links.map((l) => l.location)).toEqual(['자료!A1']);
    // 원본 뼈대 캐시(30초)가 지나도 해석 결과 캐시로 답한다
    vi.advanceTimersByTime(60_000);
    expect(await getFileExtras(docId)).toBe(first);
    expect(read).toHaveBeenCalledTimes(1);

    invalidateFileExtras(docId);
    const again = await getFileExtras(docId);
    expect(again).not.toBe(first);
    expect(again.links).toEqual(first.links);
    expect(read).toHaveBeenCalledTimes(2);
  });
});
