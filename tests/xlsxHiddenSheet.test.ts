import { randomUUID } from 'node:crypto';
import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { exportXlsxFromGrid } from '../src/xlsx/collab.js';
import { ingestXlsxLazily, setDirtyCell } from '../src/xlsx/lazyGridStore.js';

async function workbook(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const a = wb.addWorksheet('앞');
  a.getCell('A1').value = '앞1';
  a.getCell('C1').value = '숨긴열';
  a.getCell('A2').value = '숨긴행';
  a.getColumn(3).hidden = true;
  a.getRow(2).hidden = true;
  const h = wb.addWorksheet('숨김');
  h.state = 'hidden';
  h.getCell('A1').value = '숨긴 시트 값';
  wb.addWorksheet('뒤').getCell('A1').value = '뒤1';
  return Buffer.from(await wb.xlsx.writeBuffer());
}

describe('숨긴 시트가 있는 엑셀 — 편집 방 시트 번호와 저장', () => {
  test('편집 방에는 보이는 시트만 탭 순서대로(저장·보기 화면과 같은 번호)', async () => {
    const store = await ingestXlsxLazily(randomUUID(), await workbook());
    expect(store.sheetMeta.map((m) => m.name)).toEqual(['앞', '뒤']);
  });

  test('저장해도 숨긴 시트·숨긴 행·열이 그대로, 편집은 맞는 시트에', async () => {
    const original = await workbook();
    const store = await ingestXlsxLazily(randomUUID(), original);
    setDirtyCell(store, 1, 0, 0, { v: '뒤에 씀' });
    const extents = store.sheetMeta.map((m) => ({ rowCount: m.rowCount, colCount: m.colCount }));
    const out = await exportXlsxFromGrid(new Uint8Array(original), store, extents, []);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out as unknown as ArrayBuffer);
    expect(wb.worksheets.map((s) => `${s.name}:${s.state}`)).toEqual(['앞:visible', '숨김:hidden', '뒤:visible']);
    expect(wb.getWorksheet('앞')!.getColumn(3).hidden).toBe(true);
    expect(wb.getWorksheet('앞')!.getRow(2).hidden).toBe(true);
    expect(wb.getWorksheet('뒤')!.getCell('A1').value).toBe('뒤에 씀');
    expect(wb.getWorksheet('숨김')!.getCell('A1').value).toBe('숨긴 시트 값');
  });
});
