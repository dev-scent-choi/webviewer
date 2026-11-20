import { randomUUID } from 'node:crypto';
import fs from 'node:fs';
import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { exportXlsxFromGrid } from '../src/xlsx/collab.js';
import { closeLazyGridStore, getRowCells, ingestXlsxLazily, setDirtyCell } from '../src/xlsx/lazyGridStore.js';
import { readSheetExtras } from '../src/xlsx/sheetExtras.js';

describe('검토 N3 — 표 스타일은 그릴 때만 덧칠, 저장되는 칸 서식에 들어가지 않는다', () => {
  test('표가 있는 파일: 값 하나 고치고 정렬(값 섞기)한 뒤 저장해도 표 칸의 직접 서식(채우기·글자 색·굵게)이 원본과 같다', async () => {
    const original = fs.readFileSync('audit/office/corpus/xlsx/a3-fidelity-grid.xlsx');
    const before = new ExcelJS.Workbook();
    await before.xlsx.load(original as unknown as ArrayBuffer);
    const sheetName = '둘째';
    const formats = (wb: ExcelJS.Workbook) => {
      const ws = wb.getWorksheet(sheetName)!;
      return ['A3', 'B3', 'A4', 'B4', 'A5', 'B5'].map((a) => {
        const c = ws.getCell(a);
        // 채우기 없음과 'none' 무늬는 같은 뜻(저장할 때 손댄 칸은 'none'으로 쓴다)
        return JSON.stringify({ fill: c.fill?.type === 'pattern' && (c.fill as ExcelJS.FillPattern).pattern !== 'none' ? (c.fill as ExcelJS.FillPattern).pattern : null, color: c.font?.color ?? null, bold: c.font?.bold ?? false });
      });
    };
    const store = await ingestXlsxLazily(randomUUID(), original);
    try {
      const s = store.sheetMeta.findIndex((m) => m.name === sheetName);
      // 편집 경로(지연 그리드)의 칸 서식에는 표 모양이 없다
      const head = (await getRowCells(store, s, 2, 2))[0];
      expect(head !== 0 && head.s !== undefined ? store.styles[head.s].bg : undefined).toBeUndefined();
      // 값 하나 고치기 + 정렬처럼 두 행의 값 바꾸기
      setDirtyCell(store, s, 3, 1, { v: '10' });
      setDirtyCell(store, s, 3, 0, { v: 'b' });
      setDirtyCell(store, s, 4, 0, { v: 'a' });
      setDirtyCell(store, s, 4, 1, { v: '1' });
      const saved = await exportXlsxFromGrid(original, store, store.sheetMeta.map((m) => ({ rowCount: m.rowCount, colCount: m.colCount })), []);
      const after = new ExcelJS.Workbook();
      await after.xlsx.load(saved as unknown as ArrayBuffer);
      expect(formats(after)).toEqual(formats(before));
      expect(after.getWorksheet(sheetName)!.getCell('B4').value).toBe(10);
    } finally {
      await closeLazyGridStore(store);
    }
  });

  test('편집 화면이 받는 표 정보(xlsx-extras): 범위·스타일·테마 색', () => {
    const extras = readSheetExtras(fs.readFileSync('audit/office/corpus/xlsx/a3-fidelity-grid.xlsx'));
    expect(extras.tables).toEqual([expect.objectContaining({ sheetName: '둘째', r0: 2, r1: 4, c0: 0, c1: 1, style: 'TableStyleMedium2' })]);
    expect(extras.themeColors[4]).toBe('4F81BD');
  });
});
