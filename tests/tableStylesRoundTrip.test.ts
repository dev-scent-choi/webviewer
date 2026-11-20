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

describe('검토 r3 — 행/열 삽입·삭제 뒤 표 범위(화면 덧칠·저장 파일의 tableN.xml)', () => {
  type Op = { kind: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols'; at: number; count: number };
  // 표는 '둘째' 시트 A3:B5(머리 줄 + 2행, 요약 행 없음)
  const cases: Array<{ name: string; ops: Op[]; ref: string | null; cols?: number; header?: [string, string] }> = [
    { name: '표 위에 행 2개 삽입 → 통째로 내려간다', ops: [{ kind: 'insertRows', at: 0, count: 2 }], ref: 'A5:B7' },
    { name: '표 안에 행 삽입 → 늘어난다', ops: [{ kind: 'insertRows', at: 3, count: 1 }], ref: 'A3:B6' },
    { name: '표 아래에 행 삽입 → 그대로', ops: [{ kind: 'insertRows', at: 5, count: 1 }], ref: 'A3:B5' },
    { name: '표의 데이터 행 삭제 → 줄어든다', ops: [{ kind: 'deleteRows', at: 3, count: 1 }], ref: 'A3:B4' },
    { name: '표 안에 열 삽입 → 열이 늘고 머리 칸에 새 열 이름', ops: [{ kind: 'insertCols', at: 1, count: 1 }], ref: 'A3:C5', cols: 3, header: ['B3', '열1'] },
    { name: '표 왼쪽에 열 삽입 → 오른쪽으로 옮긴다', ops: [{ kind: 'insertCols', at: 0, count: 1 }], ref: 'B3:C5', cols: 2 },
    { name: '표의 첫 열 삭제 → 열 하나짜리 표', ops: [{ kind: 'deleteCols', at: 0, count: 1 }], ref: 'A3:A5', cols: 1 },
    { name: '표 행을 모두 삭제 → 표를 뺀다', ops: [{ kind: 'deleteRows', at: 1, count: 5 }], ref: null },
    { name: '여러 연산을 차례로', ops: [{ kind: 'insertRows', at: 0, count: 1 }, { kind: 'insertCols', at: 1, count: 1 }, { kind: 'deleteRows', at: 4, count: 1 }], ref: 'A4:C5', cols: 3 },
  ];
  for (const c of cases) {
    test(c.name, async () => {
      const { applyColShift, applyRowShift } = await import('../src/xlsx/lazyGridStore.js');
      const { shiftTables } = await import('../src/xlsx/tableStyles.js');
      const { strFromU8, unzipSync } = await import('fflate');
      const original = fs.readFileSync('audit/office/corpus/xlsx/a3-fidelity-grid.xlsx');
      const store = await ingestXlsxLazily(randomUUID(), original);
      try {
        const s = store.sheetMeta.findIndex((m) => m.name === '둘째');
        const extents = store.sheetMeta.map((m) => ({ rowCount: m.rowCount, colCount: m.colCount }));
        const ops = c.ops.map((op) => ({ sheet: s, ...op }));
        for (const op of ops) (op.kind.endsWith('Rows') ? applyRowShift : applyColShift)(store, s, op.at, op.count, op.kind.startsWith('insert'));
        // 화면 덧칠용(xlsx-extras)과 저장 파일이 같은 범위를 말해야 한다
        const shifted = shiftTables(readSheetExtras(original).tables, ops, (i) => store.sheetMeta[i]?.name);
        const saved = await exportXlsxFromGrid(original, store, extents, ops);
        const tableXml = Object.entries(unzipSync(saved)).filter(([n]) => /^xl\/tables\/table\d+\.xml$/.test(n)).map(([, b]) => strFromU8(b));
        const reread = new ExcelJS.Workbook();
        await reread.xlsx.load(saved as unknown as ArrayBuffer);
        const ws = reread.getWorksheet('둘째')!;
        if (c.ref === null) {
          expect(shifted).toEqual([]);
          expect(tableXml).toEqual([]);
          expect(ws.getTables()).toHaveLength(0);
          return;
        }
        const toA1 = (t: { r0: number; r1: number; c0: number; c1: number }) => `${String.fromCharCode(65 + t.c0)}${t.r0 + 1}:${String.fromCharCode(65 + t.c1)}${t.r1 + 1}`;
        expect(toA1(shifted[0])).toBe(c.ref);
        expect(tableXml).toHaveLength(1);
        expect(tableXml[0]).toMatch(new RegExp(`<table [^>]*ref="${c.ref}"`));
        expect(tableXml[0]).toMatch(new RegExp(`<autoFilter ref="${c.ref}"`));
        if (c.cols) expect(tableXml[0]).toContain(`<tableColumns count="${c.cols}"`);
        // 머리 줄·필터 단추는 원본대로(ExcelJS는 headerRowCount가 없는 표를 머리 줄 없음·단추 숨김으로 저장했다)
        expect(tableXml[0]).toContain('headerRowCount="1"');
        expect(tableXml[0]).not.toContain('hiddenButton="1"');
        // ExcelJS로 다시 열어도 같은 범위이고, 다시 읽은 표 정보(편집 화면이 받는 것)도 같다
        const t = ws.getTables()[0] as unknown as { table: { tableRef: string } };
        expect(t.table.tableRef).toBe(c.ref);
        expect(toA1(readSheetExtras(saved).tables[0])).toBe(c.ref);
        if (c.header) expect(ws.getCell(c.header[0]).value).toBe(c.header[1]);
      } finally {
        await closeLazyGridStore(store);
      }
    });
  }
});
