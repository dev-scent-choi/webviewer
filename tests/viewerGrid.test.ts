import { randomUUID } from 'node:crypto';
import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { parseXlsxToGrid } from '../src/spreadsheet.js';
import { closeLazyGridStore, ingestXlsxLazily, setDirtyCell } from '../src/xlsx/lazyGridStore.js';
import { serializeGridFromLazyStore, worksheetXmlBytes } from '../src/xlsx/viewerGrid.js';

async function sample(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('표');
  ws.addRow(['이름', '수량', '금액']).font = { bold: true };
  for (let i = 1; i <= 30; i++) ws.addRow([`품목${i}`, i, { formula: `B${i + 1}*100`, result: i * 100 }]);
  ws.getColumn(3).numFmt = '#,##0';
  ws.mergeCells('E1:F2');
  ws.getCell('E1').value = '병합';
  wb.addWorksheet('둘째').getCell('B3').value = 7;
  return Buffer.from(await wb.xlsx.writeBuffer());
}

describe('보기용 그리드 — 큰 파일용 스트리밍 경로(viewerGrid.ts)', () => {
  test('통째 읽기(parseXlsxToGrid)와 같은 헤더·칸 값', async () => {
    const bytes = await sample();
    const full = await parseXlsxToGrid(bytes);
    const store = await ingestXlsxLazily(randomUUID(), bytes);
    try {
      const { buffer, sheetCount } = await serializeGridFromLazyStore(store);
      const lines = buffer.toString('utf8').split('\n');
      const header = JSON.parse(lines[0]);
      expect(sheetCount).toBe(2);
      expect(header.sheets.map((s: { name: string; rowCount: number; colCount: number }) => [s.name, s.rowCount, s.colCount])).toEqual(
        full.sheets.map((s) => [s.name, s.rowCount, s.colCount]),
      );
      expect(header.sheets[0].merges).toEqual(full.sheets[0].merges);
      const values = (row: Array<{ v: string } | 0>) => row.map((c) => (c === 0 ? '' : c.v)).join('|').replace(/\|+$/, '');
      const rows = lines.slice(1, 1 + full.sheets[0].rowCount).map((l) => values(JSON.parse(l)));
      // 병합 범위의 가려진 칸은 통째 읽기만 값을 복사해 둔다(화면은 병합으로 그려 보이지 않는다) — 비교에서 뺀다
      const covered = (r: number, c: number) => full.sheets[0].merges.some((m) => r + 1 >= m.r && r + 1 < m.r + m.rs && c + 1 >= m.c && c + 1 < m.c + m.cs && !(r + 1 === m.r && c + 1 === m.c));
      expect(rows).toEqual(full.sheets[0].rows.map((row, r) => values((row as Array<{ v: string } | 0>).map((cell, c) => (covered(r, c) ? 0 : cell)))));
      expect(rows[1]).toBe('품목1|1|100');
    } finally {
      await closeLazyGridStore(store);
    }
  });

  test('편집 방의 지금 상태(고친 칸)가 들어간다', async () => {
    const store = await ingestXlsxLazily(randomUUID(), await sample());
    try {
      setDirtyCell(store, 0, 1, 0, { v: '고친 값' });
      const { buffer } = await serializeGridFromLazyStore(store);
      expect(JSON.parse(buffer.toString('utf8').split('\n')[2])[0].v).toBe('고친 값');
    } finally {
      await closeLazyGridStore(store);
    }
  });

  test('시트 XML 크기를 압축을 풀지 않고 잰다', async () => {
    const bytes = await sample();
    expect(worksheetXmlBytes(bytes)).toBeGreaterThan(1000);
    expect(worksheetXmlBytes(Buffer.from('not a zip'))).toBe(0);
  });
});
