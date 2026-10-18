import { describe, expect, test } from 'vitest';
import ExcelJS from 'exceljs';
import { parseXlsxToGrid } from '../src/spreadsheet.js';
import { ingestXlsxLazily } from '../src/xlsx/lazyGridStore.js';
import fs from 'node:fs/promises';
import {
  RowLayout,
  defaultColWidthPx,
  defaultFontOfStylesXml,
  maxDigitWidthPx,
  encodeRowHeights,
  excelColWidthToPx,
  shiftRowHeightRuns,
} from '../src/xlsx/sheetGeometry.js';

describe('엑셀 열 너비·행 높이 → 화면 px', () => {
  test('열 너비는 OOXML 변환식(숫자 폭 7px), 기본 열은 64px', () => {
    expect([5, 9.140625, 22, 40].map((w) => excelColWidthToPx(w))).toEqual([35, 64, 154, 280]);
    expect(defaultColWidthPx()).toBe(64);
    expect(defaultColWidthPx(18)).toBe(126);
  });

  test('열 너비 단위는 통합 문서 기본 글꼴의 숫자 폭(Calibri 7·맑은 고딕 8·돋움 9px)', () => {
    expect(maxDigitWidthPx('Calibri', 11)).toBe(7);
    expect(maxDigitWidthPx('맑은 고딕', 11)).toBe(8);
    expect(maxDigitWidthPx('돋움', 11)).toBe(9);
    // 돋움 기본 문서: 너비 3.66 → 33px(엑셀 화면 실측과 같다), 기본 8.89 → 80px
    expect(excelColWidthToPx(3.6640625, 9)).toBe(33);
    expect(defaultColWidthPx(8.88671875, 8, 9)).toBe(80);
    // 한국어 엑셀 기본 열(맑은 고딕 11, 8.38자) = 72px
    expect(excelColWidthToPx(8.38 + 5 / 8, 8)).toBe(72);
    expect(defaultFontOfStylesXml('<styleSheet><fonts count="2"><font><sz val="11"/><name val="돋움"/></font><font/></fonts></styleSheet>')).toEqual({ name: '돋움', size: 11 });
  });

  test('RowLayout: 높이가 다른 행·숨긴 행(0)·필터로 숨긴 행을 반영한 위치', () => {
    const runs = encodeRowHeights([[0, 48], [3, 40], [4, 40], [8, 0]], 20);
    expect(runs).toEqual([[0, 1, 48], [3, 2, 40], [8, 1, 0]]);
    const layout = new RowLayout(20, runs, [6]);
    expect([0, 1, 3, 5, 6, 7, 8, 9, 10].map((r) => layout.top(r))).toEqual([0, 48, 88, 168, 188, 188, 208, 208, 228]);
    expect(layout.height(6)).toBe(0);
    expect(layout.rowAt(0)).toBe(0);
    expect(layout.rowAt(47)).toBe(0);
    expect(layout.rowAt(48)).toBe(1);
    expect(layout.rowAt(100)).toBe(3);
    // 숨긴 행(6·8)은 건너뛴다
    expect(layout.rowAt(188)).toBe(7);
    expect(layout.rowAt(208)).toBe(9);
    expect(layout.rowAt(1000)).toBe(9 + Math.floor((1000 - 208) / 20));
  });

  test('행 삽입·삭제에 맞춰 높이 묶음을 옮긴다', () => {
    const runs = [[2, 3, 40]] as const; // 2·3·4행
    expect(shiftRowHeightRuns(runs, 3, 2, true)).toEqual([[2, 1, 40], [5, 2, 40]]);
    expect(shiftRowHeightRuns(runs, 0, 1, true)).toEqual([[3, 3, 40]]);
    expect(shiftRowHeightRuns(runs, 3, 1, false)).toEqual([[2, 2, 40]]);
    expect(shiftRowHeightRuns(runs, 0, 2, false)).toEqual([[0, 3, 40]]);
    expect(shiftRowHeightRuns(runs, 1, 5, false)).toEqual([]);
  });
});

async function sampleWorkbook(): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('점검표');
  ws.getColumn(1).width = 5;
  ws.getColumn(2).width = 40;
  ws.getColumn(3).hidden = true;
  ws.getCell('A1').value = '제목';
  ws.getRow(1).height = 36;
  ws.getCell('B2').value = '필터 차압 확인 및 벨트 장력 점검, 이상 소음 여부 확인과 함께 설비 상태를 기록한다';
  ws.getCell('B2').alignment = { wrapText: true };
  ws.getCell('A3').value = '숨김';
  ws.getRow(3).hidden = true;
  ws.getCell('A4').value = '끝';
  ws.getCell('D4').value = '숨긴 열 뒤';
  const ws2 = wb.addWorksheet('둘째');
  ws2.properties.defaultColWidth = 18;
  ws2.getCell('A1').value = 'x';
  return Buffer.from(await wb.xlsx.writeBuffer());
}

describe('보기·편집 양쪽 추출이 같은 열 너비·행 높이를 낸다', () => {
  test('parseXlsxToGrid(보기)와 ingestXlsxLazily(편집)', async () => {
    const bytes = await sampleWorkbook();
    const grid = await parseXlsxToGrid(bytes);
    const lazy = await ingestXlsxLazily('geom-test', bytes);
    try {
      for (const sheet of [grid.sheets[0], lazy.sheetMeta[0]]) {
        expect(sheet.colWidths.slice(0, 3)).toEqual([35, 280, 0]);
        expect(sheet.defaultRowHeight).toBe(20);
        const layout = new RowLayout(sheet.defaultRowHeight, sheet.rowHeights);
        expect(layout.height(0)).toBe(48);
        expect(layout.height(1)).toBe(60); // 줄 바꿈 칸 자동 맞춤(280px 폭에 세 줄)
        expect(layout.height(2)).toBe(0); // 숨긴 행
        expect(layout.height(3)).toBe(20);
      }
      // 시트가 여럿이어도 편집 쪽에서 빠지지 않는다(스트리밍 리더가 둘째 시트를 건너뛰던 문제)
      expect(lazy.sheetMeta.map((m) => m.name)).toEqual(['점검표', '둘째']);
      expect(grid.sheets[1].defaultColWidth).toBe(126);
      expect(lazy.sheetMeta[1].defaultColWidth).toBe(126);
    } finally {
      await fs.rm(lazy.rowStorePath, { force: true });
    }
  });
});

describe('그림·격자선', () => {
  test('그림 위치·내용을 읽고, 격자선을 끈 시트를 알려 준다', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('표지', { views: [{ showGridLines: false }] });
    ws.getCell('C3').value = '제목';
    // 1x1 png
    const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64');
    const imageId = wb.addImage({ buffer: png as unknown as ExcelJS.Buffer, extension: 'png' });
    ws.addImage(imageId, { tl: { col: 1, row: 1 }, ext: { width: 120, height: 40 } });
    const bytes = Buffer.from(await wb.xlsx.writeBuffer());
    const { readSheetImages } = await import('../src/xlsx/drawingPassthrough.js');
    const images = readSheetImages(new Uint8Array(bytes));
    expect(images).toHaveLength(1);
    expect(images[0].sheetName).toBe('표지');
    expect(images[0].from).toMatchObject({ col: 1, row: 1 });
    expect(images[0].src.startsWith('data:image/png;base64,')).toBe(true);
    const grid = await parseXlsxToGrid(bytes);
    expect(grid.sheets[0].showGridLines).toBe(false);
    const lazy = await ingestXlsxLazily('geom-grid', bytes);
    try {
      expect(lazy.sheetMeta[0].showGridLines).toBe(false);
    } finally {
      await fs.rm(lazy.rowStorePath, { force: true });
    }
  });
});
