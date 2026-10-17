import fs from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { afterEach, describe, expect, test } from 'vitest';
import { computeGridCellDisplayAsync } from '../src/xlsx/formulas.js';
import { ingestXlsxLazily, type LazyGridStore } from '../src/xlsx/lazyGridStore.js';

/**
 * 회귀 테스트 — 동시편집 룸의 수식 평가(computeGridCellDisplaySync/Async)가 행 저장소의
 * "표시용" 텍스트(쉼표 천단위 등)를 정규화 없이 수식 엔진에 먹이던 버그.
 *
 * 증상(E2E로 실제 재현): #,##0 서식이 걸린 숫자 칸을 참조하는 수식이 전부 깨진다 —
 * `=C1*D1`은 #VALUE!, `=SUM(...)`은 쉼표 붙은 칸을 텍스트로 취급해 조용히 0. 같은 파일을
 * 읽기전용 뷰어로 열면 정상이었다(뷰어 경로는 buildFormulaInputGrid가 unformattedCellValue로
 * "235,000" → "235000"을 미리 되돌려 둔다 — 그 함수 문서 주석의 "한쪽만 고치면 편집기에서는
 * 맞는데 뷰어에서는 0 같은 불일치가 다시 생긴다"는 경고가 정확히 이 방향으로 재발한 것).
 */

let store: LazyGridStore | null = null;

afterEach(async () => {
  if (store) {
    await fs.rm(store.rowStorePath, { force: true, maxRetries: 5, retryDelay: 100 });
    store = null;
  }
});

async function ingestSample(): Promise<LazyGridStore> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('매출현황');
  sheet.getCell('A1').value = 12;
  sheet.getCell('B1').value = 235000;
  sheet.getCell('B1').numFmt = '#,##0'; // 표시 텍스트가 "235,000"이 되는 조건
  sheet.getCell('C1').value = { formula: 'A1*B1' };
  sheet.getCell('B2').value = 38000;
  sheet.getCell('B2').numFmt = '#,##0';
  sheet.getCell('D1').value = { formula: 'SUM(B1:B2)' };
  const buffer = await workbook.xlsx.writeBuffer();
  return ingestXlsxLazily(`lazy-display-text-${Date.now()}`, Buffer.from(buffer as ArrayBuffer));
}

describe('xlsx/formulas.ts: 쉼표 표시 형식 칸을 참조하는 수식 (LazyGridStore 경로)', () => {
  test('산술 수식이 #,##0 서식 칸을 숫자로 계산한다 (#VALUE! 회귀 방지)', async () => {
    store = await ingestSample();
    const result = await computeGridCellDisplayAsync(store, 0, 0, 2); // C1 = A1*B1
    expect(result).toBe(12 * 235000);
  });

  test('SUM이 쉼표 서식 칸들을 텍스트로 취급해 0이 되지 않는다', async () => {
    store = await ingestSample();
    const result = await computeGridCellDisplayAsync(store, 0, 0, 3); // D1 = SUM(B1:B2)
    expect(result).toBe(235000 + 38000);
  });
});
