import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { applyEditedMerges } from '../src/xlsx/collab.js';
import type { LazyGridSheetMeta } from '../src/xlsx/lazyGridStore.js';

function mergesOf(sheet: ExcelJS.Worksheet): string[] {
  return Object.values((sheet as unknown as { _merges: Record<string, { range: string }> })._merges)
    .map((m) => m.range)
    .sort();
}

function meta(merges: LazyGridSheetMeta['merges'], mergesEdited?: boolean): LazyGridSheetMeta {
  return { merges, mergesEdited } as unknown as LazyGridSheetMeta;
}

describe('편집기에서 바꾼 병합을 저장 파일에 쓰기', () => {
  test('병합을 편집한 시트만 sheetMeta.merges로 다시 건다(기존 병합은 풀린다)', () => {
    const wb = new ExcelJS.Workbook();
    const edited = wb.addWorksheet('edited');
    const untouched = wb.addWorksheet('untouched');
    edited.mergeCells('A1:B1');
    untouched.mergeCells('C3:D4');
    applyEditedMerges(
      { sheetMeta: [meta([{ r: 2, c: 1, rs: 2, cs: 2 }], true), meta([{ r: 9, c: 9, rs: 2, cs: 2 }])] },
      [edited, untouched],
      new Set(),
    );
    expect(mergesOf(edited)).toEqual(['A2:B3']);
    expect(mergesOf(untouched)).toEqual(['C3:D4']); // mergesEdited가 없으면 손대지 않는다
  });

  test('병합을 모두 풀면(셀 분할) 병합이 없어지고, 지워진 시트·한 칸짜리는 건너뛴다', () => {
    const wb = new ExcelJS.Workbook();
    const a = wb.addWorksheet('a');
    const b = wb.addWorksheet('b');
    a.mergeCells('A1:C1');
    b.mergeCells('A1:B2');
    applyEditedMerges({ sheetMeta: [meta([{ r: 5, c: 5, rs: 1, cs: 1 }], true), meta([], true)] }, [a, b], new Set([1]));
    expect(mergesOf(a)).toEqual([]);
    expect(mergesOf(b)).toEqual(['A1:B2']); // 지워진 시트 자리는 건드리지 않는다
  });

  test('왼쪽 위 값은 남는다', () => {
    const wb = new ExcelJS.Workbook();
    const s = wb.addWorksheet('s');
    s.getCell('B2').value = '제목';
    applyEditedMerges({ sheetMeta: [meta([{ r: 2, c: 2, rs: 1, cs: 3 }], true)] }, [s], new Set());
    expect(mergesOf(s)).toEqual(['B2:D2']);
    expect(s.getCell('B2').value).toBe('제목');
  });
});
