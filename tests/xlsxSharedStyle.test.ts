import ExcelJS from 'exceljs';
import { describe, expect, test } from 'vitest';
import { applyStyleToCell } from '../src/xlsx/collab.js';

describe('저장: 같은 서식을 쓰던 칸들', () => {
  test('한 칸의 서식을 고쳐도 같은 서식을 쓰던 다른 칸은 그대로다(ExcelJS 공유 서식 객체)', async () => {
    const src = new ExcelJS.Workbook();
    const ws0 = src.addWorksheet('목록');
    for (const a of ['A1', 'A2', 'A3']) {
      ws0.getCell(a).value = a;
      ws0.getCell(a).font = { name: '맑은 고딕', bold: true };
      ws0.getCell(a).alignment = { vertical: 'middle' };
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load((await src.xlsx.writeBuffer()) as ArrayBuffer);
    const ws = wb.getWorksheet('목록')!;
    // 읽은 직후 세 칸은 같은 서식 객체를 가리킨다(이게 문제의 전제)
    expect(ws.getCell('A1').style).toBe(ws.getCell('A2').style);

    applyStyleToCell(ws.getCell('A1'), { a: 'r', va: 't', b: 1, ff: '맑은 고딕' });
    applyStyleToCell(ws.getCell('A3'), { fs: 16, ff: '맑은 고딕' });

    const out = new ExcelJS.Workbook();
    await out.xlsx.load((await wb.xlsx.writeBuffer()) as ArrayBuffer);
    const o = out.getWorksheet('목록')!;
    expect(o.getCell('A1').alignment).toMatchObject({ horizontal: 'right', vertical: 'top' });
    expect(o.getCell('A1').font?.bold).toBe(true);
    // 고치지 않은 A2는 원래대로(굵게·가운데 세로), A1의 맞춤이나 A3의 크기가 번지지 않는다
    expect(o.getCell('A2').font?.bold).toBe(true);
    expect(o.getCell('A2').alignment?.horizontal).toBeUndefined();
    expect(o.getCell('A2').alignment?.vertical).toBe('middle');
    expect(o.getCell('A2').font?.size).toBeUndefined();
    expect(o.getCell('A3').font?.size).toBe(12);
    expect(o.getCell('A3').alignment?.horizontal).toBeUndefined();
  });
});
