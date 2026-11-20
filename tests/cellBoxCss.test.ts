import { describe, expect, test } from 'vitest';
import fs from 'node:fs';
import { borderSideCss, cellBoxCss, cellBoxCssText, sanitizeCellStyle } from '../src/collab/cellBoxCss.js';
import { fromCellStyleAttrs } from '../src/xlsx/collab.js';
import { parseXlsxToGrid } from '../src/spreadsheet.js';

describe('칸 상자 CSS(테두리 선 종류·색, 무늬, 회전, 들여쓰기)', () => {
  test('테두리: 굵기만이면 실선, 선 종류·색이 있으면 그대로', () => {
    expect(borderSideCss({ bt: 1 }, 't')).toBe('1px solid var(--v-text)');
    expect(borderSideCss({ bt: 1, bd: { t: ['dashed', '#C00000'] } }, 't')).toBe('1px dashed #C00000');
    expect(borderSideCss({ bb: 2, bd: { b: ['double'] } }, 'b')).toBe('3px double var(--v-text)');
    expect(borderSideCss({ bd: { t: ['dashed'] } }, 't')).toBe(''); // 굵기가 지워졌으면(툴바 '없음') 안 그린다
  });
  test('무늬·회전·세로쓰기·들여쓰기', () => {
    expect(cellBoxCss({ pat: ['gray125'] }).td['background-color']).toBe('#dfdfdf');
    expect(cellBoxCss({ pat: ['darkGrid', '#4472C4', '#FFFFFF'] }).td['background-image']).toContain('90deg');
    expect(cellBoxCss({ rot: 45 }).inner.transform).toBe('rotate(-45deg)');
    expect(cellBoxCss({ rot: 135 }).inner.transform).toBe('rotate(45deg)');
    expect(cellBoxCss({ rot: 255 }).inner['writing-mode']).toBe('vertical-rl');
    expect(cellBoxCss({ ind: 2 }).td['padding-left']).toBe('21px');
    expect(cellBoxCssText({ bt: 1, bd: { t: ['dotted', '#C00000'] } }).x).toBe('border-top:1px dotted #C00000');
  });
  test('x02: 파일의 빨간 점선·무늬·회전·셀에 맞춤을 읽는다', async () => {
    const grid = await parseXlsxToGrid(fs.readFileSync('audit/office/corpus/xlsx/x02-merge-border-fill.xlsx'));
    const s = grid.styles;
    expect(s.some((st) => st.bd?.t?.[0] === 'dashed' && st.bd.t[1] === '#C00000')).toBe(true);
    expect(s.some((st) => st.pat?.[0] === 'lightTrellis')).toBe(true);
    expect(s.some((st) => st.rot === 255)).toBe(true);
    expect(s.some((st) => st.rot === 45)).toBe(true);
    expect(s.some((st) => st.shr === 1)).toBe(true);
    expect(s.some((st) => st.ind === 2)).toBe(true);
  });
  test('보안: 편집자가 보낸 색·선 종류·무늬에 CSS를 끼워 넣을 수 없다(서버 진입점과 보기 화면 서식 목록)', () => {
    const evil = {
      col: 'red;}body{background:url(//x/?leak)}',
      bg: '#FFFFFF',
      bt: 1,
      bd: { t: ['thin', 'red}body{background:url(//x)}'], r: ['evil;}', '#C00000'] },
      pat: ['darkGrid', '#000000}*{x:y', '#FFFFFF'],
      rot: '45;}',
      ind: 99,
      ff: '맑은 고딕";}body{x:y}',
      zz: 'unknown',
    } as never;
    const safe = sanitizeCellStyle(evil) as Record<string, unknown>;
    expect(safe.col).toBeUndefined();
    expect(safe.bg).toBe('#FFFFFF');
    expect(safe.bd).toEqual({ t: ['thin'] });
    expect(safe.pat).toEqual(['darkGrid', '#000000', '#FFFFFF']);
    expect(safe.rot).toBeUndefined();
    expect(safe.ind).toBeUndefined();
    expect(String(safe.ff)).not.toMatch(/[";{}]/);
    expect(safe.zz).toBeUndefined();
    expect(Object.values(cellBoxCssText(safe)).join(';')).not.toMatch(/[{}]|url/);
    const viaRoom = fromCellStyleAttrs(evil) as Record<string, unknown>;
    expect(viaRoom.col).toBeUndefined();
    expect(viaRoom.bd).toEqual({ t: ['thin'] });
  });
});
