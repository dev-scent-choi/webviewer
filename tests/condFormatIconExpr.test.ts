import { describe, expect, test } from 'vitest';
import fs from 'node:fs';
import { computeStats, evalCfFormula, paintCell, type CondFormatEntry } from '../src/collab/condFormatShared.js';
import { parseCondFormats } from '../src/xlsx/condFormatParse.js';
import { parseViewerGrid } from '../src/xlsx/viewerGrid.js';

const values: Record<string, string> = { '0:2': '1', '1:2': '0', '0:3': '완료', '1:3': '' };
const valueAt = (r: number, c: number) => values[`${r}:${c}`] ?? '';

describe('조건부 서식 수식 규칙(expression)', () => {
  test('상대 참조는 칸마다 옮기고 $는 고정', () => {
    expect(evalCfFormula('$C1>=1', 0, 0, 0, 0, valueAt)).toBe(true);
    expect(evalCfFormula('$C1>=1', 1, 0, 1, 0, valueAt)).toBe(false);
    expect(evalCfFormula('$C$1>=1', 1, 0, 1, 0, valueAt)).toBe(true);
  });
  test('함수·글자·논리', () => {
    expect(evalCfFormula('AND($D1="완료",$C1>0)', 0, 0, 0, 0, valueAt)).toBe(true);
    expect(evalCfFormula('MOD(ROW(),2)=0', 0, 0, 1, 0, valueAt)).toBe(true);
    expect(evalCfFormula('ISBLANK(D2)', 0, 0, 0, 0, valueAt)).toBe(true);
    expect(evalCfFormula('LEN(D1)&"자"="2자"', 0, 0, 0, 0, valueAt)).toBe(true);
  });
  test('모르는 함수·범위·문법 오류는 안 맞음', () => {
    expect(evalCfFormula('SUM(A1:A3)>0', 0, 0, 0, 0, valueAt)).toBe(false);
    expect(evalCfFormula('VLOOKUP(1,A:B,2)', 0, 0, 0, 0, valueAt)).toBe(false);
    expect(evalCfFormula('1/0', 0, 0, 0, 0, valueAt)).toBe(false);
    expect(evalCfFormula('(1', 0, 0, 0, 0, valueAt)).toBe(false);
  });
});

describe('아이콘 집합(iconSet)', () => {
  const xml =
    '<conditionalFormatting sqref="A1:A5"><cfRule type="iconSet" priority="1"><iconSet iconSet="3Arrows"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet></cfRule></conditionalFormatting>';
  const entries = parseCondFormats(xml, [], []) as CondFormatEntry[];
  test('파일 규칙을 읽는다', () => {
    expect(entries[0].rule).toMatchObject({ t: 'icon', set: '3Arrows', th: [{ type: 'percent', v: 33 }, { type: 'percent', v: 67 }] });
  });
  test('최소~최대의 백분율 문턱으로 아이콘을 고른다', () => {
    const stats = new Map([[entries[0].id, computeStats(entries[0].rule, ['0', '50', '100'])]]);
    const at = (v: string) => paintCell(entries, 0, 0, v, stats, () => '')?.icon?.ch;
    expect(at('0')).toBe('↓');
    expect(at('50')).toBe('→');
    expect(at('100')).toBe('↑');
  });
});

describe('보기 화면 조건부 서식 굽기', () => {
  test('x03: 색조·막대·아이콘·셀 강조·수식 규칙이 칸 서식에 들어간다', async () => {
    const grid = await parseViewerGrid(fs.readFileSync('audit/office/corpus/xlsx/x03-conditional-formatting.xlsx'));
    const styles = grid.styles;
    const rows = grid.sheets[0].rows;
    const styleAt = (r: number, c: number) => {
      const cell = rows[r][c];
      return cell === 0 || cell.s === undefined ? {} : styles[cell.s];
    };
    const col = (c: number) => Array.from({ length: 15 }, (_, i) => styleAt(i + 1, c));
    expect(col(1).filter((s) => s.bg).length).toBeGreaterThan(10); // B: 색조
    expect(col(2).filter((s) => s.cb).length).toBeGreaterThan(10); // C: 데이터 막대
    expect(col(4).filter((s) => s.ci).length).toBeGreaterThan(10); // E: 아이콘
    expect(col(3).some((s) => s.bg === '#FFC7CE')).toBe(true); // D: 0보다 작음
    expect(col(0).some((s) => s.b && s.bg === '#C6EFCE')).toBe(true); // A: 수식 규칙 $C2>=1
  });
});
