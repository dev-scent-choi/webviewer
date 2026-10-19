import ExcelJS from 'exceljs';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import {
  computeStats,
  paintCell,
  parseSqref,
  shiftCondFormats,
  subtractRect,
  type CfStats,
  type CondFormatEntry,
} from '../src/collab/condFormatShared.js';
import { parseCondFormats, parseDxfs } from '../src/xlsx/condFormatParse.js';
import { applyEditedCondFormats, preserveUnsupportedCfRules } from '../src/xlsx/condFormatExport.js';
import type { LazyGridSheetMeta } from '../src/xlsx/lazyGridStore.js';

const STYLES = `<styleSheet><dxfs count="2">
<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
<dxf><font><b/><color theme="1"/></font></dxf>
</dxfs></styleSheet>`;
const THEME = ['FFFFFF', '000000', 'EEECE1', '1F497D', '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646'];

describe('파일의 조건부 서식 읽기', () => {
  const dxfs = parseDxfs(STYLES, THEME);
  test('dxf(강조 서식) — 글자 색·채우기·굵게, 테마 색', () => {
    expect(dxfs).toEqual([{ bg: '#FFC7CE', col: '#9C0006' }, { col: '#000000', b: 1 }]); // theme 1 = THEME[1]
  });

  const sheet = `<worksheet><sheetData/>
<conditionalFormatting sqref="A1:A10 C1:C5"><cfRule type="cellIs" dxfId="0" priority="2" operator="greaterThan"><formula>10</formula></cfRule></conditionalFormatting>
<conditionalFormatting sqref="B1:B9"><cfRule type="duplicateValues" dxfId="1" priority="1"/><cfRule type="top10" dxfId="0" priority="3" rank="3" bottom="1"/></conditionalFormatting>
<conditionalFormatting sqref="D1:D9"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFF8696B"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>
<conditionalFormatting sqref="E:E"><cfRule type="expression" dxfId="0" priority="5"><formula>MOD(ROW(),2)=0</formula></cfRule></conditionalFormatting>
<conditionalFormatting sqref="F1:F3"><cfRule type="beginsWith" dxfId="0" priority="6" operator="beginsWith" text="가&amp;나"><formula>LEFT(F1,3)="가&amp;나"</formula></cfRule></conditionalFormatting>
<extLst><ext><x14:conditionalFormattings><x14:conditionalFormatting><x14:cfRule type="dataBar" id="{X}"/></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>
</worksheet>`;
  const entries = parseCondFormats(sheet, dxfs, THEME);
  test('우선순위 순, 종류별 형식, 화면에 못 그리는 수식 규칙은 other, extLst는 읽지 않음', () => {
    expect(entries.map((e) => [e.p, e.rule.t])).toEqual([[1, 'dup'], [2, 'cell'], [3, 'top'], [4, 'scale'], [5, 'other'], [6, 'text']]);
    expect(entries[1]).toMatchObject({ ranges: [{ r0: 0, r1: 9, c0: 0, c1: 0 }, { r0: 0, r1: 4, c0: 2, c1: 2 }], rule: { t: 'cell', op: 'gt', v1: '10', s: { bg: '#FFC7CE', col: '#9C0006' } } });
    expect(entries[2].rule).toEqual({ t: 'top', n: 3, bottom: 1, s: { bg: '#FFC7CE', col: '#9C0006' } });
    expect(entries[3].rule).toEqual({ t: 'scale', colors: ['#F8696B', '#63BE7B'] });
    expect(entries[5].rule).toEqual({ t: 'text', op: 'begins', text: '가&나', s: { bg: '#FFC7CE', col: '#9C0006' } });
  });
  test('열 전체(E:E) 범위는 편집기 최대 행·열까지', () => {
    expect(entries[4].ranges[0]).toMatchObject({ c0: 4, c1: 4, r0: 0 });
    expect(parseSqref('B:C 3:4', 99, 9)).toEqual([{ r0: 0, r1: 99, c0: 1, c1: 2 }, { r0: 2, r1: 3, c0: 0, c1: 9 }]);
  });
});

describe('칸 판정(화면 표시)', () => {
  const red = { bg: '#FFC7CE', col: '#9C0006' };
  const grid: Record<string, string> = { '0:0': '5', '1:0': '15', '2:0': '', '3:0': '사과', '0:1': '12' };
  const valueAt = (r: number, c: number) => grid[`${r}:${c}`] ?? '';
  const paint = (entries: CondFormatEntry[], r: number, c: number, stats = new Map<string, CfStats>()) =>
    paintCell(entries, r, c, valueAt(r, c), stats, valueAt);
  const A = [{ r0: 0, r1: 3, c0: 0, c1: 0 }];

  test('보다 큼 — 숫자 비교, 빈 칸은 0(엑셀과 같음), 다른 칸 참조', () => {
    const gt: CondFormatEntry[] = [{ id: 'a', ranges: A, rule: { t: 'cell', op: 'gt', v1: '10', s: red } }];
    expect(paint(gt, 1, 0)?.style).toEqual(red);
    expect(paint(gt, 0, 0)).toBeNull();
    const lt: CondFormatEntry[] = [{ id: 'a', ranges: A, rule: { t: 'cell', op: 'lt', v1: '=$B$1', s: red } }];
    expect(paint(lt, 2, 0)?.style).toEqual(red); // 빈 칸 = 0 < 12
    expect(paint(lt, 1, 0)).toBeNull(); // 15 < 12 아님
    expect(paint(lt, 3, 0)).toBeNull(); // 글자는 크기 비교 안 함
  });
  test('사이·같음(글자)·텍스트 포함', () => {
    const between: CondFormatEntry[] = [{ id: 'a', ranges: A, rule: { t: 'cell', op: 'between', v1: '1', v2: '10', s: red } }];
    expect(paint(between, 0, 0)?.style).toEqual(red);
    expect(paint(between, 1, 0)).toBeNull();
    const eq: CondFormatEntry[] = [{ id: 'a', ranges: A, rule: { t: 'cell', op: 'eq', v1: '="사과"', s: red } }];
    expect(paint(eq, 3, 0)?.style).toEqual(red);
    const text: CondFormatEntry[] = [{ id: 'a', ranges: A, rule: { t: 'text', op: 'contains', text: '사', s: red } }];
    expect(paint(text, 3, 0)?.style).toEqual(red);
    expect(paint(text, 0, 0)).toBeNull();
  });
  test('중복·상위·평균·색조·막대는 범위 통계로', () => {
    const values = ['3', '7', '7', '10', '', 'x'];
    expect(computeStats({ t: 'dup', s: red }, values).counts?.get('7')).toBe(2);
    expect(computeStats({ t: 'top', n: 2, s: red }, values).threshold).toBe(7);
    expect(computeStats({ t: 'top', n: 1, bottom: 1, s: red }, values).threshold).toBe(3);
    expect(computeStats({ t: 'avg', s: red }, values).mean).toBeCloseTo(6.75);
    expect(computeStats({ t: 'bar', color: '#638EC6' }, values)).toEqual({ min: 0, max: 10 });
    const bar: CondFormatEntry[] = [{ id: 'b', ranges: A, rule: { t: 'bar', color: '#638EC6' } }];
    expect(paint(bar, 0, 0, new Map([['b', { min: 0, max: 10 }]]))?.bar).toEqual({ color: '#638EC6', ratio: 0.5 });
    const scale: CondFormatEntry[] = [{ id: 's', ranges: A, rule: { t: 'scale', colors: ['#000000', '#FFFFFF'] } }];
    expect(paint(scale, 0, 0, new Map([['s', { min: 0, max: 10 }]]))?.scale).toBe('#808080');
  });
  test('같은 칸에 두 규칙 — 앞(우선순위 높은) 규칙의 속성이 이긴다', () => {
    const entries: CondFormatEntry[] = [
      { id: 'hi', ranges: A, rule: { t: 'cell', op: 'gt', v1: '1', s: { bg: '#111111' } } },
      { id: 'lo', ranges: A, rule: { t: 'cell', op: 'gt', v1: '1', s: { bg: '#222222', col: '#333333' } } },
    ];
    expect(paint(entries, 1, 0)?.style).toEqual({ bg: '#111111', col: '#333333' });
  });
});

describe('행/열 이동·선택 지우기', () => {
  const e: CondFormatEntry = { id: 'x', ranges: [{ r0: 2, r1: 5, c0: 0, c1: 1 }], rule: { t: 'other' } };
  test('위에 삽입하면 밀리고, 안쪽에 삽입하면 늘고, 겹쳐 지우면 줄고, 다 지우면 규칙이 빠진다', () => {
    expect(shiftCondFormats([e], 'row', 0, 2, true)[0].ranges).toEqual([{ r0: 4, r1: 7, c0: 0, c1: 1 }]);
    expect(shiftCondFormats([e], 'row', 3, 1, true)[0].ranges).toEqual([{ r0: 2, r1: 6, c0: 0, c1: 1 }]);
    expect(shiftCondFormats([e], 'row', 1, 2, false)[0].ranges).toEqual([{ r0: 1, r1: 3, c0: 0, c1: 1 }]);
    expect(shiftCondFormats([e], 'col', 0, 2, false)).toEqual([]);
  });
  test('사각형 빼기 — 가운데를 빼면 네 조각', () => {
    expect(subtractRect({ r0: 0, r1: 4, c0: 0, c1: 4 }, { r0: 1, r1: 2, c0: 1, c1: 2 })).toEqual([
      { r0: 0, r1: 0, c0: 0, c1: 4 },
      { r0: 3, r1: 4, c0: 0, c1: 4 },
      { r0: 1, r1: 2, c0: 0, c1: 0 },
      { r0: 1, r1: 2, c0: 3, c1: 4 },
    ]);
    expect(subtractRect({ r0: 0, r1: 1, c0: 0, c1: 0 }, { r0: 0, r1: 9, c0: 0, c1: 9 })).toEqual([]);
  });
});

async function roundTrip(wb: ExcelJS.Workbook): Promise<string> {
  const buf = new Uint8Array(await wb.xlsx.writeBuffer());
  return strFromU8(unzipSync(buf)['xl/worksheets/sheet1.xml']);
}

describe('저장 파일에 쓰기', () => {
  test('ExcelJS가 못 쓰던 중복 값·다음으로 시작 규칙이 수식 규칙으로 보존된다(예전엔 사라짐)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 1;
    (ws as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings = [
      { ref: 'A1:A5', rules: [{ type: 'duplicateValues', priority: 1, style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] },
      { ref: 'B1:B5', rules: [{ type: 'beginsWith', operator: 'beginsWith', text: 'a', priority: 2, formulae: ['LEFT(B1,1)="a"'], style: {} }] },
    ];
    preserveUnsupportedCfRules(ws);
    const xml = await roundTrip(wb);
    // ExcelJS는 수식 안의 따옴표를 &quot;로 쓴다
    expect(xml).toContain('<formula>AND(A1&lt;&gt;&quot;&quot;,COUNTIF($A$1:$A$5,A1)&gt;1)</formula>');
    expect(xml).toContain('<formula>LEFT(B1,1)=&quot;a&quot;</formula>');
    expect((xml.match(/<cfRule /g) ?? []).length).toBe(2);
  });

  test('편집한 시트만 목록을 다시 쓴다 — 원래 규칙은 그대로(범위만 새로), 새 규칙 추가, 지운 규칙 빠짐, 모르는 규칙 유지', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 1;
    ws.addConditionalFormatting({ ref: 'A1:A5', rules: [{ type: 'expression', priority: 1, formulae: ['MOD(ROW(),2)=0'], style: { font: { bold: true } } }] });
    ws.addConditionalFormatting({ ref: 'B1:B5', rules: [{ type: 'cellIs', priority: 2, operator: 'greaterThan', formulae: ['3'], style: {} }] });
    ws.addConditionalFormatting({ ref: 'C1:C5', rules: [{ type: 'expression', priority: 9, formulae: ['TRUE'], style: {} }] }); // 편집기가 모르는 규칙
    const meta = {
      condFormatsEdited: true,
      condFormatFilePriorities: [1, 2],
      condFormats: [
        { id: 'f1', p: 1, ranges: [{ r0: 1, r1: 5, c0: 0, c1: 0 }], rule: { t: 'other' } }, // 한 행 밀림
        { id: 'n1', ranges: [{ r0: 0, r1: 9, c0: 3, c1: 3 }], rule: { t: 'cell', op: 'lt', v1: '사과', s: { bg: '#FFC7CE', col: '#9C0006' } } },
        // priority 2(B열 규칙)는 지워짐
      ],
    } as unknown as LazyGridSheetMeta;
    applyEditedCondFormats([meta], [ws], new Set());
    const xml = await roundTrip(wb);
    expect(xml).toContain('<conditionalFormatting sqref="A2:A6"><cfRule type="expression"');
    expect(xml).toContain('<formula>MOD(ROW(),2)=0</formula>');
    expect(xml).toMatch(/<conditionalFormatting sqref="D1:D10"><cfRule type="cellIs" dxfId="\d+" priority="2" operator="lessThan"><formula>&quot;사과&quot;<\/formula>/);
    expect(xml).not.toContain('sqref="B1:B5"');
    expect(xml).toContain('<conditionalFormatting sqref="C1:C5"><cfRule type="expression" dxfId="2" priority="3">');
    const back = new ExcelJS.Workbook();
    await back.xlsx.load(await wb.xlsx.writeBuffer());
    expect((back.worksheets[0] as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings).toHaveLength(3);
  });

  test('편집하지 않은 시트는 손대지 않는다', () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('S');
    const original = [{ ref: 'A1:A2', rules: [{ type: 'cellIs', priority: 1, operator: 'equal', formulae: ['1'], style: {} }] }];
    (ws as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings = original;
    applyEditedCondFormats([{ condFormats: [] } as unknown as LazyGridSheetMeta], [ws], new Set());
    expect((ws as unknown as { conditionalFormattings: unknown[] }).conditionalFormattings).toBe(original);
  });
});
