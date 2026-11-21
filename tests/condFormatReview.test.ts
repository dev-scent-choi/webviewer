import { describe, expect, test } from 'vitest';
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { computeStats, paintCell, shiftFormulaRefs, type CfStats, type CondFormatEntry } from '../src/collab/condFormatShared.js';
import { createCfFormulaFallback } from '../src/xlsx/cfFormulaEngine.js';
import { parseCondFormats } from '../src/xlsx/condFormatParse.js';
import { bakeCondFormats, CF_BAKE_BUDGET_MS, localDayKey } from '../src/xlsx/viewerCondFormat.js';
import { parseXlsxToGrid } from '../src/spreadsheet.js';
import { parseViewerGrid } from '../src/xlsx/viewerGrid.js';

const red = { bg: '#FFC7CE', col: '#9C0006' };
const green = { bg: '#C6EFCE' };
const today = new Date();
const iso = (d: Date) => d.toISOString().slice(0, 10);

function sheet(values: Record<string, string>) {
  const valueAt = (r: number, c: number) => values[`${r}:${c}`] ?? '';
  return { valueAt, paint: (entries: CondFormatEntry[], r: number, c: number, stats = new Map<string, CfStats>(), fb?: Parameters<typeof paintCell>[6]) => paintCell(entries, r, c, valueAt(r, c), stats, valueAt, fb) };
}

describe('검토 I2 — 날짜 칸과 TODAY()', () => {
  const past = new Date(Date.UTC(today.getFullYear() - 1, 0, 5));
  const future = new Date(Date.UTC(today.getFullYear() + 1, 0, 5));
  const { paint } = sheet({ '0:0': iso(past), '1:0': iso(future) });
  const A = [{ r0: 0, r1: 1, c0: 0, c1: 0 }];
  test('수식 규칙 $A1<TODAY() — 날짜 글자를 일련번호로 비교', () => {
    const e: CondFormatEntry[] = [{ id: 'x', ranges: A, rule: { t: 'expr', f: '$A1<TODAY()', s: red } }];
    expect(paint(e, 0, 0)?.style).toEqual(red);
    expect(paint(e, 1, 0)).toBeNull();
  });
  test('셀 값 < TODAY() (피연산자가 수식)', () => {
    const e: CondFormatEntry[] = [{ id: 'x', ranges: A, rule: { t: 'cell', op: 'lt', v1: 'TODAY()', s: red } }];
    expect(paint(e, 0, 0)?.style).toEqual(red);
    expect(paint(e, 1, 0)).toBeNull();
  });
});

describe('검토 I3 — 셀 값 규칙의 상대 참조·수식 피연산자', () => {
  const { paint } = sheet({ '1:0': '5', '1:1': '3', '4:0': '5', '4:1': '9', '0:5': '2' });
  const A = [{ r0: 1, r1: 9, c0: 0, c1: 0 }];
  test('"셀 값 > B2"는 A5에서 B5와 비교한다', () => {
    const e: CondFormatEntry[] = [{ id: 'x', ranges: A, rule: { t: 'cell', op: 'gt', v1: 'B2', s: red } }];
    expect(paint(e, 1, 0)?.style).toEqual(red); // 5 > 3
    expect(paint(e, 4, 0)).toBeNull(); // 5 > 9 아님
  });
  test('=$F$1*2 같은 식도 계산한다', () => {
    const e: CondFormatEntry[] = [{ id: 'x', ranges: A, rule: { t: 'cell', op: 'gt', v1: '=$F$1*2', s: red } }];
    expect(paint(e, 1, 0)?.style).toEqual(red); // 5 > 4
  });
  test('참조 옮기기: 상대만, $와 따옴표는 그대로, 열 범위도', () => {
    expect(shiftFormulaRefs('=COUNTIF($A:$A,$A2)>1', 3, 0)).toBe('=COUNTIF($A:$A,$A5)>1');
    expect(shiftFormulaRefs('=B2+$C$1+"A1"&D$4', 2, 1)).toBe('=C4+$C$1+"A1"&E$4');
    expect(shiftFormulaRefs('=SUM(A:A)', 0, 2)).toBe('=SUM(C:C)');
  });
});

describe('검토 I4 — 우선순위·true이면 중지·문턱값 종류', () => {
  const { paint } = sheet({ '0:0': '5' });
  const A = [{ r0: 0, r1: 0, c0: 0, c1: 0 }];
  test('true이면 중지: 1순위가 맞으면 2순위의 굵게가 섞이지 않는다', () => {
    const e: CondFormatEntry[] = [
      { id: '1', ranges: A, stop: 1, rule: { t: 'cell', op: 'gt', v1: '0', s: green } },
      { id: '2', ranges: A, rule: { t: 'cell', op: 'gt', v1: '1', s: { bg: '#FF0000', b: 1 } } },
    ];
    expect(paint(e, 0, 0)?.style).toEqual(green);
    const noStop = [{ ...e[0], stop: undefined }, e[1]];
    expect(paint(noStop, 0, 0)?.style).toEqual({ ...green, b: 1 });
  });
  test('색조 cfvo: num 문턱값', () => {
    const e: CondFormatEntry[] = [{ id: 's', ranges: A, rule: { t: 'scale', colors: ['#000000', '#FFFFFF'], th: [{ type: 'num', v: 0 }, { type: 'num', v: 20 }] } }];
    const stats = new Map([['s', computeStats(e[0].rule, ['5', '10'])]]);
    expect(paint(e, 0, 0, stats)?.scale).toBe('#404040'); // 5/20
  });
  test('데이터 막대 cfvo와 길이(minLength·maxLength)', () => {
    const e: CondFormatEntry[] = [{ id: 'b', ranges: A, rule: { t: 'bar', color: '#638EC6', th: [{ type: 'num', v: 0 }, { type: 'num', v: 10 }], len: [0, 100] } }];
    expect(paint(e, 0, 0, new Map([['b', computeStats(e[0].rule, ['5'])]]))?.bar?.ratio).toBeCloseTo(0.5);
  });
  test('파일의 stopIfTrue·cfvo를 읽는다', () => {
    const xml = `<conditionalFormatting sqref="A1:A9"><cfRule type="cellIs" dxfId="0" priority="1" stopIfTrue="1" operator="greaterThan"><formula>0</formula></cfRule>
<cfRule type="colorScale" priority="2"><colorScale><cfvo type="num" val="0"/><cfvo type="percentile" val="50"/><cfvo type="formula" val="$B$1"/><color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color rgb="FF63BE7B"/></colorScale></cfRule></conditionalFormatting>`;
    const entries = parseCondFormats(xml, [red], []);
    expect(entries[0].stop).toBe(1);
    expect(entries[1].rule).toMatchObject({ t: 'scale', th: [{ type: 'num', v: 0 }, { type: 'percentile', v: 50 }, { type: 'formula', f: '$B$1' }] });
  });
});

describe('검토 10 — 범위·함수가 든 수식 규칙은 수식 엔진으로', () => {
  const values: Record<string, string> = { '0:0': '사과', '1:0': '배', '2:0': '사과', '0:1': '10', '1:1': '20', '2:1': '30' };
  const { valueAt, paint } = sheet(values);
  const fb = { fallback: createCfFormulaFallback(valueAt, () => ({ rows: 3, cols: 2 })), tables: new Map(), extent: () => ({ rows: 3, cols: 2 }) };
  const A = [{ r0: 0, r1: 2, c0: 0, c1: 0 }];
  test('COUNTIF($A:$A,$A1)>1 — 중복 강조', () => {
    const e: CondFormatEntry[] = [{ id: 'x', ranges: A, rule: { t: 'expr', f: 'COUNTIF($A:$A,$A1)>1', s: red } }];
    expect(paint(e, 0, 0, undefined, fb)?.style).toEqual(red);
    expect(paint(e, 1, 0, undefined, fb)).toBeNull();
    expect(paint(e, 2, 0, undefined, fb)?.style).toEqual(red);
  });
  test('SUM 범위와 비교', () => {
    const e: CondFormatEntry[] = [{ id: 'x', ranges: A, rule: { t: 'expr', f: '$B1>AVERAGE($B$1:$B$3)', s: red } }];
    expect(paint(e, 2, 0, undefined, fb)?.style).toEqual(red);
    expect(paint(e, 0, 0, undefined, fb)).toBeNull();
  });
  test('엔진이 없으면(편집 화면에서 아직 불러오기 전) 안 맞음으로 둔다', () => {
    const e: CondFormatEntry[] = [{ id: 'x', ranges: A, rule: { t: 'expr', f: 'COUNTIF($A:$A,$A1)>1', s: red } }];
    expect(paint(e, 0, 0)).toBeNull();
  });
});

describe('검토 11 — x14 확장(아이콘 3Stars, 데이터 막대 설정)', () => {
  const xml = `<conditionalFormatting sqref="B1:B5"><cfRule type="dataBar" priority="2"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar><extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"><x14:id>{AAA}</x14:id></ext></extLst></cfRule></conditionalFormatting>
<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}"><x14:conditionalFormattings>
<x14:conditionalFormatting><x14:cfRule type="dataBar" id="{AAA}"><x14:dataBar minLength="0" maxLength="100" gradient="0"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/><x14:negativeFillColor rgb="FFFF0000"/></x14:dataBar></x14:cfRule><xm:sqref>B1:B5</xm:sqref></x14:conditionalFormatting>
<x14:conditionalFormatting><x14:cfRule type="iconSet" priority="1" id="{BBB}"><x14:iconSet iconSet="3Stars"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule><xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>
</x14:conditionalFormattings></ext></extLst>`;
  const entries = parseCondFormats(xml, [], []);
  test('x14 아이콘 집합을 읽어 우선순위대로 앞에, 저장용 priority(p)는 두지 않는다', () => {
    expect(entries[0].rule).toMatchObject({ t: 'icon', set: '3Stars', th: [{ type: 'percent', v: 33 }, { type: 'percent', v: 67 }] });
    expect(entries[0].p).toBeUndefined();
    expect(entries[0].xp).toBe(1);
  });
  test('데이터 막대의 x14 설정(길이 0~100, 자동 최소·최대, 음수 색)을 본 규칙에 붙인다', () => {
    expect(entries[1].rule).toMatchObject({ t: 'bar', len: [0, 100], th: [{ type: 'autoMin' }, { type: 'autoMax' }], neg: '#FF0000' });
  });
});

async function cfWorkbook(rules: number): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('s');
  for (let r = 1; r <= rules; r += 1) for (let c = 1; c <= 10; c += 1) ws.getCell(r, c).value = (r * c) % 97;
  const bytes = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
  const blocks = Array.from({ length: rules }, (_, i) => `<conditionalFormatting sqref="A${i + 1}:J${i + 1}"><cfRule type="cellIs" dxfId="0" priority="${i + 1}" operator="greaterThan"><formula>50</formula></cfRule></conditionalFormatting>`).join('');
  bytes['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(bytes['xl/worksheets/sheet1.xml']).replace('</sheetData>', `</sheetData>${blocks}`));
  bytes['xl/styles.xml'] = strToU8(strFromU8(bytes['xl/styles.xml']).replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'));
  return Buffer.from(zipSync(bytes));
}

describe('검토 C2 — 조각난 규칙 수천 개도 굽기가 거의 선형, 중간에 쉰다', () => {
  test('규칙 12,000개(행마다 하나) — 칠해지고 수 초 안에 끝난다, 이벤트 루프가 멈추지 않는다', async () => {
    const xlsx = await cfWorkbook(12_000);
    let ticks = 0;
    const timer = setInterval(() => (ticks += 1), 5);
    const t0 = performance.now();
    const grid = await parseViewerGrid(xlsx);
    const ms = performance.now() - t0;
    clearInterval(timer);
    const painted = grid.sheets[0].rows.flat().filter((c) => c && c.s !== undefined && grid.styles[c.s].bg === '#FFC7CE').length;
    expect(painted).toBeGreaterThan(1000);
    expect(ms).toBeLessThan(8000); // 예전 15.8초(제곱)
    expect(ticks).toBeGreaterThan(3); // 굽는 동안 타이머가 돌았다 = 이벤트 루프를 놓았다
  }, 60_000);
  test('TODAY() 규칙이 있으면 구운 날을 적는다(날짜가 바뀌면 보기 그리드를 다시 만든다)', async () => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('s');
    ws.getCell('A1').value = new Date(Date.UTC(2020, 0, 5));
    const bytes = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
    bytes['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(bytes['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1:A3"><cfRule type="expression" priority="1"><formula>$A1&lt;TODAY()</formula></cfRule></conditionalFormatting>'));
    const grid = await parseViewerGrid(Buffer.from(zipSync(bytes)));
    expect(grid.cfDay).toBe(localDayKey());
  });
});

async function bigBook(rows: number, cols: number, cf: string): Promise<Buffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('s');
  for (let r = 1; r <= rows; r += 1) for (let c = 1; c <= cols; c += 1) ws.getCell(r, c).value = c === 1 ? `코드${r % 997}` : (r * c) % 101;
  const b = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
  b['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(b['xl/worksheets/sheet1.xml']).replace('</sheetData>', `</sheetData>${cf}`));
  b['xl/styles.xml'] = strToU8(strFromU8(b['xl/styles.xml']).replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'));
  return Buffer.from(zipSync(b));
}

/** 굽기만 재기 — 걸린 시간, 가장 긴 이벤트 루프 멈춤, 힙 최고 증가 */
async function measureBake(x: Buffer) {
  const grid = await parseXlsxToGrid(x);
  const mem0 = process.memoryUsage().heapUsed;
  let peak = mem0;
  let maxGap = 0;
  let last = performance.now();
  const timer = setInterval(() => {
    const now = performance.now();
    maxGap = Math.max(maxGap, now - last);
    last = now;
    peak = Math.max(peak, process.memoryUsage().heapUsed);
  }, 1);
  const t0 = performance.now();
  const baked = await bakeCondFormats(grid, x);
  const ms = performance.now() - t0;
  clearInterval(timer);
  return { baked, ms, maxGap, memMb: (peak - mem0) / 1e6 };
}

describe('검토 N2 — 조건부 서식 굽기의 시간 예산·멈춤·메모리', () => {
  test('COUNTIF($A:$A,$A1)>1 중복 강조 2만 행: 범위 값 표 한 번 — 다 칠하고 예산 안, 멈춤 100ms 미만', async () => {
    const x = await bigBook(20_000, 2, '<conditionalFormatting sqref="A1:A20000"><cfRule type="expression" dxfId="0" priority="1"><formula>COUNTIF($A:$A,$A1)&gt;1</formula></cfRule></conditionalFormatting>');
    const { baked, ms, maxGap, memMb } = await measureBake(x);
    const painted = baked.sheets[0].rows.filter((row) => row[0] && row[0].s !== undefined && baked.styles[row[0].s].bg === '#FFC7CE').length;
    expect(painted).toBe(20_000);
    expect(baked.cfSkipped).toBeUndefined();
    expect(ms).toBeLessThan(CF_BAKE_BUDGET_MS + 1_000);
    expect(maxGap).toBeLessThan(100);
    expect(memMb).toBeLessThan(200);
  }, 120_000);
  test('겹치는 넓은 규칙 100개(3만×10): 예산에서 멈추고 "일부 생략"을 적는다, 멈춤 100ms 미만', async () => {
    const many = Array.from({ length: 100 }, (_, i) => `<conditionalFormatting sqref="A1:J30000"><cfRule type="cellIs" dxfId="0" priority="${i + 1}" operator="greaterThan"><formula>${i}</formula></cfRule></conditionalFormatting>`).join('');
    const { baked, ms, maxGap, memMb } = await measureBake(await bigBook(30_000, 10, many));
    expect(baked.cfSkipped).toBe(1);
    expect(ms).toBeLessThan(CF_BAKE_BUDGET_MS + 1_000);
    expect(maxGap).toBeLessThan(100);
    expect(memMb).toBeLessThan(300);
  }, 120_000);
  test('수식 엔진으로만 되는 무거운 규칙은 미리 어림해 건너뛰고 알린다', async () => {
    const x = await bigBook(20_000, 2, '<conditionalFormatting sqref="A1:A20000"><cfRule type="expression" dxfId="0" priority="1"><formula>MATCH($A1,$A$1:$A$20000,0)&gt;1</formula></cfRule></conditionalFormatting>');
    const { baked, ms } = await measureBake(x);
    expect(baked.cfSkipped).toBe(1);
    expect(ms).toBeLessThan(500);
  }, 120_000);
});

describe('검토 r3 — 시트 XML 풀기 실패', () => {
  test('조건부 서식만 건너뛰고(알림) 보기 그리드는 그대로', async () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('s').getCell('A1').value = 1;
    const grid = await parseXlsxToGrid(Buffer.from(await wb.xlsx.writeBuffer()));
    const out = await bakeCondFormats(grid, new Uint8Array([1, 2, 3, 4]));
    expect(out.cfSkipped).toBe(1);
    expect(out.sheets).toBe(grid.sheets);
  });
});

describe('무거운 규칙 어림(보기·편집 화면 공통 — 검토 r3 E1)', () => {
  test('2만 행을 칸마다 읽는 수식 엔진 규칙(MATCH)은 건너뛰고, COUNTIF·SUMPRODUCT(→COUNTIF)·작은 범위는 칠한다', async () => {
    const { cfRuleTooCostly, cfEngineRanges } = await import('../src/collab/condFormatShared.js');
    const entry = (f: string, r1 = 19_999) => ({ id: 'x', ranges: [{ r0: 0, c0: 0, r1, c1: 0 }], rule: { t: 'expr' as const, f, s: { bg: '#FF0000' } } });
    expect(cfRuleTooCostly(entry('MATCH($A1,$A$1:$A$20000,0)>1'), 20_000, 5)).toBe(true);
    // 검토 r4: SUMPRODUCT((범위=값)*1)는 COUNTIF로 바꿔 범위 값 표로 계산하므로 무겁지 않다
    expect(cfRuleTooCostly(entry('SUMPRODUCT(($A$1:$A$20000=$A1)*1)>1'), 20_000, 5)).toBe(false);
    expect(cfRuleTooCostly(entry('COUNTIF($A:$A,$A1)>1'), 20_000, 5)).toBe(false);
    expect(cfEngineRanges(entry('COUNTIF($A:$A,$A1)>1'), 20_000, 5)).toEqual([]);
    expect(cfRuleTooCostly(entry('MATCH($A1,$A$1:$A$50,0)>1', 49), 20_000, 5)).toBe(false);
    expect(cfEngineRanges(entry('MATCH($A1,$A$2:$A$50,0)>1'), 20_000, 5)).toEqual([{ r0: 1, r1: 49, cols: 1 }]);
    // 칸 몇 개짜리 규칙이라도 칸 하나가 너무 큰 범위를 읽으면 건너뛴다(칸 하나에 오래 걸린다)
    expect(cfRuleTooCostly(entry('MATCH($A1,$A:$C,0)>1', 0), 200_000, 5)).toBe(true);
  });
});

describe('검토 r4 — 엔진이 못 하는 규칙은 알리고, SUMPRODUCT 배열 비교·COUNTIFS는 칠한다', () => {
  const vals = ['a', 'b', 'a', 'c', 'a'];
  const valueAt = (r: number, c: number) => (c === 0 ? (vals[r] ?? '') : c === 1 ? String(r + 1) : '');
  const ext = () => ({ rows: 5, cols: 2 });
  const rule = (f: string): CondFormatEntry[] => [{ id: 'x', ranges: [{ r0: 0, c0: 0, r1: 4, c1: 0 }], rule: { t: 'expr', f, s: { bg: '#FF0000' } } }];
  const paintWith = (f: string, row: number) => {
    let unsupported = 0;
    const env = { fallback: createCfFormulaFallback(valueAt, ext), tables: new Map(), extent: ext, unsupported: () => (unsupported += 1) };
    return { bg: paintCell(rule(f), row, 0, valueAt(row, 0), new Map(), valueAt, env)?.style?.bg, unsupported };
  };
  test('SUMPRODUCT((범위=값)*1)·SUMPRODUCT(--(범위=값))는 COUNTIF로 바꿔 칠한다', async () => {
    const { rewriteSumproduct } = await import('../src/collab/condFormatShared.js');
    expect(rewriteSumproduct('SUMPRODUCT(($A$1:$A$5=$A1)*1)>1')).toBe('COUNTIF($A$1:$A$5,$A1)>1');
    expect(rewriteSumproduct('SUMPRODUCT(--($A1=$A$1:$A$5))>1')).toBe('COUNTIF($A$1:$A$5,$A1)>1');
    expect(paintWith('SUMPRODUCT(($A$1:$A$5=$A1)*1)>1', 0)).toEqual({ bg: '#FF0000', unsupported: 0 });
    expect(paintWith('SUMPRODUCT(--($A$1:$A$5=$A1))>1', 1)).toEqual({ bg: undefined, unsupported: 0 });
  });
  test('COUNTIFS는 수식 엔진으로 계산한다', () => {
    expect(paintWith('COUNTIFS($A$1:$A$5,$A1,$B$1:$B$5,">2")>0', 0).bg).toBe('#FF0000'); // a는 3·5행에도(B>2)
    expect(paintWith('COUNTIFS($A$1:$A$5,$A1,$B$1:$B$5,">4")>0', 1).bg).toBeUndefined();
  });
  test('엔진도 못 하는 규칙(다른 배열 비교·모르는 함수)은 칠하지 않고 "일부 생략"을 알린다', () => {
    expect(paintWith('SUMPRODUCT(($A$1:$A$5=$A1)*($B$1:$B$5>2))>0', 0)).toEqual({ bg: undefined, unsupported: 1 });
    expect(paintWith('NOSUCHFUNC($A1)>0', 0)).toEqual({ bg: undefined, unsupported: 1 });
    // 오류값(#N/A — 못 찾음)은 "안 맞음"일 뿐 알리지 않는다
    expect(paintWith('MATCH("zz",$A$1:$A$5,0)>0', 0)).toEqual({ bg: undefined, unsupported: 0 });
  });
  test('보기 화면 굽기: 엔진이 못 하는 규칙이 있으면 cfSkipped', async () => {
    const { bakeSheetCondFormats, CfBakeBudget, styleRegistry } = await import('../src/xlsx/viewerCondFormat.js');
    const budget = new CfBakeBudget();
    const lines = vals.map((v, r) => [{ v }, { v: String(r + 1) }] as const);
    await bakeSheetCondFormats(rule('NOSUCHFUNC($A1)>0'), { rows: 5, cols: 2 }, async (onRow) => lines.forEach((l, r) => onRow(r, l)), valueAt, styleRegistry([]), budget, () => undefined);
    expect(budget.skipped).toBe(true);
  });
});
