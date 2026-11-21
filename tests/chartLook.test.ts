import fs from 'node:fs';
import { describe, expect, test } from 'vitest';
import { parseChartXml } from '../src/xlsx/chartModel.js';
import { renderChartSvg } from '../src/xlsx/chartSvg.js';
import { readSheetCharts } from '../src/xlsx/drawingPassthrough.js';

const THEME = { accent1: '#4F81BD', accent2: '#C0504D', accent3: '#9BBB59', accent4: '#8064A2', accent5: '#4BACC6', accent6: '#F79646', dk1: '#000000', lt1: '#FFFFFF' };

const chart = (plot: string, extra = '') =>
  `<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>제목</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea>${plot}</c:plotArea>${extra}</c:chart></c:chartSpace>`;
const ser = (i: number, inner = '') =>
  `<c:ser><c:idx val="${i}"/><c:tx><c:v>계열${i}</c:v></c:tx>${inner}<c:cat><c:strLit><c:pt idx="0"><c:v>가</c:v></c:pt><c:pt idx="1"><c:v>나</c:v></c:pt><c:pt idx="2"><c:v>다</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>${i + 1}</c:v></c:pt><c:pt idx="1"><c:v>${i + 3}</c:v></c:pt><c:pt idx="2"><c:v>${i + 2}</c:v></c:pt></c:numLit></c:val></c:ser>`;

describe('차트 세부 모양(검토 13)', () => {
  test('계열 색(직접·테마+밝기), 누적, 범례 위치, 축 제목·숫자 서식, 데이터 레이블, 눈금선', () => {
    const xml = chart(
      `<c:barChart><c:barDir val="col"/><c:grouping val="stacked"/>${ser(0, '<c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr>')}${ser(1, '<c:spPr><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></c:spPr><c:dLbls><c:showVal val="1"/></c:dLbls>')}<c:axId val="1"/><c:axId val="2"/></c:barChart>` +
        '<c:catAx><c:axId val="1"/><c:title><c:tx><c:rich><a:p><a:r><a:t>월</a:t></a:r></a:p></c:rich></c:tx></c:title></c:catAx><c:valAx><c:axId val="2"/><c:majorGridlines/><c:title><c:tx><c:rich><a:p><a:r><a:t>원</a:t></a:r></a:p></c:rich></c:tx></c:title><c:numFmt formatCode="#,##0&quot;원&quot;" sourceLinked="0"/></c:valAx>',
      '<c:legend><c:legendPos val="r"/></c:legend>',
    );
    const parsed = parseChartXml(xml, THEME)!;
    expect(parsed.kind).toBe('column');
    expect(parsed.series[0].look?.color).toBe('#FF0000');
    expect(parsed.series[1].look?.color).toBe('#903C3A'); // C0504D 밝기 75%
    expect(parsed.series[1].look?.labels).toEqual({ val: 1 });
    expect(parsed.look).toMatchObject({ grouping: 'stacked', legend: 'r', catTitle: '월', valTitle: '원', valFmt: '#,##0"원"', gridlines: 1 });
    const svg = renderChartSvg(parsed.kind, { title: parsed.title, categories: ['가', '나', '다'], series: parsed.series.map((s) => ({ name: s.name, values: s.values.cache, look: s.look })), look: parsed.look }, 480, 300);
    expect(svg).toContain('fill="#FF0000"');
    expect(svg).toContain('원</text>'); // 축 제목·서식
    expect(svg).toContain('>8원<'); // 누적(최대 2+6) 눈금에 서식 "#,##0원"
  });

  test('콤보(막대+보조 축 꺾은선), 부드러운 선·표식', () => {
    const xml = chart(
      `<c:barChart><c:barDir val="col"/>${ser(0)}<c:axId val="1"/><c:axId val="2"/></c:barChart><c:lineChart>${ser(1, '<c:marker><c:symbol val="diamond"/></c:marker><c:smooth val="1"/>')}<c:axId val="3"/><c:axId val="4"/></c:lineChart>`,
    );
    const parsed = parseChartXml(xml, THEME)!;
    expect(parsed.series[1].look).toMatchObject({ kind: 'line', secondary: 1, smooth: 1, marker: 'diamond' });
    expect(parsed.look?.palette?.[0]).toBe('#4F81BD');
    const svg = renderChartSvg(parsed.kind, { title: '', categories: ['가', '나', '다'], series: parsed.series.map((s) => ({ name: s.name, values: s.values.cache, look: s.look })), look: parsed.look }, 480, 300);
    expect(svg).toMatch(/<path d="M[^"]*C/); // 부드러운 선(베지어)
    expect(svg).toContain('fill="#C0504D"/>'); // 테마 강조2 표식
  });

  test('도넛·영역·분산형을 읽고 그린다', () => {
    const doughnut = parseChartXml(chart(`<c:doughnutChart>${ser(0, '<c:dLbls><c:showPercent val="1"/></c:dLbls>')}<c:holeSize val="60"/></c:doughnutChart>`), THEME)!;
    expect(doughnut.kind).toBe('doughnut');
    expect(renderChartSvg('doughnut', { title: '', categories: ['가', '나', '다'], series: [{ name: 'a', values: [1, 1, 2], look: doughnut.series[0].look }], look: doughnut.look }, 300, 300)).toContain('50%');
    expect(parseChartXml(chart(`<c:areaChart>${ser(0)}<c:axId val="1"/><c:axId val="2"/></c:areaChart>`))?.kind).toBe('area');
    expect(parseChartXml(chart(`<c:scatterChart>${ser(0).replace(/cat>/g, 'xVal>').replace(/val>/g, 'yVal>')}<c:axId val="1"/><c:axId val="2"/></c:scatterChart>`))?.kind).toBe('scatter');
  });

  test('x06: 파일 차트의 테마 색(2007 팔레트)·부드러운 선을 읽는다', () => {
    const charts = readSheetCharts(fs.readFileSync('audit/office/corpus/xlsx/x06-charts.xlsx'));
    const line = charts.find((c) => c.kind === 'line')!;
    expect(line.look?.palette?.[0]).toBe('#4F81BD');
    expect(line.series.some((s) => s.look?.smooth)).toBe(true);
  });
  test('x06: 축을 지운 파일(c:delete)은 눈금 글자를 그리지 않는다(엑셀처럼), 축 제목은 그린다', () => {
    const charts = readSheetCharts(fs.readFileSync('audit/office/corpus/xlsx/x06-charts.xlsx'));
    const bar = charts.find((c) => c.kind === 'column')!;
    expect(bar.look).toMatchObject({ valHidden: 1, catHidden: 1, gridlines: 1 });
    const svg = renderChartSvg(bar.kind, { title: bar.title, categories: bar.categories.cache, series: bar.series.map((s) => ({ name: s.name, values: s.values.cache, look: s.look })), look: bar.look }, 480, 300);
    expect(svg).not.toContain('>200<');
    expect(svg).not.toContain('>1월<');
    expect(svg).toContain('>MWh<');
    const shown = renderChartSvg(bar.kind, { title: '', categories: ['1월'], series: [{ name: 'a', values: [200] }], look: { ...bar.look, valHidden: undefined, catHidden: undefined } }, 480, 300);
    expect(shown).toContain('>1월<');
  });
});
