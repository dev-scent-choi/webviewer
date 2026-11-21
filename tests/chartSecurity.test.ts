import { describe, expect, test } from 'vitest';
import { sanitizeGridChart, type GridChart } from '../src/xlsx/chartModel.js';
import { renderChartSvg } from '../src/xlsx/chartSvg.js';

const EVIL = '"/><img src=x onerror=alert(document.cookie)>';

function hostileChart(kind = 'column'): GridChart {
  return {
    id: 'u1',
    sheetName: 'S',
    kind: kind as GridChart['kind'],
    title: `<script>alert(1)</script>${EVIL}`,
    from: { row: 0, col: 0, dx: 0, dy: 0 },
    size: { width: 400, height: 300 },
    to: null,
    categories: { ref: null, cache: [EVIL, '나'] },
    series: [{ name: EVIL, nameRef: null, values: { ref: null, cache: [1, 2] }, look: { color: EVIL, marker: EVIL, pointColors: [EVIL], labels: { val: 1, fmt: EVIL } } as never }],
    source: 'new',
    look: { palette: [EVIL], legend: EVIL, grouping: EVIL, valTitle: EVIL } as never,
  };
}

describe('검토 N1 — 차트 SVG 저장형 XSS', () => {
  test.each(['column', 'bar', 'line', 'pie', 'doughnut', 'area', 'scatter', 'radar', 'bubble'])('%s: 악의적 색·이름·표식이 SVG를 빠져나가지 못한다', (kind) => {
    const c = hostileChart(kind);
    const svg = renderChartSvg(c.kind, { title: c.title, categories: c.categories.cache, series: c.series.map((s) => ({ name: s.name, values: s.values.cache, look: s.look })), look: c.look }, 400, 300);
    expect(svg).not.toMatch(/<img|<script/i); // 글자로는 남아도(&lt;img…) 태그·속성이 되지 않는다
    expect(svg).not.toMatch(/<[^>]*\sonerror=/i);
    for (const m of svg.matchAll(/(?:fill|stroke)="([^"]*)"/g)) expect(m[1]).toMatch(/^(#[0-9A-Fa-f]{6}|#fff|none|#[0-9a-f]{6})$/);
  });

  test('서버 검사: 편집자가 넣은 차트는 모양(look)을 버리고, 종류·수·글자를 맞춘다', () => {
    const safe = sanitizeGridChart({ ...hostileChart(), size: { width: 'x', height: 1e12 }, series: [{ name: 'a'.repeat(5000), values: { cache: ['<b>', 3] } }] }, true)!;
    expect(safe.look).toBeUndefined();
    expect(safe.series[0].look).toBeUndefined();
    expect(safe.series[0].name.length).toBe(200);
    expect(safe.series[0].values.cache).toEqual([0, 3]);
    expect(safe.size).toEqual({ width: 480, height: 10_000 });
    expect(sanitizeGridChart({ ...hostileChart(), kind: '<svg onload=x>' })).toBeNull();
  });

  test('파일 차트의 모양도 색은 #RRGGBB, 위치·누적은 알려진 이름만 남는다', () => {
    const safe = sanitizeGridChart({ ...hostileChart(), source: 'file', look: { palette: ['#123456', EVIL], legend: 'r', grouping: EVIL }, series: [{ name: 'a', values: { cache: [1] }, look: { color: EVIL, marker: 'diamond' } }] })!;
    expect(safe.look).toEqual({ palette: ['#123456'], legend: 'r' });
    expect(safe.series[0].look).toEqual({ marker: 'diamond' });
  });
});
