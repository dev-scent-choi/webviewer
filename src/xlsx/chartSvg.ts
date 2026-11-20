/**
 * 시트 차트 SVG — 편집 화면(gridCharts.ts)과 보기 화면(서버가 그림으로 내려 줌)이 같은 모양으로 그린다.
 * 브라우저·서버 둘 다 쓰므로 DOM을 쓰지 않는다(글자로 SVG를 만든다).
 */
import { cellsOf, type CellRangeRef, type ChartKind, type GridChart } from './chartModel.js';

const PALETTE = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'];

function esc(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function niceStep(range: number): number {
  if (range <= 0) return 1;
  const raw = range / 5;
  const pow = 10 ** Math.floor(Math.log10(raw));
  const unit = raw / pow;
  return (unit <= 1 ? 1 : unit <= 2 ? 2 : unit <= 5 ? 5 : 10) * pow;
}

function formatTick(value: number): string {
  const abs = Math.abs(value);
  if (abs >= 1e6) return `${Math.round(value / 1e5) / 10}M`;
  return Number.isInteger(value) ? value.toLocaleString() : String(Math.round(value * 100) / 100);
}

export interface ChartData {
  readonly title: string;
  readonly categories: readonly string[];
  readonly series: ReadonlyArray<{ name: string; values: readonly number[] }>;
}

/** 차트 SVG(오피스 기본 차트 모양 근사 — 회색 눈금선, 축 글자, 범례) */
export function renderChartSvg(kind: ChartKind, data: ChartData, width: number, height: number): string {
  const font = 'font-family="맑은 고딕, Malgun Gothic, sans-serif"';
  const parts: string[] = [`<rect width="${width}" height="${height}" fill="#fff"/>`];
  let top = 10;
  if (data.title) {
    parts.push(`<text x="${width / 2}" y="26" text-anchor="middle" font-size="15" fill="#595959" ${font}>${esc(data.title)}</text>`);
    top = 40;
  }
  const legendItems = kind === 'pie' ? data.categories : data.series.length > 1 ? data.series.map((s) => s.name) : [];
  const legendH = legendItems.length > 0 ? 24 : 0;
  if (legendItems.length > 0) {
    const itemW = Math.min(110, (width - 20) / legendItems.length);
    const startX = (width - itemW * legendItems.length) / 2;
    legendItems.forEach((label, i) => {
      const x = startX + i * itemW;
      parts.push(`<rect x="${x}" y="${height - 17}" width="9" height="9" fill="${PALETTE[i % PALETTE.length]}"/>`);
      parts.push(`<text x="${x + 13}" y="${height - 9}" font-size="11" fill="#595959" ${font}>${esc(String(label).slice(0, 12))}</text>`);
    });
  }
  if (kind === 'pie') {
    const values = (data.series[0]?.values ?? []).map((v) => Math.max(0, v));
    const total = values.reduce((a, b) => a + b, 0);
    const cx = width / 2;
    const cy = top + (height - top - legendH) / 2;
    const r = Math.max(10, Math.min(width / 2 - 20, (height - top - legendH) / 2 - 6));
    let angle = -Math.PI / 2;
    values.forEach((value, i) => {
      if (total <= 0 || value <= 0) return;
      const sweep = (value / total) * Math.PI * 2;
      const end = angle + sweep;
      const large = sweep > Math.PI ? 1 : 0;
      const path =
        sweep >= Math.PI * 2 - 1e-6
          ? `M ${cx - r} ${cy} a ${r} ${r} 0 1 0 ${2 * r} 0 a ${r} ${r} 0 1 0 ${-2 * r} 0`
          : `M ${cx} ${cy} L ${cx + r * Math.cos(angle)} ${cy + r * Math.sin(angle)} A ${r} ${r} 0 ${large} 1 ${cx + r * Math.cos(end)} ${cy + r * Math.sin(end)} Z`;
      parts.push(`<path d="${path}" fill="${PALETTE[i % PALETTE.length]}" stroke="#fff" stroke-width="1"/>`);
      angle = end;
    });
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${parts.join('')}</svg>`;
  }
  const all = data.series.flatMap((s) => s.values);
  const maxV = Math.max(0, ...all);
  const minV = Math.min(0, ...all);
  const step = niceStep(maxV - minV || 1);
  const axisMax = Math.ceil(maxV / step) * step || step;
  const axisMin = Math.floor(minV / step) * step;
  const horizontal = kind === 'bar';
  const left = horizontal ? 70 : 46;
  const right = width - 14;
  const bottom = height - legendH - (horizontal ? 22 : 26);
  const plotW = Math.max(10, right - left);
  const plotH = Math.max(10, bottom - top);
  const scale = (v: number) => (v - axisMin) / (axisMax - axisMin || 1);
  for (let v = axisMin; v <= axisMax + step / 2; v += step) {
    if (horizontal) {
      const x = left + scale(v) * plotW;
      parts.push(`<line x1="${x}" y1="${top}" x2="${x}" y2="${bottom}" stroke="#d9d9d9"/>`);
      parts.push(`<text x="${x}" y="${bottom + 15}" text-anchor="middle" font-size="10" fill="#595959" ${font}>${formatTick(v)}</text>`);
    } else {
      const y = bottom - scale(v) * plotH;
      parts.push(`<line x1="${left}" y1="${y}" x2="${right}" y2="${y}" stroke="#d9d9d9"/>`);
      parts.push(`<text x="${left - 6}" y="${y + 3.5}" text-anchor="end" font-size="10" fill="#595959" ${font}>${formatTick(v)}</text>`);
    }
  }
  const n = Math.max(1, data.categories.length);
  const band = (horizontal ? plotH : plotW) / n;
  const seriesCount = Math.max(1, data.series.length);
  const zero = scale(0);
  data.categories.forEach((label, i) => {
    const text = esc(String(label).slice(0, 10));
    if (horizontal) parts.push(`<text x="${left - 6}" y="${top + band * (i + 0.5) + 3.5}" text-anchor="end" font-size="10" fill="#595959" ${font}>${text}</text>`);
    else parts.push(`<text x="${left + band * (i + 0.5)}" y="${bottom + 16}" text-anchor="middle" font-size="10" fill="#595959" ${font}>${text}</text>`);
  });
  if (kind === 'line') {
    data.series.forEach((ser, s) => {
      const points = ser.values.map((v, i) => `${left + band * (i + 0.5)},${bottom - scale(v) * plotH}`).join(' ');
      parts.push(`<polyline points="${points}" fill="none" stroke="${PALETTE[s % PALETTE.length]}" stroke-width="2.25" stroke-linejoin="round" stroke-linecap="round"/>`);
    });
  } else {
    const groupW = band * 0.62;
    const barW = groupW / seriesCount;
    data.series.forEach((ser, s) => {
      ser.values.forEach((v, i) => {
        const a = Math.min(zero, scale(v));
        const b = Math.max(zero, scale(v));
        const offset = band * i + (band - groupW) / 2 + barW * s;
        if (horizontal) {
          parts.push(`<rect x="${left + a * plotW}" y="${top + offset}" width="${Math.max(0, (b - a) * plotW)}" height="${barW * 0.92}" fill="${PALETTE[s % PALETTE.length]}"/>`);
        } else {
          parts.push(`<rect x="${left + offset}" y="${bottom - b * plotH}" width="${barW * 0.92}" height="${Math.max(0, (b - a) * plotH)}" fill="${PALETTE[s % PALETTE.length]}"/>`);
        }
      });
    });
  }
  const axisLine = horizontal
    ? `<line x1="${left + zero * plotW}" y1="${top}" x2="${left + zero * plotW}" y2="${bottom}" stroke="#bfbfbf"/>`
    : `<line x1="${left}" y1="${bottom - zero * plotH}" x2="${right}" y2="${bottom - zero * plotH}" stroke="#bfbfbf"/>`;
  parts.push(axisLine);
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${parts.join('')}</svg>`;
}

/** 차트가 가리키는 칸의 지금 값으로 그릴 자료를 만든다(칸이 비면 파일에 저장된 값). valueText는 칸의 값 글자(서식 전) */
export function chartDataOf(chart: GridChart, valueText: (sheetName: string, row: number, col: number) => string): ChartData {
  const value = (ref: CellRangeRef, r: number, c: number) => valueText(ref.sheetName ?? chart.sheetName, r, c);
  const catRef = chart.categories.ref;
  const categories = catRef
    ? cellsOf(catRef).map(([r, c], i) => value(catRef, r, c) || chart.categories.cache[i] || '')
    : [...chart.categories.cache];
  const series = chart.series.map((ser) => {
    const ref = ser.values.ref;
    const values = ref
      ? cellsOf(ref).map(([r, c], i) => {
          const text = value(ref, r, c);
          return text === '' ? (ser.values.cache[i] ?? 0) : Number(text) || 0;
        })
      : [...ser.values.cache];
    const name = ser.nameRef ? value(ser.nameRef, ser.nameRef.r1, ser.nameRef.c1) || ser.name : ser.name;
    return { name, values };
  });
  const count = Math.max(categories.length, ...series.map((s) => s.values.length));
  const cats = categories.length > 0 ? categories : Array.from({ length: count }, (_, i) => String(i + 1));
  return { title: chart.title || '', categories: cats, series };
}
