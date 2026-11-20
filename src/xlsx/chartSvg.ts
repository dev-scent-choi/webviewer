/**
 * 시트 차트 SVG — 편집 화면(gridCharts.ts)과 보기 화면(서버가 그림으로 내려 줌)이 같은 모양으로 그린다.
 * 브라우저·서버 둘 다 쓰므로 DOM을 쓰지 않는다(글자로 SVG를 만든다).
 * 파일 차트의 모양(chartModel.ts ChartLook·SeriesLook — 계열 색·테마 색, 누적·100% 누적, 부드러운 선·표식, 데이터 레이블,
 * 축 제목, 범례 위치, 축 숫자 서식, 보조 축, 콤보, 영역·도넛·분산형)을 따른다. 새로 넣은 차트는 오피스 기본 모양.
 */
import { formatNumberCode, FILL_MARK } from '../collab/numberFormat.js';
import { cellsOf, safeChartColor, type CellRangeRef, type ChartKind, type ChartLabels, type ChartLook, type GridChart, type SeriesLook } from './chartModel.js';

const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E'];
const FONT = 'font-family="맑은 고딕, Malgun Gothic, sans-serif"';
const INK = '#595959';

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

function formatTick(value: number, fmt?: string): string {
  if (fmt) return formatNumberCode(value, fmt).text.replace(FILL_MARK, ' ').trim();
  const abs = Math.abs(value);
  if (abs >= 1e6) return `${Math.round(value / 1e5) / 10}M`;
  return Number.isInteger(value) ? value.toLocaleString('en-US') : String(Math.round(value * 100) / 100);
}

export interface ChartData {
  readonly title: string;
  readonly categories: readonly string[];
  readonly series: ReadonlyArray<{ name: string; values: readonly number[]; look?: SeriesLook; sizes?: readonly number[] }>;
  readonly look?: ChartLook;
}

const text = (x: number, y: number, s: string, opts = '') => `<text x="${x}" y="${y}" font-size="10" fill="${INK}" ${FONT} ${opts}>${esc(s)}</text>`;

/** 표식(circle·square·diamond·triangle·x …) */
function marker(symbol: string, x: number, y: number, color: string): string {
  const r = 3.5;
  if (symbol === 'square') return `<rect x="${x - r}" y="${y - r}" width="${2 * r}" height="${2 * r}" fill="${color}"/>`;
  if (symbol === 'diamond') return `<path d="M${x} ${y - r - 1}L${x + r + 1} ${y}L${x} ${y + r + 1}L${x - r - 1} ${y}Z" fill="${color}"/>`;
  if (symbol === 'triangle') return `<path d="M${x} ${y - r - 1}L${x + r + 1} ${y + r}L${x - r - 1} ${y + r}Z" fill="${color}"/>`;
  if (symbol === 'x' || symbol === 'star' || symbol === 'plus') return `<path d="M${x - r} ${y - r}L${x + r} ${y + r}M${x - r} ${y + r}L${x + r} ${y - r}" stroke="${color}" stroke-width="1.5"/>`;
  return `<circle cx="${x}" cy="${y}" r="${r}" fill="${color}"/>`;
}

/** 점들을 잇는 선(부드러운 선은 Catmull-Rom → 베지어) */
function linePath(points: ReadonlyArray<readonly [number, number]>, smooth: boolean): string {
  if (points.length === 0) return '';
  if (!smooth || points.length < 3) return `M${points.map(([x, y]) => `${x} ${y}`).join('L')}`;
  let d = `M${points[0][0]} ${points[0][1]}`;
  for (let i = 0; i < points.length - 1; i += 1) {
    const p0 = points[Math.max(0, i - 1)];
    const p1 = points[i];
    const p2 = points[i + 1];
    const p3 = points[Math.min(points.length - 1, i + 2)];
    d += `C${p1[0] + (p2[0] - p0[0]) / 6} ${p1[1] + (p2[1] - p0[1]) / 6} ${p2[0] - (p3[0] - p1[0]) / 6} ${p2[1] - (p3[1] - p1[1]) / 6} ${p2[0]} ${p2[1]}`;
  }
  return d;
}

function labelText(labels: ChartLabels, value: number, category: string, pct?: number, seriesName = ''): string {
  const parts: string[] = [];
  if (labels.ser && seriesName) parts.push(seriesName);
  if (labels.cat) parts.push(category);
  if (labels.val) parts.push(formatTick(value, labels.fmt));
  if (labels.pct && pct !== undefined) parts.push(`${Math.round(pct * 100)}%`);
  return parts.join(', ');
}

/** 차트 SVG(오피스 차트 모양 근사 — 눈금선, 축 글자, 범례) */
export function renderChartSvg(kind: ChartKind, data: ChartData, width: number, height: number): string {
  const look = data.look;
  // 색은 모두 이 두 함수를 지난다 — #RRGGBB가 아니면(공유 맵으로 들어온 망가진·악의적 값) 기본 팔레트 색
  const palette = (look?.palette ?? []).map(safeChartColor).filter((c): c is string => Boolean(c));
  const paletteAt = (i: number) => (palette.length > 0 ? palette : PALETTE)[i % (palette.length > 0 ? palette.length : PALETTE.length)];
  const colorOf = (i: number) => safeChartColor(data.series[i]?.look?.color) ?? paletteAt(i);
  const pointColor = (i: number) => safeChartColor(data.series[0]?.look?.pointColors?.[i]) ?? paletteAt(i);
  const lineLike = (i: number) => {
    const k = data.series[i]?.look?.kind ?? kind;
    return k === 'line' || k === 'radar' || (k === 'scatter' && Boolean(data.series[i]?.look?.smooth));
  };
  const parts: string[] = [`<rect width="${width}" height="${height}" fill="#fff"/>`];
  let top = 10;
  let bottom = height - 8;
  let left = 8;
  let right = width - 10;
  if (data.title) {
    parts.push(`<text x="${width / 2}" y="24" text-anchor="middle" font-size="14" fill="${INK}" ${FONT}>${esc(data.title)}</text>`);
    top = 36;
  }
  const round = kind === 'pie' || kind === 'doughnut';
  // 범례 — 파일 차트는 파일의 위치(없으면 범례 없음), 새 차트는 계열이 둘 이상이거나 원형이면 아래
  const legendPos = look ? look.legend : data.series.length > 1 || round ? 'b' : undefined;
  const legendItems = round
    ? data.categories.map((label, i) => ({ label, color: pointColor(i), line: false }))
    : data.series.map((s, i) => ({ label: s.name, color: colorOf(i), line: lineLike(i) }));
  // 범례 기호 — 꺾은선 계열은 선(엑셀과 같다), 나머지는 네모
  const legendMark = (x: number, y: number, item: { color: string; line: boolean }) =>
    item.line ? `<line x1="${x - 2}" y1="${y + 4.5}" x2="${x + 11}" y2="${y + 4.5}" stroke="${item.color}" stroke-width="2.25"/>` : `<rect x="${x}" y="${y}" width="9" height="9" fill="${item.color}"/>`;
  if (legendPos && legendItems.length > 0) {
    if (legendPos === 'r' || legendPos === 'tr' || legendPos === 'l') {
      const w = Math.min(110, width * 0.3);
      const x0 = legendPos === 'l' ? 8 : width - w;
      const y0 = Math.max(top, (height - legendItems.length * 16) / 2);
      legendItems.forEach((item, i) => {
        parts.push(legendMark(x0, y0 + i * 16, item));
        parts.push(text(x0 + 13, y0 + i * 16 + 8, String(item.label).slice(0, 14)));
      });
      if (legendPos === 'l') left = w + 12;
      else right = width - w - 6;
    } else {
      const itemW = Math.min(110, (width - 20) / legendItems.length);
      const startX = (width - itemW * legendItems.length) / 2;
      const y = legendPos === 't' ? top + 4 : height - 17;
      legendItems.forEach((item, i) => {
        const x = startX + i * itemW;
        parts.push(legendMark(x, y, item));
        parts.push(text(x + 13, y + 8, String(item.label).slice(0, 12)));
      });
      if (legendPos === 't') top += 22;
      else bottom = height - 26;
    }
  }

  if (round) {
    const values = (data.series[0]?.values ?? []).map((v) => Math.max(0, v));
    const total = values.reduce((a, b) => a + b, 0);
    const cx = (left + right) / 2;
    const cy = (top + bottom) / 2;
    const r = Math.max(10, Math.min((right - left) / 2 - 10, (bottom - top) / 2 - 6));
    const hole = kind === 'doughnut' ? r * Math.min(0.9, Math.max(0.1, (look?.hole ?? 50) / 100)) : 0;
    const labels = data.series[0]?.look?.labels ?? look?.labels;
    let angle = -Math.PI / 2;
    values.forEach((value, i) => {
      if (total <= 0 || value <= 0) return;
      const sweep = (value / total) * Math.PI * 2;
      const end = angle + sweep;
      const large = sweep > Math.PI ? 1 : 0;
      const color = pointColor(i);
      const pt = (rad: number, a: number) => `${cx + rad * Math.cos(a)} ${cy + rad * Math.sin(a)}`;
      const path =
        sweep >= Math.PI * 2 - 1e-6
          ? `M ${cx - r} ${cy} a ${r} ${r} 0 1 0 ${2 * r} 0 a ${r} ${r} 0 1 0 ${-2 * r} 0`
          : hole > 0
            ? `M ${pt(r, angle)} A ${r} ${r} 0 ${large} 1 ${pt(r, end)} L ${pt(hole, end)} A ${hole} ${hole} 0 ${large} 0 ${pt(hole, angle)} Z`
            : `M ${cx} ${cy} L ${pt(r, angle)} A ${r} ${r} 0 ${large} 1 ${pt(r, end)} Z`;
      parts.push(`<path d="${path}" fill="${color}" stroke="#fff" stroke-width="1" fill-rule="evenodd"/>`);
      if (labels) {
        const mid = angle + sweep / 2;
        const lr = hole > 0 ? (r + hole) / 2 : r * 0.65;
        parts.push(text(cx + lr * Math.cos(mid), cy + lr * Math.sin(mid) + 3, labelText(labels, value, String(data.categories[i] ?? ''), value / total, data.series[0]?.name), 'text-anchor="middle"'));
      }
      angle = end;
    });
    if (hole > 0) parts.push(`<circle cx="${cx}" cy="${cy}" r="${hole}" fill="#fff"/>`);
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${parts.join('')}</svg>`;
  }

  if (kind === 'radar') {
    // 방사형 — 항목마다 바큇살, 값 눈금은 동심 다각형, 계열은 닫힌 선
    const n = Math.max(1, data.categories.length, ...data.series.map((s) => s.values.length));
    const all = data.series.flatMap((s) => s.values);
    const step = niceStep(Math.max(0, ...all) - Math.min(0, ...all) || 1);
    const lo = look?.valMin ?? Math.min(0, Math.floor(Math.min(...all, 0) / step) * step);
    const hi = look?.valMax ?? (Math.ceil(Math.max(...all, 0) / step) * step || step);
    const cx = (left + right) / 2;
    const cy = (top + bottom) / 2 + 4;
    const R = Math.max(10, Math.min((right - left) / 2 - 30, (bottom - top) / 2 - 16));
    const at = (c: number, v: number) => {
      const a = -Math.PI / 2 + (2 * Math.PI * c) / n;
      const r = (Math.max(lo, Math.min(hi, v)) - lo) / (hi - lo || 1) * R;
      return [cx + r * Math.cos(a), cy + r * Math.sin(a)] as const;
    };
    for (let v = lo + step; v <= hi + step / 2; v += step) {
      parts.push(`<path d="M${Array.from({ length: n }, (_, c) => at(c, v).join(' ')).join('L')}Z" fill="none" stroke="#d9d9d9"/>`);
      parts.push(text(cx + 3, at(0, v)[1] + 3, formatTick(v, look?.valFmt)));
    }
    for (let c = 0; c < n; c += 1) {
      const [x, y] = at(c, hi);
      parts.push(`<line x1="${cx}" y1="${cy}" x2="${x}" y2="${y}" stroke="#d9d9d9"/>`);
      const [lx, ly] = at(c, hi + (hi - lo) * 0.12);
      parts.push(text(lx, ly + 3, String(data.categories[c] ?? '').slice(0, 10), 'text-anchor="middle"'));
    }
    data.series.forEach((ser, i) => {
      const color = colorOf(i);
      const pts = Array.from({ length: n }, (_, c) => at(c, ser.values[c] ?? 0));
      parts.push(`<path d="M${pts.map((p) => p.join(' ')).join('L')}Z" fill="none" stroke="${color}" stroke-width="2"/>`);
      const symbol = ser.look?.marker;
      if (symbol && symbol !== 'none') for (const [x, y] of pts) parts.push(marker(symbol, x, y, color));
    });
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${parts.join('')}</svg>`;
  }

  const horizontal = kind === 'bar';
  const kindOf = (i: number): ChartKind => data.series[i]?.look?.kind ?? kind;
  const secondary = (i: number) => Boolean(data.series[i]?.look?.secondary) && !horizontal;
  const stacked = look?.grouping;
  const n = Math.max(1, data.categories.length, ...data.series.map((s) => s.values.length));
  // 누적: 같은 종류(막대·영역)의 주 축 계열을 항목마다 쌓는다. 100% 누적은 합을 100%로
  const stackable = (i: number) => Boolean(stacked) && !secondary(i) && kindOf(i) !== 'line' && kindOf(i) !== 'scatter';
  const totals = Array.from({ length: n }, (_, c) => data.series.reduce((sum, s, i) => (stackable(i) ? sum + Math.abs(s.values[c] ?? 0) : sum), 0));
  const shown = (i: number, c: number) => {
    const v = data.series[i].values[c] ?? 0;
    return stacked === 'percentStacked' && stackable(i) ? (totals[c] ? v / totals[c] : 0) : v;
  };
  const base = (i: number, c: number) => {
    if (!stackable(i)) return 0;
    let b = 0;
    for (let k = 0; k < i; k += 1) if (stackable(k) && kindOf(k) === kindOf(i)) b += shown(k, c);
    return b;
  };
  const range = (axis2: boolean) => {
    const vals: number[] = [0];
    data.series.forEach((s, i) => {
      if (secondary(i) !== axis2) return;
      for (let c = 0; c < n; c += 1) {
        const b = base(i, c);
        vals.push(b + shown(i, c), b);
      }
    });
    const maxV = Math.max(...vals);
    const minV = Math.min(...vals);
    if (stacked === 'percentStacked' && !axis2) return { min: Math.min(0, minV), max: 1, step: 0.2 };
    // 파일이 정한 값 축 최소·최대(c:min·c:max)가 있으면 그대로 — 0이 아닌 축 시작
    const fixedMin = axis2 ? undefined : look?.valMin;
    const fixedMax = axis2 ? undefined : look?.valMax;
    const lo = fixedMin ?? minV;
    const hi = fixedMax ?? maxV;
    const step = niceStep(hi - lo || 1);
    return { min: fixedMin ?? Math.floor(lo / step) * step, max: fixedMax ?? (Math.ceil(hi / step) * step || step), step };
  };
  const ax1 = range(false);
  const hasAx2 = data.series.some((_s, i) => secondary(i));
  const ax2 = hasAx2 ? range(true) : null;
  const fmt1 = stacked === 'percentStacked' ? '0%' : look?.valFmt;
  const tickW = (ax: { min: number; max: number; step: number }, fmt?: string) => {
    let w = 0;
    // 눈금 글자 폭 어림(숫자·기호 0.62em, 한글 1em, 글자 크기 10px) + 축 여백 — 글자가 차트 왼쪽 끝에서 잘리지 않게
    for (let v = ax.min; v <= ax.max + ax.step / 2; v += ax.step) {
      let tw = 0;
      for (const ch of formatTick(v, fmt)) tw += (ch.codePointAt(0) ?? 0) >= 0x2e80 ? 10 : 6.2;
      w = Math.max(w, tw);
    }
    return w + 16;
  };
  if (look?.valTitle) {
    if (horizontal) bottom -= 16;
    else left += 18;
  }
  if (look?.catTitle) {
    if (horizontal) left += 18;
    else bottom -= 16;
  }
  left += horizontal ? 64 : tickW(ax1, fmt1) + 4;
  if (ax2) right -= tickW(ax2, look?.val2Fmt) + (look?.val2Title ? 18 : 0);
  bottom -= horizontal ? 18 : 20;
  const plotW = Math.max(10, right - left);
  const plotH = Math.max(10, bottom - top);
  const scale = (ax: { min: number; max: number }, v: number) => (v - ax.min) / (ax.max - ax.min || 1);
  const gridlines = look ? Boolean(look.gridlines) : true;
  for (let v = ax1.min; v <= ax1.max + ax1.step / 2; v += ax1.step) {
    const label = formatTick(v, fmt1);
    if (horizontal) {
      const x = left + scale(ax1, v) * plotW;
      if (gridlines) parts.push(`<line x1="${x}" y1="${top}" x2="${x}" y2="${bottom}" stroke="#d9d9d9"/>`);
      parts.push(text(x, bottom + 14, label, 'text-anchor="middle"'));
    } else {
      const y = bottom - scale(ax1, v) * plotH;
      if (gridlines) parts.push(`<line x1="${left}" y1="${y}" x2="${right}" y2="${y}" stroke="#d9d9d9"/>`);
      parts.push(text(left - 6, y + 3.5, label, 'text-anchor="end"'));
    }
  }
  if (ax2) {
    for (let v = ax2.min; v <= ax2.max + ax2.step / 2; v += ax2.step) {
      parts.push(text(right + 6, bottom - scale(ax2, v) * plotH + 3.5, formatTick(v, look?.val2Fmt)));
    }
  }
  if (look?.valTitle) {
    parts.push(
      horizontal
        ? text((left + right) / 2, height - 6 - (legendPos === 'b' ? 20 : 0), look.valTitle, 'text-anchor="middle"')
        : text(12, (top + bottom) / 2, look.valTitle, `text-anchor="middle" transform="rotate(-90 12 ${(top + bottom) / 2})"`),
    );
  }
  if (ax2 && look?.val2Title) parts.push(text(width - 8, (top + bottom) / 2, look.val2Title, `text-anchor="middle" transform="rotate(90 ${width - 8} ${(top + bottom) / 2})"`));
  if (look?.catTitle) {
    parts.push(
      horizontal
        ? text(12, (top + bottom) / 2, look.catTitle, `text-anchor="middle" transform="rotate(-90 12 ${(top + bottom) / 2})"`)
        : text((left + right) / 2, bottom + 34, look.catTitle, 'text-anchor="middle"'),
    );
  }
  const band = (horizontal ? plotH : plotW) / n;
  const scatter = kind === 'scatter' || kind === 'bubble';
  const xs = data.categories.map((c) => Number(c));
  const xNumeric = scatter && xs.length > 0 && xs.every((x) => Number.isFinite(x));
  const xMin = xNumeric ? Math.min(...xs) : 0;
  const xMax = xNumeric ? Math.max(...xs) : 1;
  const xAt = (c: number) => (xNumeric ? left + ((xs[c] - xMin) / (xMax - xMin || 1)) * plotW : left + band * (c + 0.5));
  data.categories.forEach((label, i) => {
    if (i >= n) return;
    const t = String(label).slice(0, 10);
    if (horizontal) parts.push(text(left - 6, top + band * (i + 0.5) + 3.5, t, 'text-anchor="end"'));
    else if (!xNumeric) parts.push(text(left + band * (i + 0.5), bottom + 16, t, 'text-anchor="middle"'));
  });
  if (xNumeric) {
    const step = niceStep(xMax - xMin || 1);
    for (let v = Math.ceil(xMin / step) * step; v <= xMax + step / 2; v += step) parts.push(text(left + ((v - xMin) / (xMax - xMin || 1)) * plotW, bottom + 16, formatTick(v), 'text-anchor="middle"'));
  }
  // 막대(세로·가로) — 묶은 막대는 계열마다 나란히, 누적은 쌓는다
  const barSeries = data.series.map((_s, i) => i).filter((i) => kindOf(i) === 'column' || kindOf(i) === 'bar');
  const slots = stacked ? 1 : Math.max(1, barSeries.length);
  const groupW = band * 0.62;
  const barW = groupW / slots;
  const labelsOf = (i: number) => data.series[i].look?.labels ?? look?.labels;
  barSeries.forEach((i, slot) => {
    const ax = secondary(i) && ax2 ? ax2 : ax1;
    const color = colorOf(i);
    for (let c = 0; c < n; c += 1) {
      const b0 = base(i, c);
      const b1 = b0 + shown(i, c);
      const a = scale(ax, Math.min(b0, b1));
      const b = scale(ax, Math.max(b0, b1));
      const offset = band * c + (band - groupW) / 2 + barW * (stacked ? 0 : slot);
      if (horizontal) parts.push(`<rect x="${left + a * plotW}" y="${top + offset}" width="${Math.max(0, (b - a) * plotW)}" height="${barW * 0.92}" fill="${color}"/>`);
      else parts.push(`<rect x="${left + offset}" y="${bottom - b * plotH}" width="${barW * 0.92}" height="${Math.max(0, (b - a) * plotH)}" fill="${color}"/>`);
      const labels = labelsOf(i);
      if (labels) {
        const v = data.series[i].values[c] ?? 0;
        const s = labelText(labels, v, String(data.categories[c] ?? ''), totals[c] ? v / totals[c] : undefined);
        if (horizontal) parts.push(text(left + b * plotW + 3, top + offset + barW * 0.6, s));
        else parts.push(text(left + offset + barW * 0.46, bottom - b * plotH - 3, s, 'text-anchor="middle"'));
      }
    }
  });
  // 영역·꺾은선·분산형
  data.series.forEach((ser, i) => {
    const k = kindOf(i);
    if (k === 'bubble') {
      // 거품형 — 거품 크기(c:bubbleSize)를 넓이에 비례하게, 가장 큰 거품이 그림 높이의 1/8
      const sizes = ser.sizes ?? [];
      const maxSize = Math.max(1e-9, ...data.series.flatMap((s) => s.sizes ?? []).map(Math.abs));
      const ax = ax1;
      const color = colorOf(i);
      for (let c = 0; c < n; c += 1) {
        const r = Math.max(2, Math.sqrt(Math.abs(sizes[c] ?? 0) / maxSize) * (plotH / 8));
        parts.push(`<circle cx="${xAt(c)}" cy="${bottom - scale(ax, ser.values[c] ?? 0) * plotH}" r="${r}" fill="${color}" fill-opacity="0.75" stroke="#fff"/>`);
      }
      return;
    }
    if (k !== 'line' && k !== 'area' && k !== 'scatter') return;
    const ax = secondary(i) && ax2 ? ax2 : ax1;
    const color = colorOf(i);
    const pts = Array.from({ length: n }, (_, c) => [xAt(c), bottom - scale(ax, base(i, c) + shown(i, c)) * plotH] as const);
    const smooth = Boolean(ser.look?.smooth);
    if (k === 'area') {
      const floor = Array.from({ length: n }, (_, c) => [xAt(n - 1 - c), bottom - scale(ax, base(i, n - 1 - c)) * plotH] as const);
      parts.push(`<path d="${linePath(pts, false)}L${floor.map(([x, y]) => `${x} ${y}`).join('L')}Z" fill="${color}" fill-opacity="0.85"/>`);
    } else if (k === 'line' || smooth) {
      parts.push(`<path d="${linePath(pts, smooth)}" fill="none" stroke="${color}" stroke-width="2.25" stroke-linejoin="round" stroke-linecap="round"/>`);
    }
    const symbol = ser.look?.marker ?? (k === 'scatter' ? 'circle' : look ? 'none' : 'none');
    if (symbol !== 'none') for (const [x, y] of pts) parts.push(marker(symbol, x, y, color));
    const labels = labelsOf(i);
    if (labels) pts.forEach(([x, y], c) => parts.push(text(x, y - 6, labelText(labels, ser.values[c] ?? 0, String(data.categories[c] ?? '')), 'text-anchor="middle"')));
  });
  const zero1 = scale(ax1, 0);
  parts.push(
    horizontal
      ? `<line x1="${left + zero1 * plotW}" y1="${top}" x2="${left + zero1 * plotW}" y2="${bottom}" stroke="#bfbfbf"/>`
      : `<line x1="${left}" y1="${bottom - zero1 * plotH}" x2="${right}" y2="${bottom - zero1 * plotH}" stroke="#bfbfbf"/>`,
  );
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
    return { name, values: values.map((v) => (Number.isFinite(Number(v)) ? Number(v) : 0)), ...(ser.look ? { look: ser.look } : {}), ...(ser.sizes ? { sizes: ser.sizes } : {}) };
  });
  const count = Math.max(categories.length, ...series.map((s) => s.values.length));
  const cats = categories.length > 0 ? categories : Array.from({ length: count }, (_, i) => String(i + 1));
  return { title: chart.title || '', categories: cats, series, ...(chart.look ? { look: chart.look } : {}) };
}
