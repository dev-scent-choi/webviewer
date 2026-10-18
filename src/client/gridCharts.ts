/**
 * 엑셀 편집기의 차트 — 시트 위에 SVG로 그린다. 파일에 있던 차트(/api/view/xlsx-charts, 저장 때 원본
 * 그대로 보존)와 삽입 탭으로 넣은 차트(공유 Y.Map GRID_CHARTS_MAP, 저장 때 차트 파트로 붙인다)를 함께
 * 그리며, 값은 항상 지금 칸 값으로 다시 계산한다. 넣은 차트는 끌어서 옮기고, 모서리로 크기를 바꾸고,
 * "차트 디자인" 탭에서 종류·제목을 바꾸거나 지울 수 있다.
 */
import type * as Y from 'yjs';
import { GRID_CHARTS_MAP, cellsOf, seriesFromRange, type CellRangeRef, type ChartKind, type GridChart } from '../xlsx/chartModel.js';

export interface GridChartHost {
  readonly ydoc: Y.Doc;
  readonly token: string;
  readonly canEdit: boolean;
  /** 차트 레이어를 넣을 스크롤 영역과 그 안의 표 */
  readonly scrollHost: HTMLElement;
  readonly table: HTMLElement;
  activeSheetName(): string;
  /** 지금 칸 표시 값(수식은 계산 결과) — 그 시트가 화면에 없으면 '' */
  valueText(sheetName: string, row: number, col: number): string;
  /** 표 왼쪽 위 기준 칸의 위치(px) */
  colLeft(col: number): number;
  rowTop(row: number): number;
  /** 위치(px) → 칸 */
  cellAt(x: number, y: number): { row: number; col: number };
  /** 지금 선택 범위(칸 하나면 r1===r2·c1===c2) */
  selection(): CellRangeRef;
  /** 시트의 행·열 수(현재 영역 찾기용) */
  extent(): { rows: number; cols: number };
  setChartTabVisible(visible: boolean): void;
}

const PALETTE = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#264478', '#9e480e'];
const KIND_LABELS: Readonly<Record<ChartKind, string>> = { column: '세로 막대형', bar: '가로 막대형', line: '꺾은선형', pie: '원형' };

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

/** 칸 하나만 선택했으면 그 칸을 둘러싼 비어 있지 않은 영역(Excel의 "현재 영역") */
export function currentRegion(start: { row: number; col: number }, isFilled: (r: number, c: number) => boolean, extent: { rows: number; cols: number }): CellRangeRef | null {
  if (!isFilled(start.row, start.col)) return null;
  let r1 = start.row;
  let r2 = start.row;
  let c1 = start.col;
  let c2 = start.col;
  let grew = true;
  const rowHas = (r: number) => {
    for (let c = Math.max(0, c1 - 1); c <= Math.min(extent.cols - 1, c2 + 1); c += 1) if (isFilled(r, c)) return true;
    return false;
  };
  const colHas = (c: number) => {
    for (let r = Math.max(0, r1 - 1); r <= Math.min(extent.rows - 1, r2 + 1); r += 1) if (isFilled(r, c)) return true;
    return false;
  };
  while (grew) {
    grew = false;
    if (r1 > 0 && rowHas(r1 - 1)) (r1 -= 1), (grew = true);
    if (r2 < extent.rows - 1 && rowHas(r2 + 1)) (r2 += 1), (grew = true);
    if (c1 > 0 && colHas(c1 - 1)) (c1 -= 1), (grew = true);
    if (c2 < extent.cols - 1 && colHas(c2 + 1)) (c2 += 1), (grew = true);
  }
  return { r1, c1, r2, c2 };
}

function newId(): string {
  return typeof crypto !== 'undefined' && 'randomUUID' in crypto ? crypto.randomUUID() : `ch${Date.now()}${Math.random().toString(36).slice(2)}`;
}

export function setupGridCharts(host: GridChartHost): { render(): void; insert(kind: ChartKind): void } {
  const map = host.ydoc.getMap<GridChart>(GRID_CHARTS_MAP);
  let fileCharts: GridChart[] = [];
  let selectedId: string | null = null;
  const layer = document.createElement('div');
  layer.className = 'grid-chart-layer';
  host.scrollHost.append(layer);

  void fetch(`/api/view/xlsx-charts?token=${encodeURIComponent(host.token)}`)
    .then((res) => res.json() as Promise<{ data?: { charts?: GridChart[] } }>)
    .then((json) => {
      fileCharts = json.data?.charts ?? [];
      render();
    })
    .catch(() => undefined);

  const allCharts = (): GridChart[] => [...fileCharts, ...map.values()];

  function dataOf(chart: GridChart): ChartData {
    const value = (ref: CellRangeRef, r: number, c: number) => host.valueText(ref.sheetName ?? chart.sheetName, r, c);
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
    return { title: chart.title || (chart.series.length === 1 && chart.source === 'file' ? '' : ''), categories: cats, series };
  }

  function boxOf(chart: GridChart): { left: number; top: number; width: number; height: number } {
    const left = host.colLeft(chart.from.col) + chart.from.dx;
    const top = host.rowTop(chart.from.row) + chart.from.dy;
    if (chart.size) return { left, top, width: chart.size.width, height: chart.size.height };
    const to = chart.to ?? { row: chart.from.row + 15, col: chart.from.col + 7, dx: 0, dy: 0 };
    return { left, top, width: Math.max(80, host.colLeft(to.col) + to.dx - left), height: Math.max(60, host.rowTop(to.row) + to.dy - top) };
  }

  function select(id: string | null): void {
    selectedId = id;
    const chart = id ? allCharts().find((c) => c.id === id) : undefined;
    host.setChartTabVisible(Boolean(chart && chart.source === 'new' && host.canEdit));
    syncDesignTab();
    for (const el of layer.querySelectorAll<HTMLElement>('.grid-chart')) el.classList.toggle('is-selected', el.dataset.chartId === id);
  }

  function syncDesignTab(): void {
    const chart = selectedId ? map.get(selectedId) : undefined;
    for (const btn of document.querySelectorAll<HTMLElement>('[data-chart-kind]')) {
      btn.setAttribute('aria-pressed', String(Boolean(chart) && btn.dataset.chartKind === chart?.kind && Boolean(btn.closest('[data-panel="chart"]'))));
    }
    const titleInput = document.getElementById('chart-title-input') as HTMLInputElement | null;
    if (titleInput && document.activeElement !== titleInput) titleInput.value = chart?.title ?? '';
  }

  function render(): void {
    const sheetName = host.activeSheetName();
    const charts = allCharts().filter((c) => c.sheetName === sheetName);
    layer.style.left = `${host.table.offsetLeft}px`;
    layer.style.top = `${host.table.offsetTop}px`;
    const keep = new Set<string>();
    for (const chart of charts) {
      keep.add(chart.id);
      let el = layer.querySelector<HTMLElement>(`[data-chart-id="${CSS.escape(chart.id)}"]`);
      if (!el) {
        el = document.createElement('div');
        el.className = 'grid-chart';
        el.dataset.chartId = chart.id;
        el.tabIndex = 0;
        bindInteractions(el);
        layer.append(el);
      }
      const box = boxOf(chart);
      el.style.left = `${box.left}px`;
      el.style.top = `${box.top}px`;
      el.style.width = `${box.width}px`;
      el.style.height = `${box.height}px`;
      el.classList.toggle('is-file', chart.source === 'file');
      el.classList.toggle('is-selected', chart.id === selectedId);
      el.title = chart.source === 'file' ? `${KIND_LABELS[chart.kind]} 차트(원본 파일 — 저장할 때 그대로 보존됩니다)` : `${KIND_LABELS[chart.kind]} 차트`;
      const svg = renderChartSvg(chart.kind, { ...dataOf(chart), title: chart.title }, Math.round(box.width), Math.round(box.height));
      const content = `${svg}${chart.source === 'new' && host.canEdit ? '<span class="grid-chart-resize" aria-hidden="true"></span>' : ''}`;
      if (el.dataset.rendered !== content) {
        el.innerHTML = content;
        el.dataset.rendered = content;
      }
    }
    for (const el of [...layer.querySelectorAll<HTMLElement>('.grid-chart')]) if (!keep.has(el.dataset.chartId ?? '')) el.remove();
    if (selectedId && !keep.has(selectedId)) select(null);
  }

  function bindInteractions(el: HTMLElement): void {
    el.addEventListener('pointerdown', (event) => {
      const id = el.dataset.chartId ?? '';
      select(id);
      el.focus();
      const chart = map.get(id);
      if (!chart || !host.canEdit) return;
      event.preventDefault();
      const resizing = (event.target as HTMLElement).classList.contains('grid-chart-resize');
      const startX = event.clientX;
      const startY = event.clientY;
      const box = boxOf(chart);
      const zoom = el.getBoundingClientRect().width / el.offsetWidth || 1;
      el.setPointerCapture(event.pointerId);
      const onMove = (e: PointerEvent) => {
        const dx = (e.clientX - startX) / zoom;
        const dy = (e.clientY - startY) / zoom;
        if (resizing) {
          el.style.width = `${Math.max(120, box.width + dx)}px`;
          el.style.height = `${Math.max(90, box.height + dy)}px`;
        } else {
          el.style.left = `${Math.max(0, box.left + dx)}px`;
          el.style.top = `${Math.max(0, box.top + dy)}px`;
        }
      };
      const onUp = (e: PointerEvent) => {
        el.releasePointerCapture(e.pointerId);
        el.removeEventListener('pointermove', onMove);
        el.removeEventListener('pointerup', onUp);
        const current = map.get(id);
        if (!current) return;
        if (resizing) {
          map.set(id, { ...current, size: { width: Math.round(el.offsetWidth), height: Math.round(el.offsetHeight) } });
        } else if (e.clientX !== startX || e.clientY !== startY) {
          const left = parseFloat(el.style.left);
          const top = parseFloat(el.style.top);
          const cell = host.cellAt(left, top);
          map.set(id, { ...current, from: { row: cell.row, col: cell.col, dx: Math.max(0, left - host.colLeft(cell.col)), dy: Math.max(0, top - host.rowTop(cell.row)) } });
        }
      };
      el.addEventListener('pointermove', onMove);
      el.addEventListener('pointerup', onUp);
    });
    el.addEventListener('keydown', (event) => {
      const id = el.dataset.chartId ?? '';
      if ((event.key === 'Delete' || event.key === 'Backspace') && map.has(id) && host.canEdit) {
        event.preventDefault();
        map.delete(id);
        select(null);
      } else if (event.key === 'Escape') {
        select(null);
        el.blur();
      }
    });
  }

  document.addEventListener('pointerdown', (event) => {
    const target = event.target as HTMLElement;
    if (!target.closest('.grid-chart') && !target.closest('#toolbar') && !target.closest('.dropdown-menu')) select(null);
  });

  function insert(kind: ChartKind): void {
    if (!host.canEdit) return;
    const sheetName = host.activeSheetName();
    const sel = host.selection();
    const single = sel.r1 === sel.r2 && sel.c1 === sel.c2;
    const range = single
      ? currentRegion({ row: sel.r1, col: sel.c1 }, (r, c) => host.valueText(sheetName, r, c).trim() !== '', host.extent())
      : sel;
    if (!range) {
      window.alert('차트로 만들 데이터를 선택하세요(머리글과 숫자가 있는 범위).');
      return;
    }
    const { categories, series } = seriesFromRange(range, (r, c) => host.valueText(sheetName, r, c));
    if (series.length === 0) {
      window.alert('차트로 그릴 숫자 열이 없습니다.');
      return;
    }
    const id = newId();
    const chart: GridChart = {
      id,
      sheetName,
      kind,
      title: series.length === 1 ? series[0].name.replace(/^계열\d+$/, '') : '',
      from: { row: range.r1, col: range.c2 + 2, dx: 0, dy: 0 },
      to: null,
      size: { width: 480, height: 288 },
      categories,
      series,
      source: 'new',
    };
    map.set(id, chart);
    render();
    select(id);
    layer.querySelector<HTMLElement>(`[data-chart-id="${CSS.escape(id)}"]`)?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  // 차트 디자인(상황별 탭)
  for (const btn of document.querySelectorAll<HTMLElement>('[data-panel="chart"] [data-chart-kind]')) {
    btn.addEventListener('click', () => {
      const chart = selectedId ? map.get(selectedId) : undefined;
      if (chart) map.set(chart.id, { ...chart, kind: btn.dataset.chartKind as ChartKind });
    });
  }
  const titleInput = document.getElementById('chart-title-input') as HTMLInputElement | null;
  titleInput?.addEventListener('input', () => {
    const chart = selectedId ? map.get(selectedId) : undefined;
    if (chart) map.set(chart.id, { ...chart, title: titleInput.value });
  });
  document.getElementById('btn-chart-delete')?.addEventListener('click', () => {
    if (selectedId && map.has(selectedId)) {
      map.delete(selectedId);
      select(null);
    }
  });

  map.observe(() => {
    render();
    syncDesignTab();
  });
  return { render, insert };
}
