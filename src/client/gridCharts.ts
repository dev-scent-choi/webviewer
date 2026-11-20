/**
 * 엑셀 편집기의 차트 — 시트 위에 SVG로 그린다. 파일에 있던 차트(/api/view/xlsx-charts, 저장 때 원본
 * 그대로 보존)와 삽입 탭으로 넣은 차트(공유 Y.Map GRID_CHARTS_MAP, 저장 때 차트 파트로 붙인다)를 함께
 * 그리며, 값은 항상 지금 칸 값으로 다시 계산한다. 넣은 차트는 끌어서 옮기고, 모서리로 크기를 바꾸고,
 * "차트 디자인" 탭에서 종류·제목을 바꾸거나 지울 수 있다.
 */
import type * as Y from 'yjs';
import { GRID_CHARTS_MAP, seriesFromRange, type CellRangeRef, type ChartKind, type GridChart } from '../xlsx/chartModel.js';
import { chartDataOf, renderChartSvg, type ChartData } from '../xlsx/chartSvg.js';
import { officeAlert } from './officeDialog.js';

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

const KIND_LABELS: Readonly<Record<ChartKind, string>> = { column: '세로 막대형', bar: '가로 막대형', line: '꺾은선형', pie: '원형', area: '영역형', doughnut: '도넛형', scatter: '분산형', radar: '방사형', bubble: '거품형' };

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

/** /api/view/xlsx-images 응답 항목(서버 drawingPassthrough.ts의 GridImage) */
interface FileImage {
  readonly id: string;
  readonly sheetName: string;
  readonly name: string;
  readonly from: GridChart['from'];
  readonly to: GridChart['to'];
  readonly size: GridChart['size'];
  readonly src: string;
}

export function setupGridCharts(host: GridChartHost): { render(): void; insert(kind: ChartKind): void } {
  const map = host.ydoc.getMap<GridChart>(GRID_CHARTS_MAP);
  let fileCharts: GridChart[] = [];
  let selectedId: string | null = null;
  const layer = document.createElement('div');
  layer.className = 'grid-chart-layer';
  host.scrollHost.append(layer);

  // 파일에 있던 그림(로고 등) — 읽기 전용으로 칸 위치에 얹는다(저장할 때는 원본 그리기 파트째 보존된다)
  let fileImages: FileImage[] = [];
  void fetch(`/api/view/xlsx-images?token=${encodeURIComponent(host.token)}`)
    .then((res) => res.json() as Promise<{ data?: { images?: FileImage[] } }>)
    .then((json) => {
      fileImages = json.data?.images ?? [];
      render();
    })
    .catch(() => undefined);

  void fetch(`/api/view/xlsx-charts?token=${encodeURIComponent(host.token)}`)
    .then((res) => res.json() as Promise<{ data?: { charts?: GridChart[] } }>)
    .then((json) => {
      fileCharts = json.data?.charts ?? [];
      render();
    })
    .catch(() => undefined);

  const allCharts = (): GridChart[] => [...fileCharts, ...map.values()];

  function dataOf(chart: GridChart): ChartData {
    return chartDataOf(chart, host.valueText);
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
      // SVG를 <img>(data: 주소)로 얹는다 — 공유 맵에서 온 차트 값이 어떤 글자든 문서 안에서 스크립트가 될 수 없다(보기 화면과 같은 방식)
      const resizable = chart.source === 'new' && host.canEdit;
      const key = `${svg}|${resizable ? 1 : 0}`;
      if (el.dataset.rendered !== key) {
        const img = document.createElement('img');
        img.className = 'grid-chart-img';
        img.alt = chart.title || KIND_LABELS[chart.kind] || '';
        img.draggable = false;
        img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
        const handle = document.createElement('span');
        handle.className = 'grid-chart-resize';
        handle.setAttribute('aria-hidden', 'true');
        el.replaceChildren(...(resizable ? [img, handle] : [img]));
        el.dataset.rendered = key;
      }
    }
    for (const el of [...layer.querySelectorAll<HTMLElement>('.grid-chart')]) if (!keep.has(el.dataset.chartId ?? '')) el.remove();
    if (selectedId && !keep.has(selectedId)) select(null);
    renderImages(sheetName);
  }

  function renderImages(sheetName: string): void {
    const keep = new Set<string>();
    for (const image of fileImages) {
      if (image.sheetName !== sheetName) continue;
      const left = host.colLeft(image.from.col) + image.from.dx;
      const top = host.rowTop(image.from.row) + image.from.dy;
      const box = image.size
        ? { left, top, width: image.size.width, height: image.size.height }
        : image.to
          ? { left, top, width: host.colLeft(image.to.col) + image.to.dx - left, height: host.rowTop(image.to.row) + image.to.dy - top }
          : null;
      if (!box || !(box.width > 0 && box.height > 0)) continue;
      keep.add(image.id);
      let img = layer.querySelector<HTMLImageElement>(`img[data-image-id="${CSS.escape(image.id)}"]`);
      if (!img) {
        img = document.createElement('img');
        img.className = 'grid-image';
        img.dataset.imageId = image.id;
        img.alt = image.name;
        img.draggable = false;
        img.src = image.src;
        layer.prepend(img);
      }
      img.style.left = `${box.left}px`;
      img.style.top = `${box.top}px`;
      img.style.width = `${box.width}px`;
      img.style.height = `${box.height}px`;
    }
    for (const img of [...layer.querySelectorAll<HTMLImageElement>('img.grid-image')]) if (!keep.has(img.dataset.imageId ?? '')) img.remove();
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
      void officeAlert('차트로 만들 데이터를 선택하세요(머리글과 숫자가 있는 범위).');
      return;
    }
    const { categories, series } = seriesFromRange(range, (r, c) => host.valueText(sheetName, r, c));
    if (series.length === 0) {
      void officeAlert('차트로 그릴 숫자 열이 없습니다.');
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
