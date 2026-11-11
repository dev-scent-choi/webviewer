/**
 * 거리 측정·도면 모드(반전) — PDF 문서에서만(지금 보기 화면과 같다, viewer.js measure·applyInvert). 다른 형식은 흐리게 + 이유.
 */
import type { ShellCtx } from './ctx.js';
import { disabledReason, measureLabel, paperMm, toFrameRatio } from './logic.js';

const $ = <T extends HTMLElement = HTMLElement>(id: string): T => document.getElementById(id) as T;
const SVG_NS = 'http://www.w3.org/2000/svg';
const INVERT_STORAGE_KEY = 'docview-invert';

interface Point {
  readonly frame: HTMLElement;
  readonly fx: number;
  readonly fy: number;
}

export function setupMeasure(ctx: ShellCtx): void {
  const measureBtn = $<HTMLButtonElement>('measure-btn');
  const invertBtn = $<HTMLButtonElement>('invert-btn');
  const reason = disabledReason('measure', ctx.info.permissions, ctx.info.ext);
  for (const btn of [measureBtn, invertBtn]) {
    btn.disabled = reason !== null;
    if (reason) btn.title = reason;
  }
  if (reason) return;
  const pages = ctx.body.container;
  const hint = $('measure-hint');
  const state = {
    mode: false,
    calibrating: false,
    pending: null as (Point & { dot: HTMLElement }) | null,
    items: [] as Array<{ paperMm: number; label: HTMLElement; els: Element[] }>,
    factor: 1,
    unit: 'mm(용지)',
  };

  const layerOf = (frame: HTMLElement): SVGSVGElement => {
    let layer = frame.querySelector<SVGSVGElement>('.measure-layer');
    if (!layer) {
      layer = document.createElementNS(SVG_NS, 'svg');
      layer.setAttribute('class', 'measure-layer');
      layer.setAttribute('viewBox', '0 0 100 100');
      layer.setAttribute('preserveAspectRatio', 'none');
      frame.append(layer);
    }
    return layer;
  };
  const dot = (frame: HTMLElement, fx: number, fy: number): HTMLElement => {
    const el = document.createElement('div');
    el.className = 'measure-dot';
    el.style.left = `${fx * 100}%`;
    el.style.top = `${fy * 100}%`;
    frame.append(el);
    return el;
  };
  const clear = () => {
    for (const item of state.items) for (const el of item.els) el.remove();
    state.items = [];
    state.pending?.dot.remove();
    state.pending = null;
  };
  const add = (a: Point & { dot: HTMLElement }, b: Point) => {
    const mm = paperMm(a, b, ctx.body.pageSize(Number(a.frame.dataset.page)));
    const line = document.createElementNS(SVG_NS, 'line');
    line.setAttribute('x1', String(a.fx * 100));
    line.setAttribute('y1', String(a.fy * 100));
    line.setAttribute('x2', String(b.fx * 100));
    line.setAttribute('y2', String(b.fy * 100));
    layerOf(a.frame).append(line);
    const dotB = dot(a.frame, b.fx, b.fy);
    const label = document.createElement('div');
    label.className = 'measure-label';
    label.style.left = `${((a.fx + b.fx) / 2) * 100}%`;
    label.style.top = `${((a.fy + b.fy) / 2) * 100}%`;
    label.textContent = measureLabel(mm, state.factor, state.unit);
    a.frame.append(label);
    state.items.push({ paperMm: mm, label, els: [line, a.dot, dotB, label] });
    if (!state.calibrating) return;
    state.calibrating = false;
    const value = Number((window.prompt('방금 잰 구간의 실제 길이를 입력하세요 (예: 1500 → mm)', '') || '').replace(/[^\d.]/g, ''));
    if (value > 0) {
      state.factor = value / mm;
      state.unit = 'mm';
      for (const item of state.items) item.label.textContent = measureLabel(item.paperMm, state.factor, state.unit);
      hint.textContent = '보정 완료 — 실제 치수(mm)로 표시됩니다';
    } else {
      hint.textContent = '보정이 취소되었습니다';
    }
  };
  const setMode = (on: boolean) => {
    state.mode = on;
    state.calibrating = false;
    measureBtn.setAttribute('aria-pressed', String(on));
    $('measure-bar').hidden = !on;
    pages.classList.toggle('measuring', on);
    if (!on) clear();
    else hint.textContent = '두 점을 클릭해 거리를 측정하세요';
  };

  measureBtn.addEventListener('click', () => setMode(!state.mode));
  $('measure-close').addEventListener('click', () => setMode(false));
  $('measure-clear').addEventListener('click', clear);
  $('measure-calibrate').addEventListener('click', () => {
    state.calibrating = true;
    hint.textContent = '실제 길이를 아는 구간(치수선 등)의 양 끝을 클릭하세요';
  });
  pages.addEventListener('click', (event) => {
    if (!state.mode) return;
    const frame = (event.target as Element).closest<HTMLElement>('.page-frame');
    if (!frame) return;
    // 보기 회전을 되돌린 쪽 비율(표시와 같은 toFrameRatio)
    const point = { frame, ...toFrameRatio(frame.getBoundingClientRect(), frame.offsetWidth, frame.offsetHeight, ctx.body.rotateDeg, event.clientX, event.clientY) };
    if (!state.pending || state.pending.frame !== frame) {
      state.pending?.dot.remove();
      state.pending = { ...point, dot: dot(frame, point.fx, point.fy) };
    } else {
      add(state.pending, point);
      state.pending = null;
    }
  });
  ctx.onEscape(10, () => {
    if (!state.mode) return false;
    setMode(false);
    return true;
  });
  // 쪽 틀이 새로 만들어지면(pdf.js → PNG) 잰 것은 사라진다 — 상태도 비운다
  ctx.body.onSourceChange(() => setMode(false));

  const applyInvert = (on: boolean) => {
    pages.classList.toggle('pages-invert', on);
    invertBtn.setAttribute('aria-pressed', String(on));
    try {
      localStorage.setItem(INVERT_STORAGE_KEY, on ? '1' : '0');
    } catch {
      // 시크릿 모드 등에서 저장 실패해도 무시
    }
  };
  invertBtn.addEventListener('click', () => applyInvert(!pages.classList.contains('pages-invert')));
  try {
    if (localStorage.getItem(INVERT_STORAGE_KEY) === '1') applyInvert(true);
  } catch {
    // 저장소를 못 쓰면 기본값
  }
}
