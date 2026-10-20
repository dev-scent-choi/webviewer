/**
 * PDF 쪽 목록 보기 — 연속 스크롤, 배율·보기 회전, 현재 쪽 추적, 레이어, 찾기 강조.
 *
 * 쪽마다 자리(div.pdf-page)는 처음부터 다 만들어 스크롤 길이를 잡되, 그림(canvas)과 글자 층은 보이는 쪽과 그 앞뒤 한 화면
 * (IntersectionObserver rootMargin)만 그린다 — 수백 쪽 문서도 처음 여는 데 시간이 들지 않고, 멀어진 쪽의 canvas는 비워
 * 메모리를 돌려준다. 쪽 크기는 처음엔 1쪽 크기로 잡아 두고 그 쪽을 실제로 읽을 때 맞춘다(pdf.js 기본 뷰어와 같은 방식).
 */
import { pdfjs, type OptionalContentConfig, type PDFDocumentProxy } from './pdfDocument.js';
import type { FindMatch } from './findText.js';

type PDFPageProxy = Awaited<ReturnType<PDFDocumentProxy['getPage']>>;
type RenderTask = ReturnType<PDFPageProxy['render']>;

interface Slot {
  readonly page: number;
  readonly el: HTMLDivElement;
  canvas: HTMLCanvasElement;
  readonly textLayer: HTMLDivElement;
  readonly hitLayer: HTMLDivElement;
  /** 배율 1·회전 0에서의 크기(pt). 아직 그 쪽을 읽지 않았으면 1쪽 크기 */
  baseW: number;
  baseH: number;
  /** 쪽 자체의 /Rotate */
  pageRotate: number;
  sized: boolean;
  /** 지금 그려진 상태의 열쇠(배율·회전·레이어) — 같으면 다시 그리지 않는다 */
  renderedKey: string | null;
  renderTask: RenderTask | null;
  rendering: Promise<void> | null;
  textDivs: HTMLElement[];
  textStrs: string[];
}

const ZOOM_STEPS = [0.25, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4];
/** canvas 한 장의 최대 화소 수 — 넘으면 해상도를 낮춰 그리고 CSS로 늘린다(큰 쪽을 4배로 볼 때 메모리 폭주 방지) */
const MAX_CANVAS_PIXELS = 16_777_216;
const PAGE_GAP = 16;

export type ScaleMode = number | 'fit-width' | 'fit-page';

export class PageView {
  private readonly slots: Slot[] = [];
  private scaleValue = 1;
  private viewRotation = 0;
  private current = 1;
  private readonly pageListeners: Array<(page: number) => void> = [];
  private readonly scaleListeners: Array<(scale: number) => void> = [];
  private observer: IntersectionObserver | null = null;
  private readonly visible = new Set<number>();
  private ocConfig: OptionalContentConfig | null = null;
  private ocVersion = 0;
  private matchesByPage = new Map<number, FindMatch[]>();
  private currentMatch: FindMatch | null = null;
  private readonly textCache = new Map<number, Promise<string[]>>();
  private renderQueue: Promise<void> = Promise.resolve();
  private scrollRaf = 0;

  constructor(
    private readonly host: HTMLElement,
    private readonly scroller: HTMLElement,
    private readonly doc: PDFDocumentProxy,
  ) {}

  get pageCount(): number {
    return this.doc.numPages;
  }
  get scale(): number {
    return this.scaleValue;
  }
  get currentPage(): number {
    return this.current;
  }
  get layers(): OptionalContentConfig | null {
    return this.ocConfig;
  }

  onPageChange(cb: (page: number) => void): void {
    this.pageListeners.push(cb);
  }
  onScaleChange(cb: (scale: number) => void): void {
    this.scaleListeners.push(cb);
  }

  async init(): Promise<void> {
    const first = await this.doc.getPage(1);
    const vp = first.getViewport({ scale: 1, rotation: 0 });
    const userUnit = (first as unknown as { userUnit?: number }).userUnit ?? 1;
    this.ocConfig = await this.doc.getOptionalContentConfig();
    for (let page = 1; page <= this.doc.numPages; page += 1) {
      const el = document.createElement('div');
      el.className = 'pdf-page';
      el.dataset.page = String(page);
      const canvas = document.createElement('canvas');
      const textLayer = document.createElement('div');
      textLayer.className = 'pdf-text-layer';
      const hitLayer = document.createElement('div');
      hitLayer.className = 'pdf-hit-layer';
      const label = document.createElement('span');
      label.className = 'pdf-page-label';
      label.textContent = String(page);
      el.append(canvas, textLayer, hitLayer, label);
      this.host.append(el);
      this.slots.push({
        page,
        el,
        canvas,
        textLayer,
        hitLayer,
        baseW: vp.width * userUnit,
        baseH: vp.height * userUnit,
        pageRotate: page === 1 ? first.rotate : 0,
        sized: page === 1,
        renderedKey: null,
        renderTask: null,
        rendering: null,
        textDivs: [],
        textStrs: [],
      });
    }
    this.layoutAll();
    this.observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          const page = Number((entry.target as HTMLElement).dataset.page);
          if (entry.isIntersecting) this.visible.add(page);
          else {
            this.visible.delete(page);
            this.release(this.slots[page - 1]);
          }
        }
        this.renderVisible();
      },
      { root: this.scroller, rootMargin: '100% 0px' },
    );
    for (const slot of this.slots) this.observer.observe(slot.el);
    this.scroller.addEventListener('scroll', () => {
      if (this.scrollRaf) return;
      this.scrollRaf = requestAnimationFrame(() => {
        this.scrollRaf = 0;
        this.updateCurrentPage();
      });
    });
  }

  // ---------- 크기·배율·회전 ----------
  private rotationOf(slot: Slot): number {
    return (slot.pageRotate + this.viewRotation) % 360;
  }
  private cssSize(slot: Slot): { w: number; h: number } {
    const turned = this.rotationOf(slot) % 180 !== 0;
    const w = (turned ? slot.baseH : slot.baseW) * this.scaleValue;
    const h = (turned ? slot.baseW : slot.baseH) * this.scaleValue;
    return { w, h };
  }
  private layout(slot: Slot): void {
    const { w, h } = this.cssSize(slot);
    slot.el.style.width = `${Math.floor(w)}px`;
    slot.el.style.height = `${Math.floor(h)}px`;
    slot.el.style.setProperty('--scale-factor', String(this.scaleValue));
  }
  private layoutAll(): void {
    for (const slot of this.slots) this.layout(slot);
  }

  /** 지금 쪽의 크기 기준으로 맞춤 배율을 계산한다 */
  private fitScale(mode: 'fit-width' | 'fit-page'): number {
    const slot = this.slots[this.current - 1] ?? this.slots[0];
    if (!slot) return 1;
    const turned = this.rotationOf(slot) % 180 !== 0;
    const w = turned ? slot.baseH : slot.baseW;
    const h = turned ? slot.baseW : slot.baseH;
    const pad = getComputedStyle(this.host);
    const availW = this.scroller.clientWidth - parseFloat(pad.paddingLeft) - parseFloat(pad.paddingRight) - 4;
    const byWidth = availW / w;
    if (mode === 'fit-width') return Math.max(0.1, byWidth);
    const availH = this.scroller.clientHeight - parseFloat(pad.paddingTop) - PAGE_GAP;
    return Math.max(0.1, Math.min(byWidth, availH / h));
  }

  setScale(mode: ScaleMode): void {
    const next = typeof mode === 'number' ? mode : this.fitScale(mode);
    const clamped = Math.min(4, Math.max(0.1, next));
    if (Math.abs(clamped - this.scaleValue) < 0.001) return;
    // 지금 보던 쪽의 같은 비율 위치를 배율이 바뀐 뒤에도 유지한다
    const anchor = this.slots[this.current - 1];
    const ratio = anchor ? (this.scroller.scrollTop - anchor.el.offsetTop) / Math.max(1, anchor.el.offsetHeight) : 0;
    this.scaleValue = clamped;
    this.layoutAll();
    if (anchor) this.scroller.scrollTop = anchor.el.offsetTop + ratio * anchor.el.offsetHeight;
    this.invalidateAll();
    for (const cb of this.scaleListeners) cb(this.scaleValue);
  }
  zoomIn(): void {
    this.setScale(ZOOM_STEPS.find((s) => s > this.scaleValue + 0.001) ?? 4);
  }
  zoomOut(): void {
    this.setScale([...ZOOM_STEPS].reverse().find((s) => s < this.scaleValue - 0.001) ?? 0.25);
  }
  rotate(delta: 90 | -90): void {
    const page = this.current;
    this.viewRotation = (this.viewRotation + delta + 360) % 360;
    this.layoutAll();
    this.invalidateAll();
    this.goTo(page);
  }

  // ---------- 쪽 이동 ----------
  goTo(page: number): void {
    const slot = this.slots[Math.min(Math.max(1, Math.round(page)), this.slots.length) - 1];
    if (!slot) return;
    this.scroller.scrollTop = slot.el.offsetTop - 12;
    this.setCurrent(slot.page);
  }
  private setCurrent(page: number): void {
    if (page === this.current) return;
    this.current = page;
    for (const cb of this.pageListeners) cb(page);
  }
  private updateCurrentPage(): void {
    const mid = this.scroller.scrollTop + this.scroller.clientHeight * 0.35;
    // 쪽 자리는 위에서 아래로 쌓이므로 이분 탐색
    let lo = 0;
    let hi = this.slots.length - 1;
    while (lo < hi) {
      const m = (lo + hi + 1) >> 1;
      if (this.slots[m].el.offsetTop <= mid) lo = m;
      else hi = m - 1;
    }
    if (this.slots[lo]) this.setCurrent(this.slots[lo].page);
  }

  // ---------- 그리기 ----------
  private keyOf(): string {
    return `${this.scaleValue.toFixed(4)}|${this.viewRotation}|${this.ocVersion}|${window.devicePixelRatio || 1}`;
  }
  private invalidateAll(): void {
    for (const slot of this.slots) {
      if (slot.renderedKey !== null) slot.renderedKey = 'stale';
    }
    this.renderVisible();
  }
  private release(slot: Slot | undefined): void {
    if (!slot || slot.renderedKey === null) return;
    slot.renderTask?.cancel();
    slot.renderTask = null;
    slot.canvas.width = 0;
    slot.canvas.height = 0;
    slot.textLayer.replaceChildren();
    slot.hitLayer.replaceChildren();
    slot.textDivs = [];
    slot.textStrs = [];
    slot.renderedKey = null;
  }
  private renderVisible(): void {
    const pages = [...this.visible].sort((a, b) => Math.abs(a - this.current) - Math.abs(b - this.current));
    for (const page of pages) {
      const slot = this.slots[page - 1];
      if (slot && slot.renderedKey !== this.keyOf()) void this.renderSlot(slot);
    }
  }

  /** 한 쪽을 그린다(같은 쪽을 동시에 두 번 그리지 않는다). 끝나면 그 쪽의 찾기 강조도 다시 그린다. */
  renderSlot(slot: Slot): Promise<void> {
    if (slot.rendering) return slot.rendering;
    const key = this.keyOf();
    slot.rendering = (async () => {
      try {
        const page = await this.doc.getPage(slot.page);
        if (!slot.sized) {
          const vp1 = page.getViewport({ scale: 1, rotation: 0 });
          const userUnit = (page as unknown as { userUnit?: number }).userUnit ?? 1;
          slot.baseW = vp1.width * userUnit;
          slot.baseH = vp1.height * userUnit;
          slot.pageRotate = page.rotate;
          slot.sized = true;
          this.layout(slot);
        }
        const viewport = page.getViewport({ scale: this.scaleValue, rotation: this.rotationOf(slot) });
        let ratio = window.devicePixelRatio || 1;
        if (viewport.width * viewport.height * ratio * ratio > MAX_CANVAS_PIXELS) {
          ratio = Math.sqrt(MAX_CANVAS_PIXELS / (viewport.width * viewport.height));
        }
        const canvas = document.createElement('canvas');
        canvas.width = Math.floor(viewport.width * ratio);
        canvas.height = Math.floor(viewport.height * ratio);
        const ctx = canvas.getContext('2d', { alpha: false });
        if (!ctx) return;
        slot.renderTask?.cancel();
        const task = page.render({
          canvasContext: ctx,
          viewport,
          transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined,
          annotationMode: pdfjs.AnnotationMode.ENABLE_FORMS,
          optionalContentConfigPromise: this.ocConfig ? Promise.resolve(this.ocConfig) : undefined,
        });
        slot.renderTask = task;
        await task.promise;
        if (slot.renderTask !== task) return;
        slot.renderTask = null;
        // 다 그린 canvas로 한 번에 바꿔 끼운다(그리는 동안 빈 쪽이 보이지 않게)
        slot.canvas.replaceWith(canvas);
        slot.canvas = canvas;

        const textDivs: HTMLElement[] = [];
        const textStrs: string[] = [];
        slot.textLayer.replaceChildren();
        const textTask = pdfjs.renderTextLayer({
          textContentSource: page.streamTextContent(),
          container: slot.textLayer,
          viewport,
          textDivs,
          textContentItemsStr: textStrs,
        });
        await textTask.promise;
        slot.textDivs = textDivs;
        slot.textStrs = textStrs;
        slot.renderedKey = key;
        this.drawHits(slot);
      } catch (error) {
        if ((error as { name?: string })?.name !== 'RenderingCancelledException') console.error(`쪽 ${slot.page} 그리기 실패:`, error);
      } finally {
        slot.rendering = null;
        // 그리는 동안 배율 등이 바뀌었으면 새 열쇠로 다시
        if (this.visible.has(slot.page) && slot.renderedKey !== this.keyOf()) queueMicrotask(() => void this.renderSlot(slot));
      }
    })();
    return slot.rendering;
  }

  /** 첫 쪽이 다 그려질 때까지 기다린다(docview:ready 신호용) */
  async firstPaint(): Promise<void> {
    const slot = this.slots[0];
    if (slot) await this.renderSlot(slot);
  }

  // ---------- 레이어 ----------
  setLayerVisibility(id: string, visible: boolean): void {
    if (!this.ocConfig) return;
    this.ocConfig.setVisibility(id, visible);
    this.ocVersion += 1;
    this.invalidateAll();
  }

  // ---------- 찾기 ----------
  /** 찾기용 쪽 글자 조각 — 글자 층과 같은 순서(pdf.js가 조각마다 span 하나) */
  textItems(page: number): Promise<string[]> {
    let cached = this.textCache.get(page);
    if (!cached) {
      cached = this.doc
        .getPage(page)
        .then((p) => p.getTextContent())
        .then((content) => content.items.map((item) => ('str' in item ? item.str : '')));
      this.textCache.set(page, cached);
    }
    return cached;
  }

  highlight(matches: readonly FindMatch[], current: FindMatch | null): void {
    this.matchesByPage = new Map();
    for (const m of matches) {
      const list = this.matchesByPage.get(m.page) ?? [];
      list.push(m);
      this.matchesByPage.set(m.page, list);
    }
    this.currentMatch = current;
    for (const slot of this.slots) if (slot.renderedKey !== null) this.drawHits(slot);
  }

  private drawHits(slot: Slot): void {
    slot.hitLayer.replaceChildren();
    const matches = this.matchesByPage.get(slot.page);
    if (!matches?.length || slot.textDivs.length === 0) return;
    const base = slot.el.getBoundingClientRect();
    for (const m of matches) {
      for (const rect of this.rectsOf(slot, m)) {
        const hit = document.createElement('div');
        hit.className = 'pdf-find-hit';
        if (m === this.currentMatch) hit.classList.add('current');
        hit.style.left = `${rect.left - base.left}px`;
        hit.style.top = `${rect.top - base.top}px`;
        hit.style.width = `${rect.width}px`;
        hit.style.height = `${rect.height}px`;
        slot.hitLayer.append(hit);
      }
    }
  }

  /** 일치 구간이 덮는 글자의 화면 사각형들(조각마다 하나) */
  private rectsOf(slot: Slot, m: FindMatch): DOMRect[] {
    const out: DOMRect[] = [];
    for (let i = m.itemStart; i <= m.itemEnd; i += 1) {
      const div = slot.textDivs[i];
      const node = div?.firstChild;
      if (!div || !node || node.nodeType !== Node.TEXT_NODE) continue;
      const len = node.textContent?.length ?? 0;
      const start = i === m.itemStart ? Math.min(m.offsetStart, len) : 0;
      const end = i === m.itemEnd ? Math.min(m.offsetEnd, len) : len;
      if (end <= start) continue;
      const range = document.createRange();
      range.setStart(node, start);
      range.setEnd(node, end);
      out.push(...Array.from(range.getClientRects()));
    }
    return out;
  }

  /** 일치한 곳이 화면 가운데쯤 오도록 — 그 쪽이 아직 안 그려졌으면 그린 뒤 */
  async scrollToMatch(m: FindMatch): Promise<void> {
    const slot = this.slots[m.page - 1];
    if (!slot) return;
    if (slot.renderedKey !== this.keyOf()) {
      this.scroller.scrollTop = slot.el.offsetTop - 12;
      this.visible.add(slot.page);
      await this.renderSlot(slot);
    }
    const rects = this.rectsOf(slot, m);
    if (rects.length === 0) return;
    const box = this.scroller.getBoundingClientRect();
    const r = rects[0];
    const dy = r.top - box.top - this.scroller.clientHeight / 3;
    const dx = r.left - box.left - this.scroller.clientWidth / 2;
    this.scroller.scrollTop += dy;
    if (r.left < box.left || r.right > box.right) this.scroller.scrollLeft += dx;
    this.setCurrent(slot.page);
  }
}
