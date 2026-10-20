/**
 * PDF 쪽 목록 보기 — 연속 스크롤, 배율·보기 회전, 현재 쪽 추적, 레이어, 찾기 강조.
 *
 * 쪽마다 자리(div.pdf-page)는 처음부터 다 만들어 스크롤 길이를 잡되, 그림(canvas)과 글자 층은 보이는 쪽과 그 앞뒤 한 화면
 * (IntersectionObserver rootMargin)만 그린다 — 수백 쪽 문서도 처음 여는 데 시간이 들지 않고, 멀어진 쪽의 canvas는 비워
 * 메모리를 돌려준다. 쪽 크기는 처음엔 1쪽 크기로 잡아 두고 그 쪽을 실제로 읽을 때 맞춘다(pdf.js 기본 뷰어와 같은 방식).
 *
 * 3단계(쪽 정리): 무엇을 어떤 순서로 보여 줄지는 쪽 목록(PdfPageEntry — 방의 pdfPages)이 정한다. 원본 쪽('o:'), 넣은 PDF의 쪽('p:'),
 * 빈 쪽('b:')을 그리고, 목록의 회전을 쪽 자체 회전에 더한다. 목록이 바뀌면(setPages) 같은 열쇠의 자리는 그대로 옮겨 쓴다.
 * 쪽 번호(1부터)는 언제나 지금 목록에서의 자리다.
 */
import { pdfjs, type OptionalContentConfig, type PDFDocumentProxy } from './pdfDocument.js';
import type { FindMatch } from './findText.js';
import { identityPages, parsePageKey, type PdfPageEntry } from '../../pdf/annotTypes.js';

type PDFPageProxy = Awaited<ReturnType<PDFDocumentProxy['getPage']>>;
type RenderTask = ReturnType<PDFPageProxy['render']>;
export type PageViewport = ReturnType<PDFPageProxy['getViewport']>;

interface Slot {
  page: number;
  entry: PdfPageEntry;
  readonly el: HTMLDivElement;
  canvas: HTMLCanvasElement;
  readonly textLayer: HTMLDivElement;
  readonly hitLayer: HTMLDivElement;
  /** 주석 층(annotLayer.ts가 채운다) — 글자 층 위 */
  readonly annotLayer: HTMLDivElement;
  readonly label: HTMLSpanElement;
  /** 지금 그려진 상태의 viewport(주석 좌표 변환용) */
  viewport: PageViewport | null;
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

/** 글 조각(pdf.js 글자 층의 span 하나) — PDF 좌표 */
export interface TextItemInfo {
  readonly str: string;
  /** [a, b, c, d, e, f] — e, f가 기준점 */
  readonly transform: readonly number[];
  readonly width: number;
  readonly height: number;
}

/** 넣은 PDF(부속)를 연다 — 편집기가 /api/view/pdf-part에서 받아 캐시한다 */
export type PartLoader = (partId: string) => Promise<PDFDocumentProxy>;

export class PageView {
  private slots: Slot[] = [];
  private readonly slotOfEl = new WeakMap<Element, Slot>();
  private scaleValue = 1;
  private viewRotation = 0;
  private current = 1;
  private readonly pageListeners: Array<(page: number) => void> = [];
  private readonly scaleListeners: Array<(scale: number) => void> = [];
  private readonly renderListeners: Array<(page: number, layer: HTMLDivElement, viewport: PageViewport) => void> = [];
  private readonly pagesListeners: Array<() => void> = [];
  private observer: IntersectionObserver | null = null;
  private readonly visible = new Set<Slot>();
  private ocConfig: OptionalContentConfig | null = null;
  private ocVersion = 0;
  private matchesByPage = new Map<number, FindMatch[]>();
  private currentMatch: FindMatch | null = null;
  private readonly textCache = new Map<string, Promise<TextItemInfo[]>>();
  private scrollRaf = 0;
  private defaultW = 612;
  private defaultH = 792;
  /** pdf.js PageViewport 생성자 — 4.2는 클래스를 내보내지 않아 첫 쪽 viewport에서 빌린다(빈 쪽 좌표 변환용) */
  private viewportCtor: (new (params: { viewBox: number[]; scale: number; rotation: number; offsetX: number; offsetY: number; dontFlip: boolean }) => PageViewport) | null = null;

  constructor(
    private readonly host: HTMLElement,
    private readonly scroller: HTMLElement,
    private readonly doc: PDFDocumentProxy,
    private readonly loadPart: PartLoader = () => Promise.reject(new Error('넣은 파일을 열 수 없습니다.')),
  ) {}

  get pageCount(): number {
    return this.slots.length;
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
  /** 지금 쪽 목록(보이는 순서) */
  get entries(): PdfPageEntry[] {
    return this.slots.map((s) => s.entry);
  }

  onPageChange(cb: (page: number) => void): void {
    this.pageListeners.push(cb);
  }
  onScaleChange(cb: (scale: number) => void): void {
    this.scaleListeners.push(cb);
  }
  /** 쪽이 (다시) 그려질 때마다 — 주석 층이 그 viewport로 주석을 다시 그린다 */
  onPageRendered(cb: (page: number, layer: HTMLDivElement, viewport: PageViewport) => void): void {
    this.renderListeners.push(cb);
  }
  /** 쪽 목록이 바뀐 뒤(순서·회전·넣기·지우기) */
  onPagesChange(cb: () => void): void {
    this.pagesListeners.push(cb);
  }
  viewportOf(page: number): PageViewport | null {
    return this.slots[page - 1]?.viewport ?? null;
  }
  pageElement(page: number): HTMLDivElement | null {
    return this.slots[page - 1]?.el ?? null;
  }
  /** 그려진 쪽 번호들 */
  renderedPages(): number[] {
    return this.slots.filter((s) => s.viewport && s.renderedKey !== null).map((s) => s.page);
  }
  /** 주석이 붙는 쪽 열쇠 */
  pageKeyOf(page: number): string {
    return this.slots[page - 1]?.entry.key ?? '';
  }
  /** 열쇠 → 지금 몇 쪽인지(목록에 없으면 null) */
  pageOfKey(pageKey: string): number | null {
    const slot = this.slots.find((s) => s.entry.key === pageKey);
    return slot ? slot.page : null;
  }
  /** 쪽의 기본 크기(pt, 쪽 자체 회전 반영 전) — 빈 쪽을 넣을 때 옆 쪽 크기를 따른다 */
  sizeOf(page: number): { w: number; h: number } {
    const slot = this.slots[page - 1];
    return slot ? { w: slot.baseW, h: slot.baseH } : { w: this.defaultW, h: this.defaultH };
  }

  async init(): Promise<void> {
    const first = await this.doc.getPage(1);
    const vp = first.getViewport({ scale: 1, rotation: 0 });
    this.viewportCtor = vp.constructor as NonNullable<PageView['viewportCtor']>;
    const userUnit = (first as unknown as { userUnit?: number }).userUnit ?? 1;
    this.defaultW = vp.width * userUnit;
    this.defaultH = vp.height * userUnit;
    this.ocConfig = await this.doc.getOptionalContentConfig();
    this.observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          const slot = this.slotOfEl.get(entry.target);
          if (!slot) continue;
          if (entry.isIntersecting) this.visible.add(slot);
          else {
            this.visible.delete(slot);
            this.release(slot);
          }
        }
        this.renderVisible();
      },
      { root: this.scroller, rootMargin: '100% 0px' },
    );
    this.slots = identityPages(this.doc.numPages).map((entry, i) => this.createSlot(entry, i + 1));
    const firstSlot = this.slots[0];
    if (firstSlot) {
      firstSlot.pageRotate = first.rotate;
      firstSlot.sized = true;
    }
    this.layoutAll();
    this.scroller.addEventListener('scroll', () => {
      if (this.scrollRaf) return;
      this.scrollRaf = requestAnimationFrame(() => {
        this.scrollRaf = 0;
        this.updateCurrentPage();
      });
    });
  }

  private createSlot(entry: PdfPageEntry, page: number): Slot {
    const el = document.createElement('div');
    el.className = 'pdf-page';
    el.dataset.page = String(page);
    el.dataset.key = entry.key;
    const canvas = document.createElement('canvas');
    const textLayer = document.createElement('div');
    textLayer.className = 'pdf-text-layer';
    const hitLayer = document.createElement('div');
    hitLayer.className = 'pdf-hit-layer';
    // 고친 글·가리기 미리 보기(editLayer.ts) — 글자 층 위
    const editLayer = document.createElement('div');
    editLayer.className = 'pdf-edit-layer';
    // 양식 입력란(formLayer.ts가 채운다) — 글자 층 위, 주석 층 아래
    const formLayer = document.createElement('div');
    formLayer.className = 'pdf-form-layer';
    const annotLayer = document.createElement('div');
    annotLayer.className = 'pdf-annot-layer';
    const label = document.createElement('span');
    label.className = 'pdf-page-label';
    label.textContent = String(page);
    el.append(canvas, textLayer, hitLayer, editLayer, formLayer, annotLayer, label);
    this.host.append(el);
    const blank = parsePageKey(entry.key)?.kind === 'blank' ? entry.blank : undefined;
    const slot: Slot = {
      page,
      entry,
      el,
      canvas,
      textLayer,
      hitLayer,
      annotLayer,
      label,
      viewport: null,
      baseW: blank?.w ?? this.defaultW,
      baseH: blank?.h ?? this.defaultH,
      pageRotate: 0,
      sized: !!blank,
      renderedKey: null,
      renderTask: null,
      rendering: null,
      textDivs: [],
      textStrs: [],
    };
    this.slotOfEl.set(el, slot);
    this.observer?.observe(el);
    return slot;
  }

  /**
   * 쪽 목록을 바꾼다. 같은 열쇠의 쪽은 자리(그림 포함)를 옮겨 쓰고, 목록의 회전이 바뀐 쪽만 다시 그린다.
   * 지금 보던 쪽(열쇠)이 남아 있으면 그 쪽을 계속 본다.
   */
  setPages(next: readonly PdfPageEntry[]): void {
    if (next.length === 0) return;
    const same = next.length === this.slots.length && next.every((e, i) => e.key === this.slots[i].entry.key && e.rotate === this.slots[i].entry.rotate);
    if (same) return;
    const anchorKey = this.slots[this.current - 1]?.entry.key;
    const anchorOffset = this.slots[this.current - 1] ? this.scroller.scrollTop - this.slots[this.current - 1].el.offsetTop : 0;
    const old = new Map(this.slots.map((s) => [s.entry.key, s]));
    const slots: Slot[] = [];
    next.forEach((entry, i) => {
      let slot = old.get(entry.key);
      if (slot) {
        old.delete(entry.key);
        if (slot.entry.rotate !== entry.rotate && slot.renderedKey !== null) slot.renderedKey = 'stale';
        slot.entry = entry;
      } else {
        slot = this.createSlot(entry, i + 1);
      }
      slot.page = i + 1;
      slot.el.dataset.page = String(i + 1);
      slot.label.textContent = String(i + 1);
      slots.push(slot);
    });
    for (const gone of old.values()) {
      this.observer?.unobserve(gone.el);
      this.visible.delete(gone);
      this.release(gone);
      gone.el.remove();
    }
    // 순서대로 다시 붙인다(append는 옮기기)
    for (const slot of slots) this.host.append(slot.el);
    this.slots = slots;
    this.layoutAll();
    const anchor = anchorKey ? slots.find((s) => s.entry.key === anchorKey) : undefined;
    if (anchor) {
      this.scroller.scrollTop = anchor.el.offsetTop + anchorOffset;
      this.current = anchor.page;
    } else {
      this.current = Math.min(this.current, slots.length);
    }
    for (const cb of this.pageListeners) cb(this.current);
    for (const cb of this.pagesListeners) cb();
    this.renderVisible();
  }

  // ---------- 크기·배율·회전 ----------
  private rotationOf(slot: Slot): number {
    return (slot.pageRotate + slot.entry.rotate + this.viewRotation) % 360;
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
    slot.annotLayer.replaceChildren();
    slot.el.querySelector('.pdf-form-layer')?.replaceChildren();
    slot.el.querySelector('.pdf-edit-layer')?.replaceChildren();
    slot.viewport = null;
    slot.textDivs = [];
    slot.textStrs = [];
    slot.renderedKey = null;
  }
  private renderVisible(): void {
    const slots = [...this.visible].sort((a, b) => Math.abs(a.page - this.current) - Math.abs(b.page - this.current));
    for (const slot of slots) if (slot.renderedKey !== this.keyOf()) void this.renderSlot(slot);
  }

  /** 열쇠의 pdf.js 쪽 — 빈 쪽이면 null */
  private async proxyOf(entry: PdfPageEntry): Promise<PDFPageProxy | null> {
    const parsed = parsePageKey(entry.key);
    if (parsed?.kind === 'orig') return this.doc.getPage(parsed.index + 1);
    if (parsed?.kind === 'part') return (await this.loadPart(parsed.partId)).getPage(parsed.index + 1);
    return null;
  }

  /** 빈 쪽의 viewport — pdf.js 쪽과 같은 좌표 변환(주석 층이 쓴다) */
  private blankViewport(slot: Slot, scale: number, rotation: number): PageViewport {
    return new this.viewportCtor!({ viewBox: [0, 0, slot.baseW, slot.baseH], scale, rotation, offsetX: 0, offsetY: 0, dontFlip: false });
  }

  /** 한 쪽을 그린다(같은 쪽을 동시에 두 번 그리지 않는다). 끝나면 그 쪽의 찾기 강조도 다시 그린다. */
  renderSlot(slot: Slot): Promise<void> {
    if (slot.rendering) return slot.rendering;
    const key = this.keyOf();
    slot.rendering = (async () => {
      try {
        const page = await this.proxyOf(slot.entry);
        if (page && !slot.sized) {
          const vp1 = page.getViewport({ scale: 1, rotation: 0 });
          const userUnit = (page as unknown as { userUnit?: number }).userUnit ?? 1;
          slot.baseW = vp1.width * userUnit;
          slot.baseH = vp1.height * userUnit;
          slot.pageRotate = page.rotate;
          slot.sized = true;
          this.layout(slot);
        }
        const viewport = page
          ? page.getViewport({ scale: this.scaleValue, rotation: this.rotationOf(slot) })
          : this.blankViewport(slot, this.scaleValue, this.rotationOf(slot));
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
        if (page) {
          const task = page.render({
            canvasContext: ctx,
            viewport,
            transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined,
            annotationMode: pdfjs.AnnotationMode.ENABLE_FORMS,
            // 레이어 설정은 원본 문서 것 — 넣은 PDF의 쪽은 그 파일 기본값으로
            optionalContentConfigPromise: this.ocConfig && parsePageKey(slot.entry.key)?.kind === 'orig' ? Promise.resolve(this.ocConfig) : undefined,
          });
          slot.renderTask = task;
          await task.promise;
          if (slot.renderTask !== task) return;
          slot.renderTask = null;
        } else {
          ctx.fillStyle = '#fff';
          ctx.fillRect(0, 0, canvas.width, canvas.height);
        }
        // 다 그린 canvas로 한 번에 바꿔 끼운다(그리는 동안 빈 쪽이 보이지 않게)
        slot.canvas.replaceWith(canvas);
        slot.canvas = canvas;

        const textDivs: HTMLElement[] = [];
        const textStrs: string[] = [];
        slot.textLayer.replaceChildren();
        if (page) {
          const textTask = pdfjs.renderTextLayer({
            textContentSource: page.streamTextContent(),
            container: slot.textLayer,
            viewport,
            textDivs,
            textContentItemsStr: textStrs,
          });
          await textTask.promise;
        }
        slot.textDivs = textDivs;
        slot.textStrs = textStrs;
        slot.renderedKey = key;
        slot.viewport = viewport;
        this.drawHits(slot);
        for (const cb of this.renderListeners) cb(slot.page, slot.annotLayer, viewport);
      } catch (error) {
        if ((error as { name?: string })?.name !== 'RenderingCancelledException') console.error(`쪽 ${slot.page} 그리기 실패:`, error);
      } finally {
        slot.rendering = null;
        // 그리는 동안 배율 등이 바뀌었으면 새 열쇠로 다시
        if (this.visible.has(slot) && this.slots.includes(slot) && slot.renderedKey !== this.keyOf()) queueMicrotask(() => void this.renderSlot(slot));
      }
    })();
    return slot.rendering;
  }

  /** 첫 쪽이 다 그려질 때까지 기다린다(docview:ready 신호용) */
  async firstPaint(): Promise<void> {
    const slot = this.slots[0];
    if (slot) await this.renderSlot(slot);
  }

  /**
   * 작은 그림(썸네일·정리 화면 카드) — 폭 width(px)에 맞춰 box 안에 그린다. 목록의 회전 반영(보기 회전은 빼고).
   * 그린 viewport(주석 겹침용)를 돌려준다.
   */
  async drawThumbnail(page: number, box: HTMLElement, width: number): Promise<PageViewport | null> {
    const slot = this.slots[page - 1];
    if (!slot) return null;
    try {
      const proxy = await this.proxyOf(slot.entry);
      const rotation = ((proxy?.rotate ?? 0) + slot.entry.rotate) % 360;
      const base = proxy ? proxy.getViewport({ scale: 1, rotation }) : this.blankViewport(slot, 1, rotation);
      const scale = width / base.width;
      const viewport = proxy ? proxy.getViewport({ scale, rotation }) : this.blankViewport(slot, scale, rotation);
      const ratio = Math.min(2, window.devicePixelRatio || 1);
      const canvas = document.createElement('canvas');
      canvas.width = Math.floor(viewport.width * ratio);
      canvas.height = Math.floor(viewport.height * ratio);
      box.style.height = `${Math.round(viewport.height)}px`;
      const ctx = canvas.getContext('2d', { alpha: false });
      if (!ctx) return null;
      if (proxy) {
        await proxy.render({ canvasContext: ctx, viewport, transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined }).promise;
      } else {
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
      }
      box.replaceChildren(canvas);
      return viewport;
    } catch (error) {
      console.error(`썸네일 ${page} 실패:`, error);
      return null;
    }
  }

  /** 그 쪽의 양식 위젯(pdf.js 주석 정보) — 원본 쪽만(넣은 PDF의 양식은 저장 때 빠지므로 다루지 않는다) */
  async widgetsOf(page: number): Promise<Array<Record<string, unknown>>> {
    const slot = this.slots[page - 1];
    if (!slot || parsePageKey(slot.entry.key)?.kind !== 'orig') return [];
    const proxy = await this.proxyOf(slot.entry);
    if (!proxy) return [];
    const annots = (await proxy.getAnnotations({ intent: 'display' })) as Array<Record<string, unknown>>;
    return annots.filter((a) => a.subtype === 'Widget');
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
  async textItems(page: number): Promise<string[]> {
    return (await this.textItemInfos(page)).map((i) => i.str);
  }

  /** 쪽의 글 조각(위치 포함) — 글자 층 span과 같은 순서 */
  textItemInfos(page: number): Promise<TextItemInfo[]> {
    const slot = this.slots[page - 1];
    if (!slot) return Promise.resolve([]);
    let cached = this.textCache.get(slot.entry.key);
    if (!cached) {
      cached = this.proxyOf(slot.entry).then(async (p) => {
        if (!p) return [];
        const content = await p.getTextContent();
        return content.items.map((item) =>
          'str' in item ? { str: item.str, transform: item.transform as number[], width: item.width, height: item.height } : { str: '', transform: [1, 0, 0, 1, 0, 0], width: 0, height: 0 },
        );
      });
      this.textCache.set(slot.entry.key, cached);
    }
    return cached;
  }

  /** 글자 층 span → (쪽, 조각 순번) */
  textSpanIndex(span: Element): { page: number; index: number } | null {
    const slot = this.slots.find((s) => s.el.contains(span));
    if (!slot) return null;
    const index = slot.textDivs.indexOf(span as HTMLElement);
    return index < 0 ? null : { page: slot.page, index };
  }
  /** 그 쪽의 글자 층 span들(그려진 경우) */
  textDivsOf(page: number): readonly HTMLElement[] {
    return this.slots[page - 1]?.textDivs ?? [];
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
      this.visible.add(slot);
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
