/**
 * 쪽 문서 안쪽(ViewBody) — 쪽마다 .page-frame(지금 보기 화면 viewer.js와 같은 틀·CSS)을 만들고, 그림은 원천(PageSource)이
 * 그린다: 서버 PNG(pngSource.ts) 또는 브라우저 pdf.js(pdfjsSource.ts). 확대(쪽 칸 폭 %)·보기 회전(CSS)·현재 쪽·가상화(보이는 쪽 +
 * 지금 쪽 ±2만 캔버스 유지)·찾기 강조는 두 원천이 같다 — 마크업·측정·찾기 겹침이 .page-frame 비율 좌표 하나로 붙는다.
 * pdf.js가 실패하면 replaceSource로 같은 몸체에 PNG 원천을 끼운다(등록한 콜백·확대는 그대로).
 */
import { DEFAULT_ZOOM_INDEX, LatestDraws, ZOOM_STEPS, clampZoomIndex, fitPagePercent, keepFinishedDraw, needsRedraw, pagesToKeep, snapZoomIndex, textLayerStale } from './logic.js';
import type { BodyKind, PageSize, SearchHit, ViewBody, ZoomMode } from './types.js';

export interface DrawOpts {
  /** 확대 %(ZOOM_STEPS) */
  readonly percent: number;
  /** 쪽 틀의 CSS 폭(px) */
  readonly cssWidth: number;
  /** min(devicePixelRatio, 2) */
  readonly dpr: number;
}

export interface PageSource {
  readonly kind: 'pages-png' | 'pages-pdfjs' | 'pages-artcraft';
  readonly pageCount: number;
  /** 글자 층을 그리는가(pdf.js + canCopy) */
  readonly textSelectable: boolean;
  /** 쪽 크기(pt) — 아직 모르면 null */
  sizeOf(page: number): PageSize | null;
  /** 같은 값이면 다시 그리지 않는다(PNG: 서버 배율, pdf.js: 화소 폭) */
  qualityKey(page: number, opts: DrawOpts): number;
  /** canvas에 그린다(크기도 정한다). 실패하면 throw, 취소면 AbortError·RenderingCancelledException */
  draw(page: number, canvas: HTMLCanvasElement, opts: DrawOpts, signal: AbortSignal): Promise<void>;
  drawText?(page: number, layer: HTMLElement, cssWidth: number): Promise<void>;
  search(q: string, signal: AbortSignal): Promise<SearchHit[]>;
  thumbnail(page: number, width: number): Promise<HTMLCanvasElement | HTMLImageElement>;
  /** 이 쪽을 다 썼다 — pdf.js는 쪽 객체의 연산 목록·풀어 놓은 그림을 놓는다(긴·스캔 PDF 메모리) */
  release?(page: number): void;
  destroy(): void;
}

const MAX_CONCURRENT_DRAWS = 3;
const KEEP_RADIUS = 2;
const DEFAULT_SIZE: PageSize = { w: 612, h: 792 };

function isCancel(error: unknown): boolean {
  const name = (error as { name?: string } | null)?.name;
  return name === 'AbortError' || name === 'RenderingCancelledException';
}

export class PagesBody implements ViewBody {
  /** #pages(.pages-container) — mount 뒤에 쓴다 */
  container!: HTMLElement;
  zoomIndex = DEFAULT_ZOOM_INDEX;
  rotateDeg = 0;
  private frames: HTMLElement[] = [];
  private current = 1;
  private readonly visible = new Set<number>();
  private readonly drawnKey = new Map<number, number>();
  /** 글자 층을 그린 쪽 폭(CSS px) — 확대로 폭이 바뀌면 글자 층만 다시 그린다 */
  private readonly textWidth = new Map<number, number>();
  private readonly pending = new LatestDraws();
  /** 글자 층 그리기 차례 — 겹쳐 그리거나 비운 쪽에 늦게 붙지 않게 마지막 것만 붙인다 */
  private readonly textGen = new Map<number, number>();
  private resizeObserver: ResizeObserver | null = null;
  private abort = new AbortController();
  private active = 0;
  private readonly waiters: Array<() => void> = [];
  private reloadTimer: ReturnType<typeof setTimeout> | null = null;
  private drawObserver: IntersectionObserver | null = null;
  private pageObserver: IntersectionObserver | null = null;
  private readonly pageCbs: Array<(page: number) => void> = [];
  private readonly zoomCbs: Array<(percent: number) => void> = [];
  private readonly drawCbs: Array<(page: number) => void> = [];
  private readonly errorCbs: Array<(error: unknown) => void> = [];
  private readonly sourceCbs: Array<() => void> = [];

  constructor(
    private source: PageSource,
    private readonly initialSizes: readonly PageSize[] | undefined,
  ) {}

  get kind(): BodyKind {
    return this.source.kind;
  }
  get textSelectable(): boolean {
    return this.source.textSelectable;
  }
  get currentPage(): number {
    return this.current;
  }
  get zoomPercent(): number {
    return ZOOM_STEPS[this.zoomIndex];
  }

  async mount(host: HTMLElement): Promise<void> {
    this.container = host;
    host.classList.add('pages-container');
    this.buildFrames();
    this.applyZoom();
    // 창·썸네일 칸 크기가 바뀌면 쪽 폭(칸의 %)도 바뀐다 — 확대와 같이 보이는 쪽과 글자 층을 새 폭으로
    let lastWidth = host.clientWidth;
    this.resizeObserver = new ResizeObserver(() => {
      if (host.clientWidth === lastWidth) return; // 쪽이 그려지며 높이만 느는 것은 무시
      lastWidth = host.clientWidth;
      this.scheduleQualityReload();
    });
    this.resizeObserver.observe(host);
  }

  pageCount(): number {
    return this.source.pageCount;
  }
  pageFrame(page: number): HTMLElement | null {
    return this.frames[page - 1] ?? null;
  }
  /** 쪽 크기(pt) — 원천이 아직 모르면 서버가 준 크기(info.pages), 그것도 없으면 Letter */
  pageSize(page: number): PageSize {
    return this.source.sizeOf(page) ?? this.initialSizes?.[page - 1] ?? this.initialSizes?.[0] ?? DEFAULT_SIZE;
  }

  onPage(cb: (page: number) => void): void {
    this.pageCbs.push(cb);
  }
  onZoom(cb: (percent: number) => void): void {
    this.zoomCbs.push(cb);
  }
  /** 쪽 그림이 (다시) 그려질 때마다 */
  onDraw(cb: (page: number) => void): void {
    this.drawCbs.push(cb);
  }
  /** 그리기 실패(취소 빼고) — view.ts가 pdf.js → PNG로 내려간다 */
  onDrawError(cb: (error: unknown) => void): void {
    this.errorCbs.push(cb);
  }
  /** replaceSource 뒤(쪽 틀을 다시 만들었다 — 겹침·썸네일을 다시 붙인다) */
  onSourceChange(cb: () => void): void {
    this.sourceCbs.push(cb);
  }

  replaceSource(next: PageSource): void {
    const page = this.current;
    this.abort.abort();
    this.abort = new AbortController();
    this.source.destroy();
    this.source = next;
    this.buildFrames();
    this.applyZoom();
    this.frames[page - 1]?.scrollIntoView({ block: 'start' });
    for (const cb of this.sourceCbs) cb();
  }

  goto(page: number): void {
    const target = Math.min(Math.max(Math.round(page), 1), this.frames.length || 1);
    this.frames[target - 1]?.scrollIntoView({ block: 'start' });
    this.setCurrent(target);
  }

  zoom(z: ZoomMode): void {
    if (z === 'fit-width' || z === 'fit-page') {
      // 1120px 읽기 폭 제한을 풀어 창 끝까지(viewer.js setupZoomFit와 같다)
      this.container.classList.add('pages-edge-to-edge');
      const toolbarHeight = document.getElementById('toolbar')?.offsetHeight ?? 0;
      const percent = z === 'fit-width' ? 100 : fitPagePercent(this.container.clientWidth, window.innerHeight - toolbarHeight, this.pageSize(this.current));
      this.zoomIndex = snapZoomIndex(percent);
    } else {
      this.zoomIndex = snapZoomIndex(z);
    }
    this.applyZoom();
  }

  setZoomIndex(index: number): void {
    const next = clampZoomIndex(index);
    if (next === this.zoomIndex) return;
    this.zoomIndex = next;
    this.applyZoom();
  }

  /** 커서가 가리키는 곳을 그대로 둔 채 확대(viewer.js zoomAtPoint 그대로 — 세로는 창, 가로는 쪽 칸이 스크롤한다) */
  zoomAtPoint(nextIndex: number, clientX: number, clientY: number): void {
    const bounded = clampZoomIndex(nextIndex);
    if (bounded === this.zoomIndex) return;
    const ratio = ZOOM_STEPS[bounded] / ZOOM_STEPS[this.zoomIndex];
    const rect = this.container.getBoundingClientRect();
    const containerDocTop = rect.top + window.scrollY;
    const contentX = this.container.scrollLeft + (clientX - rect.left);
    const contentY = clientY + window.scrollY - containerDocTop;
    this.zoomIndex = bounded;
    this.applyZoom();
    this.container.scrollLeft = contentX * ratio - (clientX - rect.left);
    window.scrollTo(window.scrollX, containerDocTop + contentY * ratio - clientY);
  }

  /** 보기 전용 회전(저장 안 함) — 틀을 CSS로 돌린다(style.css .pages-rotated) */
  rotate(deg: 90 | -90): void {
    this.rotateDeg = (this.rotateDeg + deg + 360) % 360;
    this.container.style.setProperty('--page-rotate', `${this.rotateDeg}deg`);
    this.container.classList.toggle('pages-rotated', this.rotateDeg !== 0);
  }

  search(q: string, signal: AbortSignal): Promise<SearchHit[]> {
    return this.source.search(q, signal);
  }
  async thumbnail(page: number, width: number): Promise<HTMLCanvasElement | HTMLImageElement> {
    const source = this.source;
    try {
      return await source.thumbnail(page, width);
    } finally {
      // 몸체가 그려 두지 않았고 곧 그리지도 않을(보이는 쪽·지금 쪽 ±2 밖) 쪽이면 썸네일 때문에 받은 쪽 객체를 놓는다
      if (!this.drawnKey.has(page) && !keepFinishedDraw(page, this.current, this.frames.length, KEEP_RADIUS, this.visible)) source.release?.(page);
    }
  }

  /** 찾은 곳 강조(viewer.js renderPageHitOverlays 그대로) — 지금 결과의 첫 사각형을 돌려준다 */
  showHits(hits: readonly SearchHit[], current: number): HTMLElement | null {
    this.clearHits();
    let currentEl: HTMLElement | null = null;
    hits.forEach((hit, index) => {
      const frame = this.frames[hit.page - 1];
      if (!frame) return;
      for (const rect of hit.rects) {
        const el = document.createElement('div');
        el.className = index === current ? 'search-hit-overlay current' : 'search-hit-overlay';
        el.style.left = `${rect.x * 100}%`;
        el.style.top = `${rect.y * 100}%`;
        el.style.width = `${Math.max(rect.w * 100, 0.5)}%`;
        el.style.height = `${Math.max(rect.h * 100, 0.5)}%`;
        frame.append(el);
        if (index === current && !currentEl) currentEl = el;
      }
    });
    return currentEl;
  }

  clearHits(): void {
    for (const el of this.container.querySelectorAll('.search-hit-overlay')) el.remove();
  }

  destroy(): void {
    this.abort.abort();
    this.drawObserver?.disconnect();
    this.pageObserver?.disconnect();
    this.resizeObserver?.disconnect();
    if (this.reloadTimer !== null) clearTimeout(this.reloadTimer);
    this.source.destroy();
  }

  private buildFrames(): void {
    this.drawObserver?.disconnect();
    this.pageObserver?.disconnect();
    this.visible.clear();
    this.drawnKey.clear();
    this.textWidth.clear();
    this.pending.clear();
    // 화면 근처(한 화면 위아래)에 들어오면 그리고, 멀어지면 trim()이 비운다
    this.drawObserver = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          const page = Number((entry.target as HTMLElement).dataset.page);
          if (entry.isIntersecting) {
            this.visible.add(page);
            void this.ensureDrawn(page);
          } else {
            this.visible.delete(page);
          }
        }
        this.trim();
      },
      { rootMargin: '100% 0px' },
    );
    this.pageObserver = new IntersectionObserver(
      (entries) => {
        const shown = entries.find((entry) => entry.isIntersecting);
        if (shown) this.setCurrent(Number((shown.target as HTMLElement).dataset.page));
      },
      { threshold: 0.4 },
    );
    this.frames = [];
    for (let page = 1; page <= this.source.pageCount; page += 1) {
      const frame = document.createElement('div');
      frame.className = 'page-frame page-loading';
      frame.dataset.page = String(page);
      // 크기를 미리 알면 그리기 전에도 틀 높이가 맞아 쪽 이동·스크롤이 어긋나지 않는다
      const size = this.pageSize(page);
      frame.style.aspectRatio = `${size.w} / ${size.h}`;
      frame.style.minHeight = '0';
      const canvas = document.createElement('canvas');
      canvas.dataset.page = String(page);
      const label = document.createElement('span');
      label.className = 'page-number';
      label.textContent = String(page);
      frame.append(canvas, label);
      if (this.source.textSelectable) {
        const layer = document.createElement('div');
        layer.className = 'pdf-text-layer';
        frame.append(layer);
      }
      this.frames.push(frame);
      this.drawObserver.observe(frame);
      this.pageObserver.observe(frame);
    }
    this.container.replaceChildren(...this.frames);
    // 옵저버 콜백(다음 틱)을 기다리지 않고 1쪽부터(viewer.js buildPages와 같다)
    void this.ensureDrawn(1);
  }

  private applyZoom(): void {
    this.container.style.setProperty('--page-width', `${this.zoomPercent}%`);
    for (const cb of this.zoomCbs) cb(this.zoomPercent);
    this.scheduleQualityReload();
  }

  /** 확대가 멈춘 뒤 보이는 쪽을 새 배율로(viewer.js scheduleQualityReload와 같은 350ms) */
  private scheduleQualityReload(): void {
    if (this.reloadTimer !== null) clearTimeout(this.reloadTimer);
    this.reloadTimer = setTimeout(() => {
      this.reloadTimer = null;
      for (const page of this.visible) void this.ensureDrawn(page);
    }, 350);
  }

  private setCurrent(page: number): void {
    if (page === this.current) return;
    this.current = page;
    this.trim();
    for (const cb of this.pageCbs) cb(page);
  }

  private drawOpts(frame: HTMLElement): DrawOpts {
    return { percent: this.zoomPercent, cssWidth: frame.clientWidth || this.container.clientWidth, dpr: Math.min(window.devicePixelRatio || 1, 2) };
  }

  /** 대용량 쪽 그림을 한꺼번에 받으면 스크롤·단추 반응이 늦어진다 — 동시에 3장까지(viewer.js withLoadSlot) */
  private async withSlot(task: () => Promise<void>): Promise<void> {
    if (this.active >= MAX_CONCURRENT_DRAWS) await new Promise<void>((resolve) => this.waiters.push(resolve));
    this.active += 1;
    try {
      await task();
    } finally {
      this.active -= 1;
      this.waiters.shift()?.();
    }
  }

  private async ensureDrawn(page: number): Promise<void> {
    const frame = this.frames[page - 1];
    if (!frame) return;
    const opts = this.drawOpts(frame);
    const want = this.source.qualityKey(page, opts);
    if (this.pending.isLatest(page, want)) return;
    if (!needsRedraw(this.drawnKey.get(page) ?? 0, want)) {
      // 지금 그림으로 충분하다 — 다른 배율로 그리던 것은 끝나도 붙이지 않고, 글자 층만 지금 쪽 폭에 다시 맞춘다
      this.pending.drop(page);
      if (this.drawnKey.has(page) && textLayerStale(this.textWidth.get(page), opts.cssWidth)) await this.drawTextLayer(page, frame, opts.cssWidth);
      return;
    }
    this.pending.start(page, want);
    const source = this.source;
    const signal = this.abort.signal;
    await this.withSlot(async () => {
      // 자리를 기다리는 사이 더 새 배율이 왔으면 그것이 그린다
      if (signal.aborted || source !== this.source || !this.pending.isLatest(page, want)) return;
      // 다 그린 canvas로 한 번에 바꿔 끼운다(그리는 동안 빈 쪽이 보이지 않게)
      const next = document.createElement('canvas');
      // 버리는 캔버스는 0×0으로 — iOS Safari는 캔버스 메모리를 GC 뒤에야 늦게 돌려줘 연달아 확대하면 상한에 걸린다
      const discard = () => {
        next.width = 0;
        next.height = 0;
      };
      try {
        await source.draw(page, next, opts, signal);
      } catch (error) {
        discard();
        this.pending.end(page, want);
        if (isCancel(error) || signal.aborted) return;
        frame.classList.add('page-error');
        for (const cb of this.errorCbs) cb(error);
        return;
      }
      // 그리는 사이 원천이 바뀌었거나, 더 새 요청이 왔거나(그것이 붙는다) 지금 그림을 쓰기로 했으면 버린다
      if (signal.aborted || source !== this.source || !this.pending.isLatest(page, want)) {
        discard();
        return;
      }
      this.pending.end(page, want);
      // 그리는 사이 창 밖이 된 쪽은 붙이지 않는다(다음 스크롤까지 메모리를 잡지 않게)
      if (!keepFinishedDraw(page, this.current, this.frames.length, KEEP_RADIUS, this.visible)) {
        discard();
        if (!this.drawnKey.has(page)) source.release?.(page);
        return;
      }
      next.dataset.page = String(page);
      const old = frame.querySelector('canvas');
      if (old) {
        old.width = 0;
        old.height = 0;
        old.replaceWith(next);
      }
      frame.style.aspectRatio = `${next.width} / ${next.height}`;
      frame.classList.remove('page-loading', 'page-error');
      frame.classList.add('page-ready');
      this.drawnKey.set(page, want);
      await this.drawTextLayer(page, frame, opts.cssWidth);
      for (const cb of this.drawCbs) cb(page);
    });
  }

  /** 글자 층을 새로 그려 바꿔 끼운다 — 그사이 더 새 글자 층이 시작됐거나 쪽을 비웠으면 버린다 */
  private async drawTextLayer(page: number, frame: HTMLElement, cssWidth: number): Promise<void> {
    const layer = frame.querySelector<HTMLElement>('.pdf-text-layer');
    const source = this.source;
    if (!layer || !source.drawText) return;
    const gen = (this.textGen.get(page) ?? 0) + 1;
    this.textGen.set(page, gen);
    this.textWidth.set(page, cssWidth);
    const fresh = layer.cloneNode(false) as HTMLElement;
    try {
      await source.drawText(page, fresh, cssWidth);
    } catch {
      return;
    }
    if (this.textGen.get(page) !== gen || source !== this.source || !this.drawnKey.has(page)) return;
    layer.replaceWith(fresh);
  }

  /** 멀어진 쪽의 캔버스를 비워 메모리를 돌려준다 — 틀의 가로세로 비율은 남아 스크롤 위치가 그대로다 */
  private release(page: number): void {
    const frame = this.frames[page - 1];
    if (!frame || !this.drawnKey.has(page)) return;
    const canvas = frame.querySelector('canvas');
    if (canvas) {
      canvas.width = 0;
      canvas.height = 0;
    }
    frame.querySelector('.pdf-text-layer')?.replaceChildren();
    this.drawnKey.delete(page);
    this.textWidth.delete(page);
    // 그리던 글자 층은 붙이지 않는다
    this.textGen.set(page, (this.textGen.get(page) ?? 0) + 1);
    this.source.release?.(page);
    frame.classList.remove('page-ready');
    frame.classList.add('page-loading');
  }

  private trim(): void {
    const keep = pagesToKeep(this.current, this.frames.length, KEEP_RADIUS, this.visible);
    for (const page of [...this.drawnKey.keys()]) if (!keep.has(page)) this.release(page);
  }
}
