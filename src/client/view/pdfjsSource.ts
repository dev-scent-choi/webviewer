/**
 * pdf.js 원천 — 내려받기 권한이 있을 때만(설계 V5) /api/view/derived-pdf로 받은 PDF를 브라우저에서 그린다. 여는 방법·보안 설정은
 * PDF 편집기와 같은 pdfDocument.ts(PDF 안 자바스크립트 실행 안 함, worker·cMap·표준 글꼴 자체 호스팅), 찾기 정규화는 findText.ts.
 * view.ts가 동적 import로만 부른다 — PNG 모드에서는 pdf.js 묶음을 받지 않는다.
 */
import { findInPage } from '../pdf/findText.js';
import { openPdfData, pdfjs, type PDFDocumentProxy } from '../pdf/pdfDocument.js';
import { SEARCH_LIMIT, canvasPixelRatio, matchRects, type TextItemGeom, type ViewportGeom } from './logic.js';
import type { DrawOpts, PageSource } from './pagesBody.js';
import type { PageSize, SearchHit } from './types.js';

type PageProxy = Awaited<ReturnType<PDFDocumentProxy['getPage']>>;
type Viewport = ReturnType<PageProxy['getViewport']>;

/** 받을 수 없음(403·202·423 등) — view.ts가 PNG로 내려간다 */
export class DerivedPdfUnavailable extends Error {}

export async function openPdfjsSource(tokenQuery: string, canCopy: boolean, signal: AbortSignal): Promise<PdfjsSource> {
  const res = await fetch(`/api/view/derived-pdf?${tokenQuery}`, { cache: 'no-store', signal });
  if (res.status !== 200) throw new DerivedPdfUnavailable(String(res.status));
  const doc = await openPdfData(new Uint8Array(await res.arrayBuffer()));
  return new PdfjsSource(doc, canCopy);
}

export class PdfjsSource implements PageSource {
  readonly kind = 'pages-pdfjs' as const;
  private readonly sizes: Array<PageSize | null>;
  private readonly texts = new Map<number, Promise<{ items: TextItemGeom[]; vp: ViewportGeom }>>();

  constructor(
    private readonly doc: PDFDocumentProxy,
    /** 글자 층은 복사 권한(canCopy)이 있을 때만(설계 §4) */
    readonly textSelectable: boolean,
  ) {
    this.sizes = new Array<PageSize | null>(doc.numPages).fill(null);
  }

  get pageCount(): number {
    return this.doc.numPages;
  }

  sizeOf(page: number): PageSize | null {
    return this.sizes[page - 1] ?? null;
  }

  qualityKey(_page: number, opts: DrawOpts): number {
    return Math.round(opts.cssWidth * opts.dpr);
  }

  private async open(page: number): Promise<{ proxy: PageProxy; base: Viewport }> {
    const proxy = await this.doc.getPage(page);
    const base = proxy.getViewport({ scale: 1 });
    this.sizes[page - 1] = { w: base.width, h: base.height };
    return { proxy, base };
  }

  async draw(page: number, canvas: HTMLCanvasElement, opts: DrawOpts, signal: AbortSignal): Promise<void> {
    const { proxy, base } = await this.open(page);
    const viewport = proxy.getViewport({ scale: opts.cssWidth / base.width });
    const ratio = canvasPixelRatio(viewport.width, viewport.height, opts.dpr);
    canvas.width = Math.floor(viewport.width * ratio);
    canvas.height = Math.floor(viewport.height * ratio);
    const ctx = canvas.getContext('2d', { alpha: false });
    if (!ctx) throw new Error('쪽을 그릴 캔버스를 만들지 못했습니다.');
    const task = proxy.render({ canvasContext: ctx, viewport, transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined });
    const cancel = () => task.cancel();
    signal.addEventListener('abort', cancel, { once: true });
    try {
      await task.promise;
    } finally {
      signal.removeEventListener('abort', cancel);
    }
  }

  async drawText(page: number, layer: HTMLElement, cssWidth: number): Promise<void> {
    const { proxy, base } = await this.open(page);
    const viewport = proxy.getViewport({ scale: cssWidth / base.width });
    layer.replaceChildren();
    layer.style.setProperty('--scale-factor', String(viewport.scale));
    await pdfjs.renderTextLayer({ textContentSource: proxy.streamTextContent(), container: layer, viewport, textDivs: [] }).promise;
  }

  private textOf(page: number): Promise<{ items: TextItemGeom[]; vp: ViewportGeom }> {
    let hit = this.texts.get(page);
    if (!hit) {
      hit = this.open(page).then(async ({ proxy, base }) => {
        const content = await proxy.getTextContent();
        const items: TextItemGeom[] = [];
        for (const item of content.items) {
          if ('str' in item) items.push({ str: item.str, transform: item.transform as number[], width: item.width, height: item.height });
        }
        return { items, vp: { transform: base.transform, width: base.width, height: base.height } };
      });
      this.texts.set(page, hit);
    }
    return hit;
  }

  async search(q: string, signal: AbortSignal): Promise<SearchHit[]> {
    const hits: SearchHit[] = [];
    for (let page = 1; page <= this.doc.numPages && hits.length < SEARCH_LIMIT; page += 1) {
      if (signal.aborted) throw new DOMException('찾기를 멈췄습니다.', 'AbortError');
      const { items, vp } = await this.textOf(page);
      for (const match of findInPage(page, items, q)) {
        hits.push({ page, rects: matchRects(match, items, vp) });
        if (hits.length >= SEARCH_LIMIT) break;
      }
    }
    return hits;
  }

  async thumbnail(page: number, width: number): Promise<HTMLCanvasElement> {
    const { proxy, base } = await this.open(page);
    const viewport = proxy.getViewport({ scale: width / base.width });
    const ratio = Math.min(2, window.devicePixelRatio || 1);
    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(viewport.width * ratio);
    canvas.height = Math.floor(viewport.height * ratio);
    const ctx = canvas.getContext('2d', { alpha: false });
    if (!ctx) throw new Error('썸네일을 그릴 캔버스를 만들지 못했습니다.');
    await proxy.render({ canvasContext: ctx, viewport, transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined }).promise;
    return canvas;
  }

  destroy(): void {
    void this.doc.destroy();
  }
}
