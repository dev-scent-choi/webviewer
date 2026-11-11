/**
 * 서버 PNG 원천 — 내려받기 권한이 없는 토큰(설계 V5): 쪽 그림(/api/view/pages)·썸네일(?thumb=true)·찾기(/api/view/search)를
 * 지금 보기 화면과 같은 경로로 받는다. PDF·원본은 브라우저로 오지 않는다.
 */
import { SEARCH_LIMIT, pngScale } from './logic.js';
import type { DrawOpts, PageSource } from './pagesBody.js';
import type { PageSize, SearchHit } from './types.js';

export class PngSource implements PageSource {
  readonly kind = 'pages-png' as const;
  readonly textSelectable = false;

  constructor(
    private readonly tokenQuery: string,
    readonly pageCount: number,
    private readonly sizes: readonly PageSize[] | undefined,
  ) {}

  sizeOf(page: number): PageSize | null {
    return this.sizes?.[page - 1] ?? null;
  }

  qualityKey(_page: number, opts: DrawOpts): number {
    return pngScale(opts.percent, opts.dpr);
  }

  async draw(page: number, canvas: HTMLCanvasElement, opts: DrawOpts, signal: AbortSignal): Promise<void> {
    const res = await fetch(`/api/view/pages/${page}?${this.tokenQuery}&scale=${pngScale(opts.percent, opts.dpr)}`, { signal });
    if (!res.ok) throw new Error('쪽 그림을 불러오지 못했습니다.');
    const bitmap = await createImageBitmap(await res.blob());
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    canvas.getContext('2d')?.drawImage(bitmap, 0, 0);
    bitmap.close();
  }

  async search(q: string, signal: AbortSignal): Promise<SearchHit[]> {
    const res = await fetch(`/api/view/search?${this.tokenQuery}&q=${encodeURIComponent(q)}&limit=${SEARCH_LIMIT}`, { signal });
    const json = (await res.json().catch(() => null)) as { success?: boolean; data?: { hits: SearchHit[] }; error?: string } | null;
    if (!res.ok || !json?.success || !json.data) throw new Error(json?.error || '검색에 실패했습니다.');
    return json.data.hits.map((hit) => ({ page: hit.page, rects: hit.rects }));
  }

  async thumbnail(page: number): Promise<HTMLImageElement> {
    const img = new Image();
    img.alt = `${page}쪽`;
    img.decoding = 'async';
    img.src = `/api/view/pages/${page}?${this.tokenQuery}&thumb=true`;
    await img.decode();
    return img;
  }

  destroy(): void {
    // 들고 있는 것이 없다
  }
}
