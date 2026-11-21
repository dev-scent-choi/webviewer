/**
 * 아트크래프트 원천(설계 §4.1) — 원본을 브라우저 Worker 안 ArtCraft wasm 엔진이 그린다. 서버 변환을 부르지 않는다.
 * 열기에 실패하면 ArtcraftUnavailable — view.ts가 리브레오피스 원천으로 바꾼다. 열린 뒤 그리기 오류는 PagesBody.onDrawError로 간다.
 * view.ts가 동적 import로만 부른다.
 */
import {
  ByteLru,
  JobQueue,
  drawPlan,
  pageCountSuspicious,
  quickPixelWidth,
  runHits,
  runSpanGeom,
  targetPixelWidth,
  type InitReply,
  type PageText,
  type RenderReply,
  type TextReply,
  type WorkerCall,
  type WorkerReply,
  type WorkerRequest,
} from './artcraftLogic.js';
import { SEARCH_LIMIT } from './logic.js';
import type { DrawOpts, PageSource } from './pagesBody.js';
import type { PageSize, SearchHit, ViewInfo } from './types.js';

/** 브라우저 엔진으로 열 수 없음 — view.ts가 리브레오피스로 넘어간다 */
export class ArtcraftUnavailable extends Error {}

/** 열기 = wasm·글꼴·원본 받기 + 엔진 open(설계에 값이 없어 이 계획이 정함 — 느린 회선의 4 MB 받기를 감안) */
export const OPEN_TIMEOUT_MS = 60_000;
/** 첫 쪽(과 그 뒤 모든 요청)이 이 안에 끝나지 않으면 Worker를 끄고 리브레오피스로(설계 §4.2·§6) */
export const RENDER_TIMEOUT_MS = 20_000;
/** 쪽 그림 보관 상한 — 1600×2071 한 장이 약 13 MB라 보이는 쪽 ±2 정도 */
const IMAGE_CACHE_BYTES = 64 * 1024 * 1024;

const abortError = () => new DOMException('그리기를 멈췄습니다.', 'AbortError');

interface Job {
  readonly msg: WorkerCall;
  readonly timeoutMs: number;
  readonly signal?: AbortSignal;
  readonly resolve: (reply: WorkerReply) => void;
  readonly reject: (error: unknown) => void;
  id: number;
  /** 끝나면 떼어 낸다 — 문서 수명 signal에 남으면 job → resolve → 그림 버퍼를 붙잡아 LRU가 무력해진다 */
  onAbort?: () => void;
}

/** Worker에 한 번에 하나씩 보낸다(엔진은 한 줄) — 급한 것 먼저, 취소된 것은 줄에서 뺀다, 시간 넘기면 Worker를 끈다 */
export class WorkerClient {
  private readonly queue = new JobQueue<Job>();
  private busy: Job | null = null;
  private nextId = 1;
  private dead: Error | null = null;
  private timer: ReturnType<typeof setTimeout> | null = null;

  constructor(private readonly worker: Worker) {
    worker.onmessage = (event: MessageEvent<WorkerReply>) => this.finish(event.data);
    worker.onerror = (event) => {
      event.preventDefault();
      this.die(new ArtcraftUnavailable(event.message || '브라우저 엔진이 멈췄습니다'));
    };
    worker.onmessageerror = () => this.die(new ArtcraftUnavailable('브라우저 엔진의 답을 읽지 못했습니다'));
  }

  call<R extends WorkerReply>(msg: WorkerCall, timeoutMs: number, urgent = true, signal?: AbortSignal): Promise<R> {
    if (this.dead) return Promise.reject(this.dead);
    if (signal?.aborted) return Promise.reject(abortError());
    return new Promise<R>((resolve, reject) => {
      const job: Job = { msg, timeoutMs, signal, resolve: resolve as (reply: WorkerReply) => void, reject, id: 0 };
      if (signal) {
        job.onAbort = () => {
          if (this.queue.remove(job)) this.settle(job, () => reject(abortError()));
        };
        signal.addEventListener('abort', job.onAbort, { once: true });
      }
      this.queue.push(job, urgent);
      this.pump();
    });
  }

  terminate(): void {
    this.die(new ArtcraftUnavailable('브라우저 엔진을 닫았습니다'));
  }

  private pump(): void {
    if (this.busy || this.dead) return;
    const job = this.queue.shift();
    if (!job) return;
    job.id = this.nextId++;
    this.busy = job;
    this.timer = setTimeout(() => this.die(new ArtcraftUnavailable('브라우저 엔진이 제때 답하지 않았습니다')), job.timeoutMs);
    const request: WorkerRequest = { ...job.msg, id: job.id };
    this.worker.postMessage(request);
  }

  private finish(reply: WorkerReply): void {
    const job = this.busy;
    if (!job || reply.id !== job.id) return;
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = null;
    this.busy = null;
    if ('error' in reply && reply.fatal) {
      // wasm trap 뒤의 엔진 상태는 믿을 수 없다 — 끄고 이후 요청은 바로 거절
      this.die(new ArtcraftUnavailable(reply.error), job);
      return;
    }
    this.settle(job, () => {
      if (job.signal?.aborted) job.reject(abortError());
      else if ('error' in reply) job.reject(new Error(reply.error));
      else job.resolve(reply);
    });
    this.pump();
  }

  private settle(job: Job, done: () => void): void {
    if (job.onAbort) job.signal?.removeEventListener('abort', job.onAbort);
    job.onAbort = undefined;
    done();
  }

  private die(error: Error, current: Job | null = this.busy): void {
    if (this.dead) return;
    this.dead = error;
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = null;
    this.worker.terminate();
    if (current) this.settle(current, () => current.reject(error));
    this.busy = null;
    for (let job = this.queue.shift(); job; job = this.queue.shift()) {
      const j = job;
      this.settle(j, () => j.reject(error));
    }
  }
}

function paint(canvas: HTMLCanvasElement, img: ImageData): void {
  canvas.width = img.width;
  canvas.height = img.height;
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('쪽을 그릴 캔버스를 만들지 못했습니다.');
  ctx.putImageData(img, 0, 0);
}

let measureCtx: CanvasRenderingContext2D | null = null;
function naturalWidth(text: string, fontPx: number): number {
  measureCtx ??= document.createElement('canvas').getContext('2d');
  if (!measureCtx) return 0;
  measureCtx.font = `${fontPx}px sans-serif`;
  return measureCtx.measureText(text).width;
}

export class ArtcraftSource implements PageSource {
  readonly kind = 'pages-artcraft' as const;
  private readonly images = new ByteLru<ImageData>(IMAGE_CACHE_BYTES, (img) => img.data.byteLength);
  private readonly texts = new Map<number, Promise<PageText>>();
  /** 쪽마다 진행 중인 선명하게 그리기 — release·다음 draw가 취소한다(PagesBody signal은 문서 수명이라 따로 둔다) */
  private readonly sharpening = new Map<number, AbortController>();
  /** 선명하게 그리기가 실패한 쪽 — qualityKey를 1 올려 PagesBody가 다음 ensureDrawn에서 다시 그리게 한다 */
  private readonly sharpFailed = new Set<number>();

  constructor(
    private readonly client: WorkerClient,
    private readonly sizes: readonly PageSize[],
    /** 글자 층은 복사 권한(canCopy)이 있을 때만(설계 §4.1) */
    readonly textSelectable: boolean,
  ) {}

  get pageCount(): number {
    return this.sizes.length;
  }

  sizeOf(page: number): PageSize | null {
    return this.sizes[page - 1] ?? null;
  }

  qualityKey(_page: number, opts: DrawOpts): number {
    return Math.ceil(opts.cssWidth * opts.dpr) + (this.sharpFailed.has(_page) ? 1 : 0);
  }

  /** 열 때 1쪽 1배를 미리 — 첫 쪽 20초 제한을 여기서 재고, 첫 draw는 보관한 그림을 바로 쓴다 */
  async warmUp(): Promise<void> {
    await this.render(1, quickPixelWidth(this.sizes[0]), true);
  }

  async draw(page: number, canvas: HTMLCanvasElement, opts: DrawOpts, signal: AbortSignal): Promise<void> {
    const size = this.sizes[page - 1];
    if (!size) throw new Error(`${page}쪽이 없습니다.`);
    const px = targetPixelWidth(opts.cssWidth, opts.dpr, size);
    this.sharpFailed.delete(page);
    this.sharpening.get(page)?.abort();
    this.sharpening.delete(page);
    if (drawPlan(this.pixelWidthsOf(page), px, size.w) !== 'two-stage') {
      paint(canvas, await this.render(page, px, true, signal));
      return;
    }
    paint(canvas, await this.render(page, quickPixelWidth(size), true, signal));
    // 선명한 그림은 급하지 않은 줄로 — 다 되면 같은 캔버스를 바꿔 칠한다(PagesBody가 버린 캔버스는 폭이 0)
    const sharpen = new AbortController();
    this.sharpening.set(page, sharpen);
    this.render(page, px, false, sharpen.signal).then(
      (img) => {
        if (canvas.width > 0) paint(canvas, img);
      },
      (error: unknown) => {
        // 실패하면 1배 그림이 남는다 — qualityKey가 달라져 다음 스크롤·확대 때 다시 그린다(취소는 제외)
        if (!(error instanceof DOMException && error.name === 'AbortError')) this.sharpFailed.add(page);
      },
    ).finally(() => {
      if (this.sharpening.get(page) === sharpen) this.sharpening.delete(page);
    });
  }

  async drawText(page: number, layer: HTMLElement, cssWidth: number): Promise<void> {
    const text = await this.textOf(page);
    layer.replaceChildren();
    for (const run of text.runs) {
      const g = runSpanGeom(run, text, cssWidth);
      const span = document.createElement('span');
      span.textContent = run.t;
      span.style.left = `${g.left}%`;
      span.style.top = `${g.top}%`;
      span.style.fontSize = `${g.fontPx}px`;
      span.style.fontFamily = 'sans-serif';
      const natural = naturalWidth(run.t, g.fontPx);
      if (natural > 0) span.style.transform = `scaleX(${g.widthPx / natural})`;
      layer.append(span);
    }
  }

  async search(q: string, signal: AbortSignal): Promise<SearchHit[]> {
    const hits: SearchHit[] = [];
    for (let page = 1; page <= this.sizes.length && hits.length < SEARCH_LIMIT; page += 1) {
      if (signal.aborted) throw new DOMException('찾기를 멈췄습니다.', 'AbortError');
      hits.push(...runHits(page, await this.textOf(page), q));
    }
    return hits.slice(0, SEARCH_LIMIT);
  }

  async thumbnail(page: number, width: number): Promise<HTMLCanvasElement> {
    const img = await this.render(page, Math.ceil(width * Math.min(2, window.devicePixelRatio || 1)), false);
    const canvas = document.createElement('canvas');
    paint(canvas, img);
    return canvas;
  }

  /** 이 쪽을 다 썼다 — 진행 중인 선명하게 그리기를 취소하고 보관 그림을 놓는다(설계 §4.1) */
  release(page: number): void {
    this.sharpening.get(page)?.abort();
    this.sharpening.delete(page);
    for (const px of this.pixelWidthsOf(page)) this.images.delete(`${page}:${px}`);
  }

  destroy(): void {
    this.client.terminate();
  }

  private pixelWidthsOf(page: number): number[] {
    const prefix = `${page}:`;
    return this.images
      .keys()
      .filter((key) => key.startsWith(prefix))
      .map((key) => Number(key.slice(prefix.length)));
  }

  private async render(page: number, px: number, urgent: boolean, signal?: AbortSignal): Promise<ImageData> {
    const key = `${page}:${px}`;
    const hit = this.images.get(key);
    if (hit) return hit;
    const size = this.sizes[page - 1];
    const reply = await this.client.call<RenderReply>({ cmd: 'render', page, scale: px / size.w }, RENDER_TIMEOUT_MS, urgent, signal);
    const img = new ImageData(new Uint8ClampedArray(reply.rgba), reply.w, reply.h);
    this.images.set(key, img);
    return img;
  }

  private textOf(page: number): Promise<PageText> {
    let hit = this.texts.get(page);
    if (!hit) {
      hit = this.client.call<TextReply>({ cmd: 'text', page }, RENDER_TIMEOUT_MS, false).then((reply) => reply.text);
      hit.catch(() => this.texts.delete(page)); // 실패한 것은 다음에 다시 묻는다
      this.texts.set(page, hit);
    }
    return hit;
  }
}

/** Worker를 띄워 엔진·글꼴·원본을 받고 열어 1쪽 1배까지 그린다. 어떤 실패든 ArtcraftUnavailable(설계 §4.2 전환 규칙) */
export async function openArtcraftSource(info: ViewInfo): Promise<ArtcraftSource> {
  const a = info.artcraft;
  if (!a) throw new ArtcraftUnavailable('브라우저 엔진 정보가 없습니다');
  let client: WorkerClient | null = null;
  try {
    client = new WorkerClient(new Worker(a.workerUrl, { type: 'module' }));
    const opened = await client.call<InitReply>(
      { cmd: 'init', wasm: a.assets.wasm, fonts: a.assets.fonts, original: a.originalUrl, name: `document.${a.ext}` },
      OPEN_TIMEOUT_MS,
    );
    if (opened.count === 0) throw new ArtcraftUnavailable('쪽이 없습니다');
    if (pageCountSuspicious(a.experimental, opened.count, info.pageCount)) {
      throw new ArtcraftUnavailable(`쪽 수가 서버 변환(${info.pageCount})과 많이 다릅니다(${opened.count})`);
    }
    const source = new ArtcraftSource(client, opened.sizes, info.permissions.canCopy);
    await source.warmUp();
    return source;
  } catch (error) {
    client?.terminate();
    throw error instanceof ArtcraftUnavailable ? error : new ArtcraftUnavailable(error instanceof Error ? error.message : String(error));
  }
}
