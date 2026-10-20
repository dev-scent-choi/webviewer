/**
 * OCR(8단계) — 글자가 그림으로만 있는 쪽(스캔본)을 이 브라우저에서 tesseract.js로 읽는다(한국어+영어).
 * 쪽을 pdf.js로 쪽 그림의 원래 해상도(없으면 약 216dpi)로 그려 인식하고, 낱말마다 기준선 시작·끝·높이를 PDF 좌표로 바꿔 방의 pdfOcr에 쓴다
 * (실행 취소 대상). 화면에는 바로 투명한 낱말 층(찾기·고르기·복사)으로, 저장하면 서버가 보이지 않는 글자 층으로 쓴다
 * (src/pdf/ocrLayer.ts). worker·wasm·학습 데이터는 이 서버의 /ocr/에서만 받는다 — 문서 그림이 밖으로 나가지 않는다.
 */
import type * as Y from 'yjs';
import type { OcrWord, PdfOcrPage } from '../../pdf/annotTypes.js';
import { LOCAL_ORIGIN } from './annotLayer.js';
import type { PageView } from './pageView.js';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
/** 그림이 없는 쪽의 인식 배율(72dpi × 3 ≈ 216dpi). 스캔본은 쪽 그림의 원래 해상도로 그린다(pageView.renderForOcr) */
const OCR_SCALE = 3;
/** 이보다 확신이 낮은 낱말은 버린다(그림의 무늬를 글자로 읽은 것 등) */
const MIN_CONFIDENCE = 45;

interface TessWord {
  text: string;
  confidence: number;
  bbox: { x0: number; y0: number; x1: number; y1: number };
}
interface TessLine {
  text: string;
  words: TessWord[];
}
interface TessWorker {
  recognize(image: HTMLCanvasElement, opts?: object, output?: object): Promise<{ data: { blocks?: Array<{ paragraphs: Array<{ lines: TessLine[] }> }> | null } }>;
  terminate(): Promise<unknown>;
}

/**
 * 줄의 띄어쓰기대로 낱말을 다시 만든다 — tesseract는 줄 글자(line.text)의 띄어쓰기는 맞게 내면서 낱말 목록은 한글을
 * 음절마다 쪼개 낸다("가","동","률","은", 실측). 줄 글자를 띄어쓰기로 나눈 각 낱말에 조각을 차례로 이어 붙이고 상자를 합친다.
 * 조각을 이어 붙인 글자가 줄 글자와 맞지 않으면(드문 경우) 그 줄은 조각을 그대로 쓴다.
 */
export function wordsByLineText(line: TessLine): TessWord[] {
  const pieces = line.words.filter((w) => w.text.trim());
  const tokens = line.text.trim().split(/\s+/).filter(Boolean);
  const out: TessWord[] = [];
  let i = 0;
  for (const token of tokens) {
    let text = '';
    let merged: TessWord | null = null;
    while (i < pieces.length && text.length < token.length) {
      const w = pieces[i];
      text += w.text.trim();
      merged = merged
        ? {
            text,
            confidence: Math.min(merged.confidence, w.confidence),
            bbox: { x0: Math.min(merged.bbox.x0, w.bbox.x0), y0: Math.min(merged.bbox.y0, w.bbox.y0), x1: Math.max(merged.bbox.x1, w.bbox.x1), y1: Math.max(merged.bbox.y1, w.bbox.y1) },
          }
        : { ...w, text };
      i += 1;
    }
    if (!merged || text !== token) return pieces.map((w) => ({ ...w, text: w.text.trim() }));
    out.push(merged);
  }
  return i === pieces.length ? out : pieces.map((w) => ({ ...w, text: w.text.trim() }));
}

export interface OcrUiOptions {
  readonly view: PageView;
  readonly ydoc: Y.Doc;
  readonly map: Y.Map<PdfOcrPage>;
  readonly author: string;
  readonly notice: (text: string) => void;
}

export interface OcrUi {
  open(): void;
}

export function setupOcr(opts: OcrUiOptions): OcrUi {
  const { view, map } = opts;
  const dialog = $('ocr-dialog');
  const status = $('ocr-status');
  const run = $<HTMLButtonElement>('ocr-run');
  const stop = $<HTMLButtonElement>('ocr-stop');
  const clear = $<HTMLButtonElement>('ocr-clear');
  let worker: TessWorker | null = null;
  let running = false;
  let cancelled = false;

  for (const b of dialog.querySelectorAll<HTMLButtonElement>('[data-close]')) b.addEventListener('click', () => !running && (dialog.hidden = true));
  dialog.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !running) dialog.hidden = true;
  });

  async function getWorker(): Promise<TessWorker> {
    if (worker) return worker;
    status.textContent = '글자 인식기를 준비하는 중…(처음 한 번은 몇 초 걸립니다)';
    // tesseract.js는 CommonJS — 동적 import로 받으면 함수가 default 아래에 있다
    const mod = (await import('tesseract.js')) as unknown as { createWorker?: typeof import('tesseract.js').createWorker; default?: { createWorker: typeof import('tesseract.js').createWorker } };
    const createWorker = mod.createWorker ?? mod.default!.createWorker;
    worker = (await createWorker(['kor', 'eng'], 1, {
      workerPath: '/ocr/worker.min.js',
      corePath: '/ocr/',
      langPath: '/ocr/lang',
      gzip: true,
      logger: () => undefined,
    })) as unknown as TessWorker;
    return worker;
  }

  /** 대상 쪽 번호들 */
  async function targets(scope: string): Promise<number[]> {
    const all = Array.from({ length: view.pageCount }, (_, i) => i + 1);
    if (scope === 'current') return [view.currentPage];
    if (scope === 'all') return all;
    // 글자가 없는 쪽만 — pdf.js 글 조각이 비어 있는 쪽(이미 인식한 쪽은 빼고)
    const out: number[] = [];
    for (const p of all) {
      if (map.has(view.pageKeyOf(p))) continue;
      const items = await view.textItemInfos(p);
      if (!items.some((i) => i.str.trim())) out.push(p);
    }
    return out;
  }

  async function recognizePage(page: number): Promise<number> {
    const rendered = await view.renderForOcr(page, OCR_SCALE);
    if (!rendered) return 0;
    const w = await getWorker();
    const { data } = await w.recognize(rendered.canvas, {}, { blocks: true });
    rendered.canvas.width = 0;
    const vp = rendered.viewport;
    const words: OcrWord[] = [];
    for (const block of data.blocks ?? []) {
      for (const para of block.paragraphs) {
        for (const line of para.lines) {
          for (const tw of wordsByLineText(line)) {
            const text = tw.text.trim();
            if (!text || tw.confidence < MIN_CONFIDENCE) continue;
            const { x0, y0, x1, y1 } = tw.bbox;
            // 기준선은 상자 아래쪽에서 조금 위(아래로 내려가는 글자 몫)
            const base = y1 - (y1 - y0) * 0.18;
            const [x, y] = vp.convertToPdfPoint(x0, base) as number[];
            const [ex, ey] = vp.convertToPdfPoint(x1, base) as number[];
            const [tx, ty] = vp.convertToPdfPoint(x0, y0) as number[];
            const h = Math.hypot(tx - x, ty - y);
            if (h < 1) continue;
            words.push({ t: text, x, y, ex, ey, h });
          }
        }
      }
    }
    const key = view.pageKeyOf(page);
    opts.ydoc.transact(() => {
      if (words.length) map.set(key, { words, author: opts.author, at: new Date().toISOString() });
      else map.delete(key);
    }, LOCAL_ORIGIN);
    return words.length;
  }

  run.addEventListener('click', async () => {
    if (running) return;
    const scope = dialog.querySelector<HTMLInputElement>('input[name="ocr-scope"]:checked')?.value ?? 'empty';
    running = true;
    cancelled = false;
    run.disabled = true;
    clear.disabled = true;
    stop.hidden = false;
    let total = 0;
    let done = 0;
    try {
      const pages = await targets(scope);
      if (pages.length === 0) {
        status.textContent = '인식할 쪽이 없습니다(모든 쪽에 이미 글자가 있습니다). ‘모든 쪽’을 고르면 다시 읽습니다.';
        return;
      }
      for (const p of pages) {
        if (cancelled) break;
        status.textContent = `글자를 읽는 중… ${done + 1}/${pages.length}쪽 (${p}쪽)`;
        total += await recognizePage(p);
        done += 1;
      }
      status.textContent = cancelled ? `멈췄습니다 — ${done}쪽에서 낱말 ${total}개를 읽었습니다.` : `${done}쪽에서 낱말 ${total}개를 읽었습니다. 이제 찾기·글자 고르기가 됩니다.`;
      if (!cancelled) opts.notice(`글자 인식 끝 — 낱말 ${total}개(저장하면 파일에도 들어갑니다)`);
    } catch (error) {
      console.error(error);
      status.textContent = `글자를 읽지 못했습니다: ${error instanceof Error ? error.message : String(error)}`;
    } finally {
      running = false;
      run.disabled = false;
      clear.disabled = false;
      stop.hidden = true;
    }
  });
  stop.addEventListener('click', () => {
    cancelled = true;
    status.textContent = '이 쪽까지 읽고 멈춥니다…';
  });
  clear.addEventListener('click', () => {
    const key = view.pageKeyOf(view.currentPage);
    if (!map.has(key)) {
      status.textContent = '지금 쪽에는 인식한 글자가 없습니다.';
      return;
    }
    opts.ydoc.transact(() => map.delete(key), LOCAL_ORIGIN);
    status.textContent = `${view.currentPage}쪽의 인식 결과를 지웠습니다.`;
  });

  return {
    open() {
      status.textContent = '';
      stop.hidden = true;
      dialog.hidden = false;
    },
  };
}
