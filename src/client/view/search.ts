/**
 * 문서 안 찾기 — PNG 모드는 서버(/api/view/search), pdf.js 모드는 브라우저(findText.ts)가 찾고, 강조·이동은 같다
 * (viewer.js runSearch·gotoHit·gotoPageHit·centerOnElement 그대로). Ctrl+F·F3·Shift+F3·Enter·Shift+Enter·Esc.
 */
import type { ShellCtx } from './ctx.js';
import { ZOOM_STEPS, hitCountLabel, wrapIndex } from './logic.js';
import type { SearchHit } from './types.js';

/** 강조가 이보다 작으면(도면처럼 아주 작은 글자) 읽을 수 있는 배율까지 확대한다 */
const MIN_READABLE_HIT_PX = 11;
const TARGET_HIT_PX = 16;

const $ = <T extends HTMLElement = HTMLElement>(id: string): T => document.getElementById(id) as T;

/** 요소를 화면 가운데로 — 쪽 칸은 가로, 창은 세로로 스크롤한다(viewer.js centerOnElement) */
function centerOn(container: HTMLElement, el: HTMLElement): void {
  const rect = el.getBoundingClientRect();
  const box = container.getBoundingClientRect();
  container.scrollLeft += rect.left + rect.width / 2 - (box.left + container.clientWidth / 2);
  const dy = rect.top + rect.height / 2 - window.innerHeight / 2;
  const before = container.scrollTop;
  container.scrollTop += dy;
  const consumed = container.scrollTop - before;
  if (Math.abs(dy - consumed) > 1) window.scrollBy(0, dy - consumed);
}

export function setupSearch(ctx: ShellCtx): void {
  const bar = $('search-bar');
  const input = $<HTMLInputElement>('search-input');
  const count = $('search-count');
  const prev = $<HTMLButtonElement>('search-prev');
  const next = $<HTMLButtonElement>('search-next');
  let hits: SearchHit[] = [];
  let current = -1;
  let lastQuery = '';
  let running: AbortController | null = null;

  const setCount = (text: string) => {
    count.textContent = text;
    prev.disabled = hits.length === 0;
    next.disabled = hits.length === 0;
  };
  const open = () => {
    bar.hidden = false;
    input.focus();
    input.select();
  };
  const close = () => {
    bar.hidden = true;
    running?.abort();
    ctx.body.clearHits();
    hits = [];
    current = -1;
    lastQuery = '';
    setCount('');
  };
  const show = (index: number) => {
    if (hits.length === 0) return;
    current = wrapIndex(index, hits.length);
    setCount(hitCountLabel(current, hits.length));
    const el = ctx.body.showHits(hits, current);
    if (el) {
      const height = el.getBoundingClientRect().height;
      if (height > 0 && height < MIN_READABLE_HIT_PX) {
        const target = Math.min(ZOOM_STEPS[ZOOM_STEPS.length - 1], ctx.body.zoomPercent * (TARGET_HIT_PX / height));
        const found = ZOOM_STEPS.findIndex((step) => step >= target);
        const resolved = found === -1 ? ZOOM_STEPS.length - 1 : found;
        if (resolved > ctx.body.zoomIndex) ctx.body.setZoomIndex(resolved);
      }
    }
    const target = el ?? ctx.body.pageFrame(hits[current].page);
    if (target) centerOn(ctx.body.container, target);
  };
  const run = async () => {
    const query = input.value.trim();
    if (query === '') return;
    if (query === lastQuery) {
      if (hits.length > 0) show(current + 1);
      return;
    }
    lastQuery = query;
    ctx.body.clearHits();
    hits = [];
    setCount('검색 중…');
    running?.abort();
    const ac = new AbortController();
    running = ac;
    try {
      const found = await ctx.body.search(query, ac.signal);
      if (ac.signal.aborted) return;
      hits = found;
      current = -1;
      if (hits.length === 0) setCount('결과 없음');
      else show(0);
    } catch (error) {
      if ((error as { name?: string } | null)?.name === 'AbortError') return;
      lastQuery = '';
      setCount(error instanceof Error && error.message ? error.message : '검색 실패');
    }
  };

  $('search-btn').addEventListener('click', open);
  $('search-close').addEventListener('click', close);
  prev.addEventListener('click', () => show(current - 1));
  next.addEventListener('click', () => show(current + 1));
  input.addEventListener('keydown', (event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    if (event.shiftKey) show(current - 1);
    else void run();
  });
  document.addEventListener('keydown', (event) => {
    const key = event.key.toLowerCase();
    // 브라우저 기본 Ctrl+F는 그리지 않은 쪽을 찾지 못한다 — 자체 찾기로
    if ((event.ctrlKey || event.metaKey) && key === 'f') {
      event.preventDefault();
      open();
    } else if (key === 'f3') {
      event.preventDefault();
      if (hits.length > 0) show(current + (event.shiftKey ? -1 : 1));
      else open();
    }
  });
  ctx.onEscape(90, () => {
    if (bar.hidden) return false;
    close();
    return true;
  });
  // 원천이 바뀌면(pdf.js → PNG) 쪽 틀이 새로 만들어져 강조가 사라진다 — 찾기도 처음부터
  ctx.body.onSourceChange(close);
}
