/**
 * 썸네일 칸(설계 §3 "새로") — 쪽마다 작은 그림, 누르면 그 쪽으로. 칸에 보이는 것만 그린다(PNG: 서버 ?thumb=true,
 * pdf.js: 그 자리에서 폭 120px). 640px 미만은 닫힌 채 시작하고, 닫혀 있는 동안은 아무것도 받지 않는다.
 */
import type { ShellCtx } from './ctx.js';
import { thumbsOpenAtStart, thumbWindow } from './logic.js';

const THUMB_WIDTH = 120;
/** 보이는 쪽 앞뒤로 이만큼까지만 그려 둔다 */
const KEEP_MARGIN = 10;

/** 반환값은 해제 함수(떠날 때) */
export function setupThumbs(ctx: ShellCtx): () => void {
  const btn = document.getElementById('thumbs-btn') as HTMLButtonElement;
  const panel = document.getElementById('thumbs-panel') as HTMLElement;
  let items: HTMLButtonElement[] = [];
  let marked: HTMLButtonElement | null = null;
  let observer: IntersectionObserver | null = null;
  const visible = new Set<number>();
  const drawn = new Set<number>();

  const mark = (page: number) => {
    marked?.classList.remove('current');
    marked?.removeAttribute('aria-current');
    marked = items[page - 1] ?? null;
    marked?.classList.add('current');
    marked?.setAttribute('aria-current', 'page');
    if (marked && !panel.hidden) marked.scrollIntoView({ block: 'nearest' });
  };

  const setOpen = (open: boolean) => {
    panel.hidden = !open;
    document.body.classList.toggle('thumbs-open', open);
    btn.setAttribute('aria-pressed', String(open));
    if (open) mark(ctx.body.currentPage);
  };

  const draw = (page: number) => {
    const box = items[page - 1]?.querySelector('.thumb-box') as HTMLElement | undefined;
    if (!box || drawn.has(page)) return;
    drawn.add(page);
    ctx.body
      .thumbnail(page, THUMB_WIDTH)
      .then((picture) => {
        // 그리는 사이 창 밖이 됐으면 버린다
        if (drawn.has(page) && items[page - 1]?.querySelector('.thumb-box') === box) box.replaceChildren(picture);
      })
      .catch(() => {
        drawn.delete(page);
        box.classList.add('thumb-error');
      });
  };

  // 보이는 것은 그리고, 창(보이는 쪽 ±KEEP_MARGIN) 밖으로 나간 것은 비운다 — 되돌아오면 다시 그린다
  const sweep = () => {
    const range = thumbWindow(visible, KEEP_MARGIN, items.length);
    for (const page of [...drawn]) {
      if (range && page >= range[0] && page <= range[1]) continue;
      drawn.delete(page);
      items[page - 1]?.querySelector('.thumb-box')?.replaceChildren();
    }
    for (const page of visible) draw(page);
  };

  const build = () => {
    observer?.disconnect();
    visible.clear();
    drawn.clear();
    // 칸 안에서 보이는(또는 곧 보일) 것만 그린다 — 숨겨진 칸에서는 교차가 없어 요청도 없다
    observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          const page = Number((entry.target as HTMLElement).dataset.page);
          if (entry.isIntersecting) visible.add(page);
          else visible.delete(page);
        }
        sweep();
      },
      { root: panel, rootMargin: '200px 0px' },
    );
    items = [];
    marked = null;
    for (let page = 1; page <= ctx.body.pageCount(); page += 1) {
      const item = document.createElement('button');
      item.type = 'button';
      item.className = 'thumb-item';
      item.dataset.page = String(page);
      item.setAttribute('aria-label', `${page}쪽으로 이동`);
      const size = ctx.body.pageSize(page);
      const box = document.createElement('span');
      box.className = 'thumb-box';
      box.style.aspectRatio = `${size.w} / ${size.h}`;
      const label = document.createElement('span');
      label.className = 'thumb-label';
      label.textContent = String(page);
      item.append(box, label);
      item.addEventListener('click', () => {
        ctx.body.goto(page);
        // 휴대폰에서는 고른 뒤 칸을 닫아 쪽을 넓게 보인다
        if (window.innerWidth < 640) setOpen(false);
      });
      items.push(item);
      observer.observe(item);
    }
    panel.replaceChildren(...items);
    mark(ctx.body.currentPage);
  };

  btn.addEventListener('click', () => setOpen(panel.hidden));
  // 좁은 화면에서 칸이 쪽 위에 겹쳐 있을 때 Esc로 닫고 단추로 돌아간다(찾기 90 앞)
  ctx.onEscape(80, () => {
    if (panel.hidden || window.innerWidth >= 640) return false;
    setOpen(false);
    // 단추가 도구▾ 메뉴 안에 있으면 보이지 않으니 메뉴 단추로
    (btn.offsetParent ? btn : (document.getElementById('tools-btn') ?? btn)).focus();
    return true;
  });
  ctx.body.onPage(mark);
  ctx.body.onSourceChange(build);
  build();
  setOpen(thumbsOpenAtStart(window.innerWidth));
  return () => observer?.disconnect();
}
