/**
 * 썸네일 칸(설계 §3 "새로") — 쪽마다 작은 그림, 누르면 그 쪽으로. 칸에 보이는 것만 그린다(PNG: 서버 ?thumb=true,
 * pdf.js: 그 자리에서 폭 120px). 640px 미만은 닫힌 채 시작하고, 닫혀 있는 동안은 아무것도 받지 않는다.
 */
import type { ShellCtx } from './ctx.js';
import { thumbsOpenAtStart } from './logic.js';

const THUMB_WIDTH = 120;

export function setupThumbs(ctx: ShellCtx): void {
  const btn = document.getElementById('thumbs-btn') as HTMLButtonElement;
  const panel = document.getElementById('thumbs-panel') as HTMLElement;
  let items: HTMLButtonElement[] = [];
  let marked: HTMLButtonElement | null = null;
  let observer: IntersectionObserver | null = null;

  const mark = (page: number) => {
    marked?.classList.remove('current');
    marked = items[page - 1] ?? null;
    marked?.classList.add('current');
    if (marked && !panel.hidden) marked.scrollIntoView({ block: 'nearest' });
  };

  const setOpen = (open: boolean) => {
    panel.hidden = !open;
    document.body.classList.toggle('thumbs-open', open);
    btn.setAttribute('aria-pressed', String(open));
    if (open) mark(ctx.body.currentPage);
  };

  const build = () => {
    observer?.disconnect();
    // 칸 안에서 보이는(또는 곧 보일) 것만 그린다 — 숨겨진 칸에서는 교차가 없어 요청도 없다
    observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          const item = entry.target as HTMLButtonElement;
          observer?.unobserve(item);
          const box = item.querySelector('.thumb-box') as HTMLElement;
          ctx.body
            .thumbnail(Number(item.dataset.page), THUMB_WIDTH)
            .then((picture) => box.replaceChildren(picture))
            .catch(() => box.classList.add('thumb-error'));
        }
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
  ctx.body.onPage(mark);
  ctx.body.onSourceChange(build);
  build();
  setOpen(thumbsOpenAtStart(window.innerWidth));
}
