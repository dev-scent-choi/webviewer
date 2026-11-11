/**
 * 막대 — 확대·맞춤·쪽 이동·회전·전체 화면, Ctrl+휠·두 번 누르기 확대, 끌어서 옮기기, 좁은 화면의 "도구▾" 메뉴.
 * 동작은 지금 보기 화면(viewer.js setupToolbar·setupZoomFit·setupPageJump·setupRotate·setupFullscreen·setupDrawingTools)과 같다.
 */
import type { ShellCtx } from './ctx.js';
import { ZOOM_STEPS, menuIdsFor } from './logic.js';

const $ = <T extends HTMLElement = HTMLElement>(id: string): T => document.getElementById(id) as T;

export function setupToolbar(ctx: ShellCtx): void {
  setupZoom(ctx);
  setupPageNav(ctx);
  $('rotate-btn').addEventListener('click', () => ctx.body.rotate(90));
  setupFullscreen();
  setupPointer(ctx);
  setupToolsMenu();
}

function setupZoom(ctx: ShellCtx): void {
  const zoomIn = $<HTMLButtonElement>('zoom-in');
  const zoomOut = $<HTMLButtonElement>('zoom-out');
  const level = $('zoom-level');
  const sync = (percent: number) => {
    level.textContent = `${percent}%`;
    zoomOut.disabled = ctx.body.zoomIndex === 0;
    zoomIn.disabled = ctx.body.zoomIndex === ZOOM_STEPS.length - 1;
  };
  ctx.body.onZoom(sync);
  sync(ctx.body.zoomPercent);
  zoomIn.addEventListener('click', () => ctx.body.setZoomIndex(ctx.body.zoomIndex + 1));
  zoomOut.addEventListener('click', () => ctx.body.setZoomIndex(ctx.body.zoomIndex - 1));

  // 맞춤 토글 — 단추는 "다음에 누르면 적용될" 모드를 보인다(viewer.js와 같다)
  const fit = $<HTMLButtonElement>('zoom-fit-btn');
  let next: 'fit-width' | 'fit-page' = 'fit-width';
  fit.addEventListener('click', () => {
    ctx.body.zoom(next);
    next = next === 'fit-width' ? 'fit-page' : 'fit-width';
    const pageNext = next === 'fit-page';
    // svg 요소는 hidden 프로퍼티가 속성에 반영되지 않는 엔진이 있어 속성을 직접 바꾼다
    fit.querySelector('.zoom-fit-icon-width')?.toggleAttribute('hidden', pageNext);
    fit.querySelector('.zoom-fit-icon-page')?.toggleAttribute('hidden', !pageNext);
    const label = pageNext ? '페이지에 맞추기' : '너비에 맞추기';
    fit.title = label;
    fit.setAttribute('aria-label', label);
    fit.setAttribute('aria-pressed', 'true');
    setTimeout(() => fit.setAttribute('aria-pressed', 'false'), 400);
  });
}

function setupPageNav(ctx: ShellCtx): void {
  const input = $<HTMLInputElement>('page-jump-input');
  const total = $('page-jump-total');
  const prev = $<HTMLButtonElement>('page-prev');
  const next = $<HTMLButtonElement>('page-next');
  const show = (page: number) => {
    if (document.activeElement !== input) input.value = String(page);
    total.textContent = `의 ${ctx.body.pageCount()}`;
    prev.disabled = page <= 1;
    next.disabled = page >= ctx.body.pageCount();
  };
  ctx.body.onPage(show);
  ctx.body.onSourceChange(() => show(ctx.body.currentPage));
  show(ctx.body.currentPage);
  const jump = () => {
    const typed = Math.round(Number(input.value));
    if (Number.isFinite(typed)) ctx.body.goto(typed);
    input.value = String(ctx.body.currentPage);
  };
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      jump();
      input.blur();
    } else if (event.key === 'Escape') {
      input.blur();
    }
  });
  input.addEventListener('blur', jump);
  input.addEventListener('focus', () => input.select());
  prev.addEventListener('click', () => ctx.body.goto(ctx.body.currentPage - 1));
  next.addEventListener('click', () => ctx.body.goto(ctx.body.currentPage + 1));
}

/** 전체 화면 — Esc·브라우저 메뉴로 나가는 경우도 fullscreenchange로 맞춘다(viewer.js setupFullscreen) */
function setupFullscreen(): void {
  const btn = $<HTMLButtonElement>('fullscreen-btn');
  if (!document.fullscreenEnabled) {
    btn.disabled = true;
    btn.title = '이 브라우저에서는 전체 화면을 쓸 수 없습니다';
    return;
  }
  const sync = () => {
    const on = Boolean(document.fullscreenElement);
    btn.querySelector('.fullscreen-icon-enter')?.toggleAttribute('hidden', on);
    btn.querySelector('.fullscreen-icon-exit')?.toggleAttribute('hidden', !on);
    const label = on ? '전체화면 종료' : '전체화면 보기';
    btn.title = label;
    btn.setAttribute('aria-label', label);
    btn.setAttribute('aria-pressed', String(on));
  };
  btn.addEventListener('click', () => {
    if (document.fullscreenElement) void document.exitFullscreen();
    else document.documentElement.requestFullscreen().catch(() => undefined);
  });
  document.addEventListener('fullscreenchange', sync);
  sync();
}

/** Ctrl+휠·두 번 누르기 확대, 끌어서 옮기기(viewer.js setupDrawingTools 앞부분) */
function setupPointer(ctx: ShellCtx): void {
  const pages = ctx.body.container;
  pages.addEventListener(
    'wheel',
    (event) => {
      if (!event.ctrlKey) return;
      event.preventDefault();
      ctx.body.zoomAtPoint(ctx.body.zoomIndex + (event.deltaY < 0 ? 1 : -1), event.clientX, event.clientY);
    },
    { passive: false },
  );
  pages.addEventListener('dblclick', (event) => {
    if (pages.classList.contains('measuring')) return;
    ctx.body.zoomAtPoint(ctx.body.zoomIndex + 2, event.clientX, event.clientY);
  });
  let pan: { x: number; y: number; left: number; top: number } | null = null;
  pages.addEventListener('mousedown', (event) => {
    if (pages.classList.contains('measuring')) return;
    if (event.button !== 0 && event.button !== 1) return;
    const target = event.target as Element;
    // 글자 선택이 되는 화면에서 글자 위를 누르면 끌기 대신 글자 고르기
    if (ctx.body.textSelectable && target.closest('.pdf-text-layer span')) return;
    // 칸 스크롤바를 누른 것은 브라우저에 맡긴다
    if (target === pages && (event.offsetX > pages.clientWidth || event.offsetY > pages.clientHeight)) return;
    pan = { x: event.clientX, y: event.clientY, left: pages.scrollLeft, top: window.scrollY };
    pages.classList.add('panning');
    event.preventDefault();
  });
  document.addEventListener('mousemove', (event) => {
    if (!pan) return;
    pages.scrollLeft = pan.left - (event.clientX - pan.x);
    window.scrollTo(window.scrollX, pan.top - (event.clientY - pan.y));
  });
  document.addEventListener('mouseup', () => {
    pan = null;
    pages.classList.remove('panning');
  });
}

/**
 * "도구▾" — 화면 폭에 따라 정해진 단추(logic.ts menuIdsFor)를 메뉴로 옮기고, 넓어지면 제자리로 돌린다. 요소를 그대로
 * 옮기므로 걸어 둔 클릭 처리가 그대로 산다(viewer.js toolbar-more와 같은 방법).
 */
function setupToolsMenu(): void {
  const btn = $<HTMLButtonElement>('tools-btn');
  const menu = $('tools-menu');
  const home = new Map<string, { parent: Node; next: Node | null }>();
  const close = () => {
    menu.hidden = true;
    btn.setAttribute('aria-expanded', 'false');
  };
  const layout = () => {
    const wanted = new Set(menuIdsFor(window.innerWidth));
    // 메뉴에 있을 필요가 없어진 것부터 제자리로 — 뒤에서부터 넣어 기억해 둔 다음 형제가 먼저 돌아와 있게
    for (const [id, place] of [...home.entries()].reverse()) {
      if (wanted.has(id)) continue;
      const el = document.getElementById(id);
      // 기억해 둔 다음 형제가 아직 메뉴에 있으면 끝에 붙인다(insertBefore가 던지지 않게)
      if (el) place.parent.insertBefore(el, place.next?.parentNode === place.parent ? place.next : null);
      home.delete(id);
    }
    for (const id of wanted) {
      const el = document.getElementById(id);
      if (!el || home.has(id)) continue;
      home.set(id, { parent: el.parentNode as Node, next: el.nextSibling });
      menu.append(el);
    }
    btn.hidden = wanted.size === 0;
    if (wanted.size === 0) close();
  };
  btn.addEventListener('click', (event) => {
    event.stopPropagation();
    menu.hidden = !menu.hidden;
    btn.setAttribute('aria-expanded', String(!menu.hidden));
  });
  // 메뉴 안 단추를 누르면 그 동작을 한 뒤 닫는다
  menu.addEventListener('click', (event) => {
    if ((event.target as Element).closest('button')) close();
  });
  document.addEventListener('click', (event) => {
    if (!menu.hidden && !(event.target as Element).closest('#tools-wrap')) close();
  });
  window.addEventListener('resize', layout);
  layout();
}
