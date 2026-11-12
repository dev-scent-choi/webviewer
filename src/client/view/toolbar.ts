/**
 * 막대 — 확대·맞춤·쪽 이동·회전·전체 화면, Ctrl+휠·두 번 누르기 확대, 끌어서 옮기기, 좁은 화면의 "도구▾" 메뉴.
 * 동작은 지금 보기 화면(viewer.js setupToolbar·setupZoomFit·setupPageJump·setupRotate·setupFullscreen·setupDrawingTools)과 같다.
 */
import type { ShellCtx } from './ctx.js';
import { ZOOM_STEPS, menuIdsFor } from './logic.js';

const $ = <T extends HTMLElement = HTMLElement>(id: string): T => document.getElementById(id) as T;

/** 반환값은 해제 함수(떠날 때 document·window 처리를 걷는다) */
export function setupToolbar(ctx: ShellCtx): () => void {
  const off = new AbortController();
  setupZoom(ctx);
  setupPageNav(ctx);
  $('rotate-btn').addEventListener('click', () => ctx.body.rotate(90));
  setupFullscreen(off.signal);
  setupPointer(ctx, off.signal);
  setupToolsMenu(ctx, off.signal);
  return () => off.abort();
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
function setupFullscreen(signal: AbortSignal): void {
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
  document.addEventListener('fullscreenchange', sync, { signal });
  sync();
}

/** Ctrl+휠·두 번 누르기 확대, 끌어서 옮기기(viewer.js setupDrawingTools 앞부분) */
function setupPointer(ctx: ShellCtx, signal: AbortSignal): void {
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
    // 글자 선택이 되는 화면에서는 두 번 누르기가 낱말 고르기다 — 확대하지 않는다
    if (ctx.body.textSelectable && (event.target as Element).closest('.pdf-text-layer')) return;
    ctx.body.zoomAtPoint(ctx.body.zoomIndex + 2, event.clientX, event.clientY);
  });
  let pan: { x: number; y: number; left: number; top: number } | null = null;
  // 글자 선택이 되는 화면에서는 쪽 위 끌기가 글자 고르기다 — 끌어 옮기는 법을 쪽 바깥 여백의 툴팁으로 알린다(쪽 위는 빈 title로 가린다)
  const panHint = () => {
    const on = ctx.body.textSelectable;
    pages.title = on ? '쪽 바깥 여백을 끌거나 가운데 단추로 끌면 화면이 움직입니다 — 쪽 위를 끌면 글자를 고릅니다' : '';
    for (const frame of pages.querySelectorAll('.page-frame')) frame.toggleAttribute('title', on);
  };
  ctx.body.onSourceChange(panHint);
  panHint();
  pages.addEventListener('mousedown', (event) => {
    // 측정·강조 사각형 그리기 중에는 끌어도 화면을 옮기지 않는다(사각형 높이가 0이 되지 않게)
    if (pages.classList.contains('measuring') || pages.classList.contains('markup-drawing')) return;
    if (event.button !== 0 && event.button !== 1) return;
    const target = event.target as Element;
    // 쓰던 글 상자·코멘트 창 안(글자 크기 단추·입력 칸)은 끌기가 아니다 — 아래 blur가 쓰던 글을 저장해 버린다
    if (target.closest('.text-insert-wrapper, .markup-popover')) return;
    // 글자 선택이 되는 화면에서 글자 위를 누르면 끌기 대신 글자 고르기
    // 쪽 위 왼쪽 단추는 빈 곳에서도 글자 고르기를 시작해야 하니 끌기는 쪽 밖(여백)이나 가운데 단추로만
    if (ctx.body.textSelectable && event.button === 0 && target.closest('.page-frame')) return;
    // 칸 스크롤바를 누른 것은 브라우저에 맡긴다
    if (target === pages && (event.offsetX > pages.clientWidth || event.offsetY > pages.clientHeight)) return;
    (document.activeElement as HTMLElement | null)?.blur?.(); // 쪽 번호 입력 중이면 blur로 이동을 확정
    pan = { x: event.clientX, y: event.clientY, left: pages.scrollLeft, top: window.scrollY };
    pages.classList.add('panning');
    event.preventDefault();
  });
  document.addEventListener('mousemove', (event) => {
    if (!pan) return;
    pages.scrollLeft = pan.left - (event.clientX - pan.x);
    window.scrollTo(window.scrollX, pan.top - (event.clientY - pan.y));
  }, { signal });
  document.addEventListener('mouseup', () => {
    pan = null;
    pages.classList.remove('panning');
  }, { signal });
}

/**
 * "도구▾" — 화면 폭에 따라 정해진 단추(logic.ts menuIdsFor)를 메뉴로 옮기고, 넓어지면 제자리로 돌린다. 요소를 그대로
 * 옮기므로 걸어 둔 클릭 처리가 그대로 산다(viewer.js toolbar-more와 같은 방법).
 */
function setupToolsMenu(ctx: ShellCtx, signal: AbortSignal): void {
  const btn = $<HTMLButtonElement>('tools-btn');
  const menu = $('tools-menu');
  const home = new Map<string, { parent: Node; next: Node | null }>();
  const close = () => {
    menu.hidden = true;
    btn.setAttribute('aria-expanded', 'false');
  };
  const items = () => [...menu.querySelectorAll<HTMLElement>('button:not(:disabled), input:not(:disabled)')];
  // 메뉴 안 요소는 menuitem, 제자리로 돌아가면 역할을 걷는다
  // 체크 상자를 감싼 label은 감싼 채로 두고 상자에 menuitemcheckbox를 준다(메뉴 항목 안에 상호작용 요소를 넣지 않게)
  const setRole = (el: HTMLElement, inMenu: boolean) => {
    const box = el.tagName === 'LABEL' ? el.querySelector('input') : null;
    const target = box ?? el;
    if (inMenu) target.setAttribute('role', box ? 'menuitemcheckbox' : 'menuitem');
    else target.removeAttribute('role');
  };
  const layout = () => {
    const wanted = new Set(menuIdsFor(window.innerWidth));
    // 메뉴에 있을 필요가 없어진 것부터 제자리로 — 뒤에서부터 넣어 기억해 둔 다음 형제가 먼저 돌아와 있게
    for (const [id, place] of [...home.entries()].reverse()) {
      if (wanted.has(id)) continue;
      const el = document.getElementById(id);
      // 기억해 둔 다음 형제가 아직 메뉴에 있으면 끝에 붙인다(insertBefore가 던지지 않게)
      if (el) setRole(el, false);
      if (el) place.parent.insertBefore(el, place.next?.parentNode === place.parent ? place.next : null);
      home.delete(id);
    }
    for (const id of wanted) {
      const el = document.getElementById(id);
      if (!el || home.has(id)) continue;
      home.set(id, { parent: el.parentNode as Node, next: el.nextSibling });
      setRole(el, true);
      menu.append(el);
    }
    btn.hidden = wanted.size === 0;
    if (wanted.size === 0) close();
  };
  btn.addEventListener('click', (event) => {
    event.stopPropagation();
    menu.hidden = !menu.hidden;
    btn.setAttribute('aria-expanded', String(!menu.hidden));
    if (!menu.hidden && event.detail === 0) items()[0]?.focus(); // 키보드로 열면 첫 항목으로
  });
  menu.addEventListener('keydown', (event) => {
    if (event.key !== 'ArrowDown' && event.key !== 'ArrowUp') return;
    const list = items();
    if (list.length === 0) return;
    event.preventDefault();
    const at = list.indexOf(document.activeElement as HTMLElement);
    list[(at + (event.key === 'ArrowDown' ? 1 : -1) + list.length) % list.length].focus();
  });
  // Esc — 메뉴를 닫고 도구▾로 돌아간다(우선순위 25: 쓰던 글 20 다음)
  ctx.onEscape(25, () => {
    if (menu.hidden) return false;
    close();
    btn.focus();
    return true;
  });
  // 메뉴 안 단추를 누르면 그 동작을 한 뒤 닫는다
  menu.addEventListener('click', (event) => {
    if (!(event.target as Element).closest('button')) return;
    close();
    // 숨긴 메뉴 안에 남은 초점은 도구▾로 — 항목이 다른 곳(찾기 칸 등)으로 옮겼으면 그대로 둔다
    setTimeout(() => {
      const active = document.activeElement;
      if (!active || active === document.body || menu.contains(active)) btn.focus();
    });
  });
  document.addEventListener('click', (event) => {
    if (!menu.hidden && !(event.target as Element).closest('#tools-wrap')) close();
  }, { signal });
  window.addEventListener('resize', layout, { signal });
  layout();
}
