/**
 * 오피스 리본 구조(설계 2026-11-22-office-ribbon-structure)의 브라우저 쪽 두 가지 — 마크업은 scripts/gen-ribbon.mjs가 정적으로 만든다.
 *  1) 준비 중·대리·접힌 묶음 단추의 아이콘(data-icon)을 고른 탭에서만 단다(스프라이트 /ribbon/icons.svg 하나, 탭마다 늦게).
 *  2) 좁으면 오피스처럼 줄인다 — data-collapse 작은 수부터(같으면 오른쪽부터) 큰 단추를 작은 단추 세 줄로(rb-compact), 그래도 넘치면
 *     묶음을 큰 단추 하나로 접고(rb-collapsed), 남는 자리가 생기면 다시 편다. 600px 이하는 모두 접는다.
 *     접힌 묶음 단추를 누르면 원래 단추들(.ribbon-group-buttons)이 그 아래 떠 있는 판으로 보인다(DOM은 그대로 — id·처리기 유지).
 */
export const ICON_SPRITE = '/ribbon/icons.svg';
export const NARROW_PX = 600;
const SVG_NS = 'http://www.w3.org/2000/svg';

export function collapseOrder(groups: readonly { priority: number; index: number }[]): number[] {
  return [...groups].sort((a, b) => a.priority - b.priority || b.index - a.index).map((g) => g.index);
}

export function hydrateIcons(panel: HTMLElement): void {
  for (const el of panel.querySelectorAll<HTMLElement>('[data-icon]')) {
    if (el.querySelector(':scope > svg')) continue;
    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('class', 'rb-ico');
    svg.setAttribute('aria-hidden', 'true');
    const use = document.createElementNS(SVG_NS, 'use');
    use.setAttribute('href', `${ICON_SPRITE}#rbi-${el.dataset.icon ?? ''}`);
    svg.append(use);
    el.prepend(svg);
  }
}

let openGroup: HTMLElement | null = null;
let popSeq = 0;

/** 접힌 묶음 단추에 aria-haspopup·aria-controls(판 id는 한 번만 매긴다) */
function wireToggle(group: HTMLElement): void {
  if (!group.querySelector) return;
  const toggle = group.querySelector<HTMLElement>(':scope > .rb-group-toggle');
  const body = group.querySelector<HTMLElement>(':scope > .ribbon-group-buttons');
  if (!toggle || !body) return;
  if (!body.id) body.id = `rb-pop-${++popSeq}`;
  toggle.setAttribute('aria-haspopup', 'true');
  toggle.setAttribute('aria-controls', body.id);
}

export function closeGroupPopup(): void {
  if (!openGroup) return;
  const toggle = openGroup.querySelector<HTMLElement>(':scope > .rb-group-toggle');
  // 판 안에 있던 포커스가 display:none 으로 BODY로 떨어지지 않게 토글로 되돌린다
  if (toggle && openGroup.contains(document.activeElement)) toggle.focus();
  openGroup.classList.remove('rb-open');
  toggle?.setAttribute('aria-expanded', 'false');
  openGroup = null;
}

let lastFit: { panel: HTMLElement; width: number } | null = null;

export function fitPanel(panel: HTMLElement, viewportWidth: number): void {
  const groups = [...panel.querySelectorAll<HTMLElement>(':scope > .ribbon-group[data-collapse]')];
  // 같은 판·같은 폭으로 다시 맞출 때(상황별 탭·글꼴 받음·폰 키보드 resize)는 열린 묶음 판을 닫지 않는다
  if (lastFit?.panel !== panel || lastFit.width !== viewportWidth) closeGroupPopup();
  lastFit = { panel, width: viewportWidth };
  for (const g of groups) {
    g.classList.remove('rb-collapsed', 'rb-compact');
    wireToggle(g);
  }
  if (viewportWidth <= NARROW_PX) {
    for (const g of groups) g.classList.add('rb-collapsed');
    return;
  }
  const fits = () => panel.scrollWidth <= panel.clientWidth + 1;
  // 묶음에 단계 클래스를 붙여 실제로 좁아질 때만 남긴다 — 큰 단추 하나뿐인 묶음은 작게 해도, 단추 하나짜리 묶음은 접어도 오히려 넓어진다
  const tryShrink = (g: HTMLElement, cls: 'rb-compact' | 'rb-collapsed'): boolean => {
    const before = g.offsetWidth;
    const wasCompact = g.classList.contains('rb-compact');
    if (cls === 'rb-collapsed') g.classList.remove('rb-compact'); // 접은 묶음의 떠 있는 판은 원래 크기로
    g.classList.add(cls);
    if (g.offsetWidth < before) return true;
    g.classList.remove(cls);
    if (wasCompact) g.classList.add('rb-compact');
    return false;
  };
  const order = collapseOrder(groups.map((g, i) => ({ priority: Number(g.dataset.collapse), index: i }))).map((i) => groups[i]);
  // 오피스 순서: 1) 순위대로 큰 단추 → 작은 단추 3줄(rb-compact)  2) 그래도 넘치면 순위대로 묶음을 접기(rb-collapsed)
  for (const g of order) {
    if (fits()) break;
    tryShrink(g, 'rb-compact');
  }
  const folded: HTMLElement[] = [];
  for (const g of order) {
    if (fits()) break;
    if (tryShrink(g, 'rb-collapsed')) folded.push(g);
  }
  // 나중에 접은 묶음이 공간을 넉넉히 비웠으면 먼저 접은 묶음을 다시 편다(크게, 안 되면 작게) — 기존 기능이 접혀 숨지 않게
  for (const g of folded.reverse()) {
    g.classList.remove('rb-collapsed');
    if (fits()) continue;
    if (tryShrink(g, 'rb-compact') && fits()) continue;
    g.classList.remove('rb-compact');
    g.classList.add('rb-collapsed');
  }
  // 남는 공간이 있으면 작게 줄인 묶음을 중요한 것부터 다시 크게
  for (const g of [...order].reverse()) {
    if (!g.classList.contains('rb-compact')) continue;
    g.classList.remove('rb-compact');
    if (!fits()) g.classList.add('rb-compact');
  }
  if (openGroup && !openGroup.classList.contains('rb-collapsed')) closeGroupPopup();
}

export function installGroupToggles(toolbar: HTMLElement): void {
  toolbar.addEventListener('click', (event) => {
    const toggle = (event.target as Element).closest<HTMLElement>('.rb-group-toggle');
    const group = toggle?.closest<HTMLElement>('.ribbon-group');
    if (!toggle || !group) {
      // 판 안 명령을 누르면 닫는다(오피스처럼) — 안쪽 드롭다운·선택 상자·입력을 여는 조작은 제외
      const t = event.target as Element;
      if (openGroup && t.closest('.ribbon-group-buttons') && t.closest('button') && !t.closest('.dropdown-caret, .dropdown-menu, select, input')) closeGroupPopup();
      return;
    }
    wireToggle(group);
    const wasOpen = openGroup === group;
    closeGroupPopup();
    if (wasOpen) return;
    group.classList.add('rb-open');
    toggle.setAttribute('aria-expanded', 'true');
    openGroup = group;
    const body = group.querySelector<HTMLElement>(':scope > .ribbon-group-buttons');
    if (body) {
      const r = toggle.getBoundingClientRect();
      body.style.left = `${Math.max(4, Math.min(r.left, window.innerWidth - body.offsetWidth - 4))}px`;
      body.style.top = `${r.bottom + 2}px`;
    }
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Element;
    if (openGroup && !openGroup.contains(target) && !target.closest('.dropdown-menu')) closeGroupPopup();
  });
  document.addEventListener('keydown', (event) => {
    // 안쪽 드롭다운·선택 상자가 이미 처리한 Esc는 판까지 닫지 않는다
    if (event.key === 'Escape' && !event.defaultPrevented && !(event.target as Element | null)?.closest?.('.dropdown-menu, select')) closeGroupPopup();
  });
}
