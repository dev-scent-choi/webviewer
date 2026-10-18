/**
 * 리본의 ▾ 드롭다운 메뉴 공용 처리 — 리본 줄은 가로 스크롤(overflow) 영역이라 그 안에 absolute로
 * 띄운 메뉴는 잘린다. 그래서 메뉴를 화면 기준(position: fixed)으로 ▾ 바로 아래에 띄운다.
 *
 * 마크업: <span class="dropdown-btn"> … <button class="dropdown-caret"> <div class="dropdown-menu" hidden>
 * 메뉴 항목은 [data-value]를 가진 버튼이다.
 */

let openMenu: HTMLElement | null = null;

function closeOpenMenu(): void {
  if (openMenu) openMenu.hidden = true;
  openMenu = null;
}

document.addEventListener('mousedown', (event) => {
  const target = event.target as HTMLElement;
  if (openMenu && !openMenu.contains(target) && !target.closest('.dropdown-caret')) closeOpenMenu();
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') closeOpenMenu();
});
window.addEventListener('resize', closeOpenMenu);

function positionMenu(menu: HTMLElement, anchor: HTMLElement): void {
  const rect = anchor.getBoundingClientRect();
  menu.classList.add('dropdown-menu-floating');
  menu.style.left = `${Math.max(4, Math.min(rect.left, window.innerWidth - menu.offsetWidth - 4))}px`;
  menu.style.top = `${rect.bottom + 2}px`;
}

/** 버튼 하나로 메뉴를 여닫는다(▾가 없는 "선택" 같은 단일 메뉴 버튼에도 쓴다). */
export function bindMenuButton(button: HTMLElement, menu: HTMLElement, onPick: (value: string, item: HTMLElement) => void): void {
  button.addEventListener('click', (event) => {
    event.stopPropagation();
    if (openMenu === menu) {
      closeOpenMenu();
      return;
    }
    closeOpenMenu();
    menu.hidden = false;
    openMenu = menu;
    positionMenu(menu, button.closest('.dropdown-btn') ?? button);
  });
  menu.addEventListener('click', (event) => {
    const item = (event.target as HTMLElement).closest<HTMLElement>('[data-value]');
    if (!item || !menu.contains(item)) return;
    closeOpenMenu();
    onPick(item.dataset.value ?? '', item);
  });
}

/** .dropdown-btn 안의 ▾(.dropdown-caret)와 메뉴(.dropdown-menu)를 연결한다. */
export function bindDropdown(root: HTMLElement | null, onPick: (value: string, item: HTMLElement) => void): void {
  if (!root) return;
  const caret = root.querySelector<HTMLElement>('.dropdown-caret');
  const menu = root.querySelector<HTMLElement>('.dropdown-menu');
  if (caret && menu) bindMenuButton(caret, menu, onPick);
}

export interface PaletteColor {
  readonly hex: string;
  readonly label: string;
}

/** 색 견본 격자를 메뉴에 채운다 — "색 없음"은 data-value="none". */
export function fillColorPalette(menu: HTMLElement, colors: readonly PaletteColor[], noneLabel: string): void {
  menu.classList.add('dropdown-palette');
  const grid = document.createElement('div');
  grid.className = 'palette-grid';
  for (const color of colors) {
    const swatch = document.createElement('button');
    swatch.type = 'button';
    swatch.className = 'palette-swatch';
    swatch.dataset.value = color.hex;
    swatch.title = color.label;
    swatch.setAttribute('aria-label', color.label);
    swatch.style.background = color.hex;
    grid.append(swatch);
  }
  const none = document.createElement('button');
  none.type = 'button';
  none.className = 'palette-none';
  none.dataset.value = 'none';
  none.textContent = noneLabel;
  menu.replaceChildren(grid, none);
}
