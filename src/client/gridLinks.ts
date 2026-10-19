/**
 * 엑셀 편집 화면의 하이퍼링크(바로가기)와 목록 상자(데이터 유효성 "목록").
 *
 * - 링크가 걸린 칸은 클릭하면 엑셀처럼 바로 이동한다: 문서 안 위치('자료'!A1)는 그 시트·칸으로, 바깥 주소는 새 창으로.
 *   Ctrl·Shift를 누른 채 클릭하면 이동하지 않고 칸만 고른다(링크 칸을 고쳐야 할 때).
 * - 목록이 걸린 칸을 고르면 칸 오른쪽에 ▼ 단추가 나타나고, 누르거나 Alt+↓로 목록을 열어 값을 고른다.
 *
 * 정보는 서버가 원본에서 읽어 준다(/api/view/xlsx-extras, src/xlsx/sheetExtras.ts) — 편집 화면은 읽기만 한다.
 */

interface CellRange {
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}
interface SheetLink extends CellRange {
  readonly sheetName: string;
  readonly url?: string;
  readonly location?: string;
  readonly tooltip?: string;
}
interface SheetListValidation {
  readonly sheetName: string;
  readonly ranges: readonly CellRange[];
  readonly values?: readonly string[];
  readonly source?: CellRange & { readonly sheetName?: string };
}

export interface GridLinksHost {
  readonly token: string;
  readonly tableEl: HTMLElement;
  readonly scrollHost: HTMLElement;
  /** 지금 보고 있는 시트 이름 */
  currentSheetName(): string;
  /** 시트 이름으로 그 시트·칸으로 옮겨 고른다. 없는 시트면 false */
  goTo(sheetName: string | null, row: number, col: number): boolean;
  /** 칸에 보이는 값(목록이 칸 범위를 가리킬 때 항목을 푼다) */
  valueText(sheetName: string, row: number, col: number): string;
  /** 지금 시트의 칸에 값을 넣는다(편집 권한·잠금은 호출부가 판단) */
  pickValue(td: HTMLTableCellElement, row: number, col: number, value: string): void;
  isEditing(td: HTMLTableCellElement): boolean;
}

export interface GridLinks {
  /** 칸을 그릴 때마다 — 링크 칸 표시 */
  decorate(td: HTMLTableCellElement, row: number, col: number): void;
}

const inRange = (r: CellRange, row: number, col: number) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;

/** "'자료'!A1", "요약!$C$5", "B3" → 시트 이름(없으면 null)과 칸(0부터) */
export function parseLocation(location: string): { sheetName: string | null; row: number; col: number } | null {
  const text = location.trim().replace(/^#/, '');
  const bang = text.lastIndexOf('!');
  const sheetName = bang >= 0 ? text.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : null;
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)/.exec(bang >= 0 ? text.slice(bang + 1) : text);
  if (!m) return null;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  return { sheetName, row: Number(m[2]) - 1, col: col - 1 };
}

export function setupGridLinks(host: GridLinksHost): GridLinks {
  let links: SheetLink[] = [];
  let lists: SheetListValidation[] = [];

  const linkAt = (row: number, col: number) => {
    const sheet = host.currentSheetName();
    return links.find((l) => l.sheetName === sheet && inRange(l, row, col));
  };
  const listAt = (row: number, col: number) => {
    const sheet = host.currentSheetName();
    return lists.find((l) => l.sheetName === sheet && l.ranges.some((r) => inRange(r, row, col)));
  };
  const cellOf = (target: EventTarget | null) => {
    const td = (target as HTMLElement | null)?.closest?.<HTMLTableCellElement>('td[data-row][data-col]');
    return td ? { td, row: Number(td.dataset.row), col: Number(td.dataset.col) } : null;
  };

  function redecorateVisible(): void {
    for (const td of host.tableEl.querySelectorAll<HTMLTableCellElement>('td[data-row][data-col]')) {
      decorate(td, Number(td.dataset.row), Number(td.dataset.col));
    }
  }

  function decorate(td: HTMLTableCellElement, row: number, col: number): void {
    const link = links.length ? linkAt(row, col) : undefined;
    td.classList.toggle('g-link', Boolean(link));
    if (link) {
      td.title = link.tooltip || (link.url ? `${link.url}\n클릭하면 열립니다 (Ctrl+클릭: 칸 선택)` : `클릭하면 ${link.location}(으)로 이동합니다 (Ctrl+클릭: 칸 선택)`);
      td.dataset.linkTitle = '';
    } else if (td.dataset.linkTitle !== undefined) {
      // 가상 스크롤로 같은 <td>가 다른 칸에 다시 쓰일 때 예전 링크 설명이 남지 않게
      td.removeAttribute('title');
      delete td.dataset.linkTitle;
    }
  }

  function follow(link: SheetLink): void {
    if (link.url) {
      window.open(link.url, '_blank', 'noopener,noreferrer');
      return;
    }
    const target = link.location ? parseLocation(link.location) : null;
    if (target) host.goTo(target.sheetName, target.row, target.col);
  }

  // 링크 칸 클릭 — 칸을 고르는 동작(mousedown)이 먼저 끝난 뒤 이동한다. 드래그로 여러 칸을 고른 경우는 제외.
  let downAt: { x: number; y: number } | null = null;
  host.tableEl.addEventListener('mousedown', (event) => {
    downAt = { x: event.clientX, y: event.clientY };
  }, true);
  host.tableEl.addEventListener('click', (event) => {
    if (event.button !== 0 || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
    const hit = cellOf(event.target);
    if (!hit || !hit.td.classList.contains('g-link') || host.isEditing(hit.td)) return;
    if (downAt && Math.hypot(event.clientX - downAt.x, event.clientY - downAt.y) > 4) return;
    const link = linkAt(hit.row, hit.col);
    if (link) follow(link);
  });

  // ---- 목록 상자 ----
  const dropBtn = document.createElement('button');
  dropBtn.type = 'button';
  dropBtn.className = 'grid-list-btn';
  dropBtn.hidden = true;
  dropBtn.title = '목록에서 고르기 (Alt+↓)';
  dropBtn.setAttribute('aria-label', '목록에서 고르기');
  dropBtn.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>';
  host.scrollHost.append(dropBtn);
  const menu = document.createElement('div');
  menu.className = 'dropdown-menu dropdown-menu-floating grid-list-menu';
  menu.setAttribute('role', 'listbox');
  menu.hidden = true;
  document.body.append(menu);

  let current: { td: HTMLTableCellElement; row: number; col: number; list: SheetListValidation } | null = null;

  function itemsOf(list: SheetListValidation): string[] {
    if (list.values) return [...list.values];
    const src = list.source;
    if (!src) return [];
    const sheet = src.sheetName ?? list.sheetName;
    const out: string[] = [];
    for (let r = src.r1; r <= src.r2 && out.length < 500; r += 1) {
      for (let c = src.c1; c <= src.c2 && out.length < 500; c += 1) {
        const v = host.valueText(sheet, r, c);
        if (v !== '') out.push(v);
      }
    }
    return out;
  }

  function placeButton(): void {
    if (!current || !current.td.isConnected || !current.td.classList.contains('g-editable')) {
      dropBtn.hidden = true;
      return;
    }
    const hostRect = host.scrollHost.getBoundingClientRect();
    const rect = current.td.getBoundingClientRect();
    dropBtn.style.left = `${rect.right - hostRect.left + host.scrollHost.scrollLeft + 1}px`;
    dropBtn.style.top = `${rect.top - hostRect.top + host.scrollHost.scrollTop}px`;
    dropBtn.style.height = `${Math.max(16, rect.height)}px`;
    dropBtn.hidden = false;
  }

  function closeMenu(): void {
    menu.hidden = true;
  }

  function openMenu(): void {
    if (!current) return;
    const { td, row, col, list } = current;
    const items = itemsOf(list);
    menu.replaceChildren();
    const now = td.textContent ?? '';
    for (const value of items) {
      const b = document.createElement('button');
      b.type = 'button';
      b.setAttribute('role', 'option');
      b.textContent = value;
      if (value === now) b.classList.add('is-current');
      b.addEventListener('click', () => {
        closeMenu();
        host.pickValue(td, row, col, value);
      });
      menu.append(b);
    }
    if (items.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'grid-list-empty';
      empty.textContent = '목록이 비어 있습니다';
      menu.append(empty);
    }
    const rect = td.getBoundingClientRect();
    menu.style.minWidth = `${Math.max(rect.width + 18, 120)}px`;
    menu.style.left = `${rect.left}px`;
    menu.style.top = `${rect.bottom + 2}px`;
    menu.hidden = false;
    // 화면 아래로 넘치면 칸 위로
    const mh = menu.getBoundingClientRect().height;
    if (rect.bottom + 2 + mh > window.innerHeight - 8) menu.style.top = `${Math.max(8, rect.top - 2 - mh)}px`;
    (menu.querySelector<HTMLButtonElement>('.is-current') ?? menu.querySelector<HTMLButtonElement>('button'))?.focus();
  }

  dropBtn.addEventListener('mousedown', (event) => event.preventDefault()); // 칸 포커스를 뺏지 않는다
  dropBtn.addEventListener('click', () => (menu.hidden ? openMenu() : closeMenu()));
  document.addEventListener('mousedown', (event) => {
    if (!menu.hidden && !menu.contains(event.target as Node) && event.target !== dropBtn && !dropBtn.contains(event.target as Node)) closeMenu();
  });
  menu.addEventListener('keydown', (event) => {
    const buttons = [...menu.querySelectorAll<HTMLButtonElement>('button')];
    const i = buttons.indexOf(document.activeElement as HTMLButtonElement);
    if (event.key === 'Escape') {
      event.preventDefault();
      closeMenu();
      current?.td.focus();
    } else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      event.preventDefault();
      buttons[Math.min(buttons.length - 1, Math.max(0, i + (event.key === 'ArrowDown' ? 1 : -1)))]?.focus();
    }
  });

  host.tableEl.addEventListener('focusin', (event) => {
    const hit = cellOf(event.target);
    const list = hit && lists.length ? listAt(hit.row, hit.col) : undefined;
    current = hit && list ? { ...hit, list } : null;
    closeMenu();
    placeButton();
  });
  host.tableEl.addEventListener('keydown', (event) => {
    if (current && event.altKey && event.key === 'ArrowDown' && !host.isEditing(current.td)) {
      event.preventDefault();
      event.stopPropagation();
      openMenu();
    }
  }, true);
  host.scrollHost.addEventListener('scroll', () => {
    closeMenu();
    placeButton();
  }, { passive: true });

  void fetch(`/api/view/xlsx-extras?token=${encodeURIComponent(host.token)}`)
    .then((res) => res.json() as Promise<{ data?: { links?: SheetLink[]; lists?: SheetListValidation[] } }>)
    .then((json) => {
      links = json.data?.links ?? [];
      lists = json.data?.lists ?? [];
      redecorateVisible();
    })
    .catch(() => undefined);

  return { decorate };
}
