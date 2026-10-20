/**
 * 페이지 정리 화면(pdfcraft organize와 같은 배치) — 모든 쪽을 카드로 늘어놓고 고르기(클릭·Ctrl·Shift·Ctrl+A)·끌어 옮기기,
 * 위 도구 줄로 회전·지우기·빈 쪽·파일에서 넣기·앞으로/뒤로·고른 쪽 추출·나누기. 바뀐 것은 방의 쪽 목록(pageOps.ts)에 쓰고,
 * 본문 보기(pageView.setPages)·썸네일은 그 목록을 따라 바뀐다. 파일 합치기·추출 내려받기도 여기서 한다.
 */
import type * as Y from 'yjs';
import type { PdfAnnot, PdfPageEntry } from '../../pdf/annotTypes.js';
import type { PageOps } from './pageOps.js';
import type { PageView, PageViewport } from './pageView.js';
import { CONVERTIBLE_ACCEPT } from './convertUi.js';

const CARD_WIDTH = 150;

export interface OrganizeOptions {
  readonly view: PageView;
  readonly ops: PageOps;
  readonly undoManager: Y.UndoManager;
  readonly token: string;
  readonly docName: string;
  readonly canEdit: boolean;
  readonly canDownload: boolean;
  /** 카드 위에 주석을 그린다(annotLayer.drawOverlay) */
  readonly drawOverlay: (page: number, host: HTMLElement, vp: PageViewport) => void;
  readonly notice: (text: string) => void;
}

export interface OrganizeView {
  open(hint?: string): void;
  close(): void;
  readonly isOpen: boolean;
  /** 파일 고르기 창을 열어 고른 PDF들을 at 자리(없으면 끝)에 넣는다 */
  pickFiles(at?: number): void;
}

interface Card {
  readonly el: HTMLDivElement;
  readonly img: HTMLDivElement;
  readonly num: HTMLSpanElement;
  key: string;
  rotate: number;
  drawn: boolean;
}

interface PartResponse {
  readonly partId: string;
  readonly pageCount: number;
  readonly sizes: ReadonlyArray<{ w: number; h: number; rotate: number }>;
  readonly annots: readonly PdfAnnot[];
}

export function setupOrganize(opts: OrganizeOptions): OrganizeView {
  const { view, ops, canEdit, canDownload } = opts;
  const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const root = $('organize');
  const grid = $('org-grid');
  const status = $('org-status');
  const fileInput = $<HTMLInputElement>('org-file');
  fileInput.accept = CONVERTIBLE_ACCEPT;
  const splitPop = $('org-split');
  const dropLine = document.createElement('div');
  dropLine.className = 'pdf-org-drop';
  let isOpen = false;
  const selected = new Set<string>();
  let anchorKey: string | null = null;
  let cards: Card[] = [];
  const cardByKey = new Map<string, Card>();
  let insertAt: number | null = null;

  for (const el of root.querySelectorAll<HTMLElement>('[data-edit-only]')) el.hidden = !canEdit;
  for (const el of root.querySelectorAll<HTMLElement>('[data-download-only]')) el.hidden = !canDownload;

  // ---------- 카드 ----------
  const observer = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const card = cards.find((c) => c.el === entry.target);
        if (card && !card.drawn) void drawCard(card);
      }
    },
    { root: grid, rootMargin: '300px 0px' },
  );

  async function drawCard(card: Card): Promise<void> {
    card.drawn = true;
    const page = view.pageOfKey(card.key);
    if (page === null) return;
    const vp = await view.drawThumbnail(page, card.img, CARD_WIDTH);
    if (!vp || view.pageOfKey(card.key) === null) return;
    const layer = document.createElement('div');
    layer.className = 'pdf-thumb-annots';
    card.img.append(layer);
    opts.drawOverlay(view.pageOfKey(card.key)!, layer, vp);
  }

  function makeCard(entry: PdfPageEntry): Card {
    const el = document.createElement('div');
    el.className = 'pdf-org-card';
    el.tabIndex = 0;
    el.draggable = canEdit;
    el.dataset.key = entry.key;
    const img = document.createElement('div');
    img.className = 'pdf-org-img';
    img.style.width = `${CARD_WIDTH}px`;
    img.style.height = `${Math.round(CARD_WIDTH * 1.414)}px`;
    const num = document.createElement('span');
    num.className = 'pdf-org-num';
    el.append(img, num);
    const card: Card = { el, img, num, key: entry.key, rotate: entry.rotate, drawn: false };
    el.addEventListener('click', (event) => onCardClick(card, event));
    el.addEventListener('dblclick', () => {
      const page = view.pageOfKey(card.key);
      close();
      if (page !== null) view.goTo(page);
    });
    el.addEventListener('dragstart', (event) => onDragStart(card, event));
    return card;
  }

  /** 쪽 목록에 맞춰 카드를 다시 늘어놓는다 — 같은 열쇠·같은 회전의 카드는 그림을 그대로 둔다 */
  function rebuild(): void {
    if (!isOpen) return;
    const entries = view.entries;
    const keys = new Set(entries.map((e) => e.key));
    for (const [key, card] of cardByKey) {
      if (!keys.has(key)) {
        observer.unobserve(card.el);
        card.el.remove();
        cardByKey.delete(key);
      }
    }
    for (const key of [...selected]) if (!keys.has(key)) selected.delete(key);
    cards = entries.map((entry, i) => {
      let card = cardByKey.get(entry.key);
      if (!card) {
        card = makeCard(entry);
        cardByKey.set(entry.key, card);
        observer.observe(card.el);
      } else if (card.rotate !== entry.rotate) {
        card.rotate = entry.rotate;
        card.drawn = false;
        observer.unobserve(card.el);
        observer.observe(card.el);
      }
      card.num.textContent = String(i + 1);
      grid.append(card.el);
      return card;
    });
    paintSelection();
  }

  function paintSelection(): void {
    for (const card of cards) {
      const on = selected.has(card.key);
      card.el.classList.toggle('selected', on);
      card.el.setAttribute('aria-selected', String(on));
    }
    status.textContent = selected.size ? `${view.pageCount}쪽 중 ${selected.size}쪽 고름` : `${view.pageCount}쪽`;
    for (const btn of root.querySelectorAll<HTMLButtonElement>('[data-needs-selection]')) btn.disabled = selected.size === 0;
  }

  function onCardClick(card: Card, event: MouseEvent): void {
    const order = cards.map((c) => c.key);
    if (event.shiftKey && anchorKey && order.includes(anchorKey)) {
      const [a, b] = [order.indexOf(anchorKey), order.indexOf(card.key)].sort((x, y) => x - y);
      if (!(event.ctrlKey || event.metaKey)) selected.clear();
      for (const key of order.slice(a, b + 1)) selected.add(key);
    } else if (event.ctrlKey || event.metaKey) {
      if (selected.has(card.key)) selected.delete(card.key);
      else selected.add(card.key);
      anchorKey = card.key;
    } else {
      selected.clear();
      selected.add(card.key);
      anchorKey = card.key;
    }
    paintSelection();
  }

  /** 고른 쪽 열쇠들(지금 순서대로) */
  const selectedKeys = (): string[] => cards.map((c) => c.key).filter((k) => selected.has(k));
  /** 새로 넣을 자리 — 고른 쪽 중 마지막 뒤, 없으면 끝 */
  const insertIndex = (): number => {
    const order = cards.map((c) => c.key);
    const last = Math.max(-1, ...selectedKeys().map((k) => order.indexOf(k)));
    return last < 0 ? order.length : last + 1;
  };

  // ---------- 끌어 옮기기 ----------
  let dragKeys: string[] = [];
  function onDragStart(card: Card, event: DragEvent): void {
    if (!canEdit) return;
    if (!selected.has(card.key)) {
      selected.clear();
      selected.add(card.key);
      anchorKey = card.key;
      paintSelection();
    }
    dragKeys = selectedKeys();
    event.dataTransfer?.setData('text/plain', dragKeys.join(','));
    if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move';
    root.classList.add('dragging');
  }
  /** 마우스 위치 → 끼울 자리(0..카드 수)와 표시 줄 위치 */
  function dropTarget(x: number, y: number): { index: number; card: Card | null; after: boolean } {
    let best: { index: number; card: Card | null; after: boolean; d: number } = { index: cards.length, card: null, after: true, d: Infinity };
    cards.forEach((card, i) => {
      const r = card.el.getBoundingClientRect();
      if (y < r.top - 12 || y > r.bottom + 12) return;
      const cx = r.left + r.width / 2;
      const d = Math.abs(x - cx);
      if (d < best.d) best = { index: x < cx ? i : i + 1, card, after: x >= cx, d };
    });
    return best;
  }
  grid.addEventListener('dragover', (event) => {
    if (dragKeys.length === 0) return;
    event.preventDefault();
    const t = dropTarget(event.clientX, event.clientY);
    insertAt = t.index;
    if (t.card) {
      const r = t.card.el.getBoundingClientRect();
      const g = grid.getBoundingClientRect();
      dropLine.style.left = `${(t.after ? r.right + 6 : r.left - 8) - g.left + grid.scrollLeft}px`;
      dropLine.style.top = `${r.top - g.top + grid.scrollTop}px`;
      dropLine.style.height = `${r.height}px`;
      if (!dropLine.isConnected) grid.append(dropLine);
    }
  });
  const endDrag = () => {
    dragKeys = [];
    insertAt = null;
    dropLine.remove();
    root.classList.remove('dragging');
  };
  grid.addEventListener('drop', (event) => {
    event.preventDefault();
    if (dragKeys.length && insertAt !== null) ops.move(dragKeys, insertAt);
    endDrag();
  });
  grid.addEventListener('dragend', endDrag);
  // 빈 곳을 누르면 고르기 풀기
  grid.addEventListener('click', (event) => {
    if (event.target === grid) {
      selected.clear();
      paintSelection();
    }
  });

  // ---------- 도구 줄 ----------
  const on = (id: string, fn: () => void) => $(id).addEventListener('click', fn);
  on('org-rotate-left', () => ops.rotate(selectedKeys(), -90));
  on('org-rotate-right', () => ops.rotate(selectedKeys(), 90));
  on('org-delete', removeSelected);
  on('org-blank', () => {
    const at = insertIndex();
    const key = ops.insertBlank(at, view.sizeOf(Math.max(1, Math.min(view.pageCount, at))));
    selected.clear();
    selected.add(key);
    anchorKey = key;
    paintSelection();
  });
  on('org-insert-file', () => pickFiles(insertIndex()));
  on('org-move-up', () => {
    const order = cards.map((c) => c.key);
    const first = Math.min(...selectedKeys().map((k) => order.indexOf(k)));
    if (first > 0) ops.move(selectedKeys(), first - 1);
  });
  on('org-move-down', () => {
    const order = cards.map((c) => c.key);
    const last = Math.max(...selectedKeys().map((k) => order.indexOf(k)));
    if (last < order.length - 1) ops.move(selectedKeys(), last + 2);
  });
  on('org-undo', () => opts.undoManager.undo());
  on('org-redo', () => opts.undoManager.redo());
  on('org-extract', () => void extract([selectedKeys()]));
  on('org-split-btn', () => {
    splitPop.hidden = !splitPop.hidden;
    $<HTMLInputElement>('org-split-at').disabled = selected.size === 0;
    if (selected.size === 0) $<HTMLInputElement>('org-split-every').checked = true;
  });
  on('org-split-go', () => {
    const keys = cards.map((c) => c.key);
    const groups: string[][] = [];
    if ($<HTMLInputElement>('org-split-every').checked) {
      const n = Math.max(1, Math.floor(Number($<HTMLInputElement>('org-split-n').value) || 1));
      for (let i = 0; i < keys.length; i += n) groups.push(keys.slice(i, i + n));
    } else {
      // 고른 쪽 앞에서 자른다
      let cur: string[] = [];
      for (const key of keys) {
        if (selected.has(key) && cur.length) {
          groups.push(cur);
          cur = [];
        }
        cur.push(key);
      }
      if (cur.length) groups.push(cur);
    }
    splitPop.hidden = true;
    if (groups.length < 2) {
      opts.notice('나눌 곳이 없습니다 — 쪽 수나 자를 쪽을 확인해 주세요.');
      return;
    }
    void extract(groups);
  });
  on('org-close', close);
  const syncUndo = () => {
    $<HTMLButtonElement>('org-undo').disabled = opts.undoManager.undoStack.length === 0;
    $<HTMLButtonElement>('org-redo').disabled = opts.undoManager.redoStack.length === 0;
  };
  opts.undoManager.on('stack-item-added', syncUndo);
  opts.undoManager.on('stack-item-popped', syncUndo);
  syncUndo();

  function removeSelected(): void {
    const keys = selectedKeys();
    if (keys.length === 0) return;
    if (!ops.remove(keys)) {
      opts.notice('모든 쪽을 지울 수는 없습니다.');
      return;
    }
    selected.clear();
  }

  // ---------- 파일 넣기 ----------
  let pendingAt: number | null = null;
  function pickFiles(at?: number): void {
    if (!canEdit) return;
    pendingAt = at ?? view.pageCount;
    fileInput.value = '';
    fileInput.click();
  }
  fileInput.addEventListener('change', () => {
    const files = [...(fileInput.files ?? [])];
    if (files.length) void insertFiles(files, pendingAt ?? view.pageCount);
  });

  async function insertFiles(files: File[], at: number): Promise<void> {
    let pos = at;
    const inserted: string[] = [];
    for (const [n, file] of files.entries()) {
      const converting = !/\.pdf$/i.test(file.name);
      opts.notice(`${converting ? 'PDF로 바꿔 ' : ''}넣는 중… ${files.length > 1 ? `(${n + 1}/${files.length}) ` : ''}${file.name}`);
      try {
        // PDF가 아니면(워드·한글·그림…) 서버가 PDF로 바꿔 넣는다
        const res = await fetch(`/api/view/pdf-part?token=${encodeURIComponent(opts.token)}&name=${encodeURIComponent(file.name)}`, {
          method: 'POST',
          headers: { 'content-type': 'application/octet-stream' },
          body: file,
        });
        const json = (await res.json().catch(() => ({}))) as { success?: boolean; data?: PartResponse; error?: string };
        if (!res.ok || !json.success || !json.data) throw new Error(json.error ?? `${file.name}을(를) 넣지 못했습니다.`);
        const part = json.data;
        const entries: PdfPageEntry[] = Array.from({ length: part.pageCount }, (_, i) => ({ key: `p:${part.partId}:${i}`, rotate: 0 }));
        ops.insert(pos, entries, part.annots);
        pos += entries.length;
        inserted.push(...entries.map((e) => e.key));
      } catch (error) {
        opts.notice(error instanceof Error ? error.message : String(error));
        return;
      }
    }
    if (inserted.length) {
      opts.notice(`${inserted.length}쪽을 넣었습니다.`);
      selected.clear();
      for (const key of inserted) selected.add(key);
      anchorKey = inserted[0];
      if (isOpen) rebuild();
      cardByKey.get(inserted[0])?.el.scrollIntoView({ block: 'nearest' });
    }
  }

  // ---------- 추출·나누기 내려받기 ----------
  async function extract(groups: string[][]): Promise<void> {
    if (!canDownload) return;
    const valid = groups.filter((g) => g.length > 0);
    if (valid.length === 0) {
      opts.notice('추출할 쪽을 먼저 고르세요.');
      return;
    }
    opts.notice(valid.length > 1 ? `${valid.length}개 파일로 나누는 중…` : '추출하는 중…');
    try {
      const res = await fetch(`/api/view/pdf-extract?token=${encodeURIComponent(opts.token)}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ groups: valid }),
      });
      if (!res.ok) {
        const json = (await res.json().catch(() => ({}))) as { error?: string };
        throw new Error(json.error ?? '내려받지 못했습니다.');
      }
      const disposition = res.headers.get('content-disposition') ?? '';
      const m = /filename\*=UTF-8''([^;]+)/i.exec(disposition);
      const fileName = m ? decodeURIComponent(m[1]) : valid.length > 1 ? 'split.zip' : 'extract.pdf';
      const url = URL.createObjectURL(await res.blob());
      const a = document.createElement('a');
      a.href = url;
      a.download = fileName;
      document.body.append(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
      opts.notice(`${fileName} 내려받기`);
    } catch (error) {
      opts.notice(error instanceof Error ? error.message : String(error));
    }
  }

  // ---------- 열기·닫기·단축키 ----------
  function open(hint?: string): void {
    isOpen = true;
    root.hidden = false;
    document.body.classList.add('organizing');
    rebuild();
    if (hint) opts.notice(hint);
    grid.focus();
  }
  function close(): void {
    isOpen = false;
    root.hidden = true;
    splitPop.hidden = true;
    document.body.classList.remove('organizing');
  }
  view.onPagesChange(rebuild);
  document.addEventListener('keydown', (event) => {
    if (!isOpen) return;
    const typing = (event.target as HTMLElement).closest('input, textarea, [contenteditable="true"]');
    if (typing) return;
    const mod = event.ctrlKey || event.metaKey;
    if (event.key === 'Escape') {
      event.preventDefault();
      if (!splitPop.hidden) splitPop.hidden = true;
      else close();
    } else if (mod && (event.key === 'a' || event.key === 'A')) {
      event.preventDefault();
      for (const c of cards) selected.add(c.key);
      paintSelection();
    } else if (canEdit && (event.key === 'Delete' || event.key === 'Backspace')) {
      event.preventDefault();
      removeSelected();
    }
  });

  return {
    open,
    close,
    get isOpen() {
      return isOpen;
    },
    pickFiles,
  };
}
