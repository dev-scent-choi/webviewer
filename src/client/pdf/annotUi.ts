/**
 * PDF 편집기 주석 도구 화면 — 떠 있는 도구 줄(편집 도구·묶음 메뉴), 색·굵기·투명도 상자, 메모 목록(정보 창), 실행 취소.
 * 그리기·고르기 자체는 annotLayer.ts가 한다.
 */
import * as Y from 'yjs';
import { ANNOT_LABELS, type PdfAnnot, type PdfAnnotType, type PdfPageEntry } from '../../pdf/annotTypes.js';
import { AnnotLayer, formatTime, LOCAL_ORIGIN, type ToolName } from './annotLayer.js';
import type { PageView } from './pageView.js';

const SWATCHES = ['#D93025', '#F29900', '#FFE600', '#1E8E3E', '#1A73E8', '#9334E6', '#1F1F23'];
const WIDTH_TYPES = new Set<PdfAnnotType>(['ink', 'rect', 'ellipse', 'line', 'arrow']);
const TOOL_KEYS: Record<string, ToolName> = { n: 'note', u: 'highlight', p: 'ink', t: 'freetext', r: 'rect', e: 'eraser' };

export interface AnnotUi {
  setTool(tool: ToolName): void;
  readonly undoManager: Y.UndoManager;
}

export function setupAnnotUi(opts: {
  layer: AnnotLayer;
  map: Y.Map<PdfAnnot>;
  /** 쪽 목록 — 실행 취소가 쪽 정리도 함께 되돌린다 */
  pages: Y.Array<PdfPageEntry>;
  view: PageView;
  canEdit: boolean;
  onShowComments: () => void;
}): AnnotUi {
  const { layer, map, canEdit, view } = opts;
  const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const toolsBox = $('annot-tools');
  const selectBtn = $('tool-select');
  const handBtn = $('tool-hand');
  const toolButtons = [...document.querySelectorAll<HTMLButtonElement>('#annot-tools [data-tool]')];
  toolsBox.hidden = !canEdit;

  const undoManager = new Y.UndoManager([map, opts.pages], { trackedOrigins: new Set([LOCAL_ORIGIN]), captureTimeout: 300 });
  const undoBtn = $<HTMLButtonElement>('annot-undo');
  const redoBtn = $<HTMLButtonElement>('annot-redo');
  const syncUndo = () => {
    undoBtn.disabled = undoManager.undoStack.length === 0;
    redoBtn.disabled = undoManager.redoStack.length === 0;
  };
  undoManager.on('stack-item-added', syncUndo);
  undoManager.on('stack-item-popped', syncUndo);
  undoBtn.addEventListener('click', () => undoManager.undo());
  redoBtn.addEventListener('click', () => undoManager.redo());

  // ---------- 도구 고르기 ----------
  const stylePop = $('annot-style');
  const styleBtn = $('annot-style-btn');
  function styleTarget(): PdfAnnotType | null {
    const sel = layer.selectedAnnot;
    if (sel) return sel.type;
    const t = layer.currentTool;
    return t === 'select' || t === 'hand' || t === 'eraser' ? null : t;
  }
  function setTool(tool: ToolName): void {
    layer.setTool(tool);
    selectBtn.setAttribute('aria-pressed', String(tool === 'select'));
    handBtn.setAttribute('aria-pressed', String(tool === 'hand'));
    for (const b of toolButtons) b.setAttribute('aria-pressed', String(b.dataset.tool === tool));
    closeMenus();
    syncStyle();
  }
  selectBtn.addEventListener('click', () => setTool('select'));
  handBtn.addEventListener('click', () => setTool('hand'));
  for (const b of toolButtons) b.addEventListener('click', () => setTool(b.dataset.tool as ToolName));

  // 묶음(형광펜류·도형): ▾를 누르면 종류 목록, 고르면 그 종류가 묶음 단추가 된다
  const groups = [...toolsBox.querySelectorAll<HTMLElement>('.pdf-float-group')];
  function closeMenus(): void {
    for (const g of groups) (g.querySelector('.pdf-float-menu') as HTMLElement).hidden = true;
  }
  for (const g of groups) {
    const main = g.querySelector<HTMLButtonElement>('[data-tool]')!;
    const menu = g.querySelector<HTMLElement>('.pdf-float-menu')!;
    g.querySelector('.pdf-float-more')!.addEventListener('click', (event) => {
      event.stopPropagation();
      const open = menu.hidden;
      closeMenus();
      menu.hidden = !open;
    });
    for (const pick of menu.querySelectorAll<HTMLButtonElement>('[data-pick]')) {
      pick.addEventListener('click', () => {
        const tool = pick.dataset.pick as ToolName;
        main.dataset.tool = tool;
        main.title = pick.title;
        main.setAttribute('aria-label', pick.title);
        main.innerHTML = pick.querySelector('svg')!.outerHTML;
        setTool(tool);
      });
    }
  }
  document.addEventListener('pointerdown', (event) => {
    const t = event.target as HTMLElement;
    if (!t.closest('.pdf-float-group')) closeMenus();
    if (!t.closest('#annot-style, #annot-style-btn')) stylePop.hidden = true;
  });

  // ---------- 색·굵기·투명도 ----------
  const colorsRow = stylePop.querySelector<HTMLElement>('.pdf-style-colors')!;
  const colorInput = $<HTMLInputElement>('annot-color');
  const widthInput = $<HTMLInputElement>('annot-width');
  const fontInput = $<HTMLInputElement>('annot-fontsize');
  const opacityInput = $<HTMLInputElement>('annot-opacity');
  for (const c of SWATCHES) {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'pdf-style-swatch';
    b.style.background = c;
    b.dataset.color = c;
    b.title = c;
    b.addEventListener('click', () => applyStyle({ color: c }));
    colorsRow.append(b);
  }
  function applyStyle(patch: { color?: string; width?: number; opacity?: number }): void {
    const type = styleTarget();
    if (!type) return;
    layer.setStyle(type, patch);
    syncStyle();
  }
  function syncStyle(): void {
    const type = styleTarget();
    styleBtn.hidden = !type;
    if (!type) {
      stylePop.hidden = true;
      return;
    }
    const sel = layer.selectedAnnot;
    const s = sel ? { color: sel.color, width: sel.width, opacity: sel.opacity } : layer.styleOf(type);
    styleBtn.style.setProperty('--dot', s.color);
    colorInput.value = s.color.toLowerCase();
    for (const b of colorsRow.querySelectorAll<HTMLButtonElement>('button')) b.setAttribute('aria-pressed', String(b.dataset.color === s.color.toUpperCase()));
    (widthInput.closest('.pdf-style-row') as HTMLElement).hidden = !WIDTH_TYPES.has(type);
    widthInput.value = String(s.width);
    $('annot-width-val').textContent = `${s.width}pt`;
    const fontRow = fontInput.closest('.pdf-style-row') as HTMLElement;
    fontRow.hidden = type !== 'freetext' || !sel;
    if (sel?.type === 'freetext') {
      fontInput.value = String(sel.fontSize ?? 12);
      $('annot-fontsize-val').textContent = `${sel.fontSize ?? 12}pt`;
    }
    opacityInput.value = String(Math.round(s.opacity * 100));
    $('annot-opacity-val').textContent = `${Math.round(s.opacity * 100)}%`;
  }
  styleBtn.addEventListener('click', () => {
    stylePop.hidden = !stylePop.hidden;
    stylePop.style.top = `${styleBtn.offsetTop + 8}px`;
    syncStyle();
  });
  colorInput.addEventListener('change', () => applyStyle({ color: colorInput.value.toUpperCase() }));
  widthInput.addEventListener('change', () => applyStyle({ width: Number(widthInput.value) }));
  widthInput.addEventListener('input', () => ($('annot-width-val').textContent = `${widthInput.value}pt`));
  opacityInput.addEventListener('change', () => applyStyle({ opacity: Number(opacityInput.value) / 100 }));
  opacityInput.addEventListener('input', () => ($('annot-opacity-val').textContent = `${opacityInput.value}%`));
  fontInput.addEventListener('input', () => ($('annot-fontsize-val').textContent = `${fontInput.value}pt`));
  fontInput.addEventListener('change', () => {
    const sel = layer.selectedAnnot;
    if (sel?.type === 'freetext') layer.update(sel.id, { fontSize: Number(fontInput.value) });
  });
  layer.onSelect(() => {
    syncStyle();
    renderComments();
  });

  // ---------- 메모 목록 ----------
  const list = $('comments');
  const count = $('comments-count');
  function renderComments(): void {
    // 지금 쪽 순서로 — 지운 쪽의 주석은 보이지 않는다(쪽을 되살리면 다시 나온다)
    const pageOf = new Map(layer.all().map((a) => [a.id, view.pageOfKey(a.pageKey)]));
    const annots = layer
      .all()
      .filter((a) => pageOf.get(a.id) !== null)
      .sort((a, b) => pageOf.get(a.id)! - pageOf.get(b.id)! || a.createdAt.localeCompare(b.createdAt));
    count.textContent = annots.length ? String(annots.length) : '';
    list.replaceChildren();
    if (annots.length === 0) {
      list.innerHTML = '<p class="pdf-empty">아직 주석이 없습니다.<br />편집 탭의 도구로 메모·형광펜 등을 남겨 보세요.</p>';
      return;
    }
    let lastPage = -1;
    const selectedId = layer.selectedAnnot?.id;
    for (const a of annots) {
      const page = pageOf.get(a.id)!;
      if (page !== lastPage) {
        const h = document.createElement('div');
        h.className = 'pdf-comment-page';
        h.textContent = `${page}쪽`;
        list.append(h);
        lastPage = page;
      }
      const row = document.createElement('div');
      row.className = 'pdf-comment';
      row.tabIndex = 0;
      row.dataset.id = a.id;
      if (a.id === selectedId) row.classList.add('selected');
      const dot = document.createElement('span');
      dot.className = 'pdf-comment-dot';
      dot.style.setProperty('--c', a.color);
      const body = document.createElement('div');
      body.className = 'pdf-comment-body';
      const meta = document.createElement('div');
      meta.className = 'pdf-comment-meta';
      const who = document.createElement('strong');
      who.textContent = `${ANNOT_LABELS[a.type]}${a.author ? ` · ${a.author}` : ''}`;
      const when = document.createElement('span');
      when.textContent = formatTime(a.modifiedAt || a.createdAt);
      meta.append(who, when);
      body.append(meta);
      if (a.text) {
        const t = document.createElement('div');
        t.className = 'pdf-comment-text';
        t.textContent = a.text;
        body.append(t);
      }
      row.append(dot, body);
      if (canEdit) {
        const del = document.createElement('button');
        del.type = 'button';
        del.className = 'pdf-icon-btn pdf-comment-del';
        del.title = '지우기';
        del.setAttribute('aria-label', '주석 지우기');
        del.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M5 7h14M10 7V4h4v3M7 7l1 13h8l1-13"/></svg>';
        del.addEventListener('click', (event) => {
          event.stopPropagation();
          layer.remove(a.id);
        });
        row.append(del);
      }
      row.addEventListener('click', () => layer.reveal(a.id));
      row.addEventListener('keydown', (event) => {
        if (event.key === 'Enter') layer.reveal(a.id);
      });
      list.append(row);
    }
  }
  layer.onChange(renderComments);
  view.onPagesChange(renderComments);
  renderComments();

  // ---------- 단축키 ----------
  document.addEventListener('keydown', (event) => {
    const typing = (event.target as HTMLElement).closest('input, textarea, [contenteditable="true"]');
    const mod = event.ctrlKey || event.metaKey;
    if (!canEdit || typing) return;
    if (mod && (event.key === 'z' || event.key === 'Z') && !event.shiftKey) {
      event.preventDefault();
      undoManager.undo();
    } else if (mod && (event.key === 'y' || event.key === 'Y' || ((event.key === 'z' || event.key === 'Z') && event.shiftKey))) {
      event.preventDefault();
      undoManager.redo();
    } else if (document.body.classList.contains('organizing')) {
      // 정리 화면에서는 주석 도구 단축키를 쓰지 않는다(그 화면 단축키는 organizeView.ts)
      return;
    } else if (!mod && !event.altKey && TOOL_KEYS[event.key.toLowerCase()]) {
      setTool(TOOL_KEYS[event.key.toLowerCase()]);
    } else if (!mod && (event.key === 'v' || event.key === 'V')) setTool('select');
    else if (!mod && (event.key === 'h' || event.key === 'H')) setTool('hand');
  });

  setTool('select');
  void opts.onShowComments;
  return { setTool, undoManager };
}
