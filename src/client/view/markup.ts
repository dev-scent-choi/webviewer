/**
 * 강조 표시(코멘트 사각형)·텍스트 추가 — 지금 보기 화면(viewer.js 마크업·텍스트 추가)을 옮겼다.
 * 서버 저장(/api/view/markup)·화면 클래스는 같다. 사용자 글은 textContent로만 넣는다.
 */
import type { CommentMarkupItem, MarkupItem, TextMarkupItem } from '../../markup.js';
import type { ShellCtx } from './ctx.js';
import { disabledReason, toFrameRatio } from './logic.js';

type Mutable<T> = T extends unknown ? { -readonly [K in keyof T]: T[K] } : never;
type Item = Mutable<MarkupItem>;
type TextItem = Mutable<TextMarkupItem>;
type CommentItem = Mutable<CommentMarkupItem>;
interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

interface TextStyle {
  fontSizePt: number;
  bold: boolean;
  fontFamily: string;
  baselineY: number;
  matched: boolean;
}

interface PendingText {
  wrapper: HTMLElement;
  box: HTMLTextAreaElement;
  frame: HTMLElement;
  page: number;
  xRatio: number;
  style: TextStyle;
}

/** viewer.js fetchJsonView와 같은 응답 모양({success, data, error}) — 실패면 그 문구로 throw */
async function fetchJson<T = unknown>(url: string, options: RequestInit = {}): Promise<T> {
  const res = await fetch(url, options);
  const json = (await res.json().catch(() => null)) as { success?: boolean; data?: T; error?: string } | null;
  if (!res.ok || !json?.success) throw new Error(json?.error || '요청에 실패했습니다.');
  return json.data as T;
}

const errText = (error: unknown, fallback: string): string => (error instanceof Error && error.message) || fallback;
const FONT_SIZE_MIN_PT = 4;
const FONT_SIZE_MAX_PT = 200;
const FONT_SIZE_STEP_PT = 1;
const JSON_HEADERS = { 'content-type': 'application/json' };

function clampFontSizePt(pt: number): number {
  return Math.min(FONT_SIZE_MAX_PT, Math.max(FONT_SIZE_MIN_PT, Math.round(pt)));
}

/** 서버가 감지한 글꼴 계열을 화면에서도 비슷하게 보이도록 CSS 폰트로 매핑한다(실제 PDF에 굽는 폰트와 맞춘다) */
function cssFontFamilyFor(fontFamily: string): string {
  if (fontFamily === 'serif') return '"Nanum Myeongjo", "Noto Serif KR", serif';
  return '"Pretendard Variable", "Pretendard", sans-serif';
}

function button(cls: string, text: string, title?: string): HTMLButtonElement {
  const b = document.createElement('button');
  b.type = 'button';
  b.className = cls;
  b.textContent = text;
  if (title) b.title = title;
  return b;
}

const clamp01 = (v: number): number => Math.min(Math.max(v, 0), 1);
const frameOf = (event: Event): HTMLElement | null => (event.target as Element).closest<HTMLElement>('.page-frame');

export function setupMarkup(ctx: ShellCtx): void {
  const reason = disabledReason('markup', ctx.info.permissions, ctx.info.ext);
  const addBtn = document.getElementById('markup-add-btn') as HTMLButtonElement;
  const textBtn = document.getElementById('text-tool-btn') as HTMLButtonElement;
  for (const btn of [addBtn, textBtn]) {
    btn.disabled = reason !== null;
    if (reason) btn.title = reason;
  }

  const canEdit = ctx.info.permissions.canEdit;
  const tokenQuery = ctx.tokenQuery;
  const markup = {
    visible: true,
    drawMode: false,
    items: [] as Item[],
    dragging: null as { frame: HTMLElement; page: number; fx: number; fy: number; rectEl: HTMLElement; box: Box | null } | null,
    draftRect: null as HTMLElement | null,
  };
  const textTool = {
    active: false,
    pending: null as PendingText | null,
    lastFontSizePt: 12,
    ghost: null as HTMLElement | null,
    hint: null as HTMLElement | null,
  };
  let textDragState: { item: TextItem; el: HTMLElement; frame: HTMLElement; moved: boolean; x: number; y: number } | null = null;
  const resizeTimers = new WeakMap<HTMLElement, number>();

  const pagesContainer = (): HTMLElement => ctx.body.container;

  /** 화면 좌표 → 쪽 틀 안의 비율. 보기 회전(CSS rotate)은 되돌려 계산한다 — 확대는 틀 크기에 이미 반영돼 있다 */
  const toFrame = (frame: HTMLElement, clientX: number, clientY: number): { fx: number; fy: number } =>
    toFrameRatio(frame.getBoundingClientRect(), frame.offsetWidth, frame.offsetHeight, ctx.body.rotateDeg, clientX, clientY);

  const markupLayer = (frame: HTMLElement): HTMLElement => {
    let layer = frame.querySelector<HTMLElement>('.markup-layer');
    if (!layer) {
      layer = document.createElement('div');
      layer.className = 'markup-layer';
      frame.append(layer);
    }
    return layer;
  };

  const clearMarkupOverlays = (): void => {
    for (const frame of pagesContainer().children) frame.querySelector('.markup-layer')?.replaceChildren();
  };

  const updateMarkupIncludeVisibility = (): void => {
    (document.getElementById('markup-include-wrap') as HTMLElement).hidden = markup.items.length === 0;
  };

  const deleteMarkupItem = async (item: Item, rectEl: HTMLElement): Promise<void> => {
    if (!window.confirm('이 마크업을 삭제할까요?')) return;
    try {
      await fetchJson(`/api/view/markup/${item.id}?${tokenQuery}`, { method: 'DELETE' });
      markup.items = markup.items.filter((it) => it.id !== item.id);
      rectEl.remove();
      updateMarkupIncludeVisibility();
    } catch (error) {
      window.alert(errText(error, '삭제에 실패했습니다.'));
    }
  };

  const applyFontSizeToEl = (el: HTMLElement, frame: HTMLElement, page: number, fontSizePt: number): void => {
    el.style.fontSize = `${fontSizePt * (frame.clientHeight / ctx.body.pageSize(page).h)}px`;
  };

  const persistFontSize = async (item: TextItem, fontSizePt: number): Promise<void> => {
    try {
      const updated = await fetchJson<TextItem>(`/api/view/markup/${item.id}?${tokenQuery}`, {
        method: 'PATCH',
        headers: JSON_HEADERS,
        body: JSON.stringify({ x: item.x, y: item.y, fontSizePt }),
      });
      item.fontSizePt = updated.fontSizePt;
    } catch (error) {
      window.alert(errText(error, '글자 크기 변경에 실패했습니다.'));
    }
  };

  /** 클릭/휠 조작마다 매번 저장 요청을 보내지 않도록 항목별로 살짝 늦춰서 한 번만 저장한다. */
  const scheduleFontSizePersist = (el: HTMLElement, item: TextItem, fontSizePt: number): void => {
    clearTimeout(resizeTimers.get(el));
    resizeTimers.set(el, window.setTimeout(() => void persistFontSize(item, fontSizePt), 350));
  };

  const startTextDrag = (event: MouseEvent, item: TextItem, el: HTMLElement, frame: HTMLElement): void => {
    if ((event.target as Element).closest('.text-overlay-delete')) return;
    event.preventDefault();
    event.stopPropagation();
    textDragState = { item, el, frame, moved: false, x: item.x, y: item.y };
    el.classList.add('dragging');
  };

  const renderTextOverlayItem = (frame: HTMLElement, layer: HTMLElement, item: TextItem): void => {
    const el = document.createElement('div');
    el.className = 'text-overlay-item';
    el.textContent = item.text;
    el.dataset.markupId = item.id;
    el.style.left = `${item.x * 100}%`;
    el.style.top = `${item.y * 100}%`;
    el.style.fontSize = `${item.fontSizePt * (frame.clientHeight / ctx.body.pageSize(item.page).h)}px`;
    el.style.fontWeight = item.bold ? '700' : '400';
    el.style.fontFamily = cssFontFamilyFor(item.fontFamily);

    if (canEdit) {
      const controls = document.createElement('div');
      controls.className = 'text-overlay-controls';
      const resizeBy = (delta: number): void => {
        const next = clampFontSizePt(item.fontSizePt + delta);
        item.fontSizePt = next;
        applyFontSizeToEl(el, frame, item.page, next);
        scheduleFontSizePersist(el, item, next);
      };
      const shrink = button('text-overlay-resize', '−', '글자 작게');
      shrink.addEventListener('click', (event) => {
        event.stopPropagation();
        resizeBy(-FONT_SIZE_STEP_PT);
      });
      const grow = button('text-overlay-resize', '+', '글자 크게');
      grow.addEventListener('click', (event) => {
        event.stopPropagation();
        resizeBy(FONT_SIZE_STEP_PT);
      });
      const del = button('text-overlay-delete', '×', '삭제');
      del.addEventListener('click', (event) => {
        event.stopPropagation();
        void deleteMarkupItem(item, el);
      });
      controls.append(shrink, grow, del);
      el.append(controls);
      el.addEventListener('mousedown', (event) => startTextDrag(event, item, el, frame));
      el.addEventListener(
        'wheel',
        (event) => {
          if (!event.ctrlKey) return;
          event.preventDefault();
          event.stopPropagation();
          resizeBy(event.deltaY < 0 ? FONT_SIZE_STEP_PT : -FONT_SIZE_STEP_PT);
        },
        { passive: false },
      );
    }
    layer.append(el);
  };

  const renderMarkupItem = (item: Item): void => {
    const frame = ctx.body.pageFrame(item.page);
    if (!frame) return;
    const layer = markupLayer(frame);
    if (item.kind === 'text') {
      renderTextOverlayItem(frame, layer, item);
      return;
    }
    const rectEl = document.createElement('div');
    rectEl.className = 'markup-rect';
    rectEl.style.left = `${item.x * 100}%`;
    rectEl.style.top = `${item.y * 100}%`;
    rectEl.style.width = `${item.w * 100}%`;
    rectEl.style.height = `${item.h * 100}%`;
    rectEl.tabIndex = 0;
    rectEl.dataset.markupId = item.id;

    const tooltip = document.createElement('div');
    tooltip.className = 'markup-tooltip';
    const metaLine = document.createElement('div');
    metaLine.className = 'markup-tooltip-meta';
    metaLine.textContent = `${item.author} · ${new Date(item.createdAt).toLocaleString('ko-KR')}`;
    const body = document.createElement('div');
    body.className = 'markup-tooltip-text';
    body.textContent = item.text;
    tooltip.append(metaLine, body);

    if (canEdit) {
      const del = button('markup-tooltip-delete', '삭제');
      del.addEventListener('click', (event) => {
        event.stopPropagation();
        void deleteMarkupItem(item, rectEl);
      });
      tooltip.append(del);
    }
    rectEl.append(tooltip);
    layer.append(rectEl);
  };

  const renderAllMarkup = (): void => {
    clearMarkupOverlays();
    if (!markup.visible) return;
    for (const item of markup.items) renderMarkupItem(item);
  };

  const loadMarkup = async (): Promise<void> => {
    try {
      markup.items = await fetchJson<Item[]>(`/api/view/markup?${tokenQuery}`);
      renderAllMarkup();
      updateMarkupIncludeVisibility();
    } catch {
      /* 마크업 로딩 실패는 열람 자체를 막지 않는다 */
    }
  };

  const closeMarkupPopover = (): void => {
    document.querySelector('.markup-popover')?.remove();
  };

  const cancelMarkupDraft = (): void => {
    if (markup.dragging) {
      markup.dragging.rectEl.remove();
      markup.dragging = null;
    }
    if (markup.draftRect) {
      markup.draftRect.remove();
      markup.draftRect = null;
    }
    closeMarkupPopover();
  };

  const setMarkupDrawMode = (on: boolean): void => {
    markup.drawMode = on;
    addBtn.setAttribute('aria-pressed', String(on));
    pagesContainer().classList.toggle('markup-drawing', on);
    if (on) setTextToolActive(false);
    if (!on) cancelMarkupDraft();
  };

  const openMarkupPopover = (frame: HTMLElement, page: number, box: Box): void => {
    closeMarkupPopover();
    const popover = document.createElement('div');
    popover.className = 'markup-popover';
    popover.style.left = `${Math.min(box.x + box.w, 0.7) * 100}%`;
    popover.style.top = `${box.y * 100}%`;

    const textarea = document.createElement('textarea');
    textarea.placeholder = '코멘트를 입력하세요';
    textarea.maxLength = 2000;

    const actions = document.createElement('div');
    actions.className = 'markup-popover-actions';
    const cancelBtn = button('btn-pill', '취소');
    const saveBtn = button('btn-pill btn-pill-primary', '저장');
    actions.append(cancelBtn, saveBtn);
    popover.append(textarea, actions);
    frame.append(popover);
    textarea.focus();

    cancelBtn.addEventListener('click', () => cancelMarkupDraft());
    saveBtn.addEventListener('click', async () => {
      const text = textarea.value.trim();
      if (!text) {
        textarea.focus();
        return;
      }
      saveBtn.disabled = true;
      try {
        const item = await fetchJson<CommentItem>(`/api/view/markup?${tokenQuery}`, {
          method: 'POST',
          headers: JSON_HEADERS,
          body: JSON.stringify({ page, x: box.x, y: box.y, w: box.w, h: box.h, text }),
        });
        markup.items.push(item);
        cancelMarkupDraft();
        renderAllMarkup();
        updateMarkupIncludeVisibility();
      } catch (error) {
        window.alert(errText(error, '마크업 저장에 실패했습니다.'));
        saveBtn.disabled = false;
      }
    });
  };

  /* ---------- 텍스트 추가(주변 폰트에 맞춰 삽입) ---------- */

  const setTextToolActive = (on: boolean): void => {
    textTool.active = on;
    textBtn.setAttribute('aria-pressed', String(on));
    pagesContainer().classList.toggle('text-tool-active', on);
    if (on) setMarkupDrawMode(false);
    if (!on) void commitTextInsert(); // 도구를 끌 때 쓰던 글은 저장한다(버리지 않는다)
    if (textTool.ghost) textTool.ghost.hidden = true;
    // 도구를 고른 동안 아래쪽에 쓰는 법을 띄운다 — 글을 넣은 뒤에도 도구가 남아 있어 지금 상태를 알 수 있게
    if (!textTool.hint) {
      textTool.hint = Object.assign(document.createElement('div'), {
        className: 'text-tool-hint',
        textContent: '텍스트 추가 — 누른 곳에 바로 입력 · Enter 로 마침 · 다 쓰면 Esc',
      });
      textTool.hint.setAttribute('role', 'status');
      document.body.append(textTool.hint);
    }
    textTool.hint.hidden = !on;
  };

  /** 마우스를 댄 자리에 글이 들어갈 줄(입력 막대 + '누르고 입력')을 미리 보인다 — 글은 누른 점 위에 앉는다 */
  const showTextGhost = (event: MouseEvent): void => {
    const target = event.target as Element;
    const frame = textTool.active && !textTool.pending ? frameOf(event) : null;
    if (!frame || target.closest('.text-overlay-item, .text-insert-wrapper, .markup-rect')) {
      if (textTool.ghost) textTool.ghost.hidden = true;
      return;
    }
    if (!textTool.ghost) textTool.ghost = Object.assign(document.createElement('div'), { className: 'text-tool-ghost' });
    const ghost = textTool.ghost;
    if (ghost.parentElement !== frame) frame.append(ghost);
    const { fx, fy } = toFrame(frame, event.clientX, event.clientY);
    applyFontSizeToEl(ghost, frame, Number(frame.dataset.page), textTool.lastFontSizePt);
    const h = parseFloat(ghost.style.fontSize) * 1.15 + 2;
    ghost.style.left = `${fx * frame.offsetWidth}px`;
    ghost.style.top = `${fy * frame.offsetHeight - h}px`;
    ghost.style.height = `${h}px`;
    ghost.hidden = false;
  };

  const cancelTextInsert = (): void => {
    if (textTool.pending) {
      textTool.pending.wrapper.remove();
      textTool.pending = null;
    }
  };

  const adjustPendingFontSize = (delta: number): void => {
    const pending = textTool.pending;
    if (!pending) return;
    pending.style.fontSizePt = clampFontSizePt(pending.style.fontSizePt + delta);
    applyFontSizeToEl(pending.box, pending.frame, pending.page, pending.style.fontSizePt);
    pending.box.focus();
  };

  async function commitTextInsert(): Promise<void> {
    const pending = textTool.pending;
    if (!pending) return;
    textTool.pending = null;
    const text = pending.box.value.trim();
    pending.wrapper.remove();
    if (!text) return;
    textTool.lastFontSizePt = pending.style.fontSizePt;
    try {
      const item = await fetchJson<TextItem>(`/api/view/markup?${tokenQuery}`, {
        method: 'POST',
        headers: JSON_HEADERS,
        body: JSON.stringify({
          kind: 'text',
          page: pending.page,
          x: pending.xRatio,
          y: pending.style.baselineY,
          fontSizePt: pending.style.fontSizePt,
          bold: pending.style.bold,
          fontFamily: pending.style.fontFamily,
          text,
        }),
      });
      markup.items.push(item);
      renderAllMarkup();
      updateMarkupIncludeVisibility();
    } catch (error) {
      window.alert(errText(error, '텍스트 추가에 실패했습니다.'));
    }
  }

  const openTextInsertBox = (frame: HTMLElement, page: number, xRatio: number, style: TextStyle): void => {
    cancelTextInsert();
    const wrapper = document.createElement('div');
    wrapper.className = 'text-insert-wrapper';
    wrapper.style.left = `${xRatio * 100}%`;
    wrapper.style.top = `${style.baselineY * 100}%`;

    const box = document.createElement('textarea');
    box.className = 'text-insert-input';
    box.rows = 1;
    box.maxLength = 2000;
    box.style.fontWeight = style.bold ? '700' : '400';
    box.style.fontFamily = cssFontFamilyFor(style.fontFamily);
    box.style.width = '2ch';
    applyFontSizeToEl(box, frame, page, style.fontSizePt);

    const controls = document.createElement('div');
    controls.className = 'text-overlay-controls text-insert-controls';
    const shrink = button('text-overlay-resize', '−', '글자 작게');
    shrink.addEventListener('mousedown', (event) => event.preventDefault());
    shrink.addEventListener('click', () => adjustPendingFontSize(-FONT_SIZE_STEP_PT));
    const grow = button('text-overlay-resize', '+', '글자 크게');
    grow.addEventListener('mousedown', (event) => event.preventDefault());
    grow.addEventListener('click', () => adjustPendingFontSize(FONT_SIZE_STEP_PT));
    controls.append(shrink, grow);

    wrapper.append(box, controls);
    frame.append(wrapper);
    box.focus();

    box.addEventListener('input', () => {
      box.style.width = `${Math.max(2, box.value.length + 1)}ch`;
    });
    box.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        void commitTextInsert();
      } else if (event.key === 'Escape') {
        // 첫 Esc 는 쓰던 글을 저장하고 마친다(도구는 그대로) — 두 번째 Esc 가 도구를 끈다
        event.preventDefault();
        event.stopPropagation();
        void commitTextInsert();
      }
    });
    box.addEventListener('blur', () => void commitTextInsert());
    box.addEventListener(
      'wheel',
      (event) => {
        if (!event.ctrlKey) return;
        event.preventDefault();
        adjustPendingFontSize(event.deltaY < 0 ? FONT_SIZE_STEP_PT : -FONT_SIZE_STEP_PT);
      },
      { passive: false },
    );
    textTool.pending = { wrapper, box, frame, page, xRatio, style };
  };

  if (!reason) {
    addBtn.addEventListener('click', () => setMarkupDrawMode(!markup.drawMode));
    const pages = pagesContainer();

    pages.addEventListener('mousedown', (event) => {
      if (!markup.drawMode) return;
      const frame = frameOf(event);
      if (!frame) return;
      event.preventDefault();
      const { fx, fy } = toFrame(frame, event.clientX, event.clientY);
      const rectEl = document.createElement('div');
      rectEl.className = 'markup-rect markup-rect-draft';
      markupLayer(frame).append(rectEl);
      markup.dragging = { frame, page: Number(frame.dataset.page), fx, fy, rectEl, box: null };
    });

    document.addEventListener('mousemove', (event) => {
      if (markup.dragging) {
        const { frame, fx: startFx, fy: startFy, rectEl } = markup.dragging;
        const cur = toFrame(frame, event.clientX, event.clientY);
        const curFx = clamp01(cur.fx);
        const curFy = clamp01(cur.fy);
        const x = Math.min(startFx, curFx);
        const y = Math.min(startFy, curFy);
        const w = Math.abs(curFx - startFx);
        const h = Math.abs(curFy - startFy);
        rectEl.style.left = `${x * 100}%`;
        rectEl.style.top = `${y * 100}%`;
        rectEl.style.width = `${w * 100}%`;
        rectEl.style.height = `${h * 100}%`;
        markup.dragging.box = { x, y, w, h };
      }
      if (textDragState) {
        textDragState.moved = true;
        const { frame, el } = textDragState;
        const cur = toFrame(frame, event.clientX, event.clientY);
        const x = clamp01(cur.fx);
        const y = clamp01(cur.fy);
        el.style.left = `${x * 100}%`;
        el.style.top = `${y * 100}%`;
        textDragState.x = x;
        textDragState.y = y;
      }
    });

    document.addEventListener('mouseup', async () => {
      if (markup.dragging) {
        const { frame, page, rectEl, box } = markup.dragging;
        markup.dragging = null;
        const MIN_SIZE = 0.01;
        if (!box || box.w < MIN_SIZE || box.h < MIN_SIZE) {
          rectEl.remove();
        } else {
          markup.draftRect = rectEl;
          openMarkupPopover(frame, page, box);
        }
      }
      if (textDragState) {
        const { item, el, moved, x, y } = textDragState;
        textDragState = null;
        el.classList.remove('dragging');
        if (!moved) return;
        try {
          const updated = await fetchJson<TextItem>(`/api/view/markup/${item.id}?${tokenQuery}`, {
            method: 'PATCH',
            headers: JSON_HEADERS,
            body: JSON.stringify({ x, y }),
          });
          item.x = updated.x;
          item.y = updated.y;
        } catch (error) {
          // 실패하면 원래 위치로 되돌린다
          el.style.left = `${item.x * 100}%`;
          el.style.top = `${item.y * 100}%`;
          window.alert(errText(error, '위치 이동에 실패했습니다.'));
        }
      }
    });

    textBtn.addEventListener('click', () => setTextToolActive(!textTool.active));
    pages.addEventListener('mousemove', (event) => showTextGhost(event));
    pages.addEventListener('mouseleave', () => textTool.ghost && (textTool.ghost.hidden = true));
    pages.addEventListener('click', async (event) => {
      if (!textTool.active) return;
      const target = event.target as Element;
      if (target.closest('.text-overlay-item') || target.closest('.text-insert-wrapper')) return;
      const frame = frameOf(event);
      if (!frame) return;
      // 쓰던 글이 있으면 저장하고, 누른 자리에서 바로 새 글을 시작한다(빈칸 여러 곳을 연달아 채우게)
      if (textTool.pending) void commitTextInsert();
      if (textTool.ghost) textTool.ghost.hidden = true;
      const page = Number(frame.dataset.page);
      const { fx: xRatio, fy: yRatio } = toFrame(frame, event.clientX, event.clientY);
      let style: TextStyle;
      try {
        style = await fetchJson<TextStyle>(`/api/view/pages/${page}/text-style?${tokenQuery}&x=${xRatio}&y=${yRatio}`);
      } catch {
        style = { fontSizePt: 12, bold: false, fontFamily: 'sans-serif', baselineY: yRatio, matched: false };
      }
      openTextInsertBox(frame, page, xRatio, style);
    });

    ctx.onEscape(20, () => {
      if (!textTool.pending) return false;
      void commitTextInsert();
      return true;
    });
    ctx.onEscape(30, () => {
      if (!textTool.active) return false;
      setTextToolActive(false);
      return true;
    });
    ctx.onEscape(40, () => {
      if (!markup.drawMode) return false;
      setMarkupDrawMode(false);
      return true;
    });
  }

  // 확대·창 크기로 쪽 높이가 바뀌면 넣은 글(과 쓰던 글)의 글자 크기를 쪽에 다시 맞춘다
  const refitTextFonts = (): void => {
    for (const el of pagesContainer().querySelectorAll<HTMLElement>('.text-overlay-item')) {
      const item = markup.items.find((it) => it.id === el.dataset.markupId);
      const frame = el.closest<HTMLElement>('.page-frame');
      if (item?.kind === 'text' && frame) applyFontSizeToEl(el, frame, item.page, item.fontSizePt);
    }
    const pending = textTool.pending;
    if (pending) applyFontSizeToEl(pending.box, pending.frame, pending.page, pending.style.fontSizePt);
  };
  ctx.body.onZoom(refitTextFonts);
  window.addEventListener('resize', refitTextFonts);

  // 쪽 틀이 다시 만들어지면(pdf.js → PNG) 겹침을 다시 그린다
  ctx.body.onSourceChange(() => {
    closeMarkupPopover();
    if (textTool.pending) cancelTextInsert();
    renderAllMarkup();
  });
  void loadMarkup();
}
