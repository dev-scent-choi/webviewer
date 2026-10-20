/**
 * 기존 글자 고치기·가리기(5단계) 편집기 층.
 *  - 글자 고치기 도구: 본문 글 조각(pdf.js 글자 층 span)을 누르면 그 자리에 입력란 — Enter로 고치고 Esc로 취소. 고친 글은 방의
 *    pdfTextEdits에 쓰고(실행 취소 대상), 화면에는 원래 글을 흰 상자로 덮고 새 글을 미리 보여 준다. 저장하면 서버가 원래 글자를 실제로
 *    지우고 같은 자리·글꼴(가능하면)로 새 글을 쓴다(src/pdf/textRewrite.ts).
 *  - 가리기 도구: 끌어서 영역 표시(빨간 테두리) → "가리기 적용"으로 확정. 적용한 영역은 검은 상자, 그 아래 글자는 고르거나 찾을 수 없게
 *    숨기고, 저장 파일에서는 그 내용을 실제로 지운다. 적용은 되돌릴 수 없다(실행 취소 기록도 비운다).
 */
import type * as Y from 'yjs';
import { isRedaction, isTextEdit, type PdfRedaction, type PdfTextEdit } from '../../pdf/annotTypes.js';
import { LOCAL_ORIGIN } from './annotLayer.js';
import type { PageView, PageViewport, TextItemInfo } from './pageView.js';

type Rect = [number, number, number, number];
const APPLY_ORIGIN = Symbol('pdf-redact-apply');

export interface EditLayerOptions {
  readonly view: PageView;
  readonly ydoc: Y.Doc;
  readonly edits: Y.Map<PdfTextEdit>;
  readonly redactions: Y.Map<PdfRedaction>;
  readonly canEdit: boolean;
  readonly author: string;
  readonly scroller: HTMLElement;
  readonly undoManager: Y.UndoManager;
  readonly notice: (text: string) => void;
}

/** 글 조각의 PDF 사각형 — 기준선 방향을 따라 너비, 위로 높이(아래로 조금) */
export function itemRect(it: TextItemInfo): Rect {
  const [a, b, c, d, e, f] = it.transform;
  const len = Math.hypot(a, b) || 1;
  const ux = a / len;
  const uy = b / len;
  const vx = -uy;
  const vy = ux;
  const h = it.height || Math.hypot(c, d) || 10;
  const lo = -0.22 * h;
  const hi = 0.92 * h;
  const pts = [
    [e + vx * lo, f + vy * lo],
    [e + ux * it.width + vx * lo, f + uy * it.width + vy * lo],
    [e + vx * hi, f + vy * hi],
    [e + ux * it.width + vx * hi, f + uy * it.width + vy * hi],
  ];
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}

const overlaps = (a: readonly number[], b: readonly number[]) => a[0] < b[2] && a[2] > b[0] && a[1] < b[3] && a[3] > b[1];
const editId = (pageKey: string, index: number) => `t:${pageKey}:${index}`;
const indexOfEdit = (id: string) => {
  const m = /:(\d+)$/.exec(id);
  return m ? Number(m[1]) : -1;
};

export class EditLayer {
  private editor: HTMLElement | null = null;
  private readonly countListeners: Array<(pending: number) => void> = [];

  constructor(private readonly opts: EditLayerOptions) {
    opts.view.onPageRendered((page, _l, vp) => void this.renderPage(page, vp));
    const rerender = () => {
      for (const page of opts.view.renderedPages()) {
        const vp = opts.view.viewportOf(page);
        if (vp) void this.renderPage(page, vp);
      }
      for (const cb of this.countListeners) cb(this.pendingCount());
    };
    opts.edits.observe(rerender);
    opts.redactions.observe(rerender);
    this.bindPointer();
  }

  /** 아직 적용하지 않은 가리기 표시 수 */
  pendingCount(): number {
    return [...this.opts.redactions.values()].filter((r) => isRedaction(r) && !r.applied).length;
  }
  onPendingChange(cb: (pending: number) => void): void {
    this.countListeners.push(cb);
  }

  /** 표시한 가리기를 모두 적용 — 되돌릴 수 없다 */
  applyRedactions(): number {
    const pending = [...this.opts.redactions.values()].filter((r) => isRedaction(r) && !r.applied);
    if (pending.length === 0) return 0;
    this.opts.ydoc.transact(() => {
      for (const r of pending) this.opts.redactions.set(r.id, { ...r, applied: true });
    }, APPLY_ORIGIN);
    // 표시를 만든 기록을 되돌리면 적용한 가리기가 사라지므로 실행 취소 기록을 비운다
    this.opts.undoManager.clear();
    return pending.length;
  }

  private tool(): string {
    return this.opts.scroller.dataset.tool ?? 'select';
  }

  // ---------- 그리기 ----------
  private async renderPage(page: number, vp: PageViewport): Promise<void> {
    const { view } = this.opts;
    const el = view.pageElement(page);
    const layer = el?.querySelector<HTMLElement>('.pdf-edit-layer');
    if (!layer) return;
    const key = view.pageKeyOf(page);
    const edits = [...this.opts.edits.values()].filter((e) => isTextEdit(e) && e.pageKey === key);
    const redactions = [...this.opts.redactions.values()].filter((r) => isRedaction(r) && r.pageKey === key);
    layer.replaceChildren();
    const box = (r: readonly number[]) => {
      const [x1, y1, x2, y2] = vp.convertToViewportRectangle([...r]) as number[];
      return { left: Math.min(x1, x2), top: Math.min(y1, y2), width: Math.abs(x2 - x1), height: Math.abs(y2 - y1) };
    };
    const place = (div: HTMLElement, r: readonly number[]) => {
      const b = box(r);
      div.style.left = `${b.left}px`;
      div.style.top = `${b.top}px`;
      div.style.width = `${b.width}px`;
      div.style.height = `${b.height}px`;
      return b;
    };
    // 원래 글자 층 — 고친 조각·가린 영역의 글자는 고르거나 찾을 수 없게
    const divs = view.textDivsOf(page);
    for (const d of divs) d.classList.remove('pdf-text-hidden');
    for (const e of edits) divs[indexOfEdit(e.id)]?.classList.add('pdf-text-hidden');
    const applied = redactions.filter((r) => r.applied);
    if (applied.length && divs.length) {
      const items = await view.textItemInfos(page);
      items.forEach((it, i) => {
        if (applied.some((r) => overlaps(itemRect(it), r.rect))) divs[i]?.classList.add('pdf-text-hidden');
      });
    }
    for (const e of edits) {
      const div = document.createElement('div');
      div.className = 'pdf-edit-text';
      div.dataset.edit = e.id;
      const b = place(div, e.rect);
      const span = document.createElement('span');
      span.textContent = e.newText;
      span.style.fontSize = `${Math.max(4, e.fontSize * vp.scale)}px`;
      span.style.lineHeight = `${b.height}px`;
      div.title = `원래: ${e.oldText}`;
      div.append(span);
      layer.append(div);
    }
    for (const r of redactions) {
      const div = document.createElement('div');
      div.className = r.applied ? 'pdf-redact applied' : 'pdf-redact';
      div.dataset.redact = r.id;
      place(div, r.rect);
      if (!r.applied && this.opts.canEdit) {
        const del = document.createElement('button');
        del.type = 'button';
        del.className = 'pdf-redact-del';
        del.title = '표시 지우기';
        del.setAttribute('aria-label', '가리기 표시 지우기');
        del.textContent = '×';
        del.addEventListener('click', (event) => {
          event.stopPropagation();
          this.opts.ydoc.transact(() => this.opts.redactions.delete(r.id), LOCAL_ORIGIN);
        });
        div.append(del);
      }
      layer.append(div);
    }
  }

  // ---------- 누르기·끌기 ----------
  private bindPointer(): void {
    const { scroller, view } = this.opts;
    scroller.addEventListener('pointerdown', (event) => {
      if (!this.opts.canEdit || event.button !== 0) return;
      const tool = this.tool();
      const target = event.target as HTMLElement;
      if (target.closest('.pdf-edit-input, .pdf-redact-del')) return;
      if (tool === 'textedit') {
        // 고친 글(미리 보기)을 누르면 다시 고치기, 아니면 그 자리의 글 조각
        const shown = target.closest<HTMLElement>('[data-edit]');
        const pageEl = target.closest<HTMLElement>('.pdf-page');
        if (!pageEl) return;
        const page = Number(pageEl.dataset.page);
        let index = shown ? indexOfEdit(shown.dataset.edit ?? '') : -1;
        if (index < 0) {
          const span = document
            .elementsFromPoint(event.clientX, event.clientY)
            .find((e) => e.parentElement?.classList.contains('pdf-text-layer') || e.closest('.pdf-text-layer'));
          const hit = span ? view.textSpanIndex(span.closest('.pdf-text-layer > span') ?? span) : null;
          if (!hit || hit.page !== page) return;
          index = hit.index;
        }
        event.preventDefault();
        void this.openEditor(page, index);
      } else if (tool === 'redact') {
        const pageEl = target.closest<HTMLElement>('.pdf-page');
        if (!pageEl) return;
        event.preventDefault();
        this.dragRedaction(event, pageEl);
      }
    });
  }

  private dragRedaction(event: PointerEvent, pageEl: HTMLElement): void {
    const page = Number(pageEl.dataset.page);
    const vp = this.opts.view.viewportOf(page);
    if (!vp) return;
    const origin = pageEl.getBoundingClientRect();
    const start = [event.clientX - origin.left, event.clientY - origin.top];
    const preview = document.createElement('div');
    preview.className = 'pdf-redact preview';
    pageEl.querySelector('.pdf-edit-layer')?.append(preview);
    let last = start;
    const move = (e: PointerEvent) => {
      last = [e.clientX - origin.left, e.clientY - origin.top];
      preview.style.left = `${Math.min(start[0], last[0])}px`;
      preview.style.top = `${Math.min(start[1], last[1])}px`;
      preview.style.width = `${Math.abs(last[0] - start[0])}px`;
      preview.style.height = `${Math.abs(last[1] - start[1])}px`;
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      preview.remove();
      if (Math.abs(last[0] - start[0]) < 4 || Math.abs(last[1] - start[1]) < 4) return;
      const p1 = vp.convertToPdfPoint(start[0], start[1]) as number[];
      const p2 = vp.convertToPdfPoint(last[0], last[1]) as number[];
      const rect: Rect = [Math.min(p1[0], p2[0]), Math.min(p1[1], p2[1]), Math.max(p1[0], p2[0]), Math.max(p1[1], p2[1])];
      const id = `r${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
      const r: PdfRedaction = { id, pageKey: this.opts.view.pageKeyOf(page), rect, applied: false, author: this.opts.author, createdAt: new Date().toISOString() };
      this.opts.ydoc.transact(() => this.opts.redactions.set(id, r), LOCAL_ORIGIN);
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }

  private closeEditor(): void {
    this.editor?.remove();
    this.editor = null;
  }

  private async openEditor(page: number, index: number): Promise<void> {
    this.closeEditor();
    const { view } = this.opts;
    const items = await view.textItemInfos(page);
    const it = items[index];
    const vp = view.viewportOf(page);
    const pageEl = view.pageElement(page);
    if (!it || !it.str.trim() || !vp || !pageEl) return;
    const key = view.pageKeyOf(page);
    const id = editId(key, index);
    const existing = this.opts.edits.get(id);
    const rect = itemRect(it);
    const [x1, y1, x2, y2] = vp.convertToViewportRectangle(rect) as number[];
    const wrap = document.createElement('div');
    wrap.className = 'pdf-edit-input';
    wrap.style.left = `${Math.min(x1, x2) - 3}px`;
    wrap.style.top = `${Math.min(y1, y2) - 3}px`;
    const input = document.createElement('input');
    input.type = 'text';
    input.value = isTextEdit(existing) ? existing.newText : it.str;
    input.setAttribute('aria-label', '고칠 글');
    const fontPx = Math.max(10, (it.height || 10) * vp.scale);
    input.style.fontSize = `${fontPx}px`;
    input.style.height = `${Math.abs(y2 - y1) + 6}px`;
    const minW = Math.abs(x2 - x1) + 6;
    const fit = () => (input.style.width = `${Math.max(minW, input.value.length * fontPx * 0.62 + 24)}px`);
    fit();
    input.addEventListener('input', fit);
    const actions = document.createElement('span');
    actions.className = 'pdf-edit-actions';
    const ok = document.createElement('button');
    ok.type = 'button';
    ok.textContent = '고치기';
    const revert = document.createElement('button');
    revert.type = 'button';
    revert.textContent = '원래대로';
    revert.hidden = !isTextEdit(existing);
    actions.append(ok, revert);
    wrap.append(input, actions);
    pageEl.append(wrap);
    this.editor = wrap;
    input.focus();
    input.select();
    let done = false;
    const commit = () => {
      if (done) return;
      done = true;
      const text = input.value;
      this.closeEditor();
      if (text === it.str || text === '') {
        // 원래 글과 같으면(또는 비우면) 고친 것을 없앤다
        if (isTextEdit(existing) && text === it.str) this.opts.ydoc.transact(() => this.opts.edits.delete(id), LOCAL_ORIGIN);
        if (text === '') this.opts.notice('빈 글로는 고칠 수 없습니다 — 글자를 지우려면 가리기를 쓰세요.');
        return;
      }
      const edit: PdfTextEdit = { id, pageKey: key, rect, oldText: it.str, newText: text, fontSize: it.height || Math.hypot(it.transform[2], it.transform[3]) || 10, author: this.opts.author, createdAt: new Date().toISOString() };
      this.opts.ydoc.transact(() => this.opts.edits.set(id, edit), LOCAL_ORIGIN);
    };
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        commit();
      } else if (e.key === 'Escape') {
        e.preventDefault();
        done = true;
        this.closeEditor();
      }
    });
    input.addEventListener('blur', () => setTimeout(() => !done && this.editor === wrap && commit(), 120));
    ok.addEventListener('pointerdown', (e) => e.preventDefault());
    ok.addEventListener('click', commit);
    revert.addEventListener('pointerdown', (e) => e.preventDefault());
    revert.addEventListener('click', () => {
      done = true;
      this.closeEditor();
      this.opts.ydoc.transact(() => this.opts.edits.delete(id), LOCAL_ORIGIN);
    });
  }
}
