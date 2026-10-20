/**
 * PDF 편집기 주석 층 — 협업 맵(pdfAnnots)의 주석을 쪽마다 그리고(SVG + 글상자·메모 DOM), 도구로 만들고 고친다.
 *
 * 좌표: 주석은 원래 쪽의 PDF 좌표(annotTypes.ts). 화면 좌표와는 그 쪽 viewport(pdf.js)의 convertToViewportPoint /
 * convertToPdfPoint로만 바꾼다 — 보기 회전·배율이 바뀌어도 같은 식으로 맞는다. 쪽이 다시 그려질 때마다(pageView.onPageRendered)
 * 그 쪽의 주석을 다시 그린다.
 *
 * 쓰기는 모두 LOCAL_ORIGIN 트랜잭션 — 실행 취소(Y.UndoManager)가 내 편집만 되돌리고, 서버는 그 연결의 편집으로 기록한다.
 * 끄는 동안(이동·크기·그리기)은 화면에서만 바꾸고 손을 뗄 때 한 번 쓴다(되돌리기 한 번에 한 동작).
 */
import type * as Y from 'yjs';
import { ANNOT_DEFAULTS, ANNOT_LABELS, type PdfAnnot, type PdfAnnotType } from '../../pdf/annotTypes.js';
import type { PageView, PageViewport } from './pageView.js';

export type ToolName = 'select' | 'hand' | 'eraser' | PdfAnnotType;

export interface ToolStyle {
  color: string;
  width: number;
  opacity: number;
}

export const LOCAL_ORIGIN = Symbol('pdf-local-edit');

const SVG_NS = 'http://www.w3.org/2000/svg';
const DRAW_TOOLS = new Set<ToolName>(['note', 'ink', 'freetext', 'rect', 'ellipse', 'line', 'arrow']);
const TEXT_TOOLS = new Set<ToolName>(['highlight', 'underline', 'strike']);
const BOX_TYPES = new Set<PdfAnnotType>(['rect', 'ellipse', 'freetext']);
const NOTE_SIZE = 24;

function svg<K extends keyof SVGElementTagNameMap>(tag: K, attrs: Record<string, string | number> = {}): SVGElementTagNameMap[K] {
  const el = document.createElementNS(SVG_NS, tag);
  for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, String(v));
  return el;
}

const pt = (vp: PageViewport, x: number, y: number) => vp.convertToViewportPoint(x, y) as [number, number];
const toPdf = (vp: PageViewport, x: number, y: number) => vp.convertToPdfPoint(x, y) as [number, number];

/** 주석이 차지하는 PDF 좌표 점들 */
function pointsOf(a: PdfAnnot): number[] {
  if (a.quads?.length) return [...a.quads];
  if (a.paths?.length) return a.paths.flat();
  if (a.line) return [...a.line];
  return [a.rect[0], a.rect[1], a.rect[2], a.rect[3], a.rect[0], a.rect[3], a.rect[2], a.rect[1]];
}

/** 화면 좌표의 감싸는 사각형 */
function viewBox(vp: PageViewport, a: PdfAnnot): { x: number; y: number; w: number; h: number } {
  const pts = pointsOf(a);
  const xs: number[] = [];
  const ys: number[] = [];
  for (let i = 0; i + 1 < pts.length; i += 2) {
    const [x, y] = pt(vp, pts[i], pts[i + 1]);
    xs.push(x);
    ys.push(y);
  }
  const pad = a.type === 'note' || a.type === 'freetext' ? 0 : (a.width * vp.scale) / 2 + 2;
  const x = Math.min(...xs) - pad;
  const y = Math.min(...ys) - pad;
  return { x, y, w: Math.max(...xs) + pad - x, h: Math.max(...ys) + pad - y };
}

function boundsOfPoints(pts: readonly number[], pad = 0): [number, number, number, number] {
  const xs = pts.filter((_, i) => i % 2 === 0);
  const ys = pts.filter((_, i) => i % 2 === 1);
  return [Math.min(...xs) - pad, Math.min(...ys) - pad, Math.max(...xs) + pad, Math.max(...ys) + pad];
}

/** 모든 좌표를 PDF 단위로 dx, dy만큼 */
function moved(a: PdfAnnot, dx: number, dy: number): PdfAnnot {
  const shift = (arr: readonly number[]) => arr.map((v, i) => v + (i % 2 === 0 ? dx : dy));
  return {
    ...a,
    rect: shift(a.rect) as unknown as PdfAnnot['rect'],
    ...(a.quads ? { quads: shift(a.quads) } : {}),
    ...(a.paths ? { paths: a.paths.map(shift) } : {}),
    ...(a.line ? { line: shift(a.line) as unknown as PdfAnnot['line'] } : {}),
  };
}

/** 점들을 줄인다(화면에서 1.5px 안쪽으로 붙은 점 빼기) */
function simplify(points: number[], minDist: number): number[] {
  if (points.length <= 4) return points;
  const out = [points[0], points[1]];
  for (let i = 2; i + 1 < points.length; i += 2) {
    const lx = out[out.length - 2];
    const ly = out[out.length - 1];
    if (Math.hypot(points[i] - lx, points[i + 1] - ly) >= minDist) out.push(points[i], points[i + 1]);
  }
  const lastX = points[points.length - 2];
  const lastY = points[points.length - 1];
  if (out[out.length - 2] !== lastX || out[out.length - 1] !== lastY) out.push(lastX, lastY);
  return out;
}

function newId(): string {
  return `a${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
}

function formatTime(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  const p = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

export interface AnnotLayerOptions {
  readonly view: PageView;
  readonly map: Y.Map<PdfAnnot>;
  readonly ydoc: Y.Doc;
  readonly canEdit: boolean;
  readonly author: string;
  readonly scroller: HTMLElement;
}

export class AnnotLayer {
  private tool: ToolName = 'select';
  private readonly styles = new Map<PdfAnnotType, ToolStyle>();
  private selected: string | null = null;
  private readonly selectListeners: Array<(annot: PdfAnnot | null) => void> = [];
  private readonly changeListeners: Array<() => void> = [];
  private popup: HTMLElement | null = null;
  private readonly pageChangeListeners: Array<(pages: Set<number>) => void> = [];

  constructor(private readonly opts: AnnotLayerOptions) {
    for (const [type, d] of Object.entries(ANNOT_DEFAULTS) as Array<[PdfAnnotType, ToolStyle]>) this.styles.set(type, { ...d });
    opts.view.onPageRendered((page) => this.renderPage(page));
    opts.map.observe((event) => {
      const pages = new Set<number>();
      event.changes.keys.forEach((change, key) => {
        const now = opts.map.get(key);
        const before = change.oldValue as PdfAnnot | undefined;
        for (const a of [now, before]) {
          const page = a ? opts.view.pageOfKey(a.pageKey) : null;
          if (page) pages.add(page);
        }
        if (change.action === 'delete' && key === this.selected) this.select(null);
      });
      for (const page of pages) this.renderPage(page);
      for (const cb of this.pageChangeListeners) cb(pages);
      if (this.selected) this.notifySelect();
      for (const cb of this.changeListeners) cb();
    });
    this.bindPointer();
    document.addEventListener('mouseup', () => {
      if (TEXT_TOOLS.has(this.tool)) setTimeout(() => this.markSelection(), 0);
    });
    document.addEventListener('keydown', (event) => {
      if (!this.opts.canEdit) return;
      const typing = (event.target as HTMLElement).closest('input, textarea, [contenteditable="true"]');
      if (typing) return;
      if ((event.key === 'Delete' || event.key === 'Backspace') && this.selected) {
        event.preventDefault();
        this.remove(this.selected);
      } else if (event.key === 'Escape') {
        this.closePopup();
        this.select(null);
      }
    });
  }

  // ---------- 상태 ----------
  get currentTool(): ToolName {
    return this.tool;
  }
  setTool(tool: ToolName): void {
    this.tool = tool;
    this.closePopup();
    if (tool !== 'select') this.select(null);
    const root = this.opts.scroller;
    root.dataset.tool = tool;
    root.classList.toggle('hand', tool === 'hand');
  }
  styleOf(type: PdfAnnotType): ToolStyle {
    return this.styles.get(type) ?? { ...ANNOT_DEFAULTS[type] };
  }
  /** 도구 기본값과, 고른 주석이 있으면 그 주석에도 */
  setStyle(type: PdfAnnotType, patch: Partial<ToolStyle>): void {
    this.styles.set(type, { ...this.styleOf(type), ...patch });
    const a = this.selected ? this.opts.map.get(this.selected) : undefined;
    if (a && this.opts.canEdit) this.update(a.id, { ...patch });
  }
  onSelect(cb: (annot: PdfAnnot | null) => void): void {
    this.selectListeners.push(cb);
  }
  onChange(cb: () => void): void {
    this.changeListeners.push(cb);
  }
  get selectedAnnot(): PdfAnnot | null {
    return (this.selected && this.opts.map.get(this.selected)) || null;
  }
  all(): PdfAnnot[] {
    return [...this.opts.map.values()];
  }

  select(id: string | null): void {
    if (this.selected === id) return;
    const before = this.selected ? this.opts.map.get(this.selected) : undefined;
    this.selected = id;
    for (const a of [before, id ? this.opts.map.get(id) : undefined]) {
      const page = a ? this.opts.view.pageOfKey(a.pageKey) : null;
      if (page) this.renderPage(page);
    }
    this.notifySelect();
  }
  private notifySelect(): void {
    const a = this.selectedAnnot;
    for (const cb of this.selectListeners) cb(a);
  }

  // ---------- 쓰기 ----------
  private write(fn: () => void): void {
    this.opts.ydoc.transact(fn, LOCAL_ORIGIN);
  }
  add(a: PdfAnnot): void {
    this.write(() => this.opts.map.set(a.id, a));
  }
  update(id: string, patch: Partial<PdfAnnot>): void {
    const a = this.opts.map.get(id);
    if (!a) return;
    this.write(() => this.opts.map.set(id, { ...a, ...patch, modifiedAt: new Date().toISOString() }));
  }
  remove(id: string): void {
    if (!this.opts.map.has(id)) return;
    this.write(() => this.opts.map.delete(id));
  }
  private make(type: PdfAnnotType, pageKey: string, geometry: Partial<PdfAnnot> & { rect: PdfAnnot['rect'] }): PdfAnnot {
    const style = this.styleOf(type);
    const now = new Date().toISOString();
    return { id: newId(), pageKey, type, color: style.color, width: style.width, opacity: style.opacity, author: this.opts.author, createdAt: now, modifiedAt: now, ...geometry };
  }

  /** 주석이 있는 쪽으로 가서 고른다(메모 목록에서) */
  reveal(id: string): void {
    const a = this.opts.map.get(id);
    const page = a ? this.opts.view.pageOfKey(a.pageKey) : null;
    if (!a || !page) return;
    this.opts.view.goTo(page);
    // 쪽이 그려진 뒤 그 주석이 보이게
    requestAnimationFrame(() => {
      const vp = this.opts.view.viewportOf(page);
      const el = this.opts.view.pageElement(page);
      if (vp && el) {
        const box = viewBox(vp, a);
        this.opts.scroller.scrollTop = el.offsetTop + box.y - this.opts.scroller.clientHeight / 3;
      }
      this.setTool('select');
      this.select(id);
      if (a.type === 'note') this.openNote(id);
    });
  }

  /** 썸네일 같은 다른 자리에 이 쪽의 주석을 그린다(고르기·조절점 없이) */
  drawOverlay(page: number, host: HTMLElement, vp: PageViewport): void {
    const key = this.opts.view.pageKeyOf(page);
    const root = svg('svg', { width: vp.width, height: vp.height, class: 'pdf-annot-svg' });
    host.append(root);
    for (const a of [...this.opts.map.values()].filter((x) => x.pageKey === key).sort((x, y) => x.createdAt.localeCompare(y.createdAt))) {
      this.drawAnnot(a, vp, root, host as HTMLDivElement);
    }
  }

  /** 주석이 바뀐 쪽들 — 썸네일 갱신용 */
  onPagesChanged(cb: (pages: Set<number>) => void): void {
    this.pageChangeListeners.push(cb);
  }

  // ---------- 그리기 ----------
  renderAll(): void {
    for (const page of this.opts.view.renderedPages()) this.renderPage(page);
  }

  private renderPage(page: number): void {
    const vp = this.opts.view.viewportOf(page);
    const el = this.opts.view.pageElement(page);
    const layer = el?.querySelector<HTMLDivElement>('.pdf-annot-layer');
    if (!vp || !layer) return;
    layer.replaceChildren();
    const key = this.opts.view.pageKeyOf(page);
    const annots = [...this.opts.map.values()].filter((a) => a.pageKey === key).sort((a, b) => a.createdAt.localeCompare(b.createdAt));
    const root = svg('svg', { width: vp.width, height: vp.height, class: 'pdf-annot-svg' });
    layer.append(root);
    for (const a of annots) this.drawAnnot(a, vp, root, layer);
    const sel = this.selectedAnnot;
    if (sel && sel.pageKey === key) this.drawSelection(sel, vp, layer);
  }

  private drawAnnot(a: PdfAnnot, vp: PageViewport, root: SVGSVGElement, layer: HTMLDivElement): void {
    const s = vp.scale;
    const g = svg('g', { 'data-id': a.id, class: `pdf-annot pdf-annot-${a.type}`, opacity: a.opacity });
    const stroke = { stroke: a.color, 'stroke-width': Math.max(0.5, a.width * s), fill: 'none', 'stroke-linecap': 'round', 'stroke-linejoin': 'round' };
    const hit = (shape: SVGElement) => {
      // 가는 선도 잘 눌리게 투명한 굵은 테두리를 한 겹 더
      const clone = shape.cloneNode() as SVGElement;
      clone.setAttribute('stroke', 'transparent');
      clone.setAttribute('stroke-width', String(Math.max(10, a.width * s + 8)));
      clone.setAttribute('class', 'pdf-annot-hit');
      g.append(clone);
    };
    const poly = (pts: number[]) => {
      const out: string[] = [];
      for (let i = 0; i + 1 < pts.length; i += 2) out.push(pt(vp, pts[i], pts[i + 1]).map((v) => v.toFixed(2)).join(','));
      return out.join(' ');
    };
    switch (a.type) {
      case 'highlight': {
        g.setAttribute('class', `${g.getAttribute('class')} pdf-annot-multiply`);
        const q = a.quads ?? [];
        for (let i = 0; i + 7 < q.length; i += 8) {
          g.append(svg('polygon', { points: poly([q[i], q[i + 1], q[i + 2], q[i + 3], q[i + 6], q[i + 7], q[i + 4], q[i + 5]]), fill: a.color, class: 'pdf-annot-fill' }));
        }
        break;
      }
      case 'underline':
      case 'strike': {
        const q = a.quads ?? [];
        for (let i = 0; i + 7 < q.length; i += 8) {
          const top = (q[i + 1] + q[i + 3]) / 2;
          const bottom = (q[i + 5] + q[i + 7]) / 2;
          const thick = Math.max(0.8, Math.abs(top - bottom) * 0.07);
          const y = a.type === 'underline' ? bottom + thick : (top + bottom) / 2;
          const [x1, y1] = pt(vp, q[i + 4], y);
          const [x2, y2] = pt(vp, q[i + 6], y);
          const line = svg('line', { x1, y1, x2, y2, stroke: a.color, 'stroke-width': Math.max(1, thick * s) });
          g.append(line);
          hit(line);
        }
        break;
      }
      case 'ink':
        for (const p of a.paths ?? []) {
          const line = svg('polyline', { points: poly([...p]), ...stroke });
          g.append(line);
          hit(line);
        }
        break;
      case 'rect': {
        const [x1, y1, x2, y2] = a.rect;
        const inset = a.width / 2;
        const shape = svg('polygon', { points: poly([x1 + inset, y1 + inset, x2 - inset, y1 + inset, x2 - inset, y2 - inset, x1 + inset, y2 - inset]), ...stroke, fill: a.fill ?? 'none' });
        g.append(shape);
        hit(shape);
        break;
      }
      case 'ellipse': {
        const b = viewBox(vp, { ...a, type: 'freetext' });
        const shape = svg('ellipse', { cx: b.x + b.w / 2, cy: b.y + b.h / 2, rx: Math.max(1, b.w / 2 - (a.width * s) / 2), ry: Math.max(1, b.h / 2 - (a.width * s) / 2), ...stroke, fill: a.fill ?? 'none' });
        g.append(shape);
        hit(shape);
        break;
      }
      case 'line':
      case 'arrow': {
        const [lx1, ly1, lx2, ly2] = a.line ?? [a.rect[0], a.rect[1], a.rect[2], a.rect[3]];
        const [x1, y1] = pt(vp, lx1, ly1);
        const [x2, y2] = pt(vp, lx2, ly2);
        const line = svg('line', { x1, y1, x2, y2, ...stroke });
        g.append(line);
        hit(line);
        if (a.type === 'arrow') {
          const len = Math.max(6, a.width * 4) * s;
          const ang = Math.atan2(y2 - y1, x2 - x1);
          const p1 = [x2 + len * Math.cos(ang + Math.PI - Math.PI / 7), y2 + len * Math.sin(ang + Math.PI - Math.PI / 7)];
          const p2 = [x2 + len * Math.cos(ang + Math.PI + Math.PI / 7), y2 + len * Math.sin(ang + Math.PI + Math.PI / 7)];
          g.append(svg('polyline', { points: `${p1.join(',')} ${x2},${y2} ${p2.join(',')}`, ...stroke }));
        }
        break;
      }
      case 'note': {
        const b = viewBox(vp, a);
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'pdf-note-icon';
        btn.dataset.id = a.id;
        btn.title = `${a.author || '메모'}: ${a.text ?? ''}`;
        btn.style.left = `${b.x}px`;
        btn.style.top = `${b.y}px`;
        btn.style.width = `${NOTE_SIZE * s}px`;
        btn.style.height = `${NOTE_SIZE * s}px`;
        btn.style.setProperty('--note', a.color);
        btn.innerHTML =
          '<svg viewBox="0 0 24 24"><path d="M2 2h20v14H10l-5 6 1-6H2z" fill="var(--note)" stroke="#3f3214" stroke-width="1.2" stroke-linejoin="round"/><path d="M6 7h12M6 11h9" stroke="#3f3214" stroke-width="1.2" stroke-linecap="round"/></svg>';
        layer.append(btn);
        return;
      }
      case 'freetext': {
        const [x1, y1, x2, y2] = a.rect;
        const [tx, ty] = pt(vp, x1, y2);
        const div = document.createElement('div');
        div.className = 'pdf-freetext';
        div.dataset.id = a.id;
        div.style.left = `${tx}px`;
        div.style.top = `${ty}px`;
        div.style.width = `${(x2 - x1) * s}px`;
        div.style.height = `${(y2 - y1) * s}px`;
        div.style.fontSize = `${(a.fontSize ?? 12) * s}px`;
        div.style.color = a.color;
        div.style.opacity = String(a.opacity);
        div.style.transform = vp.rotation ? `rotate(${vp.rotation}deg)` : '';
        div.textContent = a.text ?? '';
        layer.append(div);
        return;
      }
    }
    root.append(g);
  }

  private drawSelection(a: PdfAnnot, vp: PageViewport, layer: HTMLDivElement): void {
    const b = viewBox(vp, a);
    const box = document.createElement('div');
    box.className = 'pdf-annot-selection';
    box.dataset.id = a.id;
    box.style.left = `${b.x - 3}px`;
    box.style.top = `${b.y - 3}px`;
    box.style.width = `${b.w + 6}px`;
    box.style.height = `${b.h + 6}px`;
    if (this.opts.canEdit && BOX_TYPES.has(a.type)) {
      const handle = document.createElement('span');
      handle.className = 'pdf-annot-handle';
      handle.dataset.resize = a.id;
      box.append(handle);
    }
    layer.append(box);
  }

  // ---------- 누르기·끌기 ----------
  private pageAt(target: EventTarget | null): { page: number; el: HTMLElement; vp: PageViewport } | null {
    const el = (target as HTMLElement | null)?.closest<HTMLElement>('.pdf-page');
    const page = el ? Number(el.dataset.page) : NaN;
    const vp = Number.isFinite(page) ? this.opts.view.viewportOf(page) : null;
    return el && vp ? { page, el, vp } : null;
  }
  private local(event: PointerEvent | MouseEvent, el: HTMLElement): [number, number] {
    const r = el.getBoundingClientRect();
    return [event.clientX - r.left, event.clientY - r.top];
  }

  private bindPointer(): void {
    const scroller = this.opts.scroller;
    scroller.addEventListener('pointerdown', (event) => {
      if (event.button !== 0) return;
      const target = event.target as HTMLElement;
      if (target.closest('.pdf-note-popup, .pdf-freetext-editor')) return;
      const hitPage = this.pageAt(target);
      const annotEl = target.closest<HTMLElement>('[data-id]');
      const annotId = annotEl?.dataset.id ?? null;
      const resizeId = target.dataset.resize ?? null;

      if (this.tool === 'eraser') {
        if (annotId && this.opts.canEdit) {
          event.preventDefault();
          this.remove(annotId);
        }
        return;
      }
      if (this.tool === 'select') {
        if (resizeId && hitPage && this.opts.canEdit) {
          event.preventDefault();
          this.dragResize(event, resizeId, hitPage);
          return;
        }
        if (annotId) {
          event.preventDefault();
          this.select(annotId);
          const a = this.opts.map.get(annotId);
          if (a && hitPage && this.opts.canEdit) this.dragMove(event, a, hitPage, annotEl!);
          else if (a?.type === 'note') this.openNote(annotId);
          return;
        }
        this.select(null);
        this.closePopup();
        return;
      }
      if (!this.opts.canEdit || !hitPage || !DRAW_TOOLS.has(this.tool)) return;
      event.preventDefault();
      this.drawNew(event, hitPage);
    });
    scroller.addEventListener('dblclick', (event) => {
      const id = (event.target as HTMLElement).closest<HTMLElement>('[data-id]')?.dataset.id;
      const a = id ? this.opts.map.get(id) : undefined;
      if (a?.type === 'freetext' && this.opts.canEdit) this.editFreetext(a.id);
      if (a?.type === 'note') this.openNote(a.id);
    });
  }

  /** 끌어 옮기기 — 화면에서만 옮기다 손을 떼면 한 번 쓴다 */
  private dragMove(event: PointerEvent, a: PdfAnnot, at: { el: HTMLElement; vp: PageViewport }, annotEl: HTMLElement): void {
    const start = this.local(event, at.el);
    const layer = at.el.querySelector<HTMLElement>('.pdf-annot-layer')!;
    const moving = [...layer.querySelectorAll<HTMLElement | SVGElement>(`[data-id="${CSS.escape(a.id)}"]`)];
    let dx = 0;
    let dy = 0;
    let movedEnough = false;
    const onMove = (e: PointerEvent) => {
      const [x, y] = this.local(e, at.el);
      dx = x - start[0];
      dy = y - start[1];
      if (Math.hypot(dx, dy) > 3) movedEnough = true;
      if (!movedEnough) return;
      for (const el of moving) {
        if (el instanceof SVGElement) el.setAttribute('transform', `translate(${dx} ${dy})`);
        else el.style.translate = `${dx}px ${dy}px`;
      }
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      if (!movedEnough) {
        if (a.type === 'note') this.openNote(a.id);
        return;
      }
      const [px0, py0] = toPdf(at.vp, start[0], start[1]);
      const [px1, py1] = toPdf(at.vp, start[0] + dx, start[1] + dy);
      const next = moved(a, px1 - px0, py1 - py0);
      this.update(a.id, { rect: next.rect, quads: next.quads, paths: next.paths, line: next.line });
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  }

  /** 상자 주석 크기 바꾸기 — 오른쪽 아래 조절점, 반대쪽 모서리를 고정한다 */
  private dragResize(event: PointerEvent, id: string, at: { el: HTMLElement; vp: PageViewport }): void {
    const a = this.opts.map.get(id);
    if (!a) return;
    const b = viewBox(at.vp, a);
    const anchor = toPdf(at.vp, b.x, b.y);
    const box = at.el.querySelector<HTMLElement>(`.pdf-annot-selection[data-id="${CSS.escape(id)}"]`);
    let last: [number, number] = this.local(event, at.el);
    const onMove = (e: PointerEvent) => {
      last = this.local(e, at.el);
      if (box) {
        box.style.width = `${Math.max(8, last[0] - b.x + 3)}px`;
        box.style.height = `${Math.max(8, last[1] - b.y + 3)}px`;
      }
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      const p = toPdf(at.vp, Math.max(b.x + 8, last[0]), Math.max(b.y + 8, last[1]));
      this.update(id, { rect: [Math.min(anchor[0], p[0]), Math.min(anchor[1], p[1]), Math.max(anchor[0], p[0]), Math.max(anchor[1], p[1])] });
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  }

  /** 그리는 도구로 새 주석 — 끄는 동안 미리 보기 */
  private drawNew(event: PointerEvent, at: { page: number; el: HTMLElement; vp: PageViewport }): void {
    const tool = this.tool as PdfAnnotType;
    const pageKey = this.opts.view.pageKeyOf(at.page);
    const start = this.local(event, at.el);
    const style = this.styleOf(tool);
    const layer = at.el.querySelector<HTMLElement>('.pdf-annot-layer')!;
    const preview = svg('svg', { width: at.vp.width, height: at.vp.height, class: 'pdf-annot-preview' });
    layer.append(preview);
    const strokeAttrs = { stroke: style.color, 'stroke-width': Math.max(1, style.width * at.vp.scale), fill: 'none', 'stroke-linecap': 'round', 'stroke-linejoin': 'round', opacity: style.opacity };
    const points: number[] = [...start];
    let shape: SVGElement | null = null;
    if (tool === 'ink') {
      shape = svg('polyline', { ...strokeAttrs, points: start.join(',') });
      preview.append(shape);
    } else if (tool !== 'note') {
      shape = svg(tool === 'ellipse' ? 'ellipse' : tool === 'line' || tool === 'arrow' ? 'line' : 'rect', tool === 'freetext' ? { stroke: '#1a73e8', 'stroke-dasharray': '4 3', fill: 'rgba(26,115,232,0.06)' } : strokeAttrs);
      preview.append(shape);
    }
    let last = start;
    const onMove = (e: PointerEvent) => {
      last = this.local(e, at.el);
      if (!shape) return;
      if (tool === 'ink') {
        points.push(last[0], last[1]);
        shape.setAttribute('points', points.join(' ').replace(/(\S+) (\S+)/g, '$1,$2'));
        return;
      }
      const x = Math.min(start[0], last[0]);
      const y = Math.min(start[1], last[1]);
      const w = Math.abs(last[0] - start[0]);
      const h = Math.abs(last[1] - start[1]);
      if (tool === 'ellipse') Object.entries({ cx: x + w / 2, cy: y + h / 2, rx: w / 2, ry: h / 2 }).forEach(([k, v]) => shape!.setAttribute(k, String(v)));
      else if (tool === 'line' || tool === 'arrow') Object.entries({ x1: start[0], y1: start[1], x2: last[0], y2: last[1] }).forEach(([k, v]) => shape!.setAttribute(k, String(v)));
      else Object.entries({ x, y, width: w, height: h }).forEach(([k, v]) => shape!.setAttribute(k, String(v)));
    };
    const onUp = () => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onUp);
      preview.remove();
      const vp = at.vp;
      const tiny = Math.hypot(last[0] - start[0], last[1] - start[1]) < 4;
      if (tool === 'note') {
        const [x, y] = toPdf(vp, start[0], start[1]);
        const annot = this.make('note', pageKey, { rect: [x, y - NOTE_SIZE, x + NOTE_SIZE, y], text: '' });
        this.add(annot);
        this.setTool('select');
        this.select(annot.id);
        this.openNote(annot.id, true);
        return;
      }
      if (tool === 'ink') {
        if (points.length < 4) points.push(points[0] + 0.5, points[1]);
        const pdfPts: number[] = [];
        const simple = simplify(points, 1.5);
        for (let i = 0; i + 1 < simple.length; i += 2) pdfPts.push(...toPdf(vp, simple[i], simple[i + 1]));
        this.add(this.make('ink', pageKey, { rect: boundsOfPoints(pdfPts, style.width), paths: [pdfPts] }));
        return;
      }
      if (tool === 'line' || tool === 'arrow') {
        if (tiny) return;
        const [x1, y1] = toPdf(vp, start[0], start[1]);
        const [x2, y2] = toPdf(vp, last[0], last[1]);
        this.add(this.make(tool, pageKey, { rect: boundsOfPoints([x1, y1, x2, y2]), line: [x1, y1, x2, y2] }));
        return;
      }
      // 상자 종류: 너무 작게 끌었으면(=클릭) 글상자는 기본 크기, 도형은 만들지 않는다
      let end = last;
      if (tiny) {
        if (tool !== 'freetext') return;
        end = [start[0] + 200 * vp.scale, start[1] + 40 * vp.scale];
      }
      const p1 = toPdf(vp, start[0], start[1]);
      const p2 = toPdf(vp, end[0], end[1]);
      const rect: PdfAnnot['rect'] = [Math.min(p1[0], p2[0]), Math.min(p1[1], p2[1]), Math.max(p1[0], p2[0]), Math.max(p1[1], p2[1])];
      if (tool === 'freetext') {
        const annot = this.make('freetext', pageKey, { rect, text: '', fontSize: 14 });
        this.add(annot);
        this.setTool('select');
        this.select(annot.id);
        this.editFreetext(annot.id, true);
        return;
      }
      this.add(this.make(tool, pageKey, { rect, fill: null }));
    };
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onUp);
  }

  /** 형광펜·밑줄·취소선 — 글자를 끌어 고른 줄들로 */
  private markSelection(): void {
    if (!this.opts.canEdit || !TEXT_TOOLS.has(this.tool)) return;
    const selection = window.getSelection();
    if (!selection || selection.isCollapsed || selection.rangeCount === 0) return;
    const rects = [...selection.getRangeAt(0).getClientRects()].filter((r) => r.width > 1 && r.height > 1);
    if (rects.length === 0) return;
    // 쪽마다, 같은 줄(세로 가운데가 비슷)끼리 합친다
    const byPage = new Map<number, DOMRect[]>();
    for (const r of rects) {
      const el = document.elementsFromPoint(r.left + r.width / 2, r.top + r.height / 2).find((e) => (e as HTMLElement).classList?.contains('pdf-page')) as HTMLElement | undefined;
      const page = el ? Number(el.dataset.page) : NaN;
      if (!Number.isFinite(page)) continue;
      const list = byPage.get(page) ?? [];
      list.push(r);
      byPage.set(page, list);
    }
    const type = this.tool as PdfAnnotType;
    for (const [page, list] of byPage) {
      const el = this.opts.view.pageElement(page);
      const vp = this.opts.view.viewportOf(page);
      if (!el || !vp) continue;
      const origin = el.getBoundingClientRect();
      const lines: Array<{ l: number; t: number; r: number; b: number }> = [];
      for (const r of list.sort((a, b) => a.top - b.top || a.left - b.left)) {
        const box = { l: r.left - origin.left, t: r.top - origin.top, r: r.right - origin.left, b: r.bottom - origin.top };
        const same = lines.find((x) => Math.abs((x.t + x.b) / 2 - (box.t + box.b) / 2) < Math.min(x.b - x.t, box.b - box.t) * 0.6);
        if (same) {
          same.l = Math.min(same.l, box.l);
          same.r = Math.max(same.r, box.r);
          same.t = Math.min(same.t, box.t);
          same.b = Math.max(same.b, box.b);
        } else lines.push(box);
      }
      const quads: number[] = [];
      for (const ln of lines) {
        const corners = [toPdf(vp, ln.l, ln.t), toPdf(vp, ln.r, ln.t), toPdf(vp, ln.l, ln.b), toPdf(vp, ln.r, ln.b)];
        const xs = corners.map((c) => c[0]);
        const ys = corners.map((c) => c[1]);
        const [x1, x2, y1, y2] = [Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)];
        quads.push(x1, y2, x2, y2, x1, y1, x2, y1);
      }
      if (quads.length) this.add(this.make(type, this.opts.view.pageKeyOf(page), { rect: boundsOfPoints(quads), quads }));
    }
    selection.removeAllRanges();
  }

  // ---------- 메모 창·글상자 입력 ----------
  closePopup(): void {
    this.popup?.remove();
    this.popup = null;
  }

  openNote(id: string, fresh = false): void {
    const a = this.opts.map.get(id);
    const page = a ? this.opts.view.pageOfKey(a.pageKey) : null;
    const el = page ? this.opts.view.pageElement(page) : null;
    const vp = page ? this.opts.view.viewportOf(page) : null;
    if (!a || !el || !vp) return;
    this.closePopup();
    const b = viewBox(vp, a);
    const pop = document.createElement('div');
    pop.className = 'pdf-note-popup';
    pop.style.left = `${Math.min(b.x + b.w + 6, el.clientWidth - 260)}px`;
    pop.style.top = `${b.y}px`;
    const head = document.createElement('div');
    head.className = 'pdf-note-head';
    head.innerHTML = '<strong></strong><span></span>';
    head.querySelector('strong')!.textContent = a.author || '메모';
    head.querySelector('span')!.textContent = formatTime(a.modifiedAt || a.createdAt);
    const area = document.createElement('textarea');
    area.value = a.text ?? '';
    area.placeholder = '메모를 입력하세요';
    area.readOnly = !this.opts.canEdit;
    const actions = document.createElement('div');
    actions.className = 'pdf-note-actions';
    const save = () => {
      const latest = this.opts.map.get(id);
      if (latest && this.opts.canEdit && area.value !== (latest.text ?? '')) this.update(id, { text: area.value });
    };
    if (this.opts.canEdit) {
      const del = document.createElement('button');
      del.type = 'button';
      del.textContent = '삭제';
      del.addEventListener('click', () => {
        this.remove(id);
        this.closePopup();
      });
      const ok = document.createElement('button');
      ok.type = 'button';
      ok.className = 'primary';
      ok.textContent = '저장';
      ok.addEventListener('click', () => {
        save();
        this.closePopup();
      });
      actions.append(del, ok);
    } else {
      const close = document.createElement('button');
      close.type = 'button';
      close.textContent = '닫기';
      close.addEventListener('click', () => this.closePopup());
      actions.append(close);
    }
    area.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) {
        save();
        this.closePopup();
      } else if (e.key === 'Escape') {
        // 막 만든 빈 메모를 Esc로 닫으면 지운다
        if (fresh && !area.value.trim()) this.remove(id);
        this.closePopup();
      }
    });
    pop.append(head, area, actions);
    el.append(pop);
    this.popup = pop;
    area.focus();
  }

  editFreetext(id: string, fresh = false): void {
    const a = this.opts.map.get(id);
    const page = a ? this.opts.view.pageOfKey(a.pageKey) : null;
    const el = page ? this.opts.view.pageElement(page) : null;
    const vp = page ? this.opts.view.viewportOf(page) : null;
    if (!a || !el || !vp) return;
    this.closePopup();
    const [x1, y1, x2, y2] = a.rect;
    const [tx, ty] = pt(vp, x1, y2);
    const area = document.createElement('textarea');
    area.className = 'pdf-freetext-editor';
    area.value = a.text ?? '';
    area.style.left = `${tx}px`;
    area.style.top = `${ty}px`;
    area.style.width = `${(x2 - x1) * vp.scale}px`;
    area.style.height = `${(y2 - y1) * vp.scale}px`;
    area.style.fontSize = `${(a.fontSize ?? 12) * vp.scale}px`;
    area.style.color = a.color;
    if (vp.rotation) area.style.transform = `rotate(${vp.rotation}deg)`;
    const commit = () => {
      if (!area.isConnected) return;
      const text = area.value;
      area.remove();
      this.popup = null;
      if (!text.trim()) {
        if (fresh || !(a.text ?? '').trim()) this.remove(id);
        return;
      }
      if (text !== a.text) this.update(id, { text });
    };
    area.addEventListener('blur', commit);
    area.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' || (e.key === 'Enter' && (e.ctrlKey || e.metaKey))) {
        e.preventDefault();
        area.blur();
      }
    });
    el.append(area);
    this.popup = area;
    area.focus();
  }
}

export { ANNOT_LABELS, formatTime };
