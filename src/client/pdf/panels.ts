/**
 * PDF 편집기 오른쪽 정보 창 — 페이지(썸네일)·책갈피·레이어. 맨 오른쪽 줄의 단추로 바꾸고, 같은 단추를 다시 누르면 창을 닫는다
 * (pdfcraft의 Inspector와 같은 동작). 썸네일은 보이는 것만 작은 배율로 그린다.
 */
import { readLayers, readOutline, type OutlineNode, type PDFDocumentProxy } from './pdfDocument.js';
import type { PageView } from './pageView.js';

export type PanelName = 'thumbs' | 'outline' | 'layers';

const THUMB_WIDTH = 140;

export interface Panels {
  show(panel: PanelName | null): void;
  readonly current: PanelName | null;
}

export function setupPanels(opts: { doc: PDFDocumentProxy; view: PageView }): Panels {
  const { doc, view } = opts;
  const inspector = document.getElementById('inspector') as HTMLElement;
  const railButtons = [...document.querySelectorAll<HTMLButtonElement>('.pdf-rail-btn[data-panel]')];
  let current: PanelName | null = 'thumbs';

  function show(panel: PanelName | null): void {
    current = panel;
    inspector.hidden = panel === null;
    inspector.classList.toggle('forced-open', panel !== null && window.matchMedia('(max-width: 900px)').matches);
    for (const section of inspector.querySelectorAll<HTMLElement>('.pdf-panel')) section.hidden = section.dataset.panel !== panel;
    for (const btn of railButtons) btn.setAttribute('aria-pressed', String(btn.dataset.panel === panel));
    if (panel === 'thumbs') thumbs.reveal(view.currentPage);
  }
  for (const btn of railButtons) {
    btn.addEventListener('click', () => {
      const panel = btn.dataset.panel as PanelName;
      show(current === panel && !inspector.hidden ? null : panel);
    });
  }
  for (const close of inspector.querySelectorAll<HTMLButtonElement>('[data-close-panel]')) close.addEventListener('click', () => show(null));
  // 좁은 화면은 처음부터 닫아 둔다(CSS도 접지만 단추 눌림 상태를 맞춘다)
  if (window.matchMedia('(max-width: 900px)').matches) current = null;

  const thumbs = setupThumbs(doc, view);
  void setupOutline(doc, view);
  setupLayers(view);
  view.onPageChange((page) => thumbs.markCurrent(page));
  show(current);
  return {
    show,
    get current() {
      return current;
    },
  };
}

function setupThumbs(doc: PDFDocumentProxy, view: PageView): { markCurrent(page: number): void; reveal(page: number): void } {
  const host = document.getElementById('thumbs') as HTMLElement;
  const buttons: HTMLButtonElement[] = [];
  const drawn = new Set<number>();
  const observer = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const page = Number((entry.target as HTMLElement).dataset.page);
        if (drawn.has(page)) continue;
        drawn.add(page);
        void drawThumb(doc, page, entry.target.querySelector('.pdf-thumb-img') as HTMLElement);
      }
    },
    { root: host, rootMargin: '200px 0px' },
  );
  for (let page = 1; page <= doc.numPages; page += 1) {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'pdf-thumb';
    btn.dataset.page = String(page);
    btn.title = `${page}쪽으로 이동`;
    const img = document.createElement('div');
    img.className = 'pdf-thumb-img';
    img.style.width = `${THUMB_WIDTH}px`;
    img.style.height = `${Math.round(THUMB_WIDTH * 1.414)}px`;
    const label = document.createElement('span');
    label.textContent = String(page);
    btn.append(img, label);
    btn.addEventListener('click', () => view.goTo(page));
    host.append(btn);
    buttons.push(btn);
    observer.observe(btn);
  }
  function markCurrent(page: number): void {
    for (const b of buttons) b.classList.toggle('current', Number(b.dataset.page) === page);
  }
  function reveal(page: number): void {
    const btn = buttons[page - 1];
    if (!btn || host.offsetParent === null) return;
    const top = btn.offsetTop - host.offsetTop;
    if (top < host.scrollTop || top + btn.offsetHeight > host.scrollTop + host.clientHeight) host.scrollTop = top - 40;
  }
  markCurrent(1);
  view.onPageChange((page) => reveal(page));
  return { markCurrent, reveal };
}

async function drawThumb(doc: PDFDocumentProxy, pageNumber: number, box: HTMLElement): Promise<void> {
  try {
    const page = await doc.getPage(pageNumber);
    const base = page.getViewport({ scale: 1 });
    const scale = THUMB_WIDTH / base.width;
    const viewport = page.getViewport({ scale });
    const ratio = Math.min(2, window.devicePixelRatio || 1);
    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(viewport.width * ratio);
    canvas.height = Math.floor(viewport.height * ratio);
    box.style.height = `${Math.round(viewport.height)}px`;
    const ctx = canvas.getContext('2d', { alpha: false });
    if (!ctx) return;
    await page.render({ canvasContext: ctx, viewport, transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined }).promise;
    box.replaceChildren(canvas);
  } catch (error) {
    console.error(`썸네일 ${pageNumber} 실패:`, error);
  }
}

async function setupOutline(doc: PDFDocumentProxy, view: PageView): Promise<void> {
  const host = document.getElementById('outline') as HTMLElement;
  let nodes: OutlineNode[] = [];
  try {
    nodes = await readOutline(doc);
  } catch (error) {
    console.error('책갈피 읽기 실패:', error);
  }
  if (nodes.length === 0) {
    host.innerHTML = '<p class="pdf-empty">이 문서에는 책갈피가 없습니다.</p>';
    return;
  }
  const build = (items: readonly OutlineNode[], depth: number): HTMLUListElement => {
    const ul = document.createElement('ul');
    for (const node of items) {
      const li = document.createElement('li');
      const row = document.createElement('div');
      row.className = 'pdf-outline-row';
      const toggle = document.createElement('button');
      toggle.type = 'button';
      toggle.className = 'pdf-outline-toggle';
      const link = document.createElement('button');
      link.type = 'button';
      link.className = 'pdf-outline-link';
      const title = document.createElement('span');
      title.textContent = node.title;
      link.append(title);
      if (node.page !== null) {
        const pageLabel = document.createElement('em');
        pageLabel.textContent = String(node.page);
        link.append(pageLabel);
        link.addEventListener('click', () => view.goTo(node.page!));
      }
      row.append(toggle, link);
      li.append(row);
      if (node.children.length > 0) {
        const child = build(node.children, depth + 1);
        // 첫 단계만 펼쳐 둔다
        child.hidden = depth >= 1;
        toggle.textContent = child.hidden ? '▸' : '▾';
        toggle.setAttribute('aria-label', '펼치기/접기');
        toggle.addEventListener('click', () => {
          child.hidden = !child.hidden;
          toggle.textContent = child.hidden ? '▸' : '▾';
        });
        li.append(child);
      } else {
        toggle.hidden = true;
      }
      ul.append(li);
    }
    return ul;
  };
  host.replaceChildren(build(nodes, 0));
}

function setupLayers(view: PageView): void {
  const host = document.getElementById('layers') as HTMLElement;
  const config = view.layers;
  const layers = config ? readLayers(config) : [];
  if (layers.length === 0) {
    host.innerHTML = '<p class="pdf-empty">이 문서에는 레이어가 없습니다.</p>';
    return;
  }
  for (const layer of layers) {
    const row = document.createElement('label');
    row.className = 'pdf-layer-row';
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.checked = layer.visible;
    box.dataset.layer = layer.id;
    box.addEventListener('change', () => view.setLayerVisibility(layer.id, box.checked));
    const name = document.createElement('span');
    name.textContent = layer.name;
    row.append(box, name);
    host.append(row);
  }
}
