/**
 * PDF 편집기 오른쪽 정보 창 — 페이지(썸네일)·책갈피·레이어. 맨 오른쪽 줄의 단추로 바꾸고, 같은 단추를 다시 누르면 창을 닫는다
 * (pdfcraft의 Inspector와 같은 동작). 썸네일은 보이는 것만 작은 배율로 그린다.
 */
import { readLayers, readOutline, type OutlineNode, type PDFDocumentProxy } from './pdfDocument.js';
import type { PageView } from './pageView.js';

export type PanelName = 'thumbs' | 'outline' | 'layers' | 'comments';

const THUMB_WIDTH = 140;

export interface Panels {
  show(panel: PanelName | null): void;
  readonly current: PanelName | null;
  /** 썸네일 위에 주석을 그리는 함수를 단다(annotLayer.drawOverlay) — 그려 둔 썸네일도 바로 다시 그린다 */
  setThumbOverlay(draw: ThumbOverlay): void;
  /** 이 쪽들의 썸네일 주석을 다시 그린다(주석이 바뀌었을 때) */
  refreshThumbOverlays(pages: Iterable<number>): void;
}

export type ThumbOverlay = (page: number, host: HTMLElement, viewport: ReturnType<Awaited<ReturnType<PDFDocumentProxy['getPage']>>['getViewport']>) => void;

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

  const thumbs = setupThumbs(view);
  void setupOutline(doc, view);
  setupLayers(view);
  view.onPageChange((page) => thumbs.markCurrent(page));
  show(current);
  return {
    show,
    get current() {
      return current;
    },
    setThumbOverlay: (draw) => thumbs.setOverlay(draw),
    refreshThumbOverlays: (pages) => thumbs.refreshOverlays(pages),
  };
}

function setupThumbs(
  view: PageView,
): { markCurrent(page: number): void; reveal(page: number): void; setOverlay(draw: ThumbOverlay): void; refreshOverlays(pages: Iterable<number>): void } {
  const host = document.getElementById('thumbs') as HTMLElement;
  let buttons: HTMLButtonElement[] = [];
  let drawn = new Set<number>();
  // 썸네일마다 그린 viewport — 주석 겹침 층을 다시 그릴 때 쓴다
  let viewports = new Map<number, Parameters<ThumbOverlay>[2]>();
  let overlay: ThumbOverlay | null = null;
  let generation = 0;
  const paintOverlay = (page: number) => {
    const vp = viewports.get(page);
    const box = buttons[page - 1]?.querySelector<HTMLElement>('.pdf-thumb-img');
    if (!overlay || !vp || !box) return;
    let layer = box.querySelector<HTMLElement>('.pdf-thumb-annots');
    if (!layer) {
      layer = document.createElement('div');
      layer.className = 'pdf-thumb-annots';
      box.append(layer);
    }
    layer.replaceChildren();
    overlay(page, layer, vp);
  };
  const observer = new IntersectionObserver(
    (entries) => {
      for (const entry of entries) {
        if (!entry.isIntersecting) continue;
        const page = Number((entry.target as HTMLElement).dataset.page);
        if (drawn.has(page)) continue;
        drawn.add(page);
        const gen = generation;
        void view.drawThumbnail(page, entry.target.querySelector('.pdf-thumb-img') as HTMLElement, THUMB_WIDTH).then((vp) => {
          if (!vp || gen !== generation) return;
          viewports.set(page, vp);
          paintOverlay(page);
        });
      }
    },
    { root: host, rootMargin: '200px 0px' },
  );
  // 쪽 목록이 바뀌면(정리 화면) 통째로 다시 짠다 — 썸네일은 보이는 것만 다시 그린다
  function build(): void {
    generation += 1;
    observer.disconnect();
    host.replaceChildren();
    buttons = [];
    drawn = new Set();
    viewports = new Map();
    for (let page = 1; page <= view.pageCount; page += 1) {
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
    markCurrent(view.currentPage);
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
  build();
  view.onPagesChange(build);
  view.onPageChange((page) => reveal(page));
  return {
    markCurrent,
    reveal,
    setOverlay(draw) {
      overlay = draw;
      for (const page of viewports.keys()) paintOverlay(page);
    },
    refreshOverlays(pages) {
      for (const page of pages) paintOverlay(page);
    },
  };
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
        // 책갈피는 원본 쪽을 가리킨다 — 쪽 정리 뒤에는 지금 몇 쪽인지로(지운 쪽이면 흐리게)
        const origKey = `o:${node.page - 1}`;
        const pageLabel = document.createElement('em');
        const sync = () => {
          const now = view.pageOfKey(origKey);
          pageLabel.textContent = now === null ? '—' : String(now);
          link.disabled = now === null;
        };
        sync();
        view.onPagesChange(sync);
        link.append(pageLabel);
        link.addEventListener('click', () => {
          const now = view.pageOfKey(origKey);
          if (now !== null) view.goTo(now);
        });
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
