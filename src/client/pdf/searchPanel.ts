/**
 * PDF 편집기 오른쪽 '찾기' 패널 — 3단 검색(① 이 페이지 ② 전체 문서 ③ 결과 안에서), 찾을 곳(모든 글/제목만),
 * 쪽별로 묶은 결과 목록(속한 제목 + 앞뒤 문장), 고른 결과의 문단 테두리, 문단·쪽 전체 글 미리보기.
 * 줄·문단·제목 나누기는 textStructure.ts, 찾기는 searchEngine.ts(둘 다 순수 함수).
 */
import { normalizeForFind } from './findText.js';
import { readOutline, type OutlineNode, type PDFDocumentProxy } from './pdfDocument.js';
import type { PageView, TextItemInfo } from './pageView.js';
import { blocksOf, groupByPage, searchPage, type SearchResult, type SearchScope, type SearchTarget } from './searchEngine.js';
import { analyzePage, type PageStructure, type StructItem } from './textStructure.js';

interface Stage {
  readonly query: string;
  readonly target: SearchTarget;
  readonly results: readonly SearchResult[];
}

export interface SearchPanel {
  /** 패널을 열고(검색어가 있으면 바로 찾고) 입력칸에 초점 */
  open(initial?: string): void;
  close(): void;
}

const MAX_LIST = 400;
const $ = <T extends HTMLElement = HTMLElement>(id: string) => document.getElementById(id) as T;
const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]!);

export function setupSearchPanel(opts: {
  view: PageView;
  doc: PDFDocumentProxy;
  /** 패널 보이기/닫기(정보 창의 다른 패널과 자리를 나눈다) */
  showPanel: (open: boolean) => void;
  /** 글자가 없는 스캔 문서일 때 '문자 인식' 단추가 부르는 것 */
  onRunOcr: () => void;
}): SearchPanel {
  const { view, doc } = opts;
  const section = $('panel-search');
  const input = $<HTMLInputElement>('find-input');
  const count = $('find-count');
  const list = $('sx-results');
  const empty = $('sx-empty');
  const trail = $('sx-trail');
  const preview = $('sx-preview');
  const previewHeading = $('sx-preview-heading');
  const previewBody = $('sx-preview-body');
  const pager = $('sx-pager');
  const pageNum = $('sx-page-num');
  const scopeButtons = [...section.querySelectorAll<HTMLButtonElement>('[data-scope]')];
  const headingsOnly = $<HTMLInputElement>('sx-headings');
  const totalBadge = $('sx-total');
  const tabButtons = [...section.querySelectorAll<HTMLButtonElement>('[data-view]')];

  let scope: SearchScope = 'doc';
  let baseScope: Exclude<SearchScope, 'within'> = 'doc';
  let target: SearchTarget = 'all';
  let stages: Stage[] = [];
  let results: SearchResult[] = [];
  let index = -1;
  let generation = 0;
  let searching = false;
  let timer = 0;
  let previewMode: 'block' | 'page' = 'block';
  let previewPage = 1;
  let listLimit = MAX_LIST;

  // 쪽 구조는 글 조각 배열이 같을 때만 다시 쓴다(문자 인식으로 글이 더해지면 새 배열이 온다)
  const structures = new Map<number, { items: readonly TextItemInfo[]; structure: PageStructure }>();
  let outlineTitles: Set<string> = new Set();
  void readOutline(doc)
    .then((nodes) => {
      const walk = (ns: readonly OutlineNode[]): string[] => ns.flatMap((n) => [normalizeForFind(n.title).trim(), ...walk(n.children)]);
      outlineTitles = new Set(walk(nodes));
      structures.clear();
    })
    .catch(() => undefined);

  async function structureOf(page: number): Promise<{ items: readonly TextItemInfo[]; structure: PageStructure }> {
    const items = await view.textItemInfos(page);
    const hit = structures.get(page);
    if (hit && hit.items === items) return hit;
    // 쪽 앞에서 이어지는 제목 — 앞 쪽 구조가 있으면 그 마지막 제목
    const prev = page > 1 ? structures.get(page - 1)?.structure.lastHeading ?? '' : '';
    // 글꼴 굵기·태그 구조 역할을 조각에 붙여 제목 점수에 쓴다(못 받으면 크기·번호 같은 단서로만)
    const meta = await view.textMeta(page);
    const styled: StructItem[] = items.map((it, i) => ({
      ...it,
      bold: it.fontName ? meta.fonts[it.fontName]?.bold ?? false : false,
      role: meta.roles[i] ?? null,
    }));
    const entry = { items, structure: analyzePage(page, styled, prev, outlineTitles, { pageHeight: view.sizeOf(page).h }) };
    structures.set(page, entry);
    return entry;
  }

  // ---------- 상태 그리기 ----------
  function setChecked(buttons: HTMLButtonElement[], attr: string, value: string): void {
    for (const b of buttons) b.setAttribute('aria-checked', String(b.dataset[attr] === value));
  }

  function renderControls(): void {
    setChecked(scopeButtons, 'scope', scope);
    headingsOnly.checked = target === 'headings';
    const within = scopeButtons.find((b) => b.dataset.scope === 'within')!;
    within.disabled = scope !== 'within' && results.length === 0;
    input.placeholder = scope === 'within' ? '결과 안에서 찾을 내용' : scope === 'page' ? '이 페이지에서 찾을 내용' : '찾을 내용';
    trail.hidden = stages.length === 0;
    trail.innerHTML = stages.length
      ? stages.map((s, i) => `<button type="button" data-stage="${i}" title="이 단계로 돌아가기">${esc(s.query)} <em>${s.results.length}</em></button><span aria-hidden="true">›</span>`).join('') +
        `<span class="now">${esc(input.value.trim() || '…')}</span>`
      : '';
  }

  function renderCount(): void {
    const total = results.length;
    const pages = new Set(results.map((r) => r.page)).size;
    const position = index >= 0 ? `${index + 1}/${total}` : total > 0 ? `0/${total}` : '';
    if (!input.value.trim()) count.textContent = '';
    else if (total === 0) count.textContent = searching ? '찾는 중…' : '결과 없음';
    else count.textContent = `${position}${searching ? '…' : ''}`;
    count.dataset.pages = total > 0 ? `${pages}쪽` : '';
    totalBadge.textContent = total > 0 ? String(total) : '';
  }

  function markHtml(text: string, query: string): string {
    const q = normalizeForFind(query).trim();
    if (!q) return esc(text);
    const norm = normalizeForFind(text);
    // 정규화로 글자 수가 바뀌면 위치를 맞출 수 없다 — 그때는 표시 없이 그대로
    if (norm.length !== text.length) return esc(text);
    let out = '';
    let from = 0;
    for (let at = norm.indexOf(q); at !== -1; at = norm.indexOf(q, at + q.length)) {
      out += esc(text.slice(from, at)) + `<mark>${esc(text.slice(at, at + q.length))}</mark>`;
      from = at + q.length;
    }
    return out + esc(text.slice(from));
  }

  function renderList(): void {
    if (results.length === 0) {
      list.innerHTML = '';
      return;
    }
    const shown = results.slice(0, listLimit);
    const groups = groupByPage(shown);
    let n = 0;
    list.innerHTML =
      groups
        .map((g) => {
          const total = results.filter((r) => r.page === g.page).length;
          return (
            `<div class="pdf-sx-group" role="group" aria-label="${g.page}쪽"><div class="pdf-sx-page">${g.page}쪽 <em>${total}건</em></div>` +
            g.results
              .map((r) => {
                const i = n++;
                const head = r.isHeading ? `<span class="tag">${r.level === 1 ? '제목' : '소제목'}</span>` : r.heading ? esc(r.heading) : '';
                return `<button type="button" role="option" class="pdf-sx-item" data-i="${i}" id="sx-opt-${i}" aria-selected="${i === index}">${head ? `<span class="pdf-sx-head">${head}</span>` : ''}<span class="pdf-sx-snip">${esc(r.before)}<mark>${esc(r.hit)}</mark>${esc(r.after)}</span></button>`;
              })
              .join('') +
            '</div>'
          );
        })
        .join('') + (results.length > listLimit ? `<button type="button" class="pdf-sx-more" data-more>나머지 ${results.length - listLimit}건 더 보기</button>` : '');
  }

  function markSelected(): void {
    for (const el of list.querySelectorAll<HTMLElement>('.pdf-sx-item')) {
      const on = Number(el.dataset.i) === index;
      el.setAttribute('aria-selected', String(on));
      if (on) el.scrollIntoView({ block: 'nearest' });
    }
    if (index >= 0) input.setAttribute('aria-activedescendant', `sx-opt-${index}`);
    else input.removeAttribute('aria-activedescendant');
  }

  function renderEmpty(message: string, withOcr = false): void {
    empty.hidden = !message;
    empty.innerHTML = message ? `<p>${message}</p>${withOcr ? '<button type="button" class="pdf-sx-ocr" data-ocr>문자 인식(OCR) 실행</button>' : ''}` : '';
  }

  async function renderPreview(): Promise<void> {
    const current = results[index];
    if (!current && previewMode === 'block') {
      preview.hidden = true;
      return;
    }
    preview.hidden = false;
    for (const t of tabButtons) t.setAttribute('aria-selected', String(t.dataset.view === previewMode));
    const q = input.value.trim();
    if (previewMode === 'block' && current) {
      const { structure } = await structureOf(current.page);
      const block = structure.blocks[current.block];
      previewHeading.textContent = block.isHeading ? `${current.page}쪽 · ${block.level === 1 ? '제목' : '소제목'}` : `${current.page}쪽${block.heading ? ` · ${block.heading}` : ''}`;
      previewBody.innerHTML = markHtml(block.text, q).replace(/\n/g, '<br>');
      pager.hidden = true;
    } else {
      const total = view.pageCount;
      previewPage = Math.min(Math.max(previewPage, 1), total);
      const { structure } = await structureOf(previewPage);
      previewHeading.textContent = `${previewPage}쪽 전체 글`;
      previewBody.innerHTML = structure.text
        ? structure.blocks.map((b) => `<p class="${b.isHeading ? 'is-heading' : ''}">${markHtml(b.text, q).replace(/\n/g, '<br>')}</p>`).join('')
        : '<p class="none">이 쪽에는 글자 정보가 없습니다.</p>';
      pager.hidden = false;
      pageNum.textContent = `${previewPage} / ${total}쪽`;
      ($('sx-page-prev') as HTMLButtonElement).disabled = previewPage <= 1;
      ($('sx-page-next') as HTMLButtonElement).disabled = previewPage >= total;
    }
    previewBody.scrollTop = 0;
    previewBody.querySelector('mark')?.scrollIntoView({ block: 'nearest' });
  }

  function paintHits(): void {
    view.highlight(results.map((r) => r.match), results[index]?.match ?? null);
  }

  async function select(i: number, scroll = true): Promise<void> {
    if (results.length === 0) return;
    index = ((i % results.length) + results.length) % results.length;
    const r = results[index];
    if (index >= listLimit) {
      listLimit = Math.ceil((index + 1) / MAX_LIST) * MAX_LIST;
      renderList();
    }
    markSelected();
    renderCount();
    paintHits();
    const { structure } = await structureOf(r.page);
    view.setFocusBlock(r.page, structure.blocks[r.block]?.bbox ?? null);
    if (scroll) await view.scrollToMatch(r.match);
    previewPage = r.page;
    await renderPreview();
  }

  // ---------- 찾기 ----------
  async function run(): Promise<void> {
    const gen = ++generation;
    const query = input.value;
    results = [];
    index = -1;
    listLimit = MAX_LIST;
    searching = query.trim() !== '';
    view.setFocusBlock(null);
    renderControls();
    renderCount();
    renderList();
    renderEmpty('');
    paintHits();
    if (!searching) {
      preview.hidden = previewMode === 'block';
      return;
    }
    const allowed = scope === 'within' && stages.length ? blocksOf(stages[stages.length - 1].results) : null;
    const pages = scope === 'page' ? [view.currentPage] : Array.from({ length: view.pageCount }, (_, k) => k + 1);
    let chars = 0;
    let lastPaint = 0;
    for (const page of pages) {
      // 제목 이어받기가 맞도록 앞 쪽 구조부터 차례로 만든다(이 페이지만 찾을 때도 바로 앞 쪽은 본다)
      if (scope === 'page' && page > 1 && !structures.has(page - 1)) await structureOf(page - 1);
      const { items, structure } = await structureOf(page);
      if (gen !== generation) return;
      chars += items.reduce((n, it) => n + it.str.trim().length, 0);
      const found = searchPage(structure, items, query, target, allowed);
      if (found.length) {
        results.push(...found);
        if (index < 0) void select(0);
        // 쪽마다 다시 그리면 긴 문서에서 느려 — 0.15초마다 한 번
        if (performance.now() - lastPaint > 150) {
          lastPaint = performance.now();
          renderList();
          markSelected();
          renderCount();
          paintHits();
        }
      }
    }
    if (gen !== generation) return;
    searching = false;
    renderControls();
    renderList();
    markSelected();
    renderCount();
    paintHits();
    if (results.length === 0) {
      if (chars === 0) renderEmpty(scope === 'page' ? '이 쪽에는 글자 정보가 없습니다. 스캔한 쪽이면 문자 인식을 하면 찾을 수 있습니다.' : '이 문서에는 글자 정보가 없습니다(스캔 문서). 문자 인식을 하면 글자를 찾을 수 있습니다.', true);
      else renderEmpty(scope === 'within' ? '앞 단계 결과 안에는 없습니다. 위 단계를 눌러 돌아갈 수 있습니다.' : target === 'headings' ? '제목에서는 찾지 못했습니다. ‘모든 글’로 바꿔 보세요.' : scope === 'page' ? '이 페이지에는 없습니다. ‘전체 문서’로 넓혀 보세요.' : '찾는 내용이 없습니다.');
      if (previewMode === 'block') preview.hidden = true;
    }
  }

  function schedule(): void {
    clearTimeout(timer);
    timer = window.setTimeout(() => void run(), 200);
  }

  // ---------- 입력 ----------
  input.addEventListener('input', () => {
    renderControls();
    schedule();
  });
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      if (results.length === 0) void run();
      else void select(index + (event.shiftKey ? -1 : 1));
    } else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
      if (results.length === 0) return;
      event.preventDefault();
      void select(index + (event.key === 'ArrowDown' ? 1 : -1));
    } else if (event.key === 'Escape') {
      event.preventDefault();
      api.close();
    }
  });
  $('find-next').addEventListener('click', () => void select(index + 1));
  $('find-prev').addEventListener('click', () => void select(index - 1));

  for (const b of scopeButtons) {
    b.addEventListener('click', () => {
      const next = b.dataset.scope as SearchScope;
      if (next === 'within') {
        if (results.length === 0) return;
        // 지금 결과를 한 단계로 쌓고 새 검색어를 받는다
        stages.push({ query: input.value.trim(), target, results: [...results] });
        scope = 'within';
        input.value = '';
        results = [];
        index = -1;
        renderControls();
        renderCount();
        renderList();
        renderEmpty('앞 단계 결과 안에서 찾을 내용을 넣으세요.');
        input.focus();
        return;
      }
      stages = [];
      scope = baseScope = next;
      renderControls();
      void run();
      input.focus();
    });
  }
  headingsOnly.addEventListener('change', () => {
    target = headingsOnly.checked ? 'headings' : 'all';
    renderControls();
    void run();
  });
  trail.addEventListener('click', (event) => {
    const chip = (event.target as HTMLElement).closest<HTMLButtonElement>('[data-stage]');
    if (!chip) return;
    const k = Number(chip.dataset.stage);
    const stage = stages[k];
    stages = stages.slice(0, k);
    scope = k === 0 ? baseScope : 'within';
    target = stage.target;
    input.value = stage.query;
    void run();
    input.focus();
  });
  list.addEventListener('click', (event) => {
    const more = (event.target as HTMLElement).closest('[data-more]');
    if (more) {
      listLimit += MAX_LIST;
      renderList();
      markSelected();
      return;
    }
    const item = (event.target as HTMLElement).closest<HTMLElement>('.pdf-sx-item');
    if (item) void select(Number(item.dataset.i));
  });
  empty.addEventListener('click', (event) => {
    if ((event.target as HTMLElement).closest('[data-ocr]')) opts.onRunOcr();
  });
  for (const t of tabButtons) {
    t.addEventListener('click', () => {
      previewMode = t.dataset.view as 'block' | 'page';
      if (previewMode === 'page' && !results[index]) previewPage = view.currentPage;
      void renderPreview();
    });
  }
  $('sx-page-prev').addEventListener('click', () => {
    previewPage -= 1;
    view.goTo(previewPage);
    void renderPreview();
  });
  $('sx-page-next').addEventListener('click', () => {
    previewPage += 1;
    view.goTo(previewPage);
    void renderPreview();
  });
  $('sx-copy').addEventListener('click', async () => {
    const text = previewBody.innerText.trim();
    if (!text) return;
    const btn = $('sx-copy');
    try {
      await navigator.clipboard.writeText(text);
      btn.classList.add('done');
      btn.title = '복사했습니다';
      setTimeout(() => {
        btn.classList.remove('done');
        btn.title = '글 복사';
      }, 1500);
    } catch {
      /* 클립보드를 쓸 수 없는 환경 — 글을 직접 골라 복사할 수 있다 */
    }
  });

  // 쪽 정리로 쪽이 바뀌면 구조를 버리고 다시 찾는다
  view.onPagesChange(() => {
    structures.clear();
    stages = [];
    if (scope === 'within') scope = baseScope;
    if (!section.hidden && input.value.trim()) void run();
  });
  // '이 페이지' 범위는 쪽을 넘기면 그 쪽에서 다시
  view.onPageChange(() => {
    if (scope === 'page' && !section.hidden && input.value.trim()) schedule();
  });
  // 다른 패널로 바뀌거나 닫히면 강조를 지운다
  new MutationObserver(() => {
    if (section.hidden) {
      view.highlight([], null);
      view.setFocusBlock(null);
    } else if (results.length) {
      paintHits();
      void select(index < 0 ? 0 : index, false);
    }
  }).observe(section, { attributes: true, attributeFilter: ['hidden'] });

  const api: SearchPanel = {
    open(initial?: string) {
      opts.showPanel(true);
      if (initial !== undefined && initial.trim() && initial !== input.value) {
        stages = [];
        scope = baseScope;
        input.value = initial;
        void run();
      }
      renderControls();
      input.focus();
      input.select();
    },
    close() {
      generation += 1;
      opts.showPanel(false);
    },
  };
  renderControls();
  return api;
}
