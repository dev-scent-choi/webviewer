/**
 * PDF 편집기 오른쪽 '찾기' 패널 — 3단 검색(① 이 페이지 ② 전체 문서 ③ 결과 안에서 = 앞 결과가 속한 절 안에서),
 * 찾을 곳(전체/제목/본문), 쪽별로 묶은 결과 목록(속한 제목 + 앞뒤 문장), 고른 결과의 덩어리 테두리, 절·쪽 전체 글 미리보기.
 * 줄·문단·제목 나누기는 textStructure.ts, 찾기는 searchEngine.ts(둘 다 순수 함수).
 */
import { normalizeForFind } from './findText.js';
import { readOutline, type OutlineNode, type PDFDocumentProxy } from './pdfDocument.js';
import type { PageView, TextItemInfo } from './pageView.js';
import { groupByPage, searchPage, sectionsOf, sectionText, type SearchResult, type SearchScope, type SearchTarget } from './searchEngine.js';
import { analyzePage, buildProfile, type DocProfile, type PageStructure, type StructItem } from './textStructure.js';

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
/** 문서 기준(본문 서식·머리말)을 구할 때 고루 뽑아 보는 쪽 수 */
const PROFILE_PAGES = 30;
/** '이 페이지'만 찾을 때 이어받을 제목을 찾아 거슬러 보는 쪽 수 */
const LOOKBACK = 8;
/** 절 미리보기가 따라가는 최대 쪽 수 */
const SECTION_PAGES = 20;
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
  const targetButtons = [...section.querySelectorAll<HTMLButtonElement>('[data-target]')];
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

  // 쪽 구조는 글 조각 배열·이어받은 제목이 같을 때만 다시 쓴다(문자 인식으로 글이 더해지면 새 배열이 온다)
  interface Entry {
    items: readonly TextItemInfo[];
    structure: PageStructure;
    /** 만들 때 앞 쪽에서 이어받은 제목·절 */
    carried: string;
  }
  const structures = new Map<number, Entry>();
  let profile: Promise<DocProfile> | null = null;
  let outlineTitles: Set<string> = new Set();
  void readOutline(doc)
    .then((nodes) => {
      const walk = (ns: readonly OutlineNode[]): string[] => ns.flatMap((n) => [normalizeForFind(n.title).trim(), ...walk(n.children)]);
      outlineTitles = new Set(walk(nodes));
      structures.clear();
    })
    .catch(() => undefined);

  /** 글꼴 굵기·태그 구조 역할을 조각에 붙인다(못 받으면 크기·번호 같은 단서로만 가린다) */
  async function styledItems(page: number, items: readonly TextItemInfo[]): Promise<StructItem[]> {
    const meta = await view.textMeta(page);
    return items.map((it, i) => ({ ...it, bold: it.fontName ? meta.fonts[it.fontName]?.bold ?? false : false, role: meta.roles[i] ?? null }));
  }

  /** 문서 기준 — 쪽 하나만 보면 조문 상자·해설이 섞인 쪽에서 본문 서식이 흔들려 여러 쪽에서 한 번 구한다 */
  function profileOf(): Promise<DocProfile> {
    profile ??= (async () => {
      const total = view.pageCount;
      const n = Math.min(PROFILE_PAGES, total);
      const pages = [...new Set(Array.from({ length: n }, (_, k) => 1 + Math.floor((k * total) / n)))];
      const sample = await Promise.all(
        pages.map(async (p) => ({ items: await styledItems(p, await view.textItemInfos(p)), pageHeight: view.sizeOf(p).h })),
      );
      const found = buildProfile(sample);
      // 글이 하나도 없던 스캔 문서는 문자 인식 뒤에 다시 구한다
      if (!found.bodyStyle) profile = null;
      return found;
    })();
    return profile;
  }

  /** 쪽 구조 — 앞 쪽 끝의 제목·절을 이어받는다(앞 쪽이 아직 없으면 몇 쪽 거슬러 올라가 만들어 둔다) */
  async function structureOf(page: number): Promise<Entry> {
    const items = await view.textItemInfos(page);
    let prevHeading = '';
    let prevSection = '';
    if (page > 1) {
      if (!structures.has(page - 1)) {
        // 그보다 앞은 '전체 문서'로 찾을 때 차례로 채워진다
        let from = page - 1;
        while (from > 1 && page - from < LOOKBACK && !structures.has(from - 1)) from -= 1;
        for (let p = from; p < page; p += 1) await structureOf(p);
      }
      const prev = structures.get(page - 1)!.structure;
      prevHeading = prev.lastHeading;
      prevSection = prev.lastSection;
    }
    const carried = `${prevHeading}\u0000${prevSection}`;
    const hit = structures.get(page);
    if (hit && hit.items === items && hit.carried === carried) return hit;
    const [styled, docProfile] = await Promise.all([styledItems(page, items), profileOf()]);
    const structure = analyzePage(page, styled, prevHeading, outlineTitles, { pageHeight: view.sizeOf(page).h, profile: docProfile }, prevSection);
    const entry = { items, structure, carried };
    structures.set(page, entry);
    return entry;
  }

  /** 절 하나의 글 — 절이 다음 쪽으로 이어지면 그 쪽 구조도 만들어 따라간다 */
  async function sectionOf(key: string): Promise<{ heading: string; text: string; pages: number[] }> {
    const start = Number(key.replace(/^b:/, '').split(':')[0]);
    for (let upto = start; ; upto += 1) {
      await structureOf(upto);
      const found = sectionText(key, new Map([...structures].map(([p, e]) => [p, e.structure])));
      const last = found.pages[found.pages.length - 1] ?? start;
      if (key.startsWith('b:') || last < upto || upto >= view.pageCount || upto - start >= SECTION_PAGES) return found;
    }
  }

  // ---------- 상태 그리기 ----------
  function setChecked(buttons: HTMLButtonElement[], attr: string, value: string): void {
    for (const b of buttons) b.setAttribute('aria-checked', String(b.dataset[attr] === value));
  }

  function renderControls(): void {
    setChecked(scopeButtons, 'scope', scope);
    setChecked(targetButtons, 'target', target);
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
      const sec = await sectionOf(current.section);
      if (results[index] !== current) return;
      const span = sec.pages.length > 1 ? `${sec.pages[0]}~${sec.pages[sec.pages.length - 1]}쪽` : `${sec.pages[0] ?? current.page}쪽`;
      previewHeading.textContent = sec.heading ? `${span} · ${sec.heading}` : `${span} · 첫 제목 앞의 글`;
      previewBody.innerHTML = sec.text
        .split('\n\n')
        .map((t, k) => `<p class="${k === 0 && sec.heading && !current.section.startsWith('b:') ? 'is-heading' : ''}">${markHtml(t.replace(/\n/g, ' '), q)}</p>`)
        .join('');
      pager.hidden = true;
    } else {
      const total = view.pageCount;
      previewPage = Math.min(Math.max(previewPage, 1), total);
      const { structure } = await structureOf(previewPage);
      previewHeading.textContent = `${previewPage}쪽 전체 글`;
      previewBody.innerHTML = structure.text
        ? structure.blocks
            .filter((b) => !b.furniture)
            // PDF 줄바꿈은 지면 폭 때문이라 미리보기에선 한 칸으로 잇는다(줄을 넘은 검색어도 표시된다)
            .map((b) => `<p class="${b.isHeading ? 'is-heading' : ''}">${markHtml(b.text.replace(/\n/g, ' '), q)}</p>`)
            .join('')
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
    // '결과 안에서'는 앞 단계 결과가 속한 절(제목부터 다음 제목 전까지) 안에서만 찾는다
    const allowed = scope === 'within' && stages.length ? sectionsOf(stages[stages.length - 1].results) : null;
    const pages = scope === 'page' ? [view.currentPage] : Array.from({ length: view.pageCount }, (_, k) => k + 1);
    let chars = 0;
    let lastPaint = 0;
    for (const page of pages) {
      // 제목·절 이어받기가 맞도록 앞 쪽부터 차례로(이 페이지만 찾을 때는 structureOf 가 앞 몇 쪽을 거슬러 본다)
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
      else
        renderEmpty(
          scope === 'within'
            ? '앞 단계 결과가 속한 절 안에는 없습니다. 위 단계를 눌러 돌아갈 수 있습니다.'
            : target === 'headings'
              ? '제목에서는 찾지 못했습니다. 찾을 곳을 ‘전체’로 바꿔 보세요.'
              : target === 'body'
                ? '본문에서는 찾지 못했습니다(제목·머리말·쪽 번호는 빼고 찾았습니다). 찾을 곳을 ‘전체’로 바꿔 보세요.'
                : scope === 'page'
                  ? '이 페이지에는 없습니다. ‘전체 문서’로 넓혀 보세요.'
                  : '찾는 내용이 없습니다.',
        );
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
  for (const b of targetButtons) {
    b.addEventListener('click', () => {
      target = b.dataset.target as SearchTarget;
      renderControls();
      void run();
    });
  }
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
    profile = null;
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
