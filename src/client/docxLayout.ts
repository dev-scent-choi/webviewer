import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { ReplaceStep } from 'prosemirror-transform';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { blocksDOM, fontStack, type ParagraphIndent, type ParagraphSpacing, type TabStop } from '../docx/schema.js';
import { NumberingCounter, type NumberingDefs } from '../docx/numbering.js';
import type { DocxRenderExtras, DocxSectionInfo, DocxStyleRule } from '../docx/import.js';

/**
 * Word 문서 모양 중 문서 내용(ProseMirror 노드)이 아니라 편집기가 덧그리는 것:
 *  - 번호 매기기 글자(1. / 1.1.1. / 가) / (ㄱ) / ●) — 원본 numbering.xml 정의로 문서 순서대로 센다
 *  - 쪽 경계(쪽 나누기·구역 나누기·문단 앞 쪽 나눔)마다 각주·바닥글·쪽 간격·머리글(PAGE·NUMPAGES 필드는 쪽 번호로)
 *  - 탭 정지(왼쪽·가운데·오른쪽·소수점, 점선 채움) — 그려진 글자 위치를 재어 탭 폭을 정한다
 *  - 문서 스타일(제목·부제·인용·문자 스타일 등)의 색·크기·글꼴·간격 → CSS 규칙
 * 모두 읽기 전용 장식이라 저장 내용에 들어가지 않는다(머리글·각주 원본 파트는 저장 때 그대로 남는다).
 * 쪽은 명시적 나눔(쪽 나누기·구역·문단 앞 쪽 나눔)과, 그려진 블록 높이를 재어 쪽 본문 높이를 넘는 곳(블록 단위 — 한 블록 안에서는
 * 나누지 않는다)에서 바뀐다. 구역마다 본문 폭이 다르면(세로·가로 쪽) 그 구역 블록의 폭을 그 구역 본문 폭으로 줄인다.
 *
 * 성능(400쪽 문서에서 글자 하나 16ms 안): 글자 입력처럼 한 문단 안에서만 바뀐 편집은 장식을 다시 만들지 않고 옮기기만 하고(번호·쪽은
 * 그대로), 탭 폭은 바뀐 문단만 다시 잰다. 잴 때는 모든 쓰기를 먼저 하고 읽기를 모아서 해 문서 전체 다시 배치를 한 번으로 줄인다.
 * 넘침 쪽 나눔은 입력이 멈춘 뒤(0.4초) 한 번 잰다.
 */

const TWIPS_PER_PX = 15;

interface LayoutState {
  readonly deco: DecorationSet;
  /** 넘침으로 생긴 쪽 경계(그 앞에서 쪽이 바뀌는 본문 직속 블록의 위치) */
  readonly autoBreaks: readonly number[];
  /** 이번 거래로 다시 재야 할 범위 — 'all'이면 문서 전체 */
  readonly dirty: 'all' | ReadonlyArray<readonly [number, number]>;
}
const layoutKey = new PluginKey<LayoutState>('docxLayout');

function el(tag: string, className: string, text?: string): HTMLElement {
  const node = document.createElement(tag);
  node.className = className;
  node.contentEditable = 'false';
  if (text !== undefined) node.textContent = text;
  return node;
}

/** 머리글·바닥글 블록 — 필드(PAGE·NUMPAGES)를 이 쪽 번호로 바꿔 그린다 */
function headerFooterDOM(content: unknown[] | null | undefined, className: string, page: number, total: number): HTMLElement {
  const box = el('div', className);
  if (!content) return box;
  const inner = blocksDOM(content, 'div');
  for (const field of inner.querySelectorAll<HTMLElement>('[data-field]')) {
    if (field.dataset.field === 'PAGE') field.textContent = String(page);
    else if (field.dataset.field === 'NUMPAGES' || field.dataset.field === 'SECTIONPAGES') field.textContent = String(total);
  }
  box.append(inner);
  return box;
}

/** 이 쪽의 머리글·바닥글 — 구역 첫 쪽(titlePg면 first), 짝수 쪽(settings의 evenAndOddHeaders면 even), 아니면 default */
function headerFooterOf(extras: DocxRenderExtras, section: DocxSectionInfo | undefined, which: 'header' | 'footer', firstOfSection: boolean, page: number): unknown[] | null {
  if (!section) return null;
  const set = section[which];
  if (section.titlePg && firstOfSection) return set.first;
  return extras.evenAndOddHeaders && page % 2 === 0 ? set.even : set.default;
}

interface PageCursor {
  page: number;
  section: number;
  firstOfSection: boolean;
  notes: PMNode[];
}

/** 쪽 하나가 끝나는 곳의 각주·바닥글 */
function pageEndDOM(extras: DocxRenderExtras, cur: PageCursor, total: number, last: boolean): HTMLElement {
  const end = el('div', `docx-page-end${last ? ' docx-page-end-last' : ''}`);
  if (cur.notes.length > 0) {
    const notes = el('div', 'docx-footnotes');
    for (const ref of cur.notes) notes.append(blocksDOM(extras.notes.footnote[String(ref.attrs.noteId)] ?? [], 'div'));
    end.append(notes);
  }
  const section = extras.sections[cur.section];
  if (section?.page) end.style.minHeight = `${section.page.marginBottomPx}px`;
  end.append(headerFooterDOM(headerFooterOf(extras, section, 'footer', cur.firstOfSection, cur.page), 'docx-footer', cur.page, total));
  return end;
}

function pageStartDOM(extras: DocxRenderExtras, cur: PageCursor, total: number, first: boolean): HTMLElement {
  const start = el('div', `docx-page-start${first ? ' docx-page-start-first' : ''}`);
  const section = extras.sections[cur.section];
  if (section) start.style.setProperty('--hdr-top', `${section.headerPx}px`);
  if (section?.page && !first) start.style.height = `${section.page.marginTopPx}px`;
  start.append(headerFooterDOM(headerFooterOf(extras, section, 'header', cur.firstOfSection, cur.page), 'docx-header', cur.page, total));
  return start;
}

/** 구역 본문 폭(px) */
const contentWidthOf = (section: DocxSectionInfo | undefined): number | null =>
  section?.page ? section.page.widthPx - section.page.marginLeftPx - section.page.marginRightPx : null;

/** 번호 글자 + 번호 수준의 들여쓰기(문단에 직접 들여쓰기가 없을 때) */
function numberingDecorations(doc: PMNode, defs: NumberingDefs | null, out: Decoration[]): void {
  const counter = new NumberingCounter(defs);
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    const num = node.attrs.num as { numId: string; ilvl: number } | null;
    const result = num ? counter.next(String(num.numId), Number(num.ilvl) || 0) : null;
    if (!result) return false;
    const { label, level } = result;
    const indent = node.attrs.indent as ParagraphIndent | null;
    const hangingPx = (indent ? indent.hanging : level.hanging) / TWIPS_PER_PX;
    if (!indent && (level.left || level.hanging || level.firstLine)) {
      const textIndent = level.hanging ? -level.hanging : level.firstLine;
      out.push(Decoration.node(pos, pos + node.nodeSize, { style: `margin-left: ${level.left / TWIPS_PER_PX}px; text-indent: ${textIndent / TWIPS_PER_PX}px` }));
    }
    out.push(
      Decoration.widget(pos + 1, () => {
        const span = el('span', 'docx-num', level.suff === 'space' ? `${label} ` : label);
        if (level.suff === 'tab' && hangingPx > 0) {
          span.style.minWidth = `${hangingPx}px`;
          span.dataset.tab = String(hangingPx); // 번호가 내어쓰기보다 길면 layoutTabs가 다음 탭 위치까지 늘린다
        } else if (level.suff === 'tab') span.style.paddingRight = '0.5em';
        if (level.bold) span.style.fontWeight = '700';
        if (level.color) span.style.color = level.color;
        return span;
      }, { side: -1, key: `num-${label}-${level.suff}-${hangingPx}-${level.bold}-${level.color}` }),
    );
    return false;
  });
}

/** 본문 직속 블록(다단 구역 안 블록 포함) — after: 이 블록 뒤 쪽 경계를 둘 자리(다단 구역의 마지막 블록이면 다단 상자 밖) */
function bodyBlocksOf(doc: PMNode): Array<{ node: PMNode; pos: number; after: number }> {
  const blocks: Array<{ node: PMNode; pos: number; after: number }> = [];
  doc.forEach((node, offset) => {
    if (node.type.name === 'columns') {
      node.forEach((child, inner, index) => {
        const pos = offset + 1 + inner;
        blocks.push({ node: child, pos, after: index === node.childCount - 1 ? offset + node.nodeSize : pos + child.nodeSize });
      });
    } else blocks.push({ node, pos: offset, after: offset + node.nodeSize });
  });
  return blocks;
}

/** 쪽 경계 장식(머리글·바닥글·각주)과 구역 폭 — autoBreaks: 넘침으로 생긴 쪽 경계(블록 위치) */
function pageDecorations(doc: PMNode, extras: DocxRenderExtras, totalPages: number | null, autoBreaks: readonly number[], out: Decoration[]): void {
  if (extras.sections.length === 0) return;
  const hostContent = Math.max(0, ...extras.sections.map((s) => contentWidthOf(s) ?? 0));
  const auto = new Set(autoBreaks);
  const blocks = bodyBlocksOf(doc);
  // 쪽 수: 보기 화면(LibreOffice)이 센 실제 쪽 수가 있으면 그것, 없으면 명시적 나눔으로 센 수
  let counted = 1;
  const boundaries: Array<{ pos: number; nextSection: number | null }> = [];
  let section = 0;
  for (const { node, pos, after } of blocks) {
    if ((node.attrs.breakBefore || auto.has(pos)) && pos > 0) boundaries.push({ pos, nextSection: null });
    // 구역 본문 폭이 편집 영역(가장 넓은 구역)보다 좁으면 그 폭으로 — 세로 쪽 글이 가로 쪽 폭으로 길게 흐르지 않게
    const width = contentWidthOf(extras.sections[section]);
    if (width && width < hostContent - 0.5 && node.type.name !== 'table') {
      const indent = node.attrs.indent as ParagraphIndent | null;
      const max = width - ((indent?.left ?? 0) + (indent?.right ?? 0)) / TWIPS_PER_PX;
      out.push(Decoration.node(pos, pos + node.nodeSize, { style: `max-width: ${Math.max(40, max)}px` }));
    }
    let hasBreak = false;
    node.descendants((child) => {
      if (child.type.name === 'page_break') hasBreak = true;
      return !hasBreak;
    });
    const sect = node.attrs.sect;
    if (typeof sect === 'number') {
      section = sect + 1;
      if ((extras.sections[section]?.type ?? 'nextPage') !== 'continuous') boundaries.push({ pos: after, nextSection: section });
      else boundaries.push({ pos: after, nextSection: -section - 1 }); // 이어지는 구역: 쪽은 그대로, 구역만 바뀜
    } else if (hasBreak) boundaries.push({ pos: after, nextSection: null });
  }
  counted += boundaries.filter((b) => b.nextSection === null || b.nextSection >= 0).length;
  const total = totalPages ?? counted;

  const cur: PageCursor = { page: 1, section: 0, firstOfSection: true, notes: [] };
  out.push(Decoration.widget(0, () => pageStartDOM(extras, { ...cur }, total, true), { side: -1, key: 'page-1', ignoreSelection: true }));
  let next = 0;
  const notesUntil = (limit: number) => {
    doc.nodesBetween(Math.min(next, doc.content.size), Math.min(limit, doc.content.size), (node, pos) => {
      if (pos >= next && node.type.name === 'ooxml' && node.attrs.kind === 'note' && node.attrs.noteType === 'footnote') cur.notes.push(node);
    });
    next = limit;
  };
  for (const boundary of boundaries) {
    notesUntil(boundary.pos);
    if (boundary.nextSection !== null && boundary.nextSection < 0) {
      cur.section = -boundary.nextSection - 1;
      continue;
    }
    const ending = { ...cur, notes: [...cur.notes] };
    cur.page += 1;
    cur.notes = [];
    cur.firstOfSection = boundary.nextSection !== null;
    if (boundary.nextSection !== null) cur.section = boundary.nextSection;
    const starting = { ...cur };
    out.push(
      Decoration.widget(boundary.pos, () => {
        const wrap = el('div', 'docx-page-boundary');
        wrap.append(pageEndDOM(extras, ending, total, false), el('div', 'docx-page-gap'), pageStartDOM(extras, starting, total, false));
        return wrap;
      }, { side: -1, key: `page-${starting.page}-${starting.section}-${ending.notes.map((n) => n.attrs.noteId).join(',')}`, ignoreSelection: true }),
    );
  }
  notesUntil(doc.content.size);
  const lastPage = { ...cur, notes: [...cur.notes] };
  out.push(
    Decoration.widget(doc.content.size, () => {
      const end = pageEndDOM(extras, lastPage, total, true);
      const endnoteIds = Object.keys(extras.notes.endnote);
      if (endnoteIds.length > 0) {
        const endnotes = el('div', 'docx-footnotes docx-endnotes');
        for (const id of endnoteIds) endnotes.append(blocksDOM(extras.notes.endnote[id], 'div'));
        end.prepend(endnotes);
      }
      return end;
    }, { side: 1, key: `page-end-${lastPage.page}-${lastPage.notes.map((n) => n.attrs.noteId).join(',')}`, ignoreSelection: true }),
  );
}

function buildDecorations(doc: PMNode, extras: DocxRenderExtras, totalPages: number | null, autoBreaks: readonly number[]): DecorationSet {
  const out: Decoration[] = [];
  numberingDecorations(doc, extras.numbering, out);
  pageDecorations(doc, extras, totalPages, autoBreaks, out);
  return DecorationSet.create(doc, out);
}

/** 쪽·번호를 바꾸지 않는 편집인가 — 한 글줄 블록 안에서 글자(인라인)만 넣고 지운 것(각주 참조·쪽 나누기는 제외) */
function isLocalTextEdit(tr: Transaction): boolean {
  return tr.steps.every((step, i) => {
    if (!(step instanceof ReplaceStep)) return false;
    const { from, to, slice } = step as unknown as { from: number; to: number; slice: import('prosemirror-model').Slice };
    if (slice.openStart || slice.openEnd) return false;
    const doc = tr.docs[i];
    const $from = doc.resolve(from);
    if (!$from.parent.isTextblock || !$from.sameParent(doc.resolve(to))) return false;
    const special = (n: PMNode) => !n.isInline || n.type.name === 'page_break' || (n.type.name === 'ooxml' && n.attrs.kind === 'note');
    let ok = true;
    slice.content.forEach((n) => {
      if (special(n)) ok = false;
    });
    doc.nodesBetween(from, to, (n) => {
      if (n !== $from.parent && special(n)) ok = false;
    });
    return ok;
  });
}

/** 거래가 바꾼 범위(새 문서 위치) */
function changedRanges(tr: Transaction): Array<readonly [number, number]> {
  const ranges: Array<readonly [number, number]> = [];
  tr.steps.forEach((step, i) => {
    const after = tr.mapping.slice(i + 1);
    step.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => {
      ranges.push([after.map(newStart, -1), after.map(newEnd, 1)]);
    });
  });
  return ranges;
}

/** 탭 하나 뒤 글자 폭 — 다음 탭(또는 문단 끝)까지, 소수점 맞춤이면 첫 '.' 앞까지 */
function followingWidth(tab: HTMLElement, block: HTMLElement, untilDot: boolean): number {
  const range = document.createRange();
  range.setStartAfter(tab);
  const walker = document.createTreeWalker(block, NodeFilter.SHOW_ALL);
  walker.currentNode = tab;
  let end: { node: Node; offset: number } | null = null;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node instanceof HTMLElement && node.classList.contains('docx-tab')) {
      end = { node, offset: -1 };
      break;
    }
    if (untilDot && node.nodeType === Node.TEXT_NODE) {
      const dot = (node.textContent ?? '').indexOf('.');
      if (dot >= 0) {
        end = { node, offset: dot };
        break;
      }
    }
  }
  if (end && end.offset < 0) range.setEndBefore(end.node);
  else if (end) range.setEnd(end.node, end.offset);
  else range.setEnd(block, block.childNodes.length);
  return range.getBoundingClientRect().width;
}

const BLOCK_SELECTOR = 'p, h1, h2, h3, h4, h5, h6';

function tabStopsOf(block: HTMLElement): TabStop[] {
  try {
    return JSON.parse(block.dataset.tabs ?? '[]') as TabStop[];
  } catch {
    return [];
  }
}

/**
 * 주어진 문단들의 번호 글자 폭과 탭 폭을 정한다. 앞 탭 폭이 뒤 탭 위치를 바꾸므로 문단마다 k번째 탭을 차례로 정하되, 모든 문단의
 * k번째 탭을 한꺼번에 잰다(읽기를 모두 한 뒤 쓰기를 모두) — 문서 전체 다시 배치가 탭 수가 아니라 "문단 안 탭 수의 최대값"만큼만 일어난다.
 * 예전엔 탭마다 폭을 쓰고 곧바로 위치를 읽어 400쪽 문서에서 글자 하나에 12~15초가 걸렸다.
 */
function layoutBlocks(root: HTMLElement, blocks: readonly HTMLElement[], defaultTabPx: number): void {
  if (blocks.length === 0) return;
  // 화면 배율(편집기 확대·축소 CSS zoom) — 잰 값(화면 px)을 문서 px로 되돌린다(편집 영역 전체 한 번)
  const scale = root.offsetWidth > 0 ? root.getBoundingClientRect().width / root.offsetWidth : 1;
  const own = (block: HTMLElement, selector: string) =>
    [...block.querySelectorAll<HTMLElement>(selector)].filter((t) => t.parentElement?.closest(BLOCK_SELECTOR) === block);
  const work = blocks.map((block) => ({ block, labels: own(block, '.docx-num[data-tab]'), tabs: own(block, '.docx-tab'), stops: tabStopsOf(block) }));
  // 쓰기: 모두 처음 상태로
  for (const w of work) {
    for (const label of w.labels) label.style.width = '';
    for (const tab of w.tabs) {
      tab.style.width = '0px';
      tab.className = 'docx-tab';
    }
  }
  // 읽기: 문단 원점·번호 글자(내어쓰기보다 긴 번호는 다음 기본 탭 위치까지, Word와 같다)
  const origins = work.map((w) => w.block.getBoundingClientRect().left - (parseFloat(getComputedStyle(w.block).marginLeft) || 0) * scale);
  const labelWidths: Array<[HTMLElement, number]> = [];
  work.forEach((w, i) => {
    for (const label of w.labels) {
      const minPx = Number(label.dataset.tab);
      const natural = label.scrollWidth;
      if (natural > minPx + 0.5) {
        const x = (label.getBoundingClientRect().left - origins[i]) / scale;
        labelWidths.push([label, Math.ceil((x + natural + 1) / defaultTabPx) * defaultTabPx - x]);
      }
    }
  });
  for (const [label, width] of labelWidths) label.style.width = `${width}px`;
  const rounds = Math.max(0, ...work.map((w) => w.tabs.length));
  for (let k = 0; k < rounds; k += 1) {
    // 읽기만
    const widths: Array<[HTMLElement, number, string | null]> = [];
    work.forEach((w, i) => {
      const tab = w.tabs[k];
      if (!tab) return;
      const x = (tab.getBoundingClientRect().left - origins[i]) / scale;
      const stop = w.stops.find((s) => s.pos / TWIPS_PER_PX > x + 1.5); // 1.5: 앞 탭을 1px 덜 채웠으니 같은 정지에 다시 걸리지 않게
      if (!stop) {
        widths.push([tab, Math.ceil((x + 1) / defaultTabPx) * defaultTabPx - x, null]);
        return;
      }
      const pos = stop.pos / TWIPS_PER_PX;
      const val = stop.val === 'end' ? 'right' : stop.val === 'start' ? 'left' : stop.val;
      const follow = val === 'right' || val === 'center' || val === 'decimal' ? followingWidth(tab, w.block, val === 'decimal') / scale : 0;
      const leader = stop.leader ? `docx-tab-${stop.leader === 'hyphen' || stop.leader === 'underscore' || stop.leader === 'heavy' ? 'line' : 'dot'}` : null;
      widths.push([tab, pos - x - (val === 'center' ? follow / 2 : follow), leader]);
    });
    // 쓰기만 — 1px 덜: 오른쪽 끝에 딱 맞춘 탭(목차 쪽 번호)이 소수점 차이로 다음 줄로 넘어가지 않게
    for (const [tab, width, leader] of widths) {
      tab.style.width = `${Math.max(2, width - 1)}px`;
      if (leader) tab.classList.add(leader);
    }
  }
}

/** 탭·번호 글자가 있는 문단 — root 아래 전부(머리글·바닥글 포함) */
function blocksWithTabs(root: ParentNode): HTMLElement[] {
  const set = new Set<HTMLElement>();
  for (const el of root.querySelectorAll<HTMLElement>('.docx-tab, .docx-num[data-tab]')) {
    const block = el.parentElement?.closest<HTMLElement>(BLOCK_SELECTOR);
    if (block) set.add(block);
  }
  return [...set];
}

/**
 * 넘침 쪽 나눔: 그려진 본문 블록의 위·아래를 재어(쪽 경계 장식 높이는 빼고) 쪽 본문 높이를 넘는 블록 앞에 쪽 경계를 둔다.
 * 블록 하나가 쪽보다 크면(긴 표·긴 문단) 그 블록은 나누지 않는다.
 */
function measureAutoBreaks(view: EditorView, extras: DocxRenderExtras): number[] {
  const scale = view.dom.offsetWidth > 0 ? view.dom.getBoundingClientRect().width / view.dom.offsetWidth : 1;
  const blocks = bodyBlocksOf(view.state.doc);
  const top = new Map<number, { top: number; bottom: number }>();
  // 장식(쪽 경계) 높이를 빼고 본문만 이어 붙인 좌표
  let shift = 0;
  let widgetTop: number | null = null;
  const children = [...view.dom.children] as HTMLElement[];
  const posOfDom = new Map<Element, number>();
  for (const b of blocks) {
    const dom = view.nodeDOM(b.pos);
    if (dom instanceof HTMLElement) posOfDom.set(dom, b.pos);
  }
  const visit = (list: HTMLElement[]) => {
    for (const child of list) {
      if (child.classList.contains('docx-columns')) {
        visit([...child.children] as HTMLElement[]);
        continue;
      }
      const rect = child.getBoundingClientRect();
      if (child.matches('.docx-page-boundary, .docx-page-start, .docx-page-end')) {
        widgetTop ??= rect.top;
        continue;
      }
      if (widgetTop !== null) {
        shift += rect.top - widgetTop;
        widgetTop = null;
      }
      const pos = posOfDom.get(child);
      if (pos !== undefined) top.set(pos, { top: (rect.top - shift) / scale, bottom: (rect.bottom - shift) / scale });
    }
  };
  visit(children);
  const breaks: number[] = [];
  let section = 0;
  let pageTop: number | null = null;
  for (const b of blocks) {
    const m = top.get(b.pos);
    const page = extras.sections[section]?.page;
    const explicit = b.node.attrs.breakBefore && b.pos > 0;
    if (m && (pageTop === null || explicit)) pageTop = m.top;
    else if (m && page && pageTop !== null && m.bottom - pageTop > page.heightPx - page.marginTopPx - page.marginBottomPx + 1) {
      breaks.push(b.pos);
      pageTop = m.top;
    }
    let hasBreak = false;
    b.node.descendants((child) => {
      if (child.type.name === 'page_break') hasBreak = true;
      return !hasBreak;
    });
    if (typeof b.node.attrs.sect === 'number') {
      section = b.node.attrs.sect + 1;
      if ((extras.sections[section]?.type ?? 'nextPage') !== 'continuous') pageTop = null;
    } else if (hasBreak) pageTop = null;
  }
  return breaks;
}

function cssNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

const HEX = /^#[0-9a-f]{6}$/i;

/** 스타일 규칙 하나 → CSS 선언들(문서에서 온 값이라 색·숫자·글꼴 이름만 걸러 쓴다) */
function ruleDeclarations(rule: DocxStyleRule): string[] {
  const d: string[] = [];
  if (rule.color && HEX.test(rule.color)) d.push(`color: ${rule.color}`);
  const size = cssNumber(rule.sizePt);
  if (size && size > 0 && size < 400) d.push(`font-size: ${size}pt`);
  else if (rule.kind === 'p') d.push('font-size: inherit');
  if (rule.bold !== null && rule.bold !== undefined) d.push(`font-weight: ${rule.bold ? 700 : 400}`);
  if (rule.italic !== null && rule.italic !== undefined) d.push(`font-style: ${rule.italic ? 'italic' : 'normal'}`);
  if (rule.underline) d.push('text-decoration: underline');
  if (rule.caps) d.push('text-transform: uppercase');
  if (rule.font && (rule.font.eastAsia || rule.font.ascii)) d.push(`font-family: ${fontStack([rule.font.eastAsia, rule.font.ascii])}`);
  // 문단·제목은 정렬·들여쓰기를 늘 적는다 — 문서에 있는 스타일이면 편집기 갤러리 기본 모양(인용 가운데 맞춤 등)이 새어 나오지 않게
  if (rule.kind !== 'r') d.push(`text-align: ${rule.align && /^(left|right|center|justify)$/.test(rule.align) ? rule.align : 'left'}`);
  const spacing = rule.spacing as ParagraphSpacing | null | undefined;
  if (rule.kind !== 'r') {
    const before = cssNumber(spacing?.before);
    const after = cssNumber(spacing?.after);
    const line = cssNumber(spacing?.line);
    if (before !== null) d.push(`margin-top: ${before / TWIPS_PER_PX}px`);
    if (after !== null) d.push(`margin-bottom: ${after / TWIPS_PER_PX}px`);
    if (line) {
      d.push(spacing?.lineRule === 'exact' || spacing?.lineRule === 'atLeast'
        ? `line-height: ${line / TWIPS_PER_PX}px`
        : `line-height: calc(var(--doc-single-line, 1.2) * ${Number((line / 240).toFixed(3))})`);
    } else if (rule.kind === 'h') d.push('line-height: var(--doc-line-height, 1.3)');
    const indent = rule.indent;
    const textIndent = indent ? (indent.firstLine ? indent.firstLine : indent.hanging ? -indent.hanging : 0) : 0;
    d.push(`margin-left: ${(indent?.left ?? 0) / TWIPS_PER_PX}px`, `margin-right: ${(indent?.right ?? 0) / TWIPS_PER_PX}px`, `text-indent: ${textIndent / TWIPS_PER_PX}px`);
    for (const [side, border] of [['bottom', rule.borderBottom], ['top', rule.borderTop]] as const) {
      if (border && HEX.test(border.color)) d.push(`border-${side}: ${border.widthPx}px ${border.style} ${border.color}`, `padding-${side}: 4px`);
      else d.push(`border-${side}: 0`, `padding-${side}: 0`);
    }
    if (!rule.color) d.push('color: inherit');
    if (rule.italic === null || rule.italic === undefined) d.push('font-style: normal');
    if (rule.shading && HEX.test(rule.shading)) d.push(`background-color: ${rule.shading}`);
  }
  if (rule.kind !== 'r') d.push('letter-spacing: normal', 'text-wrap: wrap');
  return d;
}

/** 문서 스타일 → CSS 규칙 목록(하나씩 insertRule). 제목은 h1~h6, 문단 스타일은 p[data-style], 문자 스타일은 span[data-rstyle]. */
export function docxStyleSheetText(rules: readonly DocxStyleRule[]): string[] {
  const scope = '.editor-surface .ProseMirror';
  return rules
    .map((rule) => {
      const key = CSS.escape(rule.key);
      const selector = rule.kind === 'h'
        ? /^[1-6]$/.test(rule.key) ? `${scope} h${rule.key}` : null
        : rule.kind === 'p' ? `${scope} p[data-style="${key}"]` : `${scope} span[data-rstyle="${key}"]`;
      const decls = ruleDeclarations(rule);
      return selector && decls.length > 0 ? `${selector} { ${decls.join('; ')}; }` : '';
    })
    .filter(Boolean);
}

const PAGINATION_META = 'docxPagination';

/** 편집기에 Word 문서 모양 장식을 붙이는 플러그인 — 번호·쪽 경계·구역 폭 장식, 탭 폭, 넘침 쪽 나눔(성능은 파일 위 설명). */
export function docxLayoutPlugin(extras: DocxRenderExtras, totalPages: number | null): Plugin<LayoutState> {
  const defaultTabPx = extras.defaultTabStopDxa / TWIPS_PER_PX;
  return new Plugin<LayoutState>({
    key: layoutKey,
    state: {
      init: (_, state) => ({ deco: buildDecorations(state.doc, extras, totalPages, []), autoBreaks: [], dirty: 'all' }),
      apply: (tr, prev, _, state) => {
        const paged = tr.getMeta(PAGINATION_META) as number[] | undefined;
        if (paged) return { deco: buildDecorations(state.doc, extras, totalPages, paged), autoBreaks: paged, dirty: 'all' };
        if (!tr.docChanged) return prev.dirty === 'all' || prev.dirty.length > 0 ? { ...prev, dirty: [] } : prev;
        const autoBreaks = prev.autoBreaks.map((p) => tr.mapping.map(p));
        if (isLocalTextEdit(tr)) return { deco: prev.deco.map(tr.mapping, tr.doc), autoBreaks, dirty: changedRanges(tr) };
        return { deco: buildDecorations(state.doc, extras, totalPages, autoBreaks), autoBreaks, dirty: 'all' };
      },
    },
    props: {
      decorations: (state) => layoutKey.getState(state)?.deco,
    },
    view: (view: EditorView) => {
      let frame = 0;
      let pending: 'all' | Set<HTMLElement> = 'all';
      let paginateTimer = 0;
      const flush = () => {
        frame = 0;
        const targets = pending === 'all' ? blocksWithTabs(view.dom) : [...pending].filter((b) => b.isConnected);
        pending = new Set();
        layoutBlocks(view.dom, targets, defaultTabPx);
      };
      const schedule = (what: 'all' | HTMLElement[]) => {
        if (what === 'all') pending = 'all';
        else if (pending !== 'all') for (const b of what) pending.add(b);
        if (!frame) frame = requestAnimationFrame(flush);
      };
      const paginate = () => {
        window.clearTimeout(paginateTimer);
        paginateTimer = window.setTimeout(() => {
          const breaks = measureAutoBreaks(view, extras);
          const current = layoutKey.getState(view.state)?.autoBreaks ?? [];
          if (breaks.join() !== current.join()) view.dispatch(view.state.tr.setMeta(PAGINATION_META, breaks).setMeta('addToHistory', false));
        }, 400);
      };
      schedule('all');
      paginate();
      // 글꼴이 늦게 내려오면 글자 폭이 바뀐다 — 다 내려온 뒤 한 번 더 잰다
      void document.fonts?.ready.then(() => {
        schedule('all');
        paginate();
      });
      return {
        update: (v, prevState) => {
          if (v.state.doc === prevState.doc && layoutKey.getState(v.state)?.deco === layoutKey.getState(prevState)?.deco) return;
          const dirty = layoutKey.getState(v.state)?.dirty ?? 'all';
          if (dirty === 'all') schedule('all');
          else {
            const touched: HTMLElement[] = [];
            for (const [from, to] of dirty) {
              v.state.doc.nodesBetween(Math.max(0, from), Math.min(v.state.doc.content.size, to), (node, pos) => {
                if (!node.isTextblock) return true;
                const dom = v.nodeDOM(pos);
                if (dom instanceof HTMLElement && dom.querySelector('.docx-tab, .docx-num[data-tab]')) touched.push(dom);
                return false;
              });
            }
            if (touched.length > 0) schedule(touched);
          }
          if (v.state.doc !== prevState.doc) paginate();
        },
        destroy: () => {
          cancelAnimationFrame(frame);
          window.clearTimeout(paginateTimer);
        },
      };
    },
  });
}
