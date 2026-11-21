import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import { ReplaceAroundStep, ReplaceStep } from 'prosemirror-transform';
import { Decoration, DecorationSet, type EditorView } from 'prosemirror-view';
import { DOMSerializer, type Node as PMNode } from 'prosemirror-model';
import { blocksDOM, fontStack, lineSpacingCss, type ParagraphIndent, type ParagraphSpacing, type TabStop } from '../docx/schema.js';
import { NumberingCounter, type NumberingDefs } from '../docx/numbering.js';
import type { DocxRenderExtras, DocxSectionInfo, DocxStyleRule } from '../docx/import.js';

/**
 * Word 문서 모양 중 문서 내용(ProseMirror 노드)이 아니라 편집기가 덧그리는 것:
 *  - 번호 매기기 글자(1. / 1.1.1. / 가) / (ㄱ) / ●) — 원본 numbering.xml 정의로 문서 순서대로 센다
 *  - 쪽 경계(쪽 나누기·구역 나누기·문단 앞 쪽 나눔)마다 각주·바닥글·쪽 간격·머리글(PAGE·NUMPAGES 필드는 쪽 번호로)
 *  - 탭 정지(왼쪽·가운데·오른쪽·소수점, 점선 채움) — 그려진 글자 위치를 재어 탭 폭을 정한다
 *  - 문서 스타일(제목·부제·인용·문자 스타일 등)의 색·크기·글꼴·간격 → CSS 규칙
 * 모두 읽기 전용 장식이라 저장 내용에 들어가지 않는다(머리글·각주 원본 파트는 저장 때 그대로 남는다).
 * 쪽은 명시적 나눔(쪽 나누기·구역·문단 앞 쪽 나눔)과 넘침(블록 높이표로 센다 — paginateFlow: 각주 높이를 빼고, 긴 문단은 줄 단위,
 * 긴 표·다단은 쪽 수만)에서 바뀐다. 구역마다 본문 폭이 다르면(세로·가로 쪽) 그 구역 블록의 폭을 그 구역 본문 폭으로 줄인다.
 *
 * 성능: 글자 입력처럼 한 문단 안에서만 바뀐 편집은 장식을 다시 만들지 않고 옮기기만 하고(번호·쪽은 그대로), 번호 들여쓰기는 노드 장식
 * 대신 CSS 규칙, 탭 폭은 바뀐 문단과 화면 안 문단만 잰다(모든 쓰기 → 모든 읽기). 화면에서 먼 블록은 그리기를 건너뛰고(content-visibility:
 * hidden), 넘침 쪽 나눔은 블록 높이가 바뀌거나 구조가 바뀐 뒤 0.8초 쉬면 한 번 잰다(그려진 위치를 읽지 않는다).
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
  notes: Array<{ node: PMNode; label: string }>;
}

/** 각주·미주 본문 — 본문 안 번호(w:footnoteRef)도 문서 순서대로 매긴 번호로(복사한 각주는 새 번호, Word와 같다) */
function noteBodyDOM(content: unknown[] | undefined, label: string): HTMLElement {
  const body = blocksDOM(content ?? [], 'div');
  for (const ref of body.querySelectorAll<HTMLElement>('.docx-note-ref')) ref.textContent = label;
  return body;
}

/** 쪽 하나가 끝나는 곳의 각주·바닥글 */
function pageEndDOM(extras: DocxRenderExtras, cur: PageCursor, total: number, last: boolean): HTMLElement {
  const end = el('div', `docx-page-end${last ? ' docx-page-end-last' : ''}`);
  if (cur.notes.length > 0) {
    const notes = el('div', 'docx-footnotes');
    for (const ref of cur.notes) notes.append(noteBodyDOM(extras.notes.footnote[String(ref.node.attrs.noteId)], ref.label));
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
/**
 * 번호 문단의 들여쓰기(문단에 직접 들여쓰기가 없을 때 번호 수준 값) — 문단마다 노드 장식을 달면 글자 하나 칠 때마다 장식 옮기기가
 * 블록 수의 제곱만큼 걸린다(400쪽 1600문단에서 글자당 약 70ms). 그래서 (번호, 수준)마다 CSS 규칙 하나로 둔다(문단의 data-numlvl).
 */
function numberingDecorations(doc: PMNode, defs: NumberingDefs | null, out: Decoration[], rules: Map<string, string>): void {
  const counter = new NumberingCounter(defs);
  doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    const num = node.attrs.num as { numId: string; ilvl: number } | null;
    const result = num ? counter.next(String(num.numId), Number(num.ilvl) || 0) : null;
    if (!result || !num) return false;
    const { label, level } = result;
    const indent = node.attrs.indent as ParagraphIndent | null;
    const hangingPx = (indent ? indent.hanging : level.hanging) / TWIPS_PER_PX;
    const key = `${num.numId}:${Number(num.ilvl) || 0}`;
    if (!rules.has(key) && /^[\w:.-]{1,80}$/.test(key) && (level.left || level.hanging || level.firstLine)) {
      const textIndent = level.hanging ? -level.hanging : level.firstLine;
      rules.set(key, `.editor-surface .ProseMirror [data-numlvl="${key}"]:not([data-indent]) { margin-left: ${level.left / TWIPS_PER_PX}px; text-indent: ${textIndent / TWIPS_PER_PX}px; }`);
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
function bodyBlocksOf(doc: PMNode): Array<{ node: PMNode; pos: number; after: number; inColumns: boolean }> {
  const blocks: Array<{ node: PMNode; pos: number; after: number; inColumns: boolean }> = [];
  doc.forEach((node, offset) => {
    if (node.type.name === 'columns') {
      node.forEach((child, inner, index) => {
        const pos = offset + 1 + inner;
        blocks.push({ node: child, pos, after: index === node.childCount - 1 ? offset + node.nodeSize : pos + child.nodeSize, inColumns: true });
      });
    } else blocks.push({ node, pos: offset, after: offset + node.nodeSize, inColumns: false });
  });
  return blocks;
}

/** 쪽 경계 장식(머리글·바닥글·각주)과 구역 폭 — autoBreaks: 넘침으로 생긴 쪽 경계(블록 위치) */
function pageDecorations(doc: PMNode, extras: DocxRenderExtras, totalPages: number | null, autoBreaks: readonly number[], out: Decoration[], numbers: PageNumbers): void {
  if (extras.sections.length === 0) return;
  const hostContent = Math.max(0, ...extras.sections.map((s) => contentWidthOf(s) ?? 0));
  const auto = [...autoBreaks].sort((a, b) => a - b);
  let ai = 0;
  const blocks = bodyBlocksOf(doc);
  // 쪽 수: 보기 화면(LibreOffice)이 센 실제 쪽 수가 있으면 그것, 없으면 명시적 나눔으로 센 수
  let counted = 1;
  // inline: 문단 안(긴 문단의 넘치는 줄 앞) · hidden: 표·다단 안 — 경계를 그리지 않고 쪽 수만 센다
  const boundaries: Array<{ pos: number; nextSection: number | null; inline?: boolean; hidden?: boolean }> = [];
  let section = 0;
  for (const { node, pos, after, inColumns } of blocks) {
    while (ai < auto.length && auto[ai] < pos) ai += 1;
    const autoHere = ai < auto.length && auto[ai] === pos;
    if (autoHere) ai += 1;
    if ((node.attrs.breakBefore || autoHere) && pos > 0) boundaries.push({ pos, nextSection: null });
    for (; ai < auto.length && auto[ai] < after; ai += 1) {
      const inline = node.isTextblock && !inColumns;
      boundaries.push({ pos: auto[ai], nextSection: null, inline, hidden: !inline });
    }
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
  numbers.total = total;
  numbers.seq.push(1);
  out.push(Decoration.widget(0, () => pageStartDOM(extras, { ...cur }, total, true), { side: -1, key: 'page-1', ignoreSelection: true }));
  // 장식 key엔 쪽 번호를 넣지 않는다 — 쪽 경계 하나가 생기거나 없어져도 뒤 쪽들의 머리글·바닥글 DOM을 다시 만들지 않게(REVIEW-r3 I-r3-2).
  // 번호(PAGE·NUMPAGES 필드)는 그린 뒤 renumberPages가 글자만 바꾼다. 짝수 쪽 머리글을 따로 쓰는 문서만 홀짝을 key에 넣는다.
  const parity = (page: number) => (extras.evenAndOddHeaders ? `${page % 2}` : '');
  let next = 0;
  // 각주·미주 번호는 문서 순서대로 다시 센다 — 같은 문서 안에서 복사한 참조도 저장 때 새 각주가 되므로(export.ts) 화면도 새 번호로
  const counters = { footnote: 0, endnote: 0 };
  const endnoteRefs: Array<{ node: PMNode; label: string }> = [];
  const notesUntil = (limit: number) => {
    doc.nodesBetween(Math.min(next, doc.content.size), Math.min(limit, doc.content.size), (node, pos) => {
      if (pos < next || node.type.name !== 'ooxml' || node.attrs.kind !== 'note' || !node.attrs.noteType) return;
      const type = node.attrs.noteType === 'endnote' ? 'endnote' : 'footnote';
      counters[type] += 1;
      const label = String(counters[type]);
      if (label !== node.attrs.text) out.push(Decoration.node(pos, pos + node.nodeSize, { 'data-label': label }));
      (type === 'footnote' ? cur.notes : endnoteRefs).push({ node, label });
    });
    next = limit;
  };
  for (const boundary of boundaries) {
    notesUntil(boundary.pos);
    if (boundary.nextSection !== null && boundary.nextSection < 0) {
      cur.section = -boundary.nextSection - 1;
      continue;
    }
    if (boundary.hidden) {
      cur.page += 1;
      continue;
    }
    const ending = { ...cur, notes: [...cur.notes] };
    numbers.seq.push(ending.page, ending.page + 1);
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
      }, { side: -1, key: `page-${ending.section}${ending.firstOfSection ? 'f' : ''}${parity(ending.page)}-${starting.section}${starting.firstOfSection ? 'f' : ''}${parity(starting.page)}-${ending.notes.map((n) => `${n.node.attrs.noteId}:${n.label}`).join(',')}`, ignoreSelection: true }),
    );
  }
  notesUntil(doc.content.size);
  const lastPage = { ...cur, notes: [...cur.notes] };
  numbers.seq.push(lastPage.page);
  out.push(
    Decoration.widget(doc.content.size, () => {
      const end = pageEndDOM(extras, lastPage, total, true);
      if (endnoteRefs.length > 0) {
        const endnotes = el('div', 'docx-footnotes docx-endnotes');
        for (const ref of endnoteRefs) endnotes.append(noteBodyDOM(extras.notes.endnote[String(ref.node.attrs.noteId)], ref.label));
        end.prepend(endnotes);
      }
      return end;
    }, { side: 1, key: `page-end-${lastPage.section}${parity(lastPage.page)}-${[...lastPage.notes, ...endnoteRefs].map((n) => `${n.node.attrs.noteId}:${n.label}`).join(',')}`, ignoreSelection: true }),
  );
}

/** 번호 수준 들여쓰기 규칙(numberingDecorations) — 문서에 쓰인 (번호, 수준)마다 하나, 편집기 하나에 한 벌 */
let numberingStyle: HTMLStyleElement | null = null;
const numberingRules = new Map<string, string>();

/** 쪽 시작·끝 장식(DOM 순서)의 쪽 번호와 전체 쪽 수 — 장식 묶음마다 */
interface PageNumbers {
  seq: number[];
  total: number;
}
const pageNumbersOf = new WeakMap<DecorationSet, PageNumbers>();

/** 쪽 시작·끝 장식의 PAGE·NUMPAGES 필드 글자를 지금 쪽 번호로 — 바뀐 것만 쓴다 */
function renumberPages(root: HTMLElement, numbers: PageNumbers | undefined): void {
  if (!numbers) return;
  root.querySelectorAll<HTMLElement>('.docx-page-start, .docx-page-end').forEach((box, i) => {
    const page = String(numbers.seq[i] ?? '');
    if (box.dataset.page === page) return;
    box.dataset.page = page;
    for (const field of box.querySelectorAll<HTMLElement>('[data-field="PAGE"]')) field.textContent = page;
  });
  const total = String(numbers.total);
  if (root.dataset.pages !== total) {
    root.dataset.pages = total;
    for (const field of root.querySelectorAll<HTMLElement>('[data-field="NUMPAGES"], [data-field="SECTIONPAGES"]')) field.textContent = total;
  }
}

function buildDecorations(doc: PMNode, extras: DocxRenderExtras, totalPages: number | null, autoBreaks: readonly number[]): DecorationSet {
  const out: Decoration[] = [];
  const numbers: PageNumbers = { seq: [], total: 0 };
  const before = numberingRules.size;
  numberingDecorations(doc, extras.numbering, out, numberingRules);
  if (numberingRules.size !== before && typeof document !== 'undefined') {
    numberingStyle ??= document.head.appendChild(document.createElement('style'));
    numberingStyle.textContent = [...numberingRules.values()].join('\n');
  }
  pageDecorations(doc, extras, totalPages, autoBreaks, out, numbers);
  const set = DecorationSet.create(doc, out);
  pageNumbersOf.set(set, numbers);
  return set;
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

/** 쪽 모양을 바꾸는 노드 — 쪽 나누기·각주 참조·구역 나눔·문단 앞 쪽 나눔·다단 상자 */
const pageShaping = (n: PMNode) =>
  n.type.name === 'page_break' || n.type.name === 'columns' || (n.type.name === 'ooxml' && n.attrs.kind === 'note') || typeof n.attrs.sect === 'number' || Boolean(n.attrs.breakBefore);

/**
 * 쪽 장식은 그대로 두고(옮기기만) 번호만 다시 맞추면 되는 편집인가 — Enter·문단 붙여넣기·문단 지우기처럼 블록이 생기고 없어지지만
 * 쪽 나누기·구역·각주는 건드리지 않은 것(REVIEW-r3 I-r3-3: 예전엔 이때마다 장식 전체를 다시 만들어 800쪽 4배 느리게 Enter 300ms대).
 */
function isBlockEditKeepingPages(tr: Transaction): boolean {
  return tr.steps.every((step, i) => {
    if (!(step instanceof ReplaceStep || step instanceof ReplaceAroundStep)) return false;
    const { from, to, slice } = step as unknown as { from: number; to: number; slice: import('prosemirror-model').Slice };
    let ok = true;
    slice.content.descendants((n) => {
      if (pageShaping(n)) ok = false;
      return ok;
    });
    if (ok) tr.docs[i].nodesBetween(from, to, (n) => {
      if (pageShaping(n)) ok = false;
      return ok;
    });
    return ok;
  });
}

/** 장식 묶음의 번호 글자 장식만 새 문서 기준으로 바꾼다(다른 장식은 옮긴 그대로) — 바뀐 번호만 빼고 넣는다 */
function refreshNumbering(set: DecorationSet, doc: PMNode, defs: NumberingDefs | null): DecorationSet {
  const fresh: Decoration[] = [];
  numberingDecorations(doc, defs, fresh, numberingRules);
  const idOf = (d: Decoration) => `${d.from}|${(d.spec as { key?: string }).key}`;
  const old = set.find(undefined, undefined, (spec) => typeof spec.key === 'string' && spec.key.startsWith('num-'));
  const oldIds = new Set(old.map(idOf));
  const freshIds = new Set(fresh.map(idOf));
  const gone = old.filter((d) => !freshIds.has(idOf(d)));
  const added = fresh.filter((d) => !oldIds.has(idOf(d)));
  return gone.length === 0 && added.length === 0 ? set : set.remove(gone).add(doc, added);
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
function layoutBlocks(root: HTMLElement, blocks: readonly HTMLElement[], defaultTabPx: number, withLabels: boolean): void {
  if (blocks.length === 0) return;
  // 화면 배율(편집기 확대·축소 CSS zoom) — 잰 값(화면 px)을 문서 px로 되돌린다(편집 영역 전체 한 번)
  const scale = root.offsetWidth > 0 ? root.getBoundingClientRect().width / root.offsetWidth : 1;
  const own = (block: HTMLElement, selector: string) =>
    [...block.querySelectorAll<HTMLElement>(selector)].filter((t) => t.parentElement?.closest(BLOCK_SELECTOR) === block);
  // 번호 글자 폭은 번호와 들여쓰기에만 달렸다 — 문단 안 글자만 바뀐 때는 다시 재지 않는다(쓰고 읽기가 강제 다시 배치를 부른다)
  const work = blocks.map((block) => ({ block, labels: withLabels ? own(block, '.docx-num[data-tab]') : [], tabs: own(block, '.docx-tab'), stops: tabStopsOf(block) }));
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

/** 넘침 쪽 나눔의 재료 — 본문 직속 블록 하나(다단 구역은 통째로 하나) */
export interface FlowItem {
  /** 마지막으로 그려졌을 때의 높이·위아래 여백(px, ResizeObserver가 적어 둔 값). 아직 모르면 null */
  readonly height: number | null;
  readonly marginTop: number;
  readonly marginBottom: number;
  /** para: 줄 단위로 나눈다(줄 높이 lineHeight) · split: 아무 데서나(표의 행·다단 — 위치는 셀 수만 있다) · block: 나누지 않는다 */
  readonly kind: 'para' | 'split' | 'block';
  readonly lineHeight: number;
  /** 이 블록 안 각주 본문 높이 합(px) — 각주는 참조가 있는 쪽 아래에 놓이므로 그 쪽 본문이 그만큼 줄어든다 */
  readonly notes: number;
  /** 이 블록이 속한 구역(쪽 크기) */
  readonly section: number;
  /** 문단 앞 쪽 나눔 · 블록 뒤에서 쪽이 바뀜(쪽 나누기·다음 쪽 구역) */
  readonly breakBefore: boolean;
  readonly breakAfter: boolean;
}

/** 넘침 쪽 경계: index번째 블록의 위에서 offset(px)만큼 내려간 곳 — 0이면 블록 앞 */
export interface FlowBreak {
  readonly index: number;
  readonly offset: number;
}

/** 쪽에 남은 자리(room px)에 둘 수 있는 블록 조각의 높이 — 문단은 줄 단위로, 앞뒤 2줄씩은 남긴다(Word 기본 '외톨이 줄 방지') */
function fitPiece(item: FlowItem, room: number, rest: number): number {
  if (item.kind === 'block' || room <= 0) return 0;
  if (item.kind === 'split') return room >= 20 ? room : 0;
  const lh = item.lineHeight;
  if (!(lh > 0)) return 0;
  const lines = Math.round(rest / lh);
  const fit = Math.min(Math.floor((room + 1) / lh), lines - 2);
  return fit >= 2 ? fit * lh : 0;
}

/**
 * 블록 높이표로 쪽을 나눈다(그려진 위치를 읽지 않는다 — 화면 밖 블록은 그리지 않고 건너뛰므로).
 * bodyHeight(구역): 그 구역 쪽의 본문 높이(쪽 높이 − 위아래 여백). 각주가 있는 쪽은 각주 높이만큼 본문이 줄어든다.
 * 쪽보다 긴 블록은 나눠 다음 쪽에 이어 센다 — 문단은 줄 단위, 표·다단은 높이로(행 경계는 모른다).
 */
export function paginateFlow(items: readonly FlowItem[], bodyHeight: (section: number) => number | null): FlowBreak[] {
  const out: FlowBreak[] = [];
  let used: number | null = null; // 이 쪽 본문에 쓴 높이 — null이면 새 쪽 처음
  let notes = 0; // 이 쪽 각주 높이
  let prevMargin = 0;
  items.forEach((item, index) => {
    if (item.breakBefore && index > 0) {
      used = null;
      notes = 0;
    }
    const limit = bodyHeight(item.section);
    if (item.height !== null && limit) {
      let top = used === null ? 0 : used + Math.max(prevMargin, item.marginTop);
      let rest = item.height;
      let done = 0;
      let notesPlaced = false;
      if (used !== null && top + rest > limit - notes - item.notes + 1) {
        const piece = fitPiece(item, limit - notes - item.notes - top, rest);
        out.push({ index, offset: piece });
        if (piece > 0) notesPlaced = true; // 참조가 앞 조각에 있다고 보고 앞 쪽에 둔다
        done += piece;
        rest -= piece;
        top = 0;
        notes = 0;
      }
      if (!notesPlaced) notes += item.notes;
      // 한 쪽보다 긴 블록 — 넘친 만큼 다음 쪽에 잇는다
      for (let guard = 0; top + rest > limit - notes + 1 && guard < 10000; guard += 1) {
        const piece = fitPiece(item, limit - notes - top, rest);
        if (piece <= 0) break;
        out.push({ index, offset: done + piece });
        done += piece;
        rest -= piece;
        top = 0;
        notes = 0;
      }
      used = top + rest;
      prevMargin = item.marginBottom;
    }
    if (item.breakAfter) {
      used = null;
      notes = 0;
      prevMargin = 0;
    }
  });
  return out;
}

interface Measured {
  readonly height: number;
  readonly marginTop: number;
  readonly marginBottom: number;
  readonly lineHeight: number;
}

/** 블록 안 각주 참조 id·쪽 나누기 — 노드는 바뀌지 않으니 노드마다 한 번만 찾는다(쪽 나눔 재기마다 문서 전체를 훑지 않게) */
const blockFacts = new WeakMap<PMNode, { noteIds: string[]; hasBreak: boolean }>();
function blockFactsOf(node: PMNode): { noteIds: string[]; hasBreak: boolean } {
  let facts = blockFacts.get(node);
  if (!facts) {
    const noteIds: string[] = [];
    let hasBreak = false;
    node.descendants((child) => {
      if (child.type.name === 'page_break') hasBreak = true;
      if (child.type.name === 'ooxml' && child.attrs.kind === 'note' && child.attrs.noteType === 'footnote') noteIds.push(String(child.attrs.noteId));
      return true;
    });
    facts = { noteIds, hasBreak };
    blockFacts.set(node, facts);
  }
  return facts;
}

/** 쪽 경계·쪽 끝 장식(본문 블록이 아니다 — 높이표에 넣지 않는다) */
const isPageWidget = (dom: Element) => dom.matches('.docx-page-boundary, .docx-page-start, .docx-page-end');

/**
 * 넘침 쪽 나눔: 본문 블록 높이표(ResizeObserver가 그릴 때마다 적어 둔 값)로 쪽을 센다. 경계는 블록 앞에, 쪽보다 긴 문단은 넘치는 줄
 * 앞에(화면에 그려진 문단은 그 줄의 글자 위치, 건너뛴 문단은 글자 수 비율로 어림 — 화면에 들어오면 다시 잰다), 표·다단 안에서는
 * 경계를 그리지 않고 쪽 수만 센다(estimated: 어림한 문단 DOM — 화면에 들어오면 다시 잰다).
 */
function measureAutoBreaks(view: EditorView, extras: DocxRenderExtras, heights: WeakMap<Element, Measured>, noteHeight: (id: string) => number, estimated: Set<HTMLElement>): number[] {
  const items: FlowItem[] = [];
  const nodes: Array<{ node: PMNode; pos: number; dom: HTMLElement | null }> = [];
  let section = 0;
  // 본문 직속 블록(다단 구역은 그 안 블록) → DOM: view.nodeDOM(pos)는 앞에서부터 세어 찾아 블록 수 제곱이 된다(400쪽 4배 느리게 글자당 약 0.3초) — 한 번 훑어 표로
  const domOf = new Map<PMNode, HTMLElement>();
  const collect = (parent: Element) => {
    for (const child of parent.children) {
      const node = (child as HTMLElement & { pmViewDesc?: { node?: PMNode } }).pmViewDesc?.node;
      if (!node || !(child instanceof HTMLElement)) continue;
      domOf.set(node, child);
      if (node.type.name === 'columns') collect(child);
    }
  };
  collect(view.dom);
  const add = (node: PMNode, pos: number, columns: number) => {
    const dom = domOf.get(node) ?? view.nodeDOM(pos);
    const el = dom instanceof HTMLElement ? dom : null;
    const m = el ? heights.get(el) : undefined;
    const { noteIds, hasBreak } = blockFactsOf(node);
    let notes = 0;
    for (const id of noteIds) notes += noteHeight(id);
    const sect = node.attrs.sect;
    let breakAfter = hasBreak;
    const own = section;
    if (typeof sect === 'number') {
      section = sect + 1;
      if ((extras.sections[section]?.type ?? 'nextPage') !== 'continuous') breakAfter = true;
    }
    // 다단 안 블록: 단 수만큼 나란히 쌓이므로 세로로는 높이 ÷ 단 수만큼 차지한다(블록 경계에서만 나눈다 — 나눈 곳엔 단 전체에 걸친 쪽 경계)
    const kind = columns > 1 ? 'block' : node.type.name === 'table' ? 'split' : node.isTextblock && !hasBreak ? 'para' : 'block';
    items.push({
      height: m ? m.height / columns : null,
      marginTop: (m?.marginTop ?? 0) / columns,
      marginBottom: (m?.marginBottom ?? 0) / columns,
      kind,
      lineHeight: m?.lineHeight ?? 0,
      notes,
      section: own,
      breakBefore: Boolean(node.attrs.breakBefore),
      breakAfter,
    });
    nodes.push({ node, pos, dom: el });
  };
  view.state.doc.forEach((node, pos) => {
    if (node.type.name !== 'columns') return add(node, pos, 1);
    const count = Math.max(1, Number(node.attrs.count) || 1);
    node.forEach((child, offset) => add(child, pos + 1 + offset, count));
  });
  const flow = paginateFlow(items, (i) => {
    const page = extras.sections[i]?.page;
    return page ? page.heightPx - page.marginTopPx - page.marginBottomPx : null;
  });
  estimated.clear();
  const breaks: number[] = [];
  const inner = new Map<number, number>();
  for (const { index, offset } of flow) {
    const { node, pos, dom } = nodes[index];
    if (offset === 0) breaks.push(pos);
    else if (items[index].kind === 'para') breaks.push(lineStartAt(view, node, pos, dom, offset, items[index].height ?? 1, estimated));
    else {
      // 표 안: 그리지 않고 쪽 수만 센다(안쪽 위치 하나에 쪽 하나 — 행 경계는 모른다)
      const k = (inner.get(pos) ?? 0) + 1;
      inner.set(pos, k);
      if (k < node.nodeSize - 1) breaks.push(pos + k);
    }
  }
  return breaks.sort((a, b) => a - b);
}

/** 문단 위에서 offset(px) 내려간 줄의 첫 글자 위치 — 그려진 문단이면 그 자리의 글자, 건너뛴(화면에서 먼, content-visibility: hidden) 문단이면 글자 수 비율로 어림 */
function lineStartAt(view: EditorView, node: PMNode, pos: number, dom: HTMLElement | null, offset: number, height: number, estimated: Set<HTMLElement>): number {
  const end = pos + node.nodeSize - 1;
  if (dom && getComputedStyle(dom).contentVisibility !== 'hidden') {
    const rect = dom.getBoundingClientRect();
    const scale = dom.offsetHeight > 0 ? rect.height / dom.offsetHeight : 1;
    const hit = view.posAtCoords({ left: rect.left + 2, top: rect.top + (offset + 2) * scale });
    if (hit && hit.pos > pos + 1 && hit.pos < end) return hit.pos;
  }
  if (dom) estimated.add(dom);
  return Math.min(end - 1, Math.max(pos + 2, pos + 1 + Math.round((node.content.size * offset) / height)));
}

/** 각주 본문 높이(px) — 각주는 편집하지 않으니 처음 한 번 숨긴 곳에 그려 잰다(쪽 끝 각주 칸과 같은 글꼴·폭) */
function noteHeights(view: EditorView, extras: DocxRenderExtras): (id: string) => number {
  const cache = new Map<string, number>();
  let box: HTMLElement | null = null;
  return (id) => {
    const known = cache.get(id);
    if (known !== undefined) return known;
    const content = extras.notes.footnote[id];
    if (!content) return 0;
    if (!box) {
      const host = view.dom.parentElement ?? document.body;
      box = el('div', 'docx-footnotes');
      const style = getComputedStyle(view.dom);
      box.style.cssText = `position:absolute;visibility:hidden;left:-99999px;top:0;width:${contentWidthOf(extras.sections[0]) ?? view.dom.clientWidth}px;font:${style.font};line-height:${style.lineHeight}`;
      box.setAttribute('aria-hidden', 'true');
      host.append(box);
    }
    const body = noteBodyDOM(content, '0');
    box.append(body);
    const height = body.getBoundingClientRect().height + (cache.size === 0 ? 12 : 0); // 처음 각주엔 구분선·위 여백
    body.remove();
    cache.set(id, height);
    return height;
  };
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
    if (line) d.push(...lineSpacingCss({ line, lineRule: spacing?.lineRule ?? null }));
    else if (rule.kind === 'h') d.push('line-height: var(--doc-line-height, 1.3)');
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
/** 화면 밖 블록 그리기 건너뛰기(content-visibility) 켜기 — 첫 쪽 나눔을 잰 뒤(그때까지 모든 블록 높이를 한 번 적어 둔다) */
const CONTENT_SKIP_META = 'docxContentSkip';
/** 화면에서 멀면 그리기를 건너뛸 본문 직속 요소 — 문단·제목(farAwareView)과 쪽 경계 장식(ProseMirror가 장식 안 변화는 보지 않는다).
 * 표는 열 너비 조절 표 보기(editor.ts)가 속성 변화를 문서 변경으로 읽어 넣지 않는다. 떠 있는 그림이 든 문단은 가두면 둘러싸기가 바뀌어 뺀다. */
const SKIPPABLE = ':is(p, h1, h2, h3, h4, h5, h6, .docx-page-boundary):not(.docx-float-host)';

/** 문단·제목 보기 — 스키마 toDOM 그대로 그리고, 화면 밖 표시(data-far)를 바꾼 것은 문서 변경으로 읽지 않는다 */
function farAwareView(node: PMNode) {
  const { dom, contentDOM } = DOMSerializer.renderSpec(document, node.type.spec.toDOM!(node));
  return { dom, contentDOM, ignoreMutation: (m: { type: string; attributeName?: string | null }) => m.type === 'attributes' && m.attributeName === 'data-far' };
}

/**
 * 커서(선택 양 끝)가 든 본문 직속 문단·제목에 data-caret을 단다 — CSS가 그 블록은 data-far여도 그린다.
 * 그리지 않는(content-visibility: hidden) 블록 안에는 Chromium이 DOM 선택을 둘 수 없어 다른 곳으로 옮기고, 그러면 Ctrl+End처럼
 * 화면 밖으로 커서를 보낸 뒤 친 글자가 엉뚱한 자리(마지막 문단 맨 앞)에 들어가 첫 글자 뒤로 순서가 뒤집혔다. 장식은 ProseMirror가
 * DOM 선택을 쓰기 전에 붙으므로 건너뛰기 범위(updateSkip, 스크롤 뒤 다음 그림)를 기다리지 않는다.
 */
export const caretBlockPlugin = new Plugin({
  props: {
    decorations: ({ doc, selection }) => {
      const deco: Decoration[] = [];
      for (const $pos of selection.anchor === selection.head ? [selection.$head] : [selection.$anchor, selection.$head]) {
        const block = $pos.depth >= 1 ? $pos.node(1) : null;
        if (block && (block.type.name === 'paragraph' || block.type.name === 'heading')) deco.push(Decoration.node($pos.before(1), $pos.after(1), { 'data-caret': '' }));
      }
      return deco.length > 0 ? DecorationSet.create(doc, deco) : null;
    },
  },
});

/** 화면에 보이는 블록 앞뒤로 그려 둘 본문 직속 블록(쪽 장식 포함) 수 */
const SKIP_MARGIN = 60;

/**
 * ProseMirror는 화면을 고칠 때마다 본문 직속 블록마다 DecorationSet.forChild를 부르는데, 그 안에서 장식 목록을 처음부터 훑는다
 * (prosemirror-view 1.42 — 블록 수 × 장식 수). 번호 1,600개·쪽 경계 800개인 400쪽 문서에서 글자 하나에 0.1초 넘게 걸린 원인이다.
 * 같은 결과를 블록 위치 → 하위 장식 표(한 번 만들어 둠)로 바로 찾는다. 줄 안 장식(inline, 이 플러그인은 만들지 않음)이 있으면 원래 방법으로.
 * shortcut: DecorationSet 내부 필드(local·children)를 읽는다 — prosemirror-view를 올릴 때 이 필드 이름이 그대로인지 확인할 것.
 */
const DecorationSetCtor = DecorationSet as unknown as new (local: Decoration[], children: unknown[]) => DecorationSet;
class IndexedDecorationSet extends DecorationSetCtor {
  private index: Map<number, DecorationSet> | null = null;
  private readonly hasInline: boolean;

  constructor(from: DecorationSet) {
    const inner = from as unknown as { local: Decoration[]; children: unknown[] };
    super(inner.local, inner.children);
    this.hasInline = inner.local.some((d) => (d as unknown as { inline: boolean }).inline);
  }

  override forChild(offset: number, node: PMNode): DecorationSet {
    if (this.hasInline) return super.forChild(offset, node) as DecorationSet;
    if (node.isLeaf) return DecorationSet.empty;
    if (!this.index) {
      this.index = new Map();
      const children = (this as unknown as { children: unknown[] }).children;
      for (let i = 0; i < children.length; i += 3) this.index.set(children[i] as number, children[i + 2] as DecorationSet);
    }
    return this.index.get(offset) ?? DecorationSet.empty;
  }
}
const indexed = new WeakMap<DecorationSet, DecorationSet>();
function indexedOf(set: DecorationSet): DecorationSet {
  let fast = indexed.get(set);
  if (!fast) {
    fast = new IndexedDecorationSet(set);
    indexed.set(set, fast);
  }
  return fast;
}

/** 편집기에 Word 문서 모양 장식을 붙이는 플러그인 — 번호·쪽 경계·구역 폭 장식, 탭 폭, 넘침 쪽 나눔(성능은 파일 위 설명). */
export function docxLayoutPlugin(extras: DocxRenderExtras, totalPages: number | null): Plugin<LayoutState> {
  const defaultTabPx = extras.defaultTabStopDxa / TWIPS_PER_PX;
  /** 화면 밖 블록 그리기 건너뛰기(CSS .docx-skip — content-visibility: hidden)가 켜졌나 */
  let skipOn = false;
  // 구역마다 본문 폭이 다르면 블록마다 폭 장식이 붙는다(pageDecorations)
  const widths = extras.sections.map((sec) => contentWidthOf(sec) ?? 0);
  const narrowSections = widths.some((w) => w > 0 && w < Math.max(0, ...widths) - 0.5);
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
        // 블록이 생기고 없어지기만 한 편집: 쪽 장식은 옮기고 번호만 맞춘다(구역마다 본문 폭이 다른 문서는 새 블록 폭 장식이 필요해 전체)
        if (!narrowSections && isBlockEditKeepingPages(tr)) return { deco: refreshNumbering(prev.deco.map(tr.mapping, tr.doc), state.doc, extras.numbering), autoBreaks, dirty: changedRanges(tr) };
        return { deco: buildDecorations(state.doc, extras, totalPages, autoBreaks), autoBreaks, dirty: 'all' };
      },
    },
    props: {
      decorations: (state) => {
        const deco = layoutKey.getState(state)?.deco;
        return deco ? indexedOf(deco) : null;
      },
      nodeViews: { paragraph: farAwareView, heading: farAwareView },
      attributes: (): Record<string, string> => (extras.sections.length > 0 ? { class: skipOn ? 'docx-paged docx-skip' : 'docx-paged' } : {}),
    },
    view: (view: EditorView) => {
      let frame = 0;
      let paginateTimer = 0;
      /**
       * 탭·번호 폭은 화면에 들어온(들어올) 문단만 잰다 — 400쪽 문서 전체를 재면 브라우저가 문서 전체를 다시 배치한다.
       * stale: 다시 재야 하는 문단, visible: 지금 화면 근처(IntersectionObserver)에 있는 문단.
       */
      const stale = new Set<HTMLElement>();
      /** 번호 글자 폭까지 다시 잴 문단(번호·들여쓰기가 바뀌었을 수 있는 'all') — 글자만 친 문단은 탭만 */
      const staleLabels = new Set<HTMLElement>();
      /** 고친 문단(커서가 있는 화면 안) — 바로 잰다 */
      const touchedNow = new Set<HTMLElement>();
      /** 화면 밖 어림으로 나눈 긴 문단 — 화면에 들어오면 그 자리 글자로 다시 잰다 */
      const estimated = new Set<HTMLElement>();
      const remeasured = new WeakSet<HTMLElement>();
      let firstPass = true;
      /**
       * 지금 화면에 보이는 문서 범위(글자 위치) — 화면 위·아래 끝 두 점의 글자 위치. IntersectionObserver로 수천 문단을 지켜보면
       * 그림마다 수 ms가 들어(400쪽 약 8ms/글자) 이렇게 한다. 화면 밖 건너뛴(content-visibility) 블록은 재지 않는다.
       */
      let scroller: Element | null | undefined;
      const scrollerOf = (): Element | null => {
        if (scroller === undefined) {
          scroller = null;
          for (let p = view.dom.parentElement; p; p = p.parentElement)
            if (/(auto|scroll)/.test(getComputedStyle(p).overflowY)) {
              scroller = p;
              break;
            }
        }
        return scroller;
      };
      const visibleRange = (): [number, number] | null => {
        const rect = view.dom.getBoundingClientRect();
        const clip = scrollerOf()?.getBoundingClientRect();
        const top = Math.max(rect.top, clip?.top ?? 0, 0) + 1;
        const bottom = Math.min(rect.bottom, clip?.bottom ?? Infinity, window.innerHeight) - 1;
        if (bottom <= top) return null;
        // posAtCoords는 큰 문서에서 본문 블록 전부의 글자 상자를 훑는다(400쪽 한 번 약 25ms) — 그 점의 요소만 집어 본문 직속 블록 위치로
        const left = rect.left + rect.width / 2;
        const from = topBlockPosAt(left, top, 'next');
        const to = topBlockPosAt(left, bottom, 'prev');
        // 못 집으면(화면 위에 다른 것이 덮였다) 고친 문단만 — 문서 전체를 재지 않는다
        if (from === null || to === null) return null;
        return [from, to + (view.state.doc.nodeAt(to)?.nodeSize ?? 0)];
      };
      /** 화면 점(x, y)에 있는 본문 직속 블록의 위치 — 쪽 장식·여백이면 dir 쪽 이웃 블록 */
      const topBlockPosAt = (x: number, y: number, dir: 'next' | 'prev'): number | null => {
        let node: Element | null = document.elementFromPoint(x, y);
        while (node && node.parentElement !== view.dom) node = node.parentElement;
        for (let guard = 0; node && guard < 50; guard += 1) {
          if (!isPageWidget(node)) {
            try {
              const pos = view.posAtDOM(node, 0) - 1;
              if (pos >= 0 && view.nodeDOM(pos) === node) return pos;
            } catch {
              // 블록이 아닌 요소
            }
          }
          node = dir === 'next' ? node.nextElementSibling : node.previousElementSibling;
        }
        return null;
      };
      /** 보이는 범위 안 블록(본문 문단·표 안 문단)의 DOM — 범위 안 블록 수만큼만 돈다 */
      const blocksIn = (range: [number, number] | null): Set<Element> => {
        const out = new Set<Element>();
        if (!range) return out;
        const doc = view.state.doc;
        doc.nodesBetween(Math.max(0, range[0]), Math.min(doc.content.size, range[1]), (node, pos) => {
          if (!node.isTextblock) return true;
          const dom = view.nodeDOM(pos);
          if (dom instanceof Element) out.add(dom);
          return false;
        });
        return out;
      };
      /** 쪽 장식(머리글·바닥글) 안 문단이 화면 안인가 — 장식은 건너뛰지 않으니 위치를 읽어도 된다 */
      const widgetVisible = (b: HTMLElement) => {
        const r = b.getBoundingClientRect();
        return r.bottom >= 0 && r.top <= window.innerHeight;
      };
      const flush = () => {
        frame = 0;
        const range = firstPass ? null : visibleRange();
        const shown = blocksIn(range);
        const targets = [...stale].filter((b) => b.isConnected && (firstPass || touchedNow.has(b) || shown.has(b) || (b.closest('.docx-page-boundary, .docx-page-start, .docx-page-end') !== null && widgetVisible(b))));
        touchedNow.clear();
        for (const b of targets) stale.delete(b);
        const labelled = targets.filter((b) => staleLabels.delete(b));
        layoutBlocks(view.dom, labelled, defaultTabPx, true);
        layoutBlocks(view.dom, targets.filter((b) => !labelled.includes(b)), defaultTabPx, false);
        firstPass = false;
        // 어림으로 나눈 문단이 화면에 들어왔다 — 그 문단마다 한 번만 다시 잰다(재도 또 어림이면 되풀이하지 않게)
        const arrived = [...estimated].filter((b) => shown.has(b) && !remeasured.has(b));
        for (const b of arrived) remeasured.add(b);
        if (arrived.length > 0) paginate();
      };
      const schedule = (what: 'all' | HTMLElement[]) => {
        if (what === 'all')
          for (const b of blocksWithTabs(view.dom)) {
            stale.add(b);
            staleLabels.add(b);
          }
        else
          for (const b of what) {
            stale.add(b);
            touchedNow.add(b);
          }
        if (!frame) frame = requestAnimationFrame(flush);
      };
      /**
       * 화면 밖 블록 그리기 건너뛰기: 화면 근처(앞뒤 SKIP_MARGIN개) 밖의 본문 직속 블록은 content-visibility: hidden(data-far 표시).
       * content-visibility: auto는 브라우저가 블록마다 화면 안인지 그림마다 재어(400쪽 2400블록,
       * 글자당 약 20ms) 쓰지 않는다. 건너뛴 블록 크기는 마지막으로 그렸을 때 크기(contain-intrinsic-size: auto)라 스크롤 위치는 그대로다.
       */
      let shownRange: [number, number] | null = null;
      const updateSkip = () => {
        if (!skipOn) return;
        const children = view.dom.children;
        const rect = view.dom.getBoundingClientRect();
        const clip = scrollerOf()?.getBoundingClientRect();
        const top = Math.max(rect.top, clip?.top ?? 0, 0) + 1;
        const bottom = Math.min(rect.bottom, clip?.bottom ?? Infinity, window.innerHeight) - 1;
        const x = rect.left + rect.width / 2;
        // 블록 사이 틈·여백(편집 영역 자체)에 걸리면 조금씩 옮겨 다시 집는다
        const indexAt = (y: number, step: number) => {
          for (let k = 0; k < 12; k += 1) {
            let node: Element | null = document.elementFromPoint(x, y + k * step);
            while (node && node.parentElement !== view.dom) node = node.parentElement;
            if (node) return Array.prototype.indexOf.call(children, node) as number;
          }
          return -1;
        };
        const first = bottom > top ? indexAt(top, 24) : -1;
        const last = bottom > top ? indexAt(bottom, -24) : -1;
        if (first < 0 || last < 0) return; // 못 집으면 그대로(다른 것이 화면을 덮었다)
        if (shownRange && first - shownRange[0] >= SKIP_MARGIN / 4 && shownRange[1] - last >= SKIP_MARGIN / 4) return; // 아직 여유 — 규칙을 바꾸지 않는다
        const lo = Math.max(0, first - SKIP_MARGIN);
        const hi = Math.min(children.length - 1, last + SKIP_MARGIN);
        shownRange = [lo, hi];
        // 블록마다 data-far 표시(바뀐 것만 쓴다) — 순서 번호 선택자(nth-child)는 Enter로 블록 하나가 끼면 뒤 블록 전부의 스타일을 다시
        // 계산하게 해(800쪽 4배 느리게 Enter마다 약 250ms) 쓰지 않는다. 문단·제목은 farAwareView가 이 속성 변화를 ProseMirror에 숨긴다.
        for (let i = 0; i < children.length; i += 1) {
          const child = children[i];
          const want = (i < lo || i > hi) && child.matches(SKIPPABLE);
          if (want !== child.hasAttribute('data-far')) {
            if (want) child.setAttribute('data-far', '');
            else child.removeAttribute('data-far');
          }
        }
      };
      let skipFrame = 0;
      // 굴리면 건너뛰기 범위를 옮기고, 새로 들어온 문단의 탭 폭을 잰다(다시 잴 것이 있을 때만)
      let lastSkipY = Number.NaN;
      const onScroll = () => {
        const y = scrollerOf()?.scrollTop ?? window.scrollY;
        const moved = !(Math.abs(y - lastSkipY) < window.innerHeight / 3);
        if (moved) lastSkipY = y;
        if (skipOn && moved && !skipFrame)
          skipFrame = requestAnimationFrame(() => {
            skipFrame = 0;
            updateSkip();
          });
        if ((stale.size > 0 || estimated.size > 0) && !frame) frame = requestAnimationFrame(flush);
      };
      window.addEventListener('scroll', onScroll, { capture: true, passive: true });
      // 본문 블록 높이표 — 블록이 그려질 때마다(크기가 바뀔 때만) 적는다. 쪽 나눔은 이 표로만 센다(화면 밖 블록 위치를 읽지 않는다)
      const heights = new WeakMap<Element, Measured>();
      const ro = typeof ResizeObserver === 'function'
        ? new ResizeObserver((entries) => {
          for (const e of entries) {
            const target = e.target as HTMLElement;
            if (!target.isConnected) continue;
            const style = getComputedStyle(target);
            const fontSize = parseFloat(style.fontSize) || 16;
            // 긴 문단 안에 둔 쪽 경계(inline 장식)는 본문이 아니다 — 빼야 다음 재기에서 문단이 그만큼 길어 보이지 않는다
            let widgets = 0;
            for (const w of target.querySelectorAll<HTMLElement>(':scope > .docx-page-boundary')) widgets += w.offsetHeight;
            heights.set(target, {
              height: target.offsetHeight - widgets,
              marginTop: parseFloat(style.marginTop) || 0,
              marginBottom: parseFloat(style.marginBottom) || 0,
              lineHeight: parseFloat(style.lineHeight) || fontSize * 1.3,
            });
          }
          paginate();
        })
        : null;
      const observeBlocks = (list: Iterable<Node>) => {
        for (const child of list) {
          if (!(child instanceof HTMLElement) || isPageWidget(child)) continue;
          if (child.classList.contains('docx-columns')) {
            // 다단 상자는 그 안 블록을 잰다(쪽 나눔이 단 안 블록 경계에서 나눈다)
            mo.observe(child, { childList: true });
            observeBlocks(child.children);
          } else ro?.observe(child);
        }
      };
      const mo = new MutationObserver((records) => {
        for (const r of records) {
          observeBlocks(r.addedNodes);
          for (const gone of r.removedNodes) if (gone instanceof HTMLElement) ro?.unobserve(gone);
        }
      });
      mo.observe(view.dom, { childList: true });
      observeBlocks(view.dom.children);
      const noteHeight = noteHeights(view, extras);
      const paginate = () => {
        window.clearTimeout(paginateTimer);
        paginateTimer = window.setTimeout(() => {
          const breaks = measureAutoBreaks(view, extras, heights, noteHeight, estimated);
          const current = layoutKey.getState(view.state)?.autoBreaks ?? [];
          // 첫 쪽 나눔을 잰 뒤부터 화면 밖 블록 그리기를 건너뛴다(그 전에 한 번은 모두 그려 높이표를 채운다)
          const tr = view.state.tr;
          if (!skipOn && view.dom.classList.contains('docx-paged')) {
            skipOn = true;
            tr.setMeta(CONTENT_SKIP_META, true);
          }
          if (breaks.join() !== current.join()) tr.setMeta(PAGINATION_META, breaks);
          if (tr.getMeta(PAGINATION_META) || tr.getMeta(CONTENT_SKIP_META)) view.dispatch(tr.setMeta('addToHistory', false));
          shownRange = null; // 쪽 경계가 바뀌면 순서 번호도 바뀐다
          updateSkip();
        }, 800);
      };
      schedule('all');
      // 글꼴이 늦게 내려오면 글자 폭이 바뀐다 — 다 내려온 뒤 한 번 더 잰다(높이는 ResizeObserver가 다시 적는다)
      void document.fonts?.ready.then(() => {
        if (!skipOn) firstPass = true;
        schedule('all');
      });
      return {
        update: (v, prevState) => {
          if (v.state.doc === prevState.doc && layoutKey.getState(v.state)?.deco === layoutKey.getState(prevState)?.deco) return;
          const dirty = layoutKey.getState(v.state)?.dirty ?? 'all';
          if (dirty === 'all') {
            renumberPages(v.dom, pageNumbersOf.get(layoutKey.getState(v.state)!.deco));
            schedule('all');
          }
          else {
            const touched: HTMLElement[] = [];
            for (const [from, to] of dirty) {
              v.state.doc.nodesBetween(Math.max(0, from), Math.min(v.state.doc.content.size, to), (node, pos) => {
                if (!node.isTextblock) return true;
                const dom = v.nodeDOM(pos);
                if (dom instanceof HTMLElement && dom.querySelector('.docx-tab')) touched.push(dom);
                return false;
              });
            }
            if (touched.length > 0) schedule(touched);
          }
          // 글자만 친 편집은 높이가 바뀔 때 ResizeObserver가 부른다 — 글자마다 쪽 나눔을 다시 재지 않는다
          if (v.state.doc !== prevState.doc && dirty === 'all') paginate();
        },
        destroy: () => {
          cancelAnimationFrame(frame);
          window.clearTimeout(paginateTimer);
          window.removeEventListener('scroll', onScroll, { capture: true });
          cancelAnimationFrame(skipFrame);
          ro?.disconnect();
          mo.disconnect();
        },
      };
    },
  });
}
