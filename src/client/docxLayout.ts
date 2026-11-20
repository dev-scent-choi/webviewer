import { Plugin, PluginKey, type EditorState } from 'prosemirror-state';
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
 * 쪽은 명시적 나눔에서만 바뀐다 — 내용이 넘쳐서 생기는 자동 쪽 나눔은 계산하지 않는다(그건 보기 화면이 Word와 같다).
 */

const TWIPS_PER_PX = 15;
const layoutKey = new PluginKey<DecorationSet>('docxLayout');

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

function headerFooterOf(section: DocxSectionInfo | undefined, which: 'header' | 'footer', firstOfSection: boolean): unknown[] | null {
  if (!section) return null;
  const set = section[which];
  return section.titlePg && firstOfSection ? set.first : set.default;
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
  end.append(headerFooterDOM(headerFooterOf(extras.sections[cur.section], 'footer', cur.firstOfSection), 'docx-footer', cur.page, total));
  return end;
}

function pageStartDOM(extras: DocxRenderExtras, cur: PageCursor, total: number, first: boolean): HTMLElement {
  const start = el('div', `docx-page-start${first ? ' docx-page-start-first' : ''}`);
  const section = extras.sections[cur.section];
  if (section) start.style.setProperty('--hdr-top', `${section.headerPx}px`);
  start.append(headerFooterDOM(headerFooterOf(section, 'header', cur.firstOfSection), 'docx-header', cur.page, total));
  return start;
}

/** 번호 글자 + 번호 수준의 들여쓰기(문단에 직접 들여쓰기가 없을 때) */
function numberingDecorations(doc: PMNode, defs: NumberingDefs | null, out: Decoration[]): void {
  if (!defs) return;
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

/** 쪽 경계 장식(머리글·바닥글·각주) — 본문 직속 블록(다단 구역 안 블록 포함)을 문서 순서대로 훑는다 */
function pageDecorations(doc: PMNode, extras: DocxRenderExtras, totalPages: number | null, out: Decoration[]): void {
  if (extras.sections.length === 0) return;
  // after: 이 블록 뒤 쪽 경계를 둘 자리 — 다단 구역의 마지막 블록이면 다단 상자 밖(안에 두면 머리글까지 단으로 흐른다)
  const blocks: Array<{ node: PMNode; pos: number; after: number }> = [];
  doc.forEach((node, offset) => {
    if (node.type.name === 'columns') {
      node.forEach((child, inner, index) => {
        const pos = offset + 1 + inner;
        blocks.push({ node: child, pos, after: index === node.childCount - 1 ? offset + node.nodeSize : pos + child.nodeSize });
      });
    } else blocks.push({ node, pos: offset, after: offset + node.nodeSize });
  });
  // 쪽 수: 보기 화면(LibreOffice)이 센 실제 쪽 수가 있으면 그것, 없으면 명시적 나눔으로 센 수
  let counted = 1;
  const boundaries: Array<{ pos: number; nextSection: number | null }> = [];
  let section = 0;
  for (const { node, pos, after } of blocks) {
    if (node.attrs.breakBefore && pos > 0) boundaries.push({ pos, nextSection: null });
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

function buildDecorations(state: EditorState, extras: DocxRenderExtras, totalPages: number | null): DecorationSet {
  const out: Decoration[] = [];
  numberingDecorations(state.doc, extras.numbering, out);
  pageDecorations(state.doc, extras, totalPages, out);
  return DecorationSet.create(state.doc, out);
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

/** 문단마다 탭 폭을 정한다 — 앞 탭 폭이 뒤 탭 위치를 바꾸므로 앞에서부터 차례로 잰다 */
function layoutTabs(root: HTMLElement, defaultTabPx: number): void {
  // 화면 배율(편집기 확대·축소 CSS zoom) — 잰 값(화면 px)을 문서 px로 되돌린다. 위치를 잡은(absolute) 머리글 안에서는
  // 요소마다 offsetWidth가 배율을 따르기도 해서, 편집 영역 전체 한 번으로 정한다.
  const scale = root.offsetWidth > 0 ? root.getBoundingClientRect().width / root.offsetWidth : 1;
  // 번호 글자가 내어쓰기 폭보다 길면(1.1.1.) Word는 번호 뒤 탭을 다음 기본 탭 위치로 보낸다 — 붙어 보이지 않게
  for (const label of root.querySelectorAll<HTMLElement>('.docx-num[data-tab]')) {
    const minPx = Number(label.dataset.tab);
    label.style.width = '';
    const natural = label.scrollWidth;
    if (natural > minPx + 0.5) {
      const block = label.closest<HTMLElement>('p, h1, h2, h3, h4, h5, h6');
      const origin = block ? block.getBoundingClientRect().left - (parseFloat(getComputedStyle(block).marginLeft) || 0) * scale : 0;
      const x = (label.getBoundingClientRect().left - origin) / scale;
      label.style.width = `${Math.ceil((x + natural + 1) / defaultTabPx) * defaultTabPx - x}px`;
    }
  }
  const blocks = new Set<HTMLElement>();
  for (const tab of root.querySelectorAll<HTMLElement>('.docx-tab')) {
    const block = tab.parentElement?.closest<HTMLElement>('p, h1, h2, h3, h4, h5, h6');
    if (block) blocks.add(block);
  }
  for (const block of blocks) {
    const stops = (() => {
      try {
        return JSON.parse(block.dataset.tabs ?? '[]') as TabStop[];
      } catch {
        return [];
      }
    })();
    const tabs = [...block.querySelectorAll<HTMLElement>('.docx-tab')].filter((t) => t.parentElement?.closest('p, h1, h2, h3, h4, h5, h6') === block);
    for (const tab of tabs) {
      tab.style.width = '0px';
      tab.className = 'docx-tab';
    }
    const style = getComputedStyle(block);
    const rect = block.getBoundingClientRect();
    const origin = rect.left - (parseFloat(style.marginLeft) || 0) * scale;
    for (const tab of tabs) {
      const x = (tab.getBoundingClientRect().left - origin) / scale;
      const stop = stops.find((s) => s.pos / TWIPS_PER_PX > x + 1.5); // 1.5: 앞 탭을 1px 덜 채웠으니 같은 정지에 다시 걸리지 않게
      let width: number;
      if (!stop) width = Math.ceil((x + 1) / defaultTabPx) * defaultTabPx - x;
      else {
        const pos = stop.pos / TWIPS_PER_PX;
        const val = stop.val === 'end' ? 'right' : stop.val === 'start' ? 'left' : stop.val;
        const follow = val === 'right' || val === 'center' || val === 'decimal' ? followingWidth(tab, block, val === 'decimal') / scale : 0;
        width = pos - x - (val === 'center' ? follow / 2 : follow);
        if (stop.leader) tab.classList.add(`docx-tab-${stop.leader === 'hyphen' || stop.leader === 'underscore' || stop.leader === 'heavy' ? 'line' : 'dot'}`);
      }
      // 1px 덜 — 오른쪽 끝에 딱 맞춘 탭(목차 쪽 번호)이 소수점 차이로 다음 줄로 넘어가지 않게
      tab.style.width = `${Math.max(2, width - 1)}px`;
    }
  }
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

/** 문서 스타일 → <style> 한 장. 제목은 h1~h6, 문단 스타일은 p[data-style], 문자 스타일은 span[data-rstyle]. */
export function docxStyleSheetText(rules: readonly DocxStyleRule[]): string {
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
    .filter(Boolean)
    .join('\n');
}

/** 편집기에 Word 문서 모양 장식을 붙이는 플러그인 — 문서가 바뀔 때마다 번호·쪽 경계를 다시 계산하고, 그린 뒤 탭 폭을 잰다. */
export function docxLayoutPlugin(extras: DocxRenderExtras, totalPages: number | null): Plugin<DecorationSet> {
  const defaultTabPx = extras.defaultTabStopDxa / TWIPS_PER_PX;
  return new Plugin<DecorationSet>({
    key: layoutKey,
    state: {
      init: (_, state) => buildDecorations(state, extras, totalPages),
      apply: (tr, prev, _, state) => (tr.docChanged ? buildDecorations(state, extras, totalPages) : prev),
    },
    props: {
      decorations: (state) => layoutKey.getState(state),
    },
    view: (view: EditorView) => {
      let frame = 0;
      const schedule = () => {
        cancelAnimationFrame(frame);
        frame = requestAnimationFrame(() => layoutTabs(view.dom, defaultTabPx));
      };
      schedule();
      // 글꼴이 늦게 내려오면 글자 폭이 바뀐다 — 다 내려온 뒤 한 번 더 잰다
      void document.fonts?.ready.then(schedule);
      return {
        update: (v, prevState) => {
          if (v.state.doc !== prevState.doc) schedule();
        },
        destroy: () => cancelAnimationFrame(frame),
      };
    },
  });
}
