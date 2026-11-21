import { strFromU8 } from 'fflate';
import { guardedUnzipSync } from '../zipGuard.js';
import { type Mark, Node as PMNode, type NodeType } from 'prosemirror-model';
import { IMAGE_MIME_BY_EXT } from '../pptx/slideContent.js';
import {
  docxSchema,
  type CellBorderSide,
  type CellBorders,
  type CellPadding,
  type PageSetup,
  type ParagraphIndent,
  type ParagraphSpacing,
  type DocxComment,
  type RowHeight,
  type TabStop,
  GALLERY_STYLE_NAMES,
  HIGHLIGHT_COLORS,
} from './schema.js';
import {
  attrsOf,
  childrenOf,
  deepTextOf,
  extrasOf,
  findAll,
  fragmentHash,
  fragmentRef,
  refHashOf,
  TRACKED_WRAP_TAGS,
  trackedWrapKey,
  type FragmentStore,
  findDeep,
  findFirst,
  onOff,
  parser,
  tagOf,
  textOf,
  xmlOf,
  type XmlNode,
} from './ooxml.js';
import { parseNumberingXml, type NumberingDefs } from './numbering.js';

/** 문서 파트 하나(본문·머리글·각주 등)를 읽는 동안 함께 쓰는 값 */
interface ImportCtx {
  readonly styles: readonly XmlNode[];
  readonly defaultAlign: string | null;
  readonly styleTableCellMargin: CellMarginSides | null;
  /** 이 파트의 관계(r:id → 대상) — 그림·하이퍼링크 주소 */
  readonly rels: ReadonlyMap<string, { readonly target: string; readonly external: boolean }>;
  readonly contentWidthPx: number | null;
  readonly themeXml: string | undefined;
  /** 각주·미주 번호(문서에 나온 순서) */
  readonly noteCount: { footnote: number; endnote: number };
  /** 각주·미주 본문을 읽는 중이면 그 번호(w:footnoteRef가 보여 줄 글자) */
  readonly noteLabel: string | null;
  /** 보존 조각 저장소(해시 → 원본 XML) — ooxml.ts fragmentRef */
  readonly fragments: FragmentStore;
  /** 브라우저가 못 그려 서버가 그림으로 바꿔 보여 줄 조각(EMF/WMF·차트·도형)의 해시 — drawingImage.ts가 한 번에 그린다 */
  readonly serverDrawn: Set<string>;
}

/** 표 스타일이 칸 안 글자에 물려주는 값(굵게·색·글꼴)과 문단 간격 */
interface RunDefaults {
  readonly strong?: boolean;
  readonly em?: boolean;
  readonly color?: string | null;
  readonly font?: { readonly ascii: string | null; readonly eastAsia: string | null } | null;
}
interface CellDefaults {
  readonly run: RunDefaults;
  readonly spacing: ParagraphSpacing | null;
}
const NO_CELL_DEFAULTS: CellDefaults = { run: {}, spacing: null };

const HEX6 = /^[0-9A-Fa-f]{6}$/;
const hexColor = (val: string | undefined): string | null => (val && HEX6.test(val) ? `#${val.toLowerCase()}` : null);

// ── 스타일 ───────────────────────────────────────────────────────────────

// 문단마다 스타일 목록을 다시 훑지 않게 styles 배열 인스턴스별로 "종류:id → 노드" Map을 한 번만 만든다
const styleMapCache = new WeakMap<readonly XmlNode[], ReadonlyMap<string, XmlNode>>();

function styleMapOf(styles: readonly XmlNode[]): ReadonlyMap<string, XmlNode> {
  const cached = styleMapCache.get(styles);
  if (cached) return cached;
  const map = new Map<string, XmlNode>();
  for (const s of findAll(styles, 'w:style')) {
    const attrs = attrsOf(s);
    const key = `${attrs['@_w:type']}:${attrs['@_w:styleId']}`;
    if (!map.has(key)) map.set(key, s); // 같은 id가 또 있으면 먼저 나온 것(Word와 같다)
  }
  styleMapCache.set(styles, map);
  return map;
}

function styleOf(styles: readonly XmlNode[], type: string, id: string): XmlNode | undefined {
  return styleMapOf(styles).get(`${type}:${id}`);
}

/** 스타일과 그 기반(w:basedOn) 사슬 — 기반이 앞, 파생이 뒤(뒤 것이 앞 것을 덮는다). */
function styleChainOf(styles: readonly XmlNode[], type: string, id: string | undefined): XmlNode[] {
  const chain: XmlNode[] = [];
  let current = id ? styleOf(styles, type, id) : undefined;
  while (current && chain.length < 12 && !chain.includes(current)) {
    chain.unshift(current);
    const basedOn = findFirst(childrenOf(current), 'w:basedOn');
    current = basedOn ? styleOf(styles, type, attrsOf(basedOn)['@_w:val'] ?? '') : undefined;
  }
  return chain;
}

/** 사슬의 각 스타일에서 w:pPr(또는 w:rPr) 자식 목록 — 기반부터 */
function chainPr(chain: readonly XmlNode[], prTag: string): XmlNode[][] {
  return chain.map((s) => {
    const pr = findFirst(childrenOf(s), prTag);
    return pr ? childrenOf(pr) : [];
  });
}

/** 여러 단계의 자식 목록 중 tag가 마지막으로 나온 요소(파생이 이긴다) */
function lastOf(levels: readonly (readonly XmlNode[])[], tag: string): XmlNode | undefined {
  for (let i = levels.length - 1; i >= 0; i -= 1) {
    const el = findFirst(levels[i], tag);
    if (el) return el;
  }
  return undefined;
}

/** w:pStyle 값 → 제목 수준. 스타일 이름이 "heading N"이면(한국어 Word는 id가 "1"이다) 또는 id가 HeadingN이면. */
function headingLevelOf(styles: readonly XmlNode[], styleId: string | undefined): number | null {
  if (!styleId) return null;
  const byId = /^Heading([1-6])$/i.exec(styleId);
  if (byId) return Number(byId[1]);
  const style = styleOf(styles, 'paragraph', styleId);
  const name = style ? attrsOf(findFirst(childrenOf(style), 'w:name') ?? {})['@_w:val'] ?? '' : '';
  const byName = /^heading ([1-6])$/i.exec(name);
  return byName ? Number(byName[1]) : null;
}

const JC_TO_ALIGN: Record<string, string> = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' };

/** w:jc val="center"/"right"/"both" → ProseMirror align 속성. left/미지정은 null(기본값)로 둔다. */
function alignOf(pPrChildren: readonly XmlNode[]): string | null {
  const jc = findFirst(pPrChildren, 'w:jc');
  if (!jc) return null;
  const val = attrsOf(jc)['@_w:val'] ?? '';
  return JC_TO_ALIGN[val] ?? null;
}

interface NamedStyleInfo {
  readonly align: string | null;
  readonly bold: boolean;
}

/**
 * 문단이 참조하는 스타일(w:pStyle)이 직접 지정한 정렬(jc)·굵기(rPr/b) — 문단에 직접 얹는다(표 제목 칸 등이 직접 서식
 * 없이 이런 스타일만 참조하는 실제 사례). 스타일의 나머지 모양(색·크기·글꼴·간격, 기반 스타일 포함)은 편집기가 스타일 CSS로
 * 입힌다(docxStyleSheetOf). 문단 자신의 직접 서식이 여전히 이 값보다 우선한다.
 */
function namedStyleInfoOf(styleId: string, styles: readonly XmlNode[]): NamedStyleInfo | null {
  const style = styleOf(styles, 'paragraph', styleId);
  if (!style) return null;
  const styleChildren = childrenOf(style);
  const pPr = findFirst(styleChildren, 'w:pPr');
  const rPr = findFirst(styleChildren, 'w:rPr');
  return {
    align: alignOf(pPr ? childrenOf(pPr) : []),
    bold: rPr ? onOff(childrenOf(rPr), 'w:b') === true : false,
  };
}

/** 문단 스타일 사슬과 문단 자신의 w:pPr에서 번호·쪽 나눔·탭 정지를 합친다(파생·직접이 이긴다). */
function paragraphLayoutOf(pPrChildren: readonly XmlNode[], styles: readonly XmlNode[], styleId: string | undefined) {
  const levels = [...chainPr(styleChainOf(styles, 'paragraph', styleId), 'w:pPr'), pPrChildren];
  let numId: string | null = null;
  let ilvl: number | null = null;
  let fromStyle = true;
  const tabs = new Map<number, TabStop>();
  let breakBefore = false;
  for (const level of levels) {
    const numPr = findFirst(level, 'w:numPr');
    if (numPr) {
      const id = attrsOf(findFirst(childrenOf(numPr), 'w:numId') ?? {})['@_w:val'];
      const lv = attrsOf(findFirst(childrenOf(numPr), 'w:ilvl') ?? {})['@_w:val'];
      if (id !== undefined) numId = id;
      if (lv !== undefined) ilvl = Number(lv);
      fromStyle = level !== pPrChildren;
    }
    const pbb = onOff(level, 'w:pageBreakBefore');
    if (pbb !== undefined) breakBefore = pbb;
    const tabsEl = findFirst(level, 'w:tabs');
    for (const tab of tabsEl ? findAll(childrenOf(tabsEl), 'w:tab') : []) {
      const a = attrsOf(tab);
      const pos = Number(a['@_w:pos']);
      if (!Number.isFinite(pos)) continue;
      if (a['@_w:val'] === 'clear') tabs.delete(pos);
      else tabs.set(pos, { pos, val: a['@_w:val'] ?? 'left', leader: a['@_w:leader'] && a['@_w:leader'] !== 'none' ? a['@_w:leader'] : null });
    }
  }
  // 번호 {numId, ilvl, fromStyle}: 직접 지정(fromStyle=false)은 저장 때 w:numPr로 다시 쓴다(numId "0" = 스타일 번호 끄기)
  const num = numId && (numId !== '0' || !fromStyle) ? { numId, ilvl: ilvl ?? 0, fromStyle } : null;
  const tabList = [...tabs.values()].filter((t) => t.val !== 'num' && t.val !== 'bar').sort((a, b) => a.pos - b.pos);
  return { num, breakBefore, tabs: tabList.length > 0 ? tabList : null };
}

// ── 글자(실행문) ──────────────────────────────────────────────────────────

interface RunMarks {
  readonly strong: boolean;
  readonly em: boolean;
  readonly underline: boolean;
  readonly strike: boolean;
  readonly color: string | null;
  readonly fontSizePt: number | null;
  readonly font: { readonly ascii: string | null; readonly eastAsia: string | null } | null;
  readonly vertAlign: 'subscript' | 'superscript' | null;
  readonly highlight: string | null;
  /** 모델링하지 않는 w:rPr 자식(원본 XML)과 문자 스타일 */
  readonly extra: { readonly refs: string[] | null; readonly rStyle: string | null } | null;
  readonly comments?: readonly string[];
  /** 감싸는 요소(하이퍼링크·변경 내용 삽입·스마트 태그 등) — 저장 때 같은 요소로 다시 감싼다 */
  readonly wraps?: readonly Mark[];
}

const NO_MARKS: RunMarks = {
  strong: false, em: false, underline: false, strike: false, color: null, fontSizePt: null, font: null, vertAlign: null, highlight: null, extra: null,
};

/** w:shd@w:fill — 6자리 hex만 색으로 본다("auto"·무늬 채움 등은 음영 없음). */
function shadingFillOf(children: readonly XmlNode[]): string | null {
  const shd = findFirst(children, 'w:shd');
  return hexColor(shd ? attrsOf(shd)['@_w:fill'] : undefined);
}

function runMarksOf(runChildren: readonly XmlNode[], inherited: RunDefaults, ctx: ImportCtx): RunMarks {
  const rPr = findFirst(runChildren, 'w:rPr');
  const rPrChildren = rPr ? childrenOf(rPr) : [];
  const underline = findFirst(rPrChildren, 'w:u');
  const sz = findFirst(rPrChildren, 'w:sz');
  const szVal = sz ? attrsOf(sz)['@_w:val'] : undefined;
  // 직접 이름(literal)이 있는 글꼴만 글꼴 마크로 — 테마 글꼴만 있으면 원본 그대로 두고(저장 때 되돌림) 화면은 문서 기본 글꼴을 물려받는다
  const rFonts = findFirst(rPrChildren, 'w:rFonts');
  const fontAttrs = rFonts ? attrsOf(rFonts) : {};
  const ascii = fontAttrs['@_w:ascii'] ?? fontAttrs['@_w:hAnsi'] ?? null;
  const eastAsia = fontAttrs['@_w:eastAsia'] ?? null;
  const vertAlignEl = findFirst(rPrChildren, 'w:vertAlign');
  const vertAlignVal = vertAlignEl ? attrsOf(vertAlignEl)['@_w:val'] : undefined;
  const highlightName = attrsOf(findFirst(rPrChildren, 'w:highlight') ?? {})['@_w:val'];
  const namedHighlight = highlightName ? HIGHLIGHT_COLORS[highlightName] : undefined;
  const shdHighlight = namedHighlight ? null : shadingFillOf(rPrChildren);
  const bold = onOff(rPrChildren, 'w:b');
  const italic = onOff(rPrChildren, 'w:i');
  const strike = onOff(rPrChildren, 'w:strike');
  const colorVal = attrsOf(findFirst(rPrChildren, 'w:color') ?? {})['@_w:val'];
  const color = hexColor(colorVal);
  const underlineOn = underline !== undefined && attrsOf(underline)['@_w:val'] !== 'none';
  const sizePt = szVal && Number.isFinite(Number(szVal)) ? Number(szVal) / 2 : null;
  const fontMark = ascii || eastAsia ? { ascii, eastAsia } : null;
  // 위에서 마크로 옮긴 것만 "모델링함" — 나머지(끄기 값 w:b val=0, 테마 글꼴, 문자 스타일 등)는 원본 그대로 보존한다
  const modeled = new Set<string>();
  if (fontMark) modeled.add('w:rFonts');
  if (bold) modeled.add('w:b');
  if (italic) modeled.add('w:i');
  if (strike) modeled.add('w:strike');
  if (color) modeled.add('w:color');
  if (sizePt) { modeled.add('w:sz'); modeled.add('w:szCs'); }
  if (namedHighlight) modeled.add('w:highlight');
  if (shdHighlight) modeled.add('w:shd');
  if (underlineOn) modeled.add('w:u');
  if (vertAlignVal === 'subscript' || vertAlignVal === 'superscript') modeled.add('w:vertAlign');
  const extraRefs = extrasOf(rPrChildren, modeled, ctx.fragments);
  const rStyle = attrsOf(findFirst(rPrChildren, 'w:rStyle') ?? {})['@_w:val'] ?? null;
  return {
    strong: bold ?? inherited.strong ?? false,
    em: italic ?? inherited.em ?? false,
    underline: underlineOn,
    strike: strike ?? false,
    color: color ?? (colorVal === undefined ? (inherited.color ?? null) : null),
    fontSizePt: sizePt,
    font: fontMark ?? inherited.font ?? null,
    vertAlign: vertAlignVal === 'subscript' || vertAlignVal === 'superscript' ? vertAlignVal : null,
    highlight: namedHighlight ?? shdHighlight,
    extra: extraRefs || rStyle ? { refs: extraRefs, rStyle } : null,
  };
}

function marksOf(marks: RunMarks): Mark[] {
  const m = docxSchema.marks;
  return [
    marks.strong ? m.strong.create() : null,
    marks.em ? m.em.create() : null,
    marks.underline ? m.underline.create() : null,
    marks.strike ? m.strike.create() : null,
    marks.color ? m.textColor.create({ color: marks.color }) : null,
    marks.fontSizePt ? m.fontSize.create({ size: marks.fontSizePt }) : null,
    marks.font ? m.fontFamily.create(marks.font) : null,
    marks.vertAlign ? m[marks.vertAlign].create() : null,
    marks.highlight ? m.highlight.create({ color: marks.highlight }) : null,
    marks.extra ? m.runExtra.create(marks.extra) : null,
    ...(marks.wraps ?? []),
    ...(marks.comments ?? []).map((id) => m.comment.create({ id })),
  ].filter((mark): mark is Mark => mark !== null);
}

/** 마크가 적용된 텍스트 조각을 ProseMirror text 노드로 변환한다 (빈 문자열은 제외). */
function textNodeWithMarks(text: string, marks: RunMarks): PMNode | null {
  return text === '' ? null : docxSchema.text(text, marksOf(marks));
}

const EMU_PER_PX = 9525;
const emuToPx = (value: string | undefined): number | null => {
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? Math.round((n / EMU_PER_PX) * 10) / 10 : null;
};

/** 관계 대상(word/ 기준 상대 경로) → zip 안 그림 경로. word/media/ 바로 아래 브라우저가 그릴 수 있는 그림만(그 밖은 null). */
function mediaPathOf(ctx: ImportCtx, rId: string | undefined): string | null {
  const rel = rId ? ctx.rels.get(rId) : undefined;
  if (!rel || rel.external) return null;
  // 대상은 보통 word/ 기준 상대 경로("media/image1.png"), 가끔 패키지 절대 경로("/word/media/…")
  const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`;
  const ext = path.split('.').pop()?.toLowerCase() ?? '';
  return /^word\/media\/[^/]+$/.test(path) && IMAGE_MIME_BY_EXT[ext] ? path : null;
}

type OoxmlAttrs = Record<string, unknown>;

/** 글상자 안(w:txbxContent) 문단들을 화면용 PM JSON으로 */
function textboxContentOf(list: readonly XmlNode[], ctx: ImportCtx): unknown[] | null {
  const boxes: XmlNode[] = [];
  const collect = (nodes: readonly XmlNode[]) => {
    for (const node of nodes) {
      if (tagOf(node) === 'w:txbxContent') boxes.push(node);
      else collect(childrenOf(node));
    }
  };
  collect(list);
  if (boxes.length === 0) return null;
  return boxes.flatMap((box) => blocksOf(childrenOf(box), ctx, false).map((n) => n.toJSON()));
}

/** DrawingML 채우기·선 색(a:solidFill/a:srgbClr) */
function shapeBoxOf(list: readonly XmlNode[]): { fill: string | null; line: string | null; lineW: number | null } {
  const spPr = findDeep(list, 'wps:spPr');
  const spChildren = spPr ? childrenOf(spPr) : [];
  const fill = findFirst(spChildren, 'a:solidFill');
  const ln = findFirst(spChildren, 'a:ln');
  const lnFill = ln ? findFirst(childrenOf(ln), 'a:solidFill') : undefined;
  const colorIn = (el: XmlNode | undefined) => hexColor(attrsOf(findFirst(el ? childrenOf(el) : [], 'a:srgbClr') ?? {})['@_val']);
  return { fill: colorIn(fill), line: colorIn(lnFill), lineW: ln ? Math.max(1, Math.round((emuToPx(attrsOf(ln)['@_w']) ?? 1))) : null };
}

/** w:drawing(줄 안 그림·떠 있는 그림·글상자·도형·차트) → ooxml 노드의 화면 속성 */
function drawingAttrsOf(drawing: XmlNode, ctx: ImportCtx): OoxmlAttrs {
  const children = childrenOf(drawing);
  const container = findFirst(children, 'wp:inline') ?? findFirst(children, 'wp:anchor');
  const inner = container ? childrenOf(container) : [];
  const extent = attrsOf(findFirst(inner, 'wp:extent') ?? {});
  const widthPx = emuToPx(extent['@_cx']);
  const heightPx = emuToPx(extent['@_cy']);
  let float: string | null = null;
  if (container && tagOf(container) === 'wp:anchor') {
    const posH = findFirst(inner, 'wp:positionH');
    const align = posH ? textOf(childrenOf(findFirst(childrenOf(posH), 'wp:align') ?? {})) : '';
    const offset = posH ? emuToPx(textOf(childrenOf(findFirst(childrenOf(posH), 'wp:posOffset') ?? {}))) : null;
    const half = (ctx.contentWidthPx ?? 600) / 2;
    float = align === 'right' || align === 'outside' || (!align && offset !== null && offset + (widthPx ?? 0) / 2 > half) ? 'right' : 'left';
  }
  const docPr = attrsOf(findFirst(inner, 'wp:docPr') ?? {});
  const label = docPr['@_descr'] || docPr['@_name'] || '';
  const content = textboxContentOf(inner, ctx);
  if (content) return { kind: 'textbox', widthPx, heightPx, float, content, box: shapeBoxOf(inner), text: content.length ? deepTextOf(inner) : '' };
  const blip = findDeep(inner, 'a:blip');
  if (blip) return { kind: 'image', widthPx, heightPx, float, media: mediaPathOf(ctx, attrsOf(blip)['@_r:embed']), text: label };
  const graphicData = attrsOf(findDeep(inner, 'a:graphicData') ?? {})['@_uri'] ?? '';
  return { kind: 'image', widthPx, heightPx, float, media: null, text: /chart/.test(graphicData) ? '차트' : label || '도형' };
}

/** VML(w:pict·w:object) — 옛 형식 그림·글상자. 크기는 v:shape의 style(pt). */
function vmlAttrsOf(el: XmlNode, ctx: ImportCtx): OoxmlAttrs {
  const list = childrenOf(el);
  const shape = findDeep(list, 'v:shape') ?? findDeep(list, 'v:rect') ?? findDeep(list, 'v:roundrect');
  const style = shape ? attrsOf(shape)['@_style'] ?? '' : '';
  const pt = (name: string) => {
    const m = new RegExp(`(?:^|;)\\s*${name}:\\s*([\\d.]+)(pt|in|px)?`).exec(style);
    if (!m) return null;
    const v = Number(m[1]);
    return Math.round(m[2] === 'in' ? v * 96 : m[2] === 'px' ? v : (v * 4) / 3);
  };
  const float = /mso-position-horizontal:\s*right/.test(style) ? 'right' : /position:\s*absolute/.test(style) ? 'left' : null;
  const content = textboxContentOf(list, ctx);
  if (content) return { kind: 'textbox', widthPx: pt('width'), heightPx: pt('height'), float, content, box: {}, text: deepTextOf(list) };
  const imagedata = findDeep(list, 'v:imagedata');
  return { kind: 'image', widthPx: pt('width'), heightPx: pt('height'), float, media: mediaPathOf(ctx, imagedata ? attrsOf(imagedata)['@_r:id'] : undefined), text: '그림' };
}

/** 원본 실행문(w:r)의 rPr과 자식 하나만 담은 조각의 참조 — 실행문 속성(xmlns 선언 포함)을 그대로 지닌다. 참조 태그는 그 자식의 태그. */
function runPieceRef(run: XmlNode, child: XmlNode, ctx: ImportCtx): string {
  const rPr = findFirst(childrenOf(run), 'w:rPr');
  const node: XmlNode = { 'w:r': rPr ? [rPr, child] : [child] };
  if (run[':@']) node[':@'] = run[':@'];
  return fragmentRef(ctx.fragments, tagOf(child) ?? 'w:r', xmlOf(node));
}

/** 문단 자식 하나(책갈피·단순 필드·삭제분 등)를 통째로 보존하는 참조 */
const elementRef = (el: XmlNode, ctx: ImportCtx): string => fragmentRef(ctx.fragments, tagOf(el) ?? '', xmlOf(el));

/** 그림·도형 노드 — 브라우저가 못 그리는 것(그림 경로 없음)은 서버가 그릴 목록에 넣는다 */
function drawingNode(ref: string, attrs: OoxmlAttrs, marks: RunMarks, ctx: ImportCtx): PMNode {
  if (attrs.kind === 'image' && !attrs.media) ctx.serverDrawn.add(refHashOf(ref));
  return ooxmlNode({ ref, ...attrs }, marks);
}

function ooxmlNode(attrs: OoxmlAttrs, marks: RunMarks): PMNode {
  return docxSchema.nodes.ooxml.create(attrs, null, marksOf(marks));
}

/** 실행문 안에서 버리는 것 — Word가 다시 만드는 표시(마지막 쪽 나눔 위치·맞춤법 표시)와 메모 참조(메모 마크로 다시 만든다) */
const DROPPED_RUN_CHILDREN = new Set(['w:rPr', 'w:lastRenderedPageBreak', 'w:commentReference', 'w:annotationRef']);

/** 실행문의 텍스트·줄바꿈·탭이 아닌 자식 하나 → 원본 그대로 보존하는 ooxml 노드 */
function preservedRunChild(run: XmlNode, child: XmlNode, marks: RunMarks, ctx: ImportCtx): PMNode {
  const tag = tagOf(child) ?? '';
  const ref = runPieceRef(run, child, ctx);
  if (tag === 'w:drawing') return drawingNode(ref, drawingAttrsOf(child, ctx), marks, ctx);
  if (tag === 'mc:AlternateContent') {
    const choiceDrawing = findDeep(childrenOf(findFirst(childrenOf(child), 'mc:Choice') ?? {}), 'w:drawing');
    const fallbackPict = findDeep(childrenOf(findFirst(childrenOf(child), 'mc:Fallback') ?? {}), 'w:pict');
    if (choiceDrawing) return drawingNode(ref, drawingAttrsOf(choiceDrawing, ctx), marks, ctx);
    if (fallbackPict) return drawingNode(ref, vmlAttrsOf(fallbackPict, ctx), marks, ctx);
  }
  if (tag === 'w:pict' || tag === 'w:object') return drawingNode(ref, vmlAttrsOf(child, ctx), marks, ctx);
  if (tag === 'w:footnoteReference' || tag === 'w:endnoteReference') {
    const noteType = tag === 'w:footnoteReference' ? 'footnote' : 'endnote';
    ctx.noteCount[noteType] += 1;
    return ooxmlNode({ ref, kind: 'note', noteType, noteId: attrsOf(child)['@_w:id'] ?? null, text: String(ctx.noteCount[noteType]) }, marks);
  }
  if (tag === 'w:footnoteRef' || tag === 'w:endnoteRef') return ooxmlNode({ ref, kind: 'note', text: ctx.noteLabel ?? '' }, marks);
  if (tag === 'w:sym') {
    const code = parseInt(attrsOf(child)['@_w:char'] ?? '', 16);
    return ooxmlNode({ ref, kind: 'text', text: Number.isFinite(code) ? String.fromCharCode(code >= 0xf000 ? code - 0xf000 : code) : '' }, marks);
  }
  if (tag === 'w:noBreakHyphen') return ooxmlNode({ ref, kind: 'text', text: '‑' }, marks);
  if (tag === 'w:ptab') return ooxmlNode({ ref, kind: 'text', text: '\t' }, marks);
  const hiddenTags = ['w:softHyphen', 'w:fldChar', 'w:instrText', 'w:separator', 'w:continuationSeparator', 'w:delInstrText'];
  if (hiddenTags.includes(tag)) return ooxmlNode({ ref, kind: 'hidden' }, marks);
  return ooxmlNode({ ref, kind: 'text', text: deepTextOf([child]) }, marks);
}

/**
 * w:r 하나의 자식들을 등장 순서대로: 글자(w:t)·줄바꿈(w:br·w:cr)·탭(w:tab)은 고칠 수 있는 노드로, 그 밖(그림·글상자·각주 참조·
 * 필드 문자 등)은 원본 그대로 담은 ooxml 노드로 만든다 — 예전엔 w:t·w:br 말고는 전부 버려 그림·탭이 사라졌다.
 */
function runInlineNodesOf(run: XmlNode, marks: RunMarks, ctx: ImportCtx): PMNode[] {
  const nodes: PMNode[] = [];
  for (const child of childrenOf(run)) {
    const tag = tagOf(child);
    if (!tag || tag === '#text' || DROPPED_RUN_CHILDREN.has(tag)) continue;
    if (tag === 'w:t') {
      const node = textNodeWithMarks(textOf(childrenOf(child)), marks);
      if (node) nodes.push(node);
    } else if (tag === 'w:br' || tag === 'w:cr') {
      const type = attrsOf(child)['@_w:type'];
      if (type === 'page') nodes.push(docxSchema.nodes.page_break.create());
      else if (type === 'column') nodes.push(preservedRunChild(run, child, marks, ctx)); // 단 나누기 — 원본 그대로
      else nodes.push(docxSchema.nodes.hard_break.create(null, null, marksOf(marks)));
    } else if (tag === 'w:tab') {
      nodes.push(docxSchema.nodes.tab.create(null, null, marksOf(marks)));
    } else {
      nodes.push(preservedRunChild(run, child, marks, ctx));
    }
  }
  return nodes;
}

/** 지금 열려 있는(w:commentRangeStart는 지났고 End는 아직인) 메모 id — 범위가 문단을 넘어갈 수 있어
 * importDocx 한 번 동안 문서 순서대로 이어서 쓴다. */
let openCommentIds = new Set<string>();

const fldCharTypeOf = (run: XmlNode): string[] =>
  findAll(childrenOf(run), 'w:fldChar').map((f) => attrsOf(f)['@_w:fldCharType'] ?? '');

/**
 * 복합 필드(w:fldChar begin … separate … end) 하나를 한 덩어리 ooxml 노드로 — 원본 실행문들을 그대로 담고 화면엔 필드 결과 글자를
 * 보인다(PAGE·NUMPAGES는 편집기가 쪽 번호로 바꿔 그린다). 문단 안에서 끝나지 않는 필드(목차 등)는 이렇게 묶지 않는다.
 */
function fieldNodeOf(items: readonly XmlNode[], inherited: RunDefaults, ctx: ImportCtx): PMNode {
  let depth = 0;
  let afterSeparate = false;
  let instr = '';
  let result = '';
  let resultMarks: RunMarks | null = null;
  for (const item of items) {
    if (tagOf(item) !== 'w:r') {
      if (afterSeparate && depth === 1) result += deepTextOf([item]);
      continue;
    }
    for (const child of childrenOf(item)) {
      const tag = tagOf(child);
      if (tag === 'w:fldChar') {
        const type = attrsOf(child)['@_w:fldCharType'];
        if (type === 'begin') depth += 1;
        else if (type === 'end') depth -= 1;
        else if (type === 'separate' && depth === 1) afterSeparate = true;
      } else if (tag === 'w:instrText' && depth === 1 && !afterSeparate) instr += textOf(childrenOf(child));
      else if (afterSeparate && depth === 1 && (tag === 'w:t' || tag === 'w:tab')) {
        result += tag === 'w:t' ? textOf(childrenOf(child)) : '\t';
        resultMarks ??= runMarksOf(childrenOf(item), inherited, ctx);
      }
    }
  }
  const field = /^\s*([A-Z]+)/i.exec(instr)?.[1]?.toUpperCase() ?? null;
  return ooxmlNode({ ref: fragmentRef(ctx.fragments, 'w:fldChar', items.map(xmlOf).join('')), kind: 'field', field, text: result }, resultMarks ?? NO_MARKS);
}

/** 글자를 감싸는 요소 — 하이퍼링크, 변경 내용 삽입·이동(w:ins·w:moveTo), 스마트 태그·사용자 XML(w:smartTag·w:customXml) */
const WRAPPER_TAGS = new Set(['w:hyperlink', 'w:ins', 'w:moveTo', 'w:smartTag', 'w:customXml']);

/**
 * 감싸는 요소의 속성(원본 이름 그대로)과 속성 자식(w:smartTagPr·w:customXmlPr, 원본 그대로 참조)을 wrap 마크로 — 저장 때 같은 요소로
 * 다시 감싼다(depth: 겹친 순서, 바깥이 작다). 예전엔 풀어서 일반 글자로 저장해 하이퍼링크가 사라지고 변경 내용 삽입이 "수락"됐다.
 */
function wrapMarkOf(el: XmlNode, ctx: ImportCtx, depth: number): Mark {
  const tag = tagOf(el)!;
  const attrs: Record<string, string> = {};
  for (const [key, value] of Object.entries(attrsOf(el))) attrs[key.replace(/^@_/, '')] = String(value);
  const rel = tag === 'w:hyperlink' && attrs['r:id'] ? ctx.rels.get(attrs['r:id']) : undefined;
  const href = tag !== 'w:hyperlink' ? null : rel?.external ? rel.target : attrs['w:anchor'] ? `#${attrs['w:anchor']}` : null;
  const prEl = childrenOf(el).find((c) => /^w:(smartTagPr|customXmlPr)$/.test(tagOf(c) ?? ''));
  if (TRACKED_WRAP_TAGS.has(tag)) {
    const key = trackedWrapKey(tag, attrs);
    ctx.fragments.set(fragmentHash(key), key);
  }
  return docxSchema.marks.wrap.create({ tag, attrs, pr: prEl ? elementRef(prEl, ctx) : null, depth, href });
}

/**
 * 문단 자식을 문서 순서대로 훑는다 — 실행문·하이퍼링크(링크 마크로)·변경 내용 삽입 안의 글자를 읽고, 메모 범위로 메모 마크를 달고,
 * 그 밖의 문단 자식(책갈피·단순 필드·내용 컨트롤·수식·변경 내용 삭제분 등)은 원본 그대로 담는다.
 */
function inlineContentOf(pChildren: readonly XmlNode[], inherited: RunDefaults, ctx: ImportCtx, wraps: readonly Mark[] = []): PMNode[] {
  const nodes: PMNode[] = [];
  let field: XmlNode[] | null = null;
  let depth = 0;
  const plainRun = (run: XmlNode) => {
    const marks = runMarksOf(childrenOf(run), inherited, ctx);
    nodes.push(...runInlineNodesOf(run, { ...marks, wraps, comments: [...openCommentIds] }, ctx));
  };
  for (const child of pChildren) {
    const tag = tagOf(child);
    if (!tag || tag === '#text' || tag === 'w:pPr') continue;
    if (field) {
      field.push(child);
      if (tag === 'w:r') for (const type of fldCharTypeOf(child)) depth += type === 'begin' ? 1 : type === 'end' ? -1 : 0;
      if (depth === 0) {
        nodes.push(fieldNodeOf(field, inherited, ctx));
        field = null;
      }
      continue;
    }
    if (tag === 'w:r') {
      const types = fldCharTypeOf(child);
      if (types[0] === 'begin') {
        depth = types.reduce((d, t) => d + (t === 'begin' ? 1 : t === 'end' ? -1 : 0), 0);
        if (depth > 0) {
          field = [child];
          continue;
        }
      }
      plainRun(child);
    } else if (WRAPPER_TAGS.has(tag)) {
      const inner = childrenOf(child).filter((c) => !/^w:(smartTagPr|customXmlPr)$/.test(tagOf(c) ?? ''));
      nodes.push(...inlineContentOf(inner, inherited, ctx, [...wraps, wrapMarkOf(child, ctx, wraps.length)]));
    } else if (tag === 'w:commentRangeStart') {
      const id = attrsOf(child)['@_w:id'];
      if (id !== undefined) openCommentIds.add(String(id));
    } else if (tag === 'w:commentRangeEnd') {
      openCommentIds.delete(String(attrsOf(child)['@_w:id']));
    } else if (tag === 'w:proofErr') {
      // 맞춤법 검사 표시 — Word가 다시 만든다
    } else if (tag === 'w:del' || tag === 'w:moveFrom') {
      nodes.push(ooxmlNode({ ref: elementRef(child, ctx), kind: 'del', text: deepTextOf(childrenOf(child)) }, { ...NO_MARKS, wraps }));
    } else if (tag === 'w:fldSimple') {
      const field2 = /^\s*([A-Z]+)/i.exec(attrsOf(child)['@_w:instr'] ?? '')?.[1]?.toUpperCase() ?? null;
      const firstRun = findFirst(childrenOf(child), 'w:r');
      const marks = firstRun ? runMarksOf(childrenOf(firstRun), inherited, ctx) : NO_MARKS;
      nodes.push(ooxmlNode({ ref: elementRef(child, ctx), kind: 'field', field: field2, text: deepTextOf(childrenOf(child)) }, marks));
    } else if (/^w:(bookmarkStart|bookmarkEnd|permStart|permEnd|moveFromRangeStart|moveFromRangeEnd|moveToRangeStart|moveToRangeEnd)$/.test(tag)) {
      nodes.push(ooxmlNode({ ref: elementRef(child, ctx), kind: 'hidden' }, { ...NO_MARKS, wraps }));
    } else {
      // 내용 컨트롤(w:sdt)·수식(m:oMath) 등 — 원본 그대로, 화면엔 그 안의 글자
      nodes.push(ooxmlNode({ ref: elementRef(child, ctx), kind: 'text', text: deepTextOf(childrenOf(child)) }, { ...NO_MARKS, wraps }));
    }
  }
  // 문단 안에서 끝나지 않은 필드(목차처럼 여러 문단에 걸친 것) — 묶지 않고 하나씩(필드 문자는 숨은 조각으로) 읽는다
  if (field) for (const item of field) {
    if (tagOf(item) === 'w:r') plainRun(item);
    else nodes.push(...inlineContentOf([item], inherited, ctx, wraps));
  }
  return nodes;
}

// ── 문단 ─────────────────────────────────────────────────────────────────

function paragraphFrom(pNode: XmlNode, ctx: ImportCtx, cell: CellDefaults, sect: number | null): PMNode {
  const pChildren = childrenOf(pNode);
  const pPr = findFirst(pChildren, 'w:pPr');
  const pPrChildren = pPr ? childrenOf(pPr) : [];
  const pStyleId = attrsOf(findFirst(pPrChildren, 'w:pStyle') ?? {})['@_w:val'];
  const level = headingLevelOf(ctx.styles, pStyleId);
  // 문단이 이름 있는 스타일을 참조하면(예: "테이블 제목") 그 스타일 자체가 직접 지정한 정렬·굵기를 물려받는다
  const namedStyle = pStyleId ? namedStyleInfoOf(pStyleId, ctx.styles) : null;
  // 문단 자신의 w:jc가 없고 참조 스타일도 정렬을 안 정하면 — 아무 스타일도 참조하지 않을 때만 문서 기본 정렬을 물려받는다
  const align = alignOf(pPrChildren) ?? namedStyle?.align ?? (level === null && !pStyleId ? ctx.defaultAlign : null);
  const inherited: RunDefaults = { ...cell.run, strong: namedStyle?.bold || cell.run.strong };
  const inline = inlineContentOf(pChildren, inherited, ctx);
  const indent = indentOf(pPrChildren);
  const ownSpacing = spacingOf(pPrChildren);
  // 표 스타일의 문단 간격(예: Table Grid의 뒤 0·1줄)은 칸 안 문단이 따로 안 정했을 때 직접 값으로 얹는다
  const spacing = ownSpacing ?? (pStyleId ? null : cell.spacing);
  const shading = shadingFillOf(pPrChildren);
  const modeled = new Set(['w:pStyle', 'w:jc']);
  if (indent) modeled.add('w:ind');
  if (ownSpacing) modeled.add('w:spacing');
  if (shading) modeled.add('w:shd');
  if (!alignOf(pPrChildren) && findFirst(pPrChildren, 'w:jc')) modeled.delete('w:jc'); // left·start 등 — 원본 그대로
  const layout = paragraphLayoutOf(pPrChildren, ctx.styles, pStyleId);
  // 직접 번호(w:numPr에 ilvl·numId만)는 num 속성으로 모델링한다 — 편집기 번호 단추가 켜고 끄고, 저장 때 다시 쓴다
  const numPr = findFirst(pPrChildren, 'w:numPr');
  if (numPr && layout.num && !layout.num.fromStyle && childrenOf(numPr).every((c) => /^w:(ilvl|numId)$/.test(tagOf(c) ?? '') || '#text' in c)) modeled.add('w:numPr');
  // 구역 나눔은 내용이 같아도 구역마다 따로 — 참조에 구역 번호를 붙인다(같은 참조가 두 문단에 생기면 편집으로 나뉜 것)
  const pPrExtra = extrasOf(pPrChildren, modeled, ctx.fragments)?.map((ref) => (ref.startsWith('w:sectPr@') && sect !== null ? `${ref}#${sect}` : ref)) ?? null;
  const format = {
    align, indent, spacing, shading,
    pPrExtra,
    num: layout.num,
    breakBefore: layout.breakBefore,
    tabs: layout.tabs,
    sect,
  };
  return level !== null
    ? docxSchema.nodes.heading.create({ level, ...format, styleId: pStyleId }, inline)
    : docxSchema.nodes.paragraph.create({ ...format, styleId: pStyleId ? galleryStyleIdOf(pStyleId, ctx.styles) : null }, inline);
}

/** 이름(w:name)이 갤러리 스타일과 같으면 그 표준 id로(한국어 Word의 "a5" → "Title"), 아니면 원래 id 그대로. */
function galleryStyleIdOf(styleId: string, styles: readonly XmlNode[]): string {
  const style = styleOf(styles, 'paragraph', styleId);
  const nameEl = style ? findFirst(childrenOf(style), 'w:name') : undefined;
  const name = nameEl ? (attrsOf(nameEl)['@_w:val'] ?? '').toLowerCase() : '';
  const match = (Object.entries(GALLERY_STYLE_NAMES) as Array<[string, string]>).find(([, galleryName]) => galleryName === name);
  return match ? match[0] : styleId;
}

function intAttr(attrs: Record<string, string>, ...names: string[]): number | null {
  for (const name of names) {
    const value = Number(attrs[`@_w:${name}`]);
    if (attrs[`@_w:${name}`] !== undefined && Number.isFinite(value)) return Math.round(value);
  }
  return null;
}

/** w:ind — 왼쪽 들여쓰기(w:left, 새 문서는 w:start)와 첫 줄/내어쓰기. 전부 0이면 속성 없음. */
function indentOf(pPrChildren: readonly XmlNode[]): ParagraphIndent | null {
  const ind = findFirst(pPrChildren, 'w:ind');
  if (!ind) return null;
  const attrs = attrsOf(ind);
  const indent = {
    left: intAttr(attrs, 'left', 'start') ?? 0,
    right: intAttr(attrs, 'right', 'end') ?? 0,
    firstLine: intAttr(attrs, 'firstLine') ?? 0,
    hanging: intAttr(attrs, 'hanging') ?? 0,
  };
  return indent.left || indent.right || indent.firstLine || indent.hanging ? indent : null;
}

/** w:spacing — 줄 간격과 단락 앞/뒤 간격. 하나도 없으면 속성 없음. */
function spacingOf(pPrChildren: readonly XmlNode[]): ParagraphSpacing | null {
  const el = findFirst(pPrChildren, 'w:spacing');
  if (!el) return null;
  const attrs = attrsOf(el);
  const rule = attrs['@_w:lineRule'];
  const spacing: ParagraphSpacing = {
    line: intAttr(attrs, 'line'),
    lineRule: rule === 'auto' || rule === 'exact' || rule === 'atLeast' ? rule : null,
    before: intAttr(attrs, 'before'),
    after: intAttr(attrs, 'after'),
  };
  return spacing.line !== null || spacing.before !== null || spacing.after !== null ? spacing : null;
}

/** CSS px 기준 dxa(=twip, 1/20pt) 환산 — 96px = 1440dxa(1인치) */
const DXA_PER_PX = 15;

function dxaToPx(dxa: number): number {
  return Math.round(dxa / DXA_PER_PX);
}

// ── 블록(문단·표·보존 조각) ────────────────────────────────────────────────

/**
 * 블록 목록(본문·표 칸·글상자·머리글·각주 안)을 PM 블록으로 — 문단·표는 고칠 수 있게, 내용 컨트롤(w:sdt) 등은 원본 그대로 담은
 * ooxml_block으로. topLevel이면 구역(w:sectPr) 번호를 문단에 매긴다(쪽 나눔·머리글·다단 계산용).
 */
function blocksOf(children: readonly XmlNode[], ctx: ImportCtx, topLevel: boolean, cell: CellDefaults = NO_CELL_DEFAULTS, sections?: XmlNode[]): PMNode[] {
  const blocks: PMNode[] = [];
  for (const child of children) {
    const tag = tagOf(child);
    if (!tag || tag === '#text' || tag === 'w:sectPr' || tag === 'w:tcPr') continue;
    if (tag === 'w:p') {
      const sectPr = topLevel && sections ? findFirst(childrenOf(findFirst(childrenOf(child), 'w:pPr') ?? {}), 'w:sectPr') : undefined;
      if (sectPr && sections) sections.push(sectPr);
      blocks.push(paragraphFrom(child, ctx, cell, sectPr && sections ? sections.length - 1 : null));
    } else if (tag === 'w:tbl') {
      blocks.push(tableFrom(child, ctx));
    } else {
      // 내용 컨트롤(w:sdt)·사용자 XML(w:customXml) 등 — 원본 그대로, 화면엔 안의 블록을 읽기 전용으로
      const inner = tag === 'w:sdt' ? findFirst(childrenOf(child), 'w:sdtContent') : tag === 'w:customXml' ? child : undefined;
      const content = inner ? blocksOf(childrenOf(inner).filter((c) => tagOf(c) !== 'w:customXmlPr'), ctx, false, cell).map((n) => n.toJSON()) : null;
      blocks.push(docxSchema.nodes.ooxml_block.create({ ref: elementRef(child, ctx), content, text: deepTextOf(childrenOf(child)) }));
    }
  }
  return blocks;
}

// ── 표 ───────────────────────────────────────────────────────────────────

/** w:tcPr/w:shd@w:fill — 셀 음영 색 */
function backgroundOf(tcPrChildren: readonly XmlNode[]): string | null {
  return shadingFillOf(tcPrChildren);
}

function tcPrChildrenOf(tcNode: XmlNode): XmlNode[] {
  const tcPr = findFirst(childrenOf(tcNode), 'w:tcPr');
  return tcPr ? childrenOf(tcPr) : [];
}

interface CellMarginSides {
  readonly top: number | undefined;
  readonly right: number | undefined;
  readonly bottom: number | undefined;
  readonly left: number | undefined;
}

/** w:tblCellMar 또는 w:tcMar의 자식들에서 4변을 읽는다(dxa, 없는 변은 undefined). */
function cellMarginSidesOf(marChildren: readonly XmlNode[]): CellMarginSides {
  const sideOf = (...tags: string[]): number | undefined => {
    const el = tags.map((t) => findFirst(marChildren, t)).find(Boolean);
    if (!el) return undefined;
    const val = Number(attrsOf(el)['@_w:w']);
    return Number.isFinite(val) && val >= 0 ? val : undefined;
  };
  return { top: sideOf('w:top'), right: sideOf('w:right', 'w:end'), bottom: sideOf('w:bottom'), left: sideOf('w:left', 'w:start') };
}

/** Word 자체의 표준 기본 셀 여백(관례상의 "Normal Table" 기본값) — 위/아래 0, 좌/우 108dxa(0.075in). */
const STANDARD_CELL_MARGIN_DXA: CellMarginSides = { top: 0, right: 108, bottom: 0, left: 108 };

/** w:tblPr/w:tblCellMar — 이 표 자체가 지정한 기본 셀 여백. 없으면 null. */
function tableCellMarginOf(tblPrChildren: readonly XmlNode[]): CellMarginSides | null {
  const tblCellMar = findFirst(tblPrChildren, 'w:tblCellMar');
  return tblCellMar ? cellMarginSidesOf(childrenOf(tblCellMar)) : null;
}

/** "기본" 표 스타일(w:type="table" w:default="1") 자체의 w:tblCellMar — w:tblStyle을 지정하지 않은 모든 표가 암묵적으로 상속받는다. */
function defaultTableStyleCellMarginOf(styles: readonly XmlNode[]): CellMarginSides | null {
  const style = findAll(styles, 'w:style').find((s) => {
    const attrs = attrsOf(s);
    return attrs['@_w:type'] === 'table' && attrs['@_w:default'] === '1';
  });
  const tblPr = style && findFirst(childrenOf(style), 'w:tblPr');
  const tblCellMar = tblPr && findFirst(childrenOf(tblPr), 'w:tblCellMar');
  return tblCellMar ? cellMarginSidesOf(childrenOf(tblCellMar)) : null;
}

/** w:tcPr/w:tcMar — 이 셀만의 여백 재정의(있는 변만 표 기본값을 덮어쓴다). */
function cellMarginOverrideOf(tcPrChildren: readonly XmlNode[]): CellMarginSides | null {
  const tcMar = findFirst(tcPrChildren, 'w:tcMar');
  return tcMar ? cellMarginSidesOf(childrenOf(tcMar)) : null;
}

/**
 * 표 자체 지정 > 표 스타일 > Word 표준 기본값 순으로 표 여백을 정하고, 그 위에 이 셀만의 재정의를 얹어 최종 4변 여백(px)을
 * 확정한다. 반올림하지 않는다 — 딱 맞는 칸에서 0.x px가 줄바꿈 여부를 가른다(108dxa = 정확히 7.2px).
 */
function resolveCellMargin(
  tableDefault: CellMarginSides | null,
  styleDefault: CellMarginSides | null,
  cellOverride: CellMarginSides | null,
): CellPadding {
  const pick = (side: keyof CellMarginSides): number =>
    (cellOverride?.[side] ?? tableDefault?.[side] ?? styleDefault?.[side] ?? STANDARD_CELL_MARGIN_DXA[side]!) / DXA_PER_PX;
  return { topPx: pick('top'), rightPx: pick('right'), bottomPx: pick('bottom'), leftPx: pick('left') };
}

/** w:tcPr/w:vAlign — 셀 안 내용의 세로 정렬. docx는 top/center/bottom, CSS vertical-align은 top/middle/bottom. */
function verticalAlignOf(tcPrChildren: readonly XmlNode[]): string | null {
  const vAlign = findFirst(tcPrChildren, 'w:vAlign');
  if (!vAlign) return null;
  const val = attrsOf(vAlign)['@_w:val'];
  if (val === 'center') return 'middle';
  if (val === 'bottom' || val === 'top') return val;
  return null;
}

/** w:tcPr/w:gridSpan@w:val — 가로 병합 칸 수. 없으면 1(병합 없음). */
function gridSpanOf(tcPrChildren: readonly XmlNode[]): number {
  const gridSpan = findFirst(tcPrChildren, 'w:gridSpan');
  if (!gridSpan) return 1;
  const val = Number(attrsOf(gridSpan)['@_w:val']);
  return Number.isFinite(val) && val > 0 ? val : 1;
}

/** w:tcPr/w:vMerge — 'restart'면 세로 병합의 시작 칸, 값이 없거나 "continue"면 위 칸에 이어지는 자리. 태그 자체가 없으면 null. */
function vMergeOf(tcPrChildren: readonly XmlNode[]): 'restart' | 'continue' | null {
  const vMerge = findFirst(tcPrChildren, 'w:vMerge');
  if (!vMerge) return null;
  return attrsOf(vMerge)['@_w:val'] === 'restart' ? 'restart' : 'continue';
}

/** w:tblGrid/w:gridCol@w:w 목록(dxa) — 표 전체의 열별 너비. 없으면 빈 배열. */
function gridColWidthsOf(tblChildren: readonly XmlNode[]): number[] {
  const tblGrid = findFirst(tblChildren, 'w:tblGrid');
  if (!tblGrid) return [];
  return findAll(childrenOf(tblGrid), 'w:gridCol').map((col) => Number(attrsOf(col)['@_w:w']) || 0);
}

/**
 * 열 너비(dxa) — Word는 칸 너비(w:tcW)를 따르고 w:tblGrid는 참고값이다(python-docx 등이 만든 파일은 tblGrid가 균등이고
 * tcW만 실제 값이다). 병합하지 않은 칸의 dxa tcW가 가장 많은 행을 기준으로, 모르는 열은 tblGrid 값으로 채운다.
 */
function columnWidthsDxaOf(tblChildren: readonly XmlNode[], trNodes: readonly XmlNode[]): number[] {
  const grid = gridColWidthsOf(tblChildren);
  let best: Array<number | null> = [];
  let bestCount = 0;
  for (const tr of trNodes) {
    const widths: Array<number | null> = [];
    let count = 0;
    for (const tc of findAll(childrenOf(tr), 'w:tc')) {
      const tcPr = tcPrChildrenOf(tc);
      const span = gridSpanOf(tcPr);
      const tcW = attrsOf(findFirst(tcPr, 'w:tcW') ?? {});
      const w = Number(tcW['@_w:w']);
      const isDxa = (tcW['@_w:type'] ?? 'dxa') === 'dxa' && Number.isFinite(w) && w > 0;
      if (span === 1 && isDxa) count += 1;
      for (let i = 0; i < span; i += 1) widths.push(span === 1 && isDxa ? w : null);
    }
    if (count > bestCount) {
      best = widths;
      bestCount = count;
    }
  }
  const total = Math.max(grid.length, best.length);
  return Array.from({ length: total }, (_, i) => best[i] ?? grid[i] ?? 0);
}

/** docx 테두리 스타일(w:val) → CSS border-style. 알 수 없는 값은 실선으로 대체한다. */
const OOXML_BORDER_STYLE_TO_CSS: Record<string, string> = {
  double: 'double',
  triple: 'double',
  thinThickSmallGap: 'double',
  thickThinSmallGap: 'double',
  dashed: 'dashed',
  dashSmallGap: 'dashed',
  dotDash: 'dashed',
  dotted: 'dotted',
};

/**
 * 한 변(w:top/w:bottom/w:left/w:right/w:insideH/w:insideV)을 CellBorderSide로. undefined(요소 자체가 없음 — 상위 값을 따름)와
 * null(요소는 있지만 "테두리 없음"으로 명시)을 구분해야 특정 변만 "nil"로 재정의했을 때 기본 테두리로 잘못 되돌아가지 않는다.
 */
function borderSideOf(el: XmlNode | undefined): CellBorderSide | null | undefined {
  if (!el) return undefined;
  const attrs = attrsOf(el);
  const val = attrs['@_w:val'];
  if (!val || val === 'none' || val === 'nil') return null;
  const sz = Number(attrs['@_w:sz']);
  const style = OOXML_BORDER_STYLE_TO_CSS[val] ?? 'solid';
  // w:sz는 1/8pt 단위 → px = sz/6. 0.5px 단위로 반올림하고 최소 1px(이중선이 두 줄로 보이는 폭은 화면에서 정한다 — schema.ts)
  const widthPx = Number.isFinite(sz) && sz > 0 ? Math.max(1, Math.round((sz / 6) * 2) / 2) : 1;
  return { style, widthPx, color: hexColor(attrs['@_w:color']) ?? '#000000' };
}

interface BorderSet {
  readonly top: CellBorderSide | null | undefined;
  readonly right: CellBorderSide | null | undefined;
  readonly bottom: CellBorderSide | null | undefined;
  readonly left: CellBorderSide | null | undefined;
  readonly insideH: CellBorderSide | null | undefined;
  readonly insideV: CellBorderSide | null | undefined;
}

const EMPTY_BORDER_SET: BorderSet = { top: null, right: null, bottom: null, left: null, insideH: null, insideV: null };

/** w:tblBorders 또는 w:tcBorders의 자식들에서 6개 변을 읽는다(없는 변은 undefined). */
function borderSetOf(borderChildren: readonly XmlNode[]): BorderSet {
  return {
    top: borderSideOf(findFirst(borderChildren, 'w:top')),
    right: borderSideOf(findFirst(borderChildren, 'w:right') ?? findFirst(borderChildren, 'w:end')),
    bottom: borderSideOf(findFirst(borderChildren, 'w:bottom')),
    left: borderSideOf(findFirst(borderChildren, 'w:left') ?? findFirst(borderChildren, 'w:start')),
    insideH: borderSideOf(findFirst(borderChildren, 'w:insideH')),
    insideV: borderSideOf(findFirst(borderChildren, 'w:insideV')),
  };
}

/** 뒤 묶음에 있는 변만 앞 묶음을 덮는다 */
function overlayBorders(base: BorderSet, top: BorderSet | null): BorderSet {
  if (!top) return base;
  const pick = (side: keyof BorderSet) => (top[side] !== undefined ? top[side] : base[side]);
  return { top: pick('top'), right: pick('right'), bottom: pick('bottom'), left: pick('left'), insideH: pick('insideH'), insideV: pick('insideV') };
}

function bordersIn(children: readonly XmlNode[], tag: string): BorderSet | null {
  const el = findFirst(children, tag);
  return el ? borderSetOf(childrenOf(el)) : null;
}

/** 표 스타일의 한 영역(전체·첫 행·줄무늬 등)이 정한 모양 */
interface TableStyleRegion {
  readonly borders: BorderSet | null;
  readonly shd: string | null;
  readonly run: RunDefaults;
  readonly spacing: ParagraphSpacing | null;
}

interface TableStyleInfo {
  readonly tableBorders: BorderSet | null;
  readonly cellMargin: CellMarginSides | null;
  readonly whole: TableStyleRegion;
  readonly regions: ReadonlyMap<string, TableStyleRegion>;
}

/** 스타일 묶음 하나(w:style 또는 w:tblStylePr)의 pPr·rPr·tcPr에서 영역 모양을 읽어 앞 값 위에 얹는다 */
function regionOf(children: readonly XmlNode[], base: TableStyleRegion | undefined, ctx: ImportCtx): TableStyleRegion {
  const rPrChildren = childrenOf(findFirst(children, 'w:rPr') ?? {});
  const tcPrChildren = childrenOf(findFirst(children, 'w:tcPr') ?? {});
  const pPrChildren = childrenOf(findFirst(children, 'w:pPr') ?? {});
  const bold = onOff(rPrChildren, 'w:b');
  const italic = onOff(rPrChildren, 'w:i');
  const color = hexColor(attrsOf(findFirst(rPrChildren, 'w:color') ?? {})['@_w:val']);
  const font = resolveFontFamilyFromRFonts(findFirst(rPrChildren, 'w:rFonts'), ctx.themeXml);
  const run: RunDefaults = {
    strong: bold ?? base?.run.strong,
    em: italic ?? base?.run.em,
    color: color ?? base?.run.color ?? null,
    font: font ? { ascii: font.ascii, eastAsia: font.eastAsia } : (base?.run.font ?? null),
  };
  return {
    borders: base?.borders ? overlayBorders(base.borders, bordersIn(tcPrChildren, 'w:tcBorders')) : bordersIn(tcPrChildren, 'w:tcBorders'),
    shd: shadingFillOf(tcPrChildren) ?? base?.shd ?? null,
    run,
    spacing: spacingOf(pPrChildren) ?? base?.spacing ?? null,
  };
}

/** w:tblStyle이 가리키는 표 스타일(기반 사슬 포함)의 테두리·여백·영역별(첫 행·줄무늬 등) 모양 */
function tableStyleInfoOf(styleId: string | undefined, ctx: ImportCtx): TableStyleInfo | null {
  const chain = styleChainOf(ctx.styles, 'table', styleId);
  if (chain.length === 0) return null;
  let tableBorders: BorderSet | null = null;
  let cellMargin: CellMarginSides | null = null;
  let whole: TableStyleRegion | undefined;
  const regions = new Map<string, TableStyleRegion>();
  for (const style of chain) {
    const children = childrenOf(style);
    const tblPr = childrenOf(findFirst(children, 'w:tblPr') ?? {});
    const borders = bordersIn(tblPr, 'w:tblBorders');
    if (borders) tableBorders = tableBorders ? overlayBorders(tableBorders, borders) : borders;
    cellMargin = tableCellMarginOf(tblPr) ?? cellMargin;
    whole = regionOf(children, whole, ctx);
    for (const cond of findAll(children, 'w:tblStylePr')) {
      const type = attrsOf(cond)['@_w:type'] ?? '';
      regions.set(type, regionOf(childrenOf(cond), regions.get(type), ctx));
    }
  }
  return { tableBorders, cellMargin, whole: whole!, regions };
}

/** w:tblLook — 표 스타일의 어느 영역을 쓰는지(첫 행·첫 열·줄무늬 등). 옛 문서는 16진 w:val 비트로만 적는다. */
function tableLookOf(tblPrChildren: readonly XmlNode[]) {
  const a = attrsOf(findFirst(tblPrChildren, 'w:tblLook') ?? {});
  const bits = parseInt(a['@_w:val'] ?? '04A0', 16);
  const flag = (name: string, bit: number) => (a[`@_w:${name}`] !== undefined ? a[`@_w:${name}`] === '1' || a[`@_w:${name}`] === 'true' : (bits & bit) !== 0);
  return {
    firstRow: flag('firstRow', 0x20),
    lastRow: flag('lastRow', 0x40),
    firstColumn: flag('firstColumn', 0x80),
    lastColumn: flag('lastColumn', 0x100),
    noHBand: flag('noHBand', 0x200),
    noVBand: flag('noVBand', 0x400),
  };
}

/** 칸 위치 → 적용되는 표 스타일 영역(낮은 우선순위부터: 줄무늬 열 → 줄무늬 행 → 첫/끝 열 → 첫/끝 행 → 모서리) */
function regionTypesOf(look: ReturnType<typeof tableLookOf>, row: number, col: number, lastRow: boolean, lastCol: boolean, totalRows: number): string[] {
  const types: string[] = [];
  const isFirstRow = look.firstRow && row === 0;
  const isLastRow = look.lastRow && lastRow && totalRows > 1;
  const isFirstCol = look.firstColumn && col === 0;
  const isLastCol = look.lastColumn && lastCol;
  if (!look.noVBand && !isFirstCol && !isLastCol) types.push((col - (look.firstColumn ? 1 : 0)) % 2 === 0 ? 'band1Vert' : 'band2Vert');
  if (!look.noHBand && !isFirstRow && !isLastRow) types.push((row - (look.firstRow ? 1 : 0)) % 2 === 0 ? 'band1Horz' : 'band2Horz');
  if (isFirstCol) types.push('firstCol');
  if (isLastCol) types.push('lastCol');
  if (isFirstRow) types.push('firstRow');
  if (isLastRow) types.push('lastRow');
  if (isFirstRow && isFirstCol) types.push('nwCell');
  if (isFirstRow && isLastCol) types.push('neCell');
  if (isLastRow && isFirstCol) types.push('swCell');
  if (isLastRow && isLastCol) types.push('seCell');
  return types;
}

const ROW_REGIONS = new Set(['firstRow', 'lastRow', 'band1Horz', 'band2Horz']);

/**
 * 표 기본 테두리와 영역·칸 재정의를 합쳐 이 칸의 4변을 확정한다. 표 가장자리 칸은 바깥 변을, 안쪽은 insideH/insideV를 쓴다.
 * 영역(첫 행 등)의 테두리는 그 영역을 한 덩어리로 보고 같은 규칙으로 얹는다. 이웃 칸끼리의 우선순위는 브라우저의
 * border-collapse(굵은 선이 이김)에 맡긴다.
 */
function resolveCellBorders(
  tableDefault: BorderSet,
  regions: ReadonlyArray<{ type: string; borders: BorderSet | null }>,
  cellOverride: BorderSet | null,
  pos: { row: number; col: number; colspan: number; endRow: number; totalRows: number; totalCols: number },
): CellBorders {
  const isFirstRow = pos.row === 0;
  const isLastRow = pos.endRow >= pos.totalRows - 1;
  const isFirstCol = pos.col === 0;
  const isLastCol = pos.col + pos.colspan >= pos.totalCols;
  const sideOf = (set: BorderSet, rowRegion: boolean | null) => ({
    top: rowRegion === true || isFirstRow ? set.top : set.insideH,
    bottom: rowRegion === true || isLastRow ? set.bottom : set.insideH,
    left: rowRegion === false || isFirstCol ? set.left : set.insideV,
    right: rowRegion === false || isLastCol ? set.right : set.insideV,
  });
  let result: Record<'top' | 'right' | 'bottom' | 'left', CellBorderSide | null | undefined> = sideOf(tableDefault, null);
  const layers = [...regions.map((r) => (r.borders ? sideOf(r.borders, ROW_REGIONS.has(r.type)) : null)), cellOverride];
  for (const layer of layers) {
    if (!layer) continue;
    for (const side of ['top', 'right', 'bottom', 'left'] as const) if (layer[side] !== undefined) result = { ...result, [side]: layer[side] };
  }
  return { top: result.top ?? null, right: result.right ?? null, bottom: result.bottom ?? null, left: result.left ?? null };
}

interface CellPlacement {
  readonly tc: XmlNode;
  readonly row: number;
  readonly col: number;
  readonly colspan: number;
  rowspan: number;
}

/** w:tblPr/w:jc — 표 자체가 페이지 안에서 왼쪽/가운데/오른쪽 중 어디 놓이는지(셀 안 문단 정렬과 무관). */
function tableAlignOf(tblPrChildren: readonly XmlNode[]): string | null {
  const jc = findFirst(tblPrChildren, 'w:jc');
  if (!jc) return null;
  const val = attrsOf(jc)['@_w:val'];
  return val === 'center' || val === 'right' || val === 'end' ? (val === 'end' ? 'right' : val) : null;
}

/** w:trPr/w:trHeight — 행 높이. hRule="auto"(또는 요소 자체가 없음)는 "지정 없음"과 동일하게 null. */
function rowHeightOf(trPrChildren: readonly XmlNode[]): RowHeight | null {
  const trHeight = findFirst(trPrChildren, 'w:trHeight');
  if (!trHeight) return null;
  const attrs = attrsOf(trHeight);
  const val = Number(attrs['@_w:val']);
  if (!Number.isFinite(val) || val <= 0) return null;
  const hRule = attrs['@_w:hRule'];
  if (hRule === 'auto') return null;
  return { rule: hRule === 'exact' ? 'exact' : 'atLeast', px: dxaToPx(val) };
}

/**
 * 열 폭 합이 페이지 내용 폭(contentWidthPx)보다 넓으면, Word가 실제로 그러듯 비율을 유지한 채 내용 폭에 맞춰 축소한다.
 * contentWidthPx를 모르거나 이미 내용 폭 이내면 그대로 둔다(원본보다 좁은 표를 억지로 늘리지 않는다).
 */
function fitGridColWidthsPx(gridColWidthsPx: readonly number[], contentWidthPx: number | null): number[] {
  const total = gridColWidthsPx.reduce((sum, w) => sum + w, 0);
  if (contentWidthPx === null || total <= contentWidthPx || total <= 0) return [...gridColWidthsPx];
  const scale = contentWidthPx / total;
  return gridColWidthsPx.map((w) => Math.max(1, Math.round(w * scale)));
}

/**
 * OOXML 표 → PM 표. 병합(w:gridSpan=가로, w:vMerge=세로)은 prosemirror-tables의 colspan/rowspan으로("이어지는 칸"은 만들지 않고
 * 시작 칸의 rowspan만 늘린다). 표 스타일(w:tblStyle)의 테두리·음영·줄무늬·첫 행/열 글자 모양은 칸마다 계산해 칸 속성과 글자에 얹고,
 * 원본 w:tblPr·w:trPr·w:tcPr 중 모델링하지 않는 자식은 그대로 담아 저장 때 되돌린다.
 */
function tableFrom(tblNode: XmlNode, ctx: ImportCtx): PMNode {
  const tblChildren = childrenOf(tblNode);
  const tblPrChildren = childrenOf(findFirst(tblChildren, 'w:tblPr') ?? {});
  const trNodes = findAll(tblChildren, 'w:tr');
  const colWidths = fitGridColWidthsPx(columnWidthsDxaOf(tblChildren, trNodes).map(dxaToPx), ctx.contentWidthPx);
  const styleId = attrsOf(findFirst(tblPrChildren, 'w:tblStyle') ?? {})['@_w:val'];
  const style = tableStyleInfoOf(styleId, ctx);
  const ownBorders = bordersIn(tblPrChildren, 'w:tblBorders');
  const tableBorderDefaults = overlayBorders(style?.tableBorders ?? EMPTY_BORDER_SET, ownBorders);
  const tableCellMarginDefault = tableCellMarginOf(tblPrChildren);
  const look = tableLookOf(tblPrChildren);

  const rowPlacements: CellPlacement[][] = [];
  const originByCol = new Map<number, CellPlacement>();
  trNodes.forEach((tr, row) => {
    const placements: CellPlacement[] = [];
    let col = 0;
    for (const tc of findAll(childrenOf(tr), 'w:tc')) {
      const tcPrChildren = tcPrChildrenOf(tc);
      const gridSpan = gridSpanOf(tcPrChildren);
      if (vMergeOf(tcPrChildren) === 'continue') {
        const origin = originByCol.get(col);
        if (origin) origin.rowspan += 1;
        col += gridSpan;
        continue;
      }
      const placement: CellPlacement = { tc, row, col, colspan: gridSpan, rowspan: 1 };
      placements.push(placement);
      originByCol.set(col, placement);
      col += gridSpan;
    }
    rowPlacements.push(placements);
  });

  const totalRows = rowPlacements.length;
  const totalCols = Math.max(0, ...rowPlacements.map((placements) => placements.reduce((sum, p) => sum + p.colspan, 0)));
  const TC_MODELED = new Set(['w:tcW', 'w:gridSpan', 'w:vMerge', 'w:tcBorders', 'w:shd', 'w:vAlign', 'w:tcMar']);

  const rows = rowPlacements.map((placements, row) => {
    const trPrChildren = childrenOf(findFirst(childrenOf(trNodes[row]), 'w:trPr') ?? {});
    return docxSchema.nodes.table_row.create(
      { height: rowHeightOf(trPrChildren), trPrExtra: extrasOf(trPrChildren, new Set(['w:trHeight']), ctx.fragments) },
      placements.map((p) => {
        const widths = colWidths.slice(p.col, p.col + p.colspan);
        const colwidth = widths.length === p.colspan && widths.every((w) => w > 0) ? widths : null;
        const tcPrChildren = tcPrChildrenOf(p.tc);
        const endRow = p.row + p.rowspan - 1;
        const regionTypes = style ? regionTypesOf(look, p.row, p.col, endRow >= totalRows - 1, p.col + p.colspan >= totalCols, totalRows) : [];
        const regions = regionTypes.map((type) => ({ type, region: style!.regions.get(type) })).filter((r) => r.region);
        const layered = [style?.whole, ...regions.map((r) => r.region)].filter((r): r is TableStyleRegion => !!r);
        const run = layered.reduce<RunDefaults>((acc, r) => ({
          strong: r.run.strong ?? acc.strong,
          em: r.run.em ?? acc.em,
          color: r.run.color ?? acc.color,
          font: r.run.font ?? acc.font,
        }), {});
        const spacing = layered.reduce<ParagraphSpacing | null>((acc, r) => r.spacing ?? acc, null);
        const borders = resolveCellBorders(
          tableBorderDefaults,
          regions.map((r) => ({ type: r.type, borders: r.region!.borders })),
          bordersIn(tcPrChildren, 'w:tcBorders'),
          { row: p.row, col: p.col, colspan: p.colspan, endRow, totalRows, totalCols },
        );
        const background = backgroundOf(tcPrChildren) ?? layered.reduce<string | null>((acc, r) => r.shd ?? acc, null);
        const cellPadding = resolveCellMargin(tableCellMarginDefault, style?.cellMargin ?? ctx.styleTableCellMargin, cellMarginOverrideOf(tcPrChildren));
        const content = blocksOf(childrenOf(p.tc), ctx, false, { run, spacing });
        return docxSchema.nodes.table_cell.create(
          {
            colspan: p.colspan,
            rowspan: p.rowspan,
            colwidth,
            background,
            borders,
            verticalAlign: verticalAlignOf(tcPrChildren),
            cellPadding,
            tcPrExtra: extrasOf(tcPrChildren, TC_MODELED, ctx.fragments),
          },
          content.length > 0 ? content : [docxSchema.nodes.paragraph.create()],
        );
      }),
    );
  });
  return docxSchema.nodes.table.create(
    { align: tableAlignOf(tblPrChildren), tblPrExtra: extrasOf(tblPrChildren, new Set(['w:jc']), ctx.fragments) },
    rows,
  );
}

/** word/comments.xml의 w:comment들 — 메모 본문은 문단마다 줄을 바꿔 글자만 모은다. */
function commentsOf(xml: string): DocxComment[] {
  const root = findFirst(parser.parse(xml) as XmlNode[], 'w:comments');
  if (!root) return [];
  return findAll(childrenOf(root), 'w:comment').map((comment) => {
    const attrs = attrsOf(comment);
    const text = findAll(childrenOf(comment), 'w:p')
      .map((para) => {
        const parts: string[] = [];
        const walk = (list: readonly XmlNode[]) => {
          for (const node of list) {
            const tag = tagOf(node);
            if (tag === 'w:t') parts.push(textOf(childrenOf(node)));
            else if (tag === 'w:br') parts.push('\n');
            else if (tag) walk(childrenOf(node));
          }
        };
        walk(childrenOf(para));
        return parts.join('');
      })
      .join('\n');
    return {
      id: String(attrs['@_w:id'] ?? ''),
      author: attrs['@_w:author'] ?? '',
      initials: attrs['@_w:initials'] ?? '',
      date: attrs['@_w:date'] ?? '',
      text,
    };
  });
}

export interface DocxImportResult {
  readonly doc: PMNode;
  /** 페이지 크기·여백(w:sectPr/w:pgSz·w:pgMar) — 화면에 실제 Word 페이지 비율로 그리기 위함. 없으면 null. */
  readonly pageSetup: PageSetup | null;
  /** 원본 word/comments.xml의 메모 — 본문에 범위가 남아 있는 것만. */
  readonly comments: readonly DocxComment[];
  /** 보존 조각 저장소(해시 → 원본 XML) — 방이 열려 있는 동안 유지해 저장 때 넘긴다(export.ts) */
  readonly fragments: FragmentStore;
  /** 서버가 그림으로 바꿔 보여 줄 조각 해시(EMF/WMF·차트·도형) */
  readonly serverDrawn: ReadonlySet<string>;
}

/** w:sectPr(파싱된 트리)에서 페이지 크기·여백을 읽는다. w:pgSz가 없으면(비정상 문서) null.
 * 여백이 명시 안 된 변은 Word 기본값(1인치=1440dxa)으로 채운다. */
function pageSetupOf(sectPr: XmlNode | undefined): PageSetup | null {
  if (!sectPr) return null;
  const sectPrChildren = childrenOf(sectPr);
  const pgSz = findFirst(sectPrChildren, 'w:pgSz');
  if (!pgSz) return null;
  const pgSzAttrs = attrsOf(pgSz);
  const w = Number(pgSzAttrs['@_w:w']);
  const h = Number(pgSzAttrs['@_w:h']);
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) return null;
  const pgMar = findFirst(sectPrChildren, 'w:pgMar');
  const pgMarAttrs = pgMar ? attrsOf(pgMar) : {};
  const marginOf = (key: string): number => {
    const v = Number(pgMarAttrs[key]);
    return dxaToPx(Number.isFinite(v) && v >= 0 ? v : 1440);
  };
  return {
    widthPx: dxaToPx(w),
    heightPx: dxaToPx(h),
    marginTopPx: marginOf('@_w:top'),
    marginBottomPx: marginOf('@_w:bottom'),
    marginLeftPx: marginOf('@_w:left'),
    marginRightPx: marginOf('@_w:right'),
    orientation: pgSzAttrs['@_w:orient'] === 'landscape' ? 'landscape' : 'portrait',
  };
}

/**
 * word/document.xml 원본 문자열에서 (마지막) w:sectPr 요소를 그대로 문자열로 뽑아낸다 — 우리가
 * 모델링하지 않는 세부 설정(헤더/푸터 참조, 단 나누기 등)까지 손실 없이 내보내기 위해, 파싱 결과를
 * 재조립하는 대신 원본 텍스트를 그대로 들고 있다가 export.ts가 그대로 되돌려 쓴다.
 */
export function extractSectPrXml(documentXml: string): string | null {
  const matches = documentXml.match(/<w:sectPr\b[^>]*(?:\/>|>[\s\S]*?<\/w:sectPr>)/g);
  return matches && matches.length > 0 ? matches[matches.length - 1] : null;
}

/** Word가 실제로 쓰는 기본 폰트 크기가 없을 때의 대체값(10pt) — Word 새 문서의 기본값도 이와 같다. */
const FALLBACK_DEFAULT_FONT_SIZE_PT = 10;

/** w:docDefaults/w:rPrDefault/w:rPr/w:sz — 직접 서식도, 스타일도 지정하지 않은 실행문이 실제로
 * 쓰는 폰트 크기(half-point). 없으면 10pt로 대체한다. */
function defaultFontSizePtOf(docDefaults: XmlNode | undefined): number {
  if (!docDefaults) return FALLBACK_DEFAULT_FONT_SIZE_PT;
  const rPrDefault = findFirst(childrenOf(docDefaults), 'w:rPrDefault');
  const rPr = rPrDefault && findFirst(childrenOf(rPrDefault), 'w:rPr');
  const sz = rPr && findFirst(childrenOf(rPr), 'w:sz');
  const val = sz ? Number(attrsOf(sz)['@_w:val']) : NaN;
  return Number.isFinite(val) && val > 0 ? val / 2 : FALLBACK_DEFAULT_FONT_SIZE_PT;
}

/** w:docDefaults/w:rPrDefault/w:rPr/w:rFonts 요소 자체(있으면) — docDefaultsRFontsOf에서 쓴다. */
function docDefaultsRFontsOf(stylesChildren: readonly XmlNode[]): XmlNode | undefined {
  const docDefaults = findFirst(stylesChildren, 'w:docDefaults');
  const rPrDefault = docDefaults && findFirst(childrenOf(docDefaults), 'w:rPrDefault');
  const rPr = rPrDefault && findFirst(childrenOf(rPrDefault), 'w:rPr');
  return rPr && findFirst(childrenOf(rPr), 'w:rFonts');
}

/**
 * "기본" 문단 스타일(w:type="paragraph" w:default="1") — pStyle을 지정하지 않은 모든 문단이
 * 암묵적으로 쓰는 스타일이다. 실제 문서에서 이 스타일 자체가 docDefaults와 다른 폰트 크기·
 * 서체·정렬을 직접 지정하는 경우가 흔해서(예: docDefaults는 테마 맑은 고딕을 가리키지만
 * 기본 스타일 자체는 리터럴로 굴림을 지정), 그 경우 실제로 적용되는 값은 이 스타일 쪽이다.
 */
function findDefaultParagraphStyle(stylesChildren: readonly XmlNode[]): XmlNode | undefined {
  return findAll(stylesChildren, 'w:style').find((style) => {
    const attrs = attrsOf(style);
    return attrs['@_w:type'] === 'paragraph' && attrs['@_w:default'] === '1';
  });
}

interface DefaultParagraphStyleInfo {
  /** 스타일 자체가 지정한 폰트 크기(pt). 없으면 null → docDefaults로 폴백해야 한다. */
  readonly sizePt: number | null;
  /** 스타일 자체의 w:rFonts 요소. 없으면 undefined → docDefaults로 폴백해야 한다. */
  readonly rFonts: XmlNode | undefined;
  /** 스타일 자체의 문단 정렬(ProseMirror align 값). 없으면 null. */
  readonly align: string | null;
}

function defaultParagraphStyleInfoOf(stylesChildren: readonly XmlNode[]): DefaultParagraphStyleInfo {
  const style = findDefaultParagraphStyle(stylesChildren);
  const styleChildren = style ? childrenOf(style) : [];
  const rPr = findFirst(styleChildren, 'w:rPr');
  const rPrChildren = rPr ? childrenOf(rPr) : [];
  const sz = findFirst(rPrChildren, 'w:sz');
  const szVal = sz ? Number(attrsOf(sz)['@_w:val']) : NaN;
  const pPr = findFirst(styleChildren, 'w:pPr');
  return {
    sizePt: Number.isFinite(szVal) && szVal > 0 ? szVal / 2 : null,
    rFonts: findFirst(rPrChildren, 'w:rFonts'),
    align: alignOf(pPr ? childrenOf(pPr) : []),
  };
}

function stylesChildrenOf(stylesXml: string): XmlNode[] {
  const parsed = parser.parse(stylesXml) as XmlNode[];
  const stylesEl = findFirst(parsed, 'w:styles');
  return stylesEl ? childrenOf(stylesEl) : [];
}

/**
 * word/styles.xml의 문서 기본 폰트 크기를 뽑아낸다 — 직접 서식(w:rPr/w:sz)도 스타일도 없는
 * 실행문은 이 값을 상속받아 렌더링해야 브라우저 기본값(16px)이 아니라 Word가 실제로 보여주는
 * 크기와 같아진다(client/editor.ts가 .ProseMirror 기본 font-size로 적용). 기본 문단 스타일
 * 자체가 크기를 지정하면 그 값이 docDefaults보다 우선한다(실제 Word의 상속 순서).
 */
export function getDocxDefaultFontSizePt(docxBytes: Uint8Array): number {
  const entries = guardedUnzipSync(docxBytes, 'getDocxDefaultFontSizePt', { filter: (f) => f.name === 'word/styles.xml' });
  const stylesXmlBytes = entries['word/styles.xml'];
  if (!stylesXmlBytes) return FALLBACK_DEFAULT_FONT_SIZE_PT;
  const stylesChildren = stylesChildrenOf(strFromU8(stylesXmlBytes));
  const styleSizePt = defaultParagraphStyleInfoOf(stylesChildren).sizePt;
  return styleSizePt ?? defaultFontSizePtOf(findFirst(stylesChildren, 'w:docDefaults'));
}

/**
 * 직접 간격을 지정하지 않은 문단이 실제로 쓰는 문단 간격 — Word의 상속 순서(문서 기본값 w:docDefaults/w:pPrDefault →
 * 기본 문단 스타일 "표준"의 w:pPr/w:spacing)를 속성마다 따라간다. 어디에도 없으면 Word처럼 앞 0·뒤 0·1줄.
 * 예전엔 이걸 읽지 않아, 워드 기본 서식(뒤 8pt·1.08줄 등)으로 만든 거의 모든 문서가 편집기 고정값(1.75줄)으로 보였다.
 */
export function getDocxDefaultParagraphSpacing(docxBytes: Uint8Array): ParagraphSpacing {
  const fallback: ParagraphSpacing = { line: 240, lineRule: 'auto', before: 0, after: 0 };
  const entries = guardedUnzipSync(docxBytes, 'getDocxDefaultParagraphSpacing', { filter: (f) => f.name === 'word/styles.xml' });
  const stylesXmlBytes = entries['word/styles.xml'];
  if (!stylesXmlBytes) return fallback;
  const stylesChildren = stylesChildrenOf(strFromU8(stylesXmlBytes));
  const docDefaults = findFirst(stylesChildren, 'w:docDefaults');
  const pPrDefault = docDefaults && findFirst(childrenOf(docDefaults), 'w:pPrDefault');
  const defaultPPr = pPrDefault && findFirst(childrenOf(pPrDefault), 'w:pPr');
  const fromDefaults = defaultPPr ? spacingOf(childrenOf(defaultPPr)) : null;
  const style = findDefaultParagraphStyle(stylesChildren);
  const stylePPr = style && findFirst(childrenOf(style), 'w:pPr');
  const fromStyle = stylePPr ? spacingOf(childrenOf(stylePPr)) : null;
  // 줄 간격은 값과 규칙(auto·exact·atLeast)을 같은 곳에서 함께 가져온다
  const lineFrom = fromStyle?.line != null ? fromStyle : fromDefaults?.line != null ? fromDefaults : fallback;
  return {
    line: lineFrom.line,
    lineRule: lineFrom.lineRule ?? 'auto',
    before: fromStyle?.before ?? fromDefaults?.before ?? fallback.before,
    after: fromStyle?.after ?? fromDefaults?.after ?? fallback.after,
  };
}

/** a:fontScheme/a:minorFont(또는 a:majorFont) 안에서 라틴 기본 서체와 한글(Hang) 서체를 뽑는다. */
function themeTypefaceOf(themeXml: string, slot: 'minorFont' | 'majorFont'): { latin: string | null; hang: string | null } {
  const slotMatch = new RegExp(`<a:${slot}>([\\s\\S]*?)</a:${slot}>`).exec(themeXml);
  if (!slotMatch) return { latin: null, hang: null };
  const body = slotMatch[1];
  const latin = /<a:latin typeface="([^"]*)"/.exec(body)?.[1];
  const hang = /<a:font script="Hang" typeface="([^"]*)"/.exec(body)?.[1];
  return { latin: latin || null, hang: hang || null };
}

/**
 * w:rFonts의 한 서체 지정(예: w:eastAsia/w:eastAsiaTheme, w:ascii/w:asciiTheme)을 실제 서체
 * 이름으로 푼다. 직접 이름(literal)이 있으면 그걸 그대로 쓰고, 테마 참조(예: "majorEastAsia")만
 * 있으면 word/theme/theme1.xml의 해당 슬롯(major/minor)에서 script에 맞는 서체(한글은 Hang,
 * 라틴은 latin)를 찾는다.
 */
function resolveFontRef(
  literal: string | undefined,
  themeRef: string | undefined,
  themeXml: string | undefined,
  script: 'hang' | 'latin',
): string | null {
  if (literal) return literal;
  if (themeRef && themeXml) {
    const slot = /major/i.test(themeRef) ? 'majorFont' : 'minorFont';
    const typefaces = themeTypefaceOf(themeXml, slot);
    return typefaces[script];
  }
  return null;
}

/**
 * 문서 기본 서체 — Word는 한글(eastAsia)과 영문/숫자(ascii)에 서로 다른 서체를 쓸 수 있다
 * (예: eastAsia="굴림", ascii="Arial"). 브라우저의 CSS font-family 목록은 "이 글자를 그릴 수
 * 있는 첫 폰트"로 폴백해 스크립트별 지정을 구분하지 못하므로(굴림도 숫자 글리프를 갖고 있어
 * Arial까지 안 내려간다), 두 값을 따로 둔 채 돌려준다 — 실제 CSS 적용은 client/editor.ts가
 * unicode-range로 스크립트별로 갈라 처리한다.
 */
export interface ScriptFontFamily {
  readonly eastAsia: string | null;
  readonly ascii: string | null;
}

/** w:rFonts 요소 하나에서 한글·영문 서체를 각각 푼다(테마 참조 포함). 둘 다 없으면 null. */
function resolveFontFamilyFromRFonts(rFonts: XmlNode | undefined, themeXml: string | undefined): ScriptFontFamily | null {
  if (!rFonts) return null;
  const attrs = attrsOf(rFonts);
  const eastAsia = resolveFontRef(attrs['@_w:eastAsia'], attrs['@_w:eastAsiaTheme'], themeXml, 'hang');
  const ascii = resolveFontRef(attrs['@_w:ascii'], attrs['@_w:asciiTheme'], themeXml, 'latin');
  return eastAsia || ascii ? { eastAsia, ascii } : null;
}

/**
 * word/styles.xml의 문서 기본 서체(한글/영문 각각)를 뽑아낸다 — 직접 서식이 없는 실행문은
 * 이 서체로 보여야 브라우저의 기본 UI 폰트가 아니라 Word가 실제로 쓰는 서체와 같아진다.
 * 기본 문단 스타일 자체가 서체를 지정하면(흔히 리터럴로) 그 값이 docDefaults(흔히 테마
 * 참조)보다 우선한다 — 실제 Word의 상속 순서와 같다. 지정된 서체가 전혀 없으면 null
 * (기존 CSS 폴백 유지).
 */
export function getDocxDefaultFontFamily(docxBytes: Uint8Array): ScriptFontFamily | null {
  const entries = guardedUnzipSync(docxBytes, 'getDocxDefaultFontFamily', { filter: (f) => f.name === 'word/styles.xml' || f.name === 'word/theme/theme1.xml' });
  const stylesXmlBytes = entries['word/styles.xml'];
  if (!stylesXmlBytes) return null;
  const stylesChildren = stylesChildrenOf(strFromU8(stylesXmlBytes));
  const themeXmlBytes = entries['word/theme/theme1.xml'];
  const themeXml = themeXmlBytes ? strFromU8(themeXmlBytes) : undefined;
  const styleRFonts = defaultParagraphStyleInfoOf(stylesChildren).rFonts;
  return (
    resolveFontFamilyFromRFonts(styleRFonts, themeXml) ??
    resolveFontFamilyFromRFonts(docDefaultsRFontsOf(stylesChildren), themeXml)
  );
}

/**
 * 페이지 크기·여백만 가볍게 뽑아낸다(표·문단 전체를 파싱하는 importDocx보다 훨씬 싸다) —
 * 뷰어가 "이 문서를 실제 Word 페이지 비율로 얼마나 넓게 그릴지" 계산하는 용도(src/routes/viewer.ts).
 */
export function getDocxPageSetup(docxBytes: Uint8Array): PageSetup | null {
  const entries = guardedUnzipSync(docxBytes, 'getDocxPageSetup', { filter: (f) => f.name === 'word/document.xml' });
  const documentXmlBytes = entries['word/document.xml'];
  if (!documentXmlBytes) return null;
  const sectPrXml = extractSectPrXml(strFromU8(documentXmlBytes));
  if (!sectPrXml) return null;
  const parsed = parser.parse(sectPrXml) as XmlNode[];
  const sectPr = findFirst(parsed, 'w:sectPr');
  return pageSetupOf(sectPr);
}


/** word/_rels/<파트>.rels → r:id → 대상(파트 폴더 기준 상대 경로) */
function relsOf(entries: Record<string, Uint8Array>, partPath: string): Map<string, { target: string; external: boolean }> {
  const slash = partPath.lastIndexOf('/');
  const relsBytes = entries[`${partPath.slice(0, slash)}/_rels/${partPath.slice(slash + 1)}.rels`];
  const map = new Map<string, { target: string; external: boolean }>();
  if (!relsBytes) return map;
  for (const m of strFromU8(relsBytes).matchAll(/<Relationship\b[^>]*>/g)) {
    const id = /\bId="([^"]*)"/.exec(m[0])?.[1];
    const target = /\bTarget="([^"]*)"/.exec(m[0])?.[1];
    if (id && target) map.set(id, { target: target.replace(/&amp;/g, '&'), external: /TargetMode="External"/.test(m[0]) });
  }
  return map;
}

/** 압축을 푼 docx에서 스타일·테마 등 가져오기 공통값을 한 번에 */
function baseCtxOf(entries: Record<string, Uint8Array>, partPath: string, contentWidthPx: number | null): ImportCtx {
  const stylesXmlBytes = entries['word/styles.xml'];
  const styles = stylesXmlBytes ? stylesChildrenOf(strFromU8(stylesXmlBytes)) : [];
  const themeBytes = entries['word/theme/theme1.xml'];
  return {
    styles,
    defaultAlign: stylesXmlBytes ? defaultParagraphStyleInfoOf(styles).align : null,
    styleTableCellMargin: stylesXmlBytes ? defaultTableStyleCellMarginOf(styles) : null,
    rels: relsOf(entries, partPath),
    contentWidthPx,
    themeXml: themeBytes ? strFromU8(themeBytes) : undefined,
    noteCount: { footnote: 0, endnote: 0 },
    noteLabel: null,
    fragments: new Map(),
    serverDrawn: new Set(),
  };
}

/** w:sectPr의 다단(w:cols num>1). python-docx처럼 w:cols를 여러 번 쓴 파일은 마지막 것을 쓴다(Word 결과와 같다 — d06·a2 정답 PDF로 확인). */
function columnsOf(sectPr: XmlNode | undefined): { count: number; spacePx: number; sep: boolean } | null {
  const cols = findAll(childrenOf(sectPr ?? {}), 'w:cols').map(attrsOf).pop();
  if (!cols || !(Number(cols['@_w:num']) > 1)) return null;
  return { count: Number(cols['@_w:num']), spacePx: dxaToPx(Number(cols['@_w:space'] ?? 720)), sep: cols['@_w:sep'] === '1' || cols['@_w:sep'] === 'true' };
}

/** 다단 구역에 속한 연속 블록을 columns 노드로 감싼다 — 구역 k는 sect=k인 문단에서 끝난다(마지막 구역은 본문 끝까지). */
function groupColumns(blocks: readonly PMNode[], sections: readonly XmlNode[]): PMNode[] {
  const out: PMNode[] = [];
  let section = 0;
  let pending: PMNode[] = [];
  let pendingCols: ReturnType<typeof columnsOf> = null;
  const flush = () => {
    if (pending.length > 0 && pendingCols) out.push(docxSchema.nodes.columns.create(pendingCols, pending));
    pending = [];
  };
  for (const block of blocks) {
    pendingCols ??= columnsOf(sections[section]);
    if (pendingCols) pending.push(block);
    else out.push(block);
    const sect = block.attrs.sect;
    if (typeof sect === 'number') {
      flush();
      pendingCols = null;
      section = sect + 1;
    }
  }
  flush();
  return out;
}

/** docx(zip) 바이트 → ProseMirror 문서. 직접 고치지 않는 요소(그림·글상자·필드 등)는 원본 XML 그대로 ooxml 노드에 담는다. */
export function importDocx(docxBytes: Uint8Array): DocxImportResult {
  const entries = guardedUnzipSync(docxBytes, 'importDocx');
  const documentXmlBytes = entries['word/document.xml'];
  if (!documentXmlBytes) {
    throw new Error('word/document.xml을 찾을 수 없습니다 — 올바른 docx 파일이 아닙니다.');
  }
  const documentXml = strFromU8(documentXmlBytes);
  const parsed = parser.parse(documentXml) as XmlNode[];
  const documentEl = findFirst(parsed, 'w:document');
  const bodyEl = documentEl && findFirst(childrenOf(documentEl), 'w:body');
  const bodyChildren = bodyEl ? childrenOf(bodyEl) : [];
  const finalSectPr = findFirst(bodyChildren, 'w:sectPr');
  const pageSetup = pageSetupOf(finalSectPr);
  const contentWidthPx = pageSetup ? pageSetup.widthPx - pageSetup.marginLeftPx - pageSetup.marginRightPx : null;
  const ctx = baseCtxOf(entries, 'word/document.xml', contentWidthPx);

  openCommentIds = new Set();
  const sections: XmlNode[] = [];
  const blocks = blocksOf(bodyChildren, ctx, true, NO_CELL_DEFAULTS, sections);
  if (finalSectPr) sections.push(finalSectPr);
  const grouped = groupColumns(blocks, sections);

  const content = grouped.length > 0 ? grouped : [docxSchema.nodes.paragraph.create()];
  const doc = (docxSchema.nodes.doc as NodeType).create(null, content);
  const commentsXml = entries['word/comments.xml'];
  const used = new Set<string>();
  doc.descendants((node) => {
    for (const mark of node.marks) if (mark.type === docxSchema.marks.comment) used.add(mark.attrs.id as string);
  });
  const comments = commentsXml ? commentsOf(strFromU8(commentsXml)).filter((c) => used.has(c.id)) : [];
  return { doc, pageSetup, comments, fragments: ctx.fragments, serverDrawn: ctx.serverDrawn };
}

// ── 편집기 화면용 부가 정보(머리글·바닥글·각주·번호·스타일) ─────────────────────

/** 스타일 하나의 화면 모양 — 편집기가 CSS 규칙으로 바꾼다(docxLayout.ts). kind: p 문단 스타일, h 제목 수준, r 문자 스타일 */
export interface DocxStyleRule {
  readonly kind: 'p' | 'h' | 'r';
  readonly key: string;
  readonly color?: string | null;
  readonly sizePt?: number | null;
  readonly bold?: boolean | null;
  readonly italic?: boolean | null;
  readonly underline?: boolean | null;
  readonly caps?: boolean | null;
  readonly font?: ScriptFontFamily | null;
  readonly align?: string | null;
  readonly spacing?: ParagraphSpacing | null;
  readonly indent?: ParagraphIndent | null;
  readonly borderBottom?: CellBorderSide | null;
  readonly borderTop?: CellBorderSide | null;
  readonly shading?: string | null;
}

/** 구역 하나 — 쪽 나눔 종류와 그 구역의 머리글·바닥글(앞 구역에서 물려받은 것 포함, PM JSON 블록) */
export interface DocxSectionInfo {
  readonly type: string;
  readonly titlePg: boolean;
  readonly header: Readonly<Record<'default' | 'first' | 'even', unknown[] | null>>;
  readonly footer: Readonly<Record<'default' | 'first' | 'even', unknown[] | null>>;
  /** 쪽 위 끝에서 머리글까지, 아래 끝에서 바닥글까지(px) */
  readonly headerPx: number;
  readonly footerPx: number;
  /** 이 구역의 종이 크기·여백(px) — 편집 화면의 구역별 본문 폭·넘침 쪽 나눔 */
  readonly page: PageSetup | null;
}

export interface DocxRenderExtras {
  readonly numbering: NumberingDefs | null;
  readonly styles: readonly DocxStyleRule[];
  readonly sections: readonly DocxSectionInfo[];
  /** 각주·미주 본문(id → PM JSON 블록) */
  readonly notes: Readonly<Record<'footnote' | 'endnote', Readonly<Record<string, unknown[]>>>>;
  /** 기본 탭 간격(dxa, word/settings.xml w:defaultTabStop) */
  readonly defaultTabStopDxa: number;
  /** 짝수 쪽에 짝수 쪽 머리글·바닥글(word/settings.xml w:evenAndOddHeaders) */
  readonly evenAndOddHeaders: boolean;
  /** 서버에 깔린 동아시아 글꼴의 Word 1줄 비율(fontMetrics.ts) — 정보 경로가 붙인다 */
  readonly lineRatios?: Readonly<Record<string, number>>;
}

/** 구역 나눔(w:sectPr) 목록 — 본문 직속 문단의 것(문서 순서) + 본문 끝 것. importDocx(blocksOf)가 매기는 구역 번호와 같은 규칙이다. */
function sectionNodesOf(bodyChildren: readonly XmlNode[]): XmlNode[] {
  const out: XmlNode[] = [];
  for (const p of findAll(bodyChildren, 'w:p')) {
    const sectPr = findFirst(childrenOf(findFirst(childrenOf(p), 'w:pPr') ?? {}), 'w:sectPr');
    if (sectPr) out.push(sectPr);
  }
  const last = findFirst(bodyChildren, 'w:sectPr');
  if (last) out.push(last);
  return out;
}

/** 스타일 사슬(기반부터)을 합친 화면 모양 */
function styleRuleOf(kind: DocxStyleRule['kind'], key: string, chain: readonly XmlNode[], ctx: ImportCtx, fallbackSizePt: number | null): DocxStyleRule {
  const rLevels = chainPr(chain, 'w:rPr');
  const pLevels = chainPr(chain, 'w:pPr');
  const flag = (tag: string): boolean | null => {
    for (let i = rLevels.length - 1; i >= 0; i -= 1) {
      const v = onOff(rLevels[i], tag);
      if (v !== undefined) return v;
    }
    return null;
  };
  const sz = Number(attrsOf(lastOf(rLevels, 'w:sz') ?? {})['@_w:val']);
  const u = lastOf(rLevels, 'w:u');
  const pBdr = childrenOf(lastOf(pLevels, 'w:pBdr') ?? {});
  const spacing = pLevels.reduce<ParagraphSpacing | null>((acc, level) => {
    const s = spacingOf(level);
    if (!s) return acc;
    return { line: s.line ?? acc?.line ?? null, lineRule: s.line !== null ? s.lineRule : (acc?.lineRule ?? null), before: s.before ?? acc?.before ?? null, after: s.after ?? acc?.after ?? null };
  }, null);
  let font: ScriptFontFamily | null = null;
  for (const level of rLevels) font = resolveFontFamilyFromRFonts(findFirst(level, 'w:rFonts'), ctx.themeXml) ?? font;
  return {
    kind,
    key,
    color: hexColor(attrsOf(lastOf(rLevels, 'w:color') ?? {})['@_w:val']),
    sizePt: Number.isFinite(sz) && sz > 0 ? sz / 2 : fallbackSizePt,
    bold: flag('w:b') ?? (kind === 'h' ? false : null),
    italic: flag('w:i'),
    underline: u ? attrsOf(u)['@_w:val'] !== 'none' : null,
    caps: flag('w:caps'),
    font,
    align: alignOf(pLevels.flatMap((l) => (findFirst(l, 'w:jc') ? [findFirst(l, 'w:jc')!] : [])).slice(-1)),
    spacing,
    indent: pLevels.reduce<ParagraphIndent | null>((acc, l) => indentOf(l) ?? acc, null),
    borderBottom: borderSideOf(findFirst(pBdr, 'w:bottom')) ?? null,
    borderTop: borderSideOf(findFirst(pBdr, 'w:top')) ?? null,
    shading: pLevels.reduce<string | null>((acc, l) => shadingFillOf(l) ?? acc, null),
  };
}

/** 문서의 문단·문자 스타일 → 화면 규칙. 제목은 수준(h1~h6)으로, 문단은 편집기가 쓰는 styleId(갤러리 이름 맞춤)로. */
function styleRulesOf(ctx: ImportCtx): DocxStyleRule[] {
  const defaultSizePt = defaultParagraphStyleInfoOf(ctx.styles).sizePt ?? defaultFontSizePtOf(findFirst(ctx.styles, 'w:docDefaults'));
  const rules: DocxStyleRule[] = [];
  for (const style of findAll(ctx.styles, 'w:style')) {
    const a = attrsOf(style);
    const id = a['@_w:styleId'];
    if (!id || a['@_w:default'] === '1') continue;
    if (a['@_w:type'] === 'paragraph') {
      const chain = styleChainOf(ctx.styles, 'paragraph', id);
      const level = headingLevelOf(ctx.styles, id);
      rules.push(level !== null
        ? styleRuleOf('h', String(level), chain, ctx, defaultSizePt)
        : styleRuleOf('p', galleryStyleIdOf(id, ctx.styles), chain, ctx, null));
    } else if (a['@_w:type'] === 'character') {
      rules.push(styleRuleOf('r', id, styleChainOf(ctx.styles, 'character', id), ctx, null));
    }
  }
  return rules;
}

/** 파트 하나(머리글·바닥글 등)의 뿌리 요소 자식들 */
function partRootChildrenOf(entries: Record<string, Uint8Array>, partPath: string, rootTag: string): XmlNode[] | null {
  const bytes = entries[partPath];
  const root = bytes ? findFirst(parser.parse(strFromU8(bytes)) as XmlNode[], rootTag) : undefined;
  return root ? childrenOf(root) : null;
}

/** 압축을 푼 docx → 편집기 화면용 부가 정보. 원본 파트는 건드리지 않는다(머리글·각주는 저장 때 그대로 남는다). */
export function getDocxRenderExtras(docxBytes: Uint8Array): DocxRenderExtras {
  const entries = guardedUnzipSync(docxBytes, 'getDocxRenderExtras');
  const documentXml = entries['word/document.xml'] ? strFromU8(entries['word/document.xml']) : '';
  const pageSetup = getDocxPageSetup(docxBytes);
  const contentWidthPx = pageSetup ? pageSetup.widthPx - pageSetup.marginLeftPx - pageSetup.marginRightPx : null;
  const base = baseCtxOf(entries, 'word/document.xml', contentWidthPx);
  const docRels = base.rels;
  const blocksFromPart = (partPath: string, rootTag: string): unknown[] | null => {
    const children = partRootChildrenOf(entries, partPath, rootTag);
    if (!children) return null;
    const ctx: ImportCtx = { ...base, rels: relsOf(entries, partPath), noteCount: { footnote: 0, endnote: 0 } };
    return blocksOf(children, ctx, false).map((n) => n.toJSON());
  };

  // 구역: importDocx와 같은 규칙(본문 직속 문단의 w:sectPr + 본문 끝) — 문서 파서로 센다(내용 컨트롤 안·변경 기록 안 sectPr는 세지 않는다)
  const body = findFirst(childrenOf(findFirst(parser.parse(documentXml) as XmlNode[], 'w:document') ?? {}), 'w:body');
  const sections: DocxSectionInfo[] = [];
  const parts = new Map<string, unknown[] | null>();
  const partFor = (rId: string | undefined, rootTag: string): unknown[] | null => {
    const rel = rId ? docRels.get(rId) : undefined;
    if (!rel || rel.external) return null;
    // 대상은 보통 word/ 기준 상대 경로, 가끔 패키지 절대 경로("/word/header1.xml")
    const path = rel.target.startsWith('/') ? rel.target.slice(1) : `word/${rel.target}`;
    if (!parts.has(path)) parts.set(path, blocksFromPart(path, rootTag));
    return parts.get(path) ?? null;
  };
  for (const sect of sectionNodesOf(childrenOf(body ?? {}))) {
    const children = childrenOf(sect);
    const prev = sections[sections.length - 1];
    const refs = (tag: string, rootTag: string, inherited: DocxSectionInfo['header'] | undefined) => {
      const out: Record<'default' | 'first' | 'even', unknown[] | null> = { default: inherited?.default ?? null, first: inherited?.first ?? null, even: inherited?.even ?? null };
      for (const ref of findAll(children, tag)) {
        const a = attrsOf(ref);
        const type = (a['@_w:type'] ?? 'default') as 'default' | 'first' | 'even';
        if (type in out) out[type] = partFor(a['@_r:id'], rootTag);
      }
      return out;
    };
    const pgMar = attrsOf(findFirst(children, 'w:pgMar') ?? {});
    sections.push({
      type: attrsOf(findFirst(children, 'w:type') ?? {})['@_w:val'] ?? 'nextPage',
      titlePg: onOff(children, 'w:titlePg') === true,
      header: refs('w:headerReference', 'w:hdr', prev?.header),
      footer: refs('w:footerReference', 'w:ftr', prev?.footer),
      headerPx: dxaToPx(Number(pgMar['@_w:header'] ?? 720)),
      footerPx: dxaToPx(Number(pgMar['@_w:footer'] ?? 720)),
      page: pageSetupOf(sect),
    });
  }

  // 각주·미주: 문서에 나온 순서대로 번호를 매긴다(구분선 등 특수 항목은 뺀다)
  const notes: Record<'footnote' | 'endnote', Record<string, unknown[]>> = { footnote: {}, endnote: {} };
  for (const type of ['footnote', 'endnote'] as const) {
    const bytes = entries[`word/${type}s.xml`];
    if (!bytes) continue;
    const order = [...documentXml.matchAll(new RegExp(`<w:${type}Reference\\b[^>]*w:id="(-?\\d+)"`, 'g'))].map((m) => m[1]);
    const root = findFirst(parser.parse(strFromU8(bytes)) as XmlNode[], `w:${type}s`);
    for (const note of findAll(childrenOf(root ?? {}), `w:${type}`)) {
      const a = attrsOf(note);
      const id = a['@_w:id'] ?? '';
      if (a['@_w:type'] && a['@_w:type'] !== 'normal') continue;
      const label = order.indexOf(id) >= 0 ? String(order.indexOf(id) + 1) : '';
      const ctx: ImportCtx = { ...base, rels: relsOf(entries, `word/${type}s.xml`), noteCount: { footnote: 0, endnote: 0 }, noteLabel: label };
      notes[type][id] = blocksOf(childrenOf(note), ctx, false).map((n) => n.toJSON());
    }
  }

  const settings = entries['word/settings.xml'] ? strFromU8(entries['word/settings.xml']) : '';
  const tabStop = Number(/<w:defaultTabStop\b[^>]*w:val="(\d+)"/.exec(settings)?.[1]);
  const numberingBytes = entries['word/numbering.xml'];
  return {
    numbering: numberingBytes ? parseNumberingXml(strFromU8(numberingBytes)) : null,
    styles: styleRulesOf(base),
    sections,
    notes,
    defaultTabStopDxa: Number.isFinite(tabStop) && tabStop > 0 ? tabStop : 720,
    evenAndOddHeaders: /<w:evenAndOddHeaders\b(?![^>]*w:val="(0|false|off)")/.test(settings),
  };
}
