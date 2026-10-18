import { strFromU8, strToU8, zipSync } from 'fflate';
import { guardedUnzipSync } from '../zipGuard.js';
import type { Mark, Node as PMNode } from 'prosemirror-model';
import { extractSectPrXml } from './import.js';
import {
  GALLERY_PARAGRAPH_STYLES,
  GALLERY_STYLE_NAMES,
  type GalleryStyleId,
  HIGHLIGHT_COLORS,
  type CellBorderSide,
  type CellBorders,
  type CellPadding,
  type DocxComment,
  type DocxPageSetup,
  type ParagraphIndent,
  type ParagraphSpacing,
  type RowHeight,
} from './schema.js';

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function escapeAttr(text: string): string {
  return escapeXml(text).replace(/"/g, '&quot;');
}

function hexOf(color: unknown): string | null {
  const hex = String(color ?? '').replace('#', '').toUpperCase();
  return /^[0-9A-F]{6}$/.test(hex) ? hex : null;
}

const HIGHLIGHT_NAME_BY_HEX = new Map(Object.entries(HIGHLIGHT_COLORS).map(([name, hex]) => [hex.slice(1).toUpperCase(), name]));

/** w:rPr 자식은 스키마가 정한 순서를 지켜야 한다(rFonts, b, i, strike, color, sz, highlight, u, shd, vertAlign). */
function renderRun(text: string, marks: readonly Mark[]): string {
  const markOf = (name: string) => marks.find((mark) => mark.type.name === name);
  const rPrParts: string[] = [];
  const font = markOf('fontFamily');
  if (font) {
    const ascii = font.attrs.ascii as string | null;
    const eastAsia = font.attrs.eastAsia as string | null;
    const fontAttrs = [
      ascii ? `w:ascii="${escapeAttr(ascii)}" w:hAnsi="${escapeAttr(ascii)}" w:cs="${escapeAttr(ascii)}"` : '',
      eastAsia ? `w:eastAsia="${escapeAttr(eastAsia)}"` : '',
    ].filter(Boolean);
    if (fontAttrs.length > 0) rPrParts.push(`<w:rFonts ${fontAttrs.join(' ')}/>`);
  }
  if (markOf('strong')) rPrParts.push('<w:b/>');
  if (markOf('em')) rPrParts.push('<w:i/>');
  if (markOf('strike')) rPrParts.push('<w:strike/>');
  const colorHex = hexOf(markOf('textColor')?.attrs.color);
  if (colorHex) rPrParts.push(`<w:color w:val="${colorHex}"/>`);
  const sizeMark = markOf('fontSize');
  if (sizeMark) {
    const halfPoints = Math.round(Number(sizeMark.attrs.size) * 2);
    rPrParts.push(`<w:sz w:val="${halfPoints}"/><w:szCs w:val="${halfPoints}"/>`);
  }
  const highlightHex = hexOf(markOf('highlight')?.attrs.color);
  const highlightName = highlightHex ? HIGHLIGHT_NAME_BY_HEX.get(highlightHex) : undefined;
  if (highlightName) rPrParts.push(`<w:highlight w:val="${highlightName}"/>`);
  if (markOf('underline')) rPrParts.push('<w:u w:val="single"/>');
  if (highlightHex && !highlightName) rPrParts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${highlightHex}"/>`);
  if (markOf('subscript')) rPrParts.push('<w:vertAlign w:val="subscript"/>');
  else if (markOf('superscript')) rPrParts.push('<w:vertAlign w:val="superscript"/>');
  const rPr = rPrParts.length > 0 ? `<w:rPr>${rPrParts.join('')}</w:rPr>` : '';
  return `<w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r>`;
}

const ALIGN_TO_JC: Record<string, string> = { center: 'center', right: 'right', justify: 'both' };

/** w:pPr 자식 순서: pStyle, shd, spacing, ind, jc (CT_PPr 스키마 순서). */
function pPrXml(styleXml: string, attrs: Record<string, unknown>): string {
  const parts: string[] = [styleXml];
  const shading = hexOf(attrs.shading);
  if (shading) parts.push(`<w:shd w:val="clear" w:color="auto" w:fill="${shading}"/>`);
  const spacing = attrs.spacing as ParagraphSpacing | null | undefined;
  if (spacing) {
    const spacingAttrs = [
      spacing.before !== null ? `w:before="${spacing.before}"` : '',
      spacing.after !== null ? `w:after="${spacing.after}"` : '',
      spacing.line !== null ? `w:line="${spacing.line}"` : '',
      spacing.line !== null && spacing.lineRule ? `w:lineRule="${spacing.lineRule}"` : '',
    ].filter(Boolean);
    if (spacingAttrs.length > 0) parts.push(`<w:spacing ${spacingAttrs.join(' ')}/>`);
  }
  const indent = attrs.indent as ParagraphIndent | null | undefined;
  if (indent) {
    const indentAttrs = [
      indent.left ? `w:left="${indent.left}"` : '',
      indent.right ? `w:right="${indent.right}"` : '',
      indent.firstLine ? `w:firstLine="${indent.firstLine}"` : '',
      indent.hanging ? `w:hanging="${indent.hanging}"` : '',
    ].filter(Boolean);
    if (indentAttrs.length > 0) parts.push(`<w:ind ${indentAttrs.join(' ')}/>`);
  }
  const jc = typeof attrs.align === 'string' ? ALIGN_TO_JC[attrs.align] : undefined;
  if (jc) parts.push(`<w:jc w:val="${jc}"/>`);
  const inner = parts.filter(Boolean).join('');
  return inner ? `<w:pPr>${inner}</w:pPr>` : '';
}

/**
 * 메모 범위 표시를 문서 순서대로 낸다 — 글자마다 붙은 메모 마크를 보고, 앞 글자에는 있었는데 이
 * 글자에 없는 메모는 닫고(commentRangeEnd + 참조 실행문), 새로 생긴 메모는 연다(commentRangeStart).
 * 범위는 문단을 넘어갈 수 있으므로 문서 하나를 그리는 동안 상태를 이어서 쓴다. w:id는 처음 나온
 * 순서대로 0부터 다시 매긴다(편집기 안의 id는 uuid 등 아무 문자열이다).
 */
interface CommentRenderState {
  readonly numbers: Map<string, number>;
  readonly open: Set<string>;
}
let commentState: CommentRenderState | null = null;

function commentNumber(state: CommentRenderState, id: string): number {
  let n = state.numbers.get(id);
  if (n === undefined) {
    n = state.numbers.size;
    state.numbers.set(id, n);
  }
  return n;
}

function closeCommentXml(state: CommentRenderState, id: string): string {
  const n = commentNumber(state, id);
  state.open.delete(id);
  return `<w:commentRangeEnd w:id="${n}"/><w:r><w:commentReference w:id="${n}"/></w:r>`;
}

function commentTransitionXml(marks: readonly Mark[]): string {
  const state = commentState;
  if (!state) return '';
  const ids = new Set(marks.filter((mark) => mark.type.name === 'comment').map((mark) => String(mark.attrs.id)));
  let xml = '';
  for (const id of [...state.open]) if (!ids.has(id)) xml += closeCommentXml(state, id);
  for (const id of ids) {
    if (!state.open.has(id)) {
      state.open.add(id);
      xml += `<w:commentRangeStart w:id="${commentNumber(state, id)}"/>`;
    }
  }
  return xml;
}

/** word/comments.xml — 본문에 실제로 범위가 있는 메모만, 매긴 번호 순서대로. */
function renderCommentsXml(numbers: ReadonlyMap<string, number>, comments: ReadonlyMap<string, DocxComment>): string {
  const items = [...numbers.entries()]
    .sort((a, b) => a[1] - b[1])
    .map(([id, n]) => {
      const comment = comments.get(id);
      const paragraphs = (comment?.text ?? '').split('\n').map(
        (line, index) =>
          `<w:p>${index === 0 ? '<w:r><w:annotationRef/></w:r>' : ''}<w:r><w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r></w:p>`,
      );
      const date = comment?.date ? ` w:date="${escapeAttr(comment.date)}"` : '';
      return (
        `<w:comment w:id="${n}" w:author="${escapeAttr(comment?.author ?? '')}"${date} w:initials="${escapeAttr(comment?.initials ?? '')}">` +
        `${paragraphs.join('')}</w:comment>`
      );
    });
  return `${DOCUMENT_XML_HEADER}<w:comments xmlns:w="${NS_W}">${items.join('')}</w:comments>`;
}

/** 메모 파트를 새로 쓸 때 — 원본의 확장 메모 파트(답글·해결 상태 등, 메모 문단 id를 참조)는 새 메모와
 * 맞지 않으므로 지우고, comments.xml은 콘텐츠 형식·관계에 등록한다. */
const COMMENT_EXTENSION_PARTS = ['commentsExtended.xml', 'commentsIds.xml', 'commentsExtensible.xml'];
function registerCommentsPart(entries: Record<string, Uint8Array>): void {
  for (const part of COMMENT_EXTENSION_PARTS) delete entries[`word/${part}`];
  const types = entries['[Content_Types].xml'];
  if (types) {
    let xml = strFromU8(types);
    for (const part of COMMENT_EXTENSION_PARTS) xml = xml.replace(new RegExp(`<Override PartName="/word/${part}"[^>]*/>`), '');
    if (!xml.includes('/word/comments.xml')) {
      xml = xml.replace(
        '</Types>',
        '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>',
      );
    }
    entries['[Content_Types].xml'] = strToU8(xml);
  }
  const rels = entries['word/_rels/document.xml.rels'];
  if (rels) {
    let xml = strFromU8(rels);
    for (const part of COMMENT_EXTENSION_PARTS) xml = xml.replace(new RegExp(`<Relationship [^>]*Target="${part}"[^>]*/>`), '');
    if (!/Target="comments\.xml"/.test(xml)) {
      xml = xml.replace(
        '</Relationships>',
        '<Relationship Id="rIdDocviewComments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>',
      );
    }
    entries['word/_rels/document.xml.rels'] = strToU8(xml);
  }
}

function renderInline(node: PMNode): string {
  let xml = '';
  node.forEach((child) => {
    if (child.isText) {
      xml += commentTransitionXml(child.marks);
      xml += renderRun(child.text ?? '', child.marks);
    } else if (child.type.name === 'hard_break') {
      xml += '<w:r><w:br/></w:r>';
    } else if (child.type.name === 'page_break') {
      xml += '<w:r><w:br w:type="page"/></w:r>';
    }
  });
  return xml;
}

/** 갤러리 스타일 id → 이 문서에서 실제로 쓸 w:styleId (exportDocx가 styles.xml을 보고 정한다). */
type StyleIdMap = ReadonlyMap<string, string>;

let activeStyleIds: StyleIdMap = new Map();

/** 갤러리 스타일이 문서에 없을 때 styles.xml에 추가할 정의(Word 기본 서식과 같은 값). */
const GALLERY_STYLE_DEFS: Readonly<Record<GalleryStyleId, { readonly body: string }>> = {
  NoSpacing: {
    body: '<w:name w:val="No Spacing"/><w:uiPriority w:val="1"/><w:qFormat/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>',
  },
  Title: {
    body:
      '<w:name w:val="Title"/><w:next w:val="Normal"/><w:uiPriority w:val="10"/><w:qFormat/>' +
      '<w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:contextualSpacing/></w:pPr>' +
      '<w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>' +
      '<w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/></w:rPr>',
  },
  Subtitle: {
    body:
      '<w:name w:val="Subtitle"/><w:next w:val="Normal"/><w:uiPriority w:val="11"/><w:qFormat/>' +
      '<w:pPr><w:spacing w:after="160"/></w:pPr>' +
      '<w:rPr><w:color w:val="5A5A5A"/><w:spacing w:val="15"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>',
  },
  Quote: {
    body:
      '<w:name w:val="Quote"/><w:next w:val="Normal"/><w:uiPriority w:val="29"/><w:qFormat/>' +
      '<w:pPr><w:spacing w:before="200" w:after="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr>' +
      '<w:rPr><w:i/><w:iCs/><w:color w:val="404040"/></w:rPr>',
  },
  IntenseQuote: {
    body:
      '<w:name w:val="Intense Quote"/><w:next w:val="Normal"/><w:uiPriority w:val="30"/><w:qFormat/>' +
      '<w:pPr><w:pBdr><w:top w:val="single" w:sz="4" w:space="10" w:color="4472C4"/>' +
      '<w:bottom w:val="single" w:sz="4" w:space="10" w:color="4472C4"/></w:pBdr>' +
      '<w:spacing w:before="360" w:after="360"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/></w:pPr>' +
      '<w:rPr><w:i/><w:iCs/><w:color w:val="4472C4"/></w:rPr>',
  },
};

/**
 * 문서가 쓰는 갤러리 스타일마다 실제 w:styleId를 정한다 — 같은 이름(w:name)의 스타일이 이미 있으면
 * 그 id를 쓰고(한국어 Word는 내장 스타일 id가 "a5" 같은 값이다), 없으면 styles.xml에 정의를 추가한다.
 */
function resolveGalleryStyles(doc: PMNode, stylesXml: string | null): { ids: StyleIdMap; stylesXml: string | null } {
  const used = new Set<GalleryStyleId>();
  doc.descendants((node) => {
    const id = node.attrs.styleId as string | null | undefined;
    if (id && (GALLERY_PARAGRAPH_STYLES as readonly string[]).includes(id)) used.add(id as GalleryStyleId);
  });
  const ids = new Map<string, string>();
  if (used.size === 0 || stylesXml === null) return { ids, stylesXml };
  const idByName = new Map<string, string>();
  const existingIds = new Set<string>();
  for (const match of stylesXml.matchAll(/<w:style\b[^>]*w:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g)) {
    existingIds.add(match[1]);
    const name = /<w:name\s+w:val="([^"]+)"/.exec(match[2])?.[1]?.toLowerCase();
    if (name && !idByName.has(name)) idByName.set(name, match[1]);
  }
  let added = '';
  for (const id of used) {
    const def = GALLERY_STYLE_DEFS[id];
    const existing = idByName.get(GALLERY_STYLE_NAMES[id]) ?? (existingIds.has(id) ? id : undefined);
    if (existing) {
      ids.set(id, existing);
    } else {
      ids.set(id, id);
      added += `<w:style w:type="paragraph" w:styleId="${id}">${def.body}</w:style>`;
    }
  }
  return { ids, stylesXml: added ? stylesXml.replace(/<\/w:styles>\s*$/, `${added}</w:styles>`) : stylesXml };
}

function renderParagraph(node: PMNode, prefix = '', styleIds: StyleIdMap = activeStyleIds): string {
  const prefixRun = prefix ? renderRun(prefix, []) : '';
  const styleId = node.attrs.styleId as string | null;
  const resolved = styleId ? (styleIds.get(styleId) ?? styleId) : null;
  const pPr = pPrXml(resolved ? `<w:pStyle w:val="${escapeAttr(resolved)}"/>` : '', node.attrs);
  return `<w:p>${pPr}${prefixRun}${renderInline(node)}</w:p>`;
}

function renderHeading(node: PMNode): string {
  const level = node.attrs.level as number;
  const pPr = pPrXml(`<w:pStyle w:val="Heading${level}"/>`, node.attrs);
  return `<w:p>${pPr}${renderInline(node)}</w:p>`;
}

/**
 * 목록은 word/numbering.xml 연동 없이(v1 스코프 단순화) "• "/"1. " 접두사가 붙은
 * 일반 문단으로 평탄화해 내보낸다 — Word에서 다시 열어도 텍스트는 그대로 보이지만
 * 실제 목록 서식(자동 번호 매기기 등)은 아니다.
 */
function renderList(node: PMNode, ordered: boolean): string {
  let xml = '';
  let index = 1;
  node.forEach((listItem) => {
    const prefix = ordered ? `${index}. ` : '• ';
    listItem.forEach((child) => {
      xml += child.type.name === 'paragraph' ? renderParagraph(child, prefix) : renderBlock(child);
    });
    index += 1;
  });
  return xml;
}

/** 셀에 colwidth(px)가 없을 때(예: 편집기에서 새로 만든 표)의 기본 셀 너비 */
const CELL_WIDTH_DXA = 2000;
/** CSS px 기준 dxa(=twip, 1/20pt) 환산 — 96px = 1440dxa(1인치). src/docx/import.ts와 대칭. */
const DXA_PER_PX = 15;

function pxToDxa(px: number): number {
  return Math.round(px * DXA_PER_PX);
}

/** 편집기에서 새로 만든 표처럼 셀에 borders 속성이 아예 없을 때만 쓰이는 표 기본 테두리 —
 * import.ts에서 가져온 셀은 항상 자기 borders를 명시적으로 갖고 있어 이 기본값을 안 쓴다. */
const BORDER = '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>' +
  '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>';

const CSS_BORDER_STYLE_TO_OOXML: Record<string, string> = {
  solid: 'single',
  double: 'double',
  dashed: 'dashed',
  dotted: 'dotted',
};

/** CellBorderSide 한 변을 w:top/w:left/w:bottom/w:right 중 하나로 직렬화한다. null이면 "테두리 없음"(w:val="nil"). */
function renderBorderSideXml(tag: string, side: CellBorderSide | null): string {
  if (!side) return `<w:${tag} w:val="nil"/>`;
  const val = CSS_BORDER_STYLE_TO_OOXML[side.style] ?? 'single';
  // widthPx = sz/6 (import.ts와 대칭) → sz = widthPx*6, 최소 2(1/4pt)는 되게 한다.
  const sz = Math.max(2, Math.round(side.widthPx * 6));
  const color = side.color.replace('#', '').toUpperCase();
  return `<w:${tag} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color}"/>`;
}

/** 셀의 borders 속성(4변 모두 확정된 값)을 <w:tcBorders>로 직렬화한다. 속성이 없으면 빈 문자열(표 기본 테두리를 그대로 따른다). */
function renderCellBordersXml(borders: unknown): string {
  if (!borders || typeof borders !== 'object') return '';
  const b = borders as CellBorders;
  return (
    '<w:tcBorders>' +
    renderBorderSideXml('top', b.top) +
    renderBorderSideXml('left', b.left) +
    renderBorderSideXml('bottom', b.bottom) +
    renderBorderSideXml('right', b.right) +
    '</w:tcBorders>'
  );
}

/** 셀의 cellPadding 속성(4변 모두 확정된 px 값)을 <w:tcMar>로 직렬화한다. 없으면 빈 문자열(표/스타일 기본 여백을 그대로 따른다). */
function renderCellMarginXml(padding: unknown): string {
  if (!padding || typeof padding !== 'object') return '';
  const p = padding as CellPadding;
  const side = (tag: string, px: number) => `<w:${tag} w:w="${pxToDxa(px)}" w:type="dxa"/>`;
  return `<w:tcMar>${side('top', p.topPx)}${side('left', p.leftPx)}${side('bottom', p.bottomPx)}${side('right', p.rightPx)}</w:tcMar>`;
}

/** 세로 병합 중(pending) 열의 상태 — 몇 행 더 w:vMerge 이어짐 칸을 채워야 하는지, 그 칸의 가로폭(colspan)·너비(dxa) */
interface PendingVMerge {
  remaining: number;
  readonly colspan: number;
  readonly widthDxa: number;
}

/**
 * ProseMirror 표 한 행을 docx 행으로 직렬화한다. prosemirror-tables는 세로 병합을 "이어지는
 * 칸은 그 행의 셀 목록에 아예 없다"는 방식으로 표현하므로(src/docx/import.ts의 tableFrom과
 * 동일한 모델), 이 행의 실제 셀들 사이사이에 이전 행에서 시작된 세로 병합(pendingVMerge)이
 * 남아 있는 열이 있으면 그 자리에 w:vMerge 이어짐 칸을 끼워 넣어 열 정렬을 맞춘다. 실제 셀과
 * pendingVMerge 칸은 열 위치가 절대 겹치지 않으므로(겹치면 애초에 이 행에 그 열의 실제 셀이
 * 없었을 것이다) 열 번호를 따로 계산할 필요 없이 앞에서부터 순서대로 배치하면 된다.
 */
/** table_row의 height 속성(atLeast/exact, px)을 <w:trPr><w:trHeight .../></w:trPr>로 직렬화한다. 없으면 빈 문자열(자동 높이). */
function renderRowHeightXml(height: unknown): string {
  if (!height || typeof height !== 'object') return '';
  const h = height as RowHeight;
  const hRule = h.rule === 'exact' ? 'exact' : 'atLeast';
  return `<w:trPr><w:trHeight w:val="${pxToDxa(h.px)}" w:hRule="${hRule}"/></w:trPr>`;
}

function renderTableRow(
  row: PMNode,
  pendingVMerge: Map<number, PendingVMerge>,
  columnWidthsDxa: Map<number, number>,
): string {
  const cells: PMNode[] = [];
  row.forEach((cell) => cells.push(cell));

  const trPrXml = renderRowHeightXml(row.attrs.height);
  let cellsXml = '';
  let col = 0;
  let cellIndex = 0;

  while (cellIndex < cells.length || pendingVMerge.has(col)) {
    const pending = pendingVMerge.get(col);
    if (pending) {
      const gridSpanXml = pending.colspan > 1 ? `<w:gridSpan w:val="${pending.colspan}"/>` : '';
      cellsXml +=
        `<w:tc><w:tcPr><w:tcW w:w="${pending.widthDxa}" w:type="dxa"/>` +
        `${gridSpanXml}<w:vMerge/></w:tcPr><w:p/></w:tc>`;
      pending.remaining -= 1;
      if (pending.remaining === 0) pendingVMerge.delete(col);
      col += pending.colspan;
      continue;
    }

    const cell = cells[cellIndex];
    cellIndex += 1;
    const colspan = Number(cell.attrs.colspan) || 1;
    const rowspan = Number(cell.attrs.rowspan) || 1;
    const colwidthPx = Array.isArray(cell.attrs.colwidth) ? (cell.attrs.colwidth as number[]) : null;
    // 열별 너비(colwidth)가 있으면 그대로 쓰고(열마다 tblGrid에 반영), 없으면(예: 편집기에서
    // 새로 만든 셀) 기본 너비를 colspan칸에 균등 배분한다.
    const perColumnDxa = colwidthPx
      ? colwidthPx.map(pxToDxa)
      : Array.from({ length: colspan }, () => CELL_WIDTH_DXA);
    const totalWidthDxa = perColumnDxa.reduce((sum, w) => sum + w, 0);
    perColumnDxa.forEach((w, i) => columnWidthsDxa.set(col + i, w));

    let cellBody = '';
    cell.forEach((block) => {
      cellBody += renderBlock(block);
    });
    const gridSpanXml = colspan > 1 ? `<w:gridSpan w:val="${colspan}"/>` : '';
    const vMergeXml = rowspan > 1 ? '<w:vMerge w:val="restart"/>' : '';
    const background = typeof cell.attrs.background === 'string' ? (cell.attrs.background as string) : null;
    const shdXml = background
      ? `<w:shd w:val="clear" w:color="auto" w:fill="${background.replace('#', '').toUpperCase()}"/>`
      : '';
    const bordersXml = renderCellBordersXml(cell.attrs.borders);
    const verticalAlign = cell.attrs.verticalAlign;
    // docx는 top/center/bottom, 스키마·CSS 쪽은 top/middle/bottom을 쓰므로 여기서 되돌린다.
    const vAlignXml =
      verticalAlign === 'middle'
        ? '<w:vAlign w:val="center"/>'
        : verticalAlign === 'top' || verticalAlign === 'bottom'
          ? `<w:vAlign w:val="${verticalAlign}"/>`
          : '';
    const cellMarginXml = renderCellMarginXml(cell.attrs.cellPadding);
    cellsXml +=
      `<w:tc><w:tcPr><w:tcW w:w="${totalWidthDxa}" w:type="dxa"/>` +
      `${gridSpanXml}${vMergeXml}${bordersXml}${shdXml}${vAlignXml}${cellMarginXml}</w:tcPr>${cellBody}</w:tc>`;
    if (rowspan > 1) pendingVMerge.set(col, { remaining: rowspan - 1, colspan, widthDxa: totalWidthDxa });
    col += colspan;
  }

  return `<w:tr>${trPrXml}${cellsXml}</w:tr>`;
}

function renderTable(node: PMNode): string {
  // 표 전체 열 개수 — 세로 병합으로 뒤 행은 실제 셀이 더 적을 수 있어 모든 행 중 최댓값을 쓴다.
  let colCount = 0;
  node.forEach((row) => {
    let width = 0;
    row.forEach((cell) => {
      width += Number(cell.attrs.colspan) || 1;
    });
    colCount = Math.max(colCount, width);
  });

  const pendingVMerge = new Map<number, PendingVMerge>();
  const columnWidthsDxa = new Map<number, number>();
  let rowsXml = '';
  node.forEach((row) => {
    rowsXml += renderTableRow(row, pendingVMerge, columnWidthsDxa);
  });

  const gridCols = Array.from(
    { length: colCount },
    (_, i) => `<w:gridCol w:w="${columnWidthsDxa.get(i) ?? CELL_WIDTH_DXA}"/>`,
  ).join('');

  // 표 자체가 페이지 안에서 어디 놓이는지(셀 안 문단 정렬과는 별개) — 없으면(기본 왼쪽) 생략한다.
  const align = node.attrs.align;
  const jcXml = align === 'center' || align === 'right' ? `<w:jc w:val="${align}"/>` : '';

  return (
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${jcXml}<w:tblBorders>${BORDER}</w:tblBorders></w:tblPr>` +
    `<w:tblGrid>${gridCols}</w:tblGrid>${rowsXml}</w:tbl>`
  );
}

function renderBlock(node: PMNode): string {
  switch (node.type.name) {
    case 'paragraph':
      return renderParagraph(node);
    case 'heading':
      return renderHeading(node);
    case 'bullet_list':
      return renderList(node, false);
    case 'ordered_list':
      return renderList(node, true);
    case 'table':
      return renderTable(node);
    default:
      return '';
  }
}

const DOCUMENT_XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/**
 * sectPrXml을 생략하면(주로 테스트에서) 빈 <w:sectPr/>을 쓴다 — 실제 저장 경로(exportDocx)는
 * 항상 원본 문서에서 뽑아낸 실제 sectPrXml을 넘겨 페이지 크기·여백 등을 보존한다.
 */
export function renderDocumentXml(
  doc: PMNode,
  sectPrXml: string = '<w:sectPr/>',
  styleIds: StyleIdMap = new Map(),
  commentNumbers: Map<string, number> = new Map(),
): string {
  let body = '';
  activeStyleIds = styleIds;
  commentState = { numbers: commentNumbers, open: new Set() };
  try {
    doc.forEach((node) => {
      body += renderBlock(node);
    });
    // 문서 끝까지 열린 채인 메모 범위는 마지막에 빈 문단 하나로 닫는다(참조 실행문은 문단 안에 있어야 한다)
    const state = commentState;
    if (state.open.size > 0) body += `<w:p>${[...state.open].map((id) => closeCommentXml(state, id)).join('')}</w:p>`;
  } finally {
    activeStyleIds = new Map();
    commentState = null;
  }
  // 원본 sectPr를 그대로 쓰므로 그 안의 머리글/바닥글 참조(r:id)를 위해 r 네임스페이스도 선언한다
  return (
    `${DOCUMENT_XML_HEADER}<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${body}${sectPrXml}</w:body></w:document>`
  );
}

/**
 * ProseMirror 문서를 docx로 직렬화한다. 원본 zip의 다른 파트(styles.xml, media 등)는
 * 그대로 보존하고 word/document.xml만 교체한다 — 단, 이 함수가 만드는 document.xml은
 * 스타일 정의를 참조만 할 뿐(Heading1 등) 재정의하지 않으므로, 원본에 해당 스타일이
 * 없으면(예: 빈 스타일시트로 새로 만든 docx) 기본 서식으로 표시될 수 있다.
 *
 * w:sectPr(페이지 크기·여백 등)은 이 스키마가 전혀 모델링하지 않으므로, 매번 새로 만드는 대신
 * 원본 word/document.xml에서 그대로 뽑아 재사용한다 — 그래야 편집·재저장을 거쳐도 페이지
 * 설정이 유지된다(과거엔 항상 빈 <w:sectPr/>로 덮어써서 편집 후 페이지 크기가 사라졌었다).
 */
/** styles.xml을 새로 만들었을 때 — Word가 그 파트를 찾도록 콘텐츠 형식과 문서 관계에 등록한다. */
function registerStylesPart(entries: Record<string, Uint8Array>): void {
  const types = entries['[Content_Types].xml'];
  if (types) {
    const xml = strFromU8(types);
    if (!xml.includes('/word/styles.xml')) {
      entries['[Content_Types].xml'] = strToU8(
        xml.replace(
          '</Types>',
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
        ),
      );
    }
  }
  const rels = entries['word/_rels/document.xml.rels'];
  if (rels) {
    const xml = strFromU8(rels);
    if (!xml.includes('relationships/styles"')) {
      entries['word/_rels/document.xml.rels'] = strToU8(
        xml.replace(
          '</Relationships>',
          '<Relationship Id="rIdDocviewStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        ),
      );
    }
  }
}

/**
 * 원본 w:sectPr에서 용지 크기(pgSz)와 여백(pgMar의 위·아래·왼쪽·오른쪽)만 바꾼다 — 머리글/바닥글
 * 참조, 제본 여백, 단 설정 등 나머지는 원본 그대로 둔다. CT_SectPr 순서상 pgSz·pgMar는
 * 머리글/바닥글 참조·각주 설정·구역 종류 뒤에 온다.
 */
export function applyPageSetupToSectPr(sectPrXml: string, page: DocxPageSetup): string {
  let xml = /^<w:sectPr\b[^>]*\/>$/.test(sectPrXml) ? sectPrXml.replace(/\/>$/, '></w:sectPr>') : sectPrXml;
  const orient = page.orientation === 'landscape' ? ' w:orient="landscape"' : '';
  const pgSz = `<w:pgSz w:w="${Math.round(page.width)}" w:h="${Math.round(page.height)}"${orient}/>`;
  const margins: ReadonlyArray<[string, number]> = [
    ['top', page.top],
    ['right', page.right],
    ['bottom', page.bottom],
    ['left', page.left],
  ];
  const oldPgMar = /<w:pgMar\b[^>]*\/>/.exec(xml)?.[0];
  if (oldPgMar) {
    let pgMar = oldPgMar;
    for (const [side, value] of margins) {
      const pattern = new RegExp(`\\bw:${side}="[^"]*"`);
      const attr = `w:${side}="${Math.round(value)}"`;
      pgMar = pattern.test(pgMar) ? pgMar.replace(pattern, attr) : pgMar.replace('<w:pgMar', `<w:pgMar ${attr}`);
    }
    xml = xml.replace(oldPgMar, pgMar);
  }
  const newPgMar = oldPgMar
    ? ''
    : `<w:pgMar w:top="${Math.round(page.top)}" w:right="${Math.round(page.right)}" w:bottom="${Math.round(page.bottom)}" ` +
      `w:left="${Math.round(page.left)}" w:header="851" w:footer="992" w:gutter="0"/>`;
  if (/<w:pgSz\b[^>]*\/>/.test(xml)) return xml.replace(/<w:pgSz\b[^>]*\/>/, pgSz + newPgMar);
  // pgSz 앞에 와야 하는 요소들(머리글/바닥글 참조, 각주·미주 설정, 구역 종류) 뒤에 넣는다
  const leading =
    /^<w:sectPr\b[^>]*>(?:\s*<w:(?:headerReference|footerReference|type)\b[^>]*\/>|\s*<w:(?:footnotePr|endnotePr)\b[\s\S]*?<\/w:(?:footnotePr|endnotePr)>)*/.exec(
      xml,
    )?.[0] ?? '';
  return leading + pgSz + newPgMar + xml.slice(leading.length);
}

export function exportDocx(
  originalDocxBytes: Uint8Array,
  doc: PMNode,
  page: DocxPageSetup | null = null,
  comments: readonly DocxComment[] = [],
): Uint8Array {
  const entries = guardedUnzipSync(originalDocxBytes, 'exportDocx');
  const originalDocumentXml = entries['word/document.xml'];
  const originalSectPr = (originalDocumentXml && extractSectPrXml(strFromU8(originalDocumentXml))) ?? '<w:sectPr/>';
  const sectPrXml = page ? applyPageSetupToSectPr(originalSectPr, page) : originalSectPr;
  const stylesEntry = entries['word/styles.xml'];
  const baseStyles = stylesEntry ? strFromU8(stylesEntry) : `${DOCUMENT_XML_HEADER}<w:styles xmlns:w="${NS_W}"></w:styles>`;
  const styles = resolveGalleryStyles(doc, baseStyles);
  if (styles.stylesXml !== null && (stylesEntry || styles.ids.size > 0)) {
    entries['word/styles.xml'] = strToU8(styles.stylesXml);
    if (!stylesEntry) registerStylesPart(entries);
  }
  const commentNumbers = new Map<string, number>();
  entries['word/document.xml'] = strToU8(renderDocumentXml(doc, sectPrXml, styles.ids, commentNumbers));
  // 메모가 있거나, 원본에 있던 메모를 모두 지운 경우에만 comments.xml을 새로 쓴다
  if (commentNumbers.size > 0 || entries['word/comments.xml']) {
    entries['word/comments.xml'] = strToU8(renderCommentsXml(commentNumbers, new Map(comments.map((c) => [c.id, c]))));
    registerCommentsPart(entries);
  }
  return zipSync(entries);
}
