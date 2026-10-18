import OrderedMap from 'orderedmap';
import { Schema, type MarkSpec, type Node as PMNode, type NodeSpec } from 'prosemirror-model';
import { addListNodes } from 'prosemirror-schema-list';
import { tableNodes } from 'prosemirror-tables';

/**
 * docx 실시간 동시편집 1차 스코프에 맞춘 최소 스키마.
 * 지원: 문단, 제목(1~6), 굵게/기울임/밑줄/취소선, 표, 글머리·번호 목록(표시만 — 목록은
 * word/numbering.xml 연동 없이 내보낼 때 "• "/"1. " 접두사가 붙은 일반 문단으로 평탄화된다).
 * 미지원(가져오기 시 유실): 이미지, 머리글/바닥글, 각주, 하이퍼링크, 사용자 지정 스타일.
 */
/** docx의 w:ind — 단위는 전부 dxa(twip, 1/20pt). firstLine과 hanging은 둘 중 하나만 0이 아니다. */
export interface ParagraphIndent {
  readonly left: number;
  readonly firstLine: number;
  readonly hanging: number;
  /** 오른쪽 들여쓰기(w:right) — 예전 데이터엔 없을 수 있다. */
  readonly right?: number;
}

/** docx의 w:spacing — before/after는 dxa, line은 lineRule이 auto면 240분의 1줄 단위, 아니면 dxa. */
export interface ParagraphSpacing {
  readonly line: number | null;
  readonly lineRule: 'auto' | 'exact' | 'atLeast' | null;
  readonly before: number | null;
  readonly after: number | null;
}

/** 편집기가 스타일 갤러리로 붙이는 문단 스타일 — 저장 시 원본 styles.xml의 같은 이름 스타일로
 * 연결하거나, 없으면 정의를 추가한다(export.ts의 resolveGalleryStyles). */
export const GALLERY_PARAGRAPH_STYLES = ['NoSpacing', 'Title', 'Subtitle', 'Quote', 'IntenseQuote'] as const;
/** 문서 전체 용지 설정(dxa) — 레이아웃 탭이 고친 값만 Y.Map(DOCX_PAGE_MAP)의 'setup'에 담긴다.
 * 저장 시 원본 w:sectPr의 pgSz/pgMar 값만 이것으로 바꾼다(머리글 참조 등 나머지는 그대로). */
export interface DocxPageSetup {
  readonly width: number;
  readonly height: number;
  readonly orientation: 'portrait' | 'landscape';
  readonly top: number;
  readonly bottom: number;
  readonly left: number;
  readonly right: number;
}
export const DOCX_PAGE_MAP = 'docxPage';
export const DOCX_PAGE_SETUP_KEY = 'setup';

export type GalleryStyleId = (typeof GALLERY_PARAGRAPH_STYLES)[number];
/** 갤러리 스타일의 Word 내장 이름(styles.xml w:name, 소문자) — 한국어 Word는 id가 달라도 이름은 같다. */
export const GALLERY_STYLE_NAMES: Readonly<Record<GalleryStyleId, string>> = {
  NoSpacing: 'no spacing',
  Title: 'title',
  Subtitle: 'subtitle',
  Quote: 'quote',
  IntenseQuote: 'intense quote',
};

const TWIPS_PER_PX = 15;
/** Word의 "한 줄"(240) 간격이 화면 글꼴에서 차지하는 높이 비율 — 맑은 고딕 기준 근삿값. */
const WORD_SINGLE_LINE_HEIGHT = 1.2;

function parseJsonAttr<T>(dom: HTMLElement, name: string): T | null {
  const raw = dom.getAttribute(name);
  if (!raw) return null;
  try {
    return JSON.parse(raw) as T;
  } catch {
    return null;
  }
}

/** paragraph/heading의 공통 문단 속성 — DOM에는 인라인 스타일로 보이게 하고, 붙여넣기·복원용으로
 * data-* 속성에 원래 값을 남긴다. docx로는 w:pPr(shd·spacing·ind·jc)로 나간다. */
function paragraphDOMAttrs(attrs: Record<string, unknown>): Record<string, string> {
  const styles: string[] = [];
  const out: Record<string, string> = {};
  if (attrs.align) styles.push(`text-align: ${attrs.align as string}`);
  const indent = attrs.indent as ParagraphIndent | null;
  if (indent) {
    if (indent.left) styles.push(`margin-left: ${indent.left / TWIPS_PER_PX}px`);
    if (indent.right) styles.push(`margin-right: ${indent.right / TWIPS_PER_PX}px`);
    const textIndent = indent.firstLine ? indent.firstLine : indent.hanging ? -indent.hanging : 0;
    if (textIndent) styles.push(`text-indent: ${textIndent / TWIPS_PER_PX}px`);
    out['data-indent'] = JSON.stringify(indent);
  }
  const spacing = attrs.spacing as ParagraphSpacing | null;
  if (spacing) {
    if (spacing.before !== null) styles.push(`margin-top: ${spacing.before / TWIPS_PER_PX}px`);
    if (spacing.after !== null) styles.push(`margin-bottom: ${spacing.after / TWIPS_PER_PX}px`);
    if (spacing.line) {
      styles.push(
        spacing.lineRule === 'exact' || spacing.lineRule === 'atLeast'
          ? `line-height: ${spacing.line / TWIPS_PER_PX}px`
          : `line-height: ${Number(((spacing.line / 240) * WORD_SINGLE_LINE_HEIGHT).toFixed(3))}`,
      );
    }
    out['data-spacing'] = JSON.stringify(spacing);
  }
  if (attrs.shading) {
    styles.push(`background-color: ${attrs.shading as string}`);
    out['data-shading'] = attrs.shading as string;
  }
  if (attrs.styleId) out['data-style'] = attrs.styleId as string;
  if (styles.length > 0) out.style = styles.join('; ');
  return out;
}

function paragraphAttrsFromDOM(dom: HTMLElement): Record<string, unknown> {
  const shading = dom.getAttribute('data-shading');
  return {
    align: dom.style.textAlign || null,
    indent: parseJsonAttr<ParagraphIndent>(dom, 'data-indent'),
    spacing: parseJsonAttr<ParagraphSpacing>(dom, 'data-spacing'),
    shading: shading && /^#[0-9a-fA-F]{6}$/.test(shading) ? shading : null,
  };
}

const PARAGRAPH_FORMAT_ATTRS = {
  align: { default: null },
  indent: { default: null },
  spacing: { default: null },
  shading: { default: null },
};

const baseNodes: Record<string, NodeSpec> = {
  doc: { content: 'block+' },

  paragraph: {
    content: 'inline*',
    group: 'block',
    /** styleId: 원본의 w:pStyle(제목 외 — 저장 시 그대로 되돌려 원본 스타일을 지킨다) 또는 갤러리 스타일 */
    attrs: { ...PARAGRAPH_FORMAT_ATTRS, styleId: { default: null } },
    parseDOM: [
      {
        tag: 'p',
        getAttrs: (dom) => ({
          ...paragraphAttrsFromDOM(dom as HTMLElement),
          styleId: (dom as HTMLElement).getAttribute('data-style') || null,
        }),
      },
    ],
    toDOM: (node) => ['p', paragraphDOMAttrs(node.attrs), 0],
  },

  heading: {
    attrs: { level: { default: 1 }, ...PARAGRAPH_FORMAT_ATTRS },
    content: 'inline*',
    group: 'block',
    defining: true,
    parseDOM: [1, 2, 3, 4, 5, 6].map((level) => ({
      tag: `h${level}`,
      getAttrs: (dom) => ({ level, ...paragraphAttrsFromDOM(dom as HTMLElement) }),
    })),
    toDOM: (node) => [`h${node.attrs.level as number}`, paragraphDOMAttrs(node.attrs), 0],
  },

  text: { group: 'inline' },

  /** docx의 w:br(문단 안 줄바꿈) — 새 문단이 아니라 같은 문단 안에서 줄만 바꾸는 것. */
  hard_break: {
    inline: true,
    group: 'inline',
    selectable: false,
    parseDOM: [{ tag: 'br' }],
    toDOM: () => ['br'],
  },

  /** docx의 w:br w:type="page" — 페이지 나누기. 편집 화면엔 점선 줄로 보인다. */
  page_break: {
    inline: true,
    group: 'inline',
    atom: true,
    parseDOM: [{ tag: 'span.page-break' }],
    toDOM: () => ['span', { class: 'page-break', contenteditable: 'false' }],
  },
};

/** Word 텍스트 강조색(w:highlight)의 정해진 색 이름 → '#rrggbb'. 이 밖의 색은 w:shd로 저장한다. */
export const HIGHLIGHT_COLORS: Readonly<Record<string, string>> = {
  yellow: '#ffff00',
  green: '#00ff00',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  blue: '#0000ff',
  red: '#ff0000',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
  black: '#000000',
  white: '#ffffff',
};

/** 글꼴 이름을 CSS 문자열 안에 넣어도 안전하게 — 따옴표·세미콜론·꺾쇠·역슬래시는 글꼴 이름에 쓰이지 않는다. */
export function safeFontName(value: unknown): string {
  return String(value ?? '').replace(/["';<>\\]/g, '').trim();
}

const marks: Record<string, MarkSpec> = {
  strong: {
    parseDOM: [{ tag: 'strong' }, { tag: 'b' }],
    toDOM: () => ['strong', 0],
  },
  em: {
    parseDOM: [{ tag: 'em' }, { tag: 'i' }],
    toDOM: () => ['em', 0],
  },
  underline: {
    parseDOM: [{ tag: 'u' }],
    toDOM: () => ['u', 0],
  },
  strike: {
    parseDOM: [{ tag: 's' }],
    toDOM: () => ['s', 0],
  },
  textColor: {
    attrs: { color: {} },
    parseDOM: [
      {
        style: 'color',
        getAttrs: (value) => ({ color: value as string }),
      },
    ],
    toDOM: (mark) => ['span', { style: `color: ${mark.attrs.color as string}` }, 0],
  },
  /** w:rFonts — 한글(eastAsia)과 영문·숫자(ascii) 글꼴을 따로 갖는다(Word와 같다). 리본 글꼴 상자로
   * 고르면 둘 다 같은 이름이 된다. 화면에는 한글 글꼴을 먼저 쓴다. */
  fontFamily: {
    attrs: { ascii: { default: null }, eastAsia: { default: null } },
    parseDOM: [
      {
        tag: 'span[data-font-ascii], span[data-font-east-asia]',
        getAttrs: (dom) => ({
          ascii: (dom as HTMLElement).getAttribute('data-font-ascii') || null,
          eastAsia: (dom as HTMLElement).getAttribute('data-font-east-asia') || null,
        }),
      },
    ],
    toDOM: (mark) => {
      const ascii = mark.attrs.ascii ? safeFontName(mark.attrs.ascii) : '';
      const eastAsia = mark.attrs.eastAsia ? safeFontName(mark.attrs.eastAsia) : '';
      const names = [eastAsia, ascii].filter(Boolean).map((name) => `"${name}"`);
      const attrs: Record<string, string> = { style: `font-family: ${[...names, 'var(--doc-font-family, sans-serif)'].join(', ')}` };
      if (ascii) attrs['data-font-ascii'] = ascii;
      if (eastAsia) attrs['data-font-east-asia'] = eastAsia;
      return ['span', attrs, 0];
    },
  },
  subscript: {
    excludes: 'superscript',
    parseDOM: [{ tag: 'sub' }],
    toDOM: () => ['sub', 0],
  },
  superscript: {
    excludes: 'subscript',
    parseDOM: [{ tag: 'sup' }],
    toDOM: () => ['sup', 0],
  },
  /** 텍스트 강조색 — docx의 w:highlight(정해진 색 이름) 또는 w:shd(임의 색). '#rrggbb'로 저장한다. */
  highlight: {
    attrs: { color: {} },
    parseDOM: [
      { tag: 'mark[data-highlight]', getAttrs: (dom) => ({ color: (dom as HTMLElement).getAttribute('data-highlight') }) },
    ],
    toDOM: (mark) => [
      'mark',
      { 'data-highlight': mark.attrs.color as string, style: `background-color: ${mark.attrs.color as string}; color: inherit` },
      0,
    ],
  },
  fontSize: {
    /** 포인트(pt) 단위로 저장 — docx의 w:sz(하프포인트)와는 export/import 시 ×2/÷2 변환한다 */
    attrs: { size: {} },
    parseDOM: [
      {
        style: 'font-size',
        getAttrs: (value) => ({ size: parseFloat(value as string) }),
      },
    ],
    toDOM: (mark) => ['span', { style: `font-size: ${mark.attrs.size as number}pt` }, 0],
  },
};

/**
 * docx의 w:sectPr/w:pgSz·w:pgMar — 페이지 크기·여백. import.ts가 원본에서 그대로 읽어오고,
 * export.ts는 이 스키마와 무관하게 원본 w:sectPr 자체를 그대로 보존해 내보낸다(우리가
 * 모델링하지 않는 세부 설정까지 안전하게 되돌아가도록). 여기 있는 값은 오직 "화면에 실제
 * Word 페이지와 같은 비율로 얼마나 넓게 그릴지" 계산에만 쓰인다.
 */
export interface PageSetup {
  readonly widthPx: number;
  readonly heightPx: number;
  readonly marginTopPx: number;
  readonly marginBottomPx: number;
  readonly marginLeftPx: number;
  readonly marginRightPx: number;
  readonly orientation: 'portrait' | 'landscape';
}

/** docx의 w:trPr/w:trHeight — 행 높이. rule이 'auto'면 이 속성 자체가 없음(null)과 동일하게 다룬다. */
export interface RowHeight {
  readonly rule: 'atLeast' | 'exact';
  readonly px: number;
}

function rowHeightToDOMStyle(height: unknown): Record<string, string> {
  const h = height as RowHeight | null;
  return h ? { style: `height: ${h.px}px;` } : {};
}

function rowHeightFromDOM(dom: HTMLElement): RowHeight | null {
  const px = parseFloat(dom.style.height);
  return Number.isFinite(px) && px > 0 ? { rule: 'atLeast', px } : null;
}

/** docx의 w:tblPr/w:jc — 표 자체가 페이지(콘텐츠 영역) 안에서 왼쪽/가운데/오른쪽 중 어디에 놓이는지.
 * 셀 안 문단 정렬(paragraph의 align)과는 완전히 별개 개념이다. */
function tableAlignToDOMStyle(align: unknown): Record<string, string> {
  if (align === 'center') return { style: 'margin-left: auto; margin-right: auto;' };
  if (align === 'right') return { style: 'margin-left: auto; margin-right: 0;' };
  return {};
}

/**
 * 표 자체의 실제 너비(px) — 첫 행 각 셀의 colwidth(모두 있을 때만) 합. 하나라도 없으면
 * (예: 편집기에서 새로 만든 표) null을 돌려주고, CSS 기본값(width:100%)에 맡긴다.
 * Word 원본 표가 페이지 폭보다 좁을 때 억지로 100%까지 늘어나 셀 너비가 실제보다 커지는
 * 문제(사용자가 "Cell 너비가 Word 원본과 다르다"고 지적한 부분)를 막기 위함이다.
 */
function tableWidthPx(node: PMNode): number | null {
  const firstRow = node.firstChild;
  if (!firstRow) return null;
  let total = 0;
  let ok = true;
  firstRow.forEach((cell) => {
    const colwidth = cell.attrs.colwidth as number[] | null;
    if (!Array.isArray(colwidth) || colwidth.length === 0) {
      ok = false;
      return;
    }
    total += colwidth.reduce((sum, w) => sum + w, 0);
  });
  return ok && total > 0 ? total : null;
}

/** 셀 한 변의 테두리 — docx의 w:tcBorders/w:tblBorders 한 변(top/right/bottom/left)에 대응 */
export interface CellBorderSide {
  /** CSS border-style 값(solid/dashed/dotted/double) — docx w:val(single/dashed/dotted/double 등)에서 변환 */
  readonly style: string;
  readonly widthPx: number;
  /** "#rrggbb" */
  readonly color: string;
}

export interface CellBorders {
  readonly top: CellBorderSide | null;
  readonly right: CellBorderSide | null;
  readonly bottom: CellBorderSide | null;
  readonly left: CellBorderSide | null;
}

const BORDER_SIDES = ['top', 'right', 'bottom', 'left'] as const;

function borderSideToCSS(side: CellBorderSide | null): string {
  return side ? `${side.widthPx}px ${side.style} ${side.color}` : 'none';
}

/** "1.5px solid rgb(255, 0, 0)" 같은 계산된 CSS 값을 다시 CellBorderSide로 되돌린다(HTML 붙여넣기 등 DOM 기반 입력용). */
function borderSideFromDOM(cssValue: string): CellBorderSide | null {
  if (!cssValue || cssValue === 'none') return null;
  const match = /^([\d.]+)px\s+(\S+)\s+(.+)$/.exec(cssValue.trim());
  if (!match) return null;
  const [, widthStr, style] = match;
  return { style, widthPx: parseFloat(widthStr), color: match[3] };
}

const VALID_VERTICAL_ALIGN = new Set(['top', 'middle', 'bottom']);

/** 셀 안쪽 여백(docx의 w:tblCellMar/w:tcMar에서 계산된 최종값, px) — 4변 모두 확정된 값이다. */
export interface CellPadding {
  readonly topPx: number;
  readonly rightPx: number;
  readonly bottomPx: number;
  readonly leftPx: number;
}

function cellPaddingToCSS(padding: CellPadding): string {
  return `${padding.topPx}px ${padding.rightPx}px ${padding.bottomPx}px ${padding.leftPx}px`;
}

/** 계산된 CSS padding 값("Tpx Rpx Bpx Lpx")을 다시 CellPadding으로 되돌린다(HTML 붙여넣기 등 DOM 기반 입력용). */
function cellPaddingFromDOM(cssValue: string): CellPadding | null {
  if (!cssValue) return null;
  const parts = cssValue.trim().split(/\s+/).map(parseFloat);
  if (parts.length !== 4 || parts.some((v) => Number.isNaN(v))) return null;
  const [topPx, rightPx, bottomPx, leftPx] = parts;
  return { topPx, rightPx, bottomPx, leftPx };
}

const tableNodeSpecs = tableNodes({
  tableGroup: 'block',
  cellContent: 'block+',
  cellAttributes: {
    /** 셀 음영(docx의 w:tcPr/w:shd) — "#rrggbb" 또는 없음(null) */
    background: {
      default: null,
      getFromDOM: (dom) => (dom as HTMLElement).style.backgroundColor || null,
      setDOMAttr: (value, attrs) => {
        if (value) attrs.style = `${(attrs.style as string | undefined) ?? ''}background-color: ${value as string};`;
      },
    },
    /** 셀 4변 테두리(docx의 w:tcBorders/w:tblBorders에서 계산된 최종값) — 없으면 null(에디터 기본 테두리 사용) */
    borders: {
      default: null,
      getFromDOM: (dom) => {
        const style = (dom as HTMLElement).style;
        const borders: CellBorders = {
          top: borderSideFromDOM(style.borderTop),
          right: borderSideFromDOM(style.borderRight),
          bottom: borderSideFromDOM(style.borderBottom),
          left: borderSideFromDOM(style.borderLeft),
        };
        return borders.top || borders.right || borders.bottom || borders.left ? borders : null;
      },
      setDOMAttr: (value, attrs) => {
        if (!value) return;
        const borders = value as CellBorders;
        const extra = BORDER_SIDES.map((side) => `border-${side}: ${borderSideToCSS(borders[side])};`).join(' ');
        attrs.style = `${(attrs.style as string | undefined) ?? ''}${extra}`;
      },
    },
    /** 셀 안 내용의 세로 정렬(docx의 w:tcPr/w:vAlign) — 문단의 가로 정렬(align)과는 별개 축이다 */
    verticalAlign: {
      default: null,
      getFromDOM: (dom) => {
        const v = (dom as HTMLElement).style.verticalAlign;
        return v === 'top' || v === 'middle' || v === 'bottom' ? v : null;
      },
      setDOMAttr: (value, attrs) => {
        if (value && VALID_VERTICAL_ALIGN.has(value as string)) {
          attrs.style = `${(attrs.style as string | undefined) ?? ''}vertical-align: ${value as string};`;
        }
      },
    },
    /**
     * 셀 안쪽 여백(docx의 w:tblCellMar/w:tcMar) — CSS 기본값(0.4rem 0.6rem)은 모든 셀에
     * 똑같이 적용되는 임의값이라, 원본이 이보다 좁은 여백을 쓰면 그 차이만큼 텍스트가 덜
     * 들어가 원본에선 한 줄인 게 편집기에선 줄바꿈되는 원인이 된다. 없으면 null(CSS 기본값 사용).
     */
    cellPadding: {
      default: null,
      getFromDOM: (dom) => cellPaddingFromDOM((dom as HTMLElement).style.padding),
      setDOMAttr: (value, attrs) => {
        if (value) attrs.style = `${(attrs.style as string | undefined) ?? ''}padding: ${cellPaddingToCSS(value as CellPadding)};`;
      },
    },
  },
});

/**
 * prosemirror-tables가 만들어주는 table/table_row 스펙은 attrs가 아예 없다(셀 안쪽만 다룸) —
 * "표 자체가 페이지 어디에 놓이는지(정렬)"와 "행 높이"는 표/행 자신의 속성이라 별도로 얹는다.
 */
const patchedTableNodeSpecs: Record<string, NodeSpec> = {
  ...tableNodeSpecs,
  table: {
    ...tableNodeSpecs.table,
    attrs: { align: { default: null } },
    toDOM: (node) => {
      const widthPx = tableWidthPx(node);
      const alignStyle = tableAlignToDOMStyle(node.attrs.align).style ?? '';
      const widthStyle = widthPx ? `width: ${widthPx}px;` : '';
      const style = `${widthStyle}${alignStyle}`;
      return ['table', style ? { style } : {}, ['tbody', 0]];
    },
  },
  table_row: {
    ...tableNodeSpecs.table_row,
    attrs: { height: { default: null } },
    parseDOM: [{ tag: 'tr', getAttrs: (dom) => ({ height: rowHeightFromDOM(dom as HTMLElement) }) }],
    toDOM: (node) => ['tr', rowHeightToDOMStyle(node.attrs.height), 0],
  },
};

const withLists = addListNodes(OrderedMap.from(baseNodes), 'paragraph block*', 'block');
const withTables = withLists.append(patchedTableNodeSpecs);

export const docxSchema = new Schema({ nodes: withTables, marks });
