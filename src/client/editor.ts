import { baseKeymap, chainCommands, selectAll, toggleMark } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import { EditorState, Plugin, PluginKey, Selection, TextSelection, type Command } from 'prosemirror-state';
import { Decoration, DecorationSet, EditorView } from 'prosemirror-view';
import { liftListItem, sinkListItem, splitListItem, wrapInList } from 'prosemirror-schema-list';
import {
  addColumnAfter,
  addColumnBefore,
  addRowAfter,
  addRowBefore,
  columnResizing,
  deleteColumn,
  deleteRow,
  deleteTable,
  goToNextCell,
  isInTable,
  mergeCells,
  splitCell,
  TableView,
  tableEditing,
} from 'prosemirror-tables';
import type { Node as PMNode } from 'prosemirror-model';
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { redo, undo, ySyncPlugin, yCursorPlugin, yUndoPlugin, yUndoPluginKey } from 'y-prosemirror';
import {
  DOCX_COMMENTS_MAP,
  DOCX_PAGE_MAP,
  DOCX_PAGE_SETUP_KEY,
  docxSchema,
  type DocxComment,
  type DocxPageSetup,
  type PageSetup,
} from '../docx/schema.js';
import { setupWordReview } from './wordReview.js';
import { pickPresenceColorForName, renderPresenceAvatars } from './presence.js';
import { DOC_META_MAP, DOC_META_SAVED_AT_KEY, DOC_META_SAVE_ERROR_AT_KEY } from '../collab/gridShared.js';
import { initOfficeShell } from './officeShell.js';
import { bindDropdown, fillColorPalette, type PaletteColor } from './ribbonMenus.js';
import {
  applyMark,
  changeCase,
  changeIndent,
  clearFormatting,
  clearMark,
  copyFormat,
  currentFontName,
  currentFontSize,
  currentParagraphBox,
  currentParagraphStyle,
  insertPageBreak,
  setIndentSide,
  setSpacingSide,
  multilevelList,
  pasteFormat,
  replaceAll,
  replaceRange,
  setAlign,
  setHeading,
  setLineSpacing,
  setParagraphStyle,
  setShading,
  stepFontSize,
  type CaseMode,
  type CopiedFormat,
} from './wordCommands.js';

/** 레이아웃 탭 용지 크기(dxa, 세로 기준) */
const PAPER_SIZES = {
  A4: { width: 11906, height: 16838 },
  A3: { width: 16838, height: 23811 },
  B5: { width: 10319, height: 14572 },
  B4: { width: 14572, height: 20639 },
  Letter: { width: 12240, height: 15840 },
  Legal: { width: 12240, height: 20160 },
} as const;
/** 레이아웃 탭 여백 미리 설정(dxa, 1cm = 567) */
const MARGIN_PRESETS = {
  normal: { top: 1440, bottom: 1440, left: 1440, right: 1440 },
  korean: { top: 1701, bottom: 1440, left: 1440, right: 1440 },
  narrow: { top: 720, bottom: 720, left: 720, right: 720 },
  moderate: { top: 1440, bottom: 1440, left: 1080, right: 1080 },
  wide: { top: 1440, bottom: 1440, left: 2880, right: 2880 },
} as const;

/** Word 텍스트 강조 색 15가지(이름 있는 강조색 — 저장 시 w:highlight) */
const HIGHLIGHT_PALETTE: readonly PaletteColor[] = [
  { hex: '#ffff00', label: '노랑' }, { hex: '#00ff00', label: '밝은 녹색' }, { hex: '#00ffff', label: '청록' },
  { hex: '#ff00ff', label: '분홍' }, { hex: '#0000ff', label: '파랑' }, { hex: '#ff0000', label: '빨강' },
  { hex: '#000080', label: '진한 파랑' }, { hex: '#008080', label: '진한 청록' }, { hex: '#008000', label: '녹색' },
  { hex: '#800080', label: '보라' }, { hex: '#800000', label: '진한 빨강' }, { hex: '#808000', label: '진한 노랑' },
  { hex: '#808080', label: '회색 50%' }, { hex: '#c0c0c0', label: '회색 25%' }, { hex: '#000000', label: '검정' },
];
/** 글꼴 색 — 테마 색 한 줄 + 표준 색 한 줄 */
const FONT_COLOR_PALETTE: readonly PaletteColor[] = [
  { hex: '#000000', label: '검정' }, { hex: '#404040', label: '진한 회색' }, { hex: '#7f7f7f', label: '회색' },
  { hex: '#a6a6a6', label: '연한 회색' }, { hex: '#ffffff', label: '흰색' }, { hex: '#44546a', label: '파랑 회색' },
  { hex: '#4472c4', label: '파랑, 강조 1' }, { hex: '#ed7d31', label: '주황, 강조 2' }, { hex: '#70ad47', label: '녹색, 강조 6' },
  { hex: '#5b9bd5', label: '파랑, 강조 5' },
  { hex: '#c00000', label: '진한 빨강' }, { hex: '#ff0000', label: '빨강' }, { hex: '#ffc000', label: '주황' },
  { hex: '#ffff00', label: '노랑' }, { hex: '#92d050', label: '연한 녹색' }, { hex: '#00b050', label: '녹색' },
  { hex: '#00b0f0', label: '연한 파랑' }, { hex: '#0070c0', label: '파랑' }, { hex: '#002060', label: '진한 파랑' },
  { hex: '#7030a0', label: '자주' },
];
/** 문단 음영 — 회색 단계와 연한 테마 색 */
const SHADING_PALETTE: readonly PaletteColor[] = [
  { hex: '#f2f2f2', label: '흰색, 5% 더 어둡게' }, { hex: '#d9d9d9', label: '흰색, 15% 더 어둡게' }, { hex: '#bfbfbf', label: '흰색, 25% 더 어둡게' },
  { hex: '#a6a6a6', label: '흰색, 35% 더 어둡게' }, { hex: '#808080', label: '흰색, 50% 더 어둡게' },
  { hex: '#dae3f3', label: '파랑, 80% 더 밝게' }, { hex: '#fbe5d6', label: '주황, 80% 더 밝게' }, { hex: '#e2f0d9', label: '녹색, 80% 더 밝게' },
  { hex: '#fff2cc', label: '금색, 80% 더 밝게' }, { hex: '#deebf7', label: '연한 파랑, 80% 더 밝게' },
];

interface Permissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  readonly canCopy: boolean;
  readonly canEdit: boolean;
}

interface ScriptFontFamily {
  readonly eastAsia: string | null;
  readonly ascii: string | null;
}

interface ViewInfo {
  readonly docId: string;
  readonly name: string;
  readonly viewer: string;
  readonly permissions: Permissions;
  readonly editHasUnsupportedContent: boolean;
  readonly pageSetup: PageSetup | null;
  readonly defaultFontSizePt: number | null;
  readonly defaultFontFamily: ScriptFontFamily | null;
}

const params = new URLSearchParams(location.search);
const token = params.get('token') ?? '';

/** false면 이 문서는 읽기 전용이다(동시편집 권한 없는 링크) — grid-editor.ts의 같은 이름 변수와
 * 같은 역할. main()에서 서버가 내려준 토큰 권한으로 설정된다. ProseMirror의 EditorView가
 * `editable: () => canEdit`로 이 값을 직접 참조해 키보드/붙여넣기/드롭 등 DOM에서 오는 편집
 * 이벤트 전체를 스스로 거부하고, 여기서는 그 위에 툴바(서식 리본) 자체를 숨겨 애초에 사용자가
 * 고칠 수 있다고 착각하지 않게 한다. 실제 쓰기 차단은 여전히 서버(collab/room.ts의
 * readOnlyConns)가 독립적으로 수행한다.
 *
 * 주의: `editable:false`는 ProseMirror 코어의 `editHandlers` 목록(keydown/paste/cut/drop/
 * composition류)에 있는 DOM 이벤트만 막는다 — view.dispatch()를 직접 부르는 경로(툴바 버튼 등)는
 * 이 값을 아예 안 본다(그래서 리본을 숨기는 게 장식이 아니라 실질적 방어다). 표 열 너비 드래그처럼
 * editHandlers 밖의 handleDOMEvents 경로는 현재 prosemirror-tables 자신이 내부에서 view.editable을
 * 확인해 막아주고 있을 뿐, ProseMirror 코어가 보장하는 동작이 아니다 — 이 라이브러리를 업그레이드
 * 하거나 handleDOMEvents를 쓰는 새 플러그인을 추가할 때는 그 플러그인이 editable을 스스로
 * 확인하는지 다시 확인할 것. 어느 경우든 서버 쪽 readOnlyConns 차단이 최종 방어선이라 이 가정이
 * 깨져도 실제 쓰기가 반영되지는 않는다. */
let canEdit = true;

const titleEl = document.getElementById('doc-title') as HTMLElement;
const statusEl = document.getElementById('conn-status') as HTMLElement;
const readonlyBadgeEl = document.getElementById('readonly-badge') as HTMLElement;
const presenceEl = document.getElementById('presence-list') as HTMLElement;
const warningEl = document.getElementById('unsupported-warning') as HTMLElement;
const editorHost = document.getElementById('editor') as HTMLElement;
const errorScreen = document.getElementById('error-screen') as HTMLElement;
const errorMessage = document.getElementById('error-message') as HTMLElement;
const toolbar = document.getElementById('toolbar') as HTMLElement;

/**
 * Word 원본의 실제 페이지 크기·여백으로 편집 화면의 "종이" 크기를 그린다 — 기존엔 항상
 * 고정된 CSS 값(860px, 65ch 등 문서 크기와 무관한 임의 수치)을 썼는데, 그 임의값이 곧 "용지
 * 폭이 실제보다 좁아 보인다"/"셀 너비가 원본과 다르다"는 문제의 원인이었다. 페이지 전체 폭을
 * 그대로 쓰고 여백을 padding으로 재현하면, 안쪽 콘텐츠 폭(=페이지 폭-여백)이 Word의 실제
 * 텍스트 영역 폭과 정확히 같아진다 — 표가 그 안에서 백분율이 아니라 자기 실제 폭(px)으로
 * 그려지므로(schema.ts의 table toDOM) 셀 너비도 함께 정확해진다.
 */
function applyPageSetup(pageSetup: PageSetup | null): void {
  if (!pageSetup) return; // 원본에서 못 읽으면(비정상 문서) 기존 CSS 기본값을 그대로 둔다
  editorHost.style.width = `${pageSetup.widthPx}px`;
  editorHost.style.maxWidth = `${pageSetup.widthPx}px`;
  editorHost.style.padding =
    `${pageSetup.marginTopPx}px ${pageSetup.marginRightPx}px ${pageSetup.marginBottomPx}px ${pageSetup.marginLeftPx}px`;
  // .ProseMirror의 임의 max-width:65ch 캡을 풀어서 실제 여백만큼의 콘텐츠 폭을 그대로 쓰게 한다
  editorHost.classList.add('page-sized');
}

/**
 * 직접 서식도 문단 스타일도 없는 실행문은 Word에서 word/styles.xml의 문서 기본값(대개
 * 10~11pt)을 상속받아 보여지는데, 웹뷰어는 그 상속을 계산하지 않고 브라우저 기본값(16px)을
 * 그대로 썼다 — 그래서 직접 서식이 거의 없는 문서 전체가 원본보다 커 보였다. editorHost에
 * 상속 가능한 기본 font-size로 걸어두면, 직접 서식이 있는 실행문(fontSize 마크가 만든
 * 인라인 style)은 이 상속값을 그대로 덮어써 영향받지 않는다.
 */
function applyDefaultFontSize(pt: number | null): void {
  if (!pt) return;
  editorHost.style.fontSize = `${(pt * 4) / 3}px`; // 1pt = 96/72px = 4/3px
}

function escapeXmlText(value: string): string {
  return value.replace(/[<>&'"]/g, (ch) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' })[ch] as string);
}

/** "2026-09-06 23:30:11" — toISOString은 UTC라 화면을 보는 사람의 실제 시계와 어긋나 보이므로
 * 로컬 시각 그대로 초 단위까지 조립한다. viewer.js의 동명 함수와 같은 포맷(별도 번들이라 공유 불가). */
function formatWatermarkTimestamp(date: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  return (
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ` +
    `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`
  );
}

/** 뷰어(viewer.js)와 같은 화면 열람용 워터마크를 동시편집 화면에도 띄운다 — 토큰에 서명된(위조
 * 불가) viewer 이름 + 이 화면을 연 시각을 편집 화면 전체에 옅게 반복해 깐다. */
function applyWatermark(viewerName: string): void {
  if (!viewerName) return;
  const text = `${viewerName}  ${formatWatermarkTimestamp(new Date())}`;
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="300" height="180">` +
    `<text x="150" y="100" transform="rotate(-28 150 100)" text-anchor="middle" ` +
    `font-family="sans-serif" font-size="14" fill="rgba(0,0,0,0.16)">${escapeXmlText(text)}</text>` +
    `</svg>`;
  document.documentElement.style.setProperty('--watermark-bg', `url("data:image/svg+xml,${encodeURIComponent(svg)}")`);
}

/**
 * 한글(Hangul) 코드 영역 — 이 범위의 글자만 eastAsia 서체(예: 굴림)를 쓰게 강제하고, 나머지
 * (숫자·영문·기호 등)는 ascii 서체(예: Arial)로 넘긴다. 그냥 "굴림, Arial" 순서로만 나열하면
 * 브라우저는 "이 글자를 그릴 수 있는 첫 폰트"로 폴백하는데, 굴림도 숫자 글리프를 갖고 있어
 * Arial까지 안 내려가 버린다 — 그래서 날짜("2026. 07. 24") 같은 숫자가 Word에서 보이는 것보다
 * 눈에 띄게 얇게 보였다(실제 문서로 확인한 결과). unicode-range로 스크립트별 폰트를 강제로
 * 나눠야 Word처럼 숫자는 ascii 서체로, 한글은 eastAsia 서체로 각각 그려진다.
 */
const HANGUL_UNICODE_RANGE = 'U+1100-11FF, U+3130-318F, U+A960-A97F, U+AC00-D7A3, U+D7B0-D7FF';

/**
 * 문서의 실제 기본 서체(한글 eastAsia + 영문/숫자 ascii, word/theme/theme1.xml로 해석된 값)를
 * 편집 화면 기본 폰트로 적용한다. 두 서체가 다르면 위 unicode-range로 스크립트별로 갈라 쓰는
 * 합성 폰트를 만들어 쓰고, 하나만 있거나 둘이 같으면 그냥 그 서체 하나만 쓴다.
 */
function applyDefaultFontFamily(family: ScriptFontFamily | null): void {
  if (!family || (!family.eastAsia && !family.ascii)) return; // 이게 없으면 이 앱 UI용 폰트로 그려져 Word와 아예 다른 서체가 된다
  const { eastAsia, ascii } = family;
  // .ProseMirror 자신에게 이미 font-family 선언이 있어(style.css) 상위 요소의 inline
  // font-family는 상속되지 않는다 — 그 선언이 참조하는 CSS 변수로 값을 흘려보낸다.
  if (eastAsia && ascii && eastAsia !== ascii) {
    const styleEl = document.createElement('style');
    styleEl.textContent =
      `@font-face { font-family: "doc-mixed-font"; src: local("${eastAsia}"); unicode-range: ${HANGUL_UNICODE_RANGE}; }` +
      `@font-face { font-family: "doc-mixed-font"; src: local("${ascii}"); }`;
    document.head.appendChild(styleEl);
    editorHost.style.setProperty(
      '--doc-font-family',
      `"doc-mixed-font", "${eastAsia}", "${ascii}", var(--font-sans)`,
    );
  } else {
    editorHost.style.setProperty('--doc-font-family', `"${eastAsia ?? ascii}", var(--font-sans)`);
  }
}

/**
 * columnResizing()의 기본 TableView는 <table>을 직접 document.createElement로 만들고
 * schema.ts의 table.toDOM은 아예 호출하지 않는다 — 그래서 표 자체의 정렬(w:tblPr/w:jc,
 * 셀 안 문단 정렬과는 별개)이 스키마에만 있고 화면엔 전혀 반영되지 않았다. columnResizing이
 * 공식으로 제공하는 View 옵션으로 이 클래스를 대신 넣어 정렬 스타일만 덧붙인다.
 */
class AlignedTableView extends TableView {
  constructor(node: PMNode, defaultCellMinWidth: number) {
    super(node, defaultCellMinWidth);
    this.applyAlign(node);
  }

  private applyAlign(node: PMNode): void {
    const align = node.attrs.align as string | null;
    if (align === 'center') {
      this.table.style.marginLeft = 'auto';
      this.table.style.marginRight = 'auto';
    } else if (align === 'right') {
      this.table.style.marginLeft = 'auto';
      this.table.style.marginRight = '0';
    } else {
      this.table.style.marginLeft = '';
      this.table.style.marginRight = '';
    }
  }

  override update(node: PMNode): boolean {
    const ok = super.update(node);
    if (ok) this.applyAlign(node);
    return ok;
  }
}

function showFatalError(message: string): void {
  editorHost.hidden = true;
  toolbar.hidden = true;
  errorMessage.textContent = message;
  errorScreen.hidden = false;
}

async function fetchViewInfo(): Promise<ViewInfo> {
  const res = await fetch(`/api/view/info?token=${encodeURIComponent(token)}`);
  const json = (await res.json()) as { success: boolean; data?: ViewInfo; error?: string };
  if (!res.ok || !json.success || !json.data) {
    throw new Error(json.error || '문서 정보를 불러오지 못했습니다.');
  }
  return json.data;
}

/** Word와 같은 단축키 — 글꼴 크기 단계는 문서 기본 크기를 기준으로 센다. */
function buildKeymap(defaultSizePt: number): Record<string, Command> {
  const nodes = docxSchema.nodes;
  const marks = docxSchema.marks;
  return {
    ...baseKeymap,
    'Mod-z': undo,
    'Mod-y': redo,
    'Mod-Shift-z': redo,
    'Mod-b': toggleMark(marks.strong),
    'Mod-i': toggleMark(marks.em),
    'Mod-u': toggleMark(marks.underline),
    'Mod-=': toggleMark(marks.subscript),
    'Mod-Shift-=': toggleMark(marks.superscript),
    'Mod-+': toggleMark(marks.superscript),
    'Mod-]': stepFontSize(1, defaultSizePt),
    'Mod-[': stepFontSize(-1, defaultSizePt),
    'Mod-Shift->': stepFontSize(1, defaultSizePt),
    'Mod-Shift-<': stepFontSize(-1, defaultSizePt),
    'Mod-l': setAlign('left'),
    'Mod-e': setAlign('center'),
    'Mod-r': setAlign('right'),
    'Mod-j': setAlign('justify'),
    // 브라우저 기본 Ctrl+End/Home은 표가 있으면 문서 끝/처음까지 가지 못하는 경우가 있다 — Word처럼 확실히 보낸다.
    'Mod-End': (state, dispatch) => {
      dispatch?.(state.tr.setSelection(Selection.atEnd(state.doc)).scrollIntoView());
      return true;
    },
    'Mod-Home': (state, dispatch) => {
      dispatch?.(state.tr.setSelection(Selection.atStart(state.doc)).scrollIntoView());
      return true;
    },
    'Mod-Enter': insertPageBreak,
    'Mod-Alt-1': setHeading(1),
    'Mod-Alt-2': setHeading(2),
    'Mod-Alt-3': setHeading(3),
    'Mod-Alt-0': setParagraphStyle(null),
    'Mod-Shift-n': setParagraphStyle(null),
    Tab: chainCommands(sinkListItem(nodes.list_item), goToNextCell(1)),
    'Shift-Tab': chainCommands(liftListItem(nodes.list_item), goToNextCell(-1)),
    Enter: chainCommands(splitListItem(nodes.list_item), baseKeymap.Enter),
  };
}

interface SearchMatch {
  readonly from: number;
  readonly to: number;
}

interface SearchPluginState {
  readonly matches: readonly SearchMatch[];
  readonly current: number;
}

const searchPluginKey = new PluginKey<SearchPluginState>('search');

function findMatches(doc: PMNode, query: string): SearchMatch[] {
  if (!query) return [];
  const lower = query.toLowerCase();
  const matches: SearchMatch[] = [];
  doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    const text = node.text.toLowerCase();
    let idx = text.indexOf(lower);
    while (idx !== -1) {
      matches.push({ from: pos + idx, to: pos + idx + query.length });
      idx = text.indexOf(lower, idx + 1);
    }
  });
  return matches;
}

function buildSearchPlugin(): Plugin<SearchPluginState> {
  return new Plugin<SearchPluginState>({
    key: searchPluginKey,
    state: {
      init: () => ({ matches: [], current: -1 }),
      apply(tr, prev) {
        const meta = tr.getMeta(searchPluginKey) as SearchPluginState | undefined;
        if (meta) return meta;
        if (tr.docChanged) {
          const matches = prev.matches.map((m) => ({ from: tr.mapping.map(m.from), to: tr.mapping.map(m.to) }));
          return { matches, current: prev.current };
        }
        return prev;
      },
    },
    props: {
      decorations(state) {
        const { matches, current } = searchPluginKey.getState(state) ?? { matches: [], current: -1 };
        if (matches.length === 0) return DecorationSet.empty;
        const decorations = matches.map((match, index) =>
          Decoration.inline(match.from, match.to, {
            class: index === current ? 'editor-search-hit editor-search-hit-current' : 'editor-search-hit',
          }),
        );
        return DecorationSet.create(state.doc, decorations);
      },
    },
  });
}

/** 툴바 버튼에 현재 선택 영역의 서식을 반영한다(오피스처럼 굵게/정렬 등이 눌린 상태로 보이게). */
function markActive(state: EditorState, markType: (typeof docxSchema.marks)[keyof typeof docxSchema.marks]): boolean {
  const { from, to, empty, $from } = state.selection;
  return empty ? !!markType.isInSet(state.storedMarks ?? $from.marks()) : state.doc.rangeHasMark(from, to, markType);
}

function currentAlign(state: EditorState): string | null {
  const attrs = state.selection.$from.parent.attrs as { align?: string | null };
  return attrs.align ?? null;
}

function isBlockActive(
  state: EditorState,
  nodeType: (typeof docxSchema.nodes)[keyof typeof docxSchema.nodes],
  attrs: Record<string, unknown> = {},
): boolean {
  const parent = state.selection.$from.parent;
  if (parent.type !== nodeType) return false;
  return Object.entries(attrs).every(([key, value]) => parent.attrs[key] === value);
}

function setPressed(id: string, pressed: boolean): void {
  document.getElementById(id)?.setAttribute('aria-pressed', String(pressed));
}

/** 글꼴·크기 상자에 지금 위치의 값을 보여준다(목록에 없는 값이면 그 값을 항목으로 더한다). */
function showSelectValue(select: HTMLSelectElement | null, value: string): void {
  if (!select || select.value === value) return;
  if (![...select.options].some((option) => option.value === value)) {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    option.dataset.dynamic = '';
    select.append(option);
  }
  select.value = value;
}

function syncToolbarState(state: EditorState, defaultSizePt: number, defaultFontName: string): void {
  const { marks } = docxSchema;
  setPressed('btn-bold', markActive(state, marks.strong));
  setPressed('btn-italic', markActive(state, marks.em));
  setPressed('btn-underline', markActive(state, marks.underline));
  setPressed('btn-strike', markActive(state, marks.strike));
  setPressed('btn-subscript', markActive(state, marks.subscript));
  setPressed('btn-superscript', markActive(state, marks.superscript));
  const align = currentAlign(state);
  setPressed('btn-align-left', !align || align === 'left');
  setPressed('btn-align-center', align === 'center');
  setPressed('btn-align-right', align === 'right');
  setPressed('btn-align-justify', align === 'justify');
  setPressed('btn-bullet-list', isBlockActive(state, docxSchema.nodes.paragraph) && inList(state, docxSchema.nodes.bullet_list));
  setPressed('btn-ordered-list', isBlockActive(state, docxSchema.nodes.paragraph) && inList(state, docxSchema.nodes.ordered_list));
  const style = currentParagraphStyle(state) ?? 'Normal';
  for (const item of document.querySelectorAll<HTMLElement>('.style-gallery-item[data-style]')) {
    item.setAttribute('aria-pressed', String(item.dataset.style === style));
  }
  showSelectValue(document.getElementById('font-family') as HTMLSelectElement | null, currentFontName(state) ?? defaultFontName);
  showSelectValue(document.getElementById('font-size') as HTMLSelectElement | null, String(currentFontSize(state, defaultSizePt)));
}

function inList(state: EditorState, listType: typeof docxSchema.nodes.bullet_list): boolean {
  const { $from } = state.selection;
  for (let depth = $from.depth; depth > 0; depth -= 1) {
    if ($from.node(depth).type === listType) return true;
  }
  return false;
}

function insertTable(view: EditorView, rows: number, cols: number): void {
  const { table, table_row: row, table_cell: cell, paragraph } = docxSchema.nodes;
  const makeRow = () => row.create(null, Array.from({ length: cols }, () => cell.create(null, paragraph.create())));
  const node = table.create(null, Array.from({ length: rows }, makeRow));
  // Word처럼 선택한 글자를 지우지 않는다 — 선택 끝에 표를 넣는다.
  const { state } = view;
  const tr = state.tr.setSelection(TextSelection.create(state.doc, state.selection.to)).replaceSelectionWith(node);
  view.dispatch(tr.scrollIntoView());
  view.focus();
}

function bindToolbarButton(
  view: EditorView,
  id: string,
  command: Command,
): void {
  const button = document.getElementById(id);
  if (!button) return;
  button.addEventListener('click', () => {
    command(view.state, view.dispatch, view);
    view.focus();
  });
}

/**
 * 저장 상태 표시("저장 중… / 모든 변경 내용이 저장됨 / 저장 실패") + Ctrl+S 즉시 저장 — 엑셀·PPT
 * 편집기와 같은 동작. 서버의 flushToStore가 실제로 디스크에 쓴 뒤에만 docMeta.savedAt을 갱신하므로
 * (문서 종류 공통, room.ts) 화면의 "저장됨"은 진짜 저장 완료를 뜻한다. 예전엔 docx 편집기에만 이
 * 표시가 없어 사용자가 저장 여부를 알 수 없었다(브라우저 E2E에서 확인).
 */
function setupSaveStatus(ydoc: Y.Doc, viewToken: string): void {
  // 읽기 전용 열람자에게 저장 상태는 의미가 없다(처음 동기화만으로도 과거 저장 시각 때문에 "저장됨"이 떴다).
  if (!canEdit) return;
  const saveStatusEl = document.getElementById('save-status') as HTMLElement | null;
  const textEl = saveStatusEl?.querySelector('.save-status-text') as HTMLElement | null;
  const docMeta = ydoc.getMap<number>(DOC_META_MAP);
  let lastSavedAt = docMeta.get(DOC_META_SAVED_AT_KEY) ?? 0;
  let lastErrorAt = docMeta.get(DOC_META_SAVE_ERROR_AT_KEY) ?? 0;
  const show = (status: 'saving' | 'saved' | 'error') => {
    if (!saveStatusEl) return;
    saveStatusEl.hidden = false;
    saveStatusEl.dataset.status = status;
    if (textEl) {
      textEl.textContent =
        status === 'saving' ? '저장 중…' : status === 'saved' ? '모든 변경 내용이 저장됨' : '저장 실패 — 편집을 계속하면 다시 시도합니다';
    }
  };
  if (canEdit) {
    ydoc.on('afterTransaction', (tx) => {
      if (tx.local && tx.changed.size > 0 && !tx.changed.has(docMeta as never)) show('saving');
    });
    window.addEventListener('keydown', (event) => {
      if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === 's') {
        event.preventDefault(); // 브라우저의 "페이지 저장" 대화상자 대신 문서 저장
        show('saving');
        fetch(`/api/view/save?token=${encodeURIComponent(viewToken)}`, { method: 'POST' }).catch(() => {
          // 실패해도 유휴 자동저장이 다시 시도한다
        });
      }
    });
  }
  docMeta.observe(() => {
    const savedAt = docMeta.get(DOC_META_SAVED_AT_KEY) ?? 0;
    const errorAt = docMeta.get(DOC_META_SAVE_ERROR_AT_KEY) ?? 0;
    if (errorAt > savedAt && errorAt > lastErrorAt) {
      lastErrorAt = errorAt;
      show('error');
    } else if (savedAt > lastSavedAt) {
      lastSavedAt = savedAt;
      show('saved');
    }
  });
}

async function main(): Promise<void> {
  if (!token) {
    showFatalError('열람 토큰이 없습니다.');
    return;
  }

  let info: ViewInfo;
  try {
    info = await fetchViewInfo();
  } catch (error) {
    showFatalError(error instanceof Error ? error.message : '문서 정보를 불러오지 못했습니다.');
    return;
  }

  // canEdit:false여도 이 편집기 화면 자체는 그대로 보여준다 — grid-editor.ts(xlsx)/editor-slides.ts
  // (pptx)와 같은 정책: 실제 쓰기 차단은 서버가 읽기 전용 연결로 처리하므로(collab/room.ts), 여기선
  // 서식 도구(리본) 그룹만 숨겨 애초에 사용자가 고칠 수 있다고 착각하지 않게 한다 — 뒤로가기/제목/
  // 검색/참여자 같은 상단 줄은 도구가 아니라 계속 보여야 하므로 toolbar 전체가 아니라
  // .toolbar-format만 숨긴다.
  canEdit = info.permissions.canEdit;
  if (!canEdit) readonlyBadgeEl.hidden = false;

  titleEl.textContent = info.name;
  document.title = `${info.name} — DocView 편집`;
  if (info.editHasUnsupportedContent) {
    warningEl.hidden = false;
  }
  applyPageSetup(info.pageSetup);
  applyDefaultFontSize(info.defaultFontSizePt);
  applyDefaultFontFamily(info.defaultFontFamily);
  applyWatermark(info.viewer);

  const ydoc = new Y.Doc();
  const fragment = ydoc.getXmlFragment('prosemirror');
  const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const provider = new WebsocketProvider(`${wsProtocol}//${location.host}/ws/edit`, info.docId, ydoc, {
    params: { token },
  });

  // 이름 기준으로 색을 정한다(ydoc.clientID는 접속마다 무작위라 같은 사람이 다른 브라우저로
  // 들어오면 매번 색이 달라져 버린다) — 같은 이름이면 항상 같은 색.
  const localColor = pickPresenceColorForName(info.viewer);
  provider.awareness.setLocalStateField('user', { name: info.viewer, color: localColor });

  provider.on('status', (event: { status: string }) => {
    statusEl.textContent =
      event.status === 'connected' ? '연결됨' : event.status === 'disconnected' ? '연결 끊김 — 재연결 중…' : event.status;
    statusEl.dataset.status = event.status;
  });
  // 서버가 편집 세션 자체를 거부한 경우(edit.ts의 socket.close(4xxx, 이유)) — 4400대는 재시도해도
  // 같은 이유로 다시 거부될 뿐인 영구 거부다(잘못된 토큰, 문서 없음 등). 기본 동작(지수 백오프로
  // 계속 재연결 시도)을 멈추고 서버가 준 이유를 그대로 보여준다 — grid-editor.ts와 동일한 처리.
  provider.on('connection-close', (event: CloseEvent | null) => {
    if (event && event.code >= 4400) {
      provider.shouldConnect = false;
      showFatalError(event.reason || '서버가 연결을 거부했습니다.');
    }
  });
  let readyPosted = false;
  provider.on('sync', (isSynced: boolean) => {
    if (!isSynced) return;
    statusEl.dataset.status = 'synced';
    // 부모 창(iframe으로 이 문서를 띄운 화면)이 자신의 진행 게이지를 내용이 그려지는 이 시점에 마칠 수 있게
    // 알린다(엑셀 편집기의 docview:ready와 같다) — 첫 동기화 내용이 그려진 다음 페인트 뒤에.
    if (!readyPosted && window.parent !== window) {
      readyPosted = true;
      requestAnimationFrame(() => requestAnimationFrame(() => window.parent.postMessage({ type: 'docview:ready' }, '*')));
    }
  });
  provider.awareness.on('change', () => renderPresenceAvatars(presenceEl, provider.awareness));
  renderPresenceAvatars(presenceEl, provider.awareness);
  setupSaveStatus(ydoc, token);

  // Word 리본 글꼴·크기 상자가 서식 없는 글자에 보여줄 값 — 문서 기본값(styles.xml), 없으면 Word 기본.
  const defaultSizePt = info.defaultFontSizePt ?? 10;
  const defaultFontName = info.defaultFontFamily?.eastAsia ?? info.defaultFontFamily?.ascii ?? '';
  const state = EditorState.create({
    schema: docxSchema,
    plugins: [
      ySyncPlugin(fragment),
      yCursorPlugin(provider.awareness),
      yUndoPlugin(),
      keymap(buildKeymap(defaultSizePt)),
      columnResizing({ View: AlignedTableView }),
      tableEditing(),
      buildSearchPlugin(),
    ],
  });

  // dispatchTransaction 콜백은 view를 참조하므로, new EditorView(...) 생성자 실행 도중(예:
  // ySyncPlugin의 초기 동기화가 생성자 안에서 동기적으로 트랜잭션을 dispatch하는 경우) 콜백이
  // 곧바로 불리면 아직 대입되지 않은 view를 참조해 예외가 난다 — 그러면 main()의 catch가
  // showFatalError를 호출해 이미 마운트된 .ProseMirror까지 통째로 숨겨버린다. 그래서 생성자는
  // 기본 동작으로 완료시키고, view가 확실히 대입된 뒤 setProps로 커스텀 콜백을 나중에 건다.
  // editable: 프로세미러 표준 EditorView 옵션 — false면 DOM contentEditable을 끄고, 키보드
  // 입력·IME 조합·붙여넣기·드롭·자르기 등 "DOM에서 오는" 편집 이벤트 전체를 프로세미러 자신이
  // 무시한다(prosemirror-view의 editHandlers가 view.editable을 직접 검사). 다만 이 게이트는
  // 어디까지나 DOM 이벤트 경로에만 걸리고 view.dispatch를 직접 호출하는 코드(예: 아래 툴바 버튼
  // 클릭 핸들러)는 우회한다 — 그래서 이 옵션만으로는 부족하고, 툴바(.toolbar-format) 자체를
  // 위에서 이미 숨겨 그 우회로를 원천 차단했다.
  const view = new EditorView(editorHost, { state, editable: () => canEdit });
  let syncLayoutFields: (state: EditorState) => void = () => undefined;
  let syncReview: (state: EditorState) => void = () => undefined;
  const shell = initOfficeShell({
    token,
    docName: info.name,
    canEdit,
    canDownload: info.permissions.canDownload,
    canPrint: info.permissions.canPrint,
    onUndo: () => {
      undo(view.state);
      view.focus();
    },
    onRedo: () => {
      redo(view.state);
      view.focus();
    },
  });

  // 서식 복사: 한 번 누르면 다음 선택 한 번에, 두 번 누르면 Esc를 누를 때까지 계속 칠한다(Word와 같다).
  let painter: { format: CopiedFormat; sticky: boolean } | null = null;
  const painterBtn = document.getElementById('btn-format-painter');
  function setPainter(next: typeof painter): void {
    painter = next;
    painterBtn?.setAttribute('aria-pressed', String(next !== null));
    editorHost.classList.toggle('format-painting', next !== null);
  }

  view.setProps({
    dispatchTransaction(tr) {
      const newState = view.state.apply(tr);
      view.updateState(newState);
      syncToolbarState(newState, defaultSizePt, defaultFontName);
      shell.setContextualTab('table', canEdit && isInTable(newState));
      syncLayoutFields(newState);
      syncReview(newState);
      if (tr.docChanged || tr.selectionSet) scheduleStatusUpdate();
    },
  });
  syncToolbarState(view.state, defaultSizePt, defaultFontName);

  // 리본 명령 하나가 실행 취소 한 단계다(Word와 같다) — y-prosemirror는 0.5초 안의 변경을 한 단계로
  // 묶으므로, 명령마다 묶음을 끊는다.
  const run = (command: Command) => {
    command(view.state, view.dispatch, view);
    yUndoPluginKey.getState(view.state)?.undoManager.stopCapturing();
    view.focus();
  };
  const on = (id: string, handler: () => void) => document.getElementById(id)?.addEventListener('click', handler);
  const bind = (id: string, command: Command) => on(id, () => run(command));
  const { marks: m, nodes: n } = docxSchema;

  // ---------- 홈: 클립보드 ----------
  // 브라우저는 스크립트의 클립보드 읽기를 권한으로 막을 수 있다 — 그때는 Ctrl+V를 안내한다.
  on('btn-paste', async () => {
    view.focus();
    try {
      const text = await navigator.clipboard.readText();
      if (text) view.pasteText(text);
    } catch {
      window.alert('브라우저가 클립보드 읽기를 허용하지 않았습니다. Ctrl+V로 붙여넣으세요.');
    }
  });
  on('btn-copy', () => {
    view.focus();
    document.execCommand('copy');
  });
  on('btn-cut', () => {
    view.focus();
    document.execCommand('cut');
  });
  painterBtn?.addEventListener('click', () => {
    if (painter) setPainter(null);
    else setPainter({ format: copyFormat(view.state), sticky: false });
    view.focus();
  });
  painterBtn?.addEventListener('dblclick', () => {
    setPainter({ format: copyFormat(view.state), sticky: true });
    view.focus();
  });
  editorHost.addEventListener('mouseup', () => {
    if (!painter) return;
    // 드래그 선택이 끝난 뒤의 선택 영역에 칠한다 — mouseup 직후엔 아직 선택이 반영 전이라 한 틱 미룬다.
    setTimeout(() => {
      if (!painter || view.state.selection.empty) return;
      pasteFormat(painter.format)(view.state, view.dispatch);
      if (!painter.sticky) setPainter(null);
    }, 0);
  });
  window.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && painter) setPainter(null);
  });

  // ---------- 홈: 글꼴 ----------
  const fontFamilySelect = document.getElementById('font-family') as HTMLSelectElement | null;
  fontFamilySelect?.addEventListener('change', () => {
    const name = fontFamilySelect.value;
    run(name ? applyMark(m.fontFamily, { ascii: name, eastAsia: name }) : clearMark(m.fontFamily));
  });
  const fontSizeSelect = document.getElementById('font-size') as HTMLSelectElement | null;
  fontSizeSelect?.addEventListener('change', () => {
    const size = Number(fontSizeSelect.value);
    if (size > 0) run(applyMark(m.fontSize, { size }));
  });
  bind('btn-font-grow', stepFontSize(1, defaultSizePt));
  bind('btn-font-shrink', stepFontSize(-1, defaultSizePt));
  bindDropdown(document.getElementById('change-case-dropdown'), (mode) => run(changeCase(mode as CaseMode)));
  bind('btn-clear-format', clearFormatting);
  bind('btn-bold', toggleMark(m.strong));
  bind('btn-italic', toggleMark(m.em));
  bind('btn-underline', toggleMark(m.underline));
  bind('btn-strike', toggleMark(m.strike));
  bind('btn-subscript', toggleMark(m.subscript));
  bind('btn-superscript', toggleMark(m.superscript));

  /** 색 버튼(왼쪽: 마지막 색 바로 적용, ▾: 견본 고르기) — Word 리본과 같은 두 부분 버튼. */
  function bindColorButton(
    rootId: string,
    applyId: string,
    barId: string,
    colors: readonly PaletteColor[],
    noneLabel: string,
    initial: string,
    apply: (color: string | null) => Command,
  ): void {
    const root = document.getElementById(rootId);
    const menu = root?.querySelector<HTMLElement>('.dropdown-menu');
    const bar = document.getElementById(barId);
    let last = initial;
    if (menu) fillColorPalette(menu, colors, noneLabel);
    on(applyId, () => run(apply(last)));
    bindDropdown(root, (value) => {
      if (value === 'none') {
        run(apply(null));
        return;
      }
      last = value;
      if (bar) bar.style.background = value;
      run(apply(value));
    });
  }
  bindColorButton('highlight-dropdown', 'btn-highlight-apply', 'highlight-bar', HIGHLIGHT_PALETTE, '색 없음', '#ffff00', (color) =>
    color ? applyMark(m.highlight, { color }) : clearMark(m.highlight),
  );
  bindColorButton('font-color-dropdown', 'btn-font-color-apply', 'font-color-bar', FONT_COLOR_PALETTE, '자동', '#c00000', (color) =>
    color ? applyMark(m.textColor, { color }) : clearMark(m.textColor),
  );

  // ---------- 홈: 단락 ----------
  bind('btn-bullet-list', wrapInList(n.bullet_list));
  bind('btn-ordered-list', wrapInList(n.ordered_list));
  bind('btn-multilevel-list', multilevelList);
  bind('btn-indent', changeIndent(1));
  bind('btn-outdent', changeIndent(-1));
  bind('btn-align-left', setAlign('left'));
  bind('btn-align-center', setAlign('center'));
  bind('btn-align-right', setAlign('right'));
  bind('btn-align-justify', setAlign('justify'));
  bindDropdown(document.getElementById('line-spacing-dropdown'), (value) => run(setLineSpacing(Number(value))));
  bindColorButton('shading-dropdown', 'btn-shading-apply', 'shading-bar', SHADING_PALETTE, '색 없음', '#d9d9d9', setShading);

  // ---------- 홈: 스타일 갤러리 ----------
  for (const item of document.querySelectorAll<HTMLElement>('.style-gallery-item')) {
    item.addEventListener('click', () => {
      const style = item.dataset.style;
      const charStyle = item.dataset.charStyle;
      if (style === 'Normal') run(setParagraphStyle(null));
      else if (style === 'Heading1') run(setHeading(1));
      else if (style === 'Heading2') run(setHeading(2));
      else if (style) run(setParagraphStyle(style));
      else if (charStyle === 'strong') run(applyMark(m.strong));
      else if (charStyle === 'emphasis') run(applyMark(m.em));
      else if (charStyle === 'subtle') {
        applyMark(m.em)(view.state, view.dispatch);
        run(applyMark(m.textColor, { color: '#404040' }));
      }
    });
  }

  // ---------- 삽입 ----------
  const tableGrid = document.getElementById('table-grid');
  const TABLE_GRID_ROWS = 8;
  const TABLE_GRID_COLS = 10;
  if (tableGrid) {
    const caption = tableGrid.parentElement?.querySelector('.table-grid-caption');
    for (let r = 1; r <= TABLE_GRID_ROWS; r += 1) {
      for (let c = 1; c <= TABLE_GRID_COLS; c += 1) {
        const cell = document.createElement('button');
        cell.type = 'button';
        cell.className = 'table-grid-cell';
        cell.dataset.value = `${r}x${c}`;
        cell.setAttribute('aria-label', `${c}×${r} 표`);
        cell.addEventListener('mouseenter', () => {
          for (const other of tableGrid.querySelectorAll<HTMLElement>('.table-grid-cell')) {
            const [rr, cc] = (other.dataset.value ?? '0x0').split('x').map(Number);
            other.classList.toggle('is-hot', rr <= r && cc <= c);
          }
          if (caption) caption.textContent = `${c}×${r} 표`;
        });
        tableGrid.append(cell);
      }
    }
  }
  bindDropdown(document.getElementById('table-insert-dropdown'), (value) => {
    const [rows, cols] = value.split('x').map(Number);
    if (rows > 0 && cols > 0) insertTable(view, rows, cols);
  });
  bind('btn-insert-break', (state, dispatch) => {
    if (dispatch) dispatch(state.tr.replaceSelectionWith(n.hard_break.create()).scrollIntoView());
    return true;
  });

  // ---------- 표 레이아웃(상황별 탭) ----------
  bind('btn-add-row', addRowAfter);
  bind('btn-add-row-before', addRowBefore);
  bind('btn-add-col', addColumnAfter);
  bind('btn-add-col-before', addColumnBefore);
  bind('btn-del-row', deleteRow);
  bind('btn-del-col', deleteColumn);
  bind('btn-del-table', deleteTable);
  bind('btn-merge-cells', mergeCells);
  bind('btn-split-cell', splitCell);

  // ---------- 레이아웃: 용지 설정(여백·방향·크기) — 공유 Y.Map이라 함께 편집하는 사람에게도 바로 보인다 ----------
  const pageMap = ydoc.getMap<unknown>(DOCX_PAGE_MAP);
  const TWIPS_PER_PX = 15;
  const basePage: DocxPageSetup = info.pageSetup
    ? {
        width: info.pageSetup.widthPx * TWIPS_PER_PX,
        height: info.pageSetup.heightPx * TWIPS_PER_PX,
        orientation: info.pageSetup.orientation,
        top: info.pageSetup.marginTopPx * TWIPS_PER_PX,
        bottom: info.pageSetup.marginBottomPx * TWIPS_PER_PX,
        left: info.pageSetup.marginLeftPx * TWIPS_PER_PX,
        right: info.pageSetup.marginRightPx * TWIPS_PER_PX,
      }
    : { ...PAPER_SIZES.A4, orientation: 'portrait', ...MARGIN_PRESETS.normal };
  const effectivePage = (): DocxPageSetup => (pageMap.get(DOCX_PAGE_SETUP_KEY) as DocxPageSetup | undefined) ?? basePage;
  function renderPage(): void {
    const page = effectivePage();
    applyPageSetup({
      widthPx: page.width / TWIPS_PER_PX,
      heightPx: page.height / TWIPS_PER_PX,
      marginTopPx: page.top / TWIPS_PER_PX,
      marginBottomPx: page.bottom / TWIPS_PER_PX,
      marginLeftPx: page.left / TWIPS_PER_PX,
      marginRightPx: page.right / TWIPS_PER_PX,
      orientation: page.orientation,
    });
  }
  pageMap.observe(renderPage);
  const setPage = (patch: Partial<DocxPageSetup>) => pageMap.set(DOCX_PAGE_SETUP_KEY, { ...effectivePage(), ...patch });
  bindDropdown(document.getElementById('margins-dropdown'), (value) => {
    const preset = MARGIN_PRESETS[value as keyof typeof MARGIN_PRESETS];
    if (preset) setPage(preset);
  });
  bindDropdown(document.getElementById('orientation-dropdown'), (value) => {
    const page = effectivePage();
    const long = Math.max(page.width, page.height);
    const short = Math.min(page.width, page.height);
    setPage(value === 'landscape' ? { orientation: 'landscape', width: long, height: short } : { orientation: 'portrait', width: short, height: long });
  });
  bindDropdown(document.getElementById('paper-size-dropdown'), (value) => {
    const size = PAPER_SIZES[value as keyof typeof PAPER_SIZES];
    if (!size) return;
    const landscape = effectivePage().orientation === 'landscape';
    setPage(landscape ? { width: size.height, height: size.width } : size);
  });
  bind('btn-page-break', insertPageBreak);

  // ---------- 레이아웃: 단락 들여쓰기(cm)·간격(pt) ----------
  const DXA_PER_CM = 567;
  const DXA_PER_PT = 20;
  const layoutFields: ReadonlyArray<{ id: string; unit: number; apply: (dxa: number) => Command; read: (box: ReturnType<typeof currentParagraphBox>) => number }> = [
    { id: 'indent-left', unit: DXA_PER_CM, apply: (dxa) => setIndentSide('left', dxa), read: (box) => box.left },
    { id: 'indent-right', unit: DXA_PER_CM, apply: (dxa) => setIndentSide('right', dxa), read: (box) => box.right },
    { id: 'spacing-before', unit: DXA_PER_PT, apply: (dxa) => setSpacingSide('before', dxa), read: (box) => box.before ?? 0 },
    { id: 'spacing-after', unit: DXA_PER_PT, apply: (dxa) => setSpacingSide('after', dxa), read: (box) => box.after ?? 0 },
  ];
  for (const field of layoutFields) {
    const input = document.getElementById(field.id) as HTMLInputElement | null;
    input?.addEventListener('change', () => {
      const value = Number(input.value);
      if (Number.isFinite(value) && value >= 0) run(field.apply(value * field.unit));
    });
    input?.addEventListener('keydown', (event) => {
      // Enter의 기본 동작을 막지 않으면, 값 적용 후 편집기로 돌아간 포커스에 같은 Enter가 들어가 문단이 나뉜다.
      if (event.key === 'Enter') {
        event.preventDefault();
        input.blur();
      }
    });
  }
  syncLayoutFields = (state: EditorState) => {
    const box = currentParagraphBox(state);
    for (const field of layoutFields) {
      const input = document.getElementById(field.id) as HTMLInputElement | null;
      if (input && document.activeElement !== input) input.value = String(Number((field.read(box) / field.unit).toFixed(field.unit === DXA_PER_CM ? 2 : 1)));
    }
  };
  syncLayoutFields(view.state);

  // ---------- 검토: 맞춤법·단어 개수·메모·변경 내용 ----------
  syncReview = setupWordReview({
    view,
    comments: ydoc.getMap<DocxComment>(DOCX_COMMENTS_MAP),
    token,
    viewer: info.viewer,
    canEdit,
    run,
  });

  // ---------- 상태 표시줄: 단어 수(선택하면 "선택/전체") ----------
  const statusWords = document.getElementById('status-words');
  const statusChars = document.getElementById('status-chars');
  let statusTimer: ReturnType<typeof setTimeout> | undefined;
  const countWords = (text: string) => (text.trim() ? text.trim().split(/\s+/).length : 0);
  function updateStatus(): void {
    const { doc, selection } = view.state;
    const all = doc.textBetween(0, doc.content.size, '\n', ' ');
    const words = countWords(all);
    const chars = all.replace(/\n/g, '').length;
    if (!selection.empty) {
      const selected = doc.textBetween(selection.from, selection.to, '\n', ' ');
      if (statusWords) statusWords.textContent = `단어 ${countWords(selected).toLocaleString()}/${words.toLocaleString()}개`;
    } else if (statusWords) {
      statusWords.textContent = `단어 ${words.toLocaleString()}개`;
    }
    if (statusChars) statusChars.textContent = `문자 ${chars.toLocaleString()}개(공백 포함)`;
  }
  function scheduleStatusUpdate(): void {
    clearTimeout(statusTimer);
    statusTimer = setTimeout(updateStatus, 200);
  }
  updateStatus();

  // ---------- 확대/축소(상태 표시줄 슬라이더 + 보기 탭) ----------
  const zoomHost = document.getElementById('editor-zoom-host') as HTMLElement;
  const zoomLevelEl = document.getElementById('zoom-level') as HTMLElement;
  const zoomTrack = document.getElementById('zoom-slider-track');
  const zoomFill = document.getElementById('zoom-slider-fill');
  const zoomHandle = document.getElementById('zoom-slider-handle');
  const ZOOM_MIN = 0.6;
  const ZOOM_MAX = 2;
  // 뷰어("바로 보기")는 페이지를 실제 96dpi 크기가 아니라 .pages-container(최대 1120px,
  // 좌우 1rem 패딩 → 실사용 1088px)에 맞춰 확대해서 보여준다 — 편집기는 지금까지 페이지를
  // 항상 실제 크기(zoom=100%)로만 그려서, 같은 문서인데 "종이 자체가 더 작아 보인다"는
  // 차이가 났다. 뷰어와 같은 컨테이너 폭 기준으로 초기 줌을 계산해 화면에 보이는 크기를 맞춘다.
  const VIEWER_PAGE_CONTAINER_MAX_PX = 1120;
  const VIEWER_PAGE_CONTAINER_PADDING_PX = 32; // 뷰어 .pages-container의 좌우 1rem(16px)×2
  const clampZoom = (value: number) => Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round(value * 100) / 100));
  function fitZoom(maxWidthPx: number): number {
    if (!info.pageSetup) return 1;
    const availableWidthPx = Math.min(maxWidthPx, window.innerWidth) - VIEWER_PAGE_CONTAINER_PADDING_PX;
    // 레이아웃 탭에서 용지 방향·크기를 바꿨으면 지금 용지 폭 기준으로 맞춘다
    const pageWidthPx = parseFloat(editorHost.style.width) || info.pageSetup.widthPx;
    return clampZoom(Math.round((availableWidthPx / pageWidthPx) * 10) / 10);
  }
  let zoom = fitZoom(VIEWER_PAGE_CONTAINER_MAX_PX);
  function applyZoom(): void {
    zoomHost.style.zoom = String(zoom);
    zoomLevelEl.textContent = `${Math.round(zoom * 100)}%`;
    const ratio = (zoom - ZOOM_MIN) / (ZOOM_MAX - ZOOM_MIN);
    if (zoomFill) zoomFill.style.width = `${ratio * 100}%`;
    if (zoomHandle) zoomHandle.style.left = `${ratio * 100}%`;
  }
  const setZoom = (value: number) => {
    zoom = clampZoom(value);
    applyZoom();
  };
  applyZoom();
  on('zoom-in', () => setZoom(Math.round((zoom + 0.1) * 10) / 10));
  on('zoom-out', () => setZoom(Math.round((zoom - 0.1) * 10) / 10));
  on('btn-zoom-in-ribbon', () => setZoom(Math.round((zoom + 0.1) * 10) / 10));
  on('btn-zoom-out-ribbon', () => setZoom(Math.round((zoom - 0.1) * 10) / 10));
  on('btn-zoom-100', () => setZoom(1));
  on('btn-zoom-page-width', () => setZoom(fitZoom(window.innerWidth)));
  zoomLevelEl.addEventListener('click', () => setZoom(1));
  if (zoomTrack) {
    const fromPointer = (event: PointerEvent) => {
      const rect = zoomTrack.getBoundingClientRect();
      const ratio = Math.min(1, Math.max(0, (event.clientX - rect.left) / rect.width));
      setZoom(ZOOM_MIN + ratio * (ZOOM_MAX - ZOOM_MIN));
    };
    zoomTrack.addEventListener('pointerdown', (event) => {
      zoomTrack.setPointerCapture(event.pointerId);
      fromPointer(event);
    });
    zoomTrack.addEventListener('pointermove', (event) => {
      if (zoomTrack.hasPointerCapture(event.pointerId)) fromPointer(event);
    });
  }

  // ---------- 찾기 및 바꾸기 ----------
  const searchBar = document.getElementById('search-bar') as HTMLElement;
  const searchInput = document.getElementById('search-input') as HTMLInputElement;
  const replaceRow = document.getElementById('search-replace-row') as HTMLElement;
  const replaceInput = document.getElementById('search-replace-input') as HTMLInputElement;
  const replaceOneBtn = document.getElementById('search-replace-one') as HTMLButtonElement;
  const replaceAllBtn = document.getElementById('search-replace-all') as HTMLButtonElement;
  const searchCount = document.getElementById('search-count') as HTMLElement;
  const searchPrevBtn = document.getElementById('search-prev') as HTMLButtonElement;
  const searchNextBtn = document.getElementById('search-next') as HTMLButtonElement;
  let searchMatches: SearchMatch[] = [];
  let searchCurrent = -1;

  function applySearch(jumpTo: number): void {
    searchMatches = findMatches(view.state.doc, searchInput.value);
    searchCurrent = searchMatches.length > 0 ? ((jumpTo % searchMatches.length) + searchMatches.length) % searchMatches.length : -1;
    let tr = view.state.tr.setMeta(searchPluginKey, { matches: searchMatches, current: searchCurrent });
    if (searchCurrent >= 0) {
      const match = searchMatches[searchCurrent];
      tr = tr.setSelection(TextSelection.create(tr.doc, match.from, match.to)).scrollIntoView();
    }
    view.dispatch(tr);
    searchCount.textContent = searchMatches.length > 0 ? `${searchCurrent + 1}/${searchMatches.length}` : searchInput.value ? '0/0' : '';
    searchPrevBtn.disabled = searchMatches.length === 0;
    searchNextBtn.disabled = searchMatches.length === 0;
  }
  function openSearch(withReplace: boolean): void {
    searchBar.hidden = false;
    const showReplace = withReplace && canEdit;
    replaceRow.hidden = !showReplace;
    replaceOneBtn.hidden = !showReplace;
    replaceAllBtn.hidden = !showReplace;
    const selected = view.state.selection.empty ? '' : view.state.doc.textBetween(view.state.selection.from, view.state.selection.to, ' ');
    if (selected && !selected.includes('\n') && selected.length < 100) searchInput.value = selected;
    searchInput.focus();
    searchInput.select();
    if (searchInput.value) applySearch(0);
  }
  function closeSearch(): void {
    searchBar.hidden = true;
    searchInput.value = '';
    applySearch(0);
    view.focus();
  }
  on('search-btn', () => (searchBar.hidden ? openSearch(false) : closeSearch()));
  on('btn-ribbon-find', () => openSearch(false));
  on('btn-ribbon-replace', () => openSearch(true));
  on('btn-ribbon-select', () => run(selectAll));
  on('search-close', closeSearch);
  replaceOneBtn.addEventListener('click', () => {
    const match = searchMatches[searchCurrent];
    if (!match) return;
    replaceRange(match.from, match.to, replaceInput.value)(view.state, view.dispatch);
    applySearch(searchCurrent);
  });
  replaceAllBtn.addEventListener('click', () => {
    const count = replaceAll(view.state, view.dispatch, searchInput.value, replaceInput.value);
    applySearch(0);
    searchCount.textContent = `${count}개 바꿈`;
  });
  // applySearch가 매번 findMatches로 문서 전체를 다시 훑으므로(문서가 크면 체감되는 비용), 입력
  // 중에는 짧게 모아서(디바운스) 실행한다 — Enter/이전·다음 버튼 등 명시적 액션은 즉시 반응해야
  // 하므로 이 디바운스를 적용하지 않는다.
  let searchInputDebounce: ReturnType<typeof setTimeout> | undefined;
  searchInput.addEventListener('input', () => {
    clearTimeout(searchInputDebounce);
    searchInputDebounce = setTimeout(() => applySearch(0), 150);
  });
  searchInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') applySearch(searchCurrent + (event.shiftKey ? -1 : 1));
    if (event.key === 'Escape') closeSearch();
  });
  replaceInput.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') closeSearch();
  });
  searchPrevBtn.addEventListener('click', () => applySearch(searchCurrent - 1));
  searchNextBtn.addEventListener('click', () => applySearch(searchCurrent + 1));
  window.addEventListener('keydown', (event) => {
    if (!(event.ctrlKey || event.metaKey) || event.altKey) return;
    const key = event.key.toLowerCase();
    if (key === 'f') {
      event.preventDefault();
      openSearch(false);
    } else if (key === 'h') {
      event.preventDefault();
      openSearch(true);
    }
  });

  window.addEventListener('beforeunload', () => {
    provider.disconnect();
  });
}

main().catch((error) => {
  console.error(error);
  showFatalError('편집기를 시작하지 못했습니다.');
});
