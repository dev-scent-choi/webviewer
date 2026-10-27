import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import * as encoding from 'lib0/encoding';
import * as decoding from 'lib0/decoding';
import {
  DOC_META_MAP,
  DOC_META_SAVED_AT_KEY,
  DOC_META_SAVE_ERROR_AT_KEY,
  EDIT_SOFT_COL_LIMIT,
  EDIT_SOFT_ROW_LIMIT,
  GRID_CELLS_MAP,
  GRID_COMPUTED_MAP,
  GRID_MERGES_MAP,
  GRID_STRUCTURE_MAP,
  GRID_STYLES_MAP,
  gridCellKey,
  parseGridCellKey,
  PAGE_SIZE_ROWS,
  type CellStyleAttrs,
  type GridMergeShape,
  type SheetStructure,
} from '../collab/gridShared.js';
import { dedupeByName, pickPresenceColorForName, renderPresenceAvatars, type AwarenessUserState } from './presence.js';
import { applyWatermark } from './watermark.js';
import { parseTypedNumber, serialToDateText } from '../collab/cellValueShared.js';
import { initOfficeShell } from './officeShell.js';
import { bindDropdown } from './ribbonMenus.js';
import { currentRegion, setupGridCharts } from './gridCharts.js';
import { setupGridLinks, type GridLinks } from './gridLinks.js';
import {
  CF_PRESETS,
  computeStats,
  GRID_COND_FORMATS_MAP,
  needsStats,
  paintCell,
  rectContains,
  subtractRect,
  type CfRule,
  type CfStats,
  type CondFormatEntry,
} from '../collab/condFormatShared.js';
import type { CellRangeRef, ChartKind } from '../xlsx/chartModel.js';
import { RowLayout } from '../xlsx/sheetGeometry.js';
import { compareCells, filteredRows, GRID_FILTERS_MAP, shiftFormulaRows, sortOrder, type SheetFilter } from './gridSort.js';
import { officeAlert, officeConfirm } from './officeDialog.js';

interface Permissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  readonly canCopy: boolean;
  readonly canEdit: boolean;
}

interface ViewInfo {
  readonly docId: string;
  readonly name: string;
  readonly viewer: string;
  readonly permissions: Permissions;
}

const ALIGN_TO_CSS: Record<'l' | 'c' | 'r', string> = { l: 'left', c: 'center', r: 'right' };
/** 글꼴 드롭다운에서 고른 이름이 이 컴퓨터/브라우저에 없을 때를 위한 대체 글꼴 목록. */
const DEFAULT_CELL_FONT_STACK = "'맑은 고딕', Malgun Gothic, Calibri, sans-serif";
/** CellStyleAttrs.fs는 항상 px로 저장한다(기존 관례) — 툴바 입력창은 엑셀처럼 pt로 보여주므로
 * 여기서만 서로 변환한다. */
const PT_TO_PX = 96 / 72;
const PX_TO_PT = 72 / 96;

const params = new URLSearchParams(location.search);
const token = params.get('token') ?? '';

type AutoFunction = 'SUM' | 'AVERAGE' | 'COUNT' | 'MAX' | 'MIN';

/** false면 이 문서 전체가 읽기 전용이다(예: 과거 연도 접속현황) — 잠금 여부(lk)와 무관하게 모든
 * 칸이 편집 불가여야 한다. main()에서 서버가 내려준 토큰 권한으로 설정된다. */
let canEdit = true;

/** renderTableVirtualized()가 다시 그릴 때마다 하나씩 늘어난다. 붙여넣기로 시트 범위가 넓어지거나
 * 스크롤로 가상 렌더가 다시 그리면 renderTableVirtualized()가 트랜잭션 도중 동기적으로 표 전체를
 * (또는 화면에 보이는 범위를) 다시 그리는데, 이때 포커스가 있던 셀이
 * DOM에서 떨어져 나가며 브라우저가 blur를 보낸다 — 그 blur 콜백이 실행되는 시점엔 이미 그 요소가
 * 실제로는 못 쓰게 됐어도 `isConnected`가 아직 true로 보이는 타이밍이 브라우저마다 있어(제거
 * 시점과 blur 디스패치 시점이 항상 완전히 동기적이지 않음) 그 값만으로는 안전하게 걸러낼 수 없다.
 * 대신 "이 칸이 만들어졌을 때의 세대"와 "지금 세대"를 비교하는 논리적 판정을 쓴다 — 사이에 표가
 * 다시 그려졌으면(세대가 달라졌으면) 그 blur는 이미 낡은 것이니 커밋하지 않는다. */
let renderGeneration = 0;

const titleEl = document.getElementById('doc-title') as HTMLElement;
const statusEl = document.getElementById('conn-status') as HTMLElement;
const presenceEl = document.getElementById('presence-list') as HTMLElement;
const warningEl = document.getElementById('unsupported-warning') as HTMLElement;
const editorHost = document.getElementById('grid-editor') as HTMLElement;
const zoomHost = document.getElementById('grid-zoom-host') as HTMLElement;
const tableEl = document.getElementById('grid-edit-table') as HTMLTableElement;
const scrollHostEl = document.querySelector('.grid-editor-scroll') as HTMLElement;
/** 표를 통째로 다시 그린 뒤(시트 전환·구조 변경·확대축소) — 차트 위치를 다시 맞춘다(차트 준비 전엔 아무것도 안 한다). */
let afterGridRender: () => void = () => undefined;

/** 지금 구독 중(서버가 채워 둔) 500행 페이지 — 정렬·필터는 이 안의 데이터로만 계산한다 */
let lastSubscribedPages: ReadonlySet<number> = new Set();
/** 그중 화면에 보이는(스크롤 위치 기준) 페이지 — 로딩 표시는 이것만 본다(필터 범위로 더 받는 페이지는 뒤에서) */
let visiblePages: ReadonlySet<number> = new Set();
/** 서버가 한 연결에 받아 주는 페이지 수 상한(room.ts MAX_SUBSCRIBED_PAGES_PER_CONN과 같게) */
const MAX_SUBSCRIBED_PAGES = 32;
/** 서버 → 편집기 "채움 완료" 알림 번호(room.ts와 같은 값) */
const MESSAGE_PAGES_READY = 6;
/** 시트별로 서버가 실제로 채워 준 페이지 — lastSubscribedPages는 "요청한" 것, 이건 "도착한" 것 */
const arrivedPages = new Map<number, Set<number>>();
/** 도착 알림을 받을 때마다 — 로딩 표시·필터를 다시 맞추는 쪽이 아래에서 채운다 */
let pagesArrivedHook: () => void = () => undefined;
/** 필터로 숨긴 행(지금 시트, 오름차순) — rowLayout에서 높이 0으로 계산된다 */
let hiddenRowsSorted: number[] = [];
let hiddenRowSet: ReadonlySet<number> = new Set();
/** 엑셀 기본 행 높이(15pt) — 시트에 값이 없을 때(구버전 문서) */
const DEFAULT_ROW_HEIGHT_PX = 20;
/** 지금 시트의 행 위치 계산 — 파일의 행 높이·숨김 + 필터로 숨긴 행(0px). renderTableVirtualized가 맞춘다 */
let rowLayout = new RowLayout(DEFAULT_ROW_HEIGHT_PX);
const sheetTabsEl = document.getElementById('sheet-tabs') as HTMLElement;
const cellRefBoxEl = document.getElementById('cell-ref-box') as HTMLElement;
const formulaInputEl = document.getElementById('formula-input') as HTMLInputElement;
const toolbar = document.getElementById('toolbar') as HTMLElement;
const errorScreen = document.getElementById('error-screen') as HTMLElement;
const errorMessage = document.getElementById('error-message') as HTMLElement;
const loadingScreen = document.getElementById('loading-screen') as HTMLElement;
const loadingPercentEl = document.getElementById('loading-percent') as HTMLElement;
const loadingGaugeFillEl = document.getElementById('loading-gauge-fill') as HTMLElement;

// 다른 화면에 iframe으로 임베드된 경우엔 부모 창이 이미 자기 스피너를 쓰고 있으므로(위
// g-render-pending 관련 주석 참고) 이 화면의 로딩 표시는 끈다 — 독립 접속일 때만 보인다.
const showsOwnLoadingScreen = window.parent === window;
if (!showsOwnLoadingScreen) {
  loadingScreen.hidden = true;
}

/**
 * 최초 동기화(서버가 canonical 그리드를 다시 구성하는 과정)는 진행률을 알려줄 방법이 없다 —
 * ExcelJS 자체의 압축 해제·XML 파싱이 대부분의 시간을 차지하는데 그 내부 진행 상황은 이
 * 라이브러리가 알려주지 않는다. 그래서 index.html의 업로드 변환 게이지(main.js의
 * CONVERT_CREEP_CAP/RATE)와 똑같은 방식 — 시간이 지날수록 상한(92%)에 점근하도록 부드럽게
 * 채워, "멈춘 게 아니라 진행 중"이라는 신호만 준다. 실제 완료(provider의 sync 이벤트)가
 * 오면 그 즉시 100%로 스냅한다.
 */
const LOADING_CREEP_CAP = 92;
const LOADING_CREEP_RATE = 0.04;
const LOADING_CREEP_INTERVAL_MS = 300;
let loadingDisplayed = 0;
let loadingCreepTimer: ReturnType<typeof setInterval> | null = null;
function setLoadingProgress(percent: number): void {
  loadingPercentEl.textContent = `${Math.round(percent)}%`;
  loadingGaugeFillEl.style.transform = `scaleX(${percent / 100})`;
}
if (showsOwnLoadingScreen) {
  loadingCreepTimer = setInterval(() => {
    loadingDisplayed = Math.min(LOADING_CREEP_CAP, loadingDisplayed + (LOADING_CREEP_CAP - loadingDisplayed) * LOADING_CREEP_RATE);
    setLoadingProgress(loadingDisplayed);
  }, LOADING_CREEP_INTERVAL_MS);
}
function stopLoadingCreep(): void {
  if (loadingCreepTimer !== null) {
    clearInterval(loadingCreepTimer);
    loadingCreepTimer = null;
  }
  setLoadingProgress(100);
}

function showFatalError(message: string): void {
  stopLoadingCreep();
  loadingScreen.hidden = true;
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

function colLetter(index: number): string {
  let n = index;
  let label = '';
  do {
    label = String.fromCharCode(65 + (n % 26)) + label;
    n = Math.floor(n / 26) - 1;
  } while (n >= 0);
  return label;
}

function cellRefLabel(row: number, col: number): string {
  return `${colLetter(col)}${row + 1}`;
}

/** 엑셀처럼 숫자로 보이는 값은 서식이 따로 없으면 기본적으로 오른쪽 정렬로 보여준다. */
const NUMERIC_VALUE = /^-?\d+(\.\d+)?$/;

/** 엑셀 테두리 굵기(1=가는 선, 2=굵은 선)를 CSS로. 지정이 없으면 빈 문자열 —
 * 그러면 스타일시트의 기본 점선 격자(.grid-edit-table td)가 그대로 보인다. */
function borderCss(weight: 1 | 2 | undefined): string {
  return weight ? `${weight}px solid var(--v-text)` : '';
}

const VALIGN_TO_CSS: Record<'t' | 'm' | 'b', string> = { t: 'top', m: 'middle', b: 'bottom' };

/** 조건부 서식 색 덧입히기 — main()이 연결한다. 칸 서식을 다시 칠하는 곳(applyStyleToTd)이 여러 군데라 그 끝에서 부른다. */
let condFormatOverlay: ((td: HTMLTableCellElement) => void) | null = null;

function applyStyleToTd(td: HTMLTableCellElement, style: CellStyleAttrs | undefined): void {
  td.style.fontWeight = style?.b ? '700' : '400';
  td.style.fontStyle = style?.i ? 'italic' : 'normal';
  td.style.textDecoration = [style?.u ? 'underline' : '', style?.st ? 'line-through' : ''].filter(Boolean).join(' ');
  td.style.color = style?.col ?? '';
  td.style.background = style?.bg ?? '';
  td.style.fontFamily = style?.ff ? `"${style.ff}", ${DEFAULT_CELL_FONT_STACK}` : '';
  td.style.fontSize = style?.fs ? `${style.fs}px` : '';
  td.style.verticalAlign = style?.va ? VALIGN_TO_CSS[style.va] : '';
  // 자동 줄 바꿈 — 셀 높이는 고정이라(틀고정 sticky 좌표 계산이 고정 행 높이를 전제로 함) 줄이
  // 너무 많아지면 아래가 잘릴 수 있다. 그래도 한 줄에 다 몰아넣고 말줄임표로 자르는 것보다는
  // 실제로 줄바꿈되는 쪽이 "자동 줄 바꿈"이라는 기능 자체엔 맞다.
  td.style.whiteSpace = style?.wrap ? 'normal' : '';
  if (style?.a) {
    td.style.textAlign = ALIGN_TO_CSS[style.a];
  } else if (style?.nf) {
    td.style.textAlign = 'right'; // 숫자 서식이 있으면 서식이 적용된 텍스트(₩1,000 등)라 정규식 판정이 안 먹는다
  } else {
    td.style.textAlign = NUMERIC_VALUE.test((td.textContent ?? '').trim()) ? 'right' : '';
  }
  // 이 표는 칸마다 테두리를 따로 그린다(border-collapse: separate) — 위·왼쪽 테두리를 칸 안쪽에 그리면 위 칸의
  // 아래 선·왼쪽 칸의 오른쪽 선과 나란히 붙어 엑셀보다 두 배 굵어졌다. 오른쪽·아래는 칸 자체 테두리로, 위·왼쪽은
  // 이웃 칸의 경계선 위에 겹쳐 그린다(CSS .g-bord-tl::before).
  td.style.borderTop = '';
  td.style.borderLeft = '';
  td.style.borderRight = borderCss(style?.br);
  td.style.borderBottom = borderCss(style?.bb);
  td.style.setProperty('--g-bt', borderCss(style?.bt) || 'none');
  td.style.setProperty('--g-bl', borderCss(style?.bl) || 'none');
  td.classList.toggle('g-bord-tl', Boolean(style?.bt || style?.bl));
  td.classList.toggle('g-locked', style?.lk === 1);
  condFormatOverlay?.(td);
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

  // canEdit:false여도 이 편집기 화면 자체는 그대로 보여준다(과거 연도 접속현황처럼 "현재 연도와
  // 똑같은 화면으로, 다만 손댈 수는 없게" 보여줘야 하는 경우가 있다) — 실제 쓰기 차단은 서버가
  // 읽기 전용 연결로 처리한다(collab/room.ts). 여기서는 서식 도구(리본)만 꺼서 애초에 사용자가
  // 고칠 수 있다고 착각하지 않게 한다 — 뒤로가기/제목/검색/참여자 같은 상단 줄은 도구가 아니라
  // 계속 보여야 하므로 toolbar 전체가 아니라 .toolbar-format만 숨긴다.
  canEdit = info.permissions.canEdit;
  if (!canEdit) formulaInputEl.readOnly = true;

  titleEl.textContent = info.name;
  document.title = `${info.name} — DocView 편집`;
  applyWatermark(info.viewer);

  const ydoc = new Y.Doc();
  const cells = ydoc.getMap<string>(GRID_CELLS_MAP);
  const computed = ydoc.getMap<string>(GRID_COMPUTED_MAP);
  const structure = ydoc.getMap<SheetStructure>(GRID_STRUCTURE_MAP);
  const styles = ydoc.getMap<CellStyleAttrs>(GRID_STYLES_MAP);
  const merges = ydoc.getMap<readonly GridMergeShape[]>(GRID_MERGES_MAP);
  // 조건부 서식 규칙(시트마다 우선순위 순 목록 — condFormatShared.ts)
  const condFormats = ydoc.getMap<readonly CondFormatEntry[]>(GRID_COND_FORMATS_MAP);

  const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const provider = new WebsocketProvider(`${wsProtocol}//${location.host}/ws/edit`, info.docId, ydoc, {
    params: { token },
  });
  // 서버가 "이 페이지들은 다 채웠다"고 알린다(room.ts MESSAGE_PAGES_READY = 6) — 구독(요청)한 페이지가 실제로
  // 도착했는지 이걸로 안다. 로딩 표시·필터 값 목록·정렬 가능 여부가 이 값을 본다.
  (provider.messageHandlers as unknown as Array<(e: encoding.Encoder, d: decoding.Decoder) => void>)[MESSAGE_PAGES_READY] = (_encoder, decoder) => {
    try {
      const { sheet, pages } = JSON.parse(decoding.readVarString(decoder)) as { sheet: number; pages: number[] };
      arrivedPages.set(sheet, new Set(pages));
      pagesArrivedHook();
    } catch {
      /* 형식이 깨진 알림은 무시 — 다음 알림이 바로잡는다 */
    }
  };

  // 이름 기준으로 색을 정한다(ydoc.clientID는 접속마다 무작위라 같은 사람이 다른 브라우저로
  // 들어오면 매번 색이 달라져 버린다) — 같은 이름이면 항상 같은 색.
  const localColor = pickPresenceColorForName(info.viewer);
  provider.awareness.setLocalStateField('user', { name: info.viewer, color: localColor });
  // 그리드 전체의 기본 --remote-color를 내 색으로 깔아 둔다 — 그러면 칸을 선택만 한 순간부터도
  // (편집모드에 들어가기 전부터) 이미 내 아바타 색 테두리로 보이고, 실제로 편집모드에 들어가도
  // 색이 안 바뀌니 따로 트랜지션을 신경 쓸 필요가 없다. 다른 참여자가 편집 중인 칸은 그 칸에 직접
  // 걸리는 인라인 --remote-color(renderRemoteCellHighlights)가 이 기본값보다 항상 우선한다.
  tableEl.style.setProperty('--remote-color', localColor);

  provider.on('status', (event: { status: string }) => {
    statusEl.textContent =
      event.status === 'connected' ? '연결됨' : event.status === 'disconnected' ? '연결 끊김 — 재연결 중…' : event.status;
    statusEl.dataset.status = event.status;
  });
  // 서버가 편집 세션 자체를 거부한 경우(edit.ts의 socket.close(4xxx, 이유)) — 4400대는 재시도해도
  // 같은 이유로 다시 거부될 뿐인 영구 거부다(잘못된 토큰, 문서 없음, 동시편집 룸 상한 등). 기본
  // 동작(지수 백오프로 계속 재연결 시도)을 멈추고 서버가 준 이유를 그대로 보여준다 — 특히 룸 상한
  // 거부("사용자가 많습니다…")는 그냥 두면 사용자가 원인을 알 길 없이 "연결 끊김 — 재연결 중…"만
  // 계속 보게 된다.
  provider.on('connection-close', (event: CloseEvent | null) => {
    if (event && event.code >= 4400) {
      provider.shouldConnect = false;
      showFatalError(event.reason || '서버가 연결을 거부했습니다.');
    }
  });
  provider.on('sync', (isSynced: boolean) => {
    if (isSynced) {
      statusEl.dataset.status = 'synced';
      renderSheetTabs();
      renderTableVirtualized();
      // 동기화 직후 structure.observe가 한 번 더 renderTableVirtualized()를 트리거해 폭이 살짝 바뀌는 경우가
      // 있다 — 그 재렌더까지 끝날 시간(두 번의 rAF, 즉 최소 한 번의 페인트)을 준 뒤에야 화면에
      // 드러내, 사용자가 "넓었다가 줄어드는" 깜빡임을 보지 않게 한다.
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          editorHost.classList.remove('g-render-pending');
          stopLoadingCreep();
          loadingScreen.hidden = true;
          // 부모 창(iframe으로 이 문서를 띄운 화면)이 자신의 로딩 스피너를 이 시점까지 유지하다가
          // 여기서 iframe으로 교체할 수 있도록 알려준다 — webviewer 자신은 스피너 UI를 새로 만들지
          // 않는다(부모가 이미 쓰고 있는 스피너를 그대로 재사용하는 게 목적).
          if (window.parent !== window) {
            window.parent.postMessage({ type: 'docview:ready' }, '*');
          }
        });
      });
    }
  });

  // ---------- 저장 상태 표시(Office 365 스타일 "저장 중… / 저장됨") + Ctrl+S ----------
  // docMeta.savedAt은 서버의 flushToStore가 실제로 디스크에 쓴 뒤에만 갱신하는 일반 Y.Map
  // 값이라, 이 클라이언트뿐 아니라 같이 편집 중인 다른 사람 화면에도 똑같이 반영된다.
  // 엑셀 하단 상태 표시줄의 "준비/입력/편집" 텍스트 — 지금 이 세션의 편집 동작 기준이라
  // 다른 참여자에게는 전파하지 않는다(각자 자기 화면 기준 상태).
  const gridStatusTextEl = document.getElementById('grid-status-text') as HTMLElement | null;
  function setGridStatusText(text: string): void {
    if (gridStatusTextEl) gridStatusTextEl.textContent = text;
  }
  if (!canEdit && gridStatusTextEl) gridStatusTextEl.textContent = '읽기 전용';

  const docMeta = ydoc.getMap<number>(DOC_META_MAP);
  const saveStatusEl = document.getElementById('save-status') as HTMLElement | null;
  const saveStatusTextEl = saveStatusEl?.querySelector('.save-status-text') as HTMLElement | null;
  let lastKnownSavedAt = docMeta.get(DOC_META_SAVED_AT_KEY) ?? 0;
  let lastKnownSaveErrorAt = docMeta.get(DOC_META_SAVE_ERROR_AT_KEY) ?? 0;
  function setSaveStatus(status: 'saving' | 'saved' | 'error'): void {
    if (!saveStatusEl) return;
    saveStatusEl.hidden = false;
    saveStatusEl.dataset.status = status;
    if (saveStatusTextEl) {
      saveStatusTextEl.textContent =
        status === 'saving' ? '저장 중…' : status === 'saved' ? '모든 변경 내용이 저장됨' : '저장 실패 — 편집을 계속하면 다시 시도합니다';
    }
  }
  if (canEdit) {
    ydoc.on('afterTransaction', (tx) => {
      if (!tx.local || tx.changed.size === 0) return;
      setSaveStatus('saving');
    });
  }
  // saveErrorAt이 savedAt보다 최신이면 아직 실패가 해소되지 않은 것 — 서버가 다음 저장에
  // 성공해 savedAt이 다시 앞서는 순간 별도 정리 없이도 자연스럽게 실패 표시가 사라진다.
  docMeta.observe(() => {
    const savedAt = docMeta.get(DOC_META_SAVED_AT_KEY) ?? 0;
    const saveErrorAt = docMeta.get(DOC_META_SAVE_ERROR_AT_KEY) ?? 0;
    if (saveErrorAt > savedAt && saveErrorAt > lastKnownSaveErrorAt) {
      lastKnownSaveErrorAt = saveErrorAt;
      setSaveStatus('error');
    } else if (savedAt > lastKnownSavedAt) {
      lastKnownSavedAt = savedAt;
      setSaveStatus('saved');
    }
  });
  if (canEdit) {
    window.addEventListener('keydown', (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 's') {
        event.preventDefault();
        setSaveStatus('saving');
        fetch(`/api/view/save?token=${encodeURIComponent(token)}`, { method: 'POST' }).catch(() => {
          // 네트워크 오류 등으로 실패해도 8초 유휴 자동저장이 알아서 다시 시도하므로 별도 안내는 생략한다.
        });
      }
    });
  }

  /** 이 문서를 iframe으로 임베드한 부모 창(예: 업체별 접속현황 탭)이 자기 UI 안에 접속자 아바타를
   * 직접 그릴 수 있도록, webviewer 자신의 툴바(chrome=hidden이면 안 보임)와는 별개로 접속자 목록을
   * 그대로 postMessage로도 흘려보낸다 — webviewer는 그 목록을 "어디에" 보여줄지 전혀 몰라도 된다. */
  /** 참여자 이름·칸 내용은 이 화면을 띄운 부모 창(출처)에만 보낸다 — 부모 출처를 모르면(리퍼러를 숨긴 부모) 예전처럼 '*' */
  function parentOrigin(): string {
    const ancestors = (location as Location & { ancestorOrigins?: DOMStringList }).ancestorOrigins;
    if (ancestors && ancestors.length > 0) return ancestors[0];
    try {
      if (document.referrer) return new URL(document.referrer).origin;
    } catch {
      /* 리퍼러가 주소가 아니면 아래로 */
    }
    return '*';
  }

  function postPresenceToParent(): void {
    if (window.parent === window) return;
    const states = [...provider.awareness.getStates().entries()] as Array<[number, AwarenessUserState]>;
    const users = dedupeByName(states).map(([, state]) => ({
      name: state.user!.name!,
      color: state.user!.color ?? pickPresenceColorForName(state.user!.name!),
    }));
    window.parent.postMessage({ type: 'docview:presence', users }, parentOrigin());
  }
  provider.awareness.on('change', () => {
    renderPresenceAvatars(presenceEl, provider.awareness);
    renderRemoteCellHighlights();
    postPresenceToParent();
  });
  renderPresenceAvatars(presenceEl, provider.awareness);
  postPresenceToParent();

  // 병합(merges)도 함께 — 병합은 지운 칸 값·가운데 맞춤과 한 트랜잭션이라, 빠지면 Ctrl+Z가 값만 되살리고 병합은 남는다
  const undoManager = new Y.UndoManager([cells, structure, styles, merges, condFormats]);
  const shell = initOfficeShell({
    token,
    docName: info.name,
    canEdit,
    canDownload: info.permissions.canDownload,
    canPrint: info.permissions.canPrint,
    onUndo: () => undoManager.undo(),
    onRedo: () => undoManager.redo(),
  });

  /**
   * 시트 자리 번호 → 마지막으로 관찰한 행/열 수. structure.observe가 "행/열 삽입·삭제가 실제로
   * 일어났다"를 시트 이름 변경·탭 순서 변경 같은 다른 구조 갱신과 구별하는 유일한 기준이다
   * (리뷰 지적 I-4의 undo 스택 무효화 참고).
   */
  const lastSeenSheetExtents = new Map<string, { rowCount: number; colCount: number }>();
  structure.forEach((s, key) => lastSeenSheetExtents.set(key, { rowCount: s.rowCount, colCount: s.colCount }));

  /**
   * 이번 structure 이벤트에서 기존 시트의 행/열 **수**가 바뀌었는지 — 확인하면서 기준값을 갱신한다.
   *
   * 처음 보는 키(협업 중 새로 추가된 시트)는 "바뀐 것"으로 세지 않는다: 시트는 항상 뒤에 붙고
   * 자리 번호가 밀리지 않으므로(room.ts의 removedSheetIndices 주석) 기존 셀 키가 가리키는 대상이
   * 달라지지 않는다.
   */
  function consumeSheetExtentChange(): boolean {
    let changed = false;
    structure.forEach((s, key) => {
      const prev = lastSeenSheetExtents.get(key);
      if (prev && (prev.rowCount !== s.rowCount || prev.colCount !== s.colCount)) changed = true;
      lastSeenSheetExtents.set(key, { rowCount: s.rowCount, colCount: s.colCount });
    });
    return changed;
  }

  let activeSheet = 0;
  /** 마지막으로 선택(focus)된 칸 — 편집 중이 아니어도 계속 유지된다(툴바 서식 버튼이 "마지막
   * 선택 칸"에 적용되어야 하므로). */
  let activeCellKey: string | null = null;
  /** 지금 실제로 편집 중(contenteditable=true, 텍스트 커서 보임)인 칸. 엑셀처럼 칸을 클릭만 하면
   * "선택"만 되고, F2·더블클릭·직접 타이핑을 해야 이 값이 채워지며 편집모드로 들어간다. */
  let editingKey: string | null = null;
  /** 편집 중이라 건너뛴 구조 변경 재렌더가 밀려 있는지 — structure.observe/flushPendingStructureRender 참고. */
  let pendingStructureRender = false;
  /** 다중 셀 드래그 선택의 두 꼭짓점 — 드래그 안 하고 그냥 칸 하나만 클릭/이동했으면 앵커=포커스로
   * 같아서(1×1 범위) 별도 배경 없이 기존 :focus 테두리만 보인다. */
  let selAnchorRow = 0;
  let selAnchorCol = 0;
  let selFocusRow = 0;
  let selFocusCol = 0;

  /** 셀의 표시용 값을 얻는다 — 수식 셀이면 서버가 계산해 둔 값(computed 맵)을, 아니면 원본
   * 텍스트(cells 맵)를 그대로 쓴다. 클라이언트는 더 이상 수식을 직접 계산하지 않는다(서버가
   * canonical 전체 데이터 기준으로 계산해야 화면 밖 참조도 정확하기 때문 — 설계 문서의 핵심
   * 결정 참고). */
  function computeCellDisplay(sheetIdx: number, row: number, col: number): unknown {
    const key = gridCellKey(sheetIdx, row, col);
    const computedValue = computed.get(key);
    if (computedValue !== undefined) return computedValue;
    const raw = cells.get(key) ?? '';
    if (raw.startsWith('=')) return ''; // 서버 계산이 아직 도착 전(그 페이지가 막 로드된 직후) — 잠시 빈 값
    if (raw.trim() === '') return '';
    const num = Number(raw);
    return Number.isFinite(num) ? num : raw;
  }

  function formatComputedValue(value: unknown): string {
    if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
    if (typeof value !== 'number') return value === null || value === undefined ? '' : String(value);
    // 부동소수점 오차(0.1+0.2=0.30000000000000004류)를 다듬는다
    return String(Math.round(value * 1e10) / 1e10);
  }

  /** 소수 자릿수별 숫자 서식기 — toLocaleString은 칸마다 서식기를 새로 만들어 스크롤이 느려진다 */
  const numberFormatters = new Map<number, Intl.NumberFormat>();
  function numberFormatter(decimals: number): Intl.NumberFormat {
    let fmt = numberFormatters.get(decimals);
    if (!fmt) {
      fmt = new Intl.NumberFormat('ko-KR', { minimumFractionDigits: decimals, maximumFractionDigits: decimals });
      numberFormatters.set(decimals, fmt);
    }
    return fmt;
  }

  /** 원본 값(또는 수식 계산 결과)에 셀의 숫자 표시 형식(쉼표/통화/백분율)을 입힌다 — 저장된 값
   * 자체는 절대 안 바꾸고 "화면에 어떻게 보일지"만 바꾼다(수식 계산 결과 표시와 같은 원리). */
  function formatNumberForDisplay(text: string, style: CellStyleAttrs | undefined): string {
    // 날짜 형식 칸에 숫자(날짜 일련번호 — 예: =A4+7의 계산 결과)가 오면 엑셀처럼 날짜로 보여준다.
    // 날짜 텍스트 자체("2026-07-05")는 숫자가 아니라 그대로 통과한다.
    if (style?.df && text.trim() !== '' && Number.isFinite(Number(text))) return serialToDateText(Number(text));
    if (!style?.nf || text.trim() === '') return text;
    const num = Number(text);
    if (!Number.isFinite(num)) return text;
    if (num === 0) return ''; // 값 없음(0)은 화면을 어지럽히지 않게 빈 칸으로 보여준다
    const fmt = numberFormatter(style.nd ?? 0);
    switch (style.nf) {
      case 'percent':
        return `${fmt.format(num * 100)}%`;
      case 'currency':
        return `₩${fmt.format(num)}`;
      case 'number':
        return fmt.format(num);
      default:
        return text;
    }
  }

  /** 하이퍼링크(바로가기)·목록 상자 — 아래 setupGridLinks로 채운다(그 전에 그려진 칸은 받은 뒤 다시 표시) */
  let gridLinks: GridLinks | null = null;

  /** 셀 하나를 "지금 이 셀이 포커스 중인지"에 맞춰 원본 수식 또는 계산 결과로 그린다. */
  function renderCellDisplay(td: HTMLTableCellElement, key: string): void {
    const raw = cells.get(key) ?? '';
    if (editingKey === key) {
      // 편집 중에는 계산 결과가 아니라 실제 입력한 수식이 보여야 한다 — contentEditable 커서/선택
      // 범위가 span으로 감싸면 꼬일 수 있어 이 모드에서는 스필 span을 쓰지 않고 그대로 넣는다.
      td.replaceChildren();
      td.textContent = raw;
      return;
    }
    const style = styles.get(key);
    const text = !raw.startsWith('=')
      ? formatNumberForDisplay(raw, style)
      : formatNumberForDisplay(
          formatComputedValue(
            (() => {
              const parsed = parseGridCellKey(key);
              return parsed ? computeCellDisplay(parsed.sheetIdx, parsed.row, parsed.col) : raw;
            })(),
          ),
          style,
        );
    applySpillDisplay(td, key, text);
    condFormatOverlay?.(td); // 값이 바뀌면 강조도 바뀐다(이 경로는 applyStyleToTd를 안 거칠 때가 있다)
    if (gridLinks) {
      const parsed = parseGridCellKey(key);
      if (parsed && parsed.sheetIdx === activeSheet) gridLinks.decorate(td, parsed.row, parsed.col);
    }
  }

  /** 실제 엑셀처럼, 스필(오른쪽 빈 칸 위로 값이 넘쳐 보이는 동작)은 기본이 왼쪽 정렬인 칸에서만
   * 일어난다 — 가운데/오른쪽 정렬이 걸린 칸(예: SUM 등 계산 결과에 가운데 정렬을 걸어 둔 칸)은
   * 실제 엑셀에서도 절대 옆 칸으로 넘치지 않고 자기 칸 폭 안에서 잘린다. 예전엔 이 구분 없이 값
   * 있는 칸이면 무조건 폭을 넓혔는데, 그러면 넓어진 span의 text-align이 가운데/오른쪽으로
   * 상속되면서 글자가 원래 칸에서 몇 칸이나 떨어진 위치에 그려져 "엉뚱한 칸에 값이 보인다"는
   * 사용자 신고로 이어졌다(칸 자체의 위치·데이터는 항상 맞았다 — 순수 표시 버그).
   *
   * "이 칸의 유효 정렬"을 판단할 때 명시적 정렬(style.a)이 있으면 그대로 따르고, 없으면 엑셀의
   * 기본 규칙(숫자/날짜/수식 결과는 기본 오른쪽 정렬, 일반 텍스트는 기본 왼쪽 정렬)을 흉내낸다.
   * 이 시점엔 원본 값의 타입 정보가 없고 화면에 그려질 텍스트(text)만 있으므로, 우선 스타일에
   * 적힌 숫자 서식(style.nf)을 보고, 없으면 표시 텍스트 자체가(쉼표·통화 기호·퍼센트·공백·
   * 회계식 음수 괄호를 걷어내고 보면) 유한한 숫자로 읽히는지를 폴백으로 쓴다.
   *
   * 날짜 칸도 엑셀 기본은 오른쪽 정렬이라 스필 대상이 아니다. classifyNumFmt(src/spreadsheet.ts)는
   * 의도적으로 날짜를 style.nf 범주에 넣지 않으므로(그 문서 주석 참고 — nf를 건드리면 이 파일의
   * 서식 드롭다운(cell-numfmt-select)·xlsx 재내보내기(collab.ts의 numFmtStringFor) 쪽에 더 눈에
   * 띄는 부작용이 생긴다) 여기서도 표시 텍스트의 "모양"만으로 판정한다: 이 앱이 날짜를 그리는
   * 경로는 formatDate(src/spreadsheet.ts) 단 하나뿐이고, 원본 서식 코드가 무엇이든 항상 이
   * 모양(YYYY-MM-DD, 시:분이 있으면 그 뒤에 공백+HH:MM)으로만 렌더링되므로 이 정규식과 일치하는
   * 텍스트는 항상 날짜다. */
  const DATE_LIKE_TEXT = /^\d{4}-\d{2}-\d{2}( \d{2}:\d{2})?$/;
  function isSpillEligible(style: CellStyleAttrs | undefined, text: string): boolean {
    if (style) {
      if (style.a === 'c' || style.a === 'r') return false; // 명시적 가운데/오른쪽 정렬 — 절대 안 넘친다
      if (style.a === 'l') return true; // 명시적 왼쪽 정렬 — 항상 넘친다
      if (style.nf) return false; // 숫자 서식이 걸려 있으면 기본 오른쪽 정렬
    }
    if (text === 'TRUE' || text === 'FALSE') return false; // 논리값도 엑셀 기본은 가운데 정렬
    if (DATE_LIKE_TEXT.test(text.trim())) return false; // 날짜도 엑셀 기본은 오른쪽 정렬(위 주석 참고)
    const trimmed = text.trim();
    // 회계 서식의 음수 괄호 표기 "(1,234)"도 숫자로 인식되게 부호로 풀어준다
    const unwrapped =
      trimmed.length > 1 && trimmed.startsWith('(') && trimmed.endsWith(')') ? `-${trimmed.slice(1, -1)}` : trimmed;
    const stripped = unwrapped.replace(/[₩$€£¥,%\s]/g, '');
    const looksNumeric = stripped !== '' && Number.isFinite(Number(stripped));
    return !looksNumeric;
  }

  /** 실제 엑셀처럼 — 오른쪽 이웃 칸이 비어 있으면 이 칸의 표시 내용이 그 위로 넘쳐 보이게 한다
   * (값 있는 칸이나 병합 칸을 만나면 거기서 멈춘다). td.textContent를 직접 쓰지 않고 안쪽
   * span(.g-cell-spill)의 너비만 늘려야 td 경계 밖으로 시각적으로 넘칠 수 있다(CSS 쪽 참고).
   * 단, 기본이 왼쪽 정렬인 칸만 넓힌다(isSpillEligible 참고) — getComputedStyle 같은 강제
   * 레이아웃 계산은 쓰지 않는다(아래 캐시 관련 주석과 같은 이유로, 칸마다 그런 호출이 들어가면
   * 레이아웃 스래싱으로 직결된다). 판정은 이미 들고 있는 style 객체와 문자열만으로 끝난다. */
  function applySpillDisplay(td: HTMLTableCellElement, key: string, text: string): void {
    let span = td.querySelector<HTMLSpanElement>(':scope > .g-cell-spill');
    if (!span) {
      span = document.createElement('span');
      span.className = 'g-cell-spill';
      td.replaceChildren(span);
    }
    span.textContent = text;
    span.style.width = '';
    span.style.marginLeft = '';
    if (!text.trim() || td.colSpan > 1) return; // 병합 칸은 이미 colspan으로 넓어져 있어 스필이 필요 없다
    // 줄 바꿈 칸은 넘치지 않고 자기 칸 안에서 줄을 바꾼다(엑셀과 같다). 가운데 정렬 글자는 양옆 빈 칸으로
    // 고르게, 오른쪽 정렬 글자(숫자 아님)는 왼쪽 빈 칸으로 넘친다.
    const cellStyle = styles.get(key);
    if (cellStyle?.wrap) return;
    const spillMode: 'right' | 'left' | 'center' | null =
      cellStyle?.a === 'c' ? 'center' : cellStyle?.a === 'r' ? (NUMERIC_VALUE.test(text.trim()) ? null : 'left') : isSpillEligible(cellStyle, text) ? 'right' : null;
    if (!spillMode) return;
    const parsed = parseGridCellKey(key);
    if (!parsed) return;
    const { sheetIdx, row, col } = parsed;
    const sheet = structure.get(String(sheetIdx));
    if (!sheet) return;
    const colWidthOf = (c: number) => sheet.colWidths?.[c] ?? sheet.defaultColWidth ?? DEFAULT_COL_WIDTH_PX;
    // renderCountsFor()는 scrollHostEl.clientWidth/clientHeight를 읽는데, 이 값들은 레이아웃이
    // 최신 상태여야만 정확해서 브라우저가 그 자리에서 강제로 레이아웃을 다시 계산한다(forced
    // synchronous layout). 스크롤 패치(updateScrollWindow)처럼 한 번에 수십~수백 칸을 새로 그리는
    // 도중 renderCellDisplay→applySpillDisplay가 칸마다 이걸 부르면, DOM을 막 바꾼 직후 레이아웃을
    // 다시 읽는 패턴(레이아웃 스래싱)이 칸 수만큼 반복되어 렌더 비용이 실측상 눈에 띄게 늘었다.
    // 지금 칸이 속한 시트가 마지막 실제 렌더(vw* 캐시)와 같다면 그 열 수·병합 조회표를 그대로
    // 쓴다 — 스필 폭은 어차피 "지금 화면에 그려진 레이아웃" 기준이라 매번 새로 재는 것보다 이
    // 캐시가 더 일관되게 맞다. (시트 전환 직후처럼 캐시가 아직 안 맞을 때만 예전처럼 새로 계산한다.)
    const useCache = vwValid && sheetIdx === vwSheetIdx;
    const colCount = useCache ? vwRenderColCount : renderCountsFor(sheet).cols;
    const { anchors, covered } = useCache ? { anchors: vwAnchors, covered: vwCovered } : buildMergeLookup(sheetIdx);
    const emptyAt = (cc: number) =>
      cc >= 0 &&
      cc < colCount &&
      !covered.has(`${row}:${cc}`) &&
      !anchors.has(`${row}:${cc}`) &&
      (cells.get(gridCellKey(sheetIdx, row, cc)) ?? '').trim() === '';
    let right = 0;
    if (spillMode !== 'left') for (let cc = col + 1; emptyAt(cc); cc += 1) right += colWidthOf(cc);
    let left = 0;
    if (spillMode !== 'right') for (let cc = col - 1; emptyAt(cc); cc -= 1) left += colWidthOf(cc);
    if (spillMode === 'center') left = right = Math.min(left, right);
    if (left + right > 0) span.style.width = `${colWidthOf(col) + left + right}px`;
    if (left > 0) span.style.marginLeft = `-${left}px`;
  }

  /** key가 가리키는 칸이 바뀌면, 그 칸 자신뿐 아니라 왼쪽에서 그 칸으로 넘쳐 들어오고 있었을 수
   * 있는 가장 가까운 "값 있는" 칸도 스필 폭을 다시 계산해야 한다(그 칸이 비었다가 채워지면 왼쪽
   * 칸의 넘침이 줄어들어야 하고, 반대로 비면 다시 늘어나야 한다). */
  function refreshSpillNeighbor(key: string): void {
    const parsed = parseGridCellKey(key);
    if (!parsed) return;
    const { sheetIdx, row, col } = parsed;
    if (vwValid && !vwRowEls.has(row)) return; // 그려지지 않은 행 — 다시 그릴 칸이 없다
    const { covered } = buildMergeLookup(sheetIdx);
    for (let cc = col - 1; cc >= 0; cc -= 1) {
      if (covered.has(`${row}:${cc}`)) continue;
      const leftKey = gridCellKey(sheetIdx, row, cc);
      if ((cells.get(leftKey) ?? '').trim() === '') continue;
      const leftTd = tdFor(row, cc);
      if (leftTd && editingKey !== leftKey) renderCellDisplay(leftTd, leftKey);
      break;
    }
  }

  /** 어떤 셀이든 값이 바뀌면, 그 값을 참조할 수 있는 다른 수식 셀들도 다시 계산해 보여준다
   * (의존성 그래프 없이 "지금 그려진 수식 셀 전부 다시 계산"으로 단순하게 처리 — 편집기 규모에서
   * 충분히 빠르다). */
  function recalcFormulaCells(): void {
    // 합계(N열)처럼 수식이 들어있는 칸은 대개 잠겨 있어(.g-editable이 아님) 여기서 빠지면 값이
    // 바뀌어도 화면이 갱신되지 않는다 — 잠금 여부와 무관하게 "실제 칸"이면(data-row가 있으면) 전부 훑는다.
    tableEl.querySelectorAll<HTMLTableCellElement>('td[data-row]').forEach((td) => {
      const row = Number(td.dataset.row);
      const col = Number(td.dataset.col);
      const key = gridCellKey(activeSheet, row, col);
      if (editingKey === key) return;
      if ((cells.get(key) ?? '').startsWith('=')) renderCellDisplay(td, key);
    });
  }

  /** 화면 순서는 order 필드 기준(없으면 시트 키 번호로 대체) — 키 번호는 cells/styles/merges가
   * 참조하는 안정적인 식별자라 순서가 바뀌어도 그대로 두고, order만 바꿔서 재배치한다. */
  function sortedSheetIndices(): number[] {
    return [...structure.keys()]
      .map(Number)
      .sort((a, b) => (structure.get(String(a))?.order ?? a) - (structure.get(String(b))?.order ?? b));
  }
  function tabElFor(idx: number): HTMLElement | null {
    return sheetTabsEl.querySelector(`[data-sheet-idx="${idx}"]`);
  }
  function closeSheetPopup(): void {
    document.querySelectorAll('.sheet-context-menu, .sheet-move-menu, .grid-context-menu').forEach((el) => el.remove());
  }

  function switchToSheet(idx: number): void {
    if (activeSheet === idx) return;
    activeSheet = idx;
    resetSelectionForSheetSwitch();
    renderSheetTabs();
    renderTableVirtualized();
  }

  /** 엑셀처럼 탭 줄은 스크롤바 없이 잘려 보이고, 활성 탭이 잘려 있으면 보이는 곳까지 탭 줄을 민다 */
  function revealActiveSheetTab(): void {
    const tab = tabElFor(activeSheet);
    if (!tab) return;
    const left = tab.offsetLeft; // .grid-sheet-tabs--clip 이 position:relative 라 탭 줄 기준 위치
    const right = left + tab.offsetWidth;
    if (left < sheetTabsEl.scrollLeft) sheetTabsEl.scrollLeft = left;
    else if (right > sheetTabsEl.scrollLeft + sheetTabsEl.clientWidth) sheetTabsEl.scrollLeft = right - sheetTabsEl.clientWidth;
  }

  // 맨 왼쪽 ◀ ▶ — 엑셀처럼 활성 시트는 그대로 두고 탭 줄만 누른 방향으로 한 탭씩 민다(더 갈 곳이 없으면 꺼진다).
  // + 는 탭 줄 바깥에 있어 시트가 많아도 늘 보인다.
  const sheetNavPrev = document.getElementById('sheet-nav-prev') as HTMLButtonElement | null;
  const sheetNavNext = document.getElementById('sheet-nav-next') as HTMLButtonElement | null;
  function scrollSheetTabs(direction: -1 | 1): void {
    const tabs = [...sheetTabsEl.querySelectorAll<HTMLElement>('.grid-sheet-tab')];
    const at = sheetTabsEl.scrollLeft;
    const target =
      direction > 0
        ? tabs.find((tab) => tab.offsetLeft > at + 1) // 지금 맨 앞 탭 다음 탭이 맨 앞으로
        : [...tabs].reverse().find((tab) => tab.offsetLeft < at - 1); // 왼쪽에 가려진 바로 앞 탭이 맨 앞으로
    if (!target) return;
    const max = sheetTabsEl.scrollWidth - sheetTabsEl.clientWidth;
    sheetTabsEl.scrollLeft = Math.max(0, Math.min(max, target.offsetLeft));
    updateSheetMore();
  }
  function updateSheetNav(): void {
    const max = sheetTabsEl.scrollWidth - sheetTabsEl.clientWidth;
    if (sheetNavPrev) sheetNavPrev.disabled = sheetTabsEl.scrollLeft <= 0;
    if (sheetNavNext) sheetNavNext.disabled = sheetTabsEl.scrollLeft >= max - 1;
  }
  sheetNavPrev?.addEventListener('click', () => scrollSheetTabs(-1));
  sheetNavNext?.addEventListener('click', () => scrollSheetTabs(1));
  document.getElementById('sheet-add')?.addEventListener('click', () => addSheet());
  // 엑셀처럼 탭 줄 왼쪽/오른쪽에 가려진 시트가 있으면 그쪽 끝에 "…" — 누르면 가려진 바로 옆 시트로 간다
  const sheetMoreLeft = document.getElementById('sheet-more-left') as HTMLButtonElement | null;
  const sheetMoreRight = document.getElementById('sheet-more-right') as HTMLButtonElement | null;
  function hiddenTabs(): { left: number[]; right: number[] } {
    const view = sheetTabsEl.getBoundingClientRect();
    const left: number[] = [];
    const right: number[] = [];
    sheetTabsEl.querySelectorAll<HTMLElement>('.grid-sheet-tab').forEach((tab) => {
      const r = tab.getBoundingClientRect();
      const idx = Number(tab.dataset.sheetIdx);
      if (r.left < view.left - 1) left.push(idx);
      else if (r.right > view.right + 1) right.push(idx);
    });
    return { left, right };
  }
  function updateSheetMore(reveal = false): void {
    if (!sheetMoreLeft || !sheetMoreRight) return;
    if (reveal) revealActiveSheetTab();
    updateSheetNav();
    const { left, right } = hiddenTabs();
    const showLeft = left.length > 0;
    const showRight = right.length > 0;
    if (sheetMoreLeft.hidden !== !showLeft || sheetMoreRight.hidden !== !showRight) {
      sheetMoreLeft.hidden = !showLeft;
      sheetMoreRight.hidden = !showRight;
      requestAnimationFrame(() => updateSheetMore(reveal)); // … 가 생기고 빠지면 탭 줄 폭이 바뀌므로 한 번 더 잰다
    }
  }
  sheetMoreLeft?.addEventListener('click', () => {
    const { left } = hiddenTabs();
    if (left.length) switchToSheet(left[left.length - 1]);
  });
  sheetMoreRight?.addEventListener('click', () => {
    const { right } = hiddenTabs();
    if (right.length) switchToSheet(right[0]);
  });
  sheetTabsEl.addEventListener('scroll', () => updateSheetMore());
  new ResizeObserver(() => updateSheetMore(true)).observe(sheetTabsEl.parentElement ?? sheetTabsEl);
  sheetTabsEl.addEventListener(
    'wheel',
    (event) => {
      if (sheetTabsEl.scrollWidth <= sheetTabsEl.clientWidth) return;
      event.preventDefault();
      sheetTabsEl.scrollLeft += event.deltaY || event.deltaX;
    },
    { passive: false },
  );

  function renderSheetTabs(): void {
    const indices = sortedSheetIndices();
    sheetTabsEl.replaceChildren(
      ...indices.map((idx) => {
        const sheet = structure.get(String(idx))!;
        const tab = document.createElement('button');
        tab.type = 'button';
        tab.className = 'grid-sheet-tab' + (idx === activeSheet ? ' active' : '');
        tab.textContent = sheet.name;
        tab.dataset.sheetIdx = String(idx);
        tab.addEventListener('click', () => switchToSheet(idx));
        tab.addEventListener('dblclick', () => startSheetRename(idx));
        tab.addEventListener('contextmenu', (event) => {
          event.preventDefault();
          openSheetContextMenu(idx, event.clientX, event.clientY);
        });
        return tab;
      }),
    );
    updateSheetMore(true);
  }

  /** order 필드만 다시 매긴다 — 셀 데이터는 전혀 건드리지 않으므로 시트가 몇 개든 가벼운 연산이다. */
  function applySheetOrder(orderedKeys: readonly number[]): void {
    ydoc.transact(() => {
      orderedKeys.forEach((key, i) => {
        const sheet = structure.get(String(key));
        if (sheet && sheet.order !== i) structure.set(String(key), { ...sheet, order: i });
      });
    });
  }

  /** 엑셀의 "새 시트"/시트 탭 우클릭 "삽입"과 동일 — 기존 이름과 안 겹치는 "SheetN" 이름으로 빈
   * 시트를 만들어 바로 전환한다. afterIdx를 주면 그 시트 바로 뒤에, 안 주면 맨 뒤에 넣는다. */
  function addSheet(afterIdx?: number): void {
    const indices = sortedSheetIndices();
    const newKey = indices.length ? Math.max(...indices) + 1 : 0;
    const existingNames = new Set(indices.map((idx) => structure.get(String(idx))?.name));
    let n = indices.length + 1;
    while (existingNames.has(`Sheet${n}`)) n += 1;
    const insertPos = afterIdx === undefined ? indices.length : indices.indexOf(afterIdx) + 1;
    const orderedKeys = [...indices];
    orderedKeys.splice(insertPos, 0, newKey);
    activeSheet = newKey;
    resetSelectionForSheetSwitch();
    ydoc.transact(() => {
      structure.set(String(newKey), { name: `Sheet${n}`, rowCount: 1, colCount: 1, order: insertPos });
      orderedKeys.forEach((key, i) => {
        if (key === newKey) return;
        const sheet = structure.get(String(key));
        if (sheet && sheet.order !== i) structure.set(String(key), { ...sheet, order: i });
      });
    });
  }

  /** 엑셀 시트 탭 우클릭 "삭제" — 마지막 남은 시트는 지울 수 없고, 지금 보고 있던 시트를 지우면
   * 그 자리에 있던(원래 바로 뒤였던) 시트로 자동 전환한다. */
  async function deleteSheet(idx: number): Promise<void> {
    const indices = sortedSheetIndices();
    if (indices.length <= 1) {
      await officeAlert('마지막 남은 시트는 삭제할 수 없습니다.');
      return;
    }
    const sheet = structure.get(String(idx));
    if (!(await officeConfirm(`"${sheet?.name ?? ''}" 시트를 삭제하시겠습니까? 이 시트의 모든 내용이 사라집니다.`, { title: '시트 삭제', okLabel: '삭제', danger: true }))) return;
    if (!sortedSheetIndices().includes(idx)) return; // 기다리는 동안 다른 사람이 이미 지웠다
    const pos = indices.indexOf(idx);
    const remaining = indices.filter((i) => i !== idx);
    if (activeSheet === idx) {
      activeSheet = remaining[Math.min(pos, remaining.length - 1)];
      resetSelectionForSheetSwitch();
    }
    ydoc.transact(() => {
      structure.delete(String(idx));
      remaining.forEach((key, i) => {
        const s = structure.get(String(key));
        if (s && s.order !== i) structure.set(String(key), { ...s, order: i });
      });
    });
  }

  /** 엑셀 시트 탭 우클릭/더블클릭 "이름 바꾸기" — 탭 자리에 바로 입력창을 띄운다. */
  function startSheetRename(idx: number): void {
    const tab = tabElFor(idx);
    const sheet = structure.get(String(idx));
    if (!tab || !sheet) return;
    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'grid-sheet-tab-rename';
    input.value = sheet.name;
    tab.replaceWith(input);
    input.focus();
    input.select();
    const commit = () => {
      const trimmed = input.value.trim();
      if (trimmed && trimmed !== sheet.name) {
        structure.set(String(idx), { ...sheet, name: trimmed });
      } else {
        renderSheetTabs(); // 값이 그대로면 observe가 안 불려 렌더가 안 되므로 직접 되돌린다
      }
    };
    input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') input.blur();
      if (event.key === 'Escape') {
        input.value = sheet.name;
        input.blur();
      }
    });
    input.addEventListener('blur', commit, { once: true });
  }

  /** 엑셀 시트 탭 우클릭 "이동/복사"의 단순화된 버전 — 목록에서 기준 시트를 고르면 그 앞으로
   * 옮긴다(드래그 대신 클릭 한 번으로 재배치). */
  function openMoveMenu(idx: number, anchorEl: HTMLElement): void {
    closeSheetPopup();
    const indices = sortedSheetIndices();
    const menu = document.createElement('div');
    menu.className = 'sheet-move-menu';
    indices.forEach((targetIdx) => {
      if (targetIdx === idx) return;
      const targetSheet = structure.get(String(targetIdx));
      const item = document.createElement('button');
      item.type = 'button';
      item.textContent = `"${targetSheet?.name ?? ''}" 앞으로 이동`;
      item.addEventListener('click', () => {
        const reordered = indices.filter((i) => i !== idx);
        reordered.splice(reordered.indexOf(targetIdx), 0, idx);
        applySheetOrder(reordered);
        closeSheetPopup();
      });
      menu.appendChild(item);
    });
    const toEnd = document.createElement('button');
    toEnd.type = 'button';
    toEnd.textContent = '맨 뒤로 이동';
    toEnd.addEventListener('click', () => {
      applySheetOrder([...indices.filter((i) => i !== idx), idx]);
      closeSheetPopup();
    });
    menu.appendChild(toEnd);
    document.body.appendChild(menu);
    const anchorRect = anchorEl.getBoundingClientRect();
    positionSheetPopup(menu, anchorRect.left, anchorRect.bottom + 4, anchorRect.top - 4);
    setTimeout(() => document.addEventListener('mousedown', onSheetPopupOutsideClick), 0);
  }

  /** 뷰포트 아래쪽에 붙어 있는 시트 탭(항상 그렇다)에서 메뉴를 열면 아래로 펼칠 공간이 부족한
   * 경우가 대부분이라, 실제로 렌더링된 높이를 재서 넘치면 위로 뒤집는다. */
  function positionSheetPopup(menu: HTMLElement, x: number, belowY: number, aboveY: number = belowY): void {
    const rect = menu.getBoundingClientRect();
    const fitsBelow = belowY + rect.height <= window.innerHeight;
    const top = fitsBelow ? belowY : Math.max(4, aboveY - rect.height);
    const left = Math.max(4, Math.min(x, window.innerWidth - rect.width - 4));
    menu.style.left = `${left}px`;
    menu.style.top = `${top}px`;
  }

  function onSheetPopupOutsideClick(event: MouseEvent): void {
    const popups = document.querySelectorAll('.sheet-context-menu, .sheet-move-menu, .grid-context-menu');
    const inside = [...popups].some((el) => el.contains(event.target as Node));
    if (!inside) {
      closeSheetPopup();
      document.removeEventListener('mousedown', onSheetPopupOutsideClick);
    }
  }

  /** 엑셀 시트 탭 우클릭 메뉴와 동일한 4가지: 삽입/삭제/이름 바꾸기/이동. */
  function openSheetContextMenu(idx: number, x: number, y: number): void {
    closeSheetPopup();
    const menu = document.createElement('div');
    menu.className = 'sheet-context-menu';
    const items: Array<[string, () => void]> = [
      ['삽입', () => addSheet(idx)],
      ['삭제', () => void deleteSheet(idx)],
      ['이름 바꾸기', () => startSheetRename(idx)],
      ['이동', () => openMoveMenu(idx, tabElFor(idx) ?? menu)],
    ];
    items.forEach(([label, action]) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        if (label !== '이동') closeSheetPopup();
        action();
      });
      menu.appendChild(btn);
    });
    document.body.appendChild(menu);
    positionSheetPopup(menu, x, y);
    setTimeout(() => document.addEventListener('mousedown', onSheetPopupOutsideClick), 0);
  }

  /**
   * 행/열 삽입·삭제 — 클라이언트는 **의도만 서버에 알리고** 실제 이동은 하지 않는다.
   *
   * **왜 직접 옮기지 않는가**(예전에는 shiftRowsAndMerges/shiftColsAndMerges가 Yjs의 cells/styles
   * 키를 한 칸씩 직접 옮겼다):
   *
   * 1. 그 방식은 "이 시트의 모든 셀이 Yjs에 들어 있다"를 전제한다. 지금 Yjs는 **구독 중인 500행
   *    페이지만 담는 창**이라(sendPageSubscription 참고) 창 밖의 행은 애초에 옮길 수가 없다 —
   *    삽입 지점 아래의 나머지 시트 전체가 위쪽과 한 칸씩 어긋난 채로 남는다.
   * 2. 그 루프의 상한은 시트 실제 크기가 아니라 **하드코딩된 1000**이었다(1000행 넘는 시트를
   *    지원하기 전 시대의 코드). 1001행이 넘는 시트에서는 행 1000의 내용을 덮어쓰고 1001행
   *    이후는 손대지도 않았다.
   *
   * 서버는 구독 상태와 무관하게 시트 전체를 담은 canonical 그리드를 들고 있으므로(room.ts의
   * applyGridStructureOp) 거기서 옮기면 두 문제가 동시에 사라진다. 이동 결과는 평소의 Yjs
   * 업데이트로 되돌아오고, 이미 있는 관찰자들이 화면을 맞춘다 — cells.observe가 바뀐 칸을 다시
   * 그리고, structure.observe가 새 행/열 수로 표를 다시 그린다(서버가 구조 메타데이터를 일부러
   * **마지막에** 보내므로 그 재렌더 시점엔 셀 데이터가 이미 도착해 있다).
   *
   * 알려진 대가: 행/열 삽입·삭제는 더 이상 Ctrl+Z(undoManager)로 되돌릴 수 없다 — UndoManager는
   * 로컬 트랜잭션만 추적하고 이 변경은 서버에서 온다. 삭제된 행의 내용을 되살리려면 서버 쪽
   * 역연산이 필요해 이번 범위 밖으로 둔다(예전 방식의 undo도 1000행 넘는 시트에서는 이미
   * 깨져 있었다).
   */
  /** room.ts의 MESSAGE_GRID_STRUCTURE_OP와 **같은 값이어야 한다**(0~3은 y-websocket 표준
   * 프로토콜이 이미 쓰고 있다 — 그쪽 주석 참고). */
  const MESSAGE_GRID_STRUCTURE_OP = 4;

  function sendGridStructureOp(
    kind: 'insertRows' | 'deleteRows' | 'insertCols' | 'deleteCols',
    at: number,
    count = 1,
  ): void {
    if (!canEdit) return;
    // sendPageSubscription과 완전히 같은 전송 방식(같은 WebSocket에 varUint 타입 + varString JSON).
    if (!provider.ws || provider.ws.readyState !== WebSocket.OPEN) return;
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_GRID_STRUCTURE_OP);
    encoding.writeVarString(encoder, JSON.stringify({ sheet: activeSheet, kind, at, count }));
    provider.ws.send(encoding.toUint8Array(encoder));
  }

  function insertRowAt(row: number): void {
    if (!structure.get(String(activeSheet))) return;
    sendGridStructureOp('insertRows', row);
  }
  function deleteRowAt(row: number): void {
    const sheet = structure.get(String(activeSheet));
    // 마지막 남은 행은 지우지 않는다 — 서버도 같은 검사를 하지만(applyGridStructureOp) 여기서
    // 먼저 막아 왕복 없이 곧바로 아무 일도 일어나지 않게 한다.
    if (!sheet || sheet.rowCount <= 1) return;
    sendGridStructureOp('deleteRows', row);
  }
  function insertColumnAt(col: number): void {
    if (!structure.get(String(activeSheet))) return;
    sendGridStructureOp('insertCols', col);
  }
  function deleteColumnAt(col: number): void {
    const sheet = structure.get(String(activeSheet));
    if (!sheet || sheet.colCount <= 1) return;
    sendGridStructureOp('deleteCols', col);
  }

  function openRowContextMenu(row: number, x: number, y: number): void {
    closeSheetPopup();
    const menu = document.createElement('div');
    menu.className = 'grid-context-menu sheet-context-menu';
    const items: Array<[string, () => void]> = [
      ['위에 행 삽입', () => insertRowAt(row)],
      ['행 삭제', () => deleteRowAt(row)],
    ];
    items.forEach(([label, action]) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        closeSheetPopup();
        action();
      });
      menu.appendChild(btn);
    });
    document.body.appendChild(menu);
    positionSheetPopup(menu, x, y);
    setTimeout(() => document.addEventListener('mousedown', onSheetPopupOutsideClick), 0);
  }

  function openColumnContextMenu(col: number, x: number, y: number): void {
    closeSheetPopup();
    const menu = document.createElement('div');
    menu.className = 'grid-context-menu sheet-context-menu';
    const items: Array<[string, () => void]> = [
      ['왼쪽에 열 삽입', () => insertColumnAt(col)],
      ['열 삭제', () => deleteColumnAt(col)],
      ['오름차순 정렬', () => sortWholeSheetBy(col, false)],
      ['내림차순 정렬', () => sortWholeSheetBy(col, true)],
    ];
    items.forEach(([label, action]) => {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.textContent = label;
      btn.addEventListener('click', () => {
        closeSheetPopup();
        action();
      });
      menu.appendChild(btn);
    });
    document.body.appendChild(menu);
    positionSheetPopup(menu, x, y);
    setTimeout(() => document.addEventListener('mousedown', onSheetPopupOutsideClick), 0);
  }

/** 셀 근처에 잠깐 떴다 사라지는 안내 말풍선 — 유효성 검사에 걸린 입력을 왜 안 받았는지 알려준다. */
  function showCellTooltip(td: HTMLTableCellElement, message: string): void {
    const rect = td.getBoundingClientRect();
    const tip = document.createElement('div');
    tip.className = 'g-cell-tooltip';
    tip.textContent = message;
    tip.style.left = `${rect.left}px`;
    tip.style.top = `${rect.bottom + 4}px`;
    document.body.appendChild(tip);
    setTimeout(() => tip.remove(), 2200);
  }

  /** 데이터가 없는 빈 셀에 값을 입력하면 시트 범위(structure)를 자동으로 넓힌다 — 엑셀처럼
   * "칸이 있으니까 그냥 입력하면 늘어난다"는 동작을 흉내낸다. */
  function commitCell(td: HTMLTableCellElement): void {
    const row = Number(td.dataset.row);
    const col = Number(td.dataset.col);
    const key = gridCellKey(activeSheet, row, col);
    const value = td.textContent ?? '';
    const current = cells.get(key) ?? '';
    if (value === current) return;

    // 이 칸이 숫자 전용(엑셀 데이터 유효성 검사)이면 숫자가 아닌 입력은 저장하지 않고 안내만 띄운다
    // — blur 이후 renderCellDisplay가 다시 그리면서 원래 값으로 자연히 되돌아간다.
    if (styles.get(key)?.nu === 1 && value.trim() !== '' && !NUMERIC_VALUE.test(value.trim())) {
      showCellTooltip(td, '숫자만 입력할 수 있습니다.');
      return;
    }

    ydoc.transact(() => {
      if (value === '') {
        cells.delete(key);
        return;
      }
      // 엑셀처럼 "1,000"/"50%"/"₩1,000" 입력은 텍스트가 아니라 숫자로 저장하고, 칸에 숫자 서식이
      // 없으면 입력 모양에 맞는 서식을 붙인다(예전엔 텍스트로 저장돼 이 칸을 참조하는 수식이 깨지고
      // 저장한 엑셀에서도 "텍스트 숫자"가 됐다). 이미 서식이 있는 칸은 그 서식을 존중한다.
      const typed = value.trim().startsWith('=') ? null : parseTypedNumber(value);
      if (typed) {
        cells.set(key, String(typed.value));
        const currentStyle = styles.get(key) ?? {};
        if (!currentStyle.nf && !currentStyle.df) {
          styles.set(key, { ...currentStyle, nf: typed.kind, ...(typed.decimals ? { nd: typed.decimals } : {}) });
        }
      } else {
        cells.set(key, value);
      }
      const s = currentStructure();
      if (row >= s.rowCount || col >= s.colCount) {
        structure.set(String(activeSheet), {
          ...s,
          rowCount: Math.max(s.rowCount, row + 1),
          colCount: Math.max(s.colCount, col + 1),
        });
      }
    });
  }

  function isPasteTargetEditable(row: number, col: number): boolean {
    if (!canEdit) return false;
    return styles.get(gridCellKey(activeSheet, row, col))?.lk !== 1;
  }

  /** 실제 엑셀에서 여러 칸을 복사해 붙여넣을 때를 처리한다. 이 <td>가 contentEditable이라 아무
   * 처리도 안 하면 브라우저가 클립보드의 HTML(엑셀이 함께 넣어주는 <table> 조각)을 이 한 칸 안에
   * 그대로 밀어 넣어버려 값이 한 칸에 뭉쳐 들어가거나 깨진다 — 붙여넣기를 직접 가로채
   * 클립보드의 plain text(엑셀은 항상 탭=열, 줄바꿈=행 구분의 TSV도 같이 넣어준다)를 파싱해
   * 시작 칸부터 각 칸에 정확히 나눠 넣는다. */
  function handlePaste(event: ClipboardEvent, anchorTd: HTMLTableCellElement): void {
    const text = event.clipboardData?.getData('text/plain');
    if (!text) return;
    event.preventDefault();
    pasteTextAt(text, Number(anchorTd.dataset.row), Number(anchorTd.dataset.col));
  }

  /** TSV 글을 (startRow, startCol)부터 칸마다 나눠 넣는다 — Ctrl+V(handlePaste)와 홈 탭 "붙여넣기" 단추가 함께 쓴다. */
  function pasteTextAt(text: string, startRow: number, startCol: number): void {
    const rows = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    while (rows.length > 1 && rows[rows.length - 1] === '') rows.pop();
    const grid = rows.map((line) => line.split('\t'));
    const maxCols = Math.max(...grid.map((r) => r.length));

    ydoc.transact(() => {
      grid.forEach((values, rOffset) => {
        values.forEach((rawValue, cOffset) => {
          const row = startRow + rOffset;
          const col = startCol + cOffset;
          if (!isPasteTargetEditable(row, col)) return; // 잠긴 칸·읽기 전용 문서는 조용히 건너뛴다
          const key = gridCellKey(activeSheet, row, col);
          const value = rawValue.trim();
          if (styles.get(key)?.nu === 1 && value !== '' && !NUMERIC_VALUE.test(value)) return; // 숫자 전용 칸에 숫자가 아니면 건너뜀
          if (value === '') cells.delete(key);
          else cells.set(key, value);
        });
      });
      const s = currentStructure();
      const maxRow = startRow + grid.length - 1;
      const maxCol = startCol + maxCols - 1;
      if (maxRow >= s.rowCount || maxCol >= s.colCount) {
        structure.set(String(activeSheet), {
          ...s,
          rowCount: Math.max(s.rowCount, maxRow + 1),
          colCount: Math.max(s.colCount, maxCol + 1),
        });
      }
    });

    grid.forEach((values, rOffset) => {
      values.forEach((_value, cOffset) => {
        const row = startRow + rOffset;
        const col = startCol + cOffset;
        const targetTd = tdFor(row, col);
        if (targetTd) renderCellDisplay(targetTd, gridCellKey(activeSheet, row, col));
      });
    });
    recalcFormulaCells();
  }

  /** 데이터가 있는 범위 너머로도 빈 칸을 넉넉히 그려서, 실제 엑셀처럼 "빈 시트도 격자가 있는"
   * 느낌을 준다 — 그 빈 칸에 입력하면 commitCell이 자동으로 범위를 넓혀준다. */
  const GRID_ROW_PADDING = 15;
  const GRID_COL_PADDING = 6;
  const GRID_MIN_RENDER_ROWS = 25;
  const GRID_MIN_RENDER_COLS = 12;
  /** 시트에 기본 열 폭이 없을 때(구버전 문서) — 엑셀 기본 열 폭(64px) */
  const DEFAULT_COL_WIDTH_PX = 64;
  /** .g-edit-rownum의 min-width와 맞춘 대략치 — 아래 tableEl.style.width 계산에만 쓰이고, 실제
   * 행번호 열 폭은 여전히 CSS(min-width)가 정한다. 정확히 안 맞아도 상관없다: 테이블에 "auto가
   * 아닌" width를 준다는 사실 자체가 중요하지, 합계가 1px이라도 틀리면 안 되는 게 아니다. */
  const ROWNUM_COL_WIDTH_PX = 28;
  /** 행 번호 열 폭 — 엑셀처럼 마지막 행 번호의 자릿수만큼 넓힌다(28px 고정이면 6자리 행 번호가 옆 칸 선을 넘었다).
   * 12px 고정폭 숫자 한 자리 ≈ 7px + 좌우 여백. */
  function rownumWidthFor(rowCount: number): number {
    return Math.max(ROWNUM_COL_WIDTH_PX, String(Math.max(1, rowCount)).length * 7 + 10);
  }
  /** 가상 스크롤(윈도잉 렌더링)이 화면 위아래로 실제 뷰포트보다 더 그려두는 여유 행 수 — 스크롤이
   * 살짝만 움직여도 매 프레임 다시 그릴 필요 없게 완충 역할을 한다(너무 작으면 빠르게 스크롤할 때
   * 빈 줄이 잠깐 보이고, 너무 크면 DOM 노드가 늘어나 다시 느려진다). */
  const OVERSCAN_ROWS = 50;

  /** 시트 크기만으로는 "확대율을 낮추면 화면에 빈 칸이 더 많이 보여야 한다"(실제 엑셀은 격자가
   * 사실상 무한이라 화면에 보이는 만큼 항상 그려진다)는 걸 반영 못 한다 — 지금 뷰포트/줌 배율로
   * 실제 보이는 픽셀 영역을 재서, 그걸 채우는 데 필요한 행/열 수까지 함께 반영한다. CSS zoom은
   * 확대율이 낮을수록 이 컨테이너의 clientWidth/clientHeight를(줌이 적용되기 전 좌표계 기준으로)
   * 더 크게 보고하므로, 별도로 zoom 값을 나눌 필요 없이 그대로 쓰면 된다. */
  function renderCountsFor(sheet: SheetStructure): { rows: number; cols: number } {
    const availWidth = scrollHostEl.clientWidth;
    const availHeight = scrollHostEl.clientHeight;
    let usedWidth = rownumWidthFor(sheet.rowCount + GRID_ROW_PADDING);
    let fillCols = 0;
    while (usedWidth < availWidth && fillCols < 1000) {
      usedWidth += sheet.colWidths?.[fillCols] ?? sheet.defaultColWidth ?? DEFAULT_COL_WIDTH_PX;
      fillCols += 1;
    }
    const fillRows = Math.max(0, rowLayout.rowAt(availHeight));
    return {
      rows: Math.max(sheet.rowCount + GRID_ROW_PADDING, GRID_MIN_RENDER_ROWS, fillRows),
      cols: Math.max(sheet.colCount + GRID_COL_PADDING, GRID_MIN_RENDER_COLS, fillCols),
    };
  }

  /** 병합 정보(1-based)를 0-based (row:col) 조회표로 바꾼다 — viewer.js의 buildMergeIndex와 동일한 원리.
   * anchors: 병합의 왼쪽위 칸 → {rs,cs}(rowspan/colspan). covered: 그 외 병합에 덮인 칸(<td>를 아예 안 그림). */
  function buildMergeLookup(sheetIdx: number): {
    anchors: Map<string, { rs: number; cs: number }>;
    covered: Set<string>;
  } {
    const anchors = new Map<string, { rs: number; cs: number }>();
    const covered = new Set<string>();
    const sheetMerges = merges.get(String(sheetIdx)) ?? [];
    for (const m of sheetMerges) {
      const r0 = m.r - 1;
      const c0 = m.c - 1;
      anchors.set(`${r0}:${c0}`, { rs: m.rs, cs: m.cs });
      for (let r = r0; r < r0 + m.rs; r += 1) {
        for (let c = c0; c < c0 + m.cs; c += 1) {
          if (r === r0 && c === c0) continue;
          covered.add(`${r}:${c}`);
        }
      }
    }
    return { anchors, covered };
  }

  /** ---------- 스크롤 전용 증분 패치 상태 ----------
   * renderTableVirtualized()는 스크롤이 한 번 일어날 때마다 <colgroup>/<thead>/보이는 모든
   * <tr>·<td>를 통째로 버리고 다시 만들었다 — 각 <td>에 붙는 5~8개의 이벤트 리스너(focus/
   * mousedown/blur/dblclick/input/paste/compositionstart/keydown)까지 매 프레임 떼고 다시
   * 붙이는 셈이라, 실측(Chrome longtask) 기준 한 번에 54~108ms가 걸렸다(16ms 프레임 예산의
   * 3~6배). 스크롤만으로 인한 재호출은 "화면에 보이는 행 창(window)"만 바뀔 뿐 시트/구조/
   * 틀고정/병합/열 폭은 그대로이므로, updateScrollWindow()가 직전에 그렸던 행 창과 새 행 창을
   * 비교해 새로 들어온 행만 만들고 나간 행만 지우고 나머지는 기존 DOM 노드(와 거기 걸린 리스너)를
   * 그대로 재사용한다. 시트 전환·구조 변경·확대축소처럼 실제로 다시 그려야 하는 경우는 여전히
   * renderTableVirtualized()의 전체 재빌드를 그대로 타므로(sync/시트탭 클릭/structure.observe/
   * scheduleViewportRefit 등은 지금처럼 renderTableVirtualized를 직접 부른다) 정확성 리스크가
   * 없다 — 이 캐시는 "지금 화면에 그려진 것과 정확히 같은 문맥(시트/열 폭/틀고정/병합)에서
   * 스크롤만 됐다"는 게 보장된 경우에만 쓰이고, 그 전제가 조금이라도 의심되면(vwValid===false거나
   * 시트/열·행 수가 다르면) 곧바로 전체 재빌드로 물러난다. */
  let vwValid = false;
  let vwSheetIdx = -1;
  let vwRenderColCount = 0;
  let vwRenderRowCount = 0;
  let vwFreezeRows = 0;
  let vwFreezeCols = 0;
  let vwAnchors: Map<string, { rs: number; cs: number }> = new Map();
  let vwCovered: Set<string> = new Set();
  let vwFirstVisible = -1;
  let vwLastVisible = -1;
  let vwTbody: HTMLTableSectionElement | null = null;
  let vwTopSpacer: HTMLTableRowElement | null = null;
  let vwBottomSpacer: HTMLTableRowElement | null = null;
  /** r(0-based 행 번호) → 지금 DOM에 있는 그 행의 <tr> — 틀 고정 행(0..freezeRows-1)과 지금
   * 스크롤 창 안의 행이 모두 들어있다(스페이서는 여기 포함되지 않는다 — 리스너가 없어 재사용할
   * 이유가 없으므로 updateScrollWindow는 스페이서를 매번 새로 만든다). */
  let vwRowEls: Map<number, HTMLTableRowElement> = new Map();

  /** 화면 밖(위/아래) 구간 전체를 대신 차지하는 "안 보이는" 스페이서 한 행 — buildRow 설명 참고.
   * rowCount<=0(그릴 게 없는 구간)이면 아예 만들지 않는다. 리스너가 없는 단순한 <tr>/<td> 한 쌍이라
   * updateScrollWindow()도 재사용하지 않고 매번 새로 만든다(재사용을 위한 위치 계산 비용이 굳이
   * 만드는 비용보다 크다). */
  function buildSpacerRow(height: number): HTMLTableRowElement | null {
    if (height <= 0) return null;
    const heightPx = `${height}px`;
    const tr = document.createElement('tr');
    tr.setAttribute('aria-hidden', 'true');
    tr.style.height = heightPx;
    const td = document.createElement('td');
    td.colSpan = vwRenderColCount + 1; // 행번호 열까지 포함해 표 전체 폭을 덮는다
    // .grid-edit-table th/td에 공통으로 걸린 height:1.7rem(고정 행 높이용)을 인라인 스타일로
    // 덮어써야 스페이서가 의도한 실제 높이(수천 px일 수 있음)로 그려진다 — 인라인 스타일이
    // 클래스 규칙보다 항상 이긴다.
    td.style.height = heightPx;
    td.style.padding = '0';
    td.style.border = 'none';
    tr.appendChild(td);
    return tr;
  }

  /** 행 하나(<tr>)를 만든다 — 셀 생성·서식·이벤트 바인딩은 예전 renderTable과 동일하다. 이제는
   * 틀 고정 행이든 스크롤 윈도우 안의 행이든 항상 정상 흐름(<colgroup>이 적용되는 진짜 표 행)으로
   * 그린다 — 화면에 없는 구간은 이 함수가 아니라 buildSpacerRow가 대신 높이를 차지한다.
   * renderTableVirtualized()의 전체 재빌드와 updateScrollWindow()의 증분 패치 양쪽에서 함께
   * 쓰므로, 매개변수 대신 vw*(직전 렌더가 채워 둔) 캐시를 읽는다 — 두 호출부 모두 이 함수를 부르기
   * 전에 그 캐시가 지금 그리려는 행과 같은 문맥으로 맞춰져 있음을 보장한다. */
  function buildRow(r: number): HTMLTableRowElement {
    const tr = document.createElement('tr');
    // Chromium은 서로 다른 <tr> 소속인 position:sticky 셀들 사이의 z-index 비교를 신뢰할 수 없게
    // 처리한다(고정된 행의 배경보다 스크롤된 다른 행의 셀이 위로 그려지는 버그) — 각 행을
    // position+z-index로 그 자체의 스태킹 컨텍스트로 격리하면, 셀별 z-index는 그 행 안에서만
    // 비교되고 행 간 우열은 이 tr의 z-index로만 결정되어 늘 정확하다.
    tr.style.position = 'relative';
    tr.style.zIndex = r < vwFreezeRows ? '2' : '1';
    // 파일의 행 높이(숨긴 행은 0) — 칸 높이는 CSS가 아니라 이 값이 정한다
    const rowHeight = rowLayout.height(r);
    tr.style.height = `${rowHeight}px`;
    if (rowHeight === 0) tr.classList.add('g-hidden-row');
    const rowHead = document.createElement('th');
    rowHead.className = 'g-edit-rownum';
    rowHead.textContent = String(r + 1);
    rowHead.dataset.row = String(r);
    rowHead.addEventListener('click', (event) => selectRow(r, event.shiftKey));
    if (canEdit) {
      rowHead.addEventListener('contextmenu', (event) => {
        event.preventDefault();
        openRowContextMenu(r, event.clientX, event.clientY);
      });
    }
    tr.appendChild(rowHead);
    for (let c = 0; c < vwRenderColCount; c += 1) {
      if (vwCovered.has(`${r}:${c}`)) continue; // 병합에 덮인 칸 — 앵커 쪽에서 colspan/rowspan으로 대신 차지한다
      const key = gridCellKey(activeSheet, r, c);
      const td = document.createElement('td');
      const span = vwAnchors.get(`${r}:${c}`);
      if (span) {
        if (span.rs > 1) td.rowSpan = span.rs;
        if (span.cs > 1) td.colSpan = span.cs;
      }
      // 파일에서 숨긴 열 — 폭 0으로 자리만 지킨다
      if (!(span && span.cs > 1) && structure.get(String(vwSheetIdx))?.colWidths?.[c] === 0) td.classList.add('g-hidden-col');
      const isLocked = styles.get(key)?.lk === 1;
      const canEditThisCell = !isLocked && canEdit;
      // 클릭만으로는 "선택"만 되고 편집모드로 들어가지 않는다(엑셀과 동일) — 실제 편집모드는
      // enterEditMode()가 필요할 때만 이 값을 'true'로 바꾼다. 잠긴 칸/읽기전용 문서는 애초에
      // canEditThisCell이 false라 keydown 핸들러가 편집 진입 자체를 막는다.
      td.contentEditable = 'false';
      // 잠긴 칸도 화살표 키로 지나가거나 선택해서 볼 수 있어야 하므로(엑셀도 보호된 칸을 선택은
      // 할 수 있음) 편집 가능 여부와 무관하게 모든 칸을 포커스 가능하게 만든다. tabIndex=-1은 Tab
      // 키 순서에는 안 끼지만 클릭·프로그램적 focus()는 그대로 받는다.
      td.tabIndex = -1;
      if (canEditThisCell) td.classList.add('g-editable');
      td.dataset.row = String(r);
      td.dataset.col = String(c);
      // 엑셀 "틀 고정" — 위 N행/왼쪽 N열은 스크롤해도 계속 보여야 한다. 실제 sticky 좌표(top/left)는
      // 렌더링이 끝나 실제 셀 크기를 잴 수 있게 된 뒤 applyFreezePanes()에서 한 번에 계산한다.
      if (r < vwFreezeRows) td.classList.add('g-frozen-row');
      if (c < vwFreezeCols) td.classList.add('g-frozen-col');
      renderCellDisplay(td, key);
      applyStyleToTd(td, styles.get(key));
      // 잠겨 있으면서 맨 위 행에 있고 내용이 실제로 있는 칸(제목 배너, "조회쿼리" 버튼 등)은 클릭
      // 시 부모 창(이 문서를 iframe으로 임베드한 화면)에 알림만 보낸다 — 무엇을 보여줄지(예: 연도
      // 목록 드롭다운, 쿼리 팝업)는 전적으로 부모 쪽 책임이다. webviewer 자신은 "연도"/"쿼리" 같은
      // 특정 문서의 의미를 전혀 알 필요가 없고, 클릭된 칸의 위치·텍스트만 그대로 전달한다.
      const hasContent = (cells.get(key) ?? '').trim() !== '';
      if (isLocked && r === 0 && hasContent) {
        td.classList.add('g-clickable-header');
        // 맨 왼쪽(제목 배너)만 "목록에서 고른다"는 의미의 드롭다운 화살표를 보여준다 — 그 옆의
        // 다른 클릭 가능한 칸(조회쿼리 버튼 등)은 텍스트 자체에 "(클릭)"을 넣어 알려주므로
        // 화살표가 필요 없다.
        if (c === 0) td.classList.add('g-header-dropdown');
        td.addEventListener('click', () => {
          if (window.parent !== window) {
            // 셀의 화면상 위치(이 iframe 안에서의 좌표)를 같이 보내, 부모 쪽이 임베드한 iframe의
            // 위치와 합쳐 팝오버를 셀 바로 밑에 자연스럽게 놓을 수 있게 한다 — webviewer는 "연도
            // 팝오버를 어디에 그릴지" 자체는 몰라도 되고, 그저 클릭된 칸의 위치만 알려준다.
            const rect = td.getBoundingClientRect();
            window.parent.postMessage(
              {
                type: 'docview:title-click',
                row: r,
                col: c,
                text: cells.get(key) ?? '',
                rect: { top: rect.top, left: rect.left, bottom: rect.bottom, right: rect.right, width: rect.width, height: rect.height },
              },
              parentOrigin(),
            );
          }
        });
      }
      // 이 칸이 어느 "세대"(renderGeneration)에 만들어졌는지 기록해 둔다 — 아래 delegated
      // focusout의 stale 판정에 쓰인다: 스크롤 패치(updateScrollWindow)는 살아남는 칸을 다시
      // 만들지 않으므로 renderGeneration을 새로 올리지 않고, 그 사이 전체 재빌드가 한 번도
      // 없었다면 이 값은 여전히 "지금" 세대와 같다 — 즉 이 판정은 "이 특정 DOM 노드가 방금
      // 만들어졌는가"가 아니라 "마지막 전체 재빌드 이후로 편집 세션이 안전한가"를 묻는다.
      td.dataset.gen = String(renderGeneration);
      tr.appendChild(td);
    }
    return tr;
  }

  /** 이벤트 위임(delegation)에서 실제 <td>를 찾는다 — 화면에 보일 때는 스필 span(g-cell-spill)이
   * <td> 안에 들어있어 event.target이 그 span일 수 있으므로 closest로 조상을 거슬러 올라간다. */
  function cellFromEvent(event: Event): { td: HTMLTableCellElement; row: number; col: number; key: string } | null {
    const target = event.target as HTMLElement | null;
    const td = target?.closest<HTMLTableCellElement>('td[data-row][data-col]') ?? null;
    if (!td) return null;
    const row = Number(td.dataset.row);
    const col = Number(td.dataset.col);
    return { td, row, col, key: gridCellKey(activeSheet, row, col) };
  }

  /**
   * <td>마다 5~8개씩(focus/mousedown/blur/dblclick/input/paste/compositionstart/keydown) 걸던
   * 리스너를 표 전체에 단 한 번만 거는 위임 리스너로 대체한다 — buildRow()가 스크롤 패치
   * (updateScrollWindow)로 자주 다시 불리므로, 리스너를 셀마다 붙였다 뗐다 하는 비용 자체를 렌더
   * 경로에서 없앤다. 각 핸들러는 event.target에서 실제 칸을 역추적해(cellFromEvent) 예전에
   * r/c/key/canEditThisCell 클로저로 갖고 있던 정보를 그때그때 다시 구한다(canEditThisCell은
   * buildRow가 같은 시점에 붙여 둔 g-editable 클래스로 복원한다).
   */
  tableEl.addEventListener('focusin', (event) => {
    const hit = cellFromEvent(event);
    if (!hit) return;
    const { td, row: r, col: c, key } = hit;
    const canEditThisCell = td.classList.contains('g-editable');
    activeCellKey = key;
    selAnchorRow = r;
    selAnchorCol = c;
    selFocusRow = r;
    selFocusCol = c;
    renderRangeHighlight();
    syncCellToolbarState(key);
    setActiveHeaders(r, c);
    cellRefBoxEl.textContent = cellRefLabel(r, c);
    formulaInputEl.value = cells.get(key) ?? '';
    positionFillHandle(canEditThisCell ? td : null);
  });
  // 다중 셀 드래그 선택 시작 — 편집 중인 칸 안에서는 텍스트 드래그 선택이 우선이라 건드리지 않는다.
  tableEl.addEventListener('mousedown', (event) => {
    if (event.button !== 0) return;
    const hit = cellFromEvent(event);
    if (!hit) return;
    const { td, row: r, col: c, key } = hit;
    if (editingKey === key) return;
    if (event.shiftKey && !editingKey) {
      // Shift+클릭 — 지금 칸(기준)은 그대로, 범위 끝만 누른 칸으로(엑셀과 같다). 포커스를 옮기지 않아야
      // focusin이 범위를 한 칸으로 접지 않는다. 누른 채 끌면 이어서 넓힌다.
      event.preventDefault();
      extendSelectionTo(r, c);
      document.addEventListener('mousemove', onRangeDragMove);
      document.addEventListener('mouseup', onRangeDragEnd);
      return;
    }
    // 기본 동작을 그대로 두면 브라우저가 "텍스트 드래그 선택"도 같이 시작해버린다 — 그건 사각형이
    // 아니라 문서(행) 순서로 훑어서, 계산한 범위보다 훨씬 넓게(중간 행의 다른 칸까지) 파랗게/회색으로
    // 잡혀 보인다. preventDefault로 그 기본 동작을 끄고, 그때 같이 꺼지는 포커스만 아래에서 직접
    // 다시 걸어준다.
    event.preventDefault();
    clearHeaderSelection();
    selAnchorRow = r;
    selAnchorCol = c;
    selFocusRow = r;
    selFocusCol = c;
    renderRangeHighlight();
    td.focus();
    document.addEventListener('mousemove', onRangeDragMove);
    document.addEventListener('mouseup', onRangeDragEnd);
  });
  // focus의 위임 버전인 focusin과 대칭으로, blur 대신 focusout(버블링됨)에 건다 — 이미 아래에
  // "그리드 영역을 완전히 벗어났는지"만 보는 별개의 focusout 리스너가 있고, 이건 그것과 별도로
  // "이 칸이 편집 중이었다면 커밋한다"만 담당한다.
  tableEl.addEventListener('focusout', (event) => {
    const hit = cellFromEvent(event);
    if (!hit) return;
    const { td, key } = hit;
    if (editingKey !== key) return; // 편집 중이 아니었으면(그냥 선택만) 커밋할 것이 없다
    // 붙여넣기/자동채우기로 시트 범위(rowCount/colCount)가 늘어나거나 스크롤로 가상 렌더가 다시
    // 그리면 renderTableVirtualized()가 트랜잭션 도중 동기적으로 표 전체를 다시 그리며 이 <td>를
    // DOM에서 떼어내는데, 포커스가 있던 칸이 떨어져 나가면 브라우저가 focusout을 보낸다 — 그
    // 시점의 td.textContent는 그 작업 "전" 값이라, 그대로 commitCell을 부르면 방금 쓴 값을
    // 덮어써 버린다. 제거 시점과 focusout 디스패치 시점이 완전히 동기적이진 않아 isConnected만으로는
    // 못 거르므로, "이 칸이 만들어진 세대"가 지금도 최신인지로 판정한다(buildRow의 data-gen 참고).
    const cellGeneration = Number(td.dataset.gen);
    if (cellGeneration !== renderGeneration) {
      editingKey = null;
      provider.awareness.setLocalStateField('cell', null);
      return;
    }
    commitCell(td);
    editingKey = null;
    provider.awareness.setLocalStateField('cell', null);
    td.contentEditable = 'false';
    renderCellDisplay(td, key); // 편집 종료 — 수식이면 계산 결과로 되돌린다
    recalcFormulaCells(); // 이 셀 값이 바뀌었으니 이를 참조하는 다른 수식 셀도 갱신
    setGridStatusText('준비');
    flushPendingStructureRender(); // 편집 중이라 미뤄 뒀던 구조 변경이 있으면 지금 반영한다
  });
  tableEl.addEventListener('dblclick', (event) => {
    const hit = cellFromEvent(event);
    if (!hit || !hit.td.classList.contains('g-editable')) return;
    enterEditMode(hit.td, hit.key);
  });
  tableEl.addEventListener('input', (event) => {
    const hit = cellFromEvent(event);
    if (!hit || !hit.td.classList.contains('g-editable')) return;
    if (activeCellKey === hit.key) formulaInputEl.value = hit.td.textContent ?? '';
  });
  tableEl.addEventListener('paste', (event) => {
    const hit = cellFromEvent(event);
    if (!hit || !hit.td.classList.contains('g-editable')) return;
    handlePaste(event, hit.td);
  });
  // 셀을 선택만 한 채(더블클릭 없이) 바로 한글 등 IME 문자를 치는 경우 — 조합 중인 글자는 keydown의
  // event.key가 길이 1이 아니라("Process" 등) handleCellKeydown의 "글자 하나 치면 편집모드 진입"
  // 분기를 못 타서, 셀이 그대로 선택 상태에 머물며 입력이 통째로 사라졌었다. compositionstart
  // 시점에 직접 편집모드로 들어가면(내용은 비워서 시작 — 엑셀처럼 선택된 채 타이핑하면 기존 값을
  // 덮어쓴다) 그 뒤 IME 조합이 이 칸에서 이어진다.
  tableEl.addEventListener('compositionstart', (event) => {
    const hit = cellFromEvent(event);
    if (!hit || !hit.td.classList.contains('g-editable')) return;
    if (editingKey !== hit.key) enterEditMode(hit.td, hit.key, '');
  });
  tableEl.addEventListener('keydown', (event) => {
    const hit = cellFromEvent(event);
    if (!hit) return;
    const canEditThisCell = hit.td.classList.contains('g-editable');
    handleCellKeydown(event, hit.td, hit.row, hit.col, hit.key, canEditThisCell);
  });

  /** 전체 행만큼 <tr>을 만들던 예전 renderTable()을 대체한다 — 시트가 아무리 커도 실제 DOM에는
   * "지금 보이는 범위 + 여유분(OVERSCAN_ROWS)"만큼의 행만 만든다.
   *
   * (이전 시도였던 position:absolute + top 방식은 실제 브라우저에서 깨진다는 게 리뷰에서 밝혀졌다
   * — position:absolute인 <tr>은 계산된 display가 표준상 block으로 바뀌어 그 자손 <td>가 표의 열
   * 그리드(<colgroup>/table-layout:fixed)에서 완전히 떨어져 나가고, 브라우저가 그 <td>들을 별도의
   * 익명 미니 테이블로 감싸 table-layout:auto로 재배치한다 — 그 안의 인라인 width는 "힌트"일 뿐
   * 강제되지 않아, 실제 데이터가 있는 마지막 열이 내용 크기 기준으로 크게 부풀고 그 뒤 열들이
   * 화면 밖으로 밀려났다. 셀마다 인라인 width를 주는 우회책은 근본 원인을 가리지 못했다.)
   *
   * 지금 방식(스페이서 행): 보이는 행(<tr>)은 전부 정상 흐름(position 지정 없음)으로 그려서
   * <colgroup>/table-layout:fixed 열 정렬을 그대로 받는다(틀 고정 행이 이미 하던 방식과 동일).
   * 화면에 없는 위/아래 구간은 실제 <tr>을 만드는 대신, 그 구간 전체 높이(안 그린 행 수 ×
   * 행 높이 합)를 가진 "스페이서" <tr> 하나씩으로 대신한다(buildSpacerRow). 그러면 표의 자연스러운
   * 문서 흐름만으로 전체 스크롤 높이(틀 고정 행 + 위 스페이서 + 보이는 행 + 아래 스페이서 =
   * 전체 시트 높이)가 정확히 나오고, 보이는 행들은 정상 흐름 안에서 실제로 그렸을 때와 똑같은
   * 자리에 놓인다 — 수동으로 top을 계산해 줄 필요가 없다. */
  function renderTableVirtualized(): void {
    renderTableVirtualizedInner();
    afterGridRender();
  }
  function renderTableVirtualizedInner(): void {
    renderGeneration += 1;
    // 다시 그리면 이전 <td>들은 전부 버려지므로, 그 순간 실제로 "편집 중"이던 칸도 이제 없다 —
    // blur가 안전하게 안 오는 경우에도(과거 stale-blur 버그) 여기서 확실히 리셋해 둔다.
    editingKey = null;
    // tableEl.replaceChildren()으로 표 내용을 통째로 비우면 그 순간 스크롤 가능한 실제 높이가
    // (다시 채우기 전까지) 뷰포트 높이 수준으로 줄어들어, 브라우저가 scrollHostEl.scrollTop을
    // 그 시점의 최댓값(사실상 0)으로 즉시 clamp해 버린다 — 실측 확인됨(Playwright로
    // replaceChildren() 전/후 scrollTop을 직접 비교: 67500 → 0, 표를 되돌려도 0에 머묾). 표를
    // 비우기 "전에" 지금 스크롤 위치를 미리 읽어 둬야, 아래에서 이 값으로 가상 스크롤 범위를
    // 계산할 때 방금 clamp된 값이 아니라 실제 스크롤 위치를 쓰게 된다 — 그렇지 않으면 스크롤이
    // 일어날 때마다(이 함수가 스크롤 리스너에서 다시 불릴 때마다) 화면이 맨 위로 튕겨 돌아가
    // 사실상 스크롤이 전혀 안 되는 것처럼 보인다.
    const scrollTopSnapshot = scrollHostEl.scrollTop;
    const viewportHeightSnapshot = scrollHostEl.clientHeight;
    const sheet = structure.get(String(activeSheet));
    rowLayout = new RowLayout(sheet?.defaultRowHeight ?? DEFAULT_ROW_HEIGHT_PX, sheet?.rowHeights ?? [], hiddenRowsSorted);
    // 전체 재빌드가 시작되면 지금까지의 증분 패치 캐시(vw*)는 곧 통째로 버려질 DOM을 가리키게
    // 되므로 즉시 무효화한다 — 이 함수가 아래서 일찍 return하더라도(sheet 없음) updateScrollWindow가
    // 낡은 참조로 패치를 시도하지 않고 안전하게 다시 전체 재빌드로 물러난다.
    vwValid = false;
    tableEl.replaceChildren();
    if (!sheet) return;
    tableEl.classList.toggle('g-no-gridlines', sheet.showGridLines === false);

    warningEl.hidden = !(sheet.rowCount > EDIT_SOFT_ROW_LIMIT || sheet.colCount > EDIT_SOFT_COL_LIMIT);
    const { rows: renderRowCount, cols: renderColCount } = renderCountsFor(sheet);
    const { anchors, covered } = buildMergeLookup(activeSheet);
    // freezeRows는 항상 정상 흐름으로 그려지므로, 실제 렌더 행 수를 넘지 않게 안전하게 자른다.
    const freezeRows = Math.min(sheet.freezeRows ?? 0, renderRowCount);
    const freezeCols = sheet.freezeCols ?? 0;
    // 아래 buildRow/buildSpacerRow(hoisted)는 이 문맥을 매개변수가 아니라 vw* 캐시로 읽는다 — 이
    // 재빌드가 끝난 뒤 updateScrollWindow()가 스크롤만으로 인한 재호출에서 같은 캐시를 그대로
    // 재사용해 전체 재빌드를 건너뛸 수 있게 하기 위해서다.
    vwSheetIdx = activeSheet;
    vwRenderColCount = renderColCount;
    vwRenderRowCount = renderRowCount;
    vwFreezeRows = freezeRows;
    vwFreezeCols = freezeCols;
    vwAnchors = anchors;
    vwCovered = covered;

    // table-layout:fixed에서 열 너비를 정확히 고정하려면 <colgroup>이 가장 확실하다(그 전까진
    // 첫 번째 행 셀 너비에 의존해 왔음 — 실제 엑셀 열 너비 정보가 없던 시절의 임시 방편이었다).
    // 행번호 열은 엑셀 열이 아니라 편집기가 그리는 것이라 colWidths에 대응하는 값이 없다 — 행 수 자릿수로 정한다(rownumWidthFor).
    const colgroup = document.createElement('colgroup');
    const rownumWidth = rownumWidthFor(renderRowCount);
    const rownumCol = document.createElement('col');
    rownumCol.style.width = `${rownumWidth}px`;
    colgroup.appendChild(rownumCol);
    let totalColWidthPx = rownumWidth;
    for (let c = 0; c < renderColCount; c += 1) {
      const width = sheet.colWidths?.[c] ?? sheet.defaultColWidth ?? DEFAULT_COL_WIDTH_PX;
      const col = document.createElement('col');
      col.style.width = `${width}px`;
      colgroup.appendChild(col);
      totalColWidthPx += width;
    }
    tableEl.appendChild(colgroup);
    // Chromium은 table-layout:fixed + <colgroup>이어도, 테이블 자체 width가 auto인 채로 열이
    // 많으면(대략 12~15개 이상) 개별 <col> 폭을 조용히 무시하고 자체적으로 재분배해버린다(실측
    // 확인됨 — 열이 적을 때는 정상 동작). 테이블에 명시적 width를 주면(정확한 합계가 아니어도,
    // "auto가 아니다"라는 사실만으로) 이 경로를 피하고 각 열이 지정한 폭 그대로 그려진다.
    tableEl.style.width = `${totalColWidthPx}px`;

    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    headRow.style.position = 'relative';
    headRow.style.zIndex = '3';
    const cornerTh = document.createElement('th');
    // table-layout:fixed는 "첫 번째 행"의 셀 너비로 각 열 너비를 확정한다 — 행번호 열의 실제 좁은
    // 너비(.g-edit-rownum)를 적용하려면 헤더 행의 이 모서리 칸도 같은 클래스를 받아야 한다. 이걸
    // 빼먹으면 이 칸이 일반 칸 min-width(100px)를 그대로 물려받아 행번호 열 전체가 넓어져 버린다.
    cornerTh.className = 'g-edit-rownum';
    cornerTh.style.width = cornerTh.style.minWidth = `${rownumWidth}px`;
    headRow.appendChild(cornerTh);
    for (let c = 0; c < renderColCount; c += 1) {
      const th = document.createElement('th');
      th.textContent = colLetter(c);
      th.dataset.col = String(c);
      th.addEventListener('click', (event) => selectColumn(c, event.shiftKey));
      if (canEdit) {
        th.addEventListener('contextmenu', (event) => {
          event.preventDefault();
          openColumnContextMenu(c, event.clientX, event.clientY);
        });
      }
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    tableEl.appendChild(thead);

    // ---------- 가상 스크롤 범위 계산 ----------
    // (함수 맨 위에서 표를 비우기 전에 미리 읽어 둔 스냅샷을 쓴다 — 지금 다시 읽으면 방금
    // replaceChildren()이 clamp한 값(0)이 나온다.)
    const scrollTop = scrollTopSnapshot;
    const viewportHeight = viewportHeightSnapshot;
    const firstVisible = Math.max(freezeRows, rowLayout.rowAt(scrollTop) - OVERSCAN_ROWS);
    const lastVisible = Math.min(renderRowCount - 1, rowLayout.rowAt(scrollTop + viewportHeight) + OVERSCAN_ROWS);

    const tbody = document.createElement('tbody');

    // 틀 고정된 위쪽 행은 항상 정상 흐름으로(스크롤 위치와 무관하게 계속 보이도록) 그린다. 그
    // 아래, 화면 밖(위쪽)에 있는 안 그리는 구간은 스페이서 행 하나로 대신하고, 실제로 보이는
    // 범위(+여유분)만 진짜 행으로 그린 뒤, 화면 밖(아래쪽) 나머지 구간도 스페이서 행 하나로
    // 마무리한다. firstVisible은 이미 freezeRows 이상으로 보정돼 있다. buildRow가 만든 각 <tr>은
    // vwRowEls에도 기록해 둔다 — 이후 스크롤만으로 인한 재호출(updateScrollWindow)이 이 행들을
    // 재사용할 수 있게 하기 위해서다.
    const newRowEls = new Map<number, HTMLTableRowElement>();
    for (let r = 0; r < freezeRows; r += 1) {
      const tr = buildRow(r);
      newRowEls.set(r, tr);
      tbody.appendChild(tr);
    }
    const topSpacer = buildSpacerRow(rowLayout.top(firstVisible) - rowLayout.top(freezeRows));
    if (topSpacer) tbody.appendChild(topSpacer);
    for (let r = firstVisible; r <= lastVisible; r += 1) {
      if (hiddenRowSet.has(r)) continue;
      const tr = buildRow(r);
      newRowEls.set(r, tr);
      tbody.appendChild(tr);
    }
    const bottomSpacer = buildSpacerRow(rowLayout.top(renderRowCount) - rowLayout.top(lastVisible + 1));
    if (bottomSpacer) tbody.appendChild(bottomSpacer);
    tableEl.appendChild(tbody);
    // 표 내용이 다시 채워져 스크롤 가능한 실제 높이가 복구된 지금(스페이서 행 포함), 함수 맨 위의
    // clamp(스크롤 위치가 0으로 튕겨나간 것)를 되돌린다 — 이 시점 이후로는 (스페이서 덕분에) 그
    // 위치가 실제로 유효한 범위 안이라 clamp되지 않고 그대로 반영된다. 이미 같은 값이면 실제
    // 스크롤 이벤트가 다시 발생하지 않도록(불필요한 재렌더 방지) 값이 다를 때만 대입한다.
    if (scrollHostEl.scrollTop !== scrollTopSnapshot) scrollHostEl.scrollTop = scrollTopSnapshot;
    // 이 전체 재빌드가 방금 그린 것을 스크롤 전용 증분 패치(updateScrollWindow)가 이어받을 수
    // 있도록 캐시를 채운다 — vwValid는 여기서 성공적으로 다 그린 뒤에만 true가 된다.
    vwTbody = tbody;
    vwTopSpacer = topSpacer;
    vwBottomSpacer = bottomSpacer;
    vwRowEls = newRowEls;
    vwFirstVisible = firstVisible;
    vwLastVisible = lastVisible;
    vwValid = true;
    applyFreezePanes(freezeRows, freezeCols);
    renderRemoteCellHighlights();
    // document.hasFocus()가 false면 실제 포커스는 이 문서 밖에 있다 — 이 뷰어는 호스트 페이지에
    // iframe으로 박혀 돌아가는 것을 정식으로 지원하므로(DOCVIEW_ALLOWED_ORIGINS), 사용자가 호스트
    // 쪽 입력칸에 타이핑하는 동안에도 원격 변경·창 크기 변경·확대율 변경 같은 이유로 이 전체
    // 재빌드가 돌 수 있다. 그때 focus()를 부르면 그리드가 호스트의 포커스를 소리 없이 빼앗아,
    // 이어지는 타자가 엉뚱한 셀에 들어가 그대로 커밋된다(restoreScrollWindowFocus의 같은 가드와
    // 짝이다 — 거기 주석에 이유를 더 자세히 적어 뒀다).
    if (activeCellKey && document.hasFocus()) {
      const parsed = parseGridCellKey(activeCellKey);
      // 실제 DOM 포커스를 같은 칸의 새 <td>로 되돌린다 — focus 이벤트가 setActiveHeaders/채우기
      // 핸들 위치 재계산까지 같이 해 준다. 재렌더로 그 칸이 더 이상 없으면(시트 전환, 또는 가상
      // 스크롤 범위 밖으로 밀려남 등) 그냥 둔다.
      if (parsed && parsed.sheetIdx === activeSheet) {
        // focus()가 선택을 그 한 칸으로 접고 머리글 선택도 푼다 — 다시 그리기는 선택을 바꾸는 동작이 아니므로
        // (열 폭을 끌어 바꾼 뒤에도 고른 열이 그대로 있어야 한다) restoreScrollWindowFocus처럼 되돌린다
        const snapshot = { ar: selAnchorRow, ac: selAnchorCol, fr: selFocusRow, fc: selFocusCol, header: headerSelection };
        tdFor(parsed.row, parsed.col)?.focus({ preventScroll: true });
        selAnchorRow = snapshot.ar;
        selAnchorCol = snapshot.ac;
        selFocusRow = snapshot.fr;
        selFocusCol = snapshot.fc;
        headerSelection = snapshot.header;
        renderRangeHighlight();
        paintHeaderSelection();
      }
    }
    updateHScrollbar();
    document
      .getElementById('btn-freeze-panes')
      ?.setAttribute('aria-pressed', String(freezeRows > 0 || freezeCols > 0));
    // 지금 화면에 필요한 페이지(서버가 Yjs에 채워 둬야 하는 500행 단위 구간)를 서버에 알린다 —
    // 시트 전환·구조 변경·스크롤로 인한 재렌더 등 renderTableVirtualized가 불리는 모든 경로를
    // 여기 한 곳에서 커버한다(각 호출부에 일일이 추가하는 것보다 누락 위험이 적다).
    scheduleSubscriptionUpdate();
  }

  /**
   * 스크롤"만"으로 인한 재렌더 전용 — renderTableVirtualized()의 가벼운 대체판. <colgroup>/
   * <thead>는 스크롤로 바뀌지 않으므로 아예 건드리지 않고, tbody의 "행 창"도 직전에 그렸던 범위와
   * 겹치는 부분은 기존 <tr>(과 거기 달린 이벤트 리스너)을 그대로 두고, 창 밖으로 나간 행만 지우고
   * 새로 들어온 행만 만들어 끼워 넣는다. 이번 문맥이 직전 전체 재빌드 때와 정확히 같다는 확신이
   * 없으면(시트가 바뀌었다거나, 열/행 수가 달라졌다거나, tbody가 더 이상 표 안에 없다거나) 안전하게
   * renderTableVirtualized()의 전체 재빌드로 물러난다 — 그 판정 자체는 값 비교뿐이라 비용이 거의
   * 없으므로 "혹시 몰라서" 지나치게 자주 물러나도 손해는 아니다.
   */
  function updateScrollWindow(): void {
    if (!vwValid || !vwTbody || vwTbody.parentNode !== tableEl || vwSheetIdx !== activeSheet) {
      renderTableVirtualized();
      return;
    }
    const sheet = structure.get(String(activeSheet));
    if (!sheet) {
      renderTableVirtualized();
      return;
    }
    const { rows: renderRowCount, cols: renderColCount } = renderCountsFor(sheet);
    // 열/행 렌더 수가 달라졌다는 건 구조·확대율·창 크기 중 뭔가 바뀌었다는 뜻 — 그런 변경은 전부
    // 자기 경로에서 이미 renderTableVirtualized()를 직접 부르므로(structure.observe,
    // scheduleViewportRefit 등) 정상적으로는 여기서 다를 일이 없지만, 방어적으로 확인한다.
    if (renderRowCount !== vwRenderRowCount || renderColCount !== vwRenderColCount) {
      renderTableVirtualized();
      return;
    }

    const freezeRows = vwFreezeRows;
    const scrollTop = scrollHostEl.scrollTop;
    const viewportHeight = scrollHostEl.clientHeight;
    const firstVisible = Math.max(freezeRows, rowLayout.rowAt(scrollTop) - OVERSCAN_ROWS);
    const lastVisible = Math.min(renderRowCount - 1, rowLayout.rowAt(scrollTop + viewportHeight) + OVERSCAN_ROWS);
    if (firstVisible === vwFirstVisible && lastVisible === vwLastVisible) return; // 창이 그대로면 할 일 없음

    const tbody = vwTbody;
    // 스페이서는 리스너 없는 <tr>/<td> 한 쌍이라 재사용을 위한 위치 계산보다 새로 만드는 쪽이 싸다
    // — 매번 지우고 끝에서 다시 만든다(실제로 비싼 건 아래에서 그대로 두는 데이터 행이다).
    vwTopSpacer?.remove();
    vwBottomSpacer?.remove();

    // 창 밖으로 나간 행 제거(틀 고정 행은 스크롤과 무관하게 항상 남는다) — Map 이터레이션 도중
    // 지금 방문 중이거나 이미 지나간 항목을 delete하는 것은 안전하다(사양대로 동작).
    for (const [r, tr] of vwRowEls) {
      if (r < freezeRows) continue;
      if (r < firstVisible || r > lastVisible) {
        tr.remove();
        vwRowEls.delete(r);
      }
    }

    // 남아있는 행은 건드리지 않고, 새로 창에 들어온 행만 만들어 순서를 지켜 끼워 넣는다 — 앵커를
    // 오름차순으로 이어가며 "직전에 처리한(기존 또는 방금 만든) 행 바로 뒤"에 놓는다.
    let anchor: ChildNode | null = freezeRows > 0 ? vwRowEls.get(freezeRows - 1) ?? null : null;
    for (let r = firstVisible; r <= lastVisible; r += 1) {
      if (hiddenRowSet.has(r)) continue;
      const existing = vwRowEls.get(r);
      if (existing) {
        anchor = existing;
        continue;
      }
      const tr = buildRow(r);
      if (anchor) anchor.after(tr);
      else tbody.insertBefore(tr, tbody.firstChild);
      vwRowEls.set(r, tr);
      anchor = tr;
    }

    const topSpacer = buildSpacerRow(rowLayout.top(firstVisible) - rowLayout.top(freezeRows));
    if (topSpacer) {
      const firstWindowRow = vwRowEls.get(firstVisible) ?? null;
      if (firstWindowRow) tbody.insertBefore(topSpacer, firstWindowRow);
      else tbody.appendChild(topSpacer); // 방어적 처리 — 이론상 창이 비어 있을 때만 여기로 온다
    }
    const bottomSpacer = buildSpacerRow(rowLayout.top(renderRowCount) - rowLayout.top(lastVisible + 1));
    if (bottomSpacer) tbody.appendChild(bottomSpacer); // 항상 마지막 자식이 맞다(그 뒤엔 아무것도 없다)
    vwTopSpacer = topSpacer;
    vwBottomSpacer = bottomSpacer;
    vwFirstVisible = firstVisible;
    vwLastVisible = lastVisible;

    // 틀 고정 "열"은 스크롤 창 밖에 있던 새 행의 셀에도 sticky left를 걸어야 한다(틀 고정 행은
    // 창 밖으로 나가지 않으니 다시 계산할 필요가 없다 — applyFreezePanes는 두 축을 한 번에
    // 처리하지만 freezeRows>0 branch는 이미 있는 고정 행만 다시 훑어 idempotent하다). 틀 고정을
    // 아예 안 쓰는 보통 시트에서는 이 함수가 즉시 리턴해 비용이 없다.
    if (freezeRows > 0 || vwFreezeCols > 0) applyFreezePanes(freezeRows, vwFreezeCols);
    renderRemoteCellHighlights();
    // 전체 재빌드(renderTableVirtualized)가 맨 끝에서 하던 "활성 칸 포커스 되돌리기"와, 그 포커스
    // 복원이 focusin을 통해 곁들여 해 주던 "선택 표시 다시 칠하기"를 이 경로에도 똑같이 붙인다 —
    // 이게 없어서 (1) 스크롤 왕복 뒤 키 입력이 <body>로 새어 조용히 사라지고, (2) 화면에는 아무
    // 선택도 안 보이는데 툴바 서식은 예전 범위 전체에 먹는 문제가 있었다. 순서가 중요하다:
    // 포커스 복원이 먼저고(그 안에서 focusin이 하이라이트를 한 번 지운다), 칠하기가 나중이다.
    restoreScrollWindowFocus(firstVisible, lastVisible, freezeRows);
    repaintScrollWindowSelection();
  }

  // ---------- 셀 선택 vs 편집모드, 방향키 이동 (엑셀과 동일하게) ----------
  // 칸을 클릭하면(focus) "선택"만 된다 — contenteditable은 여전히 false다. 실제로 값을 고칠 수
  // 있으려면(1) F2, (2) 더블클릭, (3) 선택된 채로 문자를 바로 타이핑, 셋 중 하나로 편집모드에
  // 들어가야 한다. 편집모드에서 Enter/Tab은 커밋 후 아래/오른쪽 칸으로 넘어가고, Escape는 원래
  // 값으로 되돌리고 선택 상태로만 남는다.

  function placeCaretAtEnd(el: HTMLElement): void {
    const range = document.createRange();
    range.selectNodeContents(el);
    range.collapse(false);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
  }

  /** 지금 이 칸을 다른 참여자가 편집 중이면 그 사람 정보를 돌려준다(없으면 null) — 있으면 이 칸은
   * 편집모드로 못 들어간다(같은 칸을 동시에 고치는 충돌을 막는다. 서버가 강제하는 진짜 락은 아니고
   * awareness 브로드캐스트 기반의 느슨한 락이라, 정확히 동시에 누르는 극단적인 경우까지 완벽히
   * 막지는 못하지만 — 실사용에서 충분하다). */
  function remoteEditorOf(row: number, col: number): { name: string; color: string } | null {
    const states = provider.awareness.getStates() as Map<number, RemoteCellState>;
    for (const [clientId, state] of states) {
      if (clientId === ydoc.clientID) continue;
      const cell = state.cell;
      if (!cell || cell.sheetIdx !== activeSheet || cell.row !== row || cell.col !== col) continue;
      const name = state.user?.name;
      if (!name) continue;
      return { name, color: state.user?.color ?? pickPresenceColorForName(name) };
    }
    return null;
  }

  function enterEditMode(td: HTMLTableCellElement, key: string, replaceWith?: string): void {
    if (!td.classList.contains('g-editable')) return; // 잠긴 칸/읽기전용 문서는 편집모드 자체가 없다
    const parsed = parseGridCellKey(key);
    if (parsed) {
      const other = remoteEditorOf(parsed.row, parsed.col);
      if (other) {
        showCellTooltip(td, `${other.name}님이 편집 중입니다.`);
        return;
      }
    }
    editingKey = key;
    if (parsed) provider.awareness.setLocalStateField('cell', parsed); // "지금 편집 중" — 다른 참여자에게 표시
    td.contentEditable = 'true';
    td.textContent = replaceWith !== undefined ? replaceWith : cells.get(key) ?? '';
    td.focus();
    placeCaretAtEnd(td);
    formulaInputEl.value = td.textContent ?? '';
    // replaceWith가 주어졌다는 건 기존 값을 지우고 새로 치는 중(엑셀의 "입력")이라는 뜻 —
    // undefined면 더블클릭 등으로 기존 내용을 그대로 두고 이어 쓰는 "편집"이다.
    setGridStatusText(replaceWith !== undefined ? '입력' : '편집');
  }

  function commitAndMove(td: HTMLTableCellElement, key: string, row: number, col: number, dRow: number, dCol: number): void {
    commitCell(td);
    editingKey = null;
    provider.awareness.setLocalStateField('cell', null);
    td.contentEditable = 'false';
    renderCellDisplay(td, key);
    recalcFormulaCells();
    moveSelection(row, col, dRow, dCol);
    setGridStatusText('준비');
    flushPendingStructureRender();
  }

  function cancelEdit(td: HTMLTableCellElement, key: string): void {
    editingKey = null;
    provider.awareness.setLocalStateField('cell', null);
    td.contentEditable = 'false';
    renderCellDisplay(td, key); // 원래 값(수정 전)으로 되돌린다
    td.focus();
    setGridStatusText('준비');
    flushPendingStructureRender();
  }

  /** td.scrollIntoView는 sticky thead/행번호 열이 실제로 화면을 가리고 있다는 걸 모른다 — 그
   * 칸의 bounding box가 스크롤 컨테이너 사각형과 조금이라도 겹치면 "이미 보인다"고 판단해 스크롤을
   * 안 해 버려서, 위로 점프한 칸이 고정 헤더 뒤에 숨어버리는 경우가 있었다(아래로 점프할 땐 그
   * 자리를 가리는 고정 요소가 없어서 문제가 안 드러났다). 고정 헤더/행번호 열의 실제 두께만큼을
   * "보이는 영역"에서 미리 빼고 직접 스크롤한다. */
  function scrollCellIntoView(td: HTMLElement): void {
    const hostRect = scrollHostEl.getBoundingClientRect();
    const cellRect = td.getBoundingClientRect();
    const theadHeight = tableEl.querySelector('thead')?.getBoundingClientRect().height ?? 0;
    const rownumWidth = tableEl.querySelector<HTMLElement>('.g-edit-rownum')?.getBoundingClientRect().width ?? 0;
    const visibleTop = hostRect.top + theadHeight;
    const visibleLeft = hostRect.left + rownumWidth;
    if (cellRect.top < visibleTop) scrollHostEl.scrollTop -= visibleTop - cellRect.top;
    else if (cellRect.bottom > hostRect.bottom) scrollHostEl.scrollTop += cellRect.bottom - hostRect.bottom;
    if (cellRect.left < visibleLeft) scrollHostEl.scrollLeft -= visibleLeft - cellRect.left;
    else if (cellRect.right > hostRect.right) scrollHostEl.scrollLeft += cellRect.right - hostRect.right;
  }

  /** 지금 칸에서 (dRow,dCol)만큼 떨어진 칸으로 선택을 옮긴다. 병합에 덮여 실제 <td>가 없는
   * 칸이면(예: 1행 제목 병합) 같은 방향으로 한 칸씩 더 건너뛴다. 시트 렌더 범위 끝에 닿으면 멈춘다. */
  function moveSelection(fromRow: number, fromCol: number, dRow: number, dCol: number): void {
    const sheet = currentStructure();
    const { rows: rowCount, cols: colCount } = renderCountsFor(sheet);
    let row = fromRow;
    let col = fromCol;
    for (let attempt = 0; attempt < 30; attempt += 1) {
      let nextRow = Math.min(Math.max(row + dRow, 0), rowCount - 1);
      // 필터로 숨긴 행은 건너뛴다(Excel과 같다)
      while (dRow !== 0 && hiddenRowSet.has(nextRow) && nextRow + dRow >= 0 && nextRow + dRow < rowCount) nextRow += dRow;
      if (hiddenRowSet.has(nextRow)) nextRow = row;
      const nextCol = Math.min(Math.max(col + dCol, 0), colCount - 1);
      if (nextRow === row && nextCol === col) return; // 끝에 닿아 더 못 감
      row = nextRow;
      col = nextCol;
      const targetTd = tdForEnsureRendered(row, col);
      if (targetTd) {
        targetTd.focus();
        scrollCellIntoView(targetTd);
        return;
      }
    }
  }

  /** 선택 범위의 끝(selFocus)만 옮긴다 — 기준 칸(selAnchor, 지금 칸)과 포커스는 그대로. 끝 칸이 화면 밖이면 그쪽으로 스크롤. */
  function extendSelectionTo(row: number, col: number): void {
    const { rows: rowCount, cols: colCount } = renderCountsFor(currentStructure());
    selFocusRow = Math.min(Math.max(row, 0), rowCount - 1);
    selFocusCol = Math.min(Math.max(col, 0), colCount - 1);
    clearHeaderSelection();
    renderRangeHighlight();
    const td = tableEl.querySelector<HTMLElement>(`td[data-row="${selFocusRow}"][data-col="${selFocusCol}"]`);
    if (td) scrollCellIntoView(td);
  }

  /** Shift+방향키(한 칸씩), Ctrl+Shift+방향키(데이터 끝까지) — 범위 끝에서부터 넓히거나 줄인다 */
  function extendSelectionBy(dRow: number, dCol: number, toDataEdge: boolean): void {
    if (!toDataEdge) {
      extendSelectionTo(selFocusRow + dRow, selFocusCol + dCol);
      return;
    }
    const [r, c] = findDataEdge(selFocusRow, selFocusCol, dRow, dCol);
    extendSelectionTo(r, dCol === 0 ? selFocusCol : c);
  }

  /** 엑셀의 Ctrl+방향키(End 모드) — 빈 칸이거나 바로 다음 칸이 비어 있으면 다음 데이터(또는
   * 시트 끝)까지 건너뛰고, 데이터가 연속돼 있으면 그 블록이 끝나는 칸까지 단번에 이동한다. */
  /**
   * 엑셀의 Ctrl+방향키가 멈출 칸 — 빈 칸이거나 바로 다음 칸이 비어 있으면 다음 데이터(또는 시트 끝)까지, 데이터가
   * 이어지면 그 블록 끝까지. 큰 시트는 화면 근처 페이지만 받아 두므로, 아직 받지 않은 행에 닿으면 알 수 없어 데이터의
   * 처음·끝 행으로 간다(예전엔 받은 페이지 끝인 500행에서 멈췄다 — 100만 행 파일에서 Ctrl+↓가 A500).
   */
  function findDataEdge(fromRow: number, fromCol: number, dRow: number, dCol: number): [number, number] {
    const sheet = currentStructure();
    const { rows: rowCount, cols: colCount } = renderCountsFor(sheet);
    const lastDataRow = Math.max(0, sheet.rowCount - 1);
    const inBounds = (r: number, c: number) => r >= 0 && c >= 0 && r < rowCount && c < colCount;
    const unknown = (r: number) => dRow !== 0 && r >= 0 && r <= lastDataRow && !rowArrived(r);
    const isEmpty = (r: number, c: number): boolean => {
      if (!inBounds(r, c)) return true;
      const v = cells.get(gridCellKey(activeSheet, r, c));
      return v === undefined || v === '';
    };
    const jumpPastUnknown = (): [number, number] => [dRow > 0 ? lastDataRow : 0, fromCol];
    let r = fromRow;
    let c = fromCol;
    if (isEmpty(r, c) || (!unknown(r + dRow) && isEmpty(r + dRow, c + dCol))) {
      r += dRow;
      c += dCol;
      while (inBounds(r, c) && isEmpty(r, c)) {
        if (unknown(r)) return jumpPastUnknown();
        r += dRow;
        c += dCol;
      }
      return [Math.min(Math.max(r, 0), rowCount - 1), Math.min(Math.max(c, 0), colCount - 1)];
    }
    while (inBounds(r + dRow, c + dCol) && (unknown(r + dRow) || !isEmpty(r + dRow, c + dCol))) {
      if (unknown(r + dRow)) return jumpPastUnknown();
      r += dRow;
      c += dCol;
    }
    return [r, c];
  }

  function moveSelectionToDataEdge(fromRow: number, fromCol: number, dRow: number, dCol: number): void {
    const { rows: rowCount, cols: colCount } = renderCountsFor(currentStructure());
    let [row, col] = findDataEdge(fromRow, fromCol, dRow, dCol);
    // 병합에 덮인 칸이라 실제 <td>가 없으면(moveSelection과 동일한 보정) 같은 방향으로 한 칸 더 간다.
    for (let attempt = 0; attempt < 30; attempt += 1) {
      const targetTd = tdForEnsureRendered(row, col);
      if (targetTd) {
        targetTd.focus();
        scrollCellIntoView(targetTd);
        return;
      }
      const nextRow = Math.min(Math.max(row + dRow, 0), rowCount - 1);
      const nextCol = Math.min(Math.max(col + dCol, 0), colCount - 1);
      if (nextRow === row && nextCol === col) return;
      row = nextRow;
      col = nextCol;
    }
  }

  function handleCellKeydown(
    event: KeyboardEvent,
    td: HTMLTableCellElement,
    row: number,
    col: number,
    key: string,
    canEditThisCell: boolean,
  ): void {
    if (editingKey === key) {
      if (event.key === 'Enter') {
        event.preventDefault();
        commitAndMove(td, key, row, col, 1, 0);
      } else if (event.key === 'Tab') {
        event.preventDefault();
        commitAndMove(td, key, row, col, 0, event.shiftKey ? -1 : 1);
      } else if (event.key === 'Escape') {
        event.preventDefault();
        cancelEdit(td, key);
      }
      return; // 다른 키(화살표 등)는 텍스트 커서 이동 등 브라우저 기본 동작 그대로 둔다
    }

    switch (event.key) {
      case 'F2':
        if (canEditThisCell) {
          event.preventDefault();
          enterEditMode(td, key);
        }
        break;
      case 'Enter':
        event.preventDefault();
        moveSelection(row, col, 1, 0);
        break;
      case 'Tab':
        event.preventDefault();
        moveSelection(row, col, 0, event.shiftKey ? -1 : 1);
        break;
      case 'ArrowUp':
        event.preventDefault();
        if (event.shiftKey) extendSelectionBy(-1, 0, event.ctrlKey || event.metaKey);
        else if (event.ctrlKey || event.metaKey) moveSelectionToDataEdge(row, col, -1, 0);
        else moveSelection(row, col, -1, 0);
        break;
      case 'ArrowDown':
        event.preventDefault();
        if (event.shiftKey) extendSelectionBy(1, 0, event.ctrlKey || event.metaKey);
        else if (event.ctrlKey || event.metaKey) moveSelectionToDataEdge(row, col, 1, 0);
        else moveSelection(row, col, 1, 0);
        break;
      case 'ArrowLeft':
        event.preventDefault();
        if (event.shiftKey) extendSelectionBy(0, -1, event.ctrlKey || event.metaKey);
        else if (event.ctrlKey || event.metaKey) moveSelectionToDataEdge(row, col, 0, -1);
        else moveSelection(row, col, 0, -1);
        break;
      case 'ArrowRight':
        event.preventDefault();
        if (event.shiftKey) extendSelectionBy(0, 1, event.ctrlKey || event.metaKey);
        else if (event.ctrlKey || event.metaKey) moveSelectionToDataEdge(row, col, 0, 1);
        else moveSelection(row, col, 0, 1);
        break;
      case 'Delete':
      case 'Backspace': {
        if (!canEdit) break;
        event.preventDefault();
        clearSelectionValues();
        break;
      }
      default:
        // Ctrl/Cmd+C — 지금 선택된 범위를 TSV로 클립보드에 복사한다.
        if ((event.ctrlKey || event.metaKey) && (event.key === 'c' || event.key === 'C')) {
          event.preventDefault();
          copySelectionToClipboard();
          break;
        }
        // 선택된 채로 문자를 바로 치면(엑셀처럼) 기존 값을 지우고 그 문자로 편집모드에 들어간다.
        if (canEditThisCell && event.key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
          event.preventDefault();
          enterEditMode(td, key, event.key);
        }
    }
  }

  // ---------- 자동 채우기 핸들 (드래그로 값 복사, 엑셀과 동일) ----------
  const FILL_HANDLE_SIZE = 7;
  const fillHandleEl = document.createElement('div');
  fillHandleEl.className = 'g-fill-handle';
  fillHandleEl.hidden = true;
  scrollHostEl.appendChild(fillHandleEl);

  /** 지금 선택된 칸(td)의 오른쪽 아래 모서리에 채우기 핸들을 옮겨 놓는다. 편집 불가한 칸이 선택되면
   * 숨긴다. scrollHostEl의 스크롤 콘텐츠 안에 핸들이 같이 들어있으므로 스크롤해도 같이 움직인다. */
  function positionFillHandle(td: HTMLTableCellElement | null): void {
    if (!td) {
      fillHandleEl.hidden = true;
      return;
    }
    const tdRect = td.getBoundingClientRect();
    const hostRect = scrollHostEl.getBoundingClientRect();
    fillHandleEl.style.left = `${tdRect.right - hostRect.left + scrollHostEl.scrollLeft - FILL_HANDLE_SIZE / 2}px`;
    fillHandleEl.style.top = `${tdRect.bottom - hostRect.top + scrollHostEl.scrollTop - FILL_HANDLE_SIZE / 2}px`;
    fillHandleEl.hidden = false;
  }

  let fillSource: { row: number; col: number; key: string } | null = null;
  let fillTargetRow = 0;
  let fillTargetCol = 0;

  /** 드래그로 지나간 칸들(시작 칸 제외)에 한 번씩 콜백을 호출한다 — 세로/가로 중 더 크게 움직인
   * 한 방향으로만 채워진다(엑셀 자동 채우기와 동일하게 대각선으로는 안 채워짐). */
  function forEachFillCell(fn: (row: number, col: number) => void): void {
    if (!fillSource) return;
    const r0 = Math.min(fillSource.row, fillTargetRow);
    const r1 = Math.max(fillSource.row, fillTargetRow);
    const c0 = Math.min(fillSource.col, fillTargetCol);
    const c1 = Math.max(fillSource.col, fillTargetCol);
    for (let r = r0; r <= r1; r += 1) {
      for (let c = c0; c <= c1; c += 1) {
        if (r === fillSource.row && c === fillSource.col) continue;
        fn(r, c);
      }
    }
  }

  function clearFillPreview(): void {
    tableEl.querySelectorAll('.g-fill-preview').forEach((el) => el.classList.remove('g-fill-preview'));
  }

  function onFillDragMove(event: MouseEvent): void {
    if (!fillSource) return;
    const el = document.elementFromPoint(event.clientX, event.clientY) as HTMLElement | null;
    const td = el?.closest('td[data-row][data-col]') as HTMLTableCellElement | null;
    if (!td) return;
    const row = Number(td.dataset.row);
    const col = Number(td.dataset.col);
    const dRow = row - fillSource.row;
    const dCol = col - fillSource.col;
    if (Math.abs(dRow) >= Math.abs(dCol)) {
      fillTargetRow = row;
      fillTargetCol = fillSource.col;
    } else {
      fillTargetRow = fillSource.row;
      fillTargetCol = col;
    }
    clearFillPreview();
    forEachFillCell((r, c) => tdFor(r, c)?.classList.add('g-fill-preview'));
  }

  function onFillDragEnd(): void {
    document.removeEventListener('mousemove', onFillDragMove);
    document.removeEventListener('mouseup', onFillDragEnd);
    clearFillPreview();
    // forEachFillCell이 module 범위의 fillSource를 그대로 읽으므로, 실제 작업이 끝나기 전까지는
    // null로 되돌리면 안 된다(먼저 비우면 forEachFillCell이 바로 빈 채로 리턴해버려 아무 칸도
    // 못 채운다) — 다 쓰고 나서 맨 마지막에만 정리한다.
    if (!fillSource || !canEdit) {
      fillSource = null;
      return;
    }
    const sourceRow = fillSource.row;
    const sourceCol = fillSource.col;
    const sourceValue = cells.get(fillSource.key) ?? '';
    const touched: Array<{ r: number; c: number }> = [];
    ydoc.transact(() => {
      forEachFillCell((r, c) => {
        const key = gridCellKey(activeSheet, r, c);
        const targetStyle = styles.get(key);
        if (targetStyle?.lk === 1) return; // 잠긴 칸은 건너뜀
        if (targetStyle?.nu === 1 && sourceValue !== '' && !NUMERIC_VALUE.test(sourceValue)) return; // 숫자 전용 칸에 숫자가 아니면 건너뜀
        if (sourceValue === '') cells.delete(key);
        else cells.set(key, sourceValue);
        touched.push({ r, c });
      });
      const s = currentStructure();
      const maxRow = Math.max(sourceRow, fillTargetRow);
      const maxCol = Math.max(sourceCol, fillTargetCol);
      if (maxRow >= s.rowCount || maxCol >= s.colCount) {
        structure.set(String(activeSheet), {
          ...s,
          rowCount: Math.max(s.rowCount, maxRow + 1),
          colCount: Math.max(s.colCount, maxCol + 1),
        });
      }
    });
    touched.forEach(({ r, c }) => {
      const td = tdFor(r, c);
      if (td) renderCellDisplay(td, gridCellKey(activeSheet, r, c));
    });
    recalcFormulaCells();
    fillSource = null;
  }

  fillHandleEl.addEventListener('mousedown', (event) => {
    event.preventDefault();
    event.stopPropagation();
    if (!activeCellKey) return;
    const parsed = parseGridCellKey(activeCellKey);
    if (!parsed) return;
    fillSource = { row: parsed.row, col: parsed.col, key: activeCellKey };
    fillTargetRow = parsed.row;
    fillTargetCol = parsed.col;
    document.addEventListener('mousemove', onFillDragMove);
    document.addEventListener('mouseup', onFillDragEnd);
  });

  /** 한 행(<tr>) 안에서의 겹침 순서 — buildRow()가 각 <tr>을 position+z-index로 자체 스태킹
   * 컨텍스트로 격리해 두므로(행 간 우열은 그 <tr>의 z-index가 정한다), 이 값들은 전부 "같은 행
   * 안에서만" 비교된다. 가로 스크롤을 하면 얼린 칸이 아닌 칸들이 실제로 행번호 열 **위를 지나가므로**,
   * 왼쪽에 고정된 것일수록 높은 값을 줘야 한다(같은 값이면 DOM 순서가 이기는데, 행번호 <th>는 늘
   * 그 행의 첫 자식이라 항상 진다 — 그게 "가로로 스크롤하면 얼린 행의 행번호가 사라지는" 버그였다).
   *   행번호 열(left:0)  > 얼린 열(left:28px~) > 얼린 행의 나머지 칸 > 보통 칸
   * 행번호 열 값은 CSS(.grid-edit-table .g-edit-rownum)의 z-index와 같아야 한다 — 틀 고정이 없는
   * 시트에서는 CSS 값만으로 같은 순서가 나와야 하기 때문. */
  const Z_ROWNUM = 5;
  /** 얼린 열 — tbody에서는 얼린 행의 나머지 칸 위에, <thead>에서는 보통 열 문자 머리글
   * (CSS z-index:2) 위에 와야 한다. 두 경우 다 "행번호 열 바로 아래 층"이라 값이 같다. */
  const Z_FROZEN_COL = 4;
  const Z_FROZEN_ROW = 3;

  /**
   * 위 freezeRows행 / 왼쪽 freezeCols열을 실제로 sticky 처리한다. 렌더링 직후 DOM에 실제로 붙은
   * 뒤에 호출해야 offsetHeight/offsetWidth로 진짜 셀 크기를 잴 수 있다 — 그래야 컬럼 letter 헤더
   * (이미 sticky top:0)와 행번호 열(이미 sticky left:0) 바로 다음에 정확히 이어붙일 수 있다.
   * 배경을 불투명하게 칠해야 스크롤된 내용이 그 밑으로 비쳐 보이지 않는다.
   */
  function applyFreezePanes(freezeRows: number, freezeCols: number): void {
    if (freezeRows === 0 && freezeCols === 0) return;

    const theadHeight = tableEl.querySelector('thead')?.getBoundingClientRect().height ?? 0;
    const rownumWidth = tableEl.querySelector<HTMLElement>('.g-edit-rownum')?.getBoundingClientRect().width ?? 0;

    if (freezeRows > 0) {
      // 행 높이는 전부 동일(CSS 고정값)하므로, 얼릴 행들의 실제 렌더 높이를 각 행 대표 셀에서 재
      // 누적해 top을 계산한다.
      let top = theadHeight;
      for (let r = 0; r < freezeRows; r += 1) {
        const rowCells = tableEl.querySelectorAll<HTMLElement>(`tbody td[data-row="${r}"], tbody th[data-row="${r}"]`);
        let rowHeight = 0;
        rowCells.forEach((cell) => {
          cell.style.position = 'sticky';
          cell.style.top = `${top}px`;
          // 행번호 <th>는 이 행에서도 계속 맨 위여야 한다 — 얼린 행의 나머지 칸은 sticky left가
          // 없어서 가로 스크롤 때 이 칸 위로 밀려 들어오는데, .g-frozen-pane이 불투명 배경이라
          // 같은 z-index면 행번호를 통째로 지워 버린다(DOM 순서상 <th>가 항상 먼저다).
          if (cell.classList.contains('g-edit-rownum')) cell.style.zIndex = String(Z_ROWNUM);
          else cell.style.zIndex = String(cell.classList.contains('g-frozen-col') ? Z_FROZEN_COL : Z_FROZEN_ROW);
          cell.classList.add('g-frozen-pane');
          rowHeight = Math.max(rowHeight, cell.getBoundingClientRect().height);
        });
        top += rowHeight;
      }
    }

    if (freezeCols > 0) {
      let left = rownumWidth;
      for (let c = 0; c < freezeCols; c += 1) {
        const colCells = tableEl.querySelectorAll<HTMLElement>(`tbody td[data-col="${c}"]`);
        let colWidth = 0;
        colCells.forEach((cell) => {
          cell.style.position = 'sticky';
          cell.style.left = `${left}px`;
          if (!cell.style.zIndex) cell.style.zIndex = String(Z_FROZEN_COL);
          cell.classList.add('g-frozen-pane');
          colWidth = Math.max(colWidth, cell.getBoundingClientRect().width);
        });
        // 얼린 "열"은 그 열의 문자 머리글(A, B…)도 같이 고정해야 한다 — 안 그러면 데이터 칸만
        // 왼쪽에 남고 머리글은 그대로 흘러가, 고정된 A·B열 위에 엉뚱한 글자(J, K…)가 붙는다.
        // .g-frozen-pane 클래스는 일부러 안 붙인다(그 배경 규칙이 머리글의 회색 배경을 덮어
        // 얼린 열 머리글만 색이 달라진다 — <thead> th는 이미 불투명 배경을 갖고 있다).
        const headCell = tableEl.querySelector<HTMLElement>(`thead th[data-col="${c}"]`);
        if (headCell) {
          headCell.style.position = 'sticky';
          headCell.style.left = `${left}px`;
          headCell.style.zIndex = String(Z_FROZEN_COL);
          // 그 열의 tbody 칸이 하나도 없는 극단적인 경우(빈 스크롤 창)에도 폭 누적이 0으로
          // 무너지지 않게, 머리글 폭을 대신 쓴다 — 두 폭은 <colgroup>이 똑같이 정한다.
          if (colWidth === 0) colWidth = headCell.getBoundingClientRect().width;
        }
        left += colWidth;
      }
    }
  }

  function clearActiveHeaders(): void {
    tableEl.querySelectorAll('.g-col-active').forEach((el) => el.classList.remove('g-col-active'));
    tableEl.querySelectorAll('.g-row-active').forEach((el) => el.classList.remove('g-row-active'));
  }
  function setActiveHeaders(row: number, col: number): void {
    clearActiveHeaders();
    clearHeaderSelection();
    tableEl.querySelector(`thead th[data-col="${col}"]`)?.classList.add('g-col-active');
    tableEl.querySelector(`tbody th[data-row="${row}"]`)?.classList.add('g-row-active');
  }

  // 열/행 머리글(A, 1 등)을 직접 클릭했을 때 — 엑셀처럼 그 열/행 전체를 선택 표시한다.
  /** 지금 머리글 클릭으로 통째로 선택돼 있는 열/행 — 클래스(.g-col-selected/.g-row-selected)는
   * 해당 <tr>이 가상 스크롤 창 밖으로 나가면 그 행과 함께 사라져 버려서, 화면만으로는 "원래
   * 무엇이 선택돼 있었는지"를 복구할 수 없다. 스크롤로 행이 다시 들어왔을 때 표시를 되살리려면
   * 클래스와 별개로 이 상태를 들고 있어야 한다. */
  /** 고른 열/행 머리글 — Shift+클릭으로 index~to 범위까지 넓힌다(to가 없으면 하나) */
  let headerSelection: { kind: 'col' | 'row'; index: number; to?: number } | null = null;
  function clearHeaderSelection(): void {
    headerSelection = null;
    tableEl.querySelectorAll('.g-col-selected, .g-row-selected').forEach((el) => {
      el.classList.remove('g-col-selected', 'g-row-selected');
    });
  }
  /** headerSelection이 가리키는 열/행에 선택 표시 클래스를 입힌다 — 머리글 클릭 직후와, 스크롤로
   * 그 행이 다시 그려진 뒤 양쪽에서 같은 코드를 쓴다. */
  function paintHeaderSelection(): void {
    if (!headerSelection) return;
    const { kind, index, to = index } = headerSelection;
    const attr = kind === 'col' ? 'data-col' : 'data-row';
    const cls = kind === 'col' ? 'g-col-selected' : 'g-row-selected';
    const lo = Math.min(index, to);
    const hi = Math.max(index, to);
    if (kind === 'row') {
      // 행 범위는 수십만 행일 수 있다 — 지금 그려진 행만 칠한다
      for (const [r, tr] of vwRowEls) {
        if (r >= lo && r <= hi) tr.querySelectorAll(`[${attr}="${r}"]`).forEach((el) => el.classList.add(cls));
      }
      return;
    }
    for (let i = lo; i <= hi; i += 1) {
      tableEl.querySelectorAll(`[${attr}="${i}"]`).forEach((el) => el.classList.add(cls));
    }
  }
  // 열/행 머리글을 클릭하면 시각적 표시만 하는 게 아니라, 서식 툴바가 실제로 적용할 범위
  // (selAnchor~selFocus)도 그 열/행 전체로 잡아야 한다 — 안 그러면 rangeBounds()가 여전히
  // 마지막으로 포커스했던 칸 하나만 가리켜, 열/행을 선택한 채 서식을 바꿔도 그 칸에만 적용되고
  // 나머지는 그대로인 것처럼 보였다(실제로는 "적용이 안 되는" 게 아니라 "한 칸에만 적용되는" 버그).
  function selectColumn(col: number, extend = false): void {
    const anchor = extend && headerSelection?.kind === 'col' ? headerSelection.index : col;
    clearActiveHeaders();
    clearHeaderSelection();
    const { rowCount } = currentStructure();
    selAnchorRow = 0;
    selFocusRow = Math.max(0, rowCount - 1);
    selAnchorCol = anchor;
    selFocusCol = col;
    activeCellKey = gridCellKey(activeSheet, selAnchorRow, anchor);
    cellRefBoxEl.textContent = cellRefLabel(selAnchorRow, anchor);
    renderRangeHighlight();
    headerSelection = { kind: 'col', index: anchor, to: col };
    paintHeaderSelection();
  }
  function selectRow(row: number, extend = false): void {
    const anchor = extend && headerSelection?.kind === 'row' ? headerSelection.index : row;
    clearActiveHeaders();
    clearHeaderSelection();
    const { colCount } = currentStructure();
    selAnchorRow = anchor;
    selFocusRow = row;
    selAnchorCol = 0;
    selFocusCol = Math.max(0, colCount - 1);
    activeCellKey = gridCellKey(activeSheet, anchor, selAnchorCol);
    cellRefBoxEl.textContent = cellRefLabel(anchor, selAnchorCol);
    renderRangeHighlight();
    headerSelection = { kind: 'row', index: anchor, to: row };
    paintHeaderSelection();
  }
  // 잠긴 칸은 focus 리스너 자체가 안 붙으므로(편집 불가) setActiveHeaders를 못 거친다 — 어떤 칸을
  // 클릭하든(잠금 여부와 무관하게) 열/행 전체 선택 표시는 항상 풀리도록 여기서 한 번 더 잡는다.
  tableEl.addEventListener('click', (event) => {
    if ((event.target as HTMLElement).tagName === 'TD') clearHeaderSelection();
  });

  // ---------- 다중 셀 드래그 선택 (엑셀처럼 클릭+드래그로 범위 선택) ----------
  function rangeBounds(): { r0: number; r1: number; c0: number; c1: number } {
    return {
      r0: Math.min(selAnchorRow, selFocusRow),
      r1: Math.max(selAnchorRow, selFocusRow),
      c0: Math.min(selAnchorCol, selFocusCol),
      c1: Math.max(selAnchorCol, selFocusCol),
    };
  }
  function clearRangeHighlight(): void {
    tableEl.querySelectorAll('.g-range-selected').forEach((el) => el.classList.remove('g-range-selected'));
  }
  /**
   * 시트를 바꿀 때 — 선택 상태(드래그 범위/머리글 선택/활성 칸)를 통째로 비운다.
   *
   * selAnchor~selFocus도 headerSelection도 "몇 번 시트의" 선택인지를 들고 있지 않다(activeCellKey만
   * 시트 번호를 갖는다). 그래서 시트를 바꿔도 이 값들이 그대로 남아 있으면 그대로 새 시트에 적용된다:
   *  (1) 스크롤 한 번이면 repaintScrollWindowSelection()이 이전 시트의 범위를 새 시트에 칠한다
   *      — 실측: Sheet1에서 C열 전체를 선택하고 Sheet2로 간 뒤 휠을 한 번 굴리니, 선택한 적도
   *      없는 Sheet2에 125칸이 선택된 것처럼 칠해졌다.
   *  (2) 더 나쁘게는, 툴바 서식 버튼이 전부 rangeBounds() + gridCellKey(activeSheet, r, c)로 키를
   *      만들기 때문에 — 칠하기 버그가 없더라도 이전 시트에서 잡아 둔 범위가 새 시트의 같은 좌표
   *      칸들에 조용히 먹는다(같은 실측에서 굵게가 Sheet2의 125칸에 그대로 먹었다).
   * 엑셀에서도 시트를 바꾸면 보고 있지 않은 시트의 선택이 따라오지 않으므로, 새 시트는 아무것도
   * 선택되지 않은 상태로 시작한다 — 사용자가 칸을 클릭하는 순간 focusin이 전부 다시 채운다.
   */
  function resetSelectionForSheetSwitch(): void {
    clearRangeHighlight();
    clearActiveHeaders();
    clearHeaderSelection();
    activeCellKey = null;
    selAnchorRow = 0;
    selAnchorCol = 0;
    selFocusRow = 0;
    selFocusCol = 0;
    // 활성 칸이 없어졌으니 그 칸 오른쪽 아래에 떠 있던 채우기 핸들도 같이 감춘다 — 안 그러면 새
    // 시트 위에 그대로 남고, 그걸 끌면 (activeCellKey가 null이라) 아무 일도 안 하는 죽은 손잡이가 된다.
    positionFillHandle(null);
    // 이전 시트에서 고른 칸의 주소·값이 수식 입력줄에 남아 새 시트의 다른 칸 값처럼 보였다
    cellRefBoxEl.textContent = '';
    formulaInputEl.value = '';
  }
  /** 앵커~포커스 사각형에 배경을 칠한다 — 1칸짜리(드래그 안 한 보통 클릭)면 기존 :focus 테두리로
   * 충분하니 별도 배경은 안 칠한다. */
  function renderRangeHighlight(): void {
    clearRangeHighlight();
    const { r0, r1, c0, c1 } = rangeBounds();
    if (r0 === r1 && c0 === c1) return;
    // 지금 그려진 행(vwRowEls)만 칠한다 — 범위의 칸마다 DOM을 찾으면 100만 행 선택(Ctrl+Shift+↓)에 30초가 걸렸다.
    // 창 밖 행은 스크롤로 들어올 때 repaintScrollWindowSelection이 칠한다.
    for (const [r, tr] of vwRowEls) {
      if (r < r0 || r > r1) continue;
      for (let i = 0; i < tr.children.length; i += 1) {
        const cell = tr.children[i] as HTMLElement;
        const col = cell.dataset.col;
        if (col === undefined) continue;
        const c = Number(col);
        if (c >= c0 && c <= c1) cell.classList.add('g-range-selected');
      }
    }
  }
  /**
   * updateScrollWindow() 전용 — 스크롤로 새로 들어온 행에 선택 관련 "칠하기"를 다시 입힌다.
   *
   * 왜 필요한가: buildRow()는 늘 아무 장식도 없는 <td>를 만든다. 전체 재빌드 경로에서는 마지막에
   * 활성 칸을 다시 focus시켜 focusin 핸들러가 renderRangeHighlight/setActiveHeaders를 도로
   * 불러줬지만, 스크롤 전용 증분 패치에는 그 마무리가 통째로 빠져 있었다. 그 결과 선택 범위를
   * 잡아 둔 채 스크롤로 한 바퀴 돌고 오면 **화면에는 아무것도 선택돼 있지 않은데 내부 선택 상태
   * (selAnchor~selFocus)는 그대로 살아 있는** 상태가 됐다 — 이때 굵게/서식 지우기/채우기색 같은
   * 툴바 버튼(전부 rangeBounds()를 읽는다)을 누르면 "보이지 않는" 범위 수십~수백 칸에 서식이
   * 조용히 적용된다(실측: 화면상 0칸 선택 상태에서 굵게가 44칸에 먹었다).
   *
   * 성능: renderRangeHighlight()는 범위의 칸마다 tableEl.querySelector를 한 번씩 돌아서,
   * selectColumn()이 잡는 "열 전체"(selFocusRow = rowCount-1, 수천 행)를 선택해 둔 채 스크롤하면
   * 스크롤 프레임마다 수천 번의 DOM 질의가 된다 — 이 함수 전체가 매 스크롤 프레임 도는 자리라
   * 그건 81be4d8이 없앤 바로 그 종류의 비용을 되살리는 셈이다. 어차피 창 밖 행에는 칠할 <td>가
   * 존재하지 않으므로, 지금 DOM에 있는 행(vwRowEls)과 범위의 교집합만 훑는다.
   */
  function repaintScrollWindowSelection(): void {
    const { r0, r1, c0, c1 } = rangeBounds();
    // 한 칸짜리 선택이면 칠할 것이 없다(renderRangeHighlight와 같은 규칙 — :focus 테두리로 충분).
    // 이때는 지울 것도 없다: 선택이 한 칸으로 줄어드는 모든 경로가 renderRangeHighlight를 거쳐
    // 이미 전부 지웠고, 새로 만들어진 행에는 애초에 클래스가 없다. 그래서 선택 범위를 안 쓰는
    // 보통의 스크롤에서는 이 함수가 DOM을 단 한 번도 건드리지 않는다.
    if (r0 !== r1 || c0 !== c1) {
      clearRangeHighlight();
      for (const [r, tr] of vwRowEls) {
        if (r < r0 || r > r1) continue;
        // 행 하나당 자식 훑기 한 번 — 칸마다 querySelector를 도는 것보다 훨씬 싸다.
        for (let i = 0; i < tr.children.length; i += 1) {
          const cell = tr.children[i] as HTMLElement;
          const col = cell.dataset.col;
          if (col === undefined) continue; // 행번호 <th>
          const c = Number(col);
          if (c >= c0 && c <= c1) cell.classList.add('g-range-selected');
        }
      }
    }

    // 머리글 클릭으로 열/행을 통째로 선택해 둔 상태(g-col-selected/g-row-selected)도 새로 들어온
    // 행에 다시 입힌다 — 선택을 안 쓰는 보통의 스크롤에서는 headerSelection이 null이라 비용이 없다.
    paintHeaderSelection();

    // 활성 칸이 있는 행의 행번호 머리글 표시(g-row-active)도 그 <tr>과 함께 사라졌다가 돌아온다.
    if (activeCellKey) {
      const parsed = parseGridCellKey(activeCellKey);
      if (parsed && parsed.sheetIdx === activeSheet) {
        tableEl.querySelector(`tbody th[data-row="${parsed.row}"]`)?.classList.add('g-row-active');
      }
    }
  }

  /**
   * updateScrollWindow() 전용 — 창 밖으로 나갔다 돌아온 활성 칸에 실제 DOM 포커스를 되돌린다.
   * 전체 재빌드 경로(renderTableVirtualized)는 맨 끝에서 이미 같은 일을 하지만 스크롤 전용 증분
   * 패치에는 그게 없었다: 활성 칸의 <tr>이 창 밖으로 제거되는 순간 포커스가 <body>로 떨어지는데,
   * 키보드 핸들러는 tableEl에 위임돼 있어 <body>에 포커스가 있는 상태로 친 키는 그리드에 **아예
   * 도달하지 못한다** — 다시 스크롤해 그 칸이 돌아와도 포커스는 <body>에 남아, 사용자가 친 글자가
   * 아무런 피드백 없이 통째로 사라졌다.
   *
   * 포커스가 정말로 "아무 데도 없을 때"(<body>)만 되돌린다 — 수식 입력줄이나 툴바 버튼, 열린
   * 대화상자에 사용자가 옮겨 둔 포커스를 스크롤이 가로채면 안 된다.
   */
  function restoreScrollWindowFocus(firstVisible: number, lastVisible: number, freezeRows: number): void {
    if (!activeCellKey) return;
    // activeElement만으로는 "이 문서 안에서 아무 데도 포커스가 없다"와 "포커스가 아예 이 문서
    // 밖(호스트 페이지·다른 탭·다른 창)에 있다"를 구분할 수 없다 — 둘 다 activeElement가 <body>다.
    // 이 뷰어는 호스트 페이지에 iframe으로 박혀 돌아가는 것을 정식으로 지원하므로
    // (DOCVIEW_ALLOWED_ORIGINS, buildRow의 window.parent.postMessage 참고), 사용자가 iframe을
    // 클릭한 적조차 없이 호스트 쪽 입력칸에 타이핑하면서 그리드 위에 휠만 굴려도 이 함수가 돌 수
    // 있다. 그때 포커스를 되돌리면 그리드가 호스트의 포커스를 소리 없이 빼앗아, 이어지는 타자가
    // 호스트 입력칸이 아니라 스프레드시트 칸에 들어가 그대로 커밋된다 — "키가 사라지는" 원래
    // 버그보다 훨씬 나쁘다. document.hasFocus()는 정확히 이 경우에만 false고, 정작 고치려던
    // 경우(이 문서는 포커스를 갖고 있는데 활성 칸의 <tr>이 스크롤 패치로 잠깐 사라져 <body>로
    // 떨어진 상태)에는 true다.
    if (!document.hasFocus()) return;
    const active = document.activeElement;
    if (active !== null && active !== document.body) return;
    const parsed = parseGridCellKey(activeCellKey);
    if (!parsed || parsed.sheetIdx !== activeSheet) return;
    const { row, col } = parsed;
    // 틀 고정 행은 창과 무관하게 항상 그려져 있다. 그 밖의 행은 지금 창 안에 들어왔을 때만.
    if (row >= freezeRows && (row < firstVisible || row > lastVisible)) return;
    const td = tdFor(row, col);
    if (!td) return;
    // focus()는 위임된 focusin 핸들러를 그 자리에서(동기적으로) 실행하는데, 그 핸들러는 "칸을 새로
    // 클릭했다"는 뜻으로 선택 범위를 그 한 칸으로 접고(selAnchor/selFocus) 머리글 선택도 풀어버린다
    // (setActiveHeaders → clearHeaderSelection). 하지만 스크롤은 선택을 바꾸는 동작이 아니므로
    // (엑셀에서도 스크롤로 선택이 풀리지 않는다) 원래 선택 상태를 스냅숏해 뒀다가 그대로 되돌린다
    // — 실제 칠하기는 호출부가 이어서 repaintScrollWindowSelection()으로 한다.
    const anchorRow = selAnchorRow;
    const anchorCol = selAnchorCol;
    const focusRow = selFocusRow;
    const focusCol = selFocusCol;
    const header = headerSelection;
    td.focus({ preventScroll: true });
    selAnchorRow = anchorRow;
    selAnchorCol = anchorCol;
    selFocusRow = focusRow;
    selFocusCol = focusCol;
    headerSelection = header;
  }

  function onRangeDragMove(event: MouseEvent): void {
    const el = document.elementFromPoint(event.clientX, event.clientY) as HTMLElement | null;
    const td = el?.closest('td[data-row][data-col]') as HTMLTableCellElement | null;
    if (!td) return;
    const row = Number(td.dataset.row);
    const col = Number(td.dataset.col);
    if (row === selFocusRow && col === selFocusCol) return;
    selFocusRow = row;
    selFocusCol = col;
    renderRangeHighlight();
  }
  function onRangeDragEnd(): void {
    document.removeEventListener('mousemove', onRangeDragMove);
    document.removeEventListener('mouseup', onRangeDragEnd);
  }

  /** 선택 범위의 값을 비운다(Delete 키·홈 탭 "잘라내기"). 잠긴 칸은 건너뛴다. */
  function clearSelectionValues(): void {
    const { r0, r1, c0, c1 } = rangeBounds();
    ydoc.transact(() => {
      for (let r = r0; r <= r1; r += 1) {
        for (let c = c0; c <= c1; c += 1) {
          const k = gridCellKey(activeSheet, r, c);
          if (styles.get(k)?.lk === 1) continue; // 잠긴 칸은 건너뜀(선택 범위에 섞여 있어도 됨)
          if ((cells.get(k) ?? '') !== '') cells.delete(k);
        }
      }
    });
  }

  /** 지금 선택된 범위를 엑셀과 같은 TSV(탭=열, 줄바꿈=행)로 클립보드에 담는다 — handlePaste가
   * 읽는 형식과 대칭이라 그리드끼리, 또는 실제 엑셀과 오가는 복사/붙여넣기가 그대로 맞는다. */
  function copySelectionToClipboard(): void {
    const { r0, r1, c0, c1 } = rangeBounds();
    const lines: string[] = [];
    for (let r = r0; r <= r1; r += 1) {
      const rowValues: string[] = [];
      for (let c = c0; c <= c1; c += 1) {
        rowValues.push(cells.get(gridCellKey(activeSheet, r, c)) ?? '');
      }
      lines.push(rowValues.join('\t'));
    }
    navigator.clipboard?.writeText(lines.join('\n')).catch(() => {
      /* 클립보드 권한이 없는 환경이면 조용히 무시 — 복사 실패를 사용자에게 굳이 방해되게 알릴 정도는 아니다 */
    });
  }

  // 셀 편집 영역을 완전히 벗어날 때만(다른 셀로 이동하는 경우는 제외) awareness의 cell을 지운다 —
  // focusout은 버블링되므로 표 전체에 한 번만 걸어 둔다.
  // activeCellKey는 일부러 지우지 않는다 — 툴바 버튼(굵게 등)을 누르면 버튼이 포커스를 가져가며
  // 셀이 blur되는데, 이때 activeCellKey까지 지우면 "마지막으로 선택했던 셀에 서식 적용"이라는
  // 스프레드시트의 기본 동작이 깨진다(클릭한 서식 버튼이 조용히 아무 일도 안 하게 됨).
  // awareness의 cell(= "지금 편집 중" 표시)만 셀 영역을 완전히 벗어났을 때 지운다.
  tableEl.addEventListener('focusout', (event) => {
    const related = event.relatedTarget as HTMLElement | null;
    if (!related || related.getAttribute('contenteditable') !== 'true') {
      provider.awareness.setLocalStateField('cell', null);
    }
  });

  /** 그려진 행(vwRowEls) 안에서만 찾는다 — 큰 시트에서 페이지 하나(500행×열 수)가 도착하면 바뀐 칸 수만 번
   * 불리는데, 표 전체를 querySelector로 뒤지면 한 번 스크롤에 2~8초가 걸렸다(화면 밖 행은 바로 null). */
  function tdFor(row: number, col: number): HTMLTableCellElement | null {
    if (!vwValid) return tableEl.querySelector<HTMLTableCellElement>(`td[data-row="${row}"][data-col="${col}"]`);
    const tr = vwRowEls.get(row);
    if (!tr || !tr.isConnected) return null;
    return tr.querySelector<HTMLTableCellElement>(`td[data-col="${col}"]`);
  }

  /** 가상 스크롤 때문에 지금 화면 창(뷰포트+여유분) 밖의 행은 <td>가 DOM에 없을 수 있다 — 키보드로
   * 먼 칸으로 바로 이동하거나(Ctrl+방향키) 검색 결과로 점프할 때처럼, 화면 밖 칸으로 곧장 이동해야
   * 하면 그 행이 보이도록 스크롤 위치를 옮기고 즉시 다시 그린 뒤 다시 찾는다. 병합에 덮인 칸처럼
   * 애초에 <td>가 존재할 수 없는 경우엔 다시 그려도 여전히 null이며, 호출부의 기존 재시도 루프가
   * 그 경우를 처리한다. */
  function tdForEnsureRendered(row: number, col: number): HTMLTableCellElement | null {
    const existing = tdFor(row, col);
    if (existing) return existing;
    const viewportHeight = scrollHostEl.clientHeight;
    scrollHostEl.scrollTop = Math.max(0, rowLayout.top(row) - viewportHeight / 2);
    renderTableVirtualized();
    return tdFor(row, col);
  }

  interface RemoteCellState extends AwarenessUserState {
    readonly cell?: { readonly sheetIdx: number; readonly row: number; readonly col: number };
  }

  /** 다른 참여자가 지금 편집 중인 셀에 색 테두리 + 작은 이름표를 띄운다("누가 어디를 수정 중인지"). */
  /** 다른 참여자가 지금 편집 중인 셀에 그 사람 색으로 테두리만 두른다(이름표는 따로 안 보여준다 —
   * 색만으로 "누군가 지금 여기 있다"는 걸 충분히 알 수 있으면 됨). */
  function renderRemoteCellHighlights(): void {
    tableEl.querySelectorAll<HTMLTableCellElement>('.g-remote-cell').forEach((td) => {
      td.classList.remove('g-remote-cell');
      td.style.removeProperty('--remote-color');
    });
    const states = provider.awareness.getStates() as Map<number, RemoteCellState>;
    for (const [clientId, state] of states) {
      if (clientId === ydoc.clientID) continue;
      const cell = state.cell;
      const name = state.user?.name;
      if (!cell || !name || cell.sheetIdx !== activeSheet) continue;
      const td = tdFor(cell.row, cell.col);
      if (!td) continue;
      const color = state.user?.color ?? pickPresenceColorForName(name);
      td.classList.add('g-remote-cell');
      td.style.setProperty('--remote-color', color);
    }
  }

  cells.observe((event) => {
    for (const key of event.keysChanged) {
      const parsed = parseGridCellKey(key);
      if (!parsed || parsed.sheetIdx !== activeSheet) continue;
      const td = tdFor(parsed.row, parsed.col);
      // 선택만 된 칸(편집 중 아님)은 이제 원격 갱신을 그대로 반영해도 된다 — 실제로 로컬에서
      // 타이핑 중인 칸(editingKey === key)만 덮지 않는다. (선택=focus일 뿐 편집과는 분리됐다.)
      if (!td || editingKey === key) continue;
      const value = cells.get(key) ?? '';
      renderCellDisplay(td, key);
      applyStyleToTd(td, styles.get(key));
      if (key === activeCellKey && document.activeElement !== formulaInputEl) formulaInputEl.value = value;
      // 이 칸이 비었다/채워졌다가 바뀌면, 왼쪽에서 이 칸으로 넘쳐 들어오고 있었을 스필 폭도
      // 다시 계산해야 한다(엑셀처럼 — 옆 칸이 채워지면 넘침이 줄고, 비면 다시 늘어난다).
      refreshSpillNeighbor(key);
    }
    recalcFormulaCells(); // 바뀐 셀을 참조하는 다른 수식 셀들도 다시 계산
  });

  styles.observe((event) => {
    for (const key of event.keysChanged) {
      const parsed = parseGridCellKey(key);
      if (!parsed || parsed.sheetIdx !== activeSheet) continue;
      const td = tdFor(parsed.row, parsed.col);
      if (!td || editingKey === key) continue;
      applyStyleToTd(td, styles.get(key));
    }
  });

  /**
   * 서버가 계산해 둔 수식 값(GRID_COMPUTED_MAP)이 도착하면 그 칸만 다시 그린다 — 리뷰 지적 C3.
   *
   * 클라이언트는 더 이상 수식을 직접 계산하지 않고 computed 맵의 값을 읽어 보여줄 뿐인데
   * (computeCellDisplay), 이 맵에는 관찰자가 하나도 없었다. cells.observe가 부르는
   * recalcFormulaCells()는 **셀 값이 바뀐 그 순간** 동기적으로 도는 것이라, 그때는 서버가 아직
   * 계산하기 전이다(계산은 왕복 한 번 뒤에 computed 맵으로 돌아온다). 그래서 수식이 참조하는 칸을
   * 고쳐도 수식 칸의 표시값이 그대로 멈춰 있고, 스크롤·클릭 등 **다른 이유로 재렌더가 일어나야만**
   * 뒤늦게 갱신됐다 — 최종 리뷰가 잡아낸 바로 그 증상이다.
   *
   * 편집 중인 칸(editingKey === key)은 건드리지 않는다: 그 칸에는 계산 결과가 아니라 사용자가
   * 지금 치고 있는 수식 원문이 보여야 하고, 재렌더는 커서/선택 범위를 잃게 한다(cells/styles
   * 관찰자가 쓰는 것과 같은 가드).
   */
  computed.observe((event) => {
    for (const key of event.keysChanged) {
      const parsed = parseGridCellKey(key);
      if (!parsed || parsed.sheetIdx !== activeSheet) continue;
      const td = tdFor(parsed.row, parsed.col);
      if (td && editingKey !== key) renderCellDisplay(td, key);
    }
  });

  /**
   * 편집이 끝난 뒤, 편집 중이라 미뤄 뒀던 구조 변경 재렌더를 따라잡는다 — 리뷰 지적 I5의 뒷부분.
   * 편집 종료 경로(blur 커밋 / commitAndMove / cancelEdit) 끝에서 부른다.
   */
  function flushPendingStructureRender(): void {
    if (!pendingStructureRender || editingKey !== null) return;
    pendingStructureRender = false;
    renderSheetTabs();
    renderTableVirtualized();
  }

  structure.observe((event) => {
    /**
     * 원격에서 온 행/열 삽입·삭제라면 undo 스택을 통째로 버린다 — 리뷰 지적 I-4.
     *
     * UndoManager는 이 구조 변경 자체는 추적하지 않는다(서버가 보낸 원격 트랜잭션이라 trackedOrigins
     * 밖이다 — 위 sendGridStructureOp 주석의 "알려진 대가"). 문제는 **스택에 이미 쌓여 있는 로컬
     * 셀 편집들**이다: 그것들은 편집 당시의 `sheet:row:col` 키로 되돌리는데, 행/열이 밀린 뒤에는 그
     * 키가 전혀 다른 행/열을 가리킨다. 그 상태로 Ctrl+Z를 누르면 사용자가 손댄 적 없는 칸에 옛 값이
     * 조용히 써진다 — 이 프로젝트가 계속 쫓아온 그 손상 클래스와 같은 모양이다.
     *
     * 스택 항목을 새 위치로 다시 매핑하는 것이 이상적이지만 훨씬 복잡하다. 되돌릴 수 없게 되는
     * 것보다 잘못된 칸을 덮어쓰는 것이 훨씬 나쁘므로, 표준적인 완화책대로 스택을 버린다.
     *
     * 판정은 두 조건의 곱이다: (1) 로컬 트랜잭션이 아니다(`transaction.local === false` — Yjs가
     * 원격 업데이트를 적용할 때 쓰는 값), (2) 기존 시트의 행/열 수가 실제로 바뀌었다. 그래서 시트
     * 이름 변경·탭 순서 변경 같은 원격 갱신이나, 아무 칸도 밀지 않는 로컬 행 추가 버튼
     * (btn-add-row 등)은 undo를 잃지 않는다. 내가 누른 삽입·삭제도 서버를 돌아 **원격 트랜잭션으로**
     * 도착하므로 여기서 같이 걸린다(의도한 대로 — 내가 눌렀든 남이 눌렀든 스택 키는 낡는다).
     */
    // 기준값 갱신은 **로컬 이벤트에서도 반드시** 해야 한다(&& 단축 평가로 건너뛰면 로컬 행 추가가
    // 기준값을 낡게 만들어, 이후의 무해한 원격 갱신 하나가 스택을 괜히 날린다).
    const sheetExtentChanged = consumeSheetExtentChange();
    if (!event.transaction.local && sheetExtentChanged) undoManager.clear();
    // 편집 중(contentEditable)에 표를 통째로 다시 그리면 그 <td>가 DOM에서 떨어져 나가며 강제
    // blur가 발생하고, 방금 타이핑하던 값이 커밋되지 못한 채 사라진다(위 blur 핸들러의 세대 판정이
    // 그런 blur를 "버림"으로 처리한다) — 스크롤 리스너가 같은 이유로 쓰는 가드와 같은 판정이다.
    //
    // 다만 스크롤과 달리 구조 변경은 "곧 또 오는" 이벤트가 아니다: 다른 참여자의 행 삽입이나 시트
    // 이름 변경은 한 번뿐이라, 그냥 건너뛰면 편집이 끝나도 그 변경이 화면에 영영 반영되지 않는다
    // (편집 종료 경로는 그 칸 하나만 다시 그린다). 그래서 건너뛰되 반드시 예약해 둔다.
    if (editingKey !== null) {
      pendingStructureRender = true;
      return;
    }
    renderSheetTabs();
    renderTableVirtualized();
  });

  // 병합이 바뀌면(내가·다른 참여자가·실행 취소로) 표를 다시 그린다 — structure.observe와 같은 편집 중 가드
  merges.observe(() => {
    if (editingKey !== null) {
      pendingStructureRender = true;
      return;
    }
    renderTableVirtualized();
    syncCellToolbarState(activeCellKey);
  });

  function currentStructure(): SheetStructure {
    return structure.get(String(activeSheet)) ?? { name: 'Sheet1', rowCount: 1, colCount: 1 };
  }

  // 홈 탭 "셀" 그룹 — 엑셀처럼 **선택한** 행/열 기준으로 삽입·삭제한다(예전엔 시트 맨 끝 행/열만 늘리고 줄였다).
  // 행/열 머리글 오른쪽 클릭 메뉴와 같은 서버 연산(sendGridStructureOp)이라 값·서식·병합·열 너비·행 높이가 함께
  // 밀리고, 저장할 때도 같은 순서로 재생된다. 선택한 줄 수만큼 한 번에 넣고 지운다. 열 전체를 고른 채 "행 삽입"처럼
  // 다른 축 머리글이 선택돼 있으면 그 축의 선택 길이(=시트 전체)가 아니라 활성 칸 한 줄만 쓴다.
  function selectedSpan(axis: 'row' | 'col'): { at: number; count: number } | null {
    if (!activeCellKey) return null;
    const active = parseGridCellKey(activeCellKey);
    if (!active || active.sheetIdx !== activeSheet) return null;
    const b = rangeBounds();
    if (axis === 'row') {
      return headerSelection?.kind === 'col' ? { at: active.row, count: 1 } : { at: b.r0, count: b.r1 - b.r0 + 1 };
    }
    return headerSelection?.kind === 'row' ? { at: active.col, count: 1 } : { at: b.c0, count: b.c1 - b.c0 + 1 };
  }
  /** 지울 줄에 (불러와 둔) 값이 하나라도 있는지 — 행/열 삭제는 Ctrl+Z로 되돌릴 수 없으니 그때만 한 번 묻는다. */
  function spanHasValues(axis: 'row' | 'col', at: number, count: number): boolean {
    const prefix = `${activeSheet}:`;
    for (const [key, value] of cells) {
      if (!key.startsWith(prefix) || value === '') continue;
      const parsed = parseGridCellKey(key);
      if (!parsed) continue;
      const index = axis === 'row' ? parsed.row : parsed.col;
      if (index >= at && index < at + count) return true;
    }
    return false;
  }
  async function confirmSpanDelete(axis: 'row' | 'col', at: number, count: number): Promise<boolean> {
    if (!spanHasValues(axis, at, count)) return true;
    const what = axis === 'row' ? '행' : '열';
    return officeConfirm(`선택한 ${count}개 ${what}을 삭제할까요? 삭제한 ${what}은 실행 취소(Ctrl+Z)로 되돌릴 수 없습니다.`, { title: `${what} 삭제`, okLabel: '삭제', danger: true });
  }
  document.getElementById('btn-add-row')?.addEventListener('click', () => {
    const span = selectedSpan('row');
    if (!span) {
      const s = currentStructure();
      structure.set(String(activeSheet), { ...s, rowCount: s.rowCount + 1 });
      return;
    }
    sendGridStructureOp('insertRows', span.at, span.count);
  });
  document.getElementById('btn-del-row')?.addEventListener('click', async () => {
    const s = currentStructure();
    const span = selectedSpan('row');
    if (!span || span.at >= s.rowCount) return;
    const count = Math.min(span.count, s.rowCount - span.at);
    if (s.rowCount - count < 1 || !(await confirmSpanDelete('row', span.at, count))) return;
    sendGridStructureOp('deleteRows', span.at, count);
  });
  document.getElementById('btn-add-col')?.addEventListener('click', () => {
    const span = selectedSpan('col');
    if (!span) {
      const s = currentStructure();
      structure.set(String(activeSheet), { ...s, colCount: s.colCount + 1 });
      return;
    }
    sendGridStructureOp('insertCols', span.at, span.count);
  });
  document.getElementById('btn-del-col')?.addEventListener('click', async () => {
    const s = currentStructure();
    const span = selectedSpan('col');
    if (!span || span.at >= s.colCount) return;
    const count = Math.min(span.count, s.colCount - span.at);
    if (s.colCount - count < 1 || !(await confirmSpanDelete('col', span.at, count))) return;
    sendGridStructureOp('deleteCols', span.at, count);
  });

  // ---------- 자동 합계 — 엑셀처럼 위쪽(없으면 왼쪽)으로 연속된 숫자 칸을 찾아 SUM 수식을 채운다 ----------
  /** 자동 합계(Σ) 계열 — 위쪽(없으면 왼쪽)으로 연속된 숫자 칸 범위에 함수를 넣는다(Excel과 같은 범위 추정). */
  function insertAutoFunction(fn: AutoFunction): void {
    if (!activeCellKey || !canEdit) return;
    const parsed = parseGridCellKey(activeCellKey);
    if (!parsed) return;
    const { row, col } = parsed;
    const isNumericCell = (r: number, c: number): boolean => {
      const v = cells.get(gridCellKey(activeSheet, r, c));
      return v !== undefined && v !== '' && (NUMERIC_VALUE.test(v.trim()) || v.startsWith('='));
    };
    // 여러 칸을 고른 경우(엑셀과 같다) — 고른 범위의 각 열 바로 아래 칸에 함수를 넣는다. 고른 범위의 마지막 행이
    // 비어 있으면 그 행에 넣는다.
    const range = rangeBounds();
    if (range.r1 > range.r0 || range.c1 > range.c0) {
      const lastRowEmpty = (() => {
        for (let c = range.c0; c <= range.c1; c += 1) if ((cells.get(gridCellKey(activeSheet, range.r1, c)) ?? '') !== '') return false;
        return range.r1 > range.r0;
      })();
      const target = lastRowEmpty ? range.r1 : range.r1 + 1;
      const end = target - 1;
      const writes: Array<[number, string]> = [];
      for (let c = range.c0; c <= range.c1; c += 1) {
        let hasNumber = false;
        for (let r = range.r0; r <= end && !hasNumber; r += 1) hasNumber = isNumericCell(r, c);
        if (!hasNumber) continue;
        const key = gridCellKey(activeSheet, target, c);
        if (styles.get(key)?.lk === 1) continue; // 잠긴 칸
        writes.push([c, `=${fn}(${cellRefLabel(range.r0, c)}:${cellRefLabel(end, c)})`]);
      }
      if (writes.length === 0) return;
      ydoc.transact(() => {
        for (const [c, formula] of writes) cells.set(gridCellKey(activeSheet, target, c), formula);
        const s = currentStructure();
        if (target >= s.rowCount) structure.set(String(activeSheet), { ...s, rowCount: target + 1 });
      });
      for (const [c] of writes) {
        const td = tdFor(target, c);
        if (td) renderCellDisplay(td, gridCellKey(activeSheet, target, c));
      }
      recalcFormulaCells();
      return;
    }
    let formula = '';
    let r = row - 1;
    while (r >= 0 && isNumericCell(r, col)) r -= 1;
    if (r < row - 1) {
      formula = `=${fn}(${cellRefLabel(r + 1, col)}:${cellRefLabel(row - 1, col)})`;
    } else {
      let c = col - 1;
      while (c >= 0 && isNumericCell(row, c)) c -= 1;
      if (c < col - 1) formula = `=${fn}(${cellRefLabel(row, c + 1)}:${cellRefLabel(row, col - 1)})`;
    }
    if (!formula) return;
    const td = tdFor(row, col);
    if (!td || !td.classList.contains('g-editable')) return;
    enterEditMode(td, activeCellKey, formula);
    commitAndMove(td, activeCellKey, row, col, 0, 0);
  }
  for (const button of document.querySelectorAll<HTMLElement>('.autosum-apply')) {
    button.addEventListener('click', () => insertAutoFunction((button.dataset.fn as AutoFunction | undefined) ?? 'SUM'));
  }
  for (const dropdown of document.querySelectorAll<HTMLElement>('.autosum-dropdown')) {
    bindDropdown(dropdown, (value) => insertAutoFunction(value as AutoFunction));
  }
  // Excel의 자동 합계 단축키
  window.addEventListener('keydown', (event) => {
    if (event.altKey && !event.ctrlKey && (event.key === '=' || event.code === 'Equal')) {
      event.preventDefault();
      insertAutoFunction('SUM');
    }
  });

  // ---------- 틀 고정 — 선택한 칸의 위쪽 행 전체 + 왼쪽 열 전체를 고정한다(다시 누르면 해제) ----------
  document.getElementById('btn-freeze-panes')?.addEventListener('click', () => {
    if (!activeCellKey) return;
    const parsed = parseGridCellKey(activeCellKey);
    if (!parsed) return;
    const s = currentStructure();
    const isFrozen = (s.freezeRows ?? 0) > 0 || (s.freezeCols ?? 0) > 0;
    structure.set(
      String(activeSheet),
      isFrozen ? { ...s, freezeRows: 0, freezeCols: 0 } : { ...s, freezeRows: parsed.row, freezeCols: parsed.col },
    );
  });

  // ---------- 서식 도구 (선택 범위 전체에 적용, 오피스처럼 눌린 상태로 반영) ----------
  function setPressed(id: string, pressed: boolean): void {
    document.getElementById(id)?.setAttribute('aria-pressed', String(pressed));
  }
  const DEFAULT_FONT_SIZE_PT = 11;
  function syncCellToolbarState(key: string | null): void {
    const style = key ? styles.get(key) : undefined;
    setPressed('btn-cell-bold', style?.b === 1);
    setPressed('btn-cell-italic', style?.i === 1);
    setPressed('btn-cell-underline', style?.u === 1);
    setPressed('btn-cell-strike', style?.st === 1);
    setPressed('btn-cell-align-l', !style?.a || style.a === 'l');
    setPressed('btn-cell-align-c', style?.a === 'c');
    setPressed('btn-cell-align-r', style?.a === 'r');
    const at = key ? parseGridCellKey(key) : null;
    setPressed('btn-merge-center', Boolean(at && at.sheetIdx === activeSheet && mergeContaining(at.row, at.col)));
    setPressed('btn-valign-top', style?.va === 't');
    setPressed('btn-valign-middle', style?.va === 'm');
    setPressed('btn-valign-bottom', !style?.va || style.va === 'b');
    setPressed('btn-wrap-text', style?.wrap === 1);
    const fontFamilySelect = document.getElementById('cell-font-family') as HTMLSelectElement | null;
    if (fontFamilySelect) fontFamilySelect.value = style?.ff ?? '';
    const fontSizeInput = document.getElementById('cell-font-size') as HTMLInputElement | null;
    if (fontSizeInput) {
      const pt = style?.fs ? Math.round(style.fs * PX_TO_PT) : DEFAULT_FONT_SIZE_PT;
      fontSizeInput.value = String(pt);
    }
    const numfmtSelect = document.getElementById('cell-numfmt-select') as HTMLSelectElement | null;
    if (numfmtSelect) numfmtSelect.value = style?.nf ?? 'general';
  }

  /** 지금 선택 범위(다중 드래그 선택이면 그 사각형 전체, 아니면 칸 하나)에 서식을 적용한다 —
   * 셀마다 "지금 그 칸의 서식"을 기준으로 mutate가 따로 호출되므로, 굵게 상태가 섞인 범위를
   * 선택해도 엑셀처럼 칸별로 자연스럽게 토글된다. */
  function updateStyle(mutate: (current: CellStyleAttrs) => CellStyleAttrs): void {
    if (!activeCellKey) return;
    const { r0, r1, c0, c1 } = rangeBounds();
    ydoc.transact(() => {
      for (let r = r0; r <= r1; r += 1) {
        for (let c = c0; c <= c1; c += 1) {
          const key = gridCellKey(activeSheet, r, c);
          const current = styles.get(key) ?? {};
          const next = mutate(current);
          // 결과가 빈 서식이어도 **삭제가 아니라 빈 객체를 쓴다** — Y.Map.delete는 그 키가 지금 이
          // 클라이언트의 Yjs 맵에 들어 있을 때만 연산을 만든다. 서식이 페이지 단위로 윈도잉된 뒤로는
          // "화면에는 그려져 있지만 아직 시딩되지 않은 칸"이 흔하고, 거기서의 delete는 아무 이벤트도
          // 만들지 못해 서버가 영영 모르는 조용한 무동작이 된다(서식 지우기가 파일에 반영되지 않음).
          // 빈 객체 쓰기는 키의 로컬 존재 여부와 무관하게 항상 실제 연산이 되고, 서버의
          // applyStylesChangeToCanonical은 빈 서식을 삭제와 똑같이 "서식만 뗀다"로 처리한다.
          styles.set(key, next);
          const td = tdFor(r, c);
          if (td) {
            applyStyleToTd(td, styles.get(key));
            renderCellDisplay(td, key); // 숫자 서식은 텍스트 자체가 바뀌므로 다시 그린다
          }
        }
      }
    });
    // 툴바 버튼을 누르면 버튼이 포커스를 가져가 셀은 이미 blur된 상태이므로 :focus로는 못 찾는다
    // — activeCellKey(범위의 기준 칸)를 좌표로 풀어 직접 찾아 눌림 상태를 갱신한다.
    syncCellToolbarState(activeCellKey);
  }
  function toggleFlag(flag: 'b' | 'i' | 'u' | 'st' | 'wrap'): void {
    updateStyle((current) => {
      const next = { ...current };
      if (next[flag]) delete next[flag];
      else next[flag] = 1;
      return next;
    });
  }
  document.getElementById('btn-cell-bold')?.addEventListener('click', () => toggleFlag('b'));
  document.getElementById('btn-cell-italic')?.addEventListener('click', () => toggleFlag('i'));
  document.getElementById('btn-cell-underline')?.addEventListener('click', () => toggleFlag('u'));
  document.getElementById('btn-cell-strike')?.addEventListener('click', () => toggleFlag('st'));
  document.getElementById('btn-wrap-text')?.addEventListener('click', () => toggleFlag('wrap'));
  document.getElementById('btn-cell-clear-format')?.addEventListener('click', () => {
    if (!activeCellKey) return;
    const { r0, r1, c0, c1 } = rangeBounds();
    ydoc.transact(() => {
      for (let r = r0; r <= r1; r += 1) {
        for (let c = c0; c <= c1; c += 1) {
          const key = gridCellKey(activeSheet, r, c);
          // delete가 아니라 "빈 서식" 표식을 쓴다 — updateStyle의 주석 참고. 열 전체/전체 선택처럼
          // 시딩되지 않은 페이지까지 걸치는 지우기에서 delete는 대부분의 칸에서 조용한 무동작이었다.
          styles.set(key, {});
          const td = tdFor(r, c);
          if (td) {
            applyStyleToTd(td, undefined);
            renderCellDisplay(td, key);
          }
        }
      }
    });
    syncCellToolbarState(activeCellKey);
  });

  // 글자 색/배경 색 — 아이콘을 누르면 막대에 표시된 마지막 색이 바로 적용되고, 캐럿을
  // 눌러야 네이티브 색상 선택기가 뜬다(엑셀과 동일하게 "적용 버튼"과 "색 고르기"를 분리).
  const cellColorInput = document.getElementById('cell-font-color') as HTMLInputElement | null;
  const cellColorBtn = document.getElementById('cell-font-color-btn') as HTMLElement | null;
  cellColorBtn?.style.setProperty('--btn-color', cellColorInput?.value ?? '#c0392b');
  cellColorInput?.addEventListener('input', () => {
    cellColorBtn?.style.setProperty('--btn-color', cellColorInput.value);
    updateStyle((current) => ({ ...current, col: cellColorInput.value }));
  });
  document.getElementById('btn-font-color-apply')?.addEventListener('click', () => {
    if (!cellColorInput) return;
    updateStyle((current) => ({ ...current, col: cellColorInput.value }));
  });
  const cellBgInput = document.getElementById('cell-bg-color') as HTMLInputElement | null;
  const cellBgBtn = document.getElementById('cell-bg-color-btn') as HTMLElement | null;
  cellBgBtn?.style.setProperty('--btn-color', cellBgInput?.value ?? '#ffff00');
  cellBgInput?.addEventListener('input', () => {
    cellBgBtn?.style.setProperty('--btn-color', cellBgInput.value);
    updateStyle((current) => ({ ...current, bg: cellBgInput.value }));
  });
  document.getElementById('btn-bg-color-apply')?.addEventListener('click', () => {
    if (!cellBgInput) return;
    updateStyle((current) => ({ ...current, bg: cellBgInput.value }));
  });

  (['l', 'c', 'r'] as const).forEach((align) => {
    document.getElementById(`btn-cell-align-${align}`)?.addEventListener('click', () => {
      updateStyle((current) => ({ ...current, a: align }));
    });
  });
  (['t', 'm', 'b'] as const).forEach((va) => {
    const id = va === 't' ? 'btn-valign-top' : va === 'm' ? 'btn-valign-middle' : 'btn-valign-bottom';
    document.getElementById(id)?.addEventListener('click', () => {
      updateStyle((current) => ({ ...current, va }));
    });
  });

  // ---------- 맞춤: 병합하고 가운데 맞춤 (엑셀과 같은 동작) ----------
  // 병합 목록(GRID_MERGES_MAP, 1부터 센 좌표)은 시트마다 통째로 들어 있어(윈도잉하지 않는다) 여기서 바로 고쳐 쓴다.
  // 서버는 이 변경을 시트 메타데이터에 반영해 두고(room.ts merges.observe), 저장할 때 그 시트의 병합을 다시 쓴다.
  type Rect = { r0: number; r1: number; c0: number; c1: number };
  // 함수 선언으로 둔다 — syncCellToolbarState(위쪽)가 이 줄이 실행되기 전에도 불릴 수 있다
  function mergeRect(m: GridMergeShape): Rect {
    return { r0: m.r - 1, r1: m.r - 2 + m.rs, c0: m.c - 1, c1: m.c - 2 + m.cs };
  }
  function rectsIntersect(a: Rect, b: Rect): boolean {
    return !(a.r0 > b.r1 || a.r1 < b.r0 || a.c0 > b.c1 || a.c1 < b.c0);
  }
  function sheetMergeList(): GridMergeShape[] {
    return [...(merges.get(String(activeSheet)) ?? [])];
  }
  function mergeContaining(row: number, col: number): GridMergeShape | undefined {
    return sheetMergeList().find((m) => rectsIntersect(mergeRect(m), { r0: row, r1: row, c0: col, c1: col }));
  }
  function unmergeSelection(): void {
    if (!canEdit || !activeCellKey) return;
    const sel = rangeBounds();
    const list = sheetMergeList();
    const kept = list.filter((m) => !rectsIntersect(mergeRect(m), sel));
    if (kept.length === list.length) return;
    ydoc.transact(() => {
      if (kept.length > 0) merges.set(String(activeSheet), kept);
      else merges.delete(String(activeSheet));
    });
    syncCellToolbarState(activeCellKey);
  }
  async function applyMerge(mode: 'center' | 'merge' | 'across'): Promise<void> {
    if (!canEdit || !activeCellKey) return;
    const list = sheetMergeList();
    const sel = rangeBounds();
    // 병합된 칸(하나의 병합 안)을 고른 채 "병합하고 가운데 맞춤"을 누르면 엑셀처럼 병합을 푼다
    const inside = list.find((m) => {
      const r = mergeRect(m);
      return sel.r0 >= r.r0 && sel.r1 <= r.r1 && sel.c0 >= r.c0 && sel.c1 <= r.c1;
    });
    if (inside && mode === 'center') {
      unmergeSelection();
      return;
    }
    // 고른 범위와 겹치는 기존 병합까지 감싸도록 넓힌다(엑셀도 병합 일부만 걸친 선택을 병합 전체로 넓힌다)
    const box: Rect = { ...sel };
    for (let grown = true; grown; ) {
      grown = false;
      for (const m of list) {
        const r = mergeRect(m);
        if (!rectsIntersect(r, box)) continue;
        if (r.r0 < box.r0 || r.r1 > box.r1 || r.c0 < box.c0 || r.c1 > box.c1) {
          box.r0 = Math.min(box.r0, r.r0);
          box.r1 = Math.max(box.r1, r.r1);
          box.c0 = Math.min(box.c0, r.c0);
          box.c1 = Math.max(box.c1, r.c1);
          grown = true;
        }
      }
    }
    const rects: Rect[] =
      mode === 'across'
        ? Array.from({ length: box.r1 - box.r0 + 1 }, (_, i) => ({ r0: box.r0 + i, r1: box.r0 + i, c0: box.c0, c1: box.c1 }))
        : [box];
    const targets = rects.filter((r) => r.r1 > r.r0 || r.c1 > r.c0);
    if (targets.length === 0) {
      if (mode === 'center') updateStyle((current) => ({ ...current, a: 'c' })); // 칸 하나 — 엑셀처럼 가운데 맞춤만
      return;
    }
    if ((box.r1 - box.r0 + 1) * (box.c1 - box.c0 + 1) > 200_000) {
      void officeAlert('범위가 너무 넓어 병합할 수 없습니다.');
      return;
    }
    const drop: string[] = [];
    for (const rect of targets) {
      for (let r = rect.r0; r <= rect.r1; r += 1) {
        for (let c = rect.c0; c <= rect.c1; c += 1) {
          const key = gridCellKey(activeSheet, r, c);
          if (styles.get(key)?.lk === 1) {
            void officeAlert('잠긴 칸이 있어 병합할 수 없습니다.');
            return;
          }
          if ((r !== rect.r0 || c !== rect.c0) && (cells.get(key) ?? '') !== '') drop.push(key);
        }
      }
    }
    if (drop.length > 0 && !(await officeConfirm('셀을 병합하면 왼쪽 위의 값만 남고 나머지 값은 지워집니다. 계속할까요?', { title: '셀 병합' }))) return;
    ydoc.transact(() => {
      for (const key of drop) cells.delete(key);
      const kept = list.filter((m) => !targets.some((t) => rectsIntersect(mergeRect(m), t)));
      for (const t of targets) kept.push({ r: t.r0 + 1, c: t.c0 + 1, rs: t.r1 - t.r0 + 1, cs: t.c1 - t.c0 + 1 });
      merges.set(String(activeSheet), kept);
      if (mode === 'center') {
        for (const t of targets) {
          const key = gridCellKey(activeSheet, t.r0, t.c0);
          styles.set(key, { ...(styles.get(key) ?? {}), a: 'c' });
        }
      }
    });
    syncCellToolbarState(activeCellKey);
  }
  // ---------- 스타일: 조건부 서식 (엑셀 "조건부 서식" 메뉴) ----------
  // 규칙은 GRID_COND_FORMATS_MAP에 시트마다 통째로 있다(파일 규칙은 서버가 룸을 열 때 채운다). 화면 칸마다 판정해
  // 기본 서식 위에 색을 덧입힌다(condFormatOverlay — applyStyleToTd·renderCellDisplay 끝에서 불린다).
  // 중복·상위·평균·색조·막대는 범위 전체 값이 필요한데, 아직 불러오지 않은 페이지의 값은 모르므로 불러온 값으로 계산한다.
  const cfEntries = (): readonly CondFormatEntry[] => condFormats.get(String(activeSheet)) ?? [];
  function cfValueAt(row: number, col: number): string {
    const raw = cells.get(gridCellKey(activeSheet, row, col)) ?? '';
    return raw.startsWith('=') ? formatComputedValue(computeCellDisplay(activeSheet, row, col)) : raw;
  }
  let cfStats: Map<string, CfStats> | null = null;
  let cfStatsSheet = -1;
  function cfStatsNow(): ReadonlyMap<string, CfStats> {
    if (cfStats && cfStatsSheet === activeSheet) return cfStats;
    const next = new Map<string, CfStats>();
    const withStats = cfEntries().filter((e) => needsStats(e.rule));
    if (withStats.length > 0) {
      const values = new Map<string, string[]>(withStats.map((e) => [e.id, []]));
      const prefix = `${activeSheet}:`;
      for (const key of cells.keys()) {
        if (!key.startsWith(prefix)) continue;
        const at = parseGridCellKey(key);
        if (!at) continue;
        for (const e of withStats) if (e.ranges.some((r) => rectContains(r, at.row, at.col))) values.get(e.id)?.push(cfValueAt(at.row, at.col));
      }
      for (const e of withStats) next.set(e.id, computeStats(e.rule, values.get(e.id) ?? []));
    }
    cfStats = next;
    cfStatsSheet = activeSheet;
    return next;
  }
  function applyCondFormatOverlay(td: HTMLTableCellElement): void {
    const row = Number(td.dataset.row);
    const col = Number(td.dataset.col);
    if (!Number.isInteger(row) || !Number.isInteger(col) || td.dataset.row === undefined) return;
    const entries = cfEntries();
    if (entries.length === 0 && td.dataset.cf === undefined) return;
    const paint = entries.length > 0 ? paintCell(entries, row, col, cfValueAt(row, col), cfStatsNow(), cfValueAt) : null;
    if (!paint && td.dataset.cf === undefined) return;
    // 기본 서식(applyStyleToTd와 같은 값)으로 되돌린 뒤 규칙 색을 얹는다 — 규칙이 더는 안 맞으면 원래 모습이 된다
    const base = styles.get(gridCellKey(activeSheet, row, col));
    td.style.background = paint?.style?.bg ?? paint?.scale ?? base?.bg ?? '';
    td.style.color = paint?.style?.col ?? base?.col ?? '';
    td.style.fontWeight = paint?.style?.b || base?.b ? '700' : '400';
    if (paint?.bar) {
      td.style.backgroundImage = `linear-gradient(90deg, ${paint.bar.color} 0%, ${paint.bar.color}33 100%)`;
      td.style.backgroundSize = `${Math.round(paint.bar.ratio * 100)}% 76%`;
      td.style.backgroundRepeat = 'no-repeat';
      td.style.backgroundPosition = 'left center';
    }
    if (paint) td.dataset.cf = '1';
    else delete td.dataset.cf;
  }
  condFormatOverlay = applyCondFormatOverlay;
  // 값·계산 결과·규칙이 바뀌면 통계를 버리고 화면 칸을 한 번에 다시 칠한다(한 칸이 바뀌어도 상위 10개·중복 판정은 다른 칸이 바뀐다)
  let cfRepaintQueued = false;
  function scheduleCfRepaint(): void {
    cfStats = null;
    if (cfRepaintQueued) return;
    if (cfEntries().length === 0 && !tableEl.querySelector('td[data-cf]')) return;
    cfRepaintQueued = true;
    requestAnimationFrame(() => {
      cfRepaintQueued = false;
      tableEl.querySelectorAll<HTMLTableCellElement>('td[data-row]').forEach(applyCondFormatOverlay);
    });
  }
  cells.observe(scheduleCfRepaint);
  computed.observe(scheduleCfRepaint);
  condFormats.observe(scheduleCfRepaint);

  function addCondFormat(rule: CfRule): void {
    if (!canEdit || !activeCellKey) return;
    const sel = rangeBounds();
    const entry: CondFormatEntry = { id: `u${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`, ranges: [sel], rule };
    // 새 규칙이 맨 위(우선순위 가장 높음) — 엑셀과 같다
    condFormats.set(String(activeSheet), [entry, ...cfEntries()]);
  }
  function clearCondFormats(scope: 'sel' | 'sheet'): void {
    if (!canEdit) return;
    const entries = cfEntries();
    if (entries.length === 0) return;
    if (scope === 'sheet') {
      condFormats.delete(String(activeSheet));
      return;
    }
    if (!activeCellKey) return;
    const sel = rangeBounds();
    const next = entries.map((e) => ({ ...e, ranges: e.ranges.flatMap((r) => subtractRect(r, sel)) })).filter((e) => e.ranges.length > 0);
    if (next.length === entries.length && next.every((e, i) => e.ranges.length === entries[i].ranges.length && e.ranges.every((r, j) => r === entries[i].ranges[j]))) return;
    if (next.length > 0) condFormats.set(String(activeSheet), next);
    else condFormats.delete(String(activeSheet));
  }

  // 규칙 입력 창(엑셀 "보다 큼" 등 대화 상자와 같은 모양)
  type CfDialogKind = 'gt' | 'lt' | 'between' | 'eq' | 'text' | 'dup' | 'top' | 'bottom' | 'aboveAvg' | 'belowAvg';
  const CF_DIALOG: Record<CfDialogKind, { title: string; sub: string; v1?: string; v2?: boolean; dup?: boolean; def?: string }> = {
    gt: { title: '보다 큼', sub: '다음 값보다 큰 셀의 서식 지정:', v1: '값' },
    lt: { title: '보다 작음', sub: '다음 값보다 작은 셀의 서식 지정:', v1: '값' },
    between: { title: '다음 값의 사이에 있음', sub: '다음 값 사이에 있는 셀의 서식 지정:', v1: '값', v2: true },
    eq: { title: '같음', sub: '다음 값과 같은 셀의 서식 지정:', v1: '값' },
    text: { title: '텍스트 포함', sub: '다음 텍스트를 포함하는 셀의 서식 지정:', v1: '텍스트' },
    dup: { title: '중복 값', sub: '다음 값을 포함하는 셀의 서식 지정:', dup: true },
    top: { title: '상위 10개 항목', sub: '상위 순위에 속하는 셀의 서식 지정:', v1: '개수', def: '10' },
    bottom: { title: '하위 10개 항목', sub: '하위 순위에 속하는 셀의 서식 지정:', v1: '개수', def: '10' },
    aboveAvg: { title: '평균 초과', sub: '선택한 범위에서 평균을 초과하는 셀의 서식 지정:' },
    belowAvg: { title: '평균 미만', sub: '선택한 범위에서 평균 미만인 셀의 서식 지정:' },
  };
  const cfDialog = document.getElementById('cf-dialog') as HTMLElement | null;
  const cfV1 = document.getElementById('cf-dialog-v1') as HTMLInputElement | null;
  const cfV2 = document.getElementById('cf-dialog-v2') as HTMLInputElement | null;
  const cfDup = document.getElementById('cf-dialog-dup') as HTMLSelectElement | null;
  const cfStyleSelect = document.getElementById('cf-dialog-style') as HTMLSelectElement | null;
  if (cfStyleSelect) {
    for (const preset of CF_PRESETS) {
      const option = document.createElement('option');
      option.value = preset.key;
      option.textContent = preset.label;
      cfStyleSelect.append(option);
    }
  }
  let cfDialogKind: CfDialogKind | null = null;
  function openCfDialog(kind: CfDialogKind): void {
    if (!canEdit || !activeCellKey || !cfDialog || !cfV1 || !cfV2) return;
    const spec = CF_DIALOG[kind];
    cfDialogKind = kind;
    (document.getElementById('cf-dialog-title') as HTMLElement).textContent = spec.title;
    (document.getElementById('cf-dialog-sub') as HTMLElement).textContent = spec.sub;
    (document.getElementById('cf-dialog-v1-row') as HTMLElement).hidden = !spec.v1;
    (document.getElementById('cf-dialog-v1-label') as HTMLElement).textContent = spec.v1 ?? '';
    (document.getElementById('cf-dialog-v2-row') as HTMLElement).hidden = !spec.v2;
    (document.getElementById('cf-dialog-dup-row') as HTMLElement).hidden = !spec.dup;
    cfV1.value = spec.def ?? '';
    cfV1.inputMode = kind === 'top' || kind === 'bottom' ? 'numeric' : 'text';
    cfV2.value = '';
    if (cfStyleSelect) cfStyleSelect.value = CF_PRESETS[0].key; // 엑셀처럼 열 때마다 기본 서식(연한 빨강 채우기)으로
    if (cfDup) cfDup.value = 'dup';
    cfDialog.hidden = false;
    (spec.v1 ? cfV1 : cfStyleSelect)?.focus();
  }
  const closeCfDialog = () => {
    if (cfDialog) cfDialog.hidden = true;
    cfDialogKind = null;
  };
  function confirmCfDialog(): void {
    const kind = cfDialogKind;
    if (!kind || !cfV1 || !cfV2 || !cfDup || !cfStyleSelect) return;
    const s = (CF_PRESETS.find((p) => p.key === cfStyleSelect.value) ?? CF_PRESETS[0]).s;
    const v1 = cfV1.value.trim();
    const v2 = cfV2.value.trim();
    let rule: CfRule | null = null;
    if (kind === 'gt' || kind === 'lt' || kind === 'eq') rule = v1 ? { t: 'cell', op: kind, v1, s } : null;
    else if (kind === 'between') rule = v1 && v2 ? { t: 'cell', op: 'between', v1, v2, s } : null;
    else if (kind === 'text') rule = v1 ? { t: 'text', op: 'contains', text: v1, s } : null;
    else if (kind === 'dup') rule = { t: 'dup', ...(cfDup.value === 'unique' ? { unique: 1 as const } : {}), s };
    else if (kind === 'top' || kind === 'bottom') {
      const n = Number(v1);
      rule = Number.isInteger(n) && n >= 1 && n <= 1000 ? { t: 'top', n, ...(kind === 'bottom' ? { bottom: 1 as const } : {}), s } : null;
    } else rule = { t: 'avg', ...(kind === 'belowAvg' ? { below: 1 as const } : {}), s };
    if (!rule) {
      cfV1.focus();
      return;
    }
    closeCfDialog();
    addCondFormat(rule);
  }
  document.getElementById('cf-dialog-ok')?.addEventListener('click', confirmCfDialog);
  document.getElementById('cf-dialog-cancel')?.addEventListener('click', closeCfDialog);
  cfDialog?.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') closeCfDialog();
    if (event.key === 'Enter' && (event.target as HTMLElement).tagName !== 'SELECT') {
      event.preventDefault();
      confirmCfDialog();
    }
  });
  bindDropdown(document.getElementById('cf-dropdown'), (value) => {
    if (value === 'clear-sel') clearCondFormats('sel');
    else if (value === 'clear-sheet') clearCondFormats('sheet');
    else if (value.startsWith('bar:')) addCondFormat({ t: 'bar', color: value.slice(4) });
    else if (value.startsWith('scale:')) addCondFormat({ t: 'scale', colors: value.slice(6).split(',') });
    else openCfDialog(value as CfDialogKind);
  });

  document.getElementById('btn-merge-center')?.addEventListener('click', () => applyMerge('center'));
  bindDropdown(document.getElementById('merge-dropdown'), (value) => {
    if (value === 'unmerge') unmergeSelection();
    else applyMerge(value as 'center' | 'merge' | 'across');
  });

  // ---------- 글꼴 이름/크기 ----------
  const fontFamilySelect = document.getElementById('cell-font-family') as HTMLSelectElement | null;
  fontFamilySelect?.addEventListener('change', () => {
    updateStyle((current) => {
      const next = { ...current };
      if (fontFamilySelect.value) next.ff = fontFamilySelect.value;
      else delete next.ff;
      return next;
    });
  });
  const fontSizeInput = document.getElementById('cell-font-size') as HTMLInputElement | null;
  function applyFontSizePt(pt: number): void {
    const clamped = Math.max(6, Math.min(96, Math.round(pt)));
    if (fontSizeInput) fontSizeInput.value = String(clamped);
    updateStyle((current) => {
      const next = { ...current };
      if (clamped === DEFAULT_FONT_SIZE_PT) delete next.fs;
      else next.fs = Math.round(clamped * PT_TO_PX);
      return next;
    });
  }
  fontSizeInput?.addEventListener('change', () => {
    const pt = Number(fontSizeInput.value);
    applyFontSizePt(Number.isFinite(pt) && pt > 0 ? pt : DEFAULT_FONT_SIZE_PT);
  });
  document.getElementById('btn-font-size-up')?.addEventListener('click', () => {
    applyFontSizePt(Number(fontSizeInput?.value ?? DEFAULT_FONT_SIZE_PT) + 1);
  });
  document.getElementById('btn-font-size-down')?.addEventListener('click', () => {
    applyFontSizePt(Number(fontSizeInput?.value ?? DEFAULT_FONT_SIZE_PT) - 1);
  });

  // ---------- 표시 형식(숫자 서식) ----------
  const numfmtSelect = document.getElementById('cell-numfmt-select') as HTMLSelectElement | null;
  numfmtSelect?.addEventListener('change', () => {
    updateStyle((current) => {
      const next = { ...current };
      if (numfmtSelect.value === 'general') {
        delete next.nf;
        delete next.nd;
      } else {
        next.nf = numfmtSelect.value as CellStyleAttrs['nf'];
        if (!next.nd) next.nd = next.nf === 'currency' ? 2 : 0;
      }
      return next;
    });
  });
  function applyQuickNumFmt(nf: NonNullable<CellStyleAttrs['nf']>, nd: number): void {
    updateStyle((current) => ({ ...current, nf, nd }));
  }
  document.getElementById('btn-numfmt-currency')?.addEventListener('click', () => applyQuickNumFmt('currency', 2));
  document.getElementById('btn-numfmt-percent')?.addEventListener('click', () => applyQuickNumFmt('percent', 0));
  document.getElementById('btn-numfmt-comma')?.addEventListener('click', () => applyQuickNumFmt('number', 0));
  function stepDecimals(delta: number): void {
    updateStyle((current) => {
      const nextDecimals = Math.max(0, (current.nd ?? 0) + delta);
      return { ...current, nf: current.nf ?? 'number', nd: nextDecimals };
    });
  }
  document.getElementById('btn-numfmt-decimal-up')?.addEventListener('click', () => stepDecimals(1));
  document.getElementById('btn-numfmt-decimal-down')?.addEventListener('click', () => stepDecimals(-1));

  // ---------- 테두리 ----------
  type BorderPreset = 'none' | 'all' | 'outside' | 'outside-thick' | 'top' | 'bottom' | 'left' | 'right';
  function applyBorderPreset(preset: BorderPreset): void {
    if (!activeCellKey) return;
    const { r0, r1, c0, c1 } = rangeBounds();
    ydoc.transact(() => {
      for (let r = r0; r <= r1; r += 1) {
        for (let c = c0; c <= c1; c += 1) {
          const key = gridCellKey(activeSheet, r, c);
          const next: { -readonly [K in keyof CellStyleAttrs]?: CellStyleAttrs[K] } = { ...(styles.get(key) ?? {}) };
          const isTop = r === r0;
          const isBottom = r === r1;
          const isLeft = c === c0;
          const isRight = c === c1;
          const weight = preset === 'outside-thick' ? 2 : 1;
          switch (preset) {
            case 'none':
              delete next.bt; delete next.br; delete next.bb; delete next.bl;
              break;
            case 'all':
              next.bt = 1; next.br = 1; next.bb = 1; next.bl = 1;
              break;
            case 'outside':
            case 'outside-thick':
              if (isTop) next.bt = weight; else delete next.bt;
              if (isBottom) next.bb = weight; else delete next.bb;
              if (isLeft) next.bl = weight; else delete next.bl;
              if (isRight) next.br = weight; else delete next.br;
              break;
            case 'top':
              if (isTop) next.bt = 1;
              break;
            case 'bottom':
              if (isBottom) next.bb = 1;
              break;
            case 'left':
              if (isLeft) next.bl = 1;
              break;
            case 'right':
              if (isRight) next.br = 1;
              break;
          }
          // 테두리 'none'으로 서식이 통째로 비어도 delete가 아니라 빈 객체를 쓴다 — updateStyle의 주석 참고.
          styles.set(key, next as CellStyleAttrs);
          const td = tdFor(r, c);
          if (td) applyStyleToTd(td, styles.get(key));
        }
      }
    });
  }
  bindDropdown(document.getElementById('border-dropdown'), (value) => applyBorderPreset(value as BorderPreset));
  document.getElementById('btn-border')?.addEventListener('click', () => applyBorderPreset('all'));

  // ---------- 삽입: 차트·표 서식 ----------
  const sheetIndexByName = (name: string): number => {
    let found = -1;
    structure.forEach((s, key) => {
      if (found < 0 && s.name === name) found = Number(key);
    });
    return found;
  };
  const valueText = (sheetName: string, row: number, col: number): string => {
    const idx = sheetIndexByName(sheetName);
    return idx < 0 ? '' : formatComputedValue(computeCellDisplay(idx, row, col));
  };
  const headerHeight = (): number => tableEl.tHead?.offsetHeight || DEFAULT_ROW_HEIGHT_PX;
  const rownumWidth = (): number => tableEl.querySelector<HTMLElement>('.g-edit-rownum')?.offsetWidth || ROWNUM_COL_WIDTH_PX;
  const colWidthAt = (col: number): number => {
    const sheet = structure.get(String(activeSheet));
    return sheet?.colWidths?.[col] ?? sheet?.defaultColWidth ?? DEFAULT_COL_WIDTH_PX;
  };
  const colLeft = (col: number): number => {
    let x = rownumWidth();
    for (let c = 0; c < col; c += 1) x += colWidthAt(c);
    return x;
  };
  const charts = setupGridCharts({
    ydoc,
    token,
    canEdit,
    scrollHost: scrollHostEl,
    table: tableEl,
    activeSheetName: () => structure.get(String(activeSheet))?.name ?? '',
    valueText,
    colLeft,
    rowTop: (row) => headerHeight() + rowLayout.top(row),
    cellAt: (x, y) => {
      let col = 0;
      let left = rownumWidth();
      while (col < 16384 && left + colWidthAt(col) <= x) {
        left += colWidthAt(col);
        col += 1;
      }
      return { row: rowLayout.rowAt(y - headerHeight()), col };
    },
    selection: () => {
      const { r0, r1, c0, c1 } = rangeBounds();
      return { r1: r0, c1: c0, r2: r1, c2: c1 };
    },
    extent: () => {
      const sheet = structure.get(String(activeSheet));
      return { rows: sheet?.rowCount ?? 1, cols: sheet?.colCount ?? 1 };
    },
    setChartTabVisible: (visible) => shell.setContextualTab('chart', visible),
  });
  afterGridRender = () => charts.render();
  gridLinks = setupGridLinks({
    token,
    tableEl,
    scrollHost: scrollHostEl,
    currentSheetName: () => structure.get(String(activeSheet))?.name ?? '',
    goTo: (sheetName, row, col) => {
      const target = sheetName === null ? activeSheet : [...structure.entries()].find(([, s]) => s.name === sheetName)?.[0];
      if (target === undefined) return false;
      const idx = Number(target);
      if (idx !== activeSheet) {
        activeSheet = idx;
        resetSelectionForSheetSwitch();
        renderSheetTabs();
        renderTableVirtualized();
      }
      const td = tdForEnsureRendered(row, col);
      if (td) {
        td.focus();
        scrollCellIntoView(td);
      }
      return true;
    },
    valueText,
    pickValue: (td, row, col, value) => {
      const key = gridCellKey(activeSheet, row, col);
      enterEditMode(td, key, value);
      if (editingKey !== key) return; // 잠긴 칸이거나 다른 사람이 편집 중
      commitAndMove(td, key, row, col, 0, 0);
    },
    isEditing: (td) => td.contentEditable === 'true',
  });
  let chartRenderScheduled = false;
  const scheduleChartRender = () => {
    if (chartRenderScheduled) return;
    chartRenderScheduled = true;
    requestAnimationFrame(() => {
      chartRenderScheduled = false;
      charts.render();
    });
  };
  cells.observe(scheduleChartRender);
  computed.observe(scheduleChartRender);
  charts.render();
  for (const button of document.querySelectorAll<HTMLElement>('[data-insert-chart]')) {
    button.addEventListener('click', () => charts.insert(button.dataset.insertChart as ChartKind));
  }

  /** 표 서식 — 머리글 행(진한 색·흰 굵은 글자)과 줄무늬 행, 모든 테두리. 칸 하나만 골랐으면 둘러싼 데이터 영역. */
  const TABLE_STYLES: Readonly<Record<string, { head: string; band: string }>> = {
    blue: { head: '#4472c4', band: '#d9e1f2' },
    orange: { head: '#ed7d31', band: '#fce4d6' },
    green: { head: '#70ad47', band: '#e2efda' },
    gold: { head: '#ffc000', band: '#fff2cc' },
    gray: { head: '#595959', band: '#ededed' },
  };
  function applyTableStyle(preset: string): void {
    const colors = TABLE_STYLES[preset];
    if (!colors || !activeCellKey || !canEdit) return;
    const sel = rangeBounds();
    const sheetName = structure.get(String(activeSheet))?.name ?? '';
    const sheet = structure.get(String(activeSheet));
    const region =
      sel.r0 === sel.r1 && sel.c0 === sel.c1
        ? currentRegion({ row: sel.r0, col: sel.c0 }, (r, c) => valueText(sheetName, r, c).trim() !== '', { rows: sheet?.rowCount ?? 1, cols: sheet?.colCount ?? 1 })
        : { r1: sel.r0, c1: sel.c0, r2: sel.r1, c2: sel.c1 };
    if (!region) {
      void officeAlert('표 서식을 입힐 데이터 범위를 선택하세요.');
      return;
    }
    ydoc.transact(() => {
      for (let r = region.r1; r <= region.r2; r += 1) {
        for (let c = region.c1; c <= region.c2; c += 1) {
          const key = gridCellKey(activeSheet, r, c);
          const current: { -readonly [K in keyof CellStyleAttrs]?: CellStyleAttrs[K] } = { ...(styles.get(key) ?? {}) };
          current.bt = 1;
          current.bb = 1;
          current.bl = 1;
          current.br = 1;
          if (r === region.r1) {
            current.bg = colors.head;
            current.col = preset === 'gold' ? '#000000' : '#ffffff';
            current.b = 1;
          } else {
            delete current.b;
            delete current.col;
            if ((r - region.r1) % 2 === 1) current.bg = colors.band;
            else delete current.bg;
          }
          styles.set(key, current as CellStyleAttrs);
          const td = tdFor(r, c);
          if (td) {
            applyStyleToTd(td, styles.get(key));
            renderCellDisplay(td, key);
          }
        }
      }
    });
  }
  bindDropdown(document.getElementById('table-style-dropdown'), applyTableStyle);
  // 홈 탭 "스타일"의 표 서식 — 삽입 탭과 같은 함수
  bindDropdown(document.getElementById('home-table-style-dropdown'), applyTableStyle);

  // ---------- 홈: 클립보드 ----------
  // 브라우저는 스크립트의 클립보드 읽기를 권한으로 막을 수 있다 — 그때는 워드 편집기처럼 Ctrl+V를 안내한다.
  document.getElementById('btn-copy')?.addEventListener('click', () => copySelectionToClipboard());
  document.getElementById('btn-cut')?.addEventListener('click', () => {
    copySelectionToClipboard();
    if (canEdit) clearSelectionValues();
  });
  document.getElementById('btn-paste')?.addEventListener('click', async () => {
    if (!canEdit) return;
    const { r0, c0 } = rangeBounds();
    let text = '';
    try {
      text = await navigator.clipboard.readText();
    } catch {
      void officeAlert('브라우저가 클립보드 읽기를 허용하지 않았습니다. Ctrl+V로 붙여넣으세요.');
      return;
    }
    if (text) pasteTextAt(text, r0, c0);
  });

  // 홈 탭 "정렬 및 필터" — 데이터 탭의 같은 단추를 누른다(기능·상태는 그 단추 하나에만 있다)
  bindDropdown(document.getElementById('home-sort-filter-dropdown'), (id) => document.getElementById(id)?.click());

  // ---------- 데이터: 정렬·필터 ----------
  const filtersMap = ydoc.getMap<SheetFilter>(GRID_FILTERS_MAP);
  const activeSheetName = () => structure.get(String(activeSheet))?.name ?? '';
  const displayAt = (r: number, c: number) => valueText(activeSheetName(), r, c);
  /** 서버가 실제로 채워 준(도착한) 페이지 안의 행인지 — 정렬·필터는 도착한 데이터로만 계산한다. 예전엔 "요청한"
   * 페이지로 판정해서, 데이터가 오기 전에 정렬하면 빈 칸을 기준으로 줄을 옮길 수 있었고 필터 값 목록도 비어 보였다. */
  const rowsLoaded = (r1: number, r2: number) => {
    const arrived = arrivedPages.get(activeSheet);
    if (!arrived) return false;
    for (let page = Math.floor(r1 / PAGE_SIZE_ROWS); page <= Math.floor(r2 / PAGE_SIZE_ROWS); page += 1) {
      if (!arrived.has(page)) return false;
    }
    return true;
  };
  const rowArrived = (r: number) => arrivedPages.get(activeSheet)?.has(Math.floor(r / PAGE_SIZE_ROWS)) ?? false;
  /** 칸 하나만 골랐으면 둘러싼 데이터 영역, 범위를 골랐으면 그 범위 */
  function dataRegion(): CellRangeRef | null {
    const sel = rangeBounds();
    const sheet = structure.get(String(activeSheet));
    if (sel.r0 !== sel.r1 || sel.c0 !== sel.c1) return { r1: sel.r0, c1: sel.c0, r2: sel.r1, c2: sel.c1 };
    return currentRegion({ row: sel.r0, col: sel.c0 }, (r, c) => displayAt(r, c).trim() !== '', { rows: sheet?.rowCount ?? 1, cols: sheet?.colCount ?? 1 });
  }
  /** 첫 행이 머리글로 보이는지(Excel처럼): 모든 칸이 차 있고 글자가 섞여 있으며, 아래 행과 모양이 다르다 */
  function looksLikeHeader(region: CellRangeRef): boolean {
    if (region.r2 <= region.r1) return false;
    // 엑셀과 같은 판단: 첫 행이 모두 채워진 글자이고, 바로 아래 행과 자료 종류가 다르거나(글자 위 숫자) 첫 행만 굵으면 머리글.
    // (예전엔 마지막 조건이 늘 참이라 글자가 든 첫 행은 무조건 머리글로 잡혀 정렬에서 빠졌다)
    const isNum = (v: string) => v !== '' && Number.isFinite(Number(v.replace(/,/g, '')));
    let allFilled = true;
    let allText = true;
    let typeDiffers = false;
    for (let c = region.c1; c <= region.c2; c += 1) {
      const head = displayAt(region.r1, c).trim();
      const below = displayAt(region.r1 + 1, c).trim();
      if (!head) allFilled = false;
      else if (isNum(head)) allText = false;
      else if (isNum(below)) typeDiffers = true;
    }
    const boldAt = (r: number) => Boolean(styles.get(gridCellKey(activeSheet, r, region.c1))?.b);
    const boldHead = boldAt(region.r1) && !boldAt(region.r1 + 1);
    return allFilled && allText && (typeDiffers || boldHead);
  }

  /** 범위의 행을 keyCol 기준으로 다시 쓴다 — 값·서식이 행째 움직이고, 수식의 상대 행 참조도 따라 옮긴다. */
  function sortRegion(region: CellRangeRef, keyCol: number, descending: boolean, hasHeader: boolean): void {
    if (!canEdit) return;
    const first = region.r1 + (hasHeader ? 1 : 0);
    if (first >= region.r2) return;
    if (!rowsLoaded(region.r1, region.r2)) {
      void officeAlert(`정렬할 범위가 너무 큽니다(${region.r2 - region.r1 + 1}행). 화면 근처 ${PAGE_SIZE_ROWS}행 단위 범위 안에서만 정렬할 수 있습니다.`);
      return;
    }
    const keys: string[] = [];
    for (let r = first; r <= region.r2; r += 1) keys.push(displayAt(r, keyCol));
    const order = sortOrder(keys, descending);
    if (order.every((from, i) => from === i)) return;
    const snapshot = new Map<string, { value: string; style: CellStyleAttrs | undefined }>();
    for (let r = first; r <= region.r2; r += 1) {
      for (let c = region.c1; c <= region.c2; c += 1) {
        const key = gridCellKey(activeSheet, r, c);
        snapshot.set(key, { value: cells.get(key) ?? '', style: styles.get(key) });
      }
    }
    ydoc.transact(() => {
      order.forEach((fromOffset, toOffset) => {
        const fromRow = first + fromOffset;
        const toRow = first + toOffset;
        for (let c = region.c1; c <= region.c2; c += 1) {
          const source = snapshot.get(gridCellKey(activeSheet, fromRow, c));
          const key = gridCellKey(activeSheet, toRow, c);
          const value = source ? shiftFormulaRows(source.value, toRow - fromRow) : '';
          if (value === '') cells.delete(key);
          else cells.set(key, value);
          styles.set(key, source?.style ?? {});
        }
      });
    });
    renderTableVirtualized();
  }

  function sortActiveColumn(descending: boolean): void {
    const region = dataRegion();
    if (!region) {
      void officeAlert('정렬할 데이터를 선택하세요.');
      return;
    }
    const keyCol = Math.min(Math.max(selAnchorCol, region.c1), region.c2);
    const filter = filtersMap.get(String(activeSheet));
    const inFilter = filter && region.r1 >= filter.r1 && region.r2 <= filter.r2;
    sortRegion(inFilter ? filter : region, keyCol, descending, inFilter ? true : looksLikeHeader(region));
  }
  /** 열 머리 메뉴 — 시트 전체를 그 열로(도착하지 않은 행이 있으면 sortRegion 이 막는다) */
  function sortWholeSheetBy(col: number, descending: boolean): void {
    const sheet = structure.get(String(activeSheet));
    if (!sheet) return;
    const region = { r1: 0, c1: 0, r2: sheet.rowCount - 1, c2: sheet.colCount - 1 };
    sortRegion(region, col, descending, looksLikeHeader(region));
  }
  document.getElementById('btn-sort-asc')?.addEventListener('click', () => sortActiveColumn(false));
  document.getElementById('btn-sort-desc')?.addEventListener('click', () => sortActiveColumn(true));

  // 사용자 지정 정렬 대화상자
  const sortDialog = document.getElementById('sort-dialog') as HTMLElement | null;
  const sortColumnSelect = document.getElementById('sort-dialog-column') as HTMLSelectElement | null;
  const sortOrderSelect = document.getElementById('sort-dialog-order') as HTMLSelectElement | null;
  const sortHeaderCheck = document.getElementById('sort-dialog-header') as HTMLInputElement | null;
  let sortDialogRegion: CellRangeRef | null = null;
  function fillSortColumns(): void {
    if (!sortDialogRegion || !sortColumnSelect || !sortHeaderCheck) return;
    const keep = sortColumnSelect.value;
    sortColumnSelect.replaceChildren();
    for (let c = sortDialogRegion.c1; c <= sortDialogRegion.c2; c += 1) {
      const option = document.createElement('option');
      option.value = String(c);
      const head = displayAt(sortDialogRegion.r1, c);
      option.textContent = sortHeaderCheck.checked && head ? head : `${colLetter(c)}열`;
      sortColumnSelect.append(option);
    }
    if (keep && [...sortColumnSelect.options].some((o) => o.value === keep)) sortColumnSelect.value = keep;
  }
  document.getElementById('btn-sort-custom')?.addEventListener('click', () => {
    sortDialogRegion = dataRegion();
    if (!sortDialogRegion || !sortDialog || !sortHeaderCheck || !sortColumnSelect) {
      void officeAlert('정렬할 데이터를 선택하세요.');
      return;
    }
    sortHeaderCheck.checked = looksLikeHeader(sortDialogRegion);
    fillSortColumns();
    sortColumnSelect.value = String(Math.min(Math.max(selAnchorCol, sortDialogRegion.c1), sortDialogRegion.c2));
    (document.getElementById('sort-dialog-range') as HTMLElement).textContent =
      `범위 ${colLetter(sortDialogRegion.c1)}${sortDialogRegion.r1 + 1}:${colLetter(sortDialogRegion.c2)}${sortDialogRegion.r2 + 1}`;
    sortDialog.hidden = false;
    sortColumnSelect.focus();
  });
  sortHeaderCheck?.addEventListener('change', fillSortColumns);
  const closeSortDialog = () => {
    if (sortDialog) sortDialog.hidden = true;
  };
  document.getElementById('sort-dialog-cancel')?.addEventListener('click', closeSortDialog);
  sortDialog?.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') closeSortDialog();
  });
  document.getElementById('sort-dialog-ok')?.addEventListener('click', () => {
    closeSortDialog();
    if (sortDialogRegion && sortColumnSelect && sortOrderSelect && sortHeaderCheck) {
      sortRegion(sortDialogRegion, Number(sortColumnSelect.value), sortOrderSelect.value === 'desc', sortHeaderCheck.checked);
    }
  });

  // ---------- 필터: 머리글 ▾ 단추, 값 목록으로 거르기, 숨긴 행 ----------
  const filterLayer = document.createElement('div');
  filterLayer.className = 'grid-filter-layer';
  scrollHostEl.append(filterLayer);
  const filterBtn = document.getElementById('btn-filter');
  let hiddenForSheet = -1;

  function recomputeHiddenRows(): void {
    const filter = filtersMap.get(String(activeSheet));
    hiddenForSheet = activeSheet;
    // 아직 도착하지 않은 행은 값을 모르니 숨기지 않는다(도착하면 pagesArrivedHook이 다시 계산한다)
    const next = filter ? filteredRows(filter, displayAt).filter(rowArrived) : [];
    const changed = next.length !== hiddenRowsSorted.length || next.some((r, i) => r !== hiddenRowsSorted[i]);
    if (changed) {
      hiddenRowsSorted = next;
      hiddenRowSet = new Set(next);
      renderTableVirtualized();
    }
    renderFilterButtons();
  }

  function renderFilterButtons(): void {
    const filter = filtersMap.get(String(activeSheet));
    filterBtn?.setAttribute('aria-pressed', String(Boolean(filter)));
    filterLayer.replaceChildren();
    if (!filter) return;
    filterLayer.style.left = `${tableEl.offsetLeft}px`;
    filterLayer.style.top = `${tableEl.offsetTop}px`;
    const top = headerHeight() + rowLayout.top(filter.r1);
    for (let c = filter.c1; c <= filter.c2; c += 1) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'grid-filter-btn';
      button.classList.toggle('is-active', Boolean(filter.criteria[String(c)]));
      button.dataset.col = String(c);
      button.title = filter.criteria[String(c)] ? '필터 적용됨 — 눌러서 바꾸기' : '정렬 및 필터';
      button.setAttribute('aria-label', `${colLetter(c)}열 필터`);
      button.innerHTML = filter.criteria[String(c)]
        ? '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M4 5h16l-6 7.5V19l-4-2v-4.5z"/></svg>'
        : '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>';
      button.style.left = `${colLeft(c + 1) - 20}px`;
      button.style.top = `${top + (rowLayout.height(filter.r1) - 18) / 2}px`;
      button.addEventListener('click', (event) => {
        event.stopPropagation();
        openFilterMenu(c, button);
      });
      filterLayer.append(button);
    }
  }

  const filterMenu = document.createElement('div');
  filterMenu.className = 'dropdown-menu dropdown-menu-floating filter-menu';
  filterMenu.hidden = true;
  document.body.append(filterMenu);
  document.addEventListener('mousedown', (event) => {
    if (!filterMenu.hidden && !filterMenu.contains(event.target as Node) && !(event.target as HTMLElement).closest('.grid-filter-btn')) filterMenu.hidden = true;
  });

  /** 필터 목록이 열려 있는 열 — 데이터가 도착하면 그 목록을 다시 그린다 */
  let openFilterCol: { col: number; anchor: HTMLElement } | null = null;

  function openFilterMenu(col: number, anchor: HTMLElement): void {
    const filter = filtersMap.get(String(activeSheet));
    if (!filter) return;
    openFilterCol = { col, anchor };
    // 필터 범위가 한 연결이 받을 수 있는 페이지 수보다 크면(약 16,000행 초과) 전부 받을 수 없다 — 받아 둔 부분만으로
    // 목록을 만들고 그렇다고 알린다
    const filterPageCount = Math.floor(filter.r2 / PAGE_SIZE_ROWS) - Math.floor(filter.r1 / PAGE_SIZE_ROWS) + 1;
    const partial = filterPageCount > MAX_SUBSCRIBED_PAGES;
    const ready = partial
      ? [...lastSubscribedPages].every((page) => arrivedPages.get(activeSheet)?.has(page))
      : rowsLoaded(filter.r1, filter.r2);
    if (!ready) {
      // 필터 범위의 데이터가 아직 다 오지 않았다 — 일부 값만으로 목록을 만들면 값이 빠져 보인다(실측: 목록이 비어 보임)
      filterMenu.replaceChildren();
      const waiting = document.createElement('div');
      waiting.className = 'filter-loading';
      waiting.innerHTML = '<span class="grid-spinner" aria-hidden="true"></span>값 목록을 불러오는 중…';
      filterMenu.append(waiting);
      positionFilterMenu(anchor);
      scheduleSubscriptionUpdate();
      return;
    }
    const values = new Set<string>();
    // 도착한 행만 — 안 온 행을 빈 칸으로 세면 목록에 없는 "(필드 값 없음)"이 끼었다
    for (let r = filter.r1 + 1; r <= filter.r2; r += 1) if (rowArrived(r)) values.add(displayAt(r, col));
    const sorted = [...values].sort((a, b) => compareCells(a, b, false));
    const allowed = filter.criteria[String(col)];
    const checked = new Set(allowed ?? sorted);
    filterMenu.replaceChildren();
    const action = (label: string, run: () => void) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.textContent = label;
      b.addEventListener('click', () => {
        filterMenu.hidden = true;
        run();
      });
      filterMenu.append(b);
    };
    action('오름차순 정렬', () => sortRegion(filter, col, false, true));
    action('내림차순 정렬', () => sortRegion(filter, col, true, true));
    if (allowed) {
      action(`"${displayAt(filter.r1, col) || colLetter(col)}"에서 필터 해제`, () => {
        const criteria = { ...filter.criteria };
        delete criteria[String(col)];
        filtersMap.set(String(activeSheet), { ...filter, criteria });
      });
    }
    if (partial) {
      const note = document.createElement('div');
      note.className = 'filter-note';
      note.textContent = `행이 많아 일부(약 ${(MAX_SUBSCRIBED_PAGES * PAGE_SIZE_ROWS).toLocaleString()}행)만 보고 만든 목록입니다`;
      filterMenu.append(note);
    }
    const list = document.createElement('div');
    list.className = 'filter-values';
    const all = document.createElement('label');
    const allBox = document.createElement('input');
    allBox.type = 'checkbox';
    allBox.checked = checked.size === sorted.length;
    all.append(allBox, ' (모두 선택)');
    list.append(all);
    const boxes: HTMLInputElement[] = [];
    for (const value of sorted) {
      const label = document.createElement('label');
      const box = document.createElement('input');
      box.type = 'checkbox';
      box.checked = checked.has(value);
      box.dataset.value = value;
      boxes.push(box);
      label.append(box, ` ${value === '' ? '(필드 값 없음)' : value}`);
      list.append(label);
      box.addEventListener('change', () => {
        allBox.checked = boxes.every((b) => b.checked);
      });
    }
    allBox.addEventListener('change', () => {
      for (const box of boxes) box.checked = allBox.checked;
    });
    filterMenu.append(list);
    const actions = document.createElement('div');
    actions.className = 'filter-actions';
    const ok = document.createElement('button');
    ok.type = 'button';
    ok.className = 'btn-mini filter-ok';
    ok.textContent = '확인';
    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.className = 'btn-mini';
    cancel.textContent = '취소';
    actions.append(ok, cancel);
    filterMenu.append(actions);
    cancel.addEventListener('click', () => {
      filterMenu.hidden = true;
    });
    ok.addEventListener('click', () => {
      filterMenu.hidden = true;
      const picked = boxes.filter((b) => b.checked).map((b) => b.dataset.value ?? '');
      const criteria = { ...filter.criteria };
      if (picked.length === sorted.length) delete criteria[String(col)];
      else criteria[String(col)] = picked;
      filtersMap.set(String(activeSheet), { ...filter, criteria });
    });
    positionFilterMenu(anchor);
  }

  function positionFilterMenu(anchor: HTMLElement): void {
    filterMenu.hidden = false;
    const rect = anchor.getBoundingClientRect();
    filterMenu.style.left = `${Math.max(4, Math.min(rect.left, window.innerWidth - filterMenu.offsetWidth - 4))}px`;
    filterMenu.style.top = `${rect.bottom + 2}px`;
  }

  filterBtn?.addEventListener('click', () => {
    if (!canEdit) return;
    const key = String(activeSheet);
    if (filtersMap.has(key)) {
      filtersMap.delete(key);
      return;
    }
    const region = dataRegion();
    if (!region || region.r2 <= region.r1) {
      void officeAlert('필터를 걸 데이터 범위(머리글 + 데이터)를 선택하세요.');
      return;
    }
    filtersMap.set(key, { r1: region.r1, c1: region.c1, r2: region.r2, c2: region.c2, criteria: {} });
  });
  document.getElementById('btn-filter-clear')?.addEventListener('click', () => {
    const filter = filtersMap.get(String(activeSheet));
    if (filter) filtersMap.set(String(activeSheet), { ...filter, criteria: {} });
  });

  filtersMap.observe(recomputeHiddenRows);
  let filterRecomputeScheduled = false;
  const scheduleFilterRecompute = () => {
    if (filterRecomputeScheduled) return;
    filterRecomputeScheduled = true;
    requestAnimationFrame(() => {
      filterRecomputeScheduled = false;
      recomputeHiddenRows();
    });
  };
  cells.observe(scheduleFilterRecompute);
  computed.observe(scheduleFilterRecompute);
  const previousAfterGridRender = afterGridRender;
  afterGridRender = () => {
    previousAfterGridRender();
    // 시트를 바꾸면 그 시트의 필터로 숨길 행을 다시 계산한다
    if (hiddenForSheet !== activeSheet) scheduleFilterRecompute();
    renderFilterButtons();
  };
  recomputeHiddenRows();

  // ---------- 확대/축소 (엑셀처럼 드래그 슬라이더 포함) ----------
  const zoomLevelEl = document.getElementById('zoom-level') as HTMLElement;
  const zoomSliderTrackEl = document.getElementById('zoom-slider-track') as HTMLElement | null;
  const zoomSliderFillEl = document.getElementById('zoom-slider-fill') as HTMLElement | null;
  const zoomSliderHandleEl = document.getElementById('zoom-slider-handle') as HTMLElement | null;
  const ZOOM_MIN = 0.6;
  const ZOOM_MAX = 2;
  let zoom = 1;
  // 확대율이 낮아질수록 화면에 보이는 실제 영역이 넓어지므로, 그 영역을 채우기에 지금 그려둔
  // 열 수(가상 스크롤 대상이 아님)와 가상 스크롤 범위(뷰포트 높이 기준)가 여전히 맞는지 다시
  // 확인한다(rAF로 묶어서 슬라이더를 드래그하는 동안 매 프레임 전체 재렌더가 도는 걸 막는다).
  // renderTableVirtualized()는 이제 화면에 보이는 범위만 그리므로(시트 전체 크기와 무관하게
  // 항상 가볍다), 예전처럼 "이미 충분히 그려져 있는지" 미리 비교할 필요 없이 그냥 다시 부른다.
  let viewportRefitScheduled = false;
  function scheduleViewportRefit(): void {
    if (viewportRefitScheduled) return;
    viewportRefitScheduled = true;
    requestAnimationFrame(() => {
      viewportRefitScheduled = false;
      renderTableVirtualized();
    });
  }
  function applyZoom(): void {
    zoomHost.style.zoom = String(zoom);
    zoomLevelEl.textContent = `${Math.round(zoom * 100)}%`;
    const ratio = (zoom - ZOOM_MIN) / (ZOOM_MAX - ZOOM_MIN);
    if (zoomSliderFillEl) zoomSliderFillEl.style.width = `${ratio * 100}%`;
    if (zoomSliderHandleEl) zoomSliderHandleEl.style.left = `${ratio * 100}%`;
    scheduleViewportRefit();
  }
  window.addEventListener('resize', scheduleViewportRefit);

  // ---------- 열 폭·행 높이 바꾸기 — 머리글 경계를 끌거나(엑셀처럼) 두 번 눌러 맞춘다 ----------
  // 표는 스크롤·구조 변경마다 다시 그려지므로 칸마다 붙이지 않고 표에 한 번만 위임한다. 값은 서버가 적용해
  // 구조 맵으로 모두에게 알리고(applyGridResizeOp), 저장할 때 바꾼 열·행만 파일에 쓴다.
  const RESIZE_EDGE_PX = 5;
  const MIN_COL_PX = 8;
  const MIN_ROW_PX = 6;
  function sendGridResizeOp(kind: 'resizeCols' | 'resizeRows', at: number, count: number, size: number): void {
    if (!canEdit || !provider.ws || provider.ws.readyState !== WebSocket.OPEN) return;
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_GRID_STRUCTURE_OP);
    encoding.writeVarString(encoder, JSON.stringify({ sheet: activeSheet, kind, at, count, size: Math.round(size) }));
    provider.ws.send(encoding.toUint8Array(encoder));
  }
  type ResizeTarget = { axis: 'col' | 'row'; index: number; edge: HTMLElement };
  /** 커서가 머리글 경계 위면 어느 열(그 경계 왼쪽)·행(경계 위쪽)을 바꿀지 */
  function resizeTargetAt(event: MouseEvent): ResizeTarget | null {
    if (!canEdit) return null;
    const el = (event.target as HTMLElement | null)?.closest<HTMLElement>('thead th[data-col], tbody .g-edit-rownum[data-row]');
    if (!el || !tableEl.contains(el)) return null;
    const rect = el.getBoundingClientRect();
    if (el.dataset.col !== undefined) {
      const col = Number(el.dataset.col);
      if (rect.right - event.clientX <= RESIZE_EDGE_PX) return { axis: 'col', index: col, edge: el };
      if (event.clientX - rect.left <= RESIZE_EDGE_PX && col > 0) return { axis: 'col', index: col - 1, edge: el };
      return null;
    }
    const row = Number(el.dataset.row);
    if (rect.bottom - event.clientY <= RESIZE_EDGE_PX) return { axis: 'row', index: row, edge: el };
    if (event.clientY - rect.top <= RESIZE_EDGE_PX && row > 0) return { axis: 'row', index: row - 1, edge: el };
    return null;
  }
  function colWidthPx(c: number): number {
    const s = currentStructure();
    return s.colWidths?.[c] ?? s.defaultColWidth ?? DEFAULT_COL_WIDTH_PX;
  }
  function rowHeightPx(r: number): number {
    return rowLayout.top(r + 1) - rowLayout.top(r);
  }
  /** 열(행) 전체를 여러 개 골라 둔 채 그중 하나의 경계를 바꾸면 고른 것 모두 같은 크기로(엑셀과 같다) */
  function resizeSpan(axis: 'col' | 'row', index: number): { at: number; count: number } {
    const s = currentStructure();
    const r1 = Math.min(selAnchorRow, selFocusRow);
    const r2 = Math.max(selAnchorRow, selFocusRow);
    const c1 = Math.min(selAnchorCol, selFocusCol);
    const c2 = Math.max(selAnchorCol, selFocusCol);
    if (axis === 'col' && r1 === 0 && r2 >= s.rowCount - 1 && index >= c1 && index <= c2) return { at: c1, count: c2 - c1 + 1 };
    if (axis === 'row' && c1 === 0 && c2 >= s.colCount - 1 && index >= r1 && index <= r2) return { at: r1, count: r2 - r1 + 1 };
    return { at: index, count: 1 };
  }
  let resizeDrag: { target: ResizeTarget; start: number; startSize: number; size: number; guide: HTMLElement; tip: HTMLElement } | null = null;
  let suppressHeaderClick = false;
  tableEl.addEventListener('mousemove', (event) => {
    if (resizeDrag) return;
    const target = resizeTargetAt(event);
    tableEl.classList.toggle('g-resize-col', target?.axis === 'col');
    tableEl.classList.toggle('g-resize-row', target?.axis === 'row');
  });
  tableEl.addEventListener('mouseleave', () => {
    if (!resizeDrag) tableEl.classList.remove('g-resize-col', 'g-resize-row');
  });
  tableEl.addEventListener('mousedown', (event) => {
    if (event.button !== 0 || event.detail > 1) return;
    const target = resizeTargetAt(event);
    if (!target) return;
    event.preventDefault();
    event.stopPropagation();
    const isCol = target.axis === 'col';
    const startSize = isCol ? colWidthPx(target.index) : rowHeightPx(target.index);
    const host = scrollHostEl.getBoundingClientRect();
    const guide = document.createElement('div');
    guide.className = `g-resize-guide ${isCol ? 'is-col' : 'is-row'}`;
    const tip = document.createElement('div');
    tip.className = 'g-resize-tip';
    document.body.append(guide, tip);
    resizeDrag = { target, start: isCol ? event.clientX : event.clientY, startSize, size: startSize, guide, tip };
    const place = (clientX: number, clientY: number) => {
      if (!resizeDrag) return;
      const delta = ((isCol ? clientX : clientY) - resizeDrag.start) / (zoom || 1);
      resizeDrag.size = Math.max(isCol ? MIN_COL_PX : MIN_ROW_PX, Math.round(startSize + delta));
      const at = resizeDrag.start + (resizeDrag.size - startSize) * (zoom || 1);
      if (isCol) Object.assign(guide.style, { left: `${at}px`, top: `${host.top}px`, height: `${host.height}px` });
      else Object.assign(guide.style, { top: `${at}px`, left: `${host.left}px`, width: `${host.width}px` });
      tip.textContent = isCol ? `너비: ${resizeDrag.size}픽셀` : `높이: ${resizeDrag.size}픽셀`;
      Object.assign(tip.style, { left: `${clientX + 12}px`, top: `${clientY - 28}px` });
    };
    place(event.clientX, event.clientY);
    const onMove = (e: MouseEvent) => place(e.clientX, e.clientY);
    const onUp = () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp, true);
      const drag = resizeDrag;
      resizeDrag = null;
      guide.remove();
      tip.remove();
      tableEl.classList.remove('g-resize-col', 'g-resize-row');
      // 끈 뒤 따라오는 click이 열/행 선택으로 바뀌지 않게
      suppressHeaderClick = true;
      setTimeout(() => (suppressHeaderClick = false), 0);
      if (!drag || drag.size === drag.startSize) return;
      const { at, count } = resizeSpan(drag.target.axis, drag.target.index);
      sendGridResizeOp(drag.target.axis === 'col' ? 'resizeCols' : 'resizeRows', at, count, drag.size);
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp, true);
  }, true);
  tableEl.addEventListener('click', (event) => {
    if (!suppressHeaderClick) return;
    event.preventDefault();
    event.stopPropagation();
  }, true);
  // 경계를 두 번 누르면 맞춤 — 열은 지금 그려진 칸 중 가장 긴 내용에, 행은 기본 높이로
  const measureCanvas = document.createElement('canvas').getContext('2d');
  tableEl.addEventListener('dblclick', (event) => {
    const target = resizeTargetAt(event);
    if (!target) return;
    event.preventDefault();
    event.stopPropagation();
    const { at, count } = resizeSpan(target.axis, target.index);
    if (target.axis === 'row') {
      sendGridResizeOp('resizeRows', at, count, currentStructure().defaultRowHeight ?? rowHeightPx(target.index));
      return;
    }
    for (let c = at; c < at + count; c += 1) {
      let widest = 0;
      for (const tr of vwRowEls.values()) {
        const td = tr.querySelector<HTMLTableCellElement>(`td[data-col="${c}"]`);
        const text = td?.textContent ?? '';
        if (!td || text.trim() === '' || td.colSpan > 1 || !measureCanvas) continue;
        const cs = getComputedStyle(td);
        measureCanvas.font = cs.font;
        const lines = text.split('\n');
        const textWidth = Math.max(...lines.map((line) => measureCanvas.measureText(line).width));
        widest = Math.max(widest, Math.ceil(textWidth + parseFloat(cs.paddingLeft) + parseFloat(cs.paddingRight) + 2));
      }
      if (widest > 0) sendGridResizeOp('resizeCols', c, 1, Math.min(widest, 2000));
    }
  }, true);
  function setZoomFromRatio(ratio: number): void {
    const clamped = Math.max(0, Math.min(1, ratio));
    zoom = Math.round((ZOOM_MIN + clamped * (ZOOM_MAX - ZOOM_MIN)) * 20) / 20; // 5%p 단위로 스냅
    applyZoom();
  }
  // 보기 탭의 확대/축소 — 상태 표시줄 컨트롤과 같은 값을 움직인다.
  document.getElementById('btn-zoom-100')?.addEventListener('click', () => {
    zoom = 1;
    applyZoom();
  });
  document.getElementById('btn-zoom-in-ribbon')?.addEventListener('click', () => document.getElementById('zoom-in')?.click());
  document.getElementById('btn-zoom-out-ribbon')?.addEventListener('click', () => document.getElementById('zoom-out')?.click());
  document.getElementById('zoom-in')?.addEventListener('click', () => {
    zoom = Math.min(ZOOM_MAX, Math.round((zoom + 0.1) * 10) / 10);
    applyZoom();
  });
  document.getElementById('zoom-out')?.addEventListener('click', () => {
    zoom = Math.max(ZOOM_MIN, Math.round((zoom - 0.1) * 10) / 10);
    applyZoom();
  });
  function ratioFromClientX(clientX: number): number {
    const rect = zoomSliderTrackEl!.getBoundingClientRect();
    return (clientX - rect.left) / rect.width;
  }
  zoomSliderTrackEl?.addEventListener('mousedown', (event) => {
    if (event.target === zoomSliderHandleEl) return; // 핸들 자체의 드래그는 아래 리스너가 처리
    setZoomFromRatio(ratioFromClientX(event.clientX));
  });
  zoomSliderHandleEl?.addEventListener('mousedown', (event) => {
    event.preventDefault();
    const onMove = (moveEvent: MouseEvent) => setZoomFromRatio(ratioFromClientX(moveEvent.clientX));
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
  applyZoom();

  // ---------- 엑셀처럼 시트 탭 옆에 붙는 커스텀 가로 스크롤바 ----------
  // 네이티브 가로 스크롤바는 CSS에서 숨기고(그리드 바로 아래에 뜨는 게 엑셀과 달라서), 이 위젯이
  // scrollHostEl.scrollLeft를 그대로 읽고 쓴다 — 장식이 아니라 실제 스크롤을 대신한다.
  const hscrollEl = document.querySelector('.grid-hscroll') as HTMLElement | null;
  const hscrollTrackEl = document.getElementById('hscroll-track') as HTMLElement | null;
  const hscrollThumbEl = document.getElementById('hscroll-thumb') as HTMLElement | null;
  function updateHScrollbar(): void {
    if (!hscrollEl || !hscrollTrackEl || !hscrollThumbEl) return;
    const { scrollWidth, clientWidth, scrollLeft } = scrollHostEl;
    const scrollable = scrollWidth > clientWidth + 1;
    hscrollEl.dataset.scrollable = String(scrollable);
    if (!scrollable) {
      hscrollThumbEl.style.width = '100%';
      hscrollThumbEl.style.left = '0';
      return;
    }
    const trackWidth = hscrollTrackEl.clientWidth;
    const thumbWidth = Math.max(16, (clientWidth / scrollWidth) * trackWidth);
    const maxThumbLeft = trackWidth - thumbWidth;
    const maxScrollLeft = scrollWidth - clientWidth;
    const thumbLeft = maxScrollLeft > 0 ? (scrollLeft / maxScrollLeft) * maxThumbLeft : 0;
    hscrollThumbEl.style.width = `${thumbWidth}px`;
    hscrollThumbEl.style.left = `${thumbLeft}px`;
  }
  scrollHostEl.addEventListener('scroll', updateHScrollbar);
  window.addEventListener('resize', updateHScrollbar);

  // ---------- 가상 스크롤 재렌더 + 페이지 구독 전송(윈도잉) ----------
  // 서버는 이제 클라이언트가 MESSAGE_PAGE_SUBSCRIBE로 알려준 500행 단위 "페이지"만 Yjs에 채운다
  // (room.ts 참고) — 스크롤로 보이는 범위가 바뀔 때마다 그 범위를 포함하는 페이지 번호를 계산해
  // 서버에 알려야, 스크롤해서 새로 보이는 행에도 실제 데이터가 채워진다.
  const MESSAGE_PAGE_SUBSCRIBE = 2;
  let subscriptionDebounceTimer: ReturnType<typeof setTimeout> | null = null;

  /** room.ts가 받는 그대로의 형식(varUint 메시지 타입 + varString JSON 페이로드)으로 같은 WebSocket
   * 연결에 얹어 보낸다. y-websocket의 WebsocketProvider는 이 커스텀 메시지 타입을 몰라 자기
   * messageHandlers로 처리하지 못하지만, 그건 수신 쪽 얘기고(서버는 room.ts에서 messageType으로
   * 직접 분기해 처리한다) — 전송은 provider가 이미 들고 있는 실제 WebSocket 인스턴스
   * (node_modules/y-websocket/src/y-websocket.js 확인: `provider.ws`가 `WebSocket | null`이고,
   * 라이브러리 자신도 내부적으로 `this.ws.send(encoding.toUint8Array(encoder))`로 syncStep1을
   * 보낸다 — 같은 패턴을 그대로 따른다)를 통해 그대로 내보내면 된다. */
  function sendPageSubscription(sheetIdx: number, pages: readonly number[]): void {
    if (!provider.ws || provider.ws.readyState !== WebSocket.OPEN) return;
    const encoder = encoding.createEncoder();
    encoding.writeVarUint(encoder, MESSAGE_PAGE_SUBSCRIBE);
    encoding.writeVarString(encoder, JSON.stringify({ sheet: sheetIdx, pages }));
    provider.ws.send(encoding.toUint8Array(encoder));
  }

  /** 지금 스크롤 위치(+여유분) 기준으로 필요한 페이지 번호를 계산해 서버에 알린다. 스크롤 중
   * 매 프레임 보낼 필요는 없으므로(서버가 그때마다 페이지를 채우고 비우는 건 낭비다) 120ms마다
   * 한 번씩만 보낸다 — renderTableVirtualized()가 호출되는 모든 경로(스크롤, 시트 전환, 구조 변경,
   * 최초 동기화)에서 공통으로 이 함수가 불린다. */
  // 디바운스가 아니라 스로틀이다 — 스크롤이 이어지는 동안 타이머를 계속 미루면 멈출 때까지 요청이 한 번도
  // 안 나가서, 계속 내리는 내내 빈 칸('불러오는 중')만 보였다. 이미 예약돼 있으면 그대로 두고, 그때의 위치로 보낸다.
  function scheduleSubscriptionUpdate(): void {
    if (subscriptionDebounceTimer) return;
    subscriptionDebounceTimer = setTimeout(() => {
      subscriptionDebounceTimer = null;
      const scrollTop = scrollHostEl.scrollTop;
      const viewportHeight = scrollHostEl.clientHeight;
      const firstRow = Math.max(0, rowLayout.rowAt(scrollTop) - OVERSCAN_ROWS);
      const lastRow = rowLayout.rowAt(scrollTop + viewportHeight) + OVERSCAN_ROWS;
      const firstPage = Math.floor(firstRow / PAGE_SIZE_ROWS);
      const lastPage = Math.floor(lastRow / PAGE_SIZE_ROWS);
      const pages = new Set<number>();
      for (let p = firstPage; p <= lastPage; p += 1) pages.add(p);
      // 틀 고정된 위쪽 행은 스크롤 위치와 무관하게 항상 화면에 그려지므로(renderTableVirtualized),
      // 그 데이터가 담긴 페이지(freezeRows는 항상 시트 맨 위부터라 언제나 0번 페이지)도 항상 같이
      // 구독해 둔다 — 안 그러면 멀리 스크롤했을 때 고정된 헤더 행이 빈 칸으로 보일 수 있다.
      const freezeRows = structure.get(String(activeSheet))?.freezeRows ?? 0;
      if (freezeRows > 0) pages.add(0);
      visiblePages = new Set(pages);
      // 필터가 걸린 시트는 필터 범위 전체를 함께 받아 둔다 — 값 목록·숨길 행은 범위 전체 데이터가 있어야 맞다.
      // 서버가 한 연결에 받아 주는 페이지 수(32)를 넘지 않게, 보이는 페이지를 먼저 두고 남는 만큼만.
      const activeFilter = filtersMap.get(String(activeSheet));
      if (activeFilter) {
        for (let p = Math.floor(activeFilter.r1 / PAGE_SIZE_ROWS); p <= Math.floor(activeFilter.r2 / PAGE_SIZE_ROWS) && pages.size < MAX_SUBSCRIBED_PAGES; p += 1) pages.add(p);
      }
      lastSubscribedPages = pages;
      sendPageSubscription(activeSheet, [...pages].sort((a, b) => a - b));
      updateLoadingIndicator();
    }, 120);
  }

  // 세로 스크롤 때마다 화면에 보이는 범위만 다시 그린다(rAF로 묶어 스크롤 중 매 프레임 다시
  // 그리는 걸 막는다) — 기존 커스텀 가로 스크롤바 리스너(updateHScrollbar)는 그대로 두고 이
  // 리스너는 추가로 붙는다(대체 아님). 편집 중(contentEditable)에는 재렌더가 그 <td>를 DOM에서
  // 떼어내며 강제로 blur를 발생시켜 방금 타이핑한 값을 잃을 수 있어(위 블러 핸들러의 세대 판정이
  // 그런 blur를 "커밋하지 않고 버림" 처리한다) 재렌더 자체를 건너뛴다 — 스크롤은 계속되고, 편집이
  // 끝나면(다음 스크롤이나 다른 재렌더 트리거에서) 화면이 다시 맞춰진다.
  //
  // renderTableVirtualized() 전체 재빌드 대신 updateScrollWindow()를 부른다 — 실측(Chrome
  // longtask) 기준 전체 재빌드는 스크롤 프레임 하나에 54~108ms(16ms 프레임 예산의 3~6배)가 걸려
  // 스크롤이 버벅이며 의도한 만큼 이동하지 않는 원인이었다. updateScrollWindow는 같은 문맥(시트/
  // 열 폭/틀고정/병합)에서 스크롤만 일어난 경우 <colgroup>/<thead>/창 안에서 안 바뀐 행은 그대로
  // 두고 창 가장자리에서 나가고 들어온 행만 patch한다 — 문맥이 조금이라도 의심되면 알아서
  // renderTableVirtualized()로 물러나므로 안전하다.
  let scrollRenderScheduled = false;
  // ---------- 로딩 표시 — 화면에 보이는 페이지의 데이터가 아직 오지 않았으면 스피너 ----------
  const loadingEl = document.createElement('div');
  loadingEl.className = 'grid-loading';
  loadingEl.hidden = true;
  loadingEl.setAttribute('role', 'status');
  loadingEl.innerHTML = '<span class="grid-spinner" aria-hidden="true"></span><span>데이터를 불러오는 중…</span>';
  scrollHostEl.parentElement?.append(loadingEl);
  let loadingTimer: ReturnType<typeof setTimeout> | null = null;
  function updateLoadingIndicator(): void {
    const arrived = arrivedPages.get(activeSheet);
    const waiting = [...visiblePages].some((page) => !arrived?.has(page)) || visiblePages.size === 0;
    if (!waiting) {
      if (loadingTimer) clearTimeout(loadingTimer);
      loadingTimer = null;
      loadingEl.hidden = true;
      setGridStatusText('준비');
      return;
    }
    // 잠깐이면(스크롤 중 바로 도착) 깜빡이지 않게 조금 기다렸다 띄운다
    if (!loadingTimer && loadingEl.hidden) {
      loadingTimer = setTimeout(() => {
        loadingTimer = null;
        const stillWaiting = [...visiblePages].some((page) => !arrivedPages.get(activeSheet)?.has(page)) || visiblePages.size === 0;
        if (stillWaiting) {
          loadingEl.hidden = false;
          setGridStatusText('불러오는 중');
        }
      }, 150);
    }
  }
  pagesArrivedHook = () => {
    updateLoadingIndicator();
    if (filtersMap.get(String(activeSheet))) recomputeHiddenRows();
    // "불러오는 중"을 보이고 있을 때만 다시 그린다 — 값 목록이 이미 떠 있으면 사용자가 고르던 체크가 날아간다
    if (openFilterCol && !filterMenu.hidden && filterMenu.querySelector('.filter-loading')) openFilterMenu(openFilterCol.col, openFilterCol.anchor);
  };
  updateLoadingIndicator();

  scrollHostEl.addEventListener('scroll', () => {
    scheduleSubscriptionUpdate();
    if (editingKey !== null) return;
    if (scrollRenderScheduled) return;
    scrollRenderScheduled = true;
    requestAnimationFrame(() => {
      scrollRenderScheduled = false;
      if (editingKey !== null) return;
      updateScrollWindow();
    });
  });
  document.getElementById('hscroll-left')?.addEventListener('click', () => {
    scrollHostEl.scrollBy({ left: -100 });
  });
  document.getElementById('hscroll-right')?.addEventListener('click', () => {
    scrollHostEl.scrollBy({ left: 100 });
  });
  hscrollThumbEl?.addEventListener('mousedown', (event) => {
    event.preventDefault();
    const startX = event.clientX;
    const startScrollLeft = scrollHostEl.scrollLeft;
    const trackWidth = hscrollTrackEl!.clientWidth;
    const thumbWidth = hscrollThumbEl.getBoundingClientRect().width;
    const maxScrollLeft = scrollHostEl.scrollWidth - scrollHostEl.clientWidth;
    const onMove = (moveEvent: MouseEvent) => {
      const deltaPx = moveEvent.clientX - startX;
      const deltaScroll = (deltaPx / (trackWidth - thumbWidth)) * maxScrollLeft;
      scrollHostEl.scrollLeft = Math.max(0, Math.min(maxScrollLeft, startScrollLeft + deltaScroll));
    };
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });
  // 트랙의 빈 부분을 클릭하면 그 방향으로 한 화면씩 스크롤한다(엑셀과 동일).
  hscrollTrackEl?.addEventListener('mousedown', (event) => {
    if (event.target !== hscrollTrackEl) return;
    const trackRect = hscrollTrackEl.getBoundingClientRect();
    const clickedBeforeThumb = event.clientX < trackRect.left + (hscrollThumbEl?.offsetLeft ?? 0);
    scrollHostEl.scrollBy({ left: clickedBeforeThumb ? -scrollHostEl.clientWidth : scrollHostEl.clientWidth });
  });

  // ---------- 문서(시트 전체) 내 검색 ----------
  interface CellMatch {
    readonly sheetIdx: number;
    readonly row: number;
    readonly col: number;
  }
  const searchBar = document.getElementById('search-bar') as HTMLElement;
  const searchInput = document.getElementById('search-input') as HTMLInputElement;
  const searchCount = document.getElementById('search-count') as HTMLElement;
  const searchNextBtn = document.getElementById('search-next') as HTMLButtonElement;
  const searchScopeHintEl = document.getElementById('search-scope-hint') as HTMLElement | null;
  const searchReplaceRow = document.getElementById('search-replace-row') as HTMLElement | null;
  const searchReplaceInput = document.getElementById('search-replace-input') as HTMLInputElement | null;
  let searchMatches: CellMatch[] = [];
  let searchCurrent = -1;
  let lastHighlighted: HTMLTableCellElement | null = null;
  let allHighlighted: HTMLTableCellElement[] = [];
  // 엑셀처럼 — 다중 셀 범위를 선택한 채로 찾기/바꾸기를 열면 그 범위 안에서만 찾는다(셀 하나만
  // 선택돼 있었으면 시트 전체). 대화상자를 여는 시점의 선택을 스냅샷으로 고정해 둔다.
  let searchScope: { sheetIdx: number; r0: number; r1: number; c0: number; c1: number } | null = null;

  function findCellMatches(query: string): CellMatch[] {
    if (!query) return [];
    const lower = query.toLowerCase();
    const matches: CellMatch[] = [];
    for (const [key, value] of cells) {
      if (!value.toLowerCase().includes(lower)) continue;
      const parsed = parseGridCellKey(key);
      if (!parsed) continue;
      if (
        searchScope &&
        (parsed.sheetIdx !== searchScope.sheetIdx ||
          parsed.row < searchScope.r0 ||
          parsed.row > searchScope.r1 ||
          parsed.col < searchScope.c0 ||
          parsed.col > searchScope.c1)
      ) {
        continue;
      }
      matches.push(parsed);
    }
    matches.sort((a, b) => a.sheetIdx - b.sheetIdx || a.row - b.row || a.col - b.col);
    return matches;
  }

  function clearHighlight(): void {
    lastHighlighted?.classList.remove('g-hit-current');
    lastHighlighted = null;
  }
  function clearAllHighlight(): void {
    allHighlighted.forEach((td) => td.classList.remove('g-hit-all'));
    allHighlighted = [];
  }

  function jumpToMatch(index: number): void {
    clearHighlight();
    const match = searchMatches[index];
    if (!match) return;
    if (match.sheetIdx !== activeSheet) {
      activeSheet = match.sheetIdx;
      // 검색 결과로 다른 시트로 건너뛰는 것도 시트 전환이다 — 이전 시트에서 잡아 둔 선택 범위를
      // 그대로 들고 가면 안 된다(탭 클릭 전환과 같은 이유, resetSelectionForSheetSwitch 주석 참고).
      resetSelectionForSheetSwitch();
      renderSheetTabs();
      renderTableVirtualized();
    }
    // 가상 스크롤 범위 밖(다른 페이지)의 매치일 수 있다 — 필요하면 스크롤해 그 행을 그린 뒤 찾는다.
    const td = tdForEnsureRendered(match.row, match.col);
    if (td) {
      td.classList.add('g-hit-current');
      td.scrollIntoView({ block: 'center', inline: 'center' });
      lastHighlighted = td;
    }
  }

  function applySearch(jumpTo: number): void {
    clearAllHighlight();
    searchMatches = findCellMatches(searchInput.value);
    searchCurrent =
      searchMatches.length > 0 ? ((jumpTo % searchMatches.length) + searchMatches.length) % searchMatches.length : -1;
    searchCount.textContent = searchMatches.length > 0 ? `${searchCurrent + 1}/${searchMatches.length}` : searchInput.value ? '0/0' : '';
    searchNextBtn.disabled = searchMatches.length === 0;
    if (searchCurrent >= 0) jumpToMatch(searchCurrent);
    else clearHighlight();
  }
  /** "모두 찾기" — 엑셀의 결과 목록 대신, 매치되는 칸을 전부(연한 노랑) 칠하고 첫 매치로 이동해
   * (진한 노랑) 개수를 보여준다 — 지금 화면(활성 시트)에 실제로 그려진 칸만 칠할 수 있다. */
  function findAllMatches(): void {
    clearAllHighlight();
    searchMatches = findCellMatches(searchInput.value);
    searchCurrent = searchMatches.length > 0 ? 0 : -1;
    searchCount.textContent = searchMatches.length > 0 ? `${searchMatches.length}개 찾음` : searchInput.value ? '0개 찾음' : '';
    searchNextBtn.disabled = searchMatches.length === 0;
    if (searchMatches.length === 0) {
      clearHighlight();
      return;
    }
    searchMatches.forEach((m) => {
      if (m.sheetIdx !== activeSheet) return;
      const td = tdFor(m.row, m.col);
      if (td) {
        td.classList.add('g-hit-all');
        allHighlighted.push(td);
      }
    });
    jumpToMatch(0);
  }
  // ---------- 찾기/바꾸기 창 드래그 이동(엑셀처럼 제목줄을 잡고 옮길 수 있게) ----------
  const searchTitlebar = searchBar.querySelector('.search-titlebar') as HTMLElement | null;
  if (searchTitlebar) {
    let dragStart: { pointerX: number; pointerY: number; left: number; top: number } | null = null;

    searchTitlebar.addEventListener('mousedown', (event) => {
      event.preventDefault();
      const rect = searchBar.getBoundingClientRect();
      // 중앙 정렬(left:50%+transform)로 열려 있던 상태라면, 드래그 시작 시점의 실제 화면 위치를
      // 고정된 left/top 픽셀값으로 바꿔둔다 — 그래야 이후 이동량을 그대로 더해도 어긋나지 않는다.
      searchBar.style.left = `${rect.left}px`;
      searchBar.style.top = `${rect.top}px`;
      searchBar.style.transform = 'none';
      dragStart = { pointerX: event.clientX, pointerY: event.clientY, left: rect.left, top: rect.top };
    });

    document.addEventListener('mousemove', (event) => {
      if (!dragStart) return;
      const rect = searchBar.getBoundingClientRect();
      const maxLeft = Math.max(window.innerWidth - rect.width, 0);
      const maxTop = Math.max(window.innerHeight - rect.height, 0);
      const nextLeft = dragStart.left + (event.clientX - dragStart.pointerX);
      const nextTop = dragStart.top + (event.clientY - dragStart.pointerY);
      searchBar.style.left = `${Math.min(Math.max(nextLeft, 0), maxLeft)}px`;
      searchBar.style.top = `${Math.min(Math.max(nextTop, 0), maxTop)}px`;
    });

    document.addEventListener('mouseup', () => {
      dragStart = null;
    });
  }

  function openSearch(withReplace = false): void {
    searchBar.hidden = false;
    const range = rangeBounds();
    const hasRealSelection = range.r0 !== range.r1 || range.c0 !== range.c1;
    searchScope = hasRealSelection ? { sheetIdx: activeSheet, ...range } : null;
    if (searchScopeHintEl) searchScopeHintEl.hidden = !hasRealSelection;
    if (withReplace && searchReplaceRow) searchReplaceRow.hidden = false;
    searchInput.focus();
    searchInput.select();
  }
  function closeSearch(): void {
    searchBar.hidden = true;
    searchInput.value = '';
    clearHighlight();
    clearAllHighlight();
    searchMatches = [];
    searchCurrent = -1;
    searchCount.textContent = '';
    searchScope = null;
    if (searchScopeHintEl) searchScopeHintEl.hidden = true;
    if (searchReplaceRow) searchReplaceRow.hidden = true;
  }
  document.getElementById('search-btn')?.addEventListener('click', () => {
    if (searchBar.hidden) openSearch();
    else closeSearch();
  });
  document.getElementById('btn-find-replace')?.addEventListener('click', () => {
    if (!canEdit) return;
    if (searchBar.hidden) openSearch(true);
    else closeSearch();
  });
  document.getElementById('search-close')?.addEventListener('click', closeSearch);
  document.getElementById('search-find-all')?.addEventListener('click', () => findAllMatches());
  searchInput.addEventListener('input', () => applySearch(0));
  searchInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') applySearch(searchCurrent + (event.shiftKey ? -1 : 1));
    if (event.key === 'Escape') closeSearch();
  });
  searchNextBtn.addEventListener('click', () => applySearch(searchCurrent + 1));

  // ---------- 찾아서 바꾸기 ----------
  function escapeForRegExp(text: string): string {
    return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  function replaceInText(text: string, find: string, replace: string): string {
    if (!find) return text;
    return text.replace(new RegExp(escapeForRegExp(find), 'gi'), replace);
  }
  document.getElementById('search-replace-one')?.addEventListener('click', () => {
    if (!canEdit) return;
    const match = searchMatches[searchCurrent];
    if (!match) return;
    const key = gridCellKey(match.sheetIdx, match.row, match.col);
    if (styles.get(key)?.lk === 1) {
      applySearch(searchCurrent + 1);
      return;
    }
    const next = replaceInText(cells.get(key) ?? '', searchInput.value, searchReplaceInput?.value ?? '');
    ydoc.transact(() => {
      if (next === '') cells.delete(key);
      else cells.set(key, next);
    });
    applySearch(searchCurrent); // 방금 바꾼 칸이 더 이상 검색어와 안 맞으면 자연히 다음 매치로 넘어간다
  });
  document.getElementById('search-replace-all')?.addEventListener('click', () => {
    if (!canEdit || !searchInput.value) return;
    const findText = searchInput.value;
    const replaceText = searchReplaceInput?.value ?? '';
    const matches = findCellMatches(findText);
    ydoc.transact(() => {
      matches.forEach((m) => {
        const key = gridCellKey(m.sheetIdx, m.row, m.col);
        if (styles.get(key)?.lk === 1) return;
        const current = cells.get(key) ?? '';
        const next = replaceInText(current, findText, replaceText);
        if (next === current) return;
        if (next === '') cells.delete(key);
        else cells.set(key, next);
      });
    });
    applySearch(0);
  });

  window.addEventListener('keydown', (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f') {
      event.preventDefault();
      openSearch();
      return;
    }
    // 엑셀과 동일하게 Ctrl+H는 "찾기 및 바꾸기"를 바꾸기 탭이 펼쳐진 채로 연다.
    if (canEdit && (event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'h') {
      event.preventDefault();
      openSearch(true);
      return;
    }
    // Ctrl/Cmd+Z(실행 취소), Ctrl/Cmd+Y(다시 실행) — 셀 편집 중(입력창에 포커스)이 아닐 때만
    // 그리드 단위 undo/redo로 처리한다. 편집 중이면 브라우저의 기본 텍스트 입력 undo가 자연스럽다.
    const target = event.target as HTMLElement | null;
    const isTextInput = target?.tagName === 'INPUT' || target?.tagName === 'TEXTAREA' || target?.isContentEditable;
    if (isTextInput) return;
    if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === 'z') {
      event.preventDefault();
      undoManager.undo();
    } else if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'y') {
      event.preventDefault();
      undoManager.redo();
    }
  });

  // ---------- 수식 입력줄: 선택된 셀 값을 그대로 편집할 수 있는 별도 입력창 ----------
  formulaInputEl.addEventListener('input', () => {
    if (!activeCellKey) return;
    const parsed = parseGridCellKey(activeCellKey);
    const td = parsed ? tdFor(parsed.row, parsed.col) : null;
    if (td) td.textContent = formulaInputEl.value;
  });
  function commitFormulaBar(): void {
    if (!activeCellKey) return;
    const key = activeCellKey;
    const parsed = parseGridCellKey(key);
    const td = parsed ? tdFor(parsed.row, parsed.col) : null;
    if (!td) return;
    commitCell(td);
    renderCellDisplay(td, key);
    recalcFormulaCells();
  }
  formulaInputEl.addEventListener('change', commitFormulaBar);
  formulaInputEl.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      commitFormulaBar();
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
