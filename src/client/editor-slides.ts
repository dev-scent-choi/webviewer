import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import * as encoding from 'lib0/encoding';
import { EditorState, type Command } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { baseKeymap, splitBlockAs, toggleMark } from 'prosemirror-commands';
import { keymap } from 'prosemirror-keymap';
import type { Mark, MarkType, Node as PMNode } from 'prosemirror-model';
import { ySyncPlugin, yCursorPlugin, yUndoPlugin, ySyncPluginKey, undo, redo } from 'y-prosemirror';
import {
  PPTX_SLIDES_ARRAY,
  PPTX_SLIDE_CONTENTS_MAP,
  PPTX_SLIDE_VISUALS_MAP,
  NEW_TEXTBOX_WIDTH_RATIO,
  NEW_TEXTBOX_HEIGHT_RATIO,
  SLIDE_WIDTH_EMU,
  SLIDE_HEIGHT_EMU,
  shouldBlockAddTextbox,
  NEW_SHAPE_GEOMS,
  type NewShapeGeom,
  type SlideEntry,
  type SlideShapeMeta,
  type SlideVisual,
} from '../collab/pptxShared.js';
import { DOC_META_MAP, DOC_META_SAVED_AT_KEY, DOC_META_SAVE_ERROR_AT_KEY } from '../collab/gridShared.js';
import { pickPresenceColorForName, renderPresenceAvatars } from './presence.js';
import { pptxTextSchema } from '../pptx/textSchema.js';

/** src/pptx/collab.ts가 서버 쪽에서 선언한 것과 같은 모양이지만, 그 파일은 node:crypto/fflate 등
 * 서버 전용 의존성을 끌고 있어 클라이언트 번들에 그대로 import할 수 없다 — Yjs 공유 타입 자체는
 * 구조적 타이핑이라 여기서 같은 모양을 독립적으로 다시 선언해도 완전히 호환된다. */
type ShapeContentMap = Y.Map<unknown>;
type PerSlideContentsMap = Y.Map<ShapeContentMap>;

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
  readonly slideSize?: { readonly cx: number; readonly cy: number } | null;
}

const token = new URLSearchParams(location.search).get('token');
const titleEl = document.getElementById('doc-title')!;
const connStatusEl = document.getElementById('conn-status')!;
const railEl = document.getElementById('slides-rail') as HTMLElement;
const canvasEl = document.getElementById('slide-canvas') as HTMLElement;
const cardTemplate = document.getElementById('slide-card-template') as HTMLTemplateElement;
const presenceListEl = document.getElementById('presence-list') as HTMLElement;

// 어떤 슬라이드가 오른쪽 캔버스에 표시 중인지 — renderCanvasForSlide()가 이 값을 읽어 실제
// 콘텐츠를 그린다.
let activeSlideId: string | null = null;
// Task 8이 남겨 둔 훅 포인트 — main()에서 renderCanvasForSlide로 연결한다(아래 참고).
let onActiveSlideChange: ((entryId: string) => void) | null = null;

function showFatalError(message: string): void {
  document.getElementById('error-message')!.textContent = message;
  document.getElementById('error-screen')!.hidden = false;
}

async function fetchViewInfo(): Promise<ViewInfo> {
  const res = await fetch(`/api/view/info?token=${encodeURIComponent(token ?? '')}`);
  const json = (await res.json()) as { success: boolean; data?: ViewInfo; error?: string };
  if (!res.ok || !json.success || !json.data) throw new Error(json.error ?? '문서 정보를 불러오지 못했습니다.');
  return json.data;
}

let canEdit = false;
let slidesArray: Y.Array<SlideEntry>;
let dragFromIndex: number | null = null;
// main()에서 대입된다(slidesArray와 같은 패턴) — subscribeSlideContent/renderCanvasForSlide가
// 모듈 스코프 함수라 provider/slideContents도 모듈 스코프에 있어야 접근할 수 있다.
let provider: WebsocketProvider;
let slideContents: Y.Map<PerSlideContentsMap>;
/** 이미지·단순 도형·표/차트/스마트아트(전부 표시 전용, 편집 불가) — slideContents와 별개 top-level
 * 맵. room.ts가 slideContents와 같은 지연 시딩 트리거(MESSAGE_SLIDE_CONTENT_SUBSCRIBE)로 채우므로
 * 클라이언트가 따로 구독을 보낼 필요는 없다(아래 observe로 도착을 감지해 다시 그리기만 한다). */
let slideVisuals: Y.Map<readonly SlideVisual[]>;

/**
 * 슬라이드 크기(EMU)를 캔버스 픽셀로 스케일한다. `/api/view/info`(ViewInfo)에는 슬라이드 크기
 * 필드가 없고, room.ts/viewer.ts 어디에도 sldSz류 값을 따로 내려주는 경로가 없다(grep으로 확인
 * 완료) — 표준 16:9 슬라이드 크기(pptxShared.ts의 SLIDE_WIDTH_EMU/SLIDE_HEIGHT_EMU, 서버의
 * 슬라이드 배경 표시용 치수와 같은 값)를 기본값으로 쓴다. 실제 프레젠테이션의 진짜 크기를
 * 반영하는 것은 이후 개선 과제로 남긴다(이번 태스크에서 서버 API를 넓히지 않는다).
 */
// main()이 /api/view/info의 slideSize(presentation.xml의 실제 크기)로 덮어쓴다 — 없으면 16:9.
let SLIDE_W_EMU = SLIDE_WIDTH_EMU;
let SLIDE_H_EMU = SLIDE_HEIGHT_EMU;
const CANVAS_WIDTH_PX = 960; // 캔버스 폭 기준(반응형은 이후 과제)
const EMU_PER_PT = 12700;

function emuToPx(emu: number): number {
  return emu * (CANVAS_WIDTH_PX / SLIDE_W_EMU);
}

/** 캔버스 위 1pt의 px — 글자 크기·여백을 슬라이드와 같은 축척으로 그리기 위해 CSS 변수로 넘긴다. */
function slidePtPx(): number {
  return CANVAS_WIDTH_PX / (SLIDE_W_EMU / EMU_PER_PT);
}

/** room.ts의 MESSAGE_SLIDE_CONTENT_SUBSCRIBE와 **같은 값이어야 한다**(grid-editor.ts의
 * MESSAGE_GRID_STRUCTURE_OP 주석과 같은 이유 — 0~3은 y-websocket 표준 프로토콜이 이미 쓰고
 * 있어 새 값은 그 범위를 벗어나게 둔다). */
const MESSAGE_SLIDE_CONTENT_SUBSCRIBE = 5;

/**
 * 슬라이드 하나를 "지금 연다"고 서버에 알린다 — 서버가 slideContents.get(entryId)를 아직 안
 * 채웠으면 원본 XML에서 지연 시딩한다(room.ts의 seedSlideContentIfNeeded). 전송 방식은
 * grid-editor.ts의 sendPageSubscription과 완전히 같은, 이미 검증된 패턴이다: WebsocketProvider가
 * 공개 필드로 들고 있는 실제 WebSocket 인스턴스(node_modules/y-websocket/dist/src/y-websocket.d.ts:
 * `ws: WebSocket | null`)에 varUint 메시지 타입 + varString JSON 페이로드를 그대로 실어 보낸다 —
 * y-websocket 자신도 내부적으로 `this.ws.send(encoding.toUint8Array(encoder))`로 syncStep1을
 * 보내는 것과 같은 경로다.
 */
// 서버에 아직 시딩 응답을 못 받은(=slideContents에 그 키가 아직 없는) 슬라이드 id 집합 —
// 같은 슬라이드 카드를 빠르게 두 번 클릭해도 두 번째 요청은 보내지 않기 위한 것이다(서버는
// 요청마다 디스크 읽기+unzip+XML 파싱을 새로 하므로, 첫 요청이 아직 응답하기 전의 중복
// 요청은 그 비용을 그대로 낭비하고 두 번째는 결국 조용히 no-op된다). slideContents.observe가
// 실제로 그 슬라이드 키가 채워진 것을 감지하면(main() 참고) 이 집합에서 제거한다.
//
// 다만 서버 쪽 시딩이 어떤 이유로든(에러 throw, duplicateOfId를 실은 읽기 전용 연결의 요청 거부,
// 연결당 동시 시딩 상한 초과, 잘못된 payload, 존재하지 않는 슬라이드 엔트리 등) 실패하면
// slideContents에는 아무것도 쓰이지 않으므로 observe가 영영 안
// 불리고, 이 집합에 남은 entryId는 세션이 끝날 때까지 제거되지 않는다 — 그러면 그 슬라이드는
// 이후 몇 번을 다시 클릭해도 이 함수 맨 위 가드에서 조용히 no-op되어 영구히 빈 캔버스로 남는다.
// 이를 막기 위해 아래 subscribeSlideContent에서 항목을 추가할 때마다 유한한 수명(타임아웃)을
// 함께 건다 — 정상적으로 시딩되면 observe가 먼저 지우고, 타임아웃이 나중에 발동해도
// Set.delete는 이미 없는 값에 대해 안전한 no-op이라 아무 문제도 없다. 시딩이 실패한 경우엔
// 타임아웃이 먼저 발동해 이 집합에서 빠지고, 사용자가 그 슬라이드를 다시 클릭하면 재시도된다.
const pendingSlideSubscriptions = new Set<string>();
// pendingSlideSubscriptions 항목 하나의 최대 수명(ms) — 정확할 필요는 없고, 정상적인 왕복
// 시간보다 넉넉히 길어서 성공 케이스에서는 절대 발동하지 않을 정도면 된다.
const PENDING_SUBSCRIPTION_TIMEOUT_MS = 9000;

// 복제(duplicate)로 만들어진 슬라이드 entryId -> 그 원본(source) 슬라이드 entryId. 복제 버튼
// 핸들러가 선제적으로 duplicateOfId를 실어 구독 요청을 한 번 보내지만(아래 setupCardActions
// 참고), 그 시점에 WebSocket이 OPEN이 아니면(재연결 중 등) subscribeSlideContent는 아무것도
// 보내지 않고 조용히 리턴한다 — 그 경우 그 복제 슬라이드에 대해 duplicateOfId를 실어 보내는
// 유일한 시도가 사라져 버리고, 이후 renderCanvasForSlide의 "아직 안 열렸다" 폴백 경로가 보통의
// (duplicateOfId 없는) 구독만 보내게 되어 서버가 소스의 살아있는 편집 내용 대신 오래된 디스크
// XML로 시딩해 버린다. 이 맵은 그 간극을 구조적으로 막는다 — subscribeSlideContent가
// duplicateOfId 없이 불릴 때마다 여기서 먼저 찾아본다(아래 참고). 실제로 시딩이 끝나면(=
// slideContents.observe가 그 키를 감지하면) 더 이상 필요 없으니 지운다.
const pendingDuplicateSources = new Map<string, string>();

function subscribeSlideContent(entryId: string, duplicateOfId?: string): void {
  if (!provider.ws || provider.ws.readyState !== WebSocket.OPEN) return;
  if (pendingSlideSubscriptions.has(entryId)) return; // 이미 이 슬라이드에 대한 요청이 응답을 기다리는 중
  // 이 호출이 duplicateOfId를 명시적으로 안 받았다면(예: renderCanvasForSlide의 폴백 경로,
  // 또는 위 타임아웃으로 풀려난 뒤의 재시도) 이 슬라이드가 복제본으로 알려져 있는지 먼저
  // 확인한다 — 있다면 그 원본 id를 그대로 duplicateOfId로 써서, 어느 경로로 결국 요청이
  // 성공하든 항상 올바른 소스에서 시딩되게 한다.
  const effectiveDuplicateOfId = duplicateOfId ?? pendingDuplicateSources.get(entryId);
  pendingSlideSubscriptions.add(entryId);
  setTimeout(() => pendingSlideSubscriptions.delete(entryId), PENDING_SUBSCRIPTION_TIMEOUT_MS);
  const encoder = encoding.createEncoder();
  encoding.writeVarUint(encoder, MESSAGE_SLIDE_CONTENT_SUBSCRIBE);
  encoding.writeVarString(encoder, JSON.stringify({ slideEntryId: entryId, duplicateOfId: effectiveDuplicateOfId }));
  provider.ws.send(encoding.toUint8Array(encoder));
}

// entryId(슬라이드) -> 그 슬라이드의 shapeId -> 마운트된 EditorView. 슬라이드를 벗어나면(다른
// 슬라이드를 클릭하면) 여기 담긴 뷰를 전부 destroy()한 뒤 비운다 — 그러지 않으면 이미 화면에서
// 사라진 슬라이드의 EditorView가 계속 ySyncPlugin으로 Y.Doc을 관찰하며 메모리/리스너가 샌다.
const mountedViews = new Map<string, EditorView>();

// 지금 포커스를 들고 있는(=가장 최근에 focus 이벤트를 받은) EditorView. 서식 툴바(굵게/기울임/
// 밑줄/색)가 "어느 텍스트박스에 적용할지" 알아야 하는데, 슬라이드 하나에 텍스트박스(=EditorView)가
// 여러 개 동시에 마운트돼 있어 이 값 없이는 알 수 없다 — 각 view 생성 시 handleDOMEvents.focus로
// 갱신한다(아래 renderCanvasForSlide 참고). 툴바 버튼을 클릭하면 그 클릭이 현재 포커스된
// contentEditable을 blur시키지만, ProseMirror의 state.selection은 DOM blur와 무관하게 그대로
// 남아 있으므로(editor.ts의 bindToolbarButton과 동일한, 이미 검증된 가정) blur 이벤트에서 이 값을
// null로 되돌리지 않는다 — 그러면 버튼 클릭 시점엔 이미 focusedView가 null이 되어 아무 데도 적용이
// 안 된다. 유일한 예외는 renderCanvasForSlide가 이 view를 destroy할 때(아래에서 별도로 null 처리).
let focusedView: EditorView | null = null;

// 지금 활성 슬라이드의 perSlide(Y.Map<shapeId, shapeMap>)에 붙여 둔 구조적 변경 구독 해제 함수.
// perSlide 자기 자신에 대한 얕은(shallow) observe는 shapeId 키가 추가/삭제될 때만(=텍스트박스
// 추가/삭제) 올라오고, 텍스트박스 "안" 내용 편집(타이핑)은 그보다 훨씬 깊은 Y.XmlFragment
// 레벨에서 일어나 여기로 전혀 올라오지 않는다 — 그래서 이 구독을 renderCanvasForSlide 재호출
// 트리거로 써도 키 입력마다 EditorView가 destroy/재생성되는 일은 없다. slideContents 자체의
// 최상위 observe(main() 참고)는 "슬라이드가 아직 시딩 전이냐"만 보므로, 이미 시딩된 슬라이드
// 안에서 텍스트박스가 추가/삭제되는 것(이 태스크의 핵심 기능)은 그 최상위 observe로는 잡히지
// 않는다 — 그래서 이 별도의 perSlide 레벨 구독이 필요하다. renderCanvasForSlide가 불릴 때마다
// (슬라이드 전환이든, 이 구독 자신이 재호출한 것이든) 맨 위에서 이전 구독을 먼저 해제하고 새로
// 붙이므로, 활성 슬라이드가 아닌 슬라이드에 대한 구독이 남아있는 일은 없다.
let activePerSlideUnobserve: (() => void) | null = null;

/**
 * y-prosemirror의 undo/redo를 실제 키(Ctrl+Z/Ctrl+Shift+Z, Ctrl+Y)에 연결한다. 이 keymap
 * 없이는 yUndoPlugin()이 만드는 Y.UndoManager가 순전히 죽은 코드다 — 어떤 키도 그 undo/redo
 * 커맨드를 부르지 않으니 Ctrl+Z를 누르면 브라우저 자신의 contentEditable 네이티브 undo가
 * 대신 실행되고, 그 결과 되돌아간 DOM을 ProseMirror가 "새 입력"으로 다시 읽어들여 모든
 * 협업자에게 진짜 편집처럼 전파해 버린다 — 그 과정에서 다른 협업자가 동시에 넣은, 전혀 무관한
 * 편집을 조용히 되돌릴 수 있다. editor.ts(docx 편집기)의 buildKeymap과 같은 패턴 — baseKeymap
 * 위에 Mod-z/Mod-y/Mod-Shift-z만 y-prosemirror의 undo/redo로 덮어쓴다(이 스키마엔 표/목록/
 * 제목이 없어 그 외 docx 전용 바인딩은 옮기지 않는다).
 */
/** 리본 동작을 키보드에서도 같은 키로 — 커맨드는 항상 "이 view에서 처리됨(true)"을 돌려
 * 브라우저 기본 동작(Ctrl+R 새로고침, Ctrl+E 주소창 등)이 끼어들지 않게 한다. */
function viewCommand(run: (view: EditorView) => void): Command {
  return (_state, _dispatch, view) => {
    if (view) run(view);
    return true;
  };
}

function buildSlideKeymap(): Record<string, Command> {
  return {
    ...baseKeymap,
    'Mod-z': undo,
    'Mod-y': redo,
    'Mod-Shift-z': redo,
    'Mod-b': toggleMark(pptxTextSchema.marks.strong),
    'Mod-i': toggleMark(pptxTextSchema.marks.em),
    'Mod-u': toggleMark(pptxTextSchema.marks.underline),
    'Mod-l': viewCommand((v) => applyAlign(v, 'l')),
    'Mod-e': viewCommand((v) => applyAlign(v, 'ctr')),
    'Mod-r': viewCommand((v) => applyAlign(v, 'r')),
    'Mod-j': viewCommand((v) => applyAlign(v, 'just')),
    'Mod->': viewCommand((v) => stepFontSize(v, 1)),
    'Mod-Shift-.': viewCommand((v) => stepFontSize(v, 1)),
    'Mod-<': viewCommand((v) => stepFontSize(v, -1)),
    'Mod-Shift-,': viewCommand((v) => stepFontSize(v, -1)),
    // PowerPoint처럼 Enter로 만든 새 문단은 지금 문단의 수준·글머리·정렬(pPr)을 이어받는다 — 기본
    // splitBlock은 문단 끝에서 Enter를 치면 속성 없는 새 문단을 만들어, 하위 글머리 목록에서 다음
    // 항목이 최상위 수준으로 튀어나왔다.
    Enter: splitBlockAs((node) => ({ type: node.type, attrs: node.attrs })),
  };
}

/**
 * (Task 9 당시 도입 — 2026-09-17 읽기 전용 시딩 허용 수정으로 현재는 호출부가 없다) 읽기 전용
 * 열람자가 "아직 아무도 열어본 적 없는" 슬라이드를 볼 때 캔버스에 보여주던 안내. 그 시점엔
 * canEdit 세션만 subscribeSlideContent로 서버 지연 시딩을 요청할 수 있었고, 읽기 전용 연결은
 * room.ts가 그 요청 자체를 서버에서 거부해 응답이 영영 오지 않았다 — 지금은 room.ts의
 * MESSAGE_SLIDE_CONTENT_SUBSCRIBE 핸들러가 읽기 전용 연결의 순수 시딩 요청도 그대로 처리하므로
 * (duplicateOfId를 실은 요청만 거부한다), renderCanvasForSlide는 canEdit 여부와 무관하게 항상
 * subscribeSlideContent를 부르고 이 함수는 더 이상 쓰이지 않는다. 시딩 자체가 실패하는(연결당
 * 동시 시딩 상한 초과 등) 드문 경우는 canEdit 세션이 겪는 것과 같은 무설명 빈 캔버스로 수렴하므로
 * 새 회귀는 아니다. 당장 도로 쓸 데는 없지만 앞으로 "시딩이 아예 막힌 경우"를 다시 구분해 알려야
 * 할 상황이 생기면 그대로 재사용할 수 있어 남겨 둔다. canvasEl.replaceChildren()이 이미 이전
 * 내용을 비운 뒤 호출된다는 전제(호출부가 생기면 지켜야 할 조건)로 여기서는 추가만 한다.
 */
function showSlideUnavailablePlaceholder(): void {
  const placeholder = document.createElement('div');
  placeholder.className = 'slide-canvas-placeholder';
  placeholder.textContent = '이 슬라이드는 아직 미리보기를 사용할 수 없습니다 (편집자가 먼저 열어야 합니다).';
  canvasEl.append(placeholder);
}

/** PowerPoint 기본 도형(prstGeom) 중 자주 쓰이는 다각형 모양 — CSS clip-path로 그 모양만 남긴다.
 * 조정 핸들(avLst)은 기본값 비율로 근사한다. 여기 없는 모양은 사각형으로 표시된다. */
const PRESET_CLIP: Record<string, string> = {
  triangle: 'polygon(50% 0, 100% 100%, 0 100%)',
  rtTriangle: 'polygon(0 0, 100% 100%, 0 100%)',
  diamond: 'polygon(50% 0, 100% 50%, 50% 100%, 0 50%)',
  flowChartDecision: 'polygon(50% 0, 100% 50%, 50% 100%, 0 50%)',
  parallelogram: 'polygon(25% 0, 100% 0, 75% 100%, 0 100%)',
  flowChartInputOutput: 'polygon(20% 0, 100% 0, 80% 100%, 0 100%)',
  trapezoid: 'polygon(25% 0, 75% 0, 100% 100%, 0 100%)',
  pentagon: 'polygon(50% 0, 100% 38%, 82% 100%, 18% 100%, 0 38%)',
  hexagon: 'polygon(25% 0, 75% 0, 100% 50%, 75% 100%, 25% 100%, 0 50%)',
  octagon: 'polygon(29% 0, 71% 0, 100% 29%, 100% 71%, 71% 100%, 29% 100%, 0 71%, 0 29%)',
  homePlate: 'polygon(0 0, 80% 0, 100% 50%, 80% 100%, 0 100%)',
  chevron: 'polygon(0 0, 80% 0, 100% 50%, 80% 100%, 0 100%, 20% 50%)',
  rightArrow: 'polygon(0 25%, 60% 25%, 60% 0, 100% 50%, 60% 100%, 60% 75%, 0 75%)',
  leftArrow: 'polygon(100% 25%, 40% 25%, 40% 0, 0 50%, 40% 100%, 40% 75%, 100% 75%)',
  upArrow: 'polygon(25% 100%, 25% 40%, 0 40%, 50% 0, 100% 40%, 75% 40%, 75% 100%)',
  downArrow: 'polygon(25% 0, 75% 0, 75% 60%, 100% 60%, 50% 100%, 0 60%, 25% 60%)',
  leftRightArrow: 'polygon(0 50%, 20% 0, 20% 25%, 80% 25%, 80% 0, 100% 50%, 80% 100%, 80% 75%, 20% 75%, 20% 100%)',
  star5: 'polygon(50% 0, 61% 35%, 98% 35%, 68% 57%, 79% 91%, 50% 70%, 21% 91%, 32% 57%, 2% 35%, 39% 35%)',
  plus: 'polygon(33% 0, 67% 0, 67% 33%, 100% 33%, 100% 67%, 67% 67%, 67% 100%, 33% 100%, 33% 67%, 0 67%, 0 33%, 33% 33%)',
  mathPlus: 'polygon(40% 0, 60% 0, 60% 40%, 100% 40%, 100% 60%, 60% 60%, 60% 100%, 40% 100%, 40% 60%, 0 60%, 0 40%, 40% 40%)',
  snip1Rect: 'polygon(0 0, 85% 0, 100% 15%, 100% 100%, 0 100%)',
  snip2SameRect: 'polygon(15% 0, 85% 0, 100% 15%, 100% 100%, 0 100%, 0 15%)',
};
/** 둥근 계열 별칭 — CSS 클래스(slide-shape-ellipse/roundRect/stadium)로 그린다. */
const ROUNDED_ALIAS: Record<string, string> = {
  flowChartConnector: 'ellipse',
  flowChartAlternateProcess: 'roundRect',
  round2SameRect: 'roundRect',
  flowChartTerminator: 'stadium',
};

/** 이미지/단순 도형/표/차트·스마트아트 자리표시자 하나를 표시 전용 엘리먼트로 그린다 —
 * 텍스트박스(.slide-textbox)와 달리 클릭해도 편집기가 뜨지 않는다(pointer-events는 CSS에서
 * .slide-visual에 none으로 꺼둔다). */
function renderSlideVisual(visual: SlideVisual): HTMLElement {
  const el = document.createElement(visual.kind === 'table' ? 'table' : 'div');
  el.className = `slide-visual slide-visual-${visual.kind}`;
  el.style.left = `${emuToPx(visual.x ?? 0)}px`;
  el.style.top = `${emuToPx(visual.y ?? 0)}px`;
  el.style.width = `${emuToPx(visual.w ?? 0)}px`;
  el.style.height = `${emuToPx(visual.h ?? 0)}px`;

  if (visual.kind === 'image' && visual.mediaPath) {
    const img = document.createElement('img');
    img.src = `/api/view/slide-media?token=${encodeURIComponent(token ?? '')}&path=${encodeURIComponent(visual.mediaPath)}`;
    img.alt = '';
    el.append(img);
  } else if (visual.kind === 'shape') {
    if (visual.fill) el.style.background = visual.fill;
    if (visual.geom) {
      const clip = PRESET_CLIP[visual.geom];
      if (clip) el.style.clipPath = clip;
      else el.classList.add(`slide-shape-${ROUNDED_ALIAS[visual.geom] ?? visual.geom}`);
    }
  } else if (visual.kind === 'table' && visual.rows) {
    for (const row of visual.rows) {
      const tr = document.createElement('tr');
      for (const cellText of row) {
        const td = document.createElement('td');
        td.textContent = cellText;
        tr.append(td);
      }
      el.append(tr);
    }
  } else if (visual.kind === 'placeholder') {
    el.textContent = visual.label ?? '콘텐츠';
  }
  return el;
}

/**
 * 오른쪽 캔버스에 activeSlideId 슬라이드의 텍스트박스를 실제 위치로 배치하고, 각 텍스트박스에
 * pptxTextSchema 기반 ProseMirror EditorView를 마운트한다(y-prosemirror 바인딩은 editor.ts의
 * 기존 docx 편집기 마운트 코드와 동일한 패턴 — ySyncPlugin/yCursorPlugin/yUndoPlugin).
 *
 * 슬라이드가 아직 한 번도 열린 적이 없으면(slideContents에 그 entryId 키가 없으면) 서버에 지연
 * 시딩을 요청만 하고 빈 캔버스로 돌아온다 — 실제 콘텐츠는 slideContents.observe가 시딩 완료를
 * 감지해 이 함수를 다시 불러줄 때 그려진다(main()의 observe 콜백 참고).
 */
function renderCanvasForSlide(entryId: string): void {
  // y-prosemirror(1.3.7)의 ySyncPlugin/yCursorPlugin은 원격 갱신·커서 이동이 있을 때마다
  // "이 view에 적용할 setMeta"를 모듈 전역 큐에 넣고 setTimeout(0)으로 나중에 한꺼번에
  // flush한다(y-prosemirror 소스의 updateMetas). updateMetas는 flush 직전
  // `syncState.binding.isDestroyed`를 확인하긴 하지만, 라이브러리 자신은 destroy() 안
  // 어디서도 그 플래그를 true로 설정하지 않는다(node_modules/y-prosemirror/dist/
  // y-prosemirror.cjs의 ProsemirrorBinding.prototype.destroy와 lib.js:29 확인 완료 — 순수
  // 업스트림 버그, 죽은 코드나 마찬가지다). 그래서 destroy()만 부르면, 이미 예약돼 있던 flush가
  // 나중에 실행되며 죽은 view에 대고 dispatch를 시도해 "Cannot read properties of null
  // (reading 'matchesNode')" 예외를 던진다 — 두 탭이 동시에 타이핑하는 중 슬라이드를 전환하는
  // 시나리오로 재현 확인했다. destroy() 호출을 한 틱 미루는 방식은 그 창을 좁힐 뿐 완전히
  // 닫지는 못한다 — 미뤄둔 그 짧은 구간에도 원격 awareness 갱신(커서 이동 등)이 도착하면
  // setMeta가 다시 예약되어 같은 충돌이 그대로 재현될 수 있다. 실제 수정은 라이브러리가
  // 빠뜨린 그 플래그를 우리가 destroy() 호출 "직전"에 직접 true로 설정해 두는 것이다 —
  // view.state(그리고 그 안의 ySyncPlugin state가 들고 있는 binding 객체)는 destroy() 이후에도
  // 참조가 살아 있으므로, 이렇게 하면 언제 예약된 flush든 실행 시점에 항상 이 가드에 걸려
  // 안전하게 no-op된다(타이밍에 기대지 않는 결정적 해결).
  const previousViews = [...mountedViews.values()];
  for (const view of previousViews) {
    const syncState = ySyncPluginKey.getState(view.state) as { binding?: { isDestroyed?: boolean } } | undefined;
    if (syncState?.binding) syncState.binding.isDestroyed = true;
    // focusedView가 지금 destroy하려는 view를 가리키고 있다면 반드시 여기서 null로 되돌린다 —
    // 그러지 않으면 (예: 이 텍스트박스에 커서를 두고 있다가 다른 슬라이드 카드를 클릭한 경우)
    // 서식 툴바가 이미 destroy된(= docView가 사라진) view에 대고 dispatch를 시도해 예외를 던진다.
    if (focusedView === view) focusedView = null;
    view.destroy();
  }
  mountedViews.clear();
  // 이전에 활성이던 슬라이드의 perSlide 구조 구독을 해제한다 — 그러지 않으면 다른 슬라이드로
  // 옮긴 뒤에도 그 슬라이드에 텍스트박스가 추가/삭제될 때마다(원격 협업자가 하는 것 포함) 이미
  // 화면에 없는 슬라이드를 위해 불필요하게 이 함수가 다시 불린다.
  activePerSlideUnobserve?.();
  activePerSlideUnobserve = null;
  canvasEl.style.width = `${emuToPx(SLIDE_W_EMU)}px`;
  canvasEl.style.height = `${emuToPx(SLIDE_H_EMU)}px`;
  canvasEl.style.setProperty('--slide-pt', `${slidePtPx()}px`);
  canvasEl.replaceChildren();

  // 이미지·단순 도형·표/차트·스마트아트 자리표시자 — 편집 대상이 아니라 텍스트박스보다 먼저
  // 그려서(DOM 순서상 뒤에 오는 형제가 위에 쌓이므로) 항상 텍스트박스 아래에 깔리게 한다.
  for (const visual of slideVisuals.get(entryId) ?? []) {
    canvasEl.append(renderSlideVisual(visual));
  }

  const perSlide = slideContents.get(entryId);
  if (!perSlide) {
    // room.ts의 MESSAGE_SLIDE_CONTENT_SUBSCRIBE 핸들러는 이제 읽기 전용 연결의 순수 시딩 요청도
    // 그대로 처리한다(duplicateOfId가 실린 요청만 거부한다 — room.ts 주석 참고): 시딩은 원본 XML
    // 바이트를 Y.Doc에 그대로 옮기는 읽기/캐시 동작일 뿐 저장되는 파일을 바꾸지 않기 때문이다.
    // 그래서 canEdit 여부와 무관하게 항상 구독을 요청하고, observe가 다시 불러줄 때까지 빈 캔버스로
    // 기다린다(이전에는 읽기 전용일 때 서버가 요청 자체를 거부해 안내 문구만 보여줬다).
    subscribeSlideContent(entryId);
    return;
  }
  // perSlide는 이미 존재하지만(슬라이드가 한 번은 시딩됨) 그 안에서 텍스트박스가 나중에
  // 추가/삭제될 수 있다("텍스트 추가" 버튼, 삭제 버튼, 또는 원격 협업자의 같은 조작) — 그 구조적
  // 변경(=perSlide 자신의 키 추가/삭제)을 감지해 다시 그린다. 얕은 observe라 텍스트박스 "안"
  // 내용 편집(타이핑)에는 반응하지 않는다(위 activePerSlideUnobserve 선언부 주석 참고).
  const onPerSlideStructureChange = () => renderCanvasForSlide(entryId);
  perSlide.observe(onPerSlideStructureChange);
  activePerSlideUnobserve = () => perSlide.unobserve(onPerSlideStructureChange);

  for (const [shapeId, shapeMap] of perSlide.entries()) {
    const meta = shapeMap.get('meta') as SlideShapeMeta;
    const box = document.createElement('div');
    box.className = 'slide-textbox';
    box.dataset.shapeId = shapeId; // 디버깅/테스트에서 개별 텍스트박스를 식별하기 위함(동작에는 영향 없음)
    box.style.left = `${emuToPx(meta.x ?? SLIDE_W_EMU * 0.2)}px`;
    box.style.top = `${emuToPx(meta.y ?? SLIDE_H_EMU * 0.4)}px`;
    box.style.width = `${emuToPx(meta.w ?? SLIDE_W_EMU * NEW_TEXTBOX_WIDTH_RATIO)}px`;
    box.style.height = `${emuToPx(meta.h ?? SLIDE_H_EMU * NEW_TEXTBOX_HEIGHT_RATIO)}px`;
    if (meta.geom) {
      // 리본 "도형"으로 넣은 도형 — 저장 파일의 prstGeom/solidFill과 같은 모양으로 그린다.
      box.classList.add('slide-shape', `slide-shape-${meta.geom}`);
      if (meta.fill && /^#[0-9A-Fa-f]{6}$/.test(meta.fill)) box.style.background = meta.fill;
    }
    canvasEl.append(box);

    if (canEdit) {
      // 호버 시에만 보이는 작은 × 삭제 버튼(CSS: .slide-textbox-delete). perSlide.delete는
      // 위에서 이 슬라이드에 붙여 둔 onPerSlideStructureChange 구독을 트리거해 renderCanvasForSlide를
      // 다시 부르므로, 남은 텍스트박스 전체가 destroy→재생성되는 이미 검증된 기존 경로(이 함수 맨
      // 위의 previousViews 정리)를 그대로 타 삭제된 shapeId의 EditorView도 함께 정리된다 — 별도
      // cleanup 로직을 새로 만들 필요가 없다.
      const deleteBtn = document.createElement('button');
      deleteBtn.type = 'button';
      deleteBtn.className = 'slide-textbox-delete';
      deleteBtn.textContent = '×';
      deleteBtn.title = '텍스트 상자 삭제';
      deleteBtn.setAttribute('aria-label', '텍스트 상자 삭제');
      deleteBtn.addEventListener('click', (event) => {
        event.stopPropagation();
        const ydoc = slidesArray.doc!;
        ydoc.transact(() => {
          perSlide.delete(shapeId);
        });
      });
      box.append(deleteBtn);
    }

    const fragment = shapeMap.get('content') as Y.XmlFragment;
    const view = new EditorView(box, {
      state: EditorState.create({
        schema: pptxTextSchema,
        plugins: [ySyncPlugin(fragment), yCursorPlugin(provider.awareness), yUndoPlugin(), keymap(buildSlideKeymap())],
      }),
      editable: () => canEdit,
      // 커서가 움직이거나 서식이 바뀔 때마다 리본의 눌림 상태(굵게/정렬 등)와 글꼴·크기 칸을
      // 지금 커서 위치의 값으로 맞춘다 — 실제 오피스 리본과 같은 피드백. 주의: 바깥 `const view`를
      // 여기서 참조하면 안 된다 — ySyncPlugin이 EditorView **생성자 안에서**(plugin view 초기화 중
      // _forceRerender) 이미 dispatch하므로 그 시점엔 view가 아직 TDZ라 ReferenceError로 텍스트
      // 상자 전체가 비어 버린다(실제 브라우저 테스트로 재현). ProseMirror는 이 콜백의 this를 view로
      // 바인딩해 주므로 this를 쓴다.
      dispatchTransaction(this: EditorView, tr) {
        this.updateState(this.state.apply(tr));
        if (focusedView === this) refreshRibbonState(this);
      },
      // 서식 툴바가 "지금 어느 텍스트박스를 편집 중인지" 알 수 있는 유일한 경로 — handleDOMEvents의
      // 반환값 false는 "이 이벤트를 ProseMirror가 가로채 막지 말고 브라우저 기본 동작(포커스 이동)을
      // 그대로 진행시켜라"는 뜻이다(true를 반환하면 기본 동작이 막힌다). focus 이벤트는 이 view가
      // DOM에 붙은 뒤 나중에(사용자가 클릭하거나 명시적으로 .focus()를 부를 때)만 발생하므로, 이
      // 콜백 안에서 아직 대입 중인 const view를 참조해도(클로저) 실제 호출 시점엔 이미 대입이 끝나
      // 있어 안전하다 — editor.ts가 dispatchTransaction에서 겪은 "생성자 안에서 view 참조" 문제와
      // 달리, focus는 생성자 실행 중 동기적으로 발생하지 않는다.
      handleDOMEvents: {
        focus: () => {
          focusedView = view;
          refreshRibbonState(view);
          return false;
        },
      },
    });
    // 레이아웃·마스터에서 물려받은 "보이는 모양"(src/pptx/inheritance.ts) — 서식이 없는 글자의 기본
    // 크기·정렬과 상자 안 세로 맞춤. 도형(geom)은 클래스가 이미 세로 가운데를 맡는다.
    if (meta.anchor && !meta.geom) box.classList.add(`slide-anchor-${meta.anchor}`);
    if (meta.defaultSizePt) view.dom.style.fontSize = `calc(${meta.defaultSizePt} * var(--slide-pt, 1pt))`;
    const alignCss = { ctr: 'center', r: 'right', just: 'justify', l: '' }[meta.defaultAlign ?? 'l'];
    if (alignCss) view.dom.style.textAlign = alignCss;
    // 목록 수준별 글머리·들여쓰기·크기 — CSS(.slide-textbox p[data-lvl])가 이 변수들을 읽는다.
    meta.levels?.forEach((level, i) => {
      view.dom.style.setProperty(`--l${i}-marl`, String(level.marLPt));
      view.dom.style.setProperty(`--l${i}-ind`, String(level.indentPt));
      view.dom.style.setProperty(`--l${i}-bu`, level.bullet === null ? 'none' : JSON.stringify(level.bullet));
      view.dom.style.setProperty(`--l${i}-sz`, `calc(${level.sizePt} * var(--slide-pt, 1pt))`);
      if (level.color && /^#[0-9A-Fa-f]{6}$/.test(level.color)) view.dom.style.setProperty(`--l${i}-color`, level.color);
    });
    viewDefaults.set(view, { sizePt: meta.defaultSizePt ?? DEFAULT_TEXT_PT, align: meta.defaultAlign ?? 'l' });
    mountedViews.set(shapeId, view);
  }
}

function cardFor(entry: SlideEntry, index: number): HTMLElement {
  const fragment = cardTemplate.content.cloneNode(true) as DocumentFragment;
  const card = fragment.querySelector('.slide-card') as HTMLElement;
  card.dataset.index = String(index);
  card.dataset.id = entry.id;
  const img = card.querySelector('img') as HTMLImageElement;
  // 복제본은 sourcePageNumber가 null이라(자기 원본 페이지가 따로 없다) 그대로 쓰면
  // /api/view/pages/null로 400이 나며 썸네일이 깨졌다 — 같은 원본 슬라이드 파트를 가리키는 항목의
  // 페이지 번호를 빌려 쓴다(빈 슬라이드를 복제했으면 원본 파트가 없으니 빈 슬라이드 그림).
  const pageNumber =
    entry.sourcePageNumber ??
    (entry.sourceSlidePart
      ? slidesArray.toArray().find((e) => e.sourceSlidePart === entry.sourceSlidePart && e.sourcePageNumber !== null)
          ?.sourcePageNumber ?? null
      : null);
  if (entry.kind === 'blank' || pageNumber === null) {
    img.src = '/img/blank-slide.svg';
  } else {
    const base = `/api/view/pages/${pageNumber}?token=${encodeURIComponent(token ?? '')}`;
    // 저장 직후 서버는 슬라이드 목록(새 페이지 번호)을 먼저 갱신하고 PDF 재변환은 그 뒤에 수 초
    // 걸려 끝난다(room.ts flushToStore — 순서를 바꾸면 동시 편집 보호 구간이 길어진다). 그 사이
    // 새로 생긴 페이지 번호의 썸네일은 404가 나므로, 잠깐 기다렸다 다시 받는다(최대 약 15초).
    let attempt = 0;
    img.addEventListener('error', () => {
      if (attempt >= 8) return;
      attempt += 1;
      setTimeout(() => {
        if (img.isConnected) img.src = `${base}&retry=${attempt}`;
      }, 1000 + attempt * 400);
    });
    img.src = base;
  }
  (card.querySelector('.slide-card-number') as HTMLElement).textContent = String(index + 1);
  if (!canEdit) {
    card.querySelector('.slide-card-actions')?.remove();
    card.draggable = false;
  }
  return card;
}

function renderSlides(): void {
  const entries = slidesArray.toArray();
  railEl.replaceChildren(...entries.map((entry, index) => cardFor(entry, index)));
  // 아직 선택된 슬라이드가 없거나(최초 로드), 선택돼 있던 슬라이드가 방금 삭제됐다면 첫 번째
  // 슬라이드로 선택을 옮긴다 — 그러지 않으면 삭제 후 레일에 활성 표시가 하나도 안 남는다.
  const activeStillExists = entries.some((entry) => entry.id === activeSlideId);
  if (!activeStillExists) {
    if (entries.length > 0) {
      selectSlide(entries[0].id);
      return; // selectSlide()가 아래와 동일한 active 클래스 갱신을 이미 하고 끝났다
    }
    activeSlideId = null;
  }
  for (const card of railEl.querySelectorAll('.slide-card')) {
    card.classList.toggle('active', (card as HTMLElement).dataset.id === activeSlideId);
  }
}

// 슬라이드 카드 클릭(또는 렌더 직후 기본 선택)으로 오른쪽 캔버스에 표시할 슬라이드를 바꾼다.
// 실제 콘텐츠 렌더링은 onActiveSlideChange 콜백(renderCanvasForSlide)이 처리한다 — 여기서는 상태/표시만.
function selectSlide(entryId: string): void {
  // 이미 활성 상태인 슬라이드를 다시 클릭한 경우 — 아무것도 바뀌지 않았으니 다시 그리지 않는다.
  // renderCanvasForSlide는 호출될 때마다 그 슬라이드의 모든 EditorView를 destroy→재생성하므로,
  // 이 가드가 없으면 재클릭할 때마다 편집 중이던 커서 위치/선택 영역이 이유 없이 날아간다.
  if (entryId === activeSlideId) return;
  activeSlideId = entryId;
  for (const card of railEl.querySelectorAll('.slide-card')) {
    card.classList.toggle('active', (card as HTMLElement).dataset.id === entryId);
  }
  onActiveSlideChange?.(entryId);
}

function setupDragAndDrop(): void {
  railEl.addEventListener('dragstart', (event) => {
    const card = (event.target as HTMLElement).closest<HTMLElement>('.slide-card');
    if (!card) return;
    dragFromIndex = Number(card.dataset.index);
    event.dataTransfer?.setData('text/plain', card.dataset.id ?? '');
  });
  railEl.addEventListener('dragover', (event) => {
    if (dragFromIndex !== null) event.preventDefault();
  });
  railEl.addEventListener('drop', (event) => {
    event.preventDefault();
    const target = (event.target as HTMLElement).closest<HTMLElement>('.slide-card');
    if (!target || dragFromIndex === null) return;
    const toIndex = Number(target.dataset.index);
    if (toIndex === dragFromIndex) return;
    const ydoc = slidesArray.doc!;
    ydoc.transact(() => {
      const [moved] = slidesArray.slice(dragFromIndex!, dragFromIndex! + 1);
      slidesArray.delete(dragFromIndex!, 1);
      // 삭제 후 배열이 한 칸 줄어들므로, 드롭 대상이 삭제 지점보다 뒤에 있었다면 그 대상의
      // 실제(삭제 후) 위치는 원래 toIndex보다 1 작다 — 항상 "대상 카드 바로 앞"에 꽂히도록
      // 두 경우 모두 대상의 삭제-후 위치에 삽입한다.
      const insertAt = toIndex > dragFromIndex! ? toIndex - 1 : toIndex;
      slidesArray.insert(insertAt, [moved]);
    });
    dragFromIndex = null;
  });
}

function setupCardActions(): void {
  railEl.addEventListener('click', (event) => {
    const target = event.target as HTMLElement;
    const card = target.closest<HTMLElement>('.slide-card');
    if (!card) return;
    // 복제/삭제 버튼이 아니라 카드 자체(썸네일 등)를 클릭했다면 선택만 하고 끝난다 — 읽기
    // 전용 열람자도 슬라이드를 넘겨 보되(canEdit 게이트 없이), 복제/삭제 버튼은 어차피
    // cardFor()에서 이미 DOM에서 제거돼 있어 아래 분기로는 내려오지 않는다.
    if (!target.closest('.slide-card-duplicate') && !target.closest('.slide-card-delete')) {
      selectSlide(card.dataset.id!);
      return;
    }
    if (!canEdit) return;
    const index = Number(card.dataset.index);
    if (target.closest('.slide-card-duplicate')) duplicateSlideAt(index);
    else if (target.closest('.slide-card-delete')) deleteSlideAt(index);
  });
  document.getElementById('btn-add-blank')?.addEventListener('click', addBlankSlide);
  document.getElementById('btn-dup-slide')?.addEventListener('click', () => {
    const index = activeSlideIndex();
    if (canEdit && index >= 0) duplicateSlideAt(index);
  });
  document.getElementById('btn-del-slide')?.addEventListener('click', () => {
    const index = activeSlideIndex();
    if (canEdit && index >= 0) deleteSlideAt(index);
  });
  document.getElementById('btn-add-textbox')?.addEventListener('click', () => addShapeToActiveSlide());
  for (const geom of NEW_SHAPE_GEOMS) {
    document.getElementById(`btn-shape-${geom.toLowerCase()}`)?.addEventListener('click', () => {
      const fill = (document.getElementById('shape-fill-picker') as HTMLInputElement | null)?.value ?? '#4472c4';
      addShapeToActiveSlide({ geom, fill });
    });
  }
}

function activeSlideIndex(): number {
  return activeSlideId ? slidesArray.toArray().findIndex((e) => e.id === activeSlideId) : -1;
}

/** 실제 프레젠테이션 프로그램처럼 새 슬라이드는 "지금 보고 있는 슬라이드 바로 뒤"에 들어가고
 * 곧바로 그 슬라이드로 이동한다(예전엔 항상 맨 끝에 붙고 선택도 안 바뀌어, 긴 문서에서는 추가된
 * 슬라이드를 찾아 스크롤해야 했다). */
function addBlankSlide(): void {
  if (!canEdit) return;
  const index = activeSlideIndex();
  const insertAt = index >= 0 ? index + 1 : slidesArray.length;
  const id = crypto.randomUUID();
  slidesArray.insert(insertAt, [{ id, kind: 'blank', sourceSlidePart: null, sourcePageNumber: null }]);
  selectSlide(id);
}

function deleteSlideAt(index: number): void {
  if (slidesArray.length <= 1) return; // 마지막 남은 슬라이드는 지울 수 없다
  const wasActive = slidesArray.get(index)?.id === activeSlideId;
  slidesArray.delete(index, 1);
  // 지운 슬라이드를 보고 있었다면 그 자리(없으면 바로 앞)의 슬라이드로 이동한다
  if (wasActive) {
    const next = slidesArray.get(Math.min(index, slidesArray.length - 1));
    if (next) selectSlide(next.id);
  }
}

function duplicateSlideAt(index: number): void {
  const entry = slidesArray.get(index);
  const newEntryId = crypto.randomUUID();
  // 아래 proactive subscribeSlideContent 호출이 그 순간 WebSocket이 OPEN이 아니라서(재연결
  // 중 등) 아무것도 안 보내고 조용히 리턴하더라도, 이 매핑은 이미 존재하게 만들어 둔다 —
  // 그래야 나중에 어떤 경로로든(renderCanvasForSlide의 폴백, 또는 pendingSlideSubscriptions
  // 타임아웃 이후의 재시도) 이 슬라이드에 대한 구독이 결국 시도될 때 duplicateOfId 없는
  // "평범한" 요청이 나가는 대신 항상 이 원본 id를 함께 실어 보낼 수 있다.
  pendingDuplicateSources.set(newEntryId, entry.id);
  const ydoc = slidesArray.doc!;
  ydoc.transact(() => {
    slidesArray.insert(index + 1, [
      {
        id: newEntryId,
        kind: 'duplicate',
        sourceSlidePart: entry.sourceSlidePart,
        sourcePageNumber: null,
      },
    ]);
  });
  // 서버가 원본 파일의(오래된) 디스크 XML이 아니라 소스 슬라이드의 "지금 세션에서 편집된"
  // 콘텐츠를 그대로 깊은 복사해 시딩하도록, 사용자가 실제로 이 복제본을 열기 전에(활성
  // 슬라이드가 되는 것과 무관하게) 미리 duplicateOfId를 실어 구독 요청을 보내 둔다. 이 전송
  // 자체가 드롭돼도(WS가 그 순간 OPEN이 아니었던 경우) 위에서 이미 채워둔
  // pendingDuplicateSources 덕분에, 이후 이 슬라이드에 대해 실제로 나가는 첫 구독 요청은
  // subscribeSlideContent 내부에서 이 매핑을 찾아 duplicateOfId를 채워 보낸다.
  subscribeSlideContent(newEntryId, entry.id);
  selectSlide(newEntryId);
}

/** 지금 보고 있는 슬라이드에 텍스트 상자(shape 없음) 또는 도형(geom/fill)을 하나 추가한다. */
function addShapeToActiveSlide(shape?: { geom: NewShapeGeom; fill: string }): void {
  if (!canEdit || !activeSlideId) return;
  const entryId = activeSlideId;
  const entry = slidesArray.toArray().find((e) => e.id === entryId);
  if (!entry) return; // 이미 삭제된 슬라이드를 향한 레이스(카드가 지워진 직후의 늦은 클릭 등) — 방어적 no-op

  // FINAL REVIEW C1(Critical): slideContents에 이 슬라이드의 perSlide가 아직 없을 때, 원인이
  // 두 가지로 갈린다 —
  //  - kind:'blank'(원래부터 도형이 없는, 사용자가 직접 추가한 빈 슬라이드): 지금 여기서 새로
  //    만드는 게 맞다(원본이라 부를 것 자체가 없다).
  //  - kind:'original'/'duplicate'인데 아직 시딩 전(구독 응답을 못 받음 — 예: 이 슬라이드를
  //    고른 순간 WS가 재연결 중이었거나 서버 응답이 아직 안 왔거나): 여기서 새 perSlide를
  //    만들면 그 안엔 방금 추가한 도형 "하나만" 담기고, src/pptx/collab.ts의 export 삭제 병합
  //    루프(Task 13)는 "baseline에는 있는데 지금 perSlide에는 없는 도형 = 사용자가 삭제했다"로
  //    해석해 다음 저장에서 이 슬라이드의 원본 도형을 전부 지워버린다 — 텍스트박스 하나
  //    추가했을 뿐인데 슬라이드의 나머지 원본 내용이 통째로 사라지는 셈이다. 그러니 이 경우엔
  //    텍스트박스를 추가하지 않고, 구독을 다시 보낸 뒤(끊겼을 수 있으므로 — subscribeSlideContent
  //    자체가 이미 진행 중인 요청과 중복 전송은 막아 준다) 클릭 자체를 조용히 무시한다. 시딩이
  //    끝나면 slideContents.observe가 캔버스를 실제 원본 내용으로 다시 그려 주므로, 사용자는
  //    그걸 본 뒤 "텍스트 추가"를 다시 누르면 된다.
  //
  // 이 조건의 boolean 공식 자체는 shouldBlockAddTextbox(pptxShared.ts)로 뽑아 뒀다 — 이 파일은
  // 모듈 최상단에서 document/location을 곧바로 읽어 브라우저 전역 없이는 import조차 안 되므로
  // (이 저장소엔 아직 jsdom 테스트 환경이 없다), 가드의 "판단" 부분만 DOM 의존성 없는 공유
  // 모듈로 옮겨 plain vitest로 직접 단위 테스트할 수 있게 했다 —
  // tests/editorSlidesAddTextboxGuard.test.ts 참고.
  if (shouldBlockAddTextbox(entry.kind, slideContents.has(entryId))) {
    subscribeSlideContent(entryId);
    return;
  }

  const ydoc = slidesArray.doc!;
  ydoc.transact(() => {
    // 슬라이드가 완전히 빈 상태(원본이 blank로 만들어짐)라면 slideContents에 이 슬라이드의
    // perSlide 항목 자체가 없을 수 있다 — 그 경우 여기서 새로 만든다(위 가드가 이미 kind !==
    // 'blank'인데 아직 시딩 전인 경우를 걸러냈으므로, 여기 도달했는데 perSlide가 없다면 항상
    // kind === 'blank'이거나 이미 시딩이 끝난 상태다).
    let perSlide = slideContents.get(entryId);
    if (!perSlide) {
      perSlide = new Y.Map<ShapeContentMap>();
      slideContents.set(entryId, perSlide);
    }
    const shapeId = `new-${crypto.randomUUID()}`;
    const shapeMap: ShapeContentMap = new Y.Map();
    // 같은 버튼을 연달아 눌러도 정확히 겹쳐 "하나만 생긴 것처럼" 보이지 않게, 이미 있는 새 도형
    // 개수만큼 실제 프로그램처럼 대각선으로 조금씩 비켜 놓는다.
    const stagger = [...perSlide.keys()].filter((k) => k.startsWith('new-')).length % 8;
    const nudge = Math.round(SLIDE_W_EMU * 0.025) * stagger;
    const w = Math.round(SLIDE_W_EMU * (shape ? 0.3 : NEW_TEXTBOX_WIDTH_RATIO));
    const h = Math.round(SLIDE_H_EMU * (shape ? 0.3 : NEW_TEXTBOX_HEIGHT_RATIO));
    const meta: SlideShapeMeta = {
      shapeId,
      x: Math.round((SLIDE_W_EMU - w) / 2) + nudge,
      y: Math.round((SLIDE_H_EMU - h) / 2) + nudge,
      w,
      h,
      ...(shape ? { geom: shape.geom, fill: shape.fill } : {}),
    };
    shapeMap.set('meta', meta);
    const content = new Y.XmlFragment();
    if (shape) {
      // 도형 글자는 실제 프로그램처럼 가로 가운데로 시작한다(세로 가운데는 저장 시 bodyPr anchor=ctr).
      // y-prosemirror는 XmlElement 이름=노드 타입, 속성=노드 attrs로 그대로 대응시킨다.
      const paragraph = new Y.XmlElement('paragraph');
      paragraph.setAttribute('align', 'ctr');
      content.insert(0, [paragraph]);
    }
    shapeMap.set('content', content);
    perSlide.set(shapeId, shapeMap);
  });
  // perSlide가 이미 있었다면 위 변경은 slideContents 자신의 키가 아니라 그 안의 perSlide
  // Y.Map에 대한 나머지 변경이라 최상위 slideContents.observe(main() 참고)로는 안 올라온다 —
  // 하지만 이 슬라이드가 지금 활성 상태였다면(대부분의 경우) renderCanvasForSlide가 이미 이
  // perSlide에 onPerSlideStructureChange 구독을 붙여 뒀으므로(위 renderCanvasForSlide 참고)
  // 그 구독이 알아서 다시 그려준다 — perSlide를 방금 새로 만든 경우(원래 shapeless)만
  // slideContents.set이 최상위 키 변경이라 그 observe가 대신 다시 그려준다. 두 경로 모두
  // 여기서 별도로 renderCanvasForSlide를 부를 필요가 없다.
}

/**
 * 서식 툴바(굵게/기울임/밑줄/색) 버튼을 지금 focusedView에 잡혀 있는 EditorView에 연결한다.
 * editor.ts(docx 편집기)의 bindToolbarButton/font-color 배선과 같은 패턴 — 일반 click 이벤트를
 * 쓰고(mousedown+preventDefault가 아니다) 커맨드 적용 뒤 view.focus()로 포커스를 되돌린다. 이
 * 패턴이 안전한 이유: 버튼 mousedown이 지금 포커스된 contentEditable을 blur시키더라도
 * ProseMirror의 EditorState.selection은 DOM 포커스와 별개로 유지되는 순수 JS 상태라 click
 * 핸들러가 실행되는 시점에도 그대로 남아 있다 — focusedView는 focus 이벤트에서만 갱신하고 blur
 * 에서는 절대 건드리지 않으므로(renderCanvasForSlide 위쪽 주석 참고) 이 사이에도 값을 잃지 않는다.
 */
function bindFormatButton(id: string, markName: 'strong' | 'em' | 'underline'): void {
  document.getElementById(id)?.addEventListener('click', () => {
    if (!focusedView) return; // 포커스된 텍스트박스가 없으면 조용히 아무 것도 하지 않는다.
    toggleMark(pptxTextSchema.marks[markName])(focusedView.state, focusedView.dispatch);
    focusedView.focus();
  });
}

/** 프레젠테이션 프로그램의 "글꼴 크기 크게/작게" 단계(pt) — 그 사이 값은 다음 단계로 올라간다. */
const FONT_SIZE_STEPS = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
/** 글자 크기 마크가 없는 텍스트의 기준 크기 — 새 텍스트 상자 본문 기본값(18pt)과 같다. */
const DEFAULT_TEXT_PT = 18;
/** 텍스트 상자별로 레이아웃·마스터에서 물려받은 기본 크기·정렬 — 리본이 "서식 없는 글자"의 실제
 * 보이는 값을 보여주고(가운데 맞춤 부제목이면 가운데 버튼이 눌림), 크게/작게도 거기서 출발한다. */
const viewDefaults = new WeakMap<EditorView, { sizePt: number; align: 'l' | 'ctr' | 'r' | 'just' }>();

/** 커서/선택 위치의 마크 목록 — 선택 영역이 없으면 "다음 입력에 붙을" stored mark를 우선한다
 * (예: 빈 줄에서 굵게를 누른 직후의 눌림 상태도 정확히 보이도록). */
function marksAtSelection(state: EditorState): readonly Mark[] {
  const { $from, from, to, empty } = state.selection;
  if (empty) return state.storedMarks ?? $from.marks();
  // 선택 영역이면 그 안의 첫 글자의 마크(실제 오피스도 선택 시작 글자의 서식을 보여준다). $from을
  // 그대로 쓰면 안 된다 — Ctrl+A(AllSelection)는 $from이 문단 밖(doc 루트, pos 0)이라 그 자리엔
  // 글자가 없어 항상 "서식 없음"으로 보였다(브라우저 테스트로 재현).
  let found: readonly Mark[] | null = null;
  state.doc.nodesBetween(from, to, (node) => {
    if (found) return false;
    if (node.isText) found = node.marks;
    return !found;
  });
  return found ?? $from.marks();
}

/** 선택 영역의 첫 문단 — AllSelection에서도 실제 문단을 찾는다(위 marksAtSelection과 같은 이유). */
function firstParagraphInSelection(state: EditorState): PMNode | null {
  const { $from, from, to } = state.selection;
  if ($from.parent.type.name === 'paragraph') return $from.parent;
  let found: PMNode | null = null;
  state.doc.nodesBetween(from, to, (node) => {
    if (found) return false;
    if (node.type.name === 'paragraph') found = node;
    return !found;
  });
  return found;
}

function markAttrAtSelection(state: EditorState, typeName: string, attr: string): unknown {
  return marksAtSelection(state).find((m) => m.type.name === typeName)?.attrs[attr];
}

/** 같은 종류의 마크를 값만 바꿔 다시 건다(글꼴/크기/색처럼 attrs를 가진 마크). attrs가 null이면
 * 그 마크를 지운다. 선택 영역이 없으면 다음 입력에 적용되는 stored mark로 처리한다. */
function setValueMark(view: EditorView, type: MarkType, attrs: Record<string, unknown> | null): void {
  const { state } = view;
  const { from, to, empty } = state.selection;
  let tr = state.tr;
  if (empty) {
    tr = tr.removeStoredMark(type);
    if (attrs) tr = tr.addStoredMark(type.create(attrs));
  } else {
    tr = tr.removeMark(from, to, type);
    if (attrs) tr = tr.addMark(from, to, type.create(attrs));
  }
  view.dispatch(tr);
  view.focus();
}

function applyAlign(view: EditorView, align: 'l' | 'ctr' | 'r' | 'just'): void {
  const { state } = view;
  const { from, to } = state.selection;
  let tr = state.tr;
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (node.type.name !== 'paragraph') return true;
    if (node.attrs.align !== align) tr = tr.setNodeMarkup(pos, undefined, { ...node.attrs, align });
    return false;
  });
  if (tr.docChanged) view.dispatch(tr);
  view.focus();
}

function stepFontSize(view: EditorView, direction: 1 | -1): void {
  const current = Number(
    markAttrAtSelection(view.state, 'fontSize', 'size') ?? viewDefaults.get(view)?.sizePt ?? DEFAULT_TEXT_PT,
  );
  const next =
    direction > 0
      ? FONT_SIZE_STEPS.find((s) => s > current) ?? current
      : [...FONT_SIZE_STEPS].reverse().find((s) => s < current) ?? current;
  if (next !== current) setValueMark(view, pptxTextSchema.marks.fontSize, { size: next });
}

/** select에 현재 값이 목록에 없으면(원본 문서의 낯선 글꼴, 15.5pt 같은 크기) 임시 항목을 끼워
 * 넣어서라도 정확한 값을 보여준다 — 목록에 없다고 엉뚱한 값을 보여주면 사용자가 오해한다. */
function showSelectValue(select: HTMLSelectElement | null, value: string): void {
  if (!select) return;
  select.querySelector('option[data-transient]')?.remove();
  if (value && ![...select.options].some((o) => o.value === value)) {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = value;
    opt.dataset.transient = '1';
    select.append(opt);
  }
  select.value = value;
}

const ALIGN_BUTTONS: ReadonlyArray<readonly [string, 'l' | 'ctr' | 'r' | 'just']> = [
  ['btn-align-left', 'l'],
  ['btn-align-center', 'ctr'],
  ['btn-align-right', 'r'],
  ['btn-align-justify', 'just'],
];

function refreshRibbonState(view: EditorView): void {
  const state = view.state;
  const marks = marksAtSelection(state);
  const has = (name: string) => marks.some((m) => m.type.name === name);
  document.getElementById('btn-bold')?.setAttribute('aria-pressed', String(has('strong')));
  document.getElementById('btn-italic')?.setAttribute('aria-pressed', String(has('em')));
  document.getElementById('btn-underline')?.setAttribute('aria-pressed', String(has('underline')));

  // 정렬 속성이 없는 문단은 왼쪽 맞춤으로 보인다(실제 렌더링과 같은 기본값)
  const align = (firstParagraphInSelection(state)?.attrs.align as string | null) ?? viewDefaults.get(view)?.align ?? 'l';
  for (const [id, value] of ALIGN_BUTTONS) {
    document.getElementById(id)?.setAttribute('aria-pressed', String(align === value));
  }

  showSelectValue(
    document.getElementById('font-family') as HTMLSelectElement | null,
    String(markAttrAtSelection(state, 'fontFamily', 'family') ?? ''),
  );
  // 크기 서식이 없으면 레이아웃·마스터에서 물려받은 실제 보이는 크기를 보여준다(오피스와 같음).
  const size = markAttrAtSelection(state, 'fontSize', 'size') ?? viewDefaults.get(view)?.sizePt;
  showSelectValue(document.getElementById('font-size') as HTMLSelectElement | null, size === undefined ? '' : String(size));
}

function setupFormatToolbar(): void {
  bindFormatButton('btn-bold', 'strong');
  bindFormatButton('btn-italic', 'em');
  bindFormatButton('btn-underline', 'underline');

  const familySelect = document.getElementById('font-family') as HTMLSelectElement | null;
  familySelect?.addEventListener('change', () => {
    if (!focusedView) return;
    const family = familySelect.value;
    setValueMark(focusedView, pptxTextSchema.marks.fontFamily, family ? { family } : null);
  });
  const sizeSelect = document.getElementById('font-size') as HTMLSelectElement | null;
  sizeSelect?.addEventListener('change', () => {
    if (!focusedView || !sizeSelect.value) return;
    setValueMark(focusedView, pptxTextSchema.marks.fontSize, { size: Number(sizeSelect.value) });
  });
  document.getElementById('btn-font-grow')?.addEventListener('click', () => {
    if (focusedView) stepFontSize(focusedView, 1);
  });
  document.getElementById('btn-font-shrink')?.addEventListener('click', () => {
    if (focusedView) stepFontSize(focusedView, -1);
  });
  document.getElementById('btn-clear-format')?.addEventListener('click', () => {
    if (!focusedView) return;
    const { from, to, empty } = focusedView.state.selection;
    const tr = empty ? focusedView.state.tr.setStoredMarks([]) : focusedView.state.tr.removeMark(from, to);
    focusedView.dispatch(tr);
    focusedView.focus();
  });
  for (const [id, value] of ALIGN_BUTTONS) {
    document.getElementById(id)?.addEventListener('click', () => {
      if (focusedView) applyAlign(focusedView, value);
    });
  }

  const colorInput = document.getElementById('text-color-picker') as HTMLInputElement | null;
  const colorBtn = document.getElementById('text-color-btn') as HTMLElement | null;
  // .color-btn-bar는 --btn-color 커스텀 프로퍼티로 자기 배경색을 그리므로(style.css), 입력의
  // 현재 값과 항상 맞춰 둔다 — 초기값과 이후 사용자가 고른 색 둘 다.
  colorBtn?.style.setProperty('--btn-color', colorInput?.value ?? '#c00000');
  colorInput?.addEventListener('input', () => {
    colorBtn?.style.setProperty('--btn-color', colorInput.value);
    if (!focusedView) return;
    setValueMark(focusedView, pptxTextSchema.marks.textColor, { color: colorInput.value });
  });

  const fillInput = document.getElementById('shape-fill-picker') as HTMLInputElement | null;
  const fillBtn = document.getElementById('shape-fill-btn') as HTMLElement | null;
  fillBtn?.style.setProperty('--btn-color', fillInput?.value ?? '#4472c4');
  fillInput?.addEventListener('input', () => fillBtn?.style.setProperty('--btn-color', fillInput.value));

  // 슬라이드 단위 단축키 — 텍스트 편집 중이 아닐 때의 Ctrl+D(복제)는 브라우저 북마크 대신
  // 슬라이드 복제로, Ctrl+M은 어디서든 새 슬라이드로 쓴다(프레젠테이션 프로그램과 같은 키).
  window.addEventListener('keydown', (event) => {
    if (!canEdit || !(event.ctrlKey || event.metaKey) || event.shiftKey || event.altKey) return;
    const target = event.target as HTMLElement | null;
    if (target?.closest('input, select, textarea')) return;
    const key = event.key.toLowerCase();
    if (key === 'm') {
      event.preventDefault();
      addBlankSlide();
    } else if (key === 'd' && !target?.closest('.slide-textbox')) {
      event.preventDefault();
      const index = activeSlideIndex();
      if (index >= 0) duplicateSlideAt(index);
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
  titleEl.textContent = info.name;
  document.title = `${info.name} — DocView 편집`;
  canEdit = info.permissions.canEdit;
  if (info.slideSize && info.slideSize.cx > 0 && info.slideSize.cy > 0) {
    SLIDE_W_EMU = info.slideSize.cx;
    SLIDE_H_EMU = info.slideSize.cy;
  }
  // 읽기 전용 토큰이어도 이 화면 자체는 그대로 보여주고 쓰기만 막는다 — grid-editor.ts와 동일한
  // 패턴(canEdit:false로 fatal error를 띄우면 "바로보기"가 동시편집 화면을 못 씀). 실제 쓰기
  // 차단은 room.ts의 readOnlyConns(서버 쪽)가 이미 하고 있고, 여기서는 UI만 맞춘다.
  if (!canEdit) {
    // 리본 전체가 편집 도구라 읽기 전용에선 통째로 숨긴다(빈 그룹 테두리만 남는 일이 없도록).
    document.getElementById('slides-ribbon')?.setAttribute('hidden', '');
    connStatusEl.insertAdjacentHTML('afterend', '<span class="conn-status" data-status="readonly">읽기 전용</span>');
  }

  const ydoc = new Y.Doc();
  slidesArray = ydoc.getArray<SlideEntry>(PPTX_SLIDES_ARRAY);
  slideContents = ydoc.getMap<PerSlideContentsMap>(PPTX_SLIDE_CONTENTS_MAP);
  slideVisuals = ydoc.getMap<readonly SlideVisual[]>(PPTX_SLIDE_VISUALS_MAP);
  const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  provider = new WebsocketProvider(`${wsProtocol}//${location.host}/ws/edit`, info.docId, ydoc, {
    params: { token },
  });

  const myColor = pickPresenceColorForName(info.viewer);
  provider.awareness.setLocalStateField('user', { name: info.viewer, color: myColor });

  provider.on('status', ({ status }: { status: string }) => {
    connStatusEl.dataset.status = status;
    connStatusEl.textContent = status === 'connected' ? '연결됨' : '연결 중…';
    // 재연결 직후, 지금 보고 있는 슬라이드가 아직 시딩되지 않았다면 구독을 다시 보낸다(FINAL
    // REVIEW C1의 결정론적 트리거 — M4). subscribeSlideContent는 WS가 OPEN이 아닌 순간엔 아무것도
    // 안 보내고 조용히 리턴하므로(이 파일 위쪽 주석 참고), 슬라이드를 고른 바로 그 순간 마침
    // 재연결 중이었다면 그 구독 요청은 영원히 사라지고 아무도 재시도하지 않는다 — 캔버스가 계속
    // 빈 채로 남고, 사용자가 그걸 "진짜로 빈 슬라이드"로 오해해 "텍스트 추가"를 누르면(아래
    // setupCardActions의 btn-add-textbox 가드가 이제 막는 바로 그 상황) 원본 손실로 이어질 수
    // 있었다. status가 'connected'로 바뀌는 시점마다(=이제 막 WS가 다시 OPEN이 된 시점) 다시
    // 시도한다 — 이미 시딩된 슬라이드라면 서버의 seedSlideContentIfNeeded가 has() 체크로 중복
    // 시딩을 막아 주므로 여러 번 보내도 안전하다.
    if (status === 'connected' && canEdit && activeSlideId && !slideContents.has(activeSlideId)) {
      subscribeSlideContent(activeSlideId);
    }
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
  provider.awareness.on('change', () => renderPresenceAvatars(presenceListEl, provider.awareness));

  // slideContents는 슬라이드별 키가 지연 시딩될 때만 top-level 변경이 일어난다(텍스트박스 안
  // 내용 편집은 그 nested Y.Map/Y.XmlFragment 안에서만 바뀌므로 여기로 이벤트가 올라오지 않는다
  // — 그건 각 텍스트박스에 마운트된 ySyncPlugin이 알아서 반영한다). event.keysChanged로 지금
  // 보고 있는 슬라이드의 키가 실제로 추가된 경우에만 다시 그린다 — 그러지 않고 무조건 다시
  // 그리면, 다른 슬라이드가 시딩될 때마다(다른 참여자가 다른 슬라이드를 여는 경우 등) 지금 이미
  // 마운트돼 편집 중인 EditorView까지 destroy→재생성돼 커서/포커스가 끊기고 만다.
  slideContents.observe((event) => {
    // 새로 채워진(=서버가 시딩을 끝낸) 슬라이드는 더 이상 "응답 대기 중"이 아니므로 in-flight
    // 집합에서 지운다 — 그러지 않으면 그 슬라이드에 대한 이후의 정당한 subscribeSlideContent
    // 호출이 (실제로는 이미 끝난) 이전 요청 때문에 영원히 무시된다. 같은 이유로
    // pendingDuplicateSources에서도 지운다 — 시딩이 실제로 끝난 이상 그 매핑은 더 이상 쓸모가
    // 없고(서버는 이미 채워진 슬라이드를 다시 덮어쓰지 않으므로 재시딩은 원래 일어나지 않지만),
    // 남겨두면 훗날 어떤 경로로 같은 entryId가 재사용되는 예외적인 경우에 오래된 매핑이 다시
    // 쓰일 위험이 있어 위생상 정리한다.
    for (const key of event.keysChanged) {
      pendingSlideSubscriptions.delete(key);
      pendingDuplicateSources.delete(key);
    }
    if (activeSlideId && event.keysChanged.has(activeSlideId)) renderCanvasForSlide(activeSlideId);
  });
  // slideVisuals는 slideContents와 별개 타이밍에(약간 늦게) 채워진다(room.ts 참고) — 텍스트박스가
  // 먼저 보이고 이미지/도형/표가 조금 뒤에 나타나는 정도는 자연스러운 점진적 렌더링이라 괜찮다.
  slideVisuals.observe((event) => {
    if (activeSlideId && event.keysChanged.has(activeSlideId)) renderCanvasForSlide(activeSlideId);
  });
  // Task 8이 남겨 둔 훅 — 슬라이드를 고를 때마다 오른쪽 캔버스를 그 슬라이드 내용으로 채운다.
  onActiveSlideChange = renderCanvasForSlide;

  slidesArray.observe(renderSlides);
  renderSlides();
  setupDragAndDrop();
  setupCardActions();
  setupFormatToolbar();

  // ---------- 저장 상태 표시(Office 365 스타일 "저장 중… / 저장됨") + Ctrl+S ----------
  // grid-editor.ts와 완전히 동일한 패턴 — docMeta.savedAt은 서버의 flushToStore가 실제로
  // 디스크에 쓴 뒤에만 갱신하는 일반 Y.Map 값이라(gridShared.ts, 문서 종류 공통), 이 pptx
  // 편집기도 별도 서버 변경 없이 그대로 재사용할 수 있다.
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
}

void main();
