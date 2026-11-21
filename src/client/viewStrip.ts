/**
 * 읽기 전용 편집기의 "보기 꾸밈"(ui=view, 설계 2026-11-21-readonly-editor-look §3.1). 판정·규칙은 순수 함수(tests/viewStrip.test.ts),
 * 막기(installViewGuards)와 정원 초과 이동(fallBackToViewer)은 DOM. 얇은 줄 단추를 다는 일은 officeShell.ts가 한다.
 * 실제 쓰기 차단은 서버(토큰 canEdit=false → room.ts가 Update를 버림) — 여기 막기는 화면이 어긋나지 않게 하는 겹이다.
 */

/** 메인 화면이 "열기" 때 문서 id를 넣어 두는 sessionStorage 키 — 같은 문서면 얇은 줄에 "편집 열기"(public/js/main.js) */
export const EDIT_DOC_KEY = 'docview-edit-doc';
export const CAPACITY_NOTICE = '지금은 보는 사람이 많습니다. 서버 변환으로 보여 줍니다';

export type StripAction = 'zoom-out' | 'zoom-in' | 'print' | 'download' | 'edit';

/** 두 겹 — 토큰이 읽기 전용이고 주소에 ui=view가 있을 때만(하나라도 아니면 지금 편집기 그대로) */
export function isViewUi(search: string, canEdit: boolean): boolean {
  return !canEdit && new URLSearchParams(search).get('ui') === 'view';
}

/** 이전/다음 단추는 편집기가 setViewNav로 따로 단다(쪽·시트·슬라이드) */
export function stripActions(p: { readonly canPrint: boolean; readonly canDownload: boolean; readonly hasZoom: boolean; readonly canOpenEdit: boolean }): StripAction[] {
  const actions: StripAction[] = [];
  if (p.hasZoom) actions.push('zoom-out', 'zoom-in');
  if (p.canPrint) actions.push('print');
  if (p.canDownload) actions.push('download');
  if (p.canOpenEdit) actions.push('edit');
  return actions;
}

const ZOOM_KEYS = new Set(['+', '=', '-', '_', '0']);

/**
 * 보기 꾸밈에서 먹이지 않을 키 — Ctrl+글자는 허용 목록(찾기 F·전체 선택 A·복사 C(복사 권한)·인쇄 P(인쇄 권한))만,
 * 확대(+ − 0)·화살표 이동은 그대로, 입력칸(찾기 상자) 안은 막지 않는다. code(KeyX)가 있으면 한글 입력 상태에서도 글자를 안다
 */
export function blockedViewKey(e: { readonly key: string; readonly code?: string; readonly ctrl: boolean; readonly inField: boolean; readonly canCopy: boolean; readonly canPrint?: boolean }): boolean {
  if (e.inField) return false;
  const key = e.key.toLowerCase();
  if (e.ctrl) {
    if (ZOOM_KEYS.has(key)) return false;
    const letter = /^Key([A-Z])$/.exec(e.code ?? '')?.[1].toLowerCase() ?? (/^[a-z]$/.test(key) ? key : null);
    if (!letter) return false;
    // r·l 은 브라우저 자체 단축키(새로 고침·주소창) — 보기 화면과 같이 막지 않는다
    return !(letter === 'f' || letter === 'a' || letter === 'r' || letter === 'l' || (letter === 'c' && e.canCopy) || (letter === 'p' && e.canPrint === true));
  }
  return key === 'delete' || key === 'backspace' || key === 'f2';
}

/** 목록 안에서 한 칸 앞/뒤(끝에서 멈춤). 지금 위치를 모르면(-1) 첫 칸, 빈 목록이면 -1 */
export function stepIndex(length: number, current: number, dir: 1 | -1): number {
  if (length === 0) return -1;
  if (current < 0) return 0;
  return Math.min(length - 1, Math.max(0, current + dir));
}

/** 워드 쪽 이동 — 쪽 경계들의 위치(도구 줄 아래 기준 px)에서 다음(아래 첫)/이전(위 마지막) 경계 번호, 없으면 -1. ±1px은 "지금 여기" */
export function pickBoundary(tops: readonly number[], dir: 1 | -1): number {
  if (dir === 1) return tops.findIndex((top) => top > 1);
  for (let i = tops.length - 1; i >= 0; i -= 1) if (tops[i] < -1) return i;
  return -1;
}

/** 큰 엑셀 정원 초과(edit.ts 4429)면 보기 꾸밈만 같은 토큰의 보기 화면으로 — 편집 화면은 지금처럼 오류 화면 */
export function capacityFallback(code: number, viewUi: boolean, viewerUrl: string | undefined): string | null {
  return code === 4429 && viewUi && viewerUrl ? viewerUrl : null;
}

function isField(target: EventTarget | null): boolean {
  return target instanceof HTMLElement && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA');
}

/**
 * 창의 잡기 단계에서 끊는다 — 편집기들의 키·우클릭 처리기는 모두 요소나 창의 거품 단계에 걸려 있어 여기서 멈추면 하나도 돌지 않는다.
 * 우클릭: 편집기 메뉴(행·열·시트 탭·칸)는 늘 없애고, 브라우저 기본 메뉴는 복사 권한이 있을 때만 둔다(보기 화면 view.ts setupDeterrents와 같은 규칙).
 */
export function installViewGuards(canCopy: boolean, canPrint: boolean): void {
  window.addEventListener(
    'keydown',
    (event) => {
      if (!blockedViewKey({ key: event.key, code: event.code, ctrl: event.ctrlKey || event.metaKey, inField: isField(event.target), canCopy, canPrint })) return;
      event.preventDefault();
      event.stopImmediatePropagation();
    },
    true,
  );
  window.addEventListener(
    'contextmenu',
    (event) => {
      event.stopImmediatePropagation();
      if (!canCopy) event.preventDefault();
    },
    true,
  );
  const types: Array<'paste' | 'cut' | 'drop' | 'copy'> = canCopy ? ['paste', 'cut', 'drop'] : ['paste', 'cut', 'drop', 'copy'];
  for (const type of types) {
    window.addEventListener(
      type,
      (event) => {
        if (isField(event.target)) return;
        event.preventDefault();
        event.stopImmediatePropagation();
      },
      true,
    );
  }
}

/** 편집 세션 영구 거부(edit.ts close 4xxx, 세 편집기 공통) — 보기 꾸밈의 정원 초과(4429)면 같은 토큰의 보기 화면으로(설계 §3.4), 그 밖은 이유를 보여 준다 */
export function handleSessionRejected(event: { readonly code: number; readonly reason: string }, canEdit: boolean, viewerUrl: string | undefined, showFatal: (message: string) => void): void {
  const fallback = capacityFallback(event.code, isViewUi(location.search, canEdit), viewerUrl);
  if (fallback) fallBackToViewer(fallback);
  else showFatal(event.reason || '서버가 연결을 거부했습니다.');
}

/** 안내를 띄우고 잠시 뒤 보기 화면으로(뒤로 가기로 다시 막힌 화면에 오지 않게 replace) */
export function fallBackToViewer(url: string): void {
  const note = document.createElement('div');
  note.className = 'view-fallback-notice';
  note.setAttribute('role', 'status');
  note.textContent = CAPACITY_NOTICE;
  document.body.append(note);
  setTimeout(() => location.replace(url), 1500);
}
