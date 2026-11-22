/**
 * 아트크래프트 앱 화면 입구(public/artcraft-app.html, 설계 2026-11-22-artcraft-app-screen §3). 앱 자체 화면(리본·도구 그대로)을 같은 출처 iframe에
 * 띄우고, 준비되면 한글 글꼴(font.add ×2) → 문서(:open)를 넣는다. 읽기 전용이면 워드 review.restrict·엑셀 protectSheet(+되돌리기 막기·다시 걸기),
 * PPT는 띠. 저장(:save)은 canEdit 토큰이면 PUT /api/view/artcraft-save(If-Match = 연 바이트의 ETag, 저장 뒤엔 보낸 바이트의 해시로 잇는다),
 * 아니면 무시하고 안내. 준비 20초 초과·:failed·wasm 받기 실패면 info.viewerUrl(워드·PPT = 엔진 보기 → 거기서도 실패하면 리브레오피스,
 * 엑셀 = 서버 격자)로 넘긴다. 제목 줄엔 읽기 전용 표지·"편집 열기"를 두지 않는다(사용자 규칙 — 읽기 전용은 앱의 막힌 편집으로 드러난다).
 */
import {
  NOTICES, OPEN_TIMEOUT_MS, READY_TIMEOUT_MS, SAVE_LABELS, contentRedirect, fontName, frameUrl, isUndoKey, openName, readAppMessage,
  readOnlyCommands, saveDecision, saveFailure, sha256Etag, toBase64, uiShowsDocument, unprotectedSheets, type AppKind, type SaveState,
} from './artcraftAppLogic.js';
import { fallBackToViewer, isContentUi } from './viewStrip.js';

interface AppInfo {
  readonly docId: string;
  readonly name: string;
  readonly ext: string;
  readonly viewerUrl: string;
  readonly permissions: { readonly canEdit: boolean };
  readonly artcraftApp?: { readonly kind: AppKind; readonly js: string; readonly fonts: readonly string[]; readonly originalUrl: string };
}
interface AppApi {
  execute(id: string, params?: Record<string, unknown>): Promise<unknown>;
  docInspect(): Promise<unknown>;
  uiInspect(): Promise<unknown>;
}

const token = new URLSearchParams(location.search).get('token') ?? '';
const el = (id: string) => document.getElementById(id) as HTMLElement;
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

function setSave(state: SaveState, detail = ''): void {
  document.body.dataset.saveState = state;
  el('save-status').textContent = detail ? `${SAVE_LABELS[state]} — ${detail}` : SAVE_LABELS[state];
}

let noticeTimer: ReturnType<typeof setTimeout> | null = null;
/** 한 줄 안내 — reload면 "다시 열기" 단추를 달고 닫히지 않게 둔다 */
function notice(text: string, reload = false): void {
  const box = el('app-notice');
  box.textContent = text;
  box.hidden = false;
  if (noticeTimer !== null) clearTimeout(noticeTimer);
  noticeTimer = null;
  if (reload) {
    const again = document.createElement('button');
    again.type = 'button';
    again.className = 'btn-pill';
    again.textContent = '다시 열기';
    again.addEventListener('click', () => location.reload());
    box.append(again);
  } else {
    noticeTimer = setTimeout(() => (box.hidden = true), 8000);
  }
}

function showBanner(text: string): void {
  const banner = el('app-banner');
  banner.textContent = text;
  banner.hidden = false;
  el('app-host').classList.add('with-banner');
}

function randomPassword(): string {
  return [...crypto.getRandomValues(new Uint8Array(16))].map((b) => b.toString(16).padStart(2, '0')).join('');
}

async function boot(): Promise<void> {
  if (!token) {
    el('doc-title').textContent = '열람 토큰이 없습니다. 발급받은 링크로 접속해 주세요.';
    return;
  }
  const res = await fetch(`/api/view/info?token=${encodeURIComponent(token)}`).catch(() => null);
  const json = (await res?.json().catch(() => null)) as { success?: boolean; data?: AppInfo; error?: string } | null;
  if (!res?.ok || !json?.success || !json.data) {
    el('doc-title').textContent = json?.error ?? '문서를 열 수 없습니다.';
    return;
  }
  const info = json.data;
  const app = info.artcraftApp;
  // 앱 화면을 못 쓰는 토큰(설치 안 됨·내려받기 권한 없음 등) — 보기 화면이 이유를 보인다
  if (!app) {
    location.replace(info.viewerUrl);
    return;
  }
  if (isContentUi(location.search)) {
    location.replace(contentRedirect(app.kind, token));
    return;
  }
  el('doc-title').textContent = info.name;
  document.title = `DocView — ${info.name}`;
  const canSave = info.permissions.canEdit && ['docx', 'xlsx', 'pptx'].includes(info.ext);
  const readOnly = !canSave;
  if (readOnly && app.kind === 'deck') showBanner(NOTICES.readOnlyDeck);

  let opened = false;
  let failed = false;
  let unsaved = false;
  /** 저장 기준 — engine-source ETag, 저장(200)마다 보낸 바이트의 해시로 바꾼다 */
  let base: string | null = null;
  const fail = (why: string): void => {
    if (failed || opened) return;
    failed = true;
    clearTimeout(readyTimer);
    console.warn('아트크래프트 앱 화면을 열지 못했습니다:', why);
    document.body.dataset.appState = 'failed';
    fallBackToViewer(info.viewerUrl, NOTICES.failed);
  };
  const readyTimer = setTimeout(() => fail('준비 20초 초과'), READY_TIMEOUT_MS);

  // 앱을 받는 동안 원본·글꼴도 함께 받는다
  const docBytes = fetch(app.originalUrl).then((r) => {
    if (!r.ok) throw new Error(`원본 ${r.status}`);
    base = r.headers.get('etag');
    return r.arrayBuffer();
  });
  const fontBytes = Promise.all(app.fonts.map(async (url) => [fontName(url), new Uint8Array(await (await fetch(url)).arrayBuffer())] as const));
  docBytes.catch(() => undefined);
  fontBytes.catch(() => undefined);

  const frame = document.createElement('iframe');
  frame.id = 'app-frame';
  frame.title = '아트크래프트 문서 화면';
  frame.src = frameUrl(app.kind, app.js, readOnly);
  el('app-host').append(frame);
  const fileName = openName(info.name, info.ext);
  const api = (): AppApi => (frame.contentWindow as unknown as Record<string, AppApi>)[`${app.kind}craft`];

  // 읽기 전용 엑셀 — 시트 보호는 되돌리기(edit.undo) 한 번에 풀린다(Task 1 측정). 키는 iframe 창에서 먼저 막고,
  // 메뉴·단추로 되돌린 경우는 보호가 풀린 시트를 찾아 다시 건다(dirty 메시지마다 + 1초마다)
  const gridGuard = readOnly && app.kind === 'grid';
  const password = randomPassword();
  let reprotecting = false;
  async function reprotect(): Promise<void> {
    if (!opened || reprotecting) return;
    reprotecting = true;
    try {
      for (const sheet of unprotectedSheets(await api().docInspect())) await api().execute('review.protectSheet', { sheet, password });
    } catch (error) {
      console.warn('시트 보호를 다시 걸지 못했습니다:', error);
    } finally {
      reprotecting = false;
    }
  }
  function blockUndoKeys(): void {
    frame.contentWindow?.addEventListener(
      'keydown',
      (event) => {
        if (!isUndoKey({ key: event.key, code: event.code, ctrl: event.ctrlKey || event.metaKey })) return;
        event.preventDefault();
        event.stopImmediatePropagation();
      },
      true,
    );
  }

  async function openDoc(): Promise<void> {
    try {
      if (gridGuard) blockUndoKeys();
      for (const [name, u8] of await fontBytes) await api().execute('font.add', { data: toBase64(u8), name });
      const bytes = await docBytes;
      frame.contentWindow!.postMessage({ type: `${app!.kind}craft:open`, name: fileName, bytes }, location.origin, [bytes]);
      let shown = false;
      for (const t0 = Date.now(); !shown && Date.now() - t0 < OPEN_TIMEOUT_MS; ) {
        await sleep(100);
        shown = uiShowsDocument(await api().uiInspect().catch(() => null), fileName);
      }
      if (!shown && readOnly) throw new Error('문서 열기 확인 시간 초과(읽기 전용 보호를 걸 수 없음)');
      if (readOnly) {
        const inspect = app!.kind === 'grid' ? ((await api().docInspect()) as { sheets?: unknown[] } | null) : null;
        const sheets = Array.isArray(inspect?.sheets) ? inspect.sheets.length : 1;
        for (const [id, params] of readOnlyCommands(app!.kind, password, sheets)) await api().execute(id, params);
      }
      opened = true;
      document.body.dataset.appState = 'open';
      if (gridGuard) setInterval(() => void reprotect(), 1000);
    } catch (error) {
      fail(error instanceof Error ? error.message : String(error));
    }
  }

  async function upload(bytes: ArrayBuffer): Promise<void> {
    setSave('saving');
    const next = sha256Etag(bytes);
    const headers: Record<string, string> = { 'content-type': 'application/octet-stream' };
    if (base) headers['if-match'] = base;
    const r = await fetch(`/api/view/artcraft-save?token=${encodeURIComponent(token)}`, { method: 'PUT', headers, body: bytes }).catch(() => null);
    const j = (await r?.json().catch(() => null)) as { success?: boolean; error?: string } | null;
    if (r?.ok && j?.success) {
      base = await next;
      unsaved = false;
      setSave('saved');
      return;
    }
    unsaved = true;
    const why = saveFailure(r?.status ?? 0, j?.error);
    // 428 = 기준(ETag)을 못 보낸 것 — 이 화면의 버그라 기록을 남긴다
    if (r?.status === 428) console.error('artcraft-save 428: If-Match 없이 저장을 보냈습니다', { base });
    setSave('failed', why.text);
    if (why.reload) notice(why.text, true);
  }

  window.addEventListener('message', (e) => {
    const msg = readAppMessage(e, location.origin, frame.contentWindow, app.kind);
    if (!msg) return;
    if (msg.kind === 'ready') {
      if (failed || opened) return;
      clearTimeout(readyTimer);
      void openDoc();
    } else if (msg.kind === 'failed') {
      // 열기 전이면 다른 화면으로, 연 뒤면(고친 내용이 있을 수 있다) 안내만
      if (opened) notice(NOTICES.appError);
      else fail(String(msg.data.error ?? ''));
    } else if (msg.kind === 'dirty') {
      // 읽기 전용이면 보호 명령이 낸 dirty:true까지 전부 무시 — 대신 엑셀은 되돌리기로 풀린 보호를 다시 건다
      if (gridGuard) void reprotect();
      else if (canSave && msg.data.dirty === true) {
        unsaved = true;
        setSave('dirty');
      }
    } else {
      const decision = saveDecision(msg.data, canSave);
      if (decision === 'readonly') notice(NOTICES.readOnlyIgnored);
      else if (decision === 'saveas') notice(NOTICES.saveAs);
      else if (decision === 'empty') setSave('failed', '빈 문서');
      else void upload(msg.data.bytes as ArrayBuffer);
    }
  });
  window.addEventListener('beforeunload', (event) => {
    if (canSave && unsaved) event.preventDefault();
  });
}

boot().catch((error) => {
  console.error('앱 화면을 열지 못했습니다:', error);
  el('doc-title').textContent = '문서를 열 수 없습니다.';
});
