/**
 * 통합 보기 껍데기(public/view.html) 입구 — 설계 docs/superpowers/specs/2026-11-12-unified-view-shell-design.md.
 * 쪽 문서(워드·PPT·PDF 등, 한글 빼고)를 같은 모양으로 보여 주고, 편집은 형식별 편집기로 넘긴다. 엑셀 격자·도면·한글은
 * 지금 보기 화면(viewer.html)으로 넘긴다(서버 viewerUrl 기준).
 */
import { renderPresenceNames } from './presence.js';
import { applyWatermark } from './watermark.js';
import type { ShellCtx } from './view/ctx.js';
import { chooseBodyKind, createEscapeChain, pathOf } from './view/logic.js';
import { PagesBody, type PageSource } from './view/pagesBody.js';
import { PngSource } from './view/pngSource.js';
import { setupToolbar } from './view/toolbar.js';
import type { ViewInfo } from './view/types.js';

const token = new URLSearchParams(location.search).get('token') ?? '';
const tokenQuery = `token=${encodeURIComponent(token)}`;
const PRESENCE_HEARTBEAT_MS = 8000;
const NOTICE_MS = 8000;

let readyPosted = false;
/** 부모 창(향기랩·연동 제품)에 "다 그렸다" — 편집기들의 docview:ready와 같다 */
function postReady(): void {
  if (readyPosted || window.parent === window) return;
  readyPosted = true;
  requestAnimationFrame(() => requestAnimationFrame(() => window.parent.postMessage({ type: 'docview:ready' }, '*')));
}

function showError(message: string): void {
  (document.getElementById('toolbar') as HTMLElement).hidden = true;
  (document.getElementById('pages') as HTMLElement).hidden = true;
  (document.getElementById('error-screen') as HTMLElement).hidden = false;
  (document.getElementById('error-message') as HTMLElement).textContent = message;
  postReady();
}

let noticeTimer: ReturnType<typeof setTimeout> | null = null;
function notice(text: string): void {
  const el = document.getElementById('view-notice') as HTMLElement;
  el.textContent = text;
  el.hidden = text === '';
  if (noticeTimer !== null) clearTimeout(noticeTimer);
  noticeTimer = text === '' ? null : setTimeout(() => (el.hidden = true), NOTICE_MS);
}

async function fetchInfo(): Promise<ViewInfo> {
  let res: Response;
  try {
    res = await fetch(`/api/view/info?${tokenQuery}`);
  } catch {
    throw new Error('서버에 연결할 수 없습니다.');
  }
  const json = (await res.json().catch(() => null)) as { success?: boolean; data?: ViewInfo; error?: string } | null;
  if (!res.ok || !json?.success || !json.data) throw new Error(json?.error || '문서를 열 수 없습니다.');
  return json.data;
}

/** "지금 이 문서를 보고 있는 사람" — 지금 보기 화면과 같은 8초 하트비트(/api/view/presence), 그림은 presence.ts */
function startPresence(): void {
  const container = document.getElementById('presence-list');
  if (!container) return;
  const sessionId = globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random()}`;
  const beat = async () => {
    try {
      const res = await fetch(`/api/view/presence?${tokenQuery}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ sessionId }),
      });
      const json = (await res.json()) as { success?: boolean; data?: { viewers: string[] } };
      if (json.success && json.data) renderPresenceNames(container, json.data.viewers);
    } catch {
      // 접속자 표시 실패는 열람을 막지 않는다
    }
  };
  void beat();
  setInterval(() => void beat(), PRESENCE_HEARTBEAT_MS);
}

/** 우클릭·저장/인쇄 단축키 막기는 억지력일 뿐 — 실제 권한은 서버가 지킨다(viewer.js setupDeterrents와 같다) */
function setupDeterrents(ctx: ShellCtx, runEscape: () => boolean): void {
  document.addEventListener('contextmenu', (event) => {
    // 글자 선택이 되는 화면(pdf.js + 복사 권한)에서는 글자 층의 오른쪽 단추(복사)를 막지 않는다
    if (ctx.body.textSelectable && (event.target as Element | null)?.closest?.('.pdf-text-layer')) return;
    event.preventDefault();
  });
  document.addEventListener('keydown', (event) => {
    const key = event.key.toLowerCase();
    if ((event.ctrlKey || event.metaKey) && (key === 'p' || key === 's')) event.preventDefault();
    if (key === 'escape') runEscape();
  });
}

const FALLBACK_NOTICE = 'PDF를 바로 그리지 못해 서버 그림으로 보여 줍니다 — 글자 선택은 되지 않습니다.';

/** 권한에 맞는 원천 — 내려받기 권한이 있으면 pdf.js(동적 import), 받지 못하면 PNG와 한 줄 안내 */
async function openSource(info: ViewInfo, png: () => PageSource): Promise<PageSource> {
  if (chooseBodyKind(info.permissions) === 'pages-png') return png();
  try {
    const { openPdfjsSource } = await import('./view/pdfjsSource.js');
    return await openPdfjsSource(tokenQuery, info.permissions.canCopy, new AbortController().signal);
  } catch (error) {
    console.warn('pdf.js로 열지 못해 서버 그림으로 보여 줍니다:', error);
    notice(FALLBACK_NOTICE);
    return png();
  }
}

async function boot(): Promise<void> {
  if (!token) {
    showError('열람 토큰이 없습니다. 발급받은 링크로 접속해 주세요.');
    return;
  }
  let info: ViewInfo;
  try {
    info = await fetchInfo();
  } catch (error) {
    showError(error instanceof Error ? error.message : '문서를 열 수 없습니다.');
    return;
  }
  // 이 껍데기가 맡지 않는 형식은 지금 보기 화면으로(같은 조회 문자열) — 기준은 서버 viewerUrl(viewPathFor)
  if (pathOf(info.viewerUrl) !== '/view.html') {
    location.replace(`/viewer.html${location.search}`);
    return;
  }
  (document.getElementById('doc-title') as HTMLElement).textContent = info.name;
  document.title = `DocView — ${info.name}`;
  applyWatermark(info.viewer);
  startPresence();

  const png = () => new PngSource(tokenQuery, info.pageCount, info.pages);
  const body = new PagesBody(await openSource(info, png), info.pages);
  const escape = createEscapeChain();
  const ctx: ShellCtx = { token, tokenQuery, info, body, onEscape: escape.add, notice };
  setupDeterrents(ctx, escape.run);
  body.onDraw(() => postReady());
  // pdf.js로 그리다 실패하면(깨진 파일·휴대폰 메모리) 같은 자리에 서버 PNG로(설계 §4·§7)
  body.onDrawError(() => {
    if (body.kind !== 'pages-pdfjs') return;
    notice(FALLBACK_NOTICE);
    body.replaceSource(png());
  });
  await body.mount(document.getElementById('pages') as HTMLElement);
  setupToolbar(ctx);
}

boot().catch((error) => {
  console.error('보기 화면을 열지 못했습니다:', error);
  showError('문서를 열 수 없습니다.');
});
