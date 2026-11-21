/**
 * 통합 보기 껍데기(public/view.html) 입구 — 설계 docs/superpowers/specs/2026-11-12-unified-view-shell-design.md.
 * 쪽 문서(워드·PPT·PDF 등, 한글 빼고)를 같은 모양으로 보여 주고, 편집은 형식별 편집기로 넘긴다. 엑셀 격자·도면·한글은
 * 지금 보기 화면(viewer.html)으로 넘긴다(서버 viewerUrl 기준).
 */
import { renderPresenceNames } from './presence.js';
import { applyWatermark } from './watermark.js';
import type { ShellCtx } from './view/ctx.js';
import { chooseBodyKind, createEscapeChain, editHref, isArtcraftUnavailable, pathOf } from './view/logic.js';
import type { ArtcraftSource } from './view/artcraftSource.js';
import { PagesBody, type PageSource } from './view/pagesBody.js';
import { PngSource } from './view/pngSource.js';
import { setupToolbar } from './view/toolbar.js';
import { setupThumbs } from './view/thumbs.js';
import { setupSearch } from './view/search.js';
import { setupFileActions } from './view/fileActions.js';
import { setupMeasure } from './view/measure.js';
import { setupMarkup } from './view/markup.js';
import { setupEditButton } from './view/editButton.js';
import { waitUntilPrepared } from './view/prepare.js';
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

function showError(message: string, editorUrl?: string): void {
  (document.getElementById('toolbar') as HTMLElement).hidden = true;
  (document.getElementById('pages') as HTMLElement).hidden = true;
  (document.getElementById('error-screen') as HTMLElement).hidden = false;
  (document.getElementById('error-message') as HTMLElement).textContent = message;
  // 쪽을 못 그려도 편집기는 원본으로 열린다(설계 §7)
  const edit = document.getElementById('error-edit') as HTMLAnchorElement;
  if (editorUrl) {
    edit.href = editHref(editorUrl, location.search);
    edit.hidden = false;
  }
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
/** 암호 PDF — 이 토큰으로는 브라우저에서 풀 수 없다(비밀번호 풀기는 PDF 편집기 권한) */
const LOCKED_NOTICE = '암호가 걸린 PDF라 서버 그림으로 보여 줍니다 — 글자 선택은 되지 않습니다.';
/** 아트크래프트(브라우저 엔진)가 열지 못했거나 그리다 멈췄다(설계 §4.2) */
const ARTCRAFT_FALLBACK_NOTICE = '브라우저 엔진이 이 문서를 열지 못해 서버 변환으로 보여 줍니다';

function showEngineBadge(info: ViewInfo, on: boolean): void {
  const badge = document.getElementById('engine-badge') as HTMLElement;
  badge.hidden = !on;
  if (on) badge.textContent = info.artcraft?.experimental ? '브라우저 엔진(아트크래프트) · 실험' : '브라우저 엔진(아트크래프트)';
}

/** 처음 원천 — 아트크래프트 토큰이면 브라우저 엔진(동적 import), 실패하면 안내하고 서버 변환. 서버가 거절 이유를 줬으면 그 안내 */
async function firstSource(info: ViewInfo, libreSource: () => Promise<PageSource>): Promise<PageSource> {
  if (info.engine === 'artcraft' && info.artcraft) {
    try {
      const { openArtcraftSource } = await import('./view/artcraftSource.js');
      const source = await openArtcraftSource(info);
      showEngineBadge(info, true);
      // verify:artcraft가 CPU×4를 환산하려고 읽는다(Worker 안 시간)
      (document.getElementById('engine-badge') as HTMLElement).dataset.workerTiming = JSON.stringify(source.timing);
      return source;
    } catch (error) {
      console.warn('브라우저 엔진으로 열지 못해 서버 변환으로 보여 줍니다:', error);
      notice(ARTCRAFT_FALLBACK_NOTICE);
    }
  } else if (info.engineReason) {
    notice(info.engineReason);
  }
  return libreSource();
}

/** 권한에 맞는 원천 — 내려받기 권한이 있으면 pdf.js(동적 import), 받지 못하면 PNG와 한 줄 안내 */
async function openSource(info: ViewInfo, png: () => PageSource): Promise<PageSource> {
  if (chooseBodyKind(info.permissions) === 'pages-png') return png();
  try {
    const { openPdfjsSource } = await import('./view/pdfjsSource.js');
    return await openPdfjsSource(tokenQuery, info.permissions.canCopy, new AbortController().signal);
  } catch (error) {
    console.warn('pdf.js로 열지 못해 서버 그림으로 보여 줍니다:', error);
    notice(error instanceof Error && error.message === '423' ? LOCKED_NOTICE : FALLBACK_NOTICE);
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
  setupEditButton({ info });
  const png = () => new PngSource(tokenQuery, info.pageCount, info.pages);
  /** 서버 변환 원천 — 첫 변환 중(editReady)이면 "쪽을 준비하는 중"을 기다린다. 아트크래프트는 이것을 부르지 않으므로 변환을 기다리지 않는다 */
  const libreSource = async (): Promise<PageSource> => {
    if (info.processing) {
      postReady();
      await waitUntilPrepared(tokenQuery);
      info = await fetchInfo();
    }
    return openSource(info, png);
  };
  let first: PageSource;
  try {
    first = await firstSource(info, libreSource);
  } catch (error) {
    showError(error instanceof Error ? error.message : '문서를 열 수 없습니다.', info.editorUrl);
    return;
  }
  const body = new PagesBody(first, info.pages);
  const escape = createEscapeChain();
  const ctx: ShellCtx = { token, tokenQuery, info, body, onEscape: escape.add, notice };
  setupDeterrents(ctx, escape.run);
  body.onDraw(() => postReady());
  // pdf.js로 그리다 실패하면(깨진 파일·휴대폰 메모리) 같은 자리에 서버 PNG로(설계 §4·§7).
  // 아트크래프트가 멈추면(ArtcraftUnavailable: Worker 오류·20초 초과·같은 쪽 연속 실패) 서버 변환 원천으로(설계 2026-11-21 §4.2) —
  // 한 번 실패한 쪽은 원천이 다음 그리기에 다시 그린다. 여러 쪽이 함께 실패해도 한 번만
  let leavingArtcraft = false;
  const leaveArtcraft = (error: unknown): void => {
    if (body.kind !== 'pages-artcraft' || !isArtcraftUnavailable(error) || leavingArtcraft) return;
    leavingArtcraft = true;
    notice(ARTCRAFT_FALLBACK_NOTICE);
    showEngineBadge(info, false);
    libreSource().then(
      (next) => body.replaceSource(next),
      (error) => showError(error instanceof Error ? error.message : '문서를 열 수 없습니다.', info.editorUrl),
    );
  };
  if (first.kind === 'pages-artcraft') (first as ArtcraftSource).onDead = leaveArtcraft;
  body.onDrawError((error) => {
    if (body.kind === 'pages-pdfjs') {
      notice(FALLBACK_NOTICE);
      body.replaceSource(png());
      return;
    }
    leaveArtcraft(error);
  });
  await body.mount(document.getElementById('pages') as HTMLElement);
  // 떠날 때 pdf.js 문서·워커를 돌려준다(bfcache 복귀 때는 쓰지 않으므로 persisted면 둔다)
  window.addEventListener('pagehide', (event) => {
    if (!event.persisted) {
      for (const dispose of disposers) dispose();
      ctx.body.destroy();
    }
  });
  const disposers = [setupToolbar(ctx), setupThumbs(ctx)];
  setupSearch(ctx);
  setupFileActions(ctx);
  setupMeasure(ctx);
  setupMarkup(ctx);
}

boot().catch((error) => {
  console.error('보기 화면을 열지 못했습니다:', error);
  showError('문서를 열 수 없습니다.');
});
