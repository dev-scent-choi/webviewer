'use strict';

const ZOOM_STEPS = [50, 67, 80, 100, 125, 150, 200, 250, 300, 400, 500];
const DEFAULT_ZOOM_INDEX = 3;

const token = new URLSearchParams(location.search).get('token');
const pagesContainer = document.getElementById('pages');
const svgView = document.getElementById('svg-view');
const svgScroll = document.getElementById('svg-scroll');
const svgCanvas = document.getElementById('svg-canvas');
/** 'pages' | 'grid' | 'svg' — 줌/회전 버튼이 문서 종류에 따라 서로 다른 렌더링(래스터 페이지 vs
 * 벡터 svg)에 적용돼야 해서 init()이 이르게 채워두고 여러 setup 함수가 같이 참조한다. */
let currentViewKind = 'pages';
/** svg-canvas에 처음 주입했을 때(확대 1배 상태)의 실제 렌더 크기(px) — 이후 확대/맞춤 계산의
 * 기준이 된다. transform:scale()이 걸린 뒤엔 getBoundingClientRect()가 그 배율까지 반영해
 * 버리므로, 주입 직후 딱 한 번만 재서 고정해 둔다. */
let svgNaturalSize = null;
const docTitle = document.getElementById('doc-title');
const pageJumpInput = document.getElementById('page-jump-input');
const pageJumpTotal = document.getElementById('page-jump-total');
const gridFormulaRefEl = document.getElementById('grid-formula-ref');
const gridFormulaValueEl = document.getElementById('grid-formula-value');
const gridZoomHostEl = document.getElementById('grid-zoom-host');
const gridZoomLevelEl = document.getElementById('grid-zoom-level');
const gridZoomSliderTrackEl = document.getElementById('grid-zoom-slider-track');
const gridZoomSliderFillEl = document.getElementById('grid-zoom-slider-fill');
const gridZoomSliderHandleEl = document.getElementById('grid-zoom-slider-handle');
const zoomLevelLabel = document.getElementById('zoom-level');
const printBtn = document.getElementById('print-btn');
const downloadBtn = document.getElementById('download-btn');
const downloadAsBtn = document.getElementById('download-as-btn');

function escapeXmlText(value) {
  return value.replace(/[<>&'"]/g, (ch) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' })[ch]);
}

/** "2026-09-06 23:30:11" — toISOString은 UTC라 화면을 보는 사람의 실제 시계와 어긋나 보이므로
 * 로컬 시각 그대로 초 단위까지 조립한다. */
function formatWatermarkTimestamp(date) {
  const pad = (n) => String(n).padStart(2, '0');
  return (
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ` +
    `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`
  );
}

/** 화면 열람용 워터마크 — 인쇄용(buildPrintPdf)과 같은 "열람자 이름 + 시각" 문구를, PDF에 굽는 대신
 * 각 페이지 위에 CSS 배경으로 반복해서 깐다. 토큰에 서명된(위조 불가) viewer 이름을 그대로 쓰므로
 * "이 화면을 캡처/유출하면 누가 몇 시에 봤었는지가 화면에 이미 찍혀 있다"는 억지력이 목적이며,
 * 인쇄 시 워터마크(서버가 PDF 픽셀에 실제로 굽는 것)를 대체하지 않는다. 이 페이지를 연 시점의
 * 스냅샷이라 인쇄 워터마크처럼 이후로 실시간으로 흐르진 않는다. */
function applyWatermark(viewerName) {
  if (!viewerName) return;
  const text = `${viewerName}  ${formatWatermarkTimestamp(new Date())}`;
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="300" height="180">` +
    `<text x="150" y="100" transform="rotate(-28 150 100)" text-anchor="middle" ` +
    `font-family="sans-serif" font-size="14" fill="rgba(0,0,0,0.16)">${escapeXmlText(text)}</text>` +
    `</svg>`;
  document.documentElement.style.setProperty('--watermark-bg', `url("data:image/svg+xml,${encodeURIComponent(svg)}")`);
}

/** 페이지/그리드 공용 상단 위치 표시. pages 뷰에서는 편집 가능(페이지 이동), grid 뷰에서는
 * 읽기 전용(현재 보이는 행 범위 표시)으로 쓴다. */
function setPageIndicator(current, total, { editable = true, suffix = '' } = {}) {
  if (!pageJumpInput || !pageJumpTotal) return;
  pageJumpInput.value = String(current);
  pageJumpInput.readOnly = !editable;
  pageJumpTotal.textContent = `의 ${total}${suffix}`;
}

let zoomIndex = DEFAULT_ZOOM_INDEX;
let pageCount = 0;
let permissions = { canPrint: false, canDownload: false };
let docName = 'document';
let pageDimensions = null; // [{w,h}] — 로딩 전 프레임 크기 고정용

/** content-length 기반 진행률과 함께 응답 본문을 받아온다. */
async function fetchWithProgress(url, onProgress) {
  const res = await fetch(url);
  if (!res.ok) {
    const json = await res.json().catch(() => null);
    throw new Error((json && json.error) || `요청 실패 (HTTP ${res.status})`);
  }
  const total = Number(res.headers.get('content-length')) || 0;
  const reader = res.body.getReader();
  const chunks = [];
  let received = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    received += value.length;
    onProgress(received, total);
  }
  return chunks;
}

function formatProgress(received, total) {
  if (total > 0) return `${Math.round((received / total) * 100)}%`;
  return `${(received / 1048576).toFixed(1)}MB`;
}

/** 버튼을 진행률 표시기로 전환했다가 끝나면 복원한다. */
async function runButtonTask(button, progressPrefix, task) {
  // .btn-label이 있는(텍스트 포함) 버튼은 그 안에 진행률을 쓰고, 아이콘만 있는 버튼은 title
  // 툴팁에 써서 아이콘 자체가 진행률 텍스트로 덮여 사라지지 않게 한다.
  const labelEl = button.querySelector('.btn-label');
  const original = labelEl ? labelEl.textContent : button.title;
  const setStatus = (text) => {
    if (labelEl) labelEl.textContent = text;
    else button.title = text;
  };
  button.disabled = true;
  try {
    await task((received, total) => {
      setStatus(`${progressPrefix} ${formatProgress(received, total)}`);
    });
    setStatus(original);
  } catch (error) {
    setStatus('실패');
    setTimeout(() => setStatus(original), 2500);
    throw error;
  } finally {
    button.disabled = false;
  }
}

function showError(message) {
  document.getElementById('toolbar').hidden = true;
  pagesContainer.hidden = true;
  document.getElementById('error-screen').hidden = false;
  document.getElementById('error-message').textContent = message;
}

function applyZoom() {
  const percent = ZOOM_STEPS[zoomIndex];
  zoomLevelLabel.textContent = `${percent}%`;
  pagesContainer.style.setProperty('--page-width', `${percent}%`);
  document.getElementById('zoom-out').disabled = zoomIndex === 0;
  document.getElementById('zoom-in').disabled = zoomIndex === ZOOM_STEPS.length - 1;
  scheduleQualityReload();
}

/** 도면(svg) 뷰 전용 — 래스터 재요청이 필요한 applyZoom()과 달리 인라인 svg는 그냥 CSS
 * transform으로 확대/회전한다(항상 벡터라 어떤 배율에서도 흐려지지 않는다). */
function applySvgTransform() {
  const percent = ZOOM_STEPS[zoomIndex];
  svgCanvas.style.transform = `scale(${percent / 100}) rotate(${rotateDeg}deg)`;
  zoomLevelLabel.textContent = `${percent}%`;
  document.getElementById('zoom-out').disabled = zoomIndex === 0;
  document.getElementById('zoom-in').disabled = zoomIndex === ZOOM_STEPS.length - 1;
}

const tokenQuery = `token=${encodeURIComponent(token)}`;

/* ==================== 접속자 표시("지금 이 문서를 보고 있는 사람") ====================
 * docx/xlsx 동시편집(Yjs awareness)과 색·이니셜 로직을 맞추려고 src/client/presence.ts와 값을
 * 동일하게 맞춘 순수 로직만 복제했다 — viewer.js는 번들 대상이 아닌 순수 JS라 import할 수 없다. */

const PRESENCE_COLORS = ['#f76707', '#2f9e44', '#1971c2', '#ae3ec9', '#e8590c', '#0c8599'];
const PRESENCE_HEARTBEAT_MS = 8000;
const CJK_FIRST_CHAR = /[ㄱ-힝一-鿿]/;

function presenceHashString(value) {
  let hash = 0;
  for (let i = 0; i < value.length; i += 1) {
    hash = (hash * 31 + value.charCodeAt(i)) | 0;
  }
  return Math.abs(hash);
}

function pickPresenceColorForName(name) {
  return PRESENCE_COLORS[presenceHashString(name) % PRESENCE_COLORS.length];
}

function initialsOf(name) {
  const trimmed = name.trim();
  if (!trimmed) return '?';
  if (CJK_FIRST_CHAR.test(trimmed[0])) return trimmed[0];
  const parts = trimmed.split(/\s+/);
  if (parts.length >= 2) return (parts[0][0] + parts[1][0]).toUpperCase();
  return trimmed.slice(0, 2).toUpperCase();
}

const presenceSessionId =
  (window.crypto && crypto.randomUUID && crypto.randomUUID()) || `${Date.now()}-${Math.random()}`;

function renderPresenceList(viewers) {
  const container = document.getElementById('presence-list');
  if (!container) return;
  container.replaceChildren(
    ...viewers.map((name) => {
      const avatar = document.createElement('span');
      avatar.className = 'presence-badge';
      avatar.style.setProperty('--presence-color', pickPresenceColorForName(name));
      avatar.textContent = initialsOf(name);
      avatar.title = name;
      return avatar;
    }),
  );
}

async function sendPresenceHeartbeat() {
  try {
    const data = await fetchJsonView(`/api/view/presence?${tokenQuery}`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ sessionId: presenceSessionId }),
    });
    renderPresenceList(data.viewers);
  } catch {
    /* 접속자 표시 실패는 열람 자체를 막지 않는다 */
  }
}

function setupPresence() {
  sendPresenceHeartbeat();
  setInterval(sendPresenceHeartbeat, PRESENCE_HEARTBEAT_MS);
}

/* 진행 중인 페이지·그리드 로딩을 뒤로가기 시 즉시 취소해 네비게이션 지연을 없앤다 */
let viewAbort = new AbortController();

function isAbort(error) {
  return error && error.name === 'AbortError';
}

/**
 * 확대 수준 × 디스플레이 픽셀 밀도에 맞는 서버 렌더링 배율.
 * 고해상도(Windows 배율 125~150%) 모니터에서는 CSS 픽셀보다 실제 픽셀이
 * 많아 배율을 더 올려야 원본처럼 선명하게 보인다.
 */
function desiredScale() {
  const percent = ZOOM_STEPS[zoomIndex];
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  const scale = Math.ceil((percent / 100) * 2 * dpr);
  return Math.min(16, Math.max(2, scale));
}

/* 대용량 페이지 이미지를 동시에 너무 많이 받으면 메인 스레드가 밀려
   스크롤·버튼 반응이 늦어지므로 동시 로딩 수를 제한한다 */
const MAX_CONCURRENT_LOADS = 3;
let activeLoads = 0;
const loadWaiters = [];

async function withLoadSlot(task) {
  if (activeLoads >= MAX_CONCURRENT_LOADS) {
    await new Promise((resolve) => loadWaiters.push(resolve));
  }
  activeLoads += 1;
  try {
    return await task();
  } finally {
    activeLoads -= 1;
    const next = loadWaiters.shift();
    if (next) next();
  }
}

async function loadPageInto(canvas, pageNumber) {
  const scale = desiredScale();
  canvas.dataset.requestedScale = String(scale);
  await withLoadSlot(async () => {
    if (viewAbort.signal.aborted) return;
    let res;
    let bitmap;
    try {
      res = await fetch(`/api/view/pages/${pageNumber}?${tokenQuery}&scale=${scale}`, {
        signal: viewAbort.signal,
      });
      if (!res.ok) {
        canvas.closest('.page-frame').classList.add('page-error');
        return;
      }
      bitmap = await createImageBitmap(await res.blob());
    } catch (error) {
      if (isAbort(error)) return; // 뒤로가기 등으로 취소됨
      canvas.closest('.page-frame').classList.add('page-error');
      return;
    }
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    canvas.getContext('2d').drawImage(bitmap, 0, 0);
    bitmap.close();
    canvas.dataset.loadedScale = String(scale);
    const frame = canvas.closest('.page-frame');
    frame.style.aspectRatio = `${canvas.width} / ${canvas.height}`;
    frame.classList.remove('page-loading');
    frame.classList.add('page-ready');
  });
}

/**
 * 화면에서 멀리 벗어난 페이지의 비트맵을 해제한다.
 * 페이지가 많을 때 GPU 메모리가 쌓여 화질 저하·뒤로가기 지연이
 * 생기는 것을 막는다. 종횡비를 고정해 스크롤 위치는 유지된다.
 */
function releaseCanvas(canvas) {
  if (!canvas.dataset.loadedScale) return;
  const frame = canvas.closest('.page-frame');
  if (canvas.width > 0) {
    frame.style.aspectRatio = `${canvas.width} / ${canvas.height}`;
  }
  canvas.width = 0;
  canvas.height = 0;
  delete canvas.dataset.loadedScale;
  delete canvas.dataset.requestedScale;
  frame.classList.remove('page-ready');
  frame.classList.add('page-loading');
}

function ensureLoaded(canvas) {
  const target = desiredScale();
  const requested = Number(canvas.dataset.requestedScale || 0);
  if (requested >= target) return;
  loadPageInto(canvas, Number(canvas.dataset.page));
}

let qualityReloadTimer = null;

/** 확대 후 화면 근처에 보이는 페이지들을 더 높은 배율로 다시 렌더링한다. */
function scheduleQualityReload() {
  if (qualityReloadTimer !== null) clearTimeout(qualityReloadTimer);
  qualityReloadTimer = setTimeout(() => {
    qualityReloadTimer = null;
    const nearViewport = window.innerHeight * 2;
    for (const canvas of pagesContainer.querySelectorAll('canvas')) {
      if (!canvas.dataset.loadedScale) continue; // 화면 밖(해제됨)은 접근 시 로딩된다
      const rect = canvas.getBoundingClientRect();
      if (rect.bottom < -nearViewport || rect.top > window.innerHeight + nearViewport) continue;
      ensureLoaded(canvas);
    }
  }, 350);
}

function buildPages() {
  // 화면 근처(±1200px)에 들어오면 로딩, 멀어지면 비트맵 해제 (가상화)
  const observer = new IntersectionObserver((entries) => {
    for (const entry of entries) {
      const canvas = entry.target.querySelector('canvas');
      if (entry.isIntersecting) {
        ensureLoaded(canvas);
      } else {
        releaseCanvas(canvas);
      }
    }
  }, { rootMargin: '1200px 0px' });

  const frames = [];
  for (let page = 1; page <= pageCount; page += 1) {
    const frame = document.createElement('div');
    frame.className = 'page-frame page-loading';
    // 페이지 크기를 미리 알면 로딩 전에도 프레임 높이가 정확해 스크롤 점프가 어긋나지 않는다
    const dims = pageDimensions && pageDimensions[page - 1];
    if (dims && dims.w > 0) {
      frame.style.aspectRatio = `${dims.w} / ${dims.h}`;
      frame.style.minHeight = '0';
    }
    const canvas = document.createElement('canvas');
    canvas.dataset.page = String(page);
    const label = document.createElement('span');
    label.className = 'page-number';
    label.textContent = String(page);
    frame.append(canvas, label);
    frames.push(frame);
    if (page === 1) {
      // 나머지 프레임을 다 만들 때까지 기다리지 않고 1페이지 이미지 요청을 먼저 흘려보낸다 —
      // IntersectionObserver 콜백은 이 함수가 끝나고 다음 틱에야 배달되므로, 그대로 두면
      // "전체 프레임 구성"이 "1페이지 요청"보다 항상 먼저 끝나야 하는 불필요한 직렬 의존이
      // 생긴다. ensureLoaded는 이미 requestedScale 가드가 있어(loadPageInto가 그 값을
      // 동기적으로 세팅) 잠시 뒤 옵저버가 같은 캔버스에 대해 또 호출해도 중복 요청은 없다.
      ensureLoaded(canvas);
    }
    observer.observe(frame);
  }
  pagesContainer.replaceChildren(...frames);
}

function trackCurrentPage() {
  const observer = new IntersectionObserver((entries) => {
    const visible = entries.find((entry) => entry.isIntersecting);
    if (visible) {
      const page = visible.target.querySelector('canvas').dataset.page;
      if (document.activeElement !== pageJumpInput) setPageIndicator(page, pageCount);
    }
  }, { threshold: 0.4 });
  for (const frame of pagesContainer.children) {
    observer.observe(frame);
  }
}

const INVERT_STORAGE_KEY = 'docview-invert';

function applyInvert(enabled) {
  pagesContainer.classList.toggle('pages-invert', enabled);
  document.getElementById('invert-btn').setAttribute('aria-pressed', String(enabled));
  try {
    localStorage.setItem(INVERT_STORAGE_KEY, enabled ? '1' : '0');
  } catch {
    /* 시크릿 모드 등에서 저장 실패해도 무시 */
  }
}

function navigateBack() {
  // 진행 중인 이미지·그리드 로딩을 즉시 취소해야 대형 문서에서도 뒤로가기가 바로 먹힌다
  viewAbort.abort();
  if (qualityReloadTimer !== null) clearTimeout(qualityReloadTimer);
  window.location.assign('/');
}

function setupToolbar() {
  const backBtn = document.getElementById('back-btn');
  backBtn.addEventListener('click', (event) => {
    event.preventDefault();
    navigateBack();
  });

  document.getElementById('zoom-in').addEventListener('click', () => {
    zoomIndex = Math.min(zoomIndex + 1, ZOOM_STEPS.length - 1);
    if (currentViewKind === 'svg') applySvgTransform(); else applyZoom();
  });
  document.getElementById('zoom-out').addEventListener('click', () => {
    zoomIndex = Math.max(zoomIndex - 1, 0);
    if (currentViewKind === 'svg') applySvgTransform(); else applyZoom();
  });

  const invertBtn = document.getElementById('invert-btn');
  invertBtn.addEventListener('click', () => {
    applyInvert(!pagesContainer.classList.contains('pages-invert'));
  });
  try {
    if (localStorage.getItem(INVERT_STORAGE_KEY) === '1') applyInvert(true);
  } catch {
    /* 저장소 접근 불가 시 기본값 유지 */
  }

  printBtn.disabled = !permissions.canPrint;
  printBtn.title = permissions.canPrint ? '워터마크가 포함된 출력용 PDF' : '출력 권한이 없습니다';
  if (permissions.canPrint) {
    printBtn.addEventListener('click', () => {
      const includeMarkup = document.getElementById('markup-include-checkbox')?.checked ? '&includeMarkup=1' : '';
      // 팝업 차단을 피하려면 사용자 클릭 스택 안에서 창을 먼저 열어둬야 한다
      const printWindow = window.open('', '_blank');
      runButtonTask(printBtn, '생성', async (onProgress) => {
        const chunks = await fetchWithProgress(`/api/view/print?${tokenQuery}${includeMarkup}`, onProgress);
        const blob = new Blob(chunks, { type: 'application/pdf' });
        const url = URL.createObjectURL(blob);
        if (printWindow && !printWindow.closed) {
          printWindow.location.href = url;
        } else {
          const anchor = document.createElement('a');
          anchor.href = url;
          anchor.target = '_blank';
          anchor.rel = 'noopener';
          anchor.click();
        }
        setTimeout(() => URL.revokeObjectURL(url), 60_000);
      }).catch(() => {
        if (printWindow && !printWindow.closed) printWindow.close();
      });
    });
  }

  downloadBtn.disabled = !permissions.canDownload;
  downloadBtn.title = permissions.canDownload ? '원본 파일 다운로드' : '저장 권한이 없습니다';
  if (permissions.canDownload) {
    downloadBtn.addEventListener('click', () => {
      const includeMarkup = document.getElementById('markup-include-checkbox')?.checked ? '&includeMarkup=1' : '';
      runButtonTask(downloadBtn, '저장', async (onProgress) => {
        const chunks = await fetchWithProgress(`/api/view/download?${tokenQuery}${includeMarkup}`, onProgress);
        const blob = new Blob(chunks);
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement('a');
        anchor.href = url;
        anchor.download = docName;
        anchor.click();
        setTimeout(() => URL.revokeObjectURL(url), 60_000);
      }).catch(() => {});
    });
  }

  if (downloadAsBtn) {
    downloadAsBtn.disabled = !permissions.canDownload;
    downloadAsBtn.title = permissions.canDownload ? '다른 이름으로 저장' : '저장 권한이 없습니다';
    if (permissions.canDownload) {
      downloadAsBtn.addEventListener('click', () => {
        const suggested = docName;
        const customName = window.prompt('저장할 파일 이름을 입력하세요', suggested);
        if (!customName) return;
        const includeMarkup = document.getElementById('markup-include-checkbox')?.checked ? '&includeMarkup=1' : '';
        runButtonTask(downloadAsBtn, '저장', async (onProgress) => {
          const chunks = await fetchWithProgress(`/api/view/download?${tokenQuery}${includeMarkup}`, onProgress);
          const blob = new Blob(chunks);
          const url = URL.createObjectURL(blob);
          const anchor = document.createElement('a');
          anchor.href = url;
          anchor.download = customName;
          anchor.click();
          setTimeout(() => URL.revokeObjectURL(url), 60_000);
        }).catch(() => {});
      });
    }
  }
}

/* ==================== 도구모음 반응형 "..." 메뉴 ====================
   왼쪽 도구 묶음(콘텐츠~마크업 포함)이 창 폭보다 넓어지면, 아이콘 버튼은 전부 flex:none이라
   flexbox가 줄일 수 있는 유일한 대상이 구분선(.toolbar-sep)뿐이었다 — 그 결과 구분선이 폭 0으로
   눌려 아예 안 보이고, 그러고도 자리가 모자라면 옆 zone(가운데 확대/축소 등)과 아이콘이 겹쳐
   보였다. 대신 우선순위가 낮은 도구부터 실제 버튼 엘리먼트를 이 메뉴 안으로 옮긴다(엘리먼트를
   그대로 옮기므로 이미 걸려 있던 클릭 리스너가 그대로 살아있다 — 복제/재구현 불필요). */
const TOOLBAR_OVERFLOW_PRIORITY = [
  'invert-btn',
  'measure-btn',
  'markup-include-wrap',
  'text-tool-btn',
  'markup-add-btn',
];

let toolbarOriginalOrder = null;

/** 구분선은 "양옆에 보이는 도구가 하나라도 있을 때만" 보인다 — 어떤 조합으로 도구가 접혀도
 * 맨 앞/맨 뒤나 두 구분선이 나란히 남는 경우가 생기지 않는다. */
function normalizeToolbarDividers(zone, moreWrap) {
  let sawContent = false;
  let pendingSep = null;
  for (const el of zone.children) {
    if (el.classList.contains('toolbar-sep')) {
      el.hidden = true;
      pendingSep = el;
      continue;
    }
    if (el === moreWrap || el.hidden) continue;
    if (sawContent && pendingSep) pendingSep.hidden = false;
    pendingSep = null;
    sawContent = true;
  }
}

function rebuildToolbarZone(zone, moreWrap, moreMenu, overflowIds) {
  for (const el of toolbarOriginalOrder) {
    if (overflowIds.has(el.id)) {
      moreMenu.appendChild(el);
    } else {
      zone.insertBefore(el, moreWrap);
    }
  }
  normalizeToolbarDividers(zone, moreWrap);
}

function closeToolbarMoreMenu() {
  const moreMenu = document.getElementById('toolbar-more-menu');
  const moreBtn = document.getElementById('toolbar-more-btn');
  if (!moreMenu || moreMenu.hidden) return;
  moreMenu.hidden = true;
  moreBtn.setAttribute('aria-expanded', 'false');
}

function updateToolbarOverflow() {
  const zone = document.getElementById('toolbar-tools-left');
  const moreWrap = document.getElementById('toolbar-more-wrap');
  const moreBtn = document.getElementById('toolbar-more-btn');
  const moreMenu = document.getElementById('toolbar-more-menu');
  if (!zone || !moreWrap) return;

  if (!toolbarOriginalOrder) {
    toolbarOriginalOrder = Array.from(zone.children).filter((el) => el !== moreWrap);
  }

  const overflowIds = new Set();
  rebuildToolbarZone(zone, moreWrap, moreMenu, overflowIds);
  for (const id of TOOLBAR_OVERFLOW_PRIORITY) {
    if (zone.scrollWidth <= zone.clientWidth + 1) break;
    if (!document.getElementById(id)) continue;
    overflowIds.add(id);
    rebuildToolbarZone(zone, moreWrap, moreMenu, overflowIds);
  }

  moreBtn.hidden = overflowIds.size === 0;
  if (overflowIds.size === 0) closeToolbarMoreMenu();
}

function setupToolbarOverflow() {
  const moreBtn = document.getElementById('toolbar-more-btn');
  const moreMenu = document.getElementById('toolbar-more-menu');
  if (!moreBtn || !moreMenu) return;

  moreBtn.addEventListener('click', (event) => {
    event.stopPropagation();
    const willOpen = moreMenu.hidden;
    moreMenu.hidden = !willOpen;
    moreBtn.setAttribute('aria-expanded', String(willOpen));
  });
  // 메뉴 안의 도구를 고르면(체크박스 포함) 바로 닫는다 — 고르고 나면 지도 위 클릭 등 다음
  // 동작으로 이어지는 도구가 많아, 메뉴가 계속 떠 있으면 그 클릭을 가로챈다.
  moreMenu.addEventListener('click', (event) => {
    if (event.target.closest('button, label')) setTimeout(closeToolbarMoreMenu, 0);
  });
  document.addEventListener('click', (event) => {
    if (moreMenu.hidden) return;
    if (event.target === moreBtn || moreBtn.contains(event.target) || moreMenu.contains(event.target)) return;
    closeToolbarMoreMenu();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') closeToolbarMoreMenu();
  });

  // 창 resize뿐 아니라, 왼쪽 zone에 배분되는 폭 자체가 바뀌는 모든 경우를 잡아야 해서 window
  // resize 이벤트만으로는 부족하다 — ResizeObserver는 이 zone의 실제 렌더 "박스" 폭이 바뀔 때마다
  // 다시 불러 준다.
  const zone = document.getElementById('toolbar-tools-left');
  let pending = false;
  const scheduleUpdate = () => {
    if (pending) return;
    pending = true;
    requestAnimationFrame(() => {
      pending = false;
      updateToolbarOverflow();
    });
  };
  if (zone && window.ResizeObserver) {
    new ResizeObserver(scheduleUpdate).observe(zone);
  } else {
    updateToolbarOverflow();
    window.addEventListener('resize', scheduleUpdate);
  }
  // "박스" 자체 폭(clientWidth)은 안 바뀌어도, 그 안의 글자(예: "마크업 포함" 라벨)가 나중에
  // 실제 웹폰트(Outfit)로 교체되며 폭만 넓어지는 경우가 있다 — ResizeObserver는 zone의 박스
  // 크기만 보므로 이건 못 잡는다(실제로 이 타이밍차 때문에 도구 하나가 덜 접힌 채로 굳어버리는
  // 문제가 있었다). 폰트 로딩이 끝나면 한 번 더 확실히 재계산한다.
  if (document.fonts && document.fonts.ready) {
    document.fonts.ready.then(scheduleUpdate).catch(() => {});
  }
}

/** 페이지 번호 칸에 직접 입력해서 해당 페이지로 이동한다(Edge PDF 뷰어와 동일한 동작). */
function setupPageJump() {
  if (!pageJumpInput) return;
  const jumpToTyped = () => {
    const target = Math.min(Math.max(Math.round(Number(pageJumpInput.value)), 1), pageCount || 1);
    if (!Number.isFinite(target)) {
      setPageIndicator(1, pageCount);
      return;
    }
    const frame = pagesContainer.children[target - 1];
    frame?.scrollIntoView({ block: 'start' });
    setPageIndicator(target, pageCount);
    pageJumpInput.blur();
  };
  pageJumpInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      jumpToTyped();
    } else if (event.key === 'Escape') {
      pageJumpInput.blur();
    }
  });
  pageJumpInput.addEventListener('blur', jumpToTyped);
  pageJumpInput.addEventListener('focus', () => pageJumpInput.select());
}

/** "너비에 맞추기"/"페이지에 맞추기" 토글 — 버튼은 지금 누르면 적용될 모드가 아니라, 방금
 * 누른 결과 다음 클릭에 적용될 "다른" 모드를 보여준다(참고한 PDF 뷰어와 동일한 동작). */
let zoomFitMode = 'width'; // 'width' | 'page'

function setupZoomFit() {
  const btn = document.getElementById('zoom-fit-btn');
  if (!btn) return;
  const widthIcon = btn.querySelector('.zoom-fit-icon-width');
  const pageIcon = btn.querySelector('.zoom-fit-icon-page');

  function snapToNearestStep(percent) {
    const clamped = Math.min(Math.max(percent, ZOOM_STEPS[0]), ZOOM_STEPS[ZOOM_STEPS.length - 1]);
    let nearestIndex = 0;
    let nearestDiff = Infinity;
    ZOOM_STEPS.forEach((step, i) => {
      const diff = Math.abs(step - clamped);
      if (diff < nearestDiff) {
        nearestDiff = diff;
        nearestIndex = i;
      }
    });
    return nearestIndex;
  }

  // --page-width:100%는 정의상 "컨테이너 폭 전체"라서(applyZoom() 참고), 폭 맞춤은 항상
  // 100%면 된다 — pagesContainer.clientWidth를 pageWidthPt(점 단위)로 나누는 예전 계산은
  // 서로 다른 두 단위를 섞어(퍼센트가 실제로는 "컨테이너 기준"인데 "포인트 기준"인 것처럼
  // 계산) 화면 폭을 훌쩍 넘겨 가로 스크롤이 생기는 버그가 있었다.
  function pageFitPercent() {
    // 페이지 전체(가로+세로)가 스크롤 없이 한 화면에 들어오도록, 100%(=컨테이너 폭 전체) 기준
    // 세로 높이가 뷰포트를 넘는지 확인해서, 넘으면 세로가 맞는 비율로 같이 축소한다(가로세로
    // 비율은 프레임의 aspect-ratio로 고정되므로 폭 percent 하나가 둘 다를 함께 줄인다).
    const containerWidth = pagesContainer.clientWidth - 32; // 좌우 여백 감안
    const dims = (pageDimensions && pageDimensions[0]) || { w: 612, h: 792 };
    const frameHeightAt100 = containerWidth * (dims.h / dims.w);
    const toolbarEl = document.getElementById('toolbar');
    const availableHeight = window.innerHeight - toolbarEl.offsetHeight - 32; // 위아래 여백 감안(폭 계산과 같은 관례)
    if (frameHeightAt100 <= availableHeight) return 100;
    return (availableHeight / frameHeightAt100) * 100;
  }

  // svg-canvas는 (pages와 달리) 자기 콘텐츠 크기로 자연스럽게 커지는 요소라 "100%=컨테이너 폭"
  // 지름길이 없다 — svgNaturalSize(줌 1배일 때 실측한 실제 렌더 크기)를 기준으로 직접 비율을 잰다.
  function svgWidthFitPercent() {
    const availableWidth = svgScroll.clientWidth - 32;
    const naturalWidth = svgNaturalSize?.w || availableWidth;
    return (availableWidth / naturalWidth) * 100;
  }
  function svgPageFitPercent() {
    const availableWidth = svgScroll.clientWidth - 32;
    const availableHeight = svgScroll.clientHeight - 32;
    const naturalWidth = svgNaturalSize?.w || availableWidth;
    const naturalHeight = svgNaturalSize?.h || availableHeight;
    return Math.min((availableWidth / naturalWidth) * 100, (availableHeight / naturalHeight) * 100);
  }

  btn.addEventListener('click', () => {
    let percent;
    if (currentViewKind === 'svg') {
      percent = zoomFitMode === 'width' ? svgWidthFitPercent() : svgPageFitPercent();
    } else {
      // 이 컨테이너의 1120px 읽기용 폭 제한을 없애 브라우저 창 끝까지 채운다 — 폭/페이지
      // 맞춤 둘 다 이 넓어진 폭을 기준으로 계산해야 "화면 끝까지"라는 의도가 맞는다.
      pagesContainer.classList.add('pages-edge-to-edge');
      percent = zoomFitMode === 'width' ? 100 : pageFitPercent();
    }
    zoomIndex = snapToNearestStep(percent);
    if (currentViewKind === 'svg') applySvgTransform(); else applyZoom();

    zoomFitMode = zoomFitMode === 'width' ? 'page' : 'width';
    const showingPageMode = zoomFitMode === 'page';
    // svg 엘리먼트는 .hidden 프로퍼티를 대입해도 hidden 속성에 반영되지 않는 경우가 있어(SVGElement의
    // HTMLOrForeignElement 믹스인 지원이 엔진마다 달라서) 속성을 직접 세팅/제거한다.
    widthIcon.toggleAttribute('hidden', showingPageMode);
    pageIcon.toggleAttribute('hidden', !showingPageMode);
    const label = showingPageMode ? '페이지에 맞추기' : '너비에 맞추기';
    btn.title = label;
    btn.setAttribute('aria-label', label);

    btn.setAttribute('aria-pressed', 'true');
    setTimeout(() => btn.setAttribute('aria-pressed', 'false'), 400);
  });
}

/** 화면에서만 페이지를 90도씩 돌려 본다(저장되지 않는 보기 전용 회전). */
let rotateDeg = 0;
function setupRotate() {
  const btn = document.getElementById('rotate-btn');
  if (!btn) return;
  btn.addEventListener('click', () => {
    rotateDeg = (rotateDeg + 90) % 360;
    if (currentViewKind === 'svg') {
      applySvgTransform();
    } else {
      pagesContainer.style.setProperty('--page-rotate', `${rotateDeg}deg`);
      pagesContainer.classList.toggle('pages-rotated', rotateDeg !== 0);
    }
  });
}

/** 전체화면 보기 — 브라우저 Fullscreen API로 문서 전체(<html>)를 전체화면 전환한다.
 * Esc나 브라우저 자체 UI로 빠져나가는 경우도 있어(버튼 클릭만이 아니라) fullscreenchange를
 * 구독해 아이콘/상태를 항상 실제 전체화면 여부에 맞춰 동기화한다. */
function setupFullscreen() {
  const btn = document.getElementById('fullscreen-btn');
  if (!btn) return;
  if (!document.fullscreenEnabled) {
    btn.hidden = true;
    return;
  }
  const enterIcon = btn.querySelector('.fullscreen-icon-enter');
  const exitIcon = btn.querySelector('.fullscreen-icon-exit');

  function syncState() {
    const isFullscreen = Boolean(document.fullscreenElement);
    // svg는 .hidden 프로퍼티 대입이 hidden 속성에 반영 안 되는 경우가 있어(zoom-fit-btn과 같은
    // 이유) 속성을 직접 세팅/제거한다.
    enterIcon.toggleAttribute('hidden', isFullscreen);
    exitIcon.toggleAttribute('hidden', !isFullscreen);
    const label = isFullscreen ? '전체화면 종료' : '전체화면 보기';
    btn.title = label;
    btn.setAttribute('aria-label', label);
    btn.setAttribute('aria-pressed', String(isFullscreen));
  }

  btn.addEventListener('click', () => {
    if (document.fullscreenElement) {
      document.exitFullscreen();
    } else {
      document.documentElement.requestFullscreen().catch(() => {});
    }
  });
  document.addEventListener('fullscreenchange', syncState);
  syncState();
}

/* 우클릭/단축키 차단은 억지력일 뿐이며 실제 권한 강제는 서버에서 이루어진다. */
function setupDeterrents() {
  document.addEventListener('contextmenu', (event) => event.preventDefault());
  document.addEventListener('keydown', (event) => {
    const key = event.key.toLowerCase();
    if ((event.ctrlKey || event.metaKey) && (key === 'p' || key === 's')) {
      event.preventDefault();
      if (key === 'p' && permissions.canPrint) {
        printBtn.click();
      }
    }
    // 브라우저 기본 Ctrl+F는 가상화된 문서를 찾지 못하므로 자체 검색으로 대체한다
    if ((event.ctrlKey || event.metaKey) && key === 'f') {
      event.preventDefault();
      openSearchBar();
    }
    if (key === 'f3') {
      event.preventDefault();
      if (searchState.hits.length > 0) gotoHit(searchState.current + (event.shiftKey ? -1 : 1));
      else openSearchBar();
    }
    if (key === 'escape') {
      if (measure.mode) setMeasureMode(false);
      else if (textTool.pending) commitTextInsert();
      else if (textTool.active) setTextToolActive(false);
      else if (markup.drawMode) setMarkupDrawMode(false);
      else closeSearchBar();
    }
  });
}

/* ==================== 문서 내 검색 ==================== */

const searchState = { hits: [], current: -1, kind: 'pages', lastQuery: '' };

function openSearchBar() {
  const bar = document.getElementById('search-bar');
  bar.hidden = false;
  document.getElementById('search-input').focus();
  document.getElementById('search-input').select();
}

function closeSearchBar() {
  document.getElementById('search-bar').hidden = true;
  clearSearchHighlights();
  searchState.hits = [];
  searchState.current = -1;
  updateSearchCount('');
}

function updateSearchCount(text) {
  document.getElementById('search-count').textContent = text;
  const hasHits = searchState.hits.length > 0;
  document.getElementById('search-prev').disabled = !hasHits;
  document.getElementById('search-next').disabled = !hasHits;
}

function clearSearchHighlights() {
  for (const el of document.querySelectorAll('.search-hit-overlay')) el.remove();
  gridSearchHits.clear();
  gridCurrentHit = null;
  if (typeof renderGridViewport === 'function' && gridState.sheets.length > 0) renderGridViewport();
}

async function runSearch() {
  const query = document.getElementById('search-input').value.trim();
  if (query === '' || query === searchState.lastQuery) {
    if (query === searchState.lastQuery && searchState.hits.length > 0) gotoHit(searchState.current + 1);
    return;
  }
  searchState.lastQuery = query;
  clearSearchHighlights();
  updateSearchCount('검색 중…');
  try {
    const data = await fetchJsonView(
      `/api/view/search?${tokenQuery}&q=${encodeURIComponent(query)}&limit=500`,
    );
    searchState.kind = data.kind;
    searchState.hits = data.hits;
    searchState.current = -1;
    if (data.hits.length === 0) {
      updateSearchCount('결과 없음');
      return;
    }
    if (data.kind === 'grid') {
      gridSearchHits = new Set(data.hits.map((h) => `${h.sheet}:${h.r}:${h.c}`));
    }
    gotoHit(0);
  } catch (error) {
    updateSearchCount(error.message || '검색 실패');
  }
}

function gotoHit(index) {
  if (searchState.hits.length === 0) return;
  const i = ((index % searchState.hits.length) + searchState.hits.length) % searchState.hits.length;
  searchState.current = i;
  const total = searchState.hits.length;
  const totalLabel = total >= 500 ? '500+' : String(total);
  updateSearchCount(`${i + 1} / ${totalLabel}`);
  const hit = searchState.hits[i];
  if (searchState.kind === 'grid') {
    gotoGridHit(hit);
  } else {
    gotoPageHit(hit);
  }
}

/** 모든 검색 결과를 연노랑으로, 현재 결과를 진한 주황 테두리로 표시한다 */
function renderPageHitOverlays() {
  for (const el of document.querySelectorAll('.search-hit-overlay')) el.remove();
  if (searchState.kind !== 'pages') return;
  searchState.hits.forEach((hit, index) => {
    const frame = pagesContainer.children[hit.page - 1];
    if (!frame) return;
    for (const rect of hit.rects) {
      const overlay = document.createElement('div');
      overlay.className =
        index === searchState.current ? 'search-hit-overlay current' : 'search-hit-overlay';
      overlay.style.left = `${rect.x * 100}%`;
      overlay.style.top = `${rect.y * 100}%`;
      overlay.style.width = `${Math.max(rect.w * 100, 0.5)}%`;
      overlay.style.height = `${Math.max(rect.h * 100, 0.5)}%`;
      frame.append(overlay);
    }
  });
}

/** 하이라이트가 이 크기(px)보다 작으면 읽을 수 있는 배율까지 자동 확대한다 */
const MIN_READABLE_HIT_PX = 11;
const TARGET_HIT_PX = 16;

function gotoPageHit(hit) {
  const frame = pagesContainer.children[hit.page - 1];
  if (!frame) return;
  renderPageHitOverlays();
  const current = frame.querySelector('.search-hit-overlay.current');

  // 도면처럼 결과 문자가 아주 작으면(전체 보기 기준 수 px) 자동으로 확대해 보여준다
  if (current) {
    const hitHeight = current.getBoundingClientRect().height;
    if (hitHeight > 0 && hitHeight < MIN_READABLE_HIT_PX) {
      const targetPercent = Math.min(
        ZOOM_STEPS[ZOOM_STEPS.length - 1],
        ZOOM_STEPS[zoomIndex] * (TARGET_HIT_PX / hitHeight),
      );
      const nextIndex = ZOOM_STEPS.findIndex((step) => step >= targetPercent);
      const resolved = nextIndex === -1 ? ZOOM_STEPS.length - 1 : nextIndex;
      if (resolved > zoomIndex) {
        zoomIndex = resolved;
        applyZoom();
      }
    }
  }

  centerOnElement(current ?? frame);
}

/**
 * 요소를 화면 정중앙으로 스크롤한다.
 * 확대 시 pages-container가 가로·세로 스크롤을 모두 갖게 되어 scrollIntoView가
 * 어긋나므로, 컨테이너 → 윈도우 순으로 직접 보정한다.
 */
function centerOnElement(el) {
  const rect = el.getBoundingClientRect();
  const containerRect = pagesContainer.getBoundingClientRect();

  const dx = rect.left + rect.width / 2 - (containerRect.left + pagesContainer.clientWidth / 2);
  pagesContainer.scrollLeft += dx;

  const dy = rect.top + rect.height / 2 - window.innerHeight / 2;
  const beforeTop = pagesContainer.scrollTop;
  pagesContainer.scrollTop += dy;
  const consumed = pagesContainer.scrollTop - beforeTop;
  if (Math.abs(dy - consumed) > 1) {
    window.scrollBy(0, dy - consumed);
  }
}

let gridSearchHits = new Set();
let gridCurrentHit = null;

function gotoGridHit(hit) {
  if (gridState.active !== hit.sheet) activateSheet(hit.sheet);
  gridCurrentHit = `${hit.sheet}:${hit.r}:${hit.c}`;
  const scroll = document.getElementById('grid-scroll');
  scroll.scrollTop = Math.max(0, gridState.layout.top(hit.r - 1) - scroll.clientHeight / 2);
  renderGridViewport();
}

function setupSearchBar() {
  document.getElementById('search-btn').addEventListener('click', openSearchBar);
  document.getElementById('search-close').addEventListener('click', closeSearchBar);
  document.getElementById('search-prev').addEventListener('click', () => gotoHit(searchState.current - 1));
  document.getElementById('search-next').addEventListener('click', () => gotoHit(searchState.current + 1));
  document.getElementById('search-input').addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      if (event.shiftKey) gotoHit(searchState.current - 1);
      else runSearch();
    }
  });
}

/* ==================== 그리드 뷰 (스프레드시트) ==================== */

/** 파일에 기본 행 높이가 없을 때(구버전 캐시 등) 쓰는 값 — 엑셀 기본 15pt */
const GRID_DEFAULT_ROW_H = 20;
const GRID_CHUNK = 200;
const GRID_OVERSCAN = 12;
// 세로 병합 앵커를 렌더 창 뒤로 당겨오는(renderGridViewport 참고) 범위의 상한. src/spreadsheet.ts의
// MAX_ROWS(300,000)급으로 큰 시트에서, 1행에 앵커를 둔 채 시트 전체를 덮는 병합이 있으면 이 상한이
// 없을 때 맨 아래로 스크롤하기만 해도 renderFirst가 1행까지 끌려가 수십만 개의 <tr>과 수천 건의
// 청크 요청을 한 번에 만들어낸다 — GRID_CHUNK(200행)의 5배로 끊어, 이 확장 때문에 한 번에 추가로
// 요청되는 청크가 최대 5개를 넘지 않게 한다. 실전 문서의 병합은 거의 항상 이 안에 들어오고, 이
// 상한을 넘는 병합은 renderGridViewport에서 "성능 우선" 쪽으로 저하 렌더링된다.
const GRID_MERGE_EXPAND_CAP = 1000;
const GRID_ROWNUM_W = 52;
// 실제 데이터 범위 너머로도 빈 칸을 넉넉히 그려서 엑셀처럼 "빈 시트도 격자가 있는" 느낌을 준다
// (열람 전용이라 편집기처럼 입력 시 확장되지는 않지만, 화면 톤은 편집기와 맞춘다).
const GRID_ROW_PADDING = 30;
const GRID_COL_PADDING = 8;
const GRID_MIN_DISPLAY_ROWS = 40;
const GRID_MIN_DISPLAY_COLS = 14;
const GRID_DEFAULT_COL_WIDTH = 64;
/** 엑셀 기본 글꼴(맑은 고딕) — 칸 글꼴이 없는 PC에서도 엑셀과 글자 폭이 비슷하게(grid-editor.ts와 같은 값) */
const GRID_FONT_STACK = "'맑은 고딕','Malgun Gothic',Calibri,sans-serif";

/**
 * 행마다 높이가 다른(숨긴 행은 0) 시트의 세로 위치 계산 — src/xlsx/sheetGeometry.ts의 RowLayout과
 * 같은 알고리즘(이 파일은 번들되지 않는 순수 JS라 옮겨 둔다). 행 번호는 0부터.
 */
class GridRowLayout {
  constructor(defaultPx, runs) {
    this.defaultPx = defaultPx;
    this.starts = [];
    this.ends = [];
    this.heights = [];
    this.tops = [];
    let delta = 0;
    for (const [start, len, px] of runs) {
      this.starts.push(start);
      this.ends.push(start + len);
      this.heights.push(px);
      this.tops.push(start * defaultPx + delta);
      delta += len * (px - defaultPx);
    }
  }
  segBefore(row) {
    let lo = 0;
    let hi = this.starts.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.starts[mid] <= row) lo = mid + 1;
      else hi = mid;
    }
    return lo - 1;
  }
  height(row) {
    const i = this.segBefore(row);
    return i >= 0 && row < this.ends[i] ? this.heights[i] : this.defaultPx;
  }
  top(row) {
    const i = this.segBefore(row);
    if (i < 0) return row * this.defaultPx;
    const s = this.starts[i];
    const e = this.ends[i];
    const h = this.heights[i];
    if (row < e) return this.tops[i] + (row - s) * h;
    return this.tops[i] + (e - s) * h + (row - e) * this.defaultPx;
  }
  rowAt(y) {
    if (y <= 0) return 0;
    let lo = 0;
    let hi = this.tops.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.tops[mid] <= y) lo = mid + 1;
      else hi = mid;
    }
    const i = lo - 1;
    if (i < 0) return Math.floor(y / this.defaultPx);
    const s = this.starts[i];
    const e = this.ends[i];
    const h = this.heights[i];
    const segEndY = this.tops[i] + (e - s) * h;
    if (h > 0 && y < segEndY) return s + Math.floor((y - this.tops[i]) / h);
    return e + Math.floor((y - segEndY) / this.defaultPx);
  }
}

function displayRowCount(sheet) {
  return Math.max(sheet.rowCount + GRID_ROW_PADDING, GRID_MIN_DISPLAY_ROWS);
}
function displayColCount(sheet) {
  return Math.max(sheet.colCount + GRID_COL_PADDING, GRID_MIN_DISPLAY_COLS);
}
function colWidthAt(sheet, c) {
  return sheet.colWidths[c - 1] ?? sheet.defaultColWidth ?? GRID_DEFAULT_COL_WIDTH;
}

const gridState = {
  sheets: [],
  styles: [],
  active: 0,
  /** 지금 시트의 행 높이 계산(GridRowLayout) — activateSheet가 만든다 */
  layout: new GridRowLayout(GRID_DEFAULT_ROW_H, []),
  /** 파일에 있는 그림(로고 등)·차트 — /api/view/xlsx-images, /api/view/xlsx-chart-images */
  images: [],
  chunks: new Map(), // `${sheet}:${chunkIdx}` -> rows | 'loading'
  mergeStarts: new Map(),
  mergeCovered: new Map(), // `${r}:${c}` -> 그 칸을 덮는 병합(m). 같은 key로 .has()만 쓰면 예전 Set과 동일하게 동작한다.
  // renderGridViewport가 매 렌더마다 새로 계산하는 "저하된"(앵커를 당겨오지 못해 개별 <td>로
  // 그려진) 병합 집합 — moveSelectionSimple/moveSelectionToEdge가 mergeCovered만 보고 "자기 칸이
  // 없는 칸"으로 오판해 건너뛰지 않도록, 가장 최근 렌더 결과를 여기 남겨 둔다. 첫 렌더 전에도
  // 참조될 수 있으니 빈 Set으로 초기화해 둔다.
  degradedMerges: new Set(),
  selected: null, // { row, col } — 1-based, 실제 시트 좌표(병합에 덮인 칸이면 앵커 좌표로 맞춰 둔다)
  /** 범위 링크(자료!B3:H3)로 고른 범위 — { r1, c1, r2, c2 } 1-based. 칸을 누르면 풀린다 */
  selRange: null,
  /** 하이퍼링크 — /api/view/xlsx-links(src/xlsx/sheetExtras.ts SheetLink, 칸 좌표 0부터). sheetLinks는 지금 시트 것만 */
  links: [],
  sheetLinks: [],
  /** 지금 시트 링크를 행 묶음(LINK_BUCKET행)별로 — 칸마다 목록 전체를 훑지 않게. 아주 긴 범위 링크는 wide에 */
  sheetLinkBuckets: new Map(),
  sheetLinksWide: [],
};

function colLetter(index) {
  let name = '';
  let n = index;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function injectGridStyles(styles) {
  const rules = styles.map((style, i) => {
    const parts = [];
    if (style.b) parts.push('font-weight:600');
    if (style.i) parts.push('font-style:italic');
    const lines = [];
    if (style.u) lines.push('underline');
    if (style.st) lines.push('line-through');
    if (lines.length > 0) parts.push(`text-decoration:${lines.join(' ')}`);
    if (style.fs) parts.push(`font-size:${style.fs}px`);
    // 칸 글꼴(돋움·굴림 등) — 없는 PC에서는 엑셀 기본 글꼴로
    if (style.ff) parts.push(`font-family:"${String(style.ff).replace(/["\<>;{}]/g, '')}",${GRID_FONT_STACK}`);
    if (style.col) parts.push(`color:${style.col}`);
    // 서식 코드가 값에 따라 정한 색([Red] 등) — 칸 글자 색보다 이긴다
    if (style.fc) parts.push(`color:${style.fc}`);
    if (style.bg) parts.push(`background:${style.bg}`);
    if (style.a === 'c') parts.push('text-align:center');
    if (style.a === 'r') parts.push('text-align:right');
    // 'l'(명시적 왼쪽 정렬)은 .grid-table td의 브라우저 기본값이 이미 왼쪽이라 생략해도 시각적
    // 차이가 없지만(.grid-table td에 다른 기본 text-align 규칙 없음, css/style.css 확인 완료),
    // extractStyle(spreadsheet.ts)이 이제 실제 파일의 명시적 왼쪽 정렬을 'l'로 읽어 들이므로
    // 의도를 명시적으로 남겨 둔다.
    if (style.a === 'l') parts.push('text-align:left');
    // 줄 바꿈·세로 맞춤은 .grid-table td의 기본값(nowrap·bottom)과 겹치므로 아래 테두리처럼
    // `.grid-table td.gsN` 특이도로 낸다(지정 없으면 엑셀처럼 아래 맞춤)
    const cellParts = [];
    if (style.w) cellParts.push('white-space:normal;word-break:break-all');
    if (style.va === 't') cellParts.push('vertical-align:top');
    if (style.va === 'm') cellParts.push('vertical-align:middle');
    // 엑셀 테두리 굵기(1=가는 선, 2=굵은 선)를 그리드 동시편집기(src/client/grid-editor.ts의
    // borderCss/applyStyleToTd)와 똑같은 규칙으로 CSS화한다: 지정된 변만 개별 border-* longhand로
    // 얹는다. 지정 없는 변은 그대로 둬서 .grid-table td의 기본 점선 격자가 그 변에서 계속 보이게
    // 한다.
    const borderParts = [];
    if (style.bt) borderParts.push(`border-top:${style.bt}px solid var(--v-text)`);
    if (style.br) borderParts.push(`border-right:${style.br}px solid var(--v-text)`);
    if (style.bb) borderParts.push(`border-bottom:${style.bb}px solid var(--v-text)`);
    if (style.bl) borderParts.push(`border-left:${style.bl}px solid var(--v-text)`);
    // 테두리 선언은 .gs{N} 단독 클래스 선택자가 아니라 `.grid-table td.gs{N}`로 낸다 — 기본 격자
    // 규칙(css/style.css의 `.grid-table td { border: 1px dotted ... }`, 클래스 1개+태그 1개 =
    // 특이도 (0,1,1))이 순수 `.gs{N}`(클래스 1개 = 특이도 (0,1,0))보다 특이도가 더 높아서, 선언
    // 순서와 무관하게 항상 이겨 버린다(먼저/나중 문제가 아니라 특이도 자체가 밀리는 문제 —
    // Playwright getComputedStyle로 직접 재현·확인함: `.gs{N}{border-top:...}`만으로는 실제
    // 화면에 전혀 반영되지 않았다). `.grid-table td.gs{N}`(클래스 2개+태그 1개 = 특이도 (0,2,1))로
    // 올려야 기본 격자 shorthand의 해당 변을 확실히 덮어쓴다. 폰트/색/정렬 등 다른 선언은 애초에
    // `.grid-table td`와 겹치는 규칙이 없어 특이도 경합이 없으므로 그대로 `.gs{N}`만으로 둔다.
    // 조건부 서식(서버가 구운 결과 — xlsx/viewerCondFormat.ts): 데이터 막대, 아이콘 집합(편집 화면 grid-editor.ts와 같은 모양)
    if (style.cb) {
      const [barColor, pct] = style.cb;
      cellParts.push(`background-image:linear-gradient(90deg,${barColor} 0%,${barColor}33 100%);background-size:${Number(pct)}% 76%;background-repeat:no-repeat;background-position:left center`);
    }
    // 테두리 선 종류·색, 무늬 채우기, 들여쓰기(x)와 글자 회전·세로쓰기(xi) — 서버가 편집 화면과 같은 함수(collab/cellBoxCss.ts)로 만든다
    if (style.x) cellParts.push(style.x);
    const cssForStyle = [`.gs${i}{${parts.join(';')}}`];
    if (style.xi) cssForStyle.push(`.grid-table td.gs${i}>.g-cell-spill{${style.xi}}`);
    if (style.ci) {
      const [ch, iconColor, hideValue] = style.ci;
      cellParts.push('position:relative');
      cssForStyle.push(`.grid-table td.gs${i}::after{content:${JSON.stringify(String(ch))};color:${iconColor};position:absolute;left:3px;bottom:1px;pointer-events:none}`);
      if (hideValue) cssForStyle.push(`.grid-table td.gs${i}>.g-cell-spill{visibility:hidden}`);
    }
    if (borderParts.length > 0 || cellParts.length > 0) cssForStyle.push(`.grid-table td.gs${i}{${[...borderParts, ...cellParts].join(';')}}`);
    return cssForStyle.join('\n');
  });
  const styleEl = document.createElement('style');
  styleEl.textContent = rules.join('\n');
  document.head.append(styleEl);
}

function buildMergeIndex(sheet) {
  gridState.mergeStarts = new Map();
  gridState.mergeCovered = new Map();
  for (const m of sheet.merges) {
    gridState.mergeStarts.set(`${m.r}:${m.c}`, m);
    for (let r = m.r; r < m.r + m.rs; r += 1) {
      for (let c = m.c; c < m.c + m.cs; c += 1) {
        if (r !== m.r || c !== m.c) gridState.mergeCovered.set(`${r}:${c}`, m);
      }
    }
  }
}

function fetchChunk(sheetIdx, chunkIdx) {
  void ensureChunkLoaded(sheetIdx, chunkIdx);
}

/** fetchChunk와 달리 청크가 실제로 준비될 때까지 기다릴 수 있다 — 이미 다른 곳(스크롤 뷰포트)이
 * 같은 청크를 불러오는 중이면 그 요청을 다시 보내지 않고 같은 Promise를 같이 기다린다. 방향키로
 * "데이터 끝까지 이동"할 때, 아직 화면에 없는(스크롤로 지나가지 않은) 구간의 칸이 비었는지 실제로
 * 확인하려면 이렇게 기다릴 수 있어야 한다. */
const chunkLoadPromises = new Map();
function ensureChunkLoaded(sheetIdx, chunkIdx) {
  const key = `${sheetIdx}:${chunkIdx}`;
  const existing = gridState.chunks.get(key);
  if (existing !== undefined && existing !== 'loading') return Promise.resolve();
  let promise = chunkLoadPromises.get(key);
  if (promise) return promise;
  gridState.chunks.set(key, 'loading');
  promise = (async () => {
    try {
      const from = chunkIdx * GRID_CHUNK + 1;
      const data = await fetchJsonView(`/api/view/grid/rows?${tokenQuery}&sheet=${sheetIdx}&from=${from}&count=${GRID_CHUNK}`);
      gridState.chunks.set(key, data.rows);
      renderGridViewport();
      // 방향키를 빠르게 눌러 아직 안 불러와진(청크 로딩 중인) 칸으로 먼저 선택이 옮겨간 경우,
      // 그 칸의 값이 지금 막 도착했을 수 있다 — 수식 입력줄 자리(값 표시)가 빈 채로 남지 않도록
      // 선택 표시도 같이 다시 맞춘다(다른 청크 로딩이면 gridState.selected와 무관해 값은 그대로다).
      updateSelectionDisplays();
    } catch {
      gridState.chunks.delete(key);
    } finally {
      chunkLoadPromises.delete(key);
    }
  })();
  chunkLoadPromises.set(key, promise);
  return promise;
}

async function fetchJsonView(url, options = {}) {
  const res = await fetch(url, { signal: viewAbort.signal, ...options });
  const json = await res.json();
  if (!res.ok || !json.success) throw new Error(json.error || `요청 실패 (HTTP ${res.status})`);
  return json.data;
}

function cellAt(sheetIdx, rowNumber, colIdx) {
  const chunkIdx = Math.floor((rowNumber - 1) / GRID_CHUNK);
  const chunk = gridState.chunks.get(`${sheetIdx}:${chunkIdx}`);
  if (!chunk || chunk === 'loading') return undefined;
  const row = chunk[(rowNumber - 1) % GRID_CHUNK];
  return row ? row[colIdx] : 0;
}

// 지금은 항상 텍스트 콘텐츠 자리(예: <span>...</span> 안쪽)에만 쓰여 따옴표 이스케이프가 없어도
// 뚫리지 않지만, 나중에 이 값이 속성 자리(예: title="...")에 쓰이게 바뀌어도 안전하도록 처음부터
// 따옴표까지 전부 이스케이프한다.
function escapeHtml(text) {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/* ---------- 하이퍼링크(바로가기) — 편집 화면(src/client/gridLinks.ts)과 같은 규칙. 이 파일은 번들 없이 로드돼 따로 들고 있다 ---------- */

/** 바깥 주소로 열어도 되는 것만(웹·메일·전화) — javascript:·file:·네트워크 경로는 열지 않는다 */
const LINK_SAFE_URL = /^(https?:|mailto:|tel:)/i;
const linkColNumber = (letters) => [...letters.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;

/** gridLinks.ts parseLocation과 같다 — "'자료'!A1", "#요약!$C$5", "자료!B3:H3" → 시트 이름(없으면 null)·칸(0부터, 범위면 row2·col2) */
function parseLinkLocation(location) {
  const text = location.trim().replace(/^#/, '');
  const bang = text.lastIndexOf('!');
  const sheetName = bang >= 0 ? text.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : null;
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(bang >= 0 ? text.slice(bang + 1) : text);
  if (!m) return null;
  const row = Number(m[2]) - 1;
  const col = linkColNumber(m[1]);
  const row2 = m[4] ? Number(m[4]) - 1 : row;
  const col2 = m[3] ? linkColNumber(m[3]) : col;
  return { sheetName, row: Math.min(row, row2), col: Math.min(col, col2), row2: Math.max(row, row2), col2: Math.max(col, col2) };
}

/** gridLinks.ts evalLinkExpr과 같다 — HYPERLINK 첫 인자("글자"와 칸 참조를 &로 이은 것만)를 누른 칸 기준으로 푼다 */
function evalLinkExpr(expr, anchor, at, valueAt) {
  const token = /\s*(?:"((?:[^"]|"")*)"|(\$?)([A-Za-z]{1,3})(\$?)(\d+))\s*/y;
  let out = '';
  let pos = 0;
  for (;;) {
    token.lastIndex = pos;
    const m = token.exec(expr);
    if (!m) return null;
    if (m[1] !== undefined) out += m[1].replace(/""/g, '"');
    else {
      const col = linkColNumber(m[3]) + (m[2] ? 0 : at.col - anchor.col);
      const row = Number(m[5]) - 1 + (m[4] ? 0 : at.row - anchor.row);
      if (row < 0 || col < 0) return null;
      out += valueAt(row, col);
    }
    pos = token.lastIndex;
    if (pos >= expr.length) return out;
    if (expr[pos] !== '&') return null;
    pos += 1;
  }
}

/** 지금 시트 칸(0부터)의 보이는 값 */
function linkCellText(r0, c0) {
  const cell = cellAt(gridState.active, r0 + 1, c0);
  return cell && cell !== 0 ? cell.v : '';
}

const LINK_BUCKET = 64;

function setSheetLinks(sheet) {
  gridState.sheetLinks = sheet ? gridState.links.filter((l) => l.sheetName === sheet.name) : [];
  gridState.sheetLinkBuckets = new Map();
  gridState.sheetLinksWide = [];
  for (const l of gridState.sheetLinks) {
    const b1 = Math.floor(l.r1 / LINK_BUCKET);
    const b2 = Math.floor(l.r2 / LINK_BUCKET);
    if (b2 - b1 > 256) {
      gridState.sheetLinksWide.push(l);
      continue;
    }
    for (let b = b1; b <= b2; b += 1) {
      const list = gridState.sheetLinkBuckets.get(b);
      if (list) list.push(l);
      else gridState.sheetLinkBuckets.set(b, [l]);
    }
  }
}

/** 칸(1부터)에 걸린 링크 — 수식 링크가 빈 값을 보이는 칸(IFERROR(…,""))이면 링크로 치지 않는다 */
function linkAt(r, c) {
  const hit = (l) => r - 1 >= l.r1 && r - 1 <= l.r2 && c - 1 >= l.c1 && c - 1 <= l.c2;
  const link = (gridState.sheetLinkBuckets.get(Math.floor((r - 1) / LINK_BUCKET)) || []).find(hit) || gridState.sheetLinksWide.find(hit);
  if (link && link.expr && linkCellText(r - 1, c - 1) === '') return undefined;
  return link;
}

function linkTargetOf(link, r, c) {
  if (!link.expr || !link.anchor) return link;
  const text = (evalLinkExpr(link.expr, link.anchor, { row: r - 1, col: c - 1 }, linkCellText) ?? '').trim();
  if (!text) return {};
  if (text.startsWith('#')) return { location: text.slice(1) };
  return LINK_SAFE_URL.test(text) ? { url: text } : {};
}

function linkTitleOf(link, r, c) {
  if (link.blocked) return `${link.tooltip ? `${link.tooltip}\n` : ''}이 주소는 웹 화면에서 열 수 없습니다(파일·네트워크 경로): ${link.blocked}`;
  const target = linkTargetOf(link, r, c);
  return link.tooltip || (target.url ? `${target.url}\n클릭하면 열립니다 (Ctrl+클릭: 칸 선택)` : `클릭하면 ${target.location ?? '연결된 곳'}(으)로 이동합니다 (Ctrl+클릭: 칸 선택)`);
}

/** 편집 화면 officeAlert와 같은 모양의 알림(.office-dialog) — 글은 textContent로만 넣는다 */
function showGridNotice(message) {
  const back = document.createElement('div');
  back.className = 'office-dialog office-dialog-runtime';
  back.setAttribute('role', 'alertdialog');
  back.setAttribute('aria-modal', 'true');
  const card = document.createElement('div');
  card.className = 'office-dialog-card';
  const title = document.createElement('h2');
  title.textContent = '알림';
  const body = document.createElement('p');
  body.className = 'office-dialog-message';
  body.textContent = message;
  const actions = document.createElement('div');
  actions.className = 'office-dialog-actions';
  const ok = document.createElement('button');
  ok.type = 'button';
  ok.className = 'btn-mini btn-mini-primary';
  ok.textContent = '확인';
  actions.append(ok);
  card.append(title, body, actions);
  back.append(card);
  const close = () => {
    back.remove();
    document.getElementById('grid-scroll')?.focus({ preventScroll: true });
  };
  ok.addEventListener('click', close);
  back.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' || event.key === 'Enter') {
      event.preventDefault();
      close();
    }
  });
  document.body.append(back);
  ok.focus();
}

/** 문서 안 위치로 — 그 시트로 바꾸고 칸(범위면 범위 전체)을 고른 뒤 보이게 스크롤한다. 없는 시트면 false */
function goToLinkTarget(t) {
  const want = t.sheetName === null ? null : t.sheetName.toLowerCase();
  const idx = want === null ? gridState.active : gridState.sheets.findIndex((s) => s.name.toLowerCase() === want);
  if (idx < 0) return false;
  if (idx !== gridState.active) activateSheet(idx);
  // 병합에 덮인 칸이면 병합의 왼쪽 위 칸(자기 <td>가 있는 칸)을 고른다
  const m = gridState.mergeCovered.get(`${t.row + 1}:${t.col + 1}`);
  gridState.selected = m ? { row: m.r, col: m.c } : { row: t.row + 1, col: t.col + 1 };
  gridState.selRange = t.row2 !== t.row || t.col2 !== t.col ? { r1: t.row + 1, c1: t.col + 1, r2: t.row2 + 1, c2: t.col2 + 1 } : null;
  updateSelectionDisplays();
  // 화면 밖 목적지는 가운데쯤으로 옮긴다(편집 화면 tdForEnsureRendered와 같다) — 가장자리에 붙이면 머리글·스크롤 막대에 가린다
  const scroll = document.getElementById('grid-scroll');
  const sheet = gridState.sheets[gridState.active];
  const { row, col } = gridState.selected;
  const top = gridState.layout.top(row - 1);
  const bottom = gridState.layout.top(row);
  // 틀 고정 행은 위를 덮고 있다 — 그 아래 칸만 보이는 자리로 친다. shortcut: 틀 고정 열은 따지지 않는다(드묾), 생기면 같은 식으로
  const freezeRows = Math.min(sheet.freezeRows || 0, sheet.rowCount);
  const frozenH = row > freezeRows ? gridState.layout.top(freezeRows) : 0;
  if (row > freezeRows && (top < scroll.scrollTop + frozenH + 30 || bottom > scroll.scrollTop + scroll.clientHeight - 20)) {
    scroll.scrollTop = Math.max(0, top - frozenH - (scroll.clientHeight - frozenH) / 2);
  }
  let left = GRID_ROWNUM_W;
  for (let c = 1; c < col; c += 1) left += colWidthAt(sheet, c);
  const right = left + colWidthAt(sheet, col);
  if (left < scroll.scrollLeft + GRID_ROWNUM_W || right > scroll.scrollLeft + scroll.clientWidth - 20) scroll.scrollLeft = Math.max(0, left - GRID_ROWNUM_W - 40);
  renderGridViewport();
  syncGridHead();
  return true;
}

function followLink(link, r, c) {
  const { url, location } = linkTargetOf(link, r, c);
  if (url && LINK_SAFE_URL.test(url.trim())) {
    window.open(url, '_blank', 'noopener,noreferrer');
    return;
  }
  const target = location ? parseLinkLocation(location) : null;
  if (target) {
    if (!goToLinkTarget(target)) showGridNotice(`참조가 올바르지 않습니다.\n'${target.sheetName}' 시트를 찾을 수 없습니다.`);
    return;
  }
  if (location) {
    showGridNotice(`참조가 올바르지 않습니다.\n'${location}' 위치를 찾을 수 없습니다.`);
    return;
  }
  const shown = link.blocked ?? url;
  if (shown) showGridNotice(`이 바로가기는 웹 화면에서 열 수 없습니다.\n${shown}\n(엑셀이 설치된 PC에서 파일을 내려받아 여세요.)`);
}

/** 셀 값 중 검색어와 일치하는 부분만 <mark>로 감싼다 (대소문자 무시) */
function markMatches(rawText, query) {
  if (!query) return escapeHtml(rawText);
  const lower = rawText.toLowerCase();
  const needle = query.toLowerCase();
  const parts = [];
  let from = 0;
  for (;;) {
    const at = lower.indexOf(needle, from);
    if (at === -1) break;
    parts.push(escapeHtml(rawText.slice(from, at)));
    parts.push(`<mark>${escapeHtml(rawText.slice(at, at + needle.length))}</mark>`);
    from = at + needle.length;
  }
  parts.push(escapeHtml(rawText.slice(from)));
  return parts.join('');
}

/** 실제 엑셀처럼, 오른쪽 빈 칸 위로 값이 "넘쳐 보이는"(스필) 동작은 기본이 왼쪽 정렬인 칸에서만
 * 일어난다 — 가운데/오른쪽 정렬이 걸린 칸(대표적으로 SUM 등 계산 결과에 가운데 정렬을 걸어 둔
 * 칸)은 실제 엑셀에서도 절대 옆 칸으로 넘치지 않고 자기 칸 폭 안에서 잘린다. 예전엔 이 구분 없이
 * 값 있는 칸이면 무조건 폭을 넓혔는데, 그러면 넓어진 span의 text-align이 가운데/오른쪽으로
 * 상속되면서 글자가 원래 칸에서 몇 칸이나 떨어진 위치에 그려져 "엉뚱한 칸에 값이 보인다"는
 * 사용자 신고로 이어졌다(칸 자체의 data-r/data-c와 실제 데이터는 항상 맞았다 — 순수 표시 버그).
 *
 * "이 칸의 유효 정렬"을 판단할 때 명시적 정렬(style.a)이 있으면 그대로 따르고, 없으면 엑셀의
 * 기본 규칙(숫자/날짜/수식 결과는 기본 오른쪽 정렬, 일반 텍스트는 기본 왼쪽 정렬)을 흉내낸다.
 * 이 시점엔 원본 셀 값의 타입 정보가 없고 화면에 그려질 텍스트(rawText, cell.v)만 있으므로,
 * 우선 스타일에 적힌 숫자 서식(style.nf — src/spreadsheet.ts의 classifyNumFmt가 매긴다)을 보고,
 * 그것도 없으면 표시 텍스트 자체가(쉼표·통화 기호·퍼센트·공백·회계식 음수 괄호를 걷어내고 보면)
 * 유한한 숫자로 읽히는지를 폴백으로 쓴다 — src/spreadsheet.ts의 unformattedCellValue와 같은
 * 발상이지만 그 함수는 "값을 서식 없는 텍스트로 복원"하는 게 목적이라 계약이 달라(여긴 값을
 * 바꿀 필요가 없고 "숫자처럼 보이는가"만 알면 된다) 그대로 재사용하지 않고 판정 전용으로
 * 다시 작성했다. getComputedStyle 같은 강제 레이아웃 계산은 절대 쓰지 않는다 — 그리드는 스크롤
 * 프레임마다 수십~수백 칸을 다시 그리므로 칸마다 그런 호출이 들어가면 레이아웃 스래싱으로
 * 직결된다(grid-editor.ts의 applySpillDisplay 문서 주석에 있는 것과 같은 교훈). 여기 판정은
 * 전부 이미 들고 있는 데이터(style 객체, 문자열)만으로 끝난다. */
// 날짜 칸도 엑셀 기본은 오른쪽 정렬이라 스필 대상이 아니다. classifyNumFmt(src/spreadsheet.ts)는
// 의도적으로 날짜를 style.nf 범주에 넣지 않으므로(그 문서 주석 참고 — nf를 건드리면 서식 드롭다운·
// xlsx 재내보내기 쪽에 더 눈에 띄는 부작용이 생긴다) 여기서 표시 텍스트의 "모양"만으로 판정한다:
// 이 앱이 날짜를 그리는 경로는 formatDate(src/spreadsheet.ts) 단 하나뿐이고, 원본 서식 코드가
// 무엇이든 항상 이 모양(YYYY-MM-DD, 또는 시:분이 있으면 그 뒤에 공백+HH:MM)으로만 렌더링되므로
// 이 정규식과 일치하는 텍스트는 항상 날짜다(리터럴 날짜 칸은 displayValue가, 수식 결과 날짜는
// formatFormulaResult가 둘 다 이 formatDate를 그대로 호출한다).
const DATE_LIKE_TEXT = /^\d{4}-\d{2}-\d{2}( \d{2}:\d{2})?$/;
/** 숫자처럼 보이는 글자(쉼표·통화·% 포함) — 엑셀은 숫자를 옆 칸으로 넘기지 않는다 */
function isGridNumericText(rawText) {
  const stripped = String(rawText).trim().replace(/[₩$€£¥,%s()]/g, '');
  return stripped !== '' && Number.isFinite(Number(stripped));
}

function isSpillEligible(style, rawText) {
  if (style) {
    if (style.a === 'c' || style.a === 'r') return false; // 명시적 가운데/오른쪽 정렬 — 절대 안 넘친다
    if (style.a === 'l') return true; // 명시적 왼쪽 정렬 — 항상 넘친다
    if (style.nf || (style.fmt && style.fmt !== '@')) return false; // 숫자·날짜 서식이 걸려 있으면 기본 오른쪽 정렬(글자 값이면 서버가 a:'l'을 박아 둔다)
  }
  if (rawText === 'TRUE' || rawText === 'FALSE') return false; // 논리값도 엑셀 기본은 가운데 정렬
  if (DATE_LIKE_TEXT.test(rawText.trim())) return false; // 날짜도 엑셀 기본은 오른쪽 정렬(위 주석 참고)
  const trimmed = rawText.trim();
  // 회계 서식의 음수 괄호 표기 "(1,234)"도 숫자로 인식되게 부호로 풀어준다
  const unwrapped =
    trimmed.length > 1 && trimmed.startsWith('(') && trimmed.endsWith(')') ? `-${trimmed.slice(1, -1)}` : trimmed;
  const stripped = unwrapped.replace(/[₩$€£¥,%\s]/g, '');
  const looksNumeric = stripped !== '' && Number.isFinite(Number(stripped));
  return !looksNumeric;
}

function renderGridViewport() {
  const sheet = gridState.sheets[gridState.active];
  if (!sheet) return;
  const scroll = document.getElementById('grid-scroll');
  const layout = gridState.layout;
  const firstVisible = Math.max(1, layout.rowAt(scroll.scrollTop) + 1 - GRID_OVERSCAN);
  const lastVisible = Math.min(displayRowCount(sheet), layout.rowAt(scroll.scrollTop + scroll.clientHeight) + 1 + GRID_OVERSCAN);
  const colCount = displayColCount(sheet);

  // 세로 병합(rowspan)의 앵커 행(rowspan 속성을 실제로 지닌 <td>가 그려지는, 병합의 맨 위 행)이
  // 가상 스크롤 창보다 위로 스크롤되어 나가 있어도, 그 병합에 덮인 나중 행들은 여전히 창 안에 보일
  // 수 있다 — table.innerHTML을 매번 통째로 새로 그리는 구조라 이전 렌더의 앵커 <td>는 남아있지
  // 않으므로, 이번 패스에서 앵커 행 자체를 함께 그려주지 않으면 그 칸이 통째로 비어(덮인 행이라는
  // 이유로 <td> 자체를 건너뛰므로) 이후 모든 열이 한 칸씩 밀려 보인다(세로로 긴 병합에서 재현되는
  // 버그). 그래서 창과 겹치는 모든 병합의 앵커 행이 실제 렌더 시작 지점 안에 들어오도록 시작
  // 지점을 앞으로 당긴다 — 한 병합을 당기며 새로 포함된 행이 또 다른(더 위에서 시작하는) 병합에
  // 덮여 있을 수도 있으므로(예: 서로 다른 열에 높이가 다른 병합 헤더가 계단식으로 겹치는 표) 더
  // 당길 병합이 없을 때까지 반복한다. (반대로 앵커가 창 안에 있고 덮인 행만 창 아래로 넘치는 경우는
  // 손댈 필요가 없다 — rowspan은 실제로 그려진 <tr> 개수만큼만 걸쳐지므로 앵커 <td>만 있으면
  // 충분하고, 창 밖으로 넘치는 나머지는 그 시점에 스크롤되어 들어올 때 다시 그려진다.)
  //
  // 단, 이 확장에는 GRID_MERGE_EXPAND_CAP 상한이 있다 — 이 함수는 위에서 설명한 것처럼 "실제 병합
  // 크기만큼만" 유계로 늘어나므로, 병합 자체가 극단적으로 크면(예: 1행에 앵커를 두고 시트 전체를
  // 덮는 병합) 그 "유계"가 시트 전체 크기(최대 MAX_ROWS=300,000행, src/spreadsheet.ts)만큼 커질 수
  // 있다 — 시트 맨 아래로 스크롤하기만 해도 renderFirst가 1행까지 끌려가 수십만 개의 <tr>과(아래
  // 청크 프리페치 루프가 함께 늘어나) 수천 건의 동시 청크 요청을 만들어내는, 이 뷰어가 다른 곳
  // (청크 분할·LazyGrid 등)에서 애써 피해온 "거대한 시트를 통째로 처리" 문제 그 자체를 재현한다.
  // 그래서 앵커가 firstVisible보다 GRID_MERGE_EXPAND_CAP행 넘게 떨어져 있으면 그 병합은 당겨오지
  // 않고 degradedMerges에 표시해둔다 — 대신 그 병합에 덮인, 지금 창에 실제로 보이는 칸들은(아래
  // 렌더 루프에서) 병합되지 않은 각자의 <td>로 그린다. 즉 이 극단적인 경우엔 "그 병합만" 하나로
  // 이어붙은 rowspan 셀처럼 보이지 않게 되는 대가로 성능을 지킨다 — 병합이 시각적으로 안 이어지는
  // 것과 열이 어긋나는 것 중, 후자만은 항상 막는다는 절충이다(열 정렬이 깨지는 원래 버그보다는
  // 훨씬 낫다). 실전 문서에서 수천 행을 넘는 단일 병합은 매우 드물다.
  let renderFirst = firstVisible;
  const minRenderFirst = Math.max(1, firstVisible - GRID_MERGE_EXPAND_CAP);
  const degradedMerges = new Set();
  for (let extending = true; extending; ) {
    extending = false;
    for (const m of sheet.merges) {
      if (m.r >= renderFirst || m.r > lastVisible) continue;
      const mergeLastRow = m.r + m.rs - 1;
      if (mergeLastRow < renderFirst) continue;
      if (degradedMerges.has(m)) continue;
      if (m.r < minRenderFirst) {
        degradedMerges.add(m);
        continue;
      }
      renderFirst = m.r;
      extending = true;
    }
  }
  // moveSelectionSimple/moveSelectionToEdge가 이번 렌더에서 어떤 병합이 저하됐는지 알 수 있도록
  // gridState에 남겨 둔다(그 두 함수는 mergeCovered만으로는 "실제로 자기 <td>가 있는 덮인 칸"과
  // "정말 건너뛰어야 할 덮인 칸"을 구분할 수 없다).
  gridState.degradedMerges = degradedMerges;

  // 필요한 청크를 비동기로 요청 — 실제 데이터 범위 너머(패딩 영역)는 서버가 빈 배열을 돌려준다
  // (renderFirst부터 요청해야 위에서 새로 끌어온 병합 앵커 행의 데이터도 함께 로딩된다. 위 상한
  // 덕분에 renderFirst 자체가 firstVisible보다 GRID_MERGE_EXPAND_CAP행 넘게 당겨지지 않으므로,
  // 이 루프가 추가로 요청하는 청크 수도 그만큼(최대 GRID_MERGE_EXPAND_CAP/GRID_CHUNK개)으로 유계다)
  const firstChunk = Math.floor((renderFirst - 1) / GRID_CHUNK);
  const lastChunk = Math.floor((Math.min(lastVisible, sheet.rowCount) - 1) / GRID_CHUNK);
  for (let c = firstChunk; c <= lastChunk; c += 1) fetchChunk(gridState.active, c);

  const cols = `<colgroup><col style="width:${GRID_ROWNUM_W}px">${Array.from({ length: colCount }, (_, i) =>
    `<col style="width:${colWidthAt(sheet, i + 1)}px">`,
  ).join('')}</colgroup>`;
  // 크롬은 table-layout:fixed + <colgroup>이어도 테이블 자체 width가 auto인 채로 열이 많으면(대략
  // 12~15개 이상) 개별 <col> 폭을 조용히 무시하고 자체적으로 재분배해버린다(grid-editor.ts의 같은
  // 문제와 동일 — 그쪽에서 이미 확인된 크롬 버그). 테이블에 명시적 width를 주면 이 경로를 피한다.
  let totalColWidthPx = GRID_ROWNUM_W;
  for (let i = 1; i <= colCount; i += 1) totalColWidthPx += colWidthAt(sheet, i);

  const rowsHtml = [];
  // 엑셀 "틀 고정" — 위 freezeRows행·왼쪽 freezeCols열은 스크롤해도 붙어 있다(position:sticky). 고정 행은 아래로 스크롤해도
  // 늘 그리고, 그 높이만큼 표를 덜 내려 그 아래 행들이 제자리에 오게 한다(고정 행이 그 위를 덮는다).
  const freezeRows = Math.min(sheet.freezeRows || 0, sheet.rowCount);
  const freezeCols = sheet.freezeCols || 0;
  const bodyFirst = Math.max(renderFirst, freezeRows + 1);
  const rowNumbers = [];
  for (let r = 1; r <= Math.min(freezeRows, renderFirst - 1); r += 1) rowNumbers.push(r);
  for (let r = renderFirst; r <= lastVisible; r += 1) rowNumbers.push(r);
  const frozenColLeft = [];
  for (let c = 1, x = GRID_ROWNUM_W; c <= freezeCols; c += 1) {
    frozenColLeft[c] = x;
    x += colWidthAt(sheet, c);
  }
  for (const r of rowNumbers) {
    const rowH = layout.height(r - 1);
    const frozenRow = r <= freezeRows;
    const rowStick = frozenRow ? `position:sticky;top:${layout.top(r - 1)}px;z-index:4` : '';
    const cells = [`<td class="g-rownum"${frozenRow ? ` style="${rowStick};z-index:5"` : ''}>${rowH > 0 ? r : ''}</td>`];
    for (let c = 1; c <= colCount; c += 1) {
      // 이 칸을 덮는 병합이 있어도, 그 병합이 degradedMerges에 있으면(위 상한 때문에 앵커를 당겨오지
      // 못한 병합) 이 칸을 "덮인 칸"으로 건너뛰지 않고 아래에서 보통 셀처럼 그린다 — mergeStarts에
      // 이 좌표의 앵커가 없으므로(covered 칸은 원래 앵커가 아니다) span은 자연히 비게 되어, 시각적
      // 병합 없이도 열 자리는 항상 채워진다(정렬 보존이 이 저하 렌더링의 유일한 목적).
      const coveringMerge = gridState.mergeCovered.get(`${r}:${c}`);
      if (coveringMerge && !degradedMerges.has(coveringMerge)) continue;
      const merge = gridState.mergeStarts.get(`${r}:${c}`);
      const cell = cellAt(gridState.active, r, c - 1);
      const isEmpty = !cell || cell === 0;
      const classes = [];
      if (!isEmpty && cell.s !== undefined) classes.push(`gs${cell.s}`);
      const hitKey = `${gridState.active}:${r}:${c}`;
      const isHit = gridSearchHits.has(hitKey);
      if (isHit) classes.push('g-hit');
      if (hitKey === gridCurrentHit) classes.push('g-hit-current');
      if (gridState.selected && gridState.selected.row === r && gridState.selected.col === c) classes.push('g-selected');
      const sr = gridState.selRange;
      if (sr && r >= sr.r1 && r <= sr.r2 && c >= sr.c1 && c <= sr.c2) classes.push('g-range-selected');
      const link = gridState.sheetLinks.length ? linkAt(r, c) : undefined;
      if (link) classes.push('g-link');
      const linkTitle = link ? ` title="${escapeHtml(linkTitleOf(link, r, c))}"` : '';
      const cellStyleObj = isEmpty || cell.s === undefined ? undefined : gridState.styles[cell.s];
      // 표시 글자는 서버가 칸 서식(numberFormat.ts — 편집 화면과 같은 서식기)대로 이미 만들어 둔다
      const shown = isEmpty ? '' : cell.v;
      // 엑셀 기본 정렬 — 정렬을 따로 지정하지 않은 칸에서 숫자·날짜는 오른쪽, 논리값은 가운데.
      // (판정은 넘침 여부와 같은 규칙: isSpillEligible이 "기본 왼쪽 정렬인 텍스트"만 true를 준다.)
      if (shown && !(cellStyleObj && cellStyleObj.a) && !isSpillEligible(cellStyleObj, shown)) {
        classes.push(shown === 'TRUE' || shown === 'FALSE' ? 'g-bool' : 'g-num');
      }
      const fillAt = isEmpty ? -1 : shown.indexOf(GRID_FILL_MARK);
      // 채움(회계 서식의 ₩ … 25,000) — 앞뒤를 칸 양 끝에 붙인다
      const value = isEmpty
        ? ''
        : fillAt >= 0
          ? `<span class="g-fill"><span>${escapeHtml(shown.slice(0, fillAt))}</span><span>${escapeHtml(shown.slice(fillAt + 1))}</span></span>`
          : isHit
            ? markMatches(shown, searchState.lastQuery)
            : escapeHtml(shown);
      const span = merge ? ` colspan="${merge.cs}"${merge.rs > 1 ? ` rowspan="${merge.rs}"` : ''}` : '';
      // 실제 엑셀처럼 빈 이웃 칸 위로 내용이 넘쳐 보이게 한다(병합 칸이나 값 있는 칸에서 멈춘다).
      // 줄 바꿈 칸은 넘치지 않고 자기 칸 안에서 줄을 바꾼다.
      let spillStyle = '';
      // 가운데 정렬 글자는 양옆 빈 칸으로 고르게, 오른쪽 정렬 글자(숫자 아님)는 왼쪽 빈 칸으로 넘친다(엑셀과 같다)
      const align = cellStyleObj && cellStyleObj.a;
      if (cellStyleObj && cellStyleObj.shr) classes.push('g-shrink');
      const spillMode = isEmpty || merge || (cellStyleObj && (cellStyleObj.w || cellStyleObj.shr))
        ? null
        : align === 'c'
          ? 'center'
          : align === 'r'
            ? (isGridNumericText(cell.v) ? null : 'left')
            : isSpillEligible(cellStyleObj, cell.v) ? 'right' : null;
      if (spillMode) {
        const emptyAt = (cc) => {
          if (cc < 1 || cc > colCount) return false;
          if (gridState.mergeCovered.has(`${r}:${cc}`) || gridState.mergeStarts.has(`${r}:${cc}`)) return false;
          const other = cellAt(gridState.active, r, cc - 1);
          return !other || other === 0 || other.v === '';
        };
        let right = 0;
        if (spillMode !== 'left') for (let cc = c + 1; emptyAt(cc); cc += 1) right += colWidthAt(sheet, cc);
        let left = 0;
        if (spillMode !== 'right') for (let cc = c - 1; emptyAt(cc); cc -= 1) left += colWidthAt(sheet, cc);
        if (spillMode === 'center') {
          const side = Math.min(left, right);
          left = side;
          right = side;
        }
        const own = colWidthAt(sheet, c);
        if (left + right > 0) spillStyle = ` style="width:${own + left + right}px${left ? `;margin-left:-${left}px` : ''}"`;
      }
      // 내용이 행 높이를 넘어도 행이 늘어나지 않게(엑셀처럼 잘린다) — 병합이면 걸친 행 높이의 합
      const cellH = merge && merge.rs > 1 ? layout.top(r - 1 + merge.rs) - layout.top(r - 1) : rowH;
      const hiddenCol = colWidthAt(sheet, c) === 0 && !(merge && merge.cs > 1);
      if (hiddenCol) classes.push('g-hidden-col');
      const boxStyle = `max-height:${Math.max(0, cellH - 1)}px`;
      const content = value && !hiddenCol && cellH > 0
        ? `<span class="g-cell-spill" style="${boxStyle}${spillStyle ? ';' + spillStyle.slice(8, -1) : ''}">${value}</span>`
        : '';
      const colStick = c <= freezeCols ? `position:sticky;left:${frozenColLeft[c]}px;z-index:${frozenRow ? 4 : 3}` : '';
      // 고정 행·열이 겹치는 왼쪽 위 칸이 맨 위(고정 행의 나머지 칸이 가로로 밀려 와도 그 위에 그려진다)
      const stick = frozenRow ? (colStick ? `${rowStick};left:${frozenColLeft[c]}px;z-index:6` : rowStick) : colStick;
      if (stick) classes.push('g-frozen');
      const clsSticky = classes.length > 0 ? ` class="${classes.join(' ')}"` : '';
      cells.push(`<td data-r="${r}" data-c="${c}"${clsSticky}${span}${linkTitle}${stick ? ` style="${stick}"` : ''}>${content}</td>`);
    }
    // 기본보다 낮은 행(g-row-short) — 행 번호 글자 높이 때문에 행이 커지지 않게(CSS)
    rowsHtml.push(`<tr${rowH === 0 ? ' class="g-hidden-row"' : rowH < 20 ? ' class="g-row-short"' : ''} style="height:${rowH}px">${cells.join('')}</tr>`);
  }

  const table = document.getElementById('grid-table');
  table.innerHTML = cols + rowsHtml.join('');
  // 셀에 맞춤 — 칸보다 넓은 글자는 글자 크기를 줄인다(편집 화면 shrinkToFit과 같은 규칙, 이런 칸은 드물다)
  for (const td of table.querySelectorAll('td.g-shrink')) {
    const span = td.querySelector('.g-cell-spill');
    const room = td.clientWidth - 6;
    if (span && room > 0 && span.scrollWidth > room) span.style.fontSize = `${(parseFloat(getComputedStyle(span).fontSize) * room) / span.scrollWidth}px`;
  }
  table.style.width = `${totalColWidthPx}px`;
  // 고정 행을 앞에 덧붙였으면 그만큼 덜 내린다 — 그래야 bodyFirst행이 제 위치(layout.top)에 온다
  const prepended = rowNumbers.length > 0 && rowNumbers[0] < renderFirst ? layout.top(Math.min(freezeRows, renderFirst - 1)) : 0;
  const tableY = layout.top((prepended ? bodyFirst : renderFirst) - 1) - prepended;
  // 고정 행·열(sticky)은 transform을 모른다(위치 계산이 transform 전 자리 기준) — 틀 고정 시트는 top으로 옮긴다
  const sticky = freezeRows > 0 || freezeCols > 0;
  table.style.transform = sticky ? 'none' : `translateY(${tableY}px)`;
  table.style.top = sticky ? `${tableY}px` : '';
}

function renderGridHead() {
  const sheet = gridState.sheets[gridState.active];
  const head = document.getElementById('grid-head');
  const scroll = document.getElementById('grid-scroll');
  const cells = [`<div class="g-head-cell g-head-corner" style="width:${GRID_ROWNUM_W}px"></div>`];
  for (let c = 1; c <= displayColCount(sheet); c += 1) {
    const frozen = c <= (sheet.freezeCols || 0) ? ' g-head-frozen' : '';
    cells.push(`<div class="g-head-cell${frozen}" style="width:${colWidthAt(sheet, c)}px">${colLetter(c)}</div>`);
  }
  head.innerHTML = `<div class="g-head-track">${cells.join('')}</div>`;
  syncGridHead();
}

/** 열 머리글을 가로 스크롤에 맞춘다 — 틀 고정 열의 머리글은 칸처럼 제자리에 남는다 */
function syncGridHead() {
  const track = document.querySelector('.g-head-track');
  const scroll = document.getElementById('grid-scroll');
  if (!track || !scroll) return;
  track.style.transform = `translateX(${-scroll.scrollLeft}px)`;
  for (const cell of track.querySelectorAll('.g-head-frozen')) cell.style.transform = `translateX(${scroll.scrollLeft}px)`;
}

/* ==================== 그리드 확대/축소 ====================
 * grid-editor.ts의 확대/축소(줌 슬라이더 + [－]/[＋])와 동일한 UI·동작·범위를 그대로 옮겨온다.
 * 편집기와 똑같이 CSS zoom을 #grid-zoom-host(헤더 행 + 스크롤 영역을 함께 감싼 래퍼)에 걸어
 * 화면에 그려지는 픽셀만 바꾼다 — renderGridViewport()가 읽는 scroll.clientHeight/scrollTop 등은
 * 줌이 적용되기 전 좌표계 기준 값을 그대로 돌려주므로(grid-editor.ts의 renderCountsFor 주석 참고),
 * 열/행 폭 계산 쪽은 전혀 손댈 필요가 없고 줌이 바뀔 때마다 보이는 범위만 다시 그리면 된다. */
const GRID_ZOOM_MIN = 0.6;
const GRID_ZOOM_MAX = 2;
let gridZoom = 1;
let gridZoomRefitScheduled = false;

function scheduleGridZoomRefit() {
  if (gridZoomRefitScheduled) return;
  gridZoomRefitScheduled = true;
  requestAnimationFrame(() => {
    gridZoomRefitScheduled = false;
    if (gridState.sheets.length === 0) return;
    renderGridViewport();
    syncGridHead();
  });
}

function applyGridZoom() {
  if (!gridZoomHostEl) return;
  gridZoomHostEl.style.zoom = String(gridZoom);
  if (gridZoomLevelEl) gridZoomLevelEl.textContent = `${Math.round(gridZoom * 100)}%`;
  const ratio = (gridZoom - GRID_ZOOM_MIN) / (GRID_ZOOM_MAX - GRID_ZOOM_MIN);
  if (gridZoomSliderFillEl) gridZoomSliderFillEl.style.width = `${ratio * 100}%`;
  if (gridZoomSliderHandleEl) gridZoomSliderHandleEl.style.left = `${ratio * 100}%`;
  scheduleGridZoomRefit();
}

function setGridZoomFromRatio(ratio) {
  const clamped = Math.max(0, Math.min(1, ratio));
  gridZoom = Math.round((GRID_ZOOM_MIN + clamped * (GRID_ZOOM_MAX - GRID_ZOOM_MIN)) * 20) / 20; // 5%p 단위로 스냅
  applyGridZoom();
}

let gridZoomSetupDone = false;
function setupGridZoom() {
  if (gridZoomSetupDone) return; // initGridView는 시트 전환 시가 아니라 문서 로드 시 한 번만 불린다
  gridZoomSetupDone = true;

  document.getElementById('grid-zoom-in')?.addEventListener('click', () => {
    gridZoom = Math.min(GRID_ZOOM_MAX, Math.round((gridZoom + 0.1) * 10) / 10);
    applyGridZoom();
  });
  document.getElementById('grid-zoom-out')?.addEventListener('click', () => {
    gridZoom = Math.max(GRID_ZOOM_MIN, Math.round((gridZoom - 0.1) * 10) / 10);
    applyGridZoom();
  });

  function ratioFromClientX(clientX) {
    const rect = gridZoomSliderTrackEl.getBoundingClientRect();
    return (clientX - rect.left) / rect.width;
  }
  gridZoomSliderTrackEl?.addEventListener('mousedown', (event) => {
    if (event.target === gridZoomSliderHandleEl) return; // 핸들 자체 드래그는 아래 리스너가 처리
    setGridZoomFromRatio(ratioFromClientX(event.clientX));
  });
  gridZoomSliderHandleEl?.addEventListener('mousedown', (event) => {
    event.preventDefault();
    const onMove = (moveEvent) => setGridZoomFromRatio(ratioFromClientX(moveEvent.clientX));
    const onUp = () => {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  });

  window.addEventListener('resize', scheduleGridZoomRefit);
  applyGridZoom();
}

/** 0-based 열 c의 왼쪽 x(px, 행 번호 열 포함) */
function gridColLeft(sheet, c) {
  let x = GRID_ROWNUM_W;
  for (let i = 1; i <= c; i += 1) x += colWidthAt(sheet, i);
  return x;
}

/** 파일에 있던 그림을 칸 위치(앵커)에 맞춰 얹는다 — 표 위에, 클릭은 칸으로 그대로 지나가게 */
function renderGridImages() {
  const sheet = gridState.sheets[gridState.active];
  const inner = document.getElementById('grid-inner');
  let layer = document.getElementById('grid-image-layer');
  if (!layer) {
    layer = document.createElement('div');
    layer.id = 'grid-image-layer';
    layer.className = 'grid-image-layer';
    inner.append(layer);
  }
  layer.replaceChildren();
  if (!sheet) return;
  const layout = gridState.layout;
  for (const image of gridState.images) {
    if (image.sheetName !== sheet.name) continue;
    const left = gridColLeft(sheet, image.from.col) + image.from.dx;
    const top = layout.top(image.from.row) + image.from.dy;
    const width = image.size ? image.size.width : gridColLeft(sheet, image.to.col) + image.to.dx - left;
    const height = image.size ? image.size.height : layout.top(image.to.row) + image.to.dy - top;
    if (!(width > 0 && height > 0)) continue;
    const img = document.createElement('img');
    img.className = 'grid-image';
    img.src = image.src;
    img.alt = image.name || '';
    img.draggable = false;
    img.style.left = `${left}px`;
    img.style.top = `${top}px`;
    img.style.width = `${width}px`;
    img.style.height = `${height}px`;
    layer.append(img);
  }
}

function activateSheet(index) {
  gridState.active = index;
  const sheet = gridState.sheets[index];
  buildMergeIndex(sheet);
  gridState.layout = new GridRowLayout(sheet.defaultRowHeight || GRID_DEFAULT_ROW_H, sheet.rowHeights || []);
  document.getElementById('grid-table').classList.toggle('g-no-gridlines', sheet.showGridLines === false);
  const inner = document.getElementById('grid-inner');
  inner.style.height = `${gridState.layout.top(displayRowCount(sheet))}px`;
  let totalWidth = 0;
  for (let c = 1; c <= displayColCount(sheet); c += 1) totalWidth += colWidthAt(sheet, c);
  inner.style.width = `${GRID_ROWNUM_W + totalWidth}px`;
  const scroll = document.getElementById('grid-scroll');
  // 시트를 바꿀 때마다 항상 A1(왼쪽 위)부터 보이도록 세로·가로 스크롤을 모두 되돌린다 — 가로
  // 스크롤(scrollLeft)을 빼먹으면, 이전 시트를 오른쪽으로 스크롤해 둔 채로 다른 시트 탭을
  // 누를 때 새 시트도 그 스크롤 위치 그대로(A1이 안 보이는 채로) 열리는 문제가 있었다.
  scroll.scrollTop = 0;
  scroll.scrollLeft = 0;
  for (const [i, tab] of [...document.querySelectorAll('.grid-sheet-tab')].entries()) {
    tab.classList.toggle('active', i === index);
    // 링크로 먼 시트에 왔을 때도 탭 줄이 그 탭을 보여 주게(엑셀과 같다)
    if (i === index) tab.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
  renderGridHead();
  renderGridImages();
  setSheetLinks(sheet);
  gridState.selRange = null;
  gridState.selected = { row: 1, col: 1 };
  renderGridViewport();
  updateSelectionDisplays();
  // A1이 시각적으로는 바로 선택돼 보이지만, 클릭을 한 번 해야만 방향키가 먹는 건 어색하다 —
  // 시트를 열거나(또는 다른 시트로 전환) 하자마자 바로 방향키가 통하도록 스크롤 컨테이너에
  // 포커스를 미리 준다(scrollIntoView는 여기서 필요 없으니 preventScroll로 막는다).
  scroll.focus({ preventScroll: true });
}

/** 선택된 칸의 전체 값 텍스트 — renderGridViewport가 <td> 안에 그리는 것과 같은 cell.v를 그대로
 * 쓴다(이 데이터 모델은 수식 원문과 계산값을 따로 보관하지 않고 하나의 표시용 문자열만 갖는다 —
 * displayValue()/src/spreadsheet.ts 참고: 계산 결과가 있으면 결과를, 없으면 "=수식" 형태의 원문을
 * 이미 합쳐서 내려준다). 그리드 칸 자체는 열 폭에 따라 옆 칸 위로 "넘쳐 보이거나" 다음 값 있는
 * 칸을 만나면 그 지점에서 잘려 보이지만, 여기는 그 폭 제한이 없으니 항상 전체 값을 보여준다. */
function selectedCellText() {
  const sel = gridState.selected;
  if (!sel) return '';
  const cell = cellAt(gridState.active, sel.row, sel.col - 1);
  if (!cell || cell === 0) return '';
  // 수식 입력줄은 엑셀처럼 값(서식 전 — (4,321)이 아니라 -4321, ####로 가린 칸도 값)
  return cell.n ?? cell.v ?? '';
}

/** 서식 코드의 채움(*) 표시 — 이 글자 앞은 칸 왼쪽 끝, 뒤는 오른쪽 끝(src/collab/numberFormat.ts FILL_MARK와 같은 글자) */
const GRID_FILL_MARK = '\u0001';

/** 선택된 칸이 바뀔 때마다 툴바 바로 밑의 수식 입력줄 자리(읽기 전용 #grid-formula-ref/
 * #grid-formula-value)를 갱신한다. */
function updateSelectionDisplays() {
  const ref = gridState.selected ? `${colLetter(gridState.selected.col)}${gridState.selected.row}` : '';
  if (gridFormulaRefEl) gridFormulaRefEl.textContent = ref;
  if (gridFormulaValueEl) gridFormulaValueEl.textContent = selectedCellText();
}

function scrollSelectedIntoView() {
  const sheet = gridState.sheets[gridState.active];
  if (!sheet || !gridState.selected) return;
  const scroll = document.getElementById('grid-scroll');
  const { row, col } = gridState.selected;
  const rowTop = gridState.layout.top(row - 1);
  const rowBottom = gridState.layout.top(row);
  if (rowTop < scroll.scrollTop) scroll.scrollTop = rowTop;
  else if (rowBottom > scroll.scrollTop + scroll.clientHeight) scroll.scrollTop = rowBottom - scroll.clientHeight;

  let colLeft = GRID_ROWNUM_W;
  for (let c = 1; c < col; c += 1) colLeft += colWidthAt(sheet, c);
  const colRight = colLeft + colWidthAt(sheet, col);
  const visibleLeft = scroll.scrollLeft + GRID_ROWNUM_W;
  const visibleRight = scroll.scrollLeft + scroll.clientWidth;
  if (colLeft < visibleLeft) scroll.scrollLeft = colLeft - GRID_ROWNUM_W;
  else if (colRight > visibleRight) scroll.scrollLeft = colRight - scroll.clientWidth;
}

function setSelected(row, col) {
  gridState.selected = { row, col };
  gridState.selRange = null;
  updateSelectionDisplays();
  renderGridViewport();
  scrollSelectedIntoView();
}

/** 지금 칸에서 (dRow,dCol)만큼 떨어진 칸으로 선택을 옮긴다 — 읽기 전용이라 grid-editor.ts의
 * moveSelection과 달리 항상 이미 로드된 화면 안에서만 움직이므로(한 칸씩이라 스크롤 범위를
 * 벗어나지 않는다) 비동기 청크 로딩이 필요 없다. 병합에 덮인 칸(자기 <td>가 없다)은 같은
 * 방향으로 한 칸 더 건너뛴다. */
function moveSelectionSimple(dRow, dCol) {
  const sheet = gridState.sheets[gridState.active];
  if (!sheet) return;
  let { row, col } = gridState.selected ?? { row: 1, col: 1 };
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const nextRow = Math.min(Math.max(row + dRow, 1), sheet.rowCount);
    const nextCol = Math.min(Math.max(col + dCol, 1), sheet.colCount);
    if (nextRow === row && nextCol === col) return; // 시트 끝
    row = nextRow;
    col = nextCol;
    // 이 칸을 덮는 병합이 있어도, 그 병합이 저하돼(degradedMerges) 개별 <td>로 그려졌다면 실제로는
    // 자기 칸이 있는 선택 가능한 칸이다 — 진짜로 건너뛸 대상은 지금도 하나로 이어붙어 그려지는
    // (저하되지 않은) 병합에 덮인 칸뿐이다.
    const coveringMerge = gridState.mergeCovered.get(`${row}:${col}`);
    if (!coveringMerge || gridState.degradedMerges.has(coveringMerge)) {
      setSelected(row, col);
      return;
    }
  }
}

async function isCellEmptyAt(sheetIdx, sheet, row, col) {
  if (row < 1 || col < 1 || row > sheet.rowCount || col > sheet.colCount) return true;
  await ensureChunkLoaded(sheetIdx, Math.floor((row - 1) / GRID_CHUNK));
  const cell = cellAt(sheetIdx, row, col - 1);
  return !cell || cell === 0;
}

/** 엑셀의 Ctrl(⌘)+방향키(End 모드) — 빈 칸이거나 바로 다음 칸이 비어 있으면 다음 데이터(또는
 * 시트 끝)까지 건너뛰고, 데이터가 연속돼 있으면 그 블록이 끝나는 칸까지 단번에 이동한다.
 * 아직 스크롤로 지나가지 않아 안 불러와진 구간도 지나가야 할 수 있어(청크 단위로 그때그때
 * 불러온다) grid-editor.ts의 동기 버전과 달리 비동기다. */
async function moveSelectionToEdge(dRow, dCol) {
  const sheet = gridState.sheets[gridState.active];
  if (!sheet) return;
  const sheetIdx = gridState.active;
  let { row, col } = gridState.selected ?? { row: 1, col: 1 };
  const inBounds = (r, c) => r >= 1 && c >= 1 && r <= sheet.rowCount && c <= sheet.colCount;

  if ((await isCellEmptyAt(sheetIdx, sheet, row, col)) || (await isCellEmptyAt(sheetIdx, sheet, row + dRow, col + dCol))) {
    let r = row + dRow;
    let c = col + dCol;
    while (inBounds(r, c) && (await isCellEmptyAt(sheetIdx, sheet, r, c))) {
      r += dRow;
      c += dCol;
    }
    row = Math.min(Math.max(r, 1), sheet.rowCount);
    col = Math.min(Math.max(c, 1), sheet.colCount);
  } else {
    let r = row;
    let c = col;
    while (inBounds(r + dRow, c + dCol) && !(await isCellEmptyAt(sheetIdx, sheet, r + dRow, c + dCol))) {
      r += dRow;
      c += dCol;
    }
    row = r;
    col = c;
  }
  // 병합에 덮인 칸(자기 <td>가 없다)에 내려앉으면 moveSelectionSimple과 같은 이유로 한 칸 더 간다 —
  // 단, 그 병합이 저하됐다면(degradedMerges) 실제로는 자기 칸이 있으므로 건너뛰지 않는다.
  for (let attempt = 0; attempt < 30; attempt += 1) {
    const coveringMerge = gridState.mergeCovered.get(`${row}:${col}`);
    if (!coveringMerge || gridState.degradedMerges.has(coveringMerge)) {
      setSelected(row, col);
      return;
    }
    const nextRow = Math.min(Math.max(row + dRow, 1), sheet.rowCount);
    const nextCol = Math.min(Math.max(col + dCol, 1), sheet.colCount);
    if (nextRow === row && nextCol === col) {
      setSelected(row, col);
      return;
    }
    row = nextRow;
    col = nextCol;
  }
}

async function initGridView() {
  document.getElementById('pages').hidden = true;
  document.getElementById('grid').hidden = false;
  document.getElementById('zoom-out').hidden = true;
  document.getElementById('zoom-in').hidden = true;
  document.getElementById('zoom-level').hidden = true;
  document.getElementById('zoom-fit-btn').hidden = true;
  document.getElementById('rotate-btn').hidden = true;
  document.getElementById('page-view-btn').hidden = true;
  document.getElementById('page-jump-input').hidden = true;
  // "N/전체행" 정보는 이제 하단 grid-status-bar가 보여준다(예전엔 여기 상단 툴바에 작게 끼어 있었다).
  pageJumpTotal.hidden = true;
  document.getElementById('invert-btn').hidden = true;
  document.getElementById('measure-btn').hidden = true;
  document.getElementById('markup-add-btn').hidden = true;
  document.getElementById('text-tool-btn').hidden = true;
  document.getElementById('markup-include-wrap').hidden = true;
  document.getElementById('download-as-btn').hidden = true;

  const meta = await fetchJsonView(`/api/view/grid/meta?${tokenQuery}`);
  gridState.sheets = meta.sheets;
  // 그림은 늦게 와도 된다 — 도착하면 지금 시트에 얹는다
  // 차트는 서버가 편집 화면과 같은 그리기(xlsx/chartSvg.ts)로 만든 SVG 그림이라 그림과 같이 얹는다
  for (const kind of ['xlsx-images', 'xlsx-chart-images']) {
    void fetchJsonView(`/api/view/${kind}?${tokenQuery}`)
      .then((data) => {
        gridState.images = gridState.images.concat((data && data.images) || []);
        renderGridImages();
      })
      .catch(() => undefined);
  }
  // 하이퍼링크도 늦게 와도 된다 — 도착하면 지금 시트 칸에 표시
  void fetchJsonView(`/api/view/xlsx-links?${tokenQuery}`)
    .then((data) => {
      gridState.links = (data && data.links) || [];
      const sheet = gridState.sheets[gridState.active];
      setSheetLinks(sheet);
      renderGridViewport();
    })
    .catch(() => undefined);
  gridState.styles = meta.styles;
  injectGridStyles(meta.styles);
  // 서버가 시간 예산·비용 상한 때문에 조건부 서식 일부를 굽지 않았으면 알린다
  const cfNotice = document.getElementById('grid-cf-notice');
  if (cfNotice) cfNotice.hidden = !meta.cfSkipped;

  const tabs = document.getElementById('sheet-tabs');
  tabs.replaceChildren(
    ...meta.sheets.map((sheet, i) => {
      const tab = document.createElement('button');
      // 그리드 동시편집기(grid-editor.ts의 renderSheetTabs)와 같은 .grid-sheet-tab 모양을
      // 그대로 쓴다 — 다만 추가(+)/이름 바꾸기/우클릭 메뉴 같은 편집 동작은 달지 않는다.
      tab.className = 'grid-sheet-tab';
      tab.type = 'button';
      tab.role = 'tab';
      tab.textContent = sheet.name;
      tab.addEventListener('click', () => activateSheet(i));
      return tab;
    }),
  );

  setupGridZoom();

  const scroll = document.getElementById('grid-scroll');
  let ticking = false;
  scroll.addEventListener('scroll', () => {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(() => {
      ticking = false;
      renderGridViewport();
      syncGridHead();
    });
  });

  // 열람 전용이라 내용은 못 고쳐도, 어느 칸인지는 클릭으로 짚어볼 수 있어야 편하다 — 가상 스크롤이라
  // <td>가 스크롤마다 다시 그려지므로 개별 셀에 리스너를 달지 않고 테이블에 위임한다.
  document.getElementById('grid-table').addEventListener('click', (event) => {
    const td = event.target.closest('td[data-r]');
    if (!td) return;
    const r = Number(td.dataset.r);
    const c = Number(td.dataset.c);
    setSelected(r, c);
    scroll.focus({ preventScroll: true });
    // 하이퍼링크 칸은 엑셀처럼 한 번 누르면 이동한다 — Ctrl·Shift·Alt를 누른 채면 칸만 고른다
    if (event.button !== 0 || event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
    const link = gridState.sheetLinks.length ? linkAt(r, c) : undefined;
    if (link) followLink(link, r, c);
  });

  // 방향키 이동 + Ctrl(⌘)+방향키로 데이터 끝까지 이동(엑셀 End 모드). 가상 스크롤이라 개별 셀이
  // 포커스를 못 받으니(스크롤마다 새로 그려짐) 스크롤 컨테이너 자체가 포커스를 갖고 키를 받는다.
  scroll.addEventListener('keydown', (event) => {
    const withEdge = event.ctrlKey || event.metaKey;
    switch (event.key) {
      case 'ArrowUp':
        event.preventDefault();
        if (withEdge) void moveSelectionToEdge(-1, 0);
        else moveSelectionSimple(-1, 0);
        break;
      case 'ArrowDown':
        event.preventDefault();
        if (withEdge) void moveSelectionToEdge(1, 0);
        else moveSelectionSimple(1, 0);
        break;
      case 'ArrowLeft':
        event.preventDefault();
        if (withEdge) void moveSelectionToEdge(0, -1);
        else moveSelectionSimple(0, -1);
        break;
      case 'ArrowRight':
        event.preventDefault();
        if (withEdge) void moveSelectionToEdge(0, 1);
        else moveSelectionSimple(0, 1);
        break;
      default:
        break;
    }
  });

  // 복사 권한이 없으면 선택·복사·우클릭을 차단한다 (서버가 원본을 주지 않는 것이 근본 방어)
  if (!permissions.canCopy) {
    const grid = document.getElementById('grid');
    grid.classList.add('grid-noselect');
    grid.addEventListener('copy', (event) => event.preventDefault());
  }

  activateSheet(0);
}

/* ==================== 도면 도구: 커서 기준 줌 · 패닝 · 거리 측정 ==================== */

/** 커서가 가리키는 지점을 유지한 채 배율을 바꾼다 (AutoCAD 휠 줌) */
function zoomAtPoint(nextIndex, clientX, clientY) {
  const bounded = Math.max(0, Math.min(ZOOM_STEPS.length - 1, nextIndex));
  if (bounded === zoomIndex) return;
  const ratio = ZOOM_STEPS[bounded] / ZOOM_STEPS[zoomIndex];
  const rect = pagesContainer.getBoundingClientRect();
  const containerDocTop = rect.top + window.scrollY;
  const contentX = pagesContainer.scrollLeft + (clientX - rect.left);
  const contentY = clientY + window.scrollY - containerDocTop;
  zoomIndex = bounded;
  applyZoom();
  pagesContainer.scrollLeft = contentX * ratio - (clientX - rect.left);
  window.scrollTo(window.scrollX, containerDocTop + contentY * ratio - clientY);
}

const measure = {
  mode: false,
  calibrating: false,
  pending: null, // { frame, fx, fy, dot }
  items: [], // { frame, a, b, paperMm, els: [] }
  factor: 1,
  unit: 'mm(용지)',
};

function measureLayer(frame) {
  let layer = frame.querySelector('.measure-layer');
  if (!layer) {
    layer = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    layer.setAttribute('class', 'measure-layer');
    layer.setAttribute('viewBox', '0 0 100 100');
    layer.setAttribute('preserveAspectRatio', 'none');
    frame.append(layer);
  }
  return layer;
}

function measureDot(frame, fx, fy) {
  const dot = document.createElement('div');
  dot.className = 'measure-dot';
  dot.style.left = `${fx * 100}%`;
  dot.style.top = `${fy * 100}%`;
  frame.append(dot);
  return dot;
}

function measurementLabelText(paperMm) {
  const value = paperMm * measure.factor;
  return `${value >= 100 ? value.toFixed(0) : value.toFixed(1)} ${measure.unit}`;
}

function addMeasurement(a, b) {
  const frame = a.frame;
  const pageIdx = [...pagesContainer.children].indexOf(frame);
  const dims = pageDimensions && pageDimensions[pageIdx];
  if (!dims) return;
  const dxPt = (b.fx - a.fx) * dims.w;
  const dyPt = (b.fy - a.fy) * dims.h;
  const paperMm = (Math.hypot(dxPt, dyPt) * 25.4) / 72;

  const layer = measureLayer(frame);
  const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
  line.setAttribute('x1', String(a.fx * 100));
  line.setAttribute('y1', String(a.fy * 100));
  line.setAttribute('x2', String(b.fx * 100));
  line.setAttribute('y2', String(b.fy * 100));
  layer.append(line);

  const dotB = measureDot(frame, b.fx, b.fy);
  const label = document.createElement('div');
  label.className = 'measure-label';
  label.style.left = `${((a.fx + b.fx) / 2) * 100}%`;
  label.style.top = `${((a.fy + b.fy) / 2) * 100}%`;
  frame.append(label);

  const item = { frame, paperMm, label, els: [line, a.dot, dotB, label] };
  label.textContent = measurementLabelText(paperMm);
  measure.items.push(item);

  if (measure.calibrating) {
    measure.calibrating = false;
    const input = window.prompt('방금 잰 구간의 실제 길이를 입력하세요 (예: 1500 → mm)', '');
    const value = Number((input || '').replace(/[^\d.]/g, ''));
    if (value > 0) {
      measure.factor = value / paperMm;
      measure.unit = 'mm';
      for (const it of measure.items) it.label.textContent = measurementLabelText(it.paperMm);
      document.getElementById('measure-hint').textContent = '보정 완료 — 실제 치수(mm)로 표시됩니다';
    } else {
      document.getElementById('measure-hint').textContent = '보정이 취소되었습니다';
    }
  }
}

function clearMeasurements() {
  for (const item of measure.items) for (const el of item.els) el?.remove();
  measure.items = [];
  if (measure.pending) {
    measure.pending.dot.remove();
    measure.pending = null;
  }
}

function setMeasureMode(on) {
  measure.mode = on;
  measure.calibrating = false;
  document.getElementById('measure-btn').setAttribute('aria-pressed', String(on));
  document.getElementById('measure-bar').hidden = !on;
  pagesContainer.classList.toggle('measuring', on);
  if (!on) clearMeasurements();
  else document.getElementById('measure-hint').textContent = '두 점을 클릭해 거리를 측정하세요';
}

/* ==================== 마크업(리뷰 코멘트) ==================== */

const markup = {
  visible: true,
  drawMode: false,
  items: [],
  dragging: null, // {frame, page, fx, fy, rectEl, box}
  draftRect: null, // 팝오버가 열려 있는 동안 유지되는 사각형 엘리먼트
};

function markupLayer(frame) {
  let layer = frame.querySelector('.markup-layer');
  if (!layer) {
    layer = document.createElement('div');
    layer.className = 'markup-layer';
    frame.append(layer);
  }
  return layer;
}

function clearMarkupOverlays() {
  for (const frame of pagesContainer.children) {
    const layer = frame.querySelector('.markup-layer');
    if (layer) layer.replaceChildren();
  }
}

async function deleteMarkupItem(item, rectEl) {
  if (!confirm('이 마크업을 삭제할까요?')) return;
  try {
    await fetchJsonView(`/api/view/markup/${item.id}?${tokenQuery}`, { method: 'DELETE' });
    markup.items = markup.items.filter((it) => it.id !== item.id);
    rectEl.remove();
    updateMarkupIncludeVisibility();
  } catch (error) {
    alert(error.message || '삭제에 실패했습니다.');
  }
}

function pageHeightPtFor(pageIdx) {
  const dims = pageDimensions && pageDimensions[pageIdx];
  return dims ? dims.h : 792;
}

const FONT_SIZE_MIN_PT = 4;
const FONT_SIZE_MAX_PT = 200;
const FONT_SIZE_STEP_PT = 1;

function clampFontSizePt(pt) {
  return Math.min(FONT_SIZE_MAX_PT, Math.max(FONT_SIZE_MIN_PT, Math.round(pt)));
}

function applyFontSizeToEl(el, frame, page, fontSizePt) {
  el.style.fontSize = `${fontSizePt * (frame.clientHeight / pageHeightPtFor(page - 1))}px`;
}

async function persistFontSize(item, fontSizePt) {
  try {
    const updated = await fetchJsonView(`/api/view/markup/${item.id}?${tokenQuery}`, {
      method: 'PATCH',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ x: item.x, y: item.y, fontSizePt }),
    });
    item.fontSizePt = updated.fontSizePt;
  } catch (error) {
    alert(error.message || '글자 크기 변경에 실패했습니다.');
  }
}

/** 클릭/휠 조작마다 매번 저장 요청을 보내지 않도록 항목별로 살짝 늦춰서 한 번만 저장한다. */
function scheduleFontSizePersist(el, item, fontSizePt) {
  clearTimeout(el._resizeTimer);
  el._resizeTimer = setTimeout(() => persistFontSize(item, fontSizePt), 350);
}

let textDragState = null;

function startTextDrag(event, item, el, frame) {
  if (event.target.closest('.text-overlay-delete')) return;
  event.preventDefault();
  event.stopPropagation();
  textDragState = { item, el, frame, moved: false, x: item.x, y: item.y };
  el.classList.add('dragging');
}

document.addEventListener('mousemove', (event) => {
  if (!textDragState) return;
  textDragState.moved = true;
  const { frame, el } = textDragState;
  const rect = frame.getBoundingClientRect();
  const x = Math.min(Math.max((event.clientX - rect.left) / rect.width, 0), 1);
  const y = Math.min(Math.max((event.clientY - rect.top) / rect.height, 0), 1);
  el.style.left = `${x * 100}%`;
  el.style.top = `${y * 100}%`;
  textDragState.x = x;
  textDragState.y = y;
});

document.addEventListener('mouseup', async () => {
  if (!textDragState) return;
  const { item, el, moved, x, y } = textDragState;
  textDragState = null;
  el.classList.remove('dragging');
  if (!moved) return;
  try {
    const updated = await fetchJsonView(`/api/view/markup/${item.id}?${tokenQuery}`, {
      method: 'PATCH',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ x, y }),
    });
    item.x = updated.x;
    item.y = updated.y;
  } catch (error) {
    // 실패하면 원래 위치로 되돌린다
    el.style.left = `${item.x * 100}%`;
    el.style.top = `${item.y * 100}%`;
    alert(error.message || '위치 이동에 실패했습니다.');
  }
});

/** 서버가 감지한 글꼴 계열을 화면에서도 비슷하게 보이도록 CSS 폰트로 매핑한다
 * (실제 PDF에 굽는 폰트와 맞춘다 — monospace는 내장 폰트가 없어 sans-serif로 근사). */
function cssFontFamilyFor(fontFamily) {
  if (fontFamily === 'serif') return '"Nanum Myeongjo", "Noto Serif KR", serif';
  return '"Pretendard Variable", "Pretendard", sans-serif';
}

function renderTextOverlayItem(frame, layer, item) {
  const el = document.createElement('div');
  el.className = 'text-overlay-item';
  el.textContent = item.text;
  el.dataset.markupId = item.id;
  el.style.left = `${item.x * 100}%`;
  el.style.top = `${item.y * 100}%`;
  el.style.fontSize = `${item.fontSizePt * (frame.clientHeight / pageHeightPtFor(item.page - 1))}px`;
  el.style.fontWeight = item.bold ? '700' : '400';
  el.style.fontFamily = cssFontFamilyFor(item.fontFamily);

  if (permissions.canEdit) {
    const controls = document.createElement('div');
    controls.className = 'text-overlay-controls';

    const resizeBy = (delta) => {
      const next = clampFontSizePt(item.fontSizePt + delta);
      item.fontSizePt = next;
      applyFontSizeToEl(el, frame, item.page, next);
      scheduleFontSizePersist(el, item, next);
    };

    const shrink = document.createElement('button');
    shrink.type = 'button';
    shrink.className = 'text-overlay-resize';
    shrink.textContent = '−';
    shrink.title = '글자 작게';
    shrink.addEventListener('click', (event) => {
      event.stopPropagation();
      resizeBy(-FONT_SIZE_STEP_PT);
    });

    const grow = document.createElement('button');
    grow.type = 'button';
    grow.className = 'text-overlay-resize';
    grow.textContent = '+';
    grow.title = '글자 크게';
    grow.addEventListener('click', (event) => {
      event.stopPropagation();
      resizeBy(FONT_SIZE_STEP_PT);
    });

    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'text-overlay-delete';
    del.textContent = '×';
    del.title = '삭제';
    del.addEventListener('click', (event) => {
      event.stopPropagation();
      deleteMarkupItem(item, el);
    });

    controls.append(shrink, grow, del);
    el.append(controls);
    el.addEventListener('mousedown', (event) => startTextDrag(event, item, el, frame));
    el.addEventListener(
      'wheel',
      (event) => {
        if (!event.ctrlKey) return;
        event.preventDefault();
        event.stopPropagation();
        resizeBy(event.deltaY < 0 ? FONT_SIZE_STEP_PT : -FONT_SIZE_STEP_PT);
      },
      { passive: false },
    );
  }

  layer.append(el);
}

function renderMarkupItem(item) {
  const frame = pagesContainer.children[item.page - 1];
  if (!frame) return;
  const layer = markupLayer(frame);
  if (item.kind === 'text') {
    renderTextOverlayItem(frame, layer, item);
    return;
  }
  const rectEl = document.createElement('div');
  rectEl.className = 'markup-rect';
  rectEl.style.left = `${item.x * 100}%`;
  rectEl.style.top = `${item.y * 100}%`;
  rectEl.style.width = `${item.w * 100}%`;
  rectEl.style.height = `${item.h * 100}%`;
  rectEl.tabIndex = 0;
  rectEl.dataset.markupId = item.id;

  const tooltip = document.createElement('div');
  tooltip.className = 'markup-tooltip';
  const metaLine = document.createElement('div');
  metaLine.className = 'markup-tooltip-meta';
  metaLine.textContent = `${item.author} · ${new Date(item.createdAt).toLocaleString('ko-KR')}`;
  const body = document.createElement('div');
  body.className = 'markup-tooltip-text';
  body.textContent = item.text;
  tooltip.append(metaLine, body);

  if (permissions.canEdit) {
    const del = document.createElement('button');
    del.type = 'button';
    del.className = 'markup-tooltip-delete';
    del.textContent = '삭제';
    del.addEventListener('click', (event) => {
      event.stopPropagation();
      deleteMarkupItem(item, rectEl);
    });
    tooltip.append(del);
  }

  rectEl.append(tooltip);
  layer.append(rectEl);
}

function renderAllMarkup() {
  clearMarkupOverlays();
  if (!markup.visible) return;
  for (const item of markup.items) renderMarkupItem(item);
}

function updateMarkupIncludeVisibility() {
  document.getElementById('markup-include-wrap').hidden = markup.items.length === 0;
}

async function loadMarkup() {
  try {
    markup.items = await fetchJsonView(`/api/view/markup?${tokenQuery}`);
    renderAllMarkup();
    updateMarkupIncludeVisibility();
  } catch {
    /* 마크업 로딩 실패는 열람 자체를 막지 않는다 */
  }
}

function closeMarkupPopover() {
  const existing = document.querySelector('.markup-popover');
  if (existing) existing.remove();
}

function cancelMarkupDraft() {
  if (markup.dragging) {
    markup.dragging.rectEl.remove();
    markup.dragging = null;
  }
  if (markup.draftRect) {
    markup.draftRect.remove();
    markup.draftRect = null;
  }
  closeMarkupPopover();
}

function setMarkupDrawMode(on) {
  markup.drawMode = on;
  document.getElementById('markup-add-btn').setAttribute('aria-pressed', String(on));
  pagesContainer.classList.toggle('markup-drawing', on);
  if (on) setTextToolActive(false);
  if (!on) cancelMarkupDraft();
}

function openMarkupPopover(frame, page, box) {
  closeMarkupPopover();
  const popover = document.createElement('div');
  popover.className = 'markup-popover';
  popover.style.left = `${Math.min(box.x + box.w, 0.7) * 100}%`;
  popover.style.top = `${box.y * 100}%`;

  const textarea = document.createElement('textarea');
  textarea.placeholder = '코멘트를 입력하세요';
  textarea.maxLength = 2000;

  const actions = document.createElement('div');
  actions.className = 'markup-popover-actions';
  const cancelBtn = document.createElement('button');
  cancelBtn.type = 'button';
  cancelBtn.className = 'btn-pill';
  cancelBtn.textContent = '취소';
  const saveBtn = document.createElement('button');
  saveBtn.type = 'button';
  saveBtn.className = 'btn-pill btn-pill-primary';
  saveBtn.textContent = '저장';
  actions.append(cancelBtn, saveBtn);
  popover.append(textarea, actions);
  frame.append(popover);
  textarea.focus();

  cancelBtn.addEventListener('click', () => cancelMarkupDraft());

  saveBtn.addEventListener('click', async () => {
    const text = textarea.value.trim();
    if (!text) {
      textarea.focus();
      return;
    }
    saveBtn.disabled = true;
    try {
      const item = await fetchJsonView(`/api/view/markup?${tokenQuery}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ page, x: box.x, y: box.y, w: box.w, h: box.h, text }),
      });
      markup.items.push(item);
      cancelMarkupDraft();
      renderAllMarkup();
      updateMarkupIncludeVisibility();
    } catch (error) {
      alert(error.message || '마크업 저장에 실패했습니다.');
      saveBtn.disabled = false;
    }
  });
}

function setupMarkupTool() {
  const addBtn = document.getElementById('markup-add-btn');
  if (permissions.canEdit) {
    addBtn.hidden = false;
    addBtn.addEventListener('click', () => setMarkupDrawMode(!markup.drawMode));
  }

  pagesContainer.addEventListener('mousedown', (event) => {
    if (!markup.drawMode) return;
    const frame = event.target.closest('.page-frame');
    if (!frame) return;
    event.preventDefault();
    const rect = frame.getBoundingClientRect();
    const fx = (event.clientX - rect.left) / rect.width;
    const fy = (event.clientY - rect.top) / rect.height;
    const rectEl = document.createElement('div');
    rectEl.className = 'markup-rect markup-rect-draft';
    markupLayer(frame).append(rectEl);
    markup.dragging = {
      frame,
      page: [...pagesContainer.children].indexOf(frame) + 1,
      fx,
      fy,
      rectEl,
      box: null,
    };
  });

  document.addEventListener('mousemove', (event) => {
    if (!markup.dragging) return;
    const { frame, fx: startFx, fy: startFy, rectEl } = markup.dragging;
    const rect = frame.getBoundingClientRect();
    const curFx = Math.min(Math.max((event.clientX - rect.left) / rect.width, 0), 1);
    const curFy = Math.min(Math.max((event.clientY - rect.top) / rect.height, 0), 1);
    const x = Math.min(startFx, curFx);
    const y = Math.min(startFy, curFy);
    const w = Math.abs(curFx - startFx);
    const h = Math.abs(curFy - startFy);
    rectEl.style.left = `${x * 100}%`;
    rectEl.style.top = `${y * 100}%`;
    rectEl.style.width = `${w * 100}%`;
    rectEl.style.height = `${h * 100}%`;
    markup.dragging.box = { x, y, w, h };
  });

  document.addEventListener('mouseup', () => {
    if (!markup.dragging) return;
    const { frame, page, rectEl, box } = markup.dragging;
    markup.dragging = null;
    const MIN_SIZE = 0.01;
    if (!box || box.w < MIN_SIZE || box.h < MIN_SIZE) {
      rectEl.remove();
      return;
    }
    markup.draftRect = rectEl;
    openMarkupPopover(frame, page, box);
  });
}

/* ==================== 텍스트 추가(주변 폰트에 맞춰 삽입) ==================== */

// lastFontSizePt: 마우스를 댈 때 보여 줄 줄 높이 — 마지막으로 넣은 글의 크기(처음엔 12pt)
const textTool = { active: false, pending: null, lastFontSizePt: 12, ghost: null, hint: null };

function setTextToolActive(on) {
  textTool.active = on;
  document.getElementById('text-tool-btn').setAttribute('aria-pressed', String(on));
  pagesContainer.classList.toggle('text-tool-active', on);
  if (on) setMarkupDrawMode(false);
  if (!on) commitTextInsert(); // 도구를 끌 때 쓰던 글은 저장한다(버리지 않는다)
  if (textTool.ghost) textTool.ghost.hidden = true;
  // 도구를 고른 동안 아래쪽에 쓰는 법을 띄운다 — 글을 넣은 뒤에도 도구가 남아 있어 지금 상태를 알 수 있게
  if (!textTool.hint) {
    textTool.hint = Object.assign(document.createElement('div'), {
      className: 'text-tool-hint',
      textContent: '텍스트 추가 — 누른 곳에 바로 입력 · Enter 로 마침 · 다 쓰면 Esc',
    });
    textTool.hint.setAttribute('role', 'status');
    document.body.append(textTool.hint);
  }
  textTool.hint.hidden = !on;
}

/** 마우스를 댄 자리에 글이 들어갈 줄(입력 막대 + '누르고 입력')을 미리 보인다 — 글은 누른 점 위에 앉는다 */
function showTextGhost(event) {
  const frame = textTool.active && !textTool.pending ? event.target.closest('.page-frame') : null;
  if (!frame || event.target.closest('.text-overlay-item, .text-insert-wrapper, .markup-rect')) {
    if (textTool.ghost) textTool.ghost.hidden = true;
    return;
  }
  if (!textTool.ghost) textTool.ghost = Object.assign(document.createElement('div'), { className: 'text-tool-ghost' });
  const ghost = textTool.ghost;
  if (ghost.parentElement !== frame) frame.append(ghost);
  const page = [...pagesContainer.children].indexOf(frame) + 1;
  const rect = frame.getBoundingClientRect();
  applyFontSizeToEl(ghost, frame, page, textTool.lastFontSizePt);
  const h = parseFloat(ghost.style.fontSize) * 1.15 + 2;
  ghost.style.left = `${event.clientX - rect.left}px`;
  ghost.style.top = `${event.clientY - rect.top - h}px`;
  ghost.style.height = `${h}px`;
  ghost.hidden = false;
}

function cancelTextInsert() {
  if (textTool.pending) {
    textTool.pending.wrapper.remove();
    textTool.pending = null;
  }
}

function adjustPendingFontSize(delta) {
  const pending = textTool.pending;
  if (!pending) return;
  pending.style.fontSizePt = clampFontSizePt(pending.style.fontSizePt + delta);
  applyFontSizeToEl(pending.box, pending.frame, pending.page, pending.style.fontSizePt);
  pending.box.focus();
}

async function commitTextInsert() {
  const pending = textTool.pending;
  if (!pending) return;
  textTool.pending = null;
  const text = pending.box.value.trim();
  pending.wrapper.remove();
  if (!text) return;
  textTool.lastFontSizePt = pending.style.fontSizePt;
  try {
    const item = await fetchJsonView(`/api/view/markup?${tokenQuery}`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({
        kind: 'text',
        page: pending.page,
        x: pending.xRatio,
        y: pending.style.baselineY,
        fontSizePt: pending.style.fontSizePt,
        bold: pending.style.bold,
        fontFamily: pending.style.fontFamily,
        text,
      }),
    });
    markup.items.push(item);
    renderAllMarkup();
    updateMarkupIncludeVisibility();
  } catch (error) {
    alert(error.message || '텍스트 추가에 실패했습니다.');
  }
}

function openTextInsertBox(frame, page, xRatio, style) {
  cancelTextInsert();
  const wrapper = document.createElement('div');
  wrapper.className = 'text-insert-wrapper';
  wrapper.style.left = `${xRatio * 100}%`;
  wrapper.style.top = `${style.baselineY * 100}%`;

  const box = document.createElement('textarea');
  box.className = 'text-insert-input';
  box.rows = 1;
  box.maxLength = 2000;
  box.style.fontWeight = style.bold ? '700' : '400';
  box.style.fontFamily = cssFontFamilyFor(style.fontFamily);
  box.style.width = '2ch';
  applyFontSizeToEl(box, frame, page, style.fontSizePt);

  const controls = document.createElement('div');
  controls.className = 'text-overlay-controls text-insert-controls';
  const shrink = document.createElement('button');
  shrink.type = 'button';
  shrink.className = 'text-overlay-resize';
  shrink.textContent = '−';
  shrink.title = '글자 작게';
  shrink.addEventListener('mousedown', (event) => event.preventDefault());
  shrink.addEventListener('click', () => adjustPendingFontSize(-FONT_SIZE_STEP_PT));
  const grow = document.createElement('button');
  grow.type = 'button';
  grow.className = 'text-overlay-resize';
  grow.textContent = '+';
  grow.title = '글자 크게';
  grow.addEventListener('mousedown', (event) => event.preventDefault());
  grow.addEventListener('click', () => adjustPendingFontSize(FONT_SIZE_STEP_PT));
  controls.append(shrink, grow);

  wrapper.append(box, controls);
  frame.append(wrapper);
  box.focus();

  box.addEventListener('input', () => {
    box.style.width = `${Math.max(2, box.value.length + 1)}ch`;
  });
  box.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      commitTextInsert();
    } else if (event.key === 'Escape') {
      // 첫 Esc 는 쓰던 글을 저장하고 마친다(도구는 그대로) — 두 번째 Esc 가 도구를 끈다(전역 단축키)
      event.preventDefault();
      event.stopPropagation();
      commitTextInsert();
    }
  });
  box.addEventListener('blur', () => commitTextInsert());
  box.addEventListener(
    'wheel',
    (event) => {
      if (!event.ctrlKey) return;
      event.preventDefault();
      adjustPendingFontSize(event.deltaY < 0 ? FONT_SIZE_STEP_PT : -FONT_SIZE_STEP_PT);
    },
    { passive: false },
  );

  textTool.pending = { wrapper, box, frame, page, xRatio, style };
}

function setupTextTool() {
  const btn = document.getElementById('text-tool-btn');
  if (!permissions.canEdit) return; // 버튼은 hidden 상태를 유지한다
  btn.hidden = false;
  btn.addEventListener('click', () => setTextToolActive(!textTool.active));
  pagesContainer.addEventListener('mousemove', (event) => showTextGhost(event));
  pagesContainer.addEventListener('mouseleave', () => textTool.ghost && (textTool.ghost.hidden = true));

  pagesContainer.addEventListener('click', async (event) => {
    if (!textTool.active) return;
    if (event.target.closest('.text-overlay-item') || event.target.closest('.text-insert-wrapper')) return;
    const frame = event.target.closest('.page-frame');
    if (!frame) return;
    // 쓰던 글이 있으면 저장하고, 누른 자리에서 바로 새 글을 시작한다(빈칸 여러 곳을 연달아 채우게)
    if (textTool.pending) commitTextInsert();
    if (textTool.ghost) textTool.ghost.hidden = true;
    const page = [...pagesContainer.children].indexOf(frame) + 1;
    const rect = frame.getBoundingClientRect();
    const xRatio = (event.clientX - rect.left) / rect.width;
    const yRatio = (event.clientY - rect.top) / rect.height;
    let style;
    try {
      style = await fetchJsonView(`/api/view/pages/${page}/text-style?${tokenQuery}&x=${xRatio}&y=${yRatio}`);
    } catch {
      style = { fontSizePt: 12, bold: false, fontFamily: 'sans-serif', baselineY: yRatio, matched: false };
    }
    openTextInsertBox(frame, page, xRatio, style);
  });
}

function setupDrawingTools() {
  // Ctrl+휠: 커서 기준 확대/축소
  pagesContainer.addEventListener(
    'wheel',
    (event) => {
      if (!event.ctrlKey) return;
      event.preventDefault();
      zoomAtPoint(zoomIndex + (event.deltaY < 0 ? 1 : -1), event.clientX, event.clientY);
    },
    { passive: false },
  );

  // 더블클릭: 해당 지점으로 두 단계 확대 (측정 중에는 비활성)
  pagesContainer.addEventListener('dblclick', (event) => {
    if (measure.mode) return;
    zoomAtPoint(zoomIndex + 2, event.clientX, event.clientY);
  });

  // 드래그 패닝 (왼쪽/휠 버튼) — AutoCAD의 팬처럼 문서를 끌어서 이동
  let pan = null;
  pagesContainer.addEventListener('mousedown', (event) => {
    if (measure.mode) return;
    if (event.button !== 0 && event.button !== 1) return;
    // 컨테이너 스크롤바 영역 클릭은 네이티브 동작(스크롤바 드래그)에 맡긴다
    if (
      event.target === pagesContainer &&
      (event.offsetX > pagesContainer.clientWidth || event.offsetY > pagesContainer.clientHeight)
    ) {
      return;
    }
    pan = {
      x: event.clientX,
      y: event.clientY,
      left: pagesContainer.scrollLeft,
      top: window.scrollY,
    };
    pagesContainer.classList.add('panning');
    event.preventDefault();
  });
  document.addEventListener('mousemove', (event) => {
    if (!pan) return;
    pagesContainer.scrollLeft = pan.left - (event.clientX - pan.x);
    window.scrollTo(window.scrollX, pan.top - (event.clientY - pan.y));
  });
  document.addEventListener('mouseup', () => {
    pan = null;
    pagesContainer.classList.remove('panning');
  });

  // 측정 모드
  document.getElementById('measure-btn').addEventListener('click', () => setMeasureMode(!measure.mode));
  document.getElementById('measure-close').addEventListener('click', () => setMeasureMode(false));
  document.getElementById('measure-clear').addEventListener('click', clearMeasurements);
  document.getElementById('measure-calibrate').addEventListener('click', () => {
    measure.calibrating = true;
    document.getElementById('measure-hint').textContent =
      '실제 길이를 아는 구간(치수선 등)의 양 끝을 클릭하세요';
  });

  pagesContainer.addEventListener('click', (event) => {
    if (!measure.mode) return;
    const frame = event.target.closest('.page-frame');
    if (!frame) return;
    const rect = frame.getBoundingClientRect();
    const point = {
      frame,
      fx: (event.clientX - rect.left) / rect.width,
      fy: (event.clientY - rect.top) / rect.height,
    };
    if (!measure.pending || measure.pending.frame !== frame) {
      if (measure.pending) measure.pending.dot.remove();
      measure.pending = { ...point, dot: measureDot(frame, point.fx, point.fy) };
    } else {
      addMeasurement(measure.pending, point);
      measure.pending = null;
    }
  });
}

/** 커서 위치를 고정한 채 확대/축소한다 — pages 뷰의 zoomAtPoint와 같은 목적이지만 svg-scroll
 * 자체가 스크롤 컨테이너라(window가 아니라) scrollLeft/Top만 다루면 된다. */
function svgZoomAtPoint(nextIndex, clientX, clientY) {
  const bounded = Math.max(0, Math.min(ZOOM_STEPS.length - 1, nextIndex));
  if (bounded === zoomIndex) return;
  const ratio = ZOOM_STEPS[bounded] / ZOOM_STEPS[zoomIndex];
  const rect = svgScroll.getBoundingClientRect();
  const contentX = svgScroll.scrollLeft + (clientX - rect.left);
  const contentY = svgScroll.scrollTop + (clientY - rect.top);
  zoomIndex = bounded;
  applySvgTransform();
  svgScroll.scrollLeft = contentX * ratio - (clientX - rect.left);
  svgScroll.scrollTop = contentY * ratio - (clientY - rect.top);
}

/** 드래그로 도면을 끌어서 이동 — pages 뷰의 setupDrawingTools() 패닝과 달리 svg-scroll 자체가
 * 스크롤 컨테이너라(pages는 window가 세로 스크롤을 맡는 구조라 서로 다르다) scrollLeft/Top만
 * 다루면 된다. */
function setupSvgPan() {
  // Ctrl+휠: 커서 기준 확대/축소 (pages 뷰와 같은 단축키)
  svgScroll.addEventListener(
    'wheel',
    (event) => {
      if (!event.ctrlKey) return;
      event.preventDefault();
      svgZoomAtPoint(zoomIndex + (event.deltaY < 0 ? 1 : -1), event.clientX, event.clientY);
    },
    { passive: false },
  );

  let pan = null;
  svgScroll.addEventListener('mousedown', (event) => {
    if (event.button !== 0 && event.button !== 1) return;
    pan = { x: event.clientX, y: event.clientY, left: svgScroll.scrollLeft, top: svgScroll.scrollTop };
    svgScroll.classList.add('panning');
    event.preventDefault();
  });
  document.addEventListener('mousemove', (event) => {
    if (!pan) return;
    svgScroll.scrollLeft = pan.left - (event.clientX - pan.x);
    svgScroll.scrollTop = pan.top - (event.clientY - pan.y);
  });
  document.addEventListener('mouseup', () => {
    pan = null;
    svgScroll.classList.remove('panning');
  });
}

/** 도면(dxf/dwg) 벡터 뷰 — LibreOffice가 만든 SVG를 그대로 주입해 보여준다. 페이지 개념이 없는
 * 단일 도면이라 페이지 이동 컨트롤은 숨기고, 확대/축소/회전은 pages 뷰와 같은 버튼을 그대로
 * 쓰되 svg-canvas에 transform을 걸어 처리한다(applySvgTransform). */
async function initSvgView() {
  document.getElementById('pages').hidden = true;
  svgView.hidden = false;
  document.getElementById('page-jump-input').hidden = true;
  pageJumpTotal.hidden = true;
  document.getElementById('page-view-btn').hidden = true;
  // 마크업(리뷰 코멘트)은 pages 뷰의 .page-frame 좌표계에 얹는 구조라 svg 뷰엔 아직 없다.
  document.getElementById('markup-add-btn').hidden = true;
  document.getElementById('text-tool-btn').hidden = true;
  document.getElementById('markup-include-wrap').hidden = true;

  const svgRes = await fetch(`/api/view/svg?${tokenQuery}`);
  if (!svgRes.ok) {
    const body = await svgRes.json().catch(() => null);
    throw new Error(body?.error || '도면을 불러오지 못했습니다.');
  }
  svgCanvas.innerHTML = await svgRes.text();

  // 확대 1배 상태의 실제 렌더 크기를 기준값으로 고정해 둔다(이후 fit 계산에 사용).
  const rect = svgCanvas.getBoundingClientRect();
  svgNaturalSize = { w: rect.width, h: rect.height };

  applySvgTransform();
  setupZoomFit();
  setupRotate();
  setupFullscreen();
  setupSvgPan();
}

async function init() {
  if (!token) {
    showError('열람 토큰이 없습니다. 발급받은 링크로 접속해 주세요.');
    return;
  }
  let res;
  try {
    res = await fetch(`/api/view/info?${tokenQuery}`);
  } catch {
    showError('서버에 연결할 수 없습니다.');
    return;
  }
  const json = await res.json();
  if (!res.ok || !json.success) {
    showError(json.error || '문서를 열 수 없습니다.');
    return;
  }
  // 쪽 문서는 통합 보기 껍데기(view.html)로 — 예전에 받아 둔 viewer.html 주소도 같은 화면으로 열린다(같은 조회 문자열).
  // 기준은 서버 viewerUrl(viewPathFor) 하나라 view.html과 서로 넘기지 않는다.
  if (json.data.viewerUrl && new URL(json.data.viewerUrl, location.href).pathname === '/view.html') {
    location.replace(`/view.html${location.search}`);
    return;
  }

  docTitle.textContent = json.data.name;
  document.title = `DocView — ${json.data.name}`;
  docName = json.data.name;
  pageCount = json.data.pageCount;
  pageDimensions = json.data.pages || null;
  permissions = json.data.permissions;
  currentViewKind = json.data.viewKind || 'pages';
  applyWatermark(json.data.viewer);

  // 측정/도면 모드는 도면·PDF 도구다 — word 등 다른 문서 종류를 pages 뷰로 볼 때는 의미가 없으므로
  // 실제 pdf 문서일 때만 보여준다 (grid 뷰는 initGridView가 이미 별도로 항상 숨긴다).
  if (json.data.ext !== 'pdf') {
    document.getElementById('measure-btn').hidden = true;
    document.getElementById('invert-btn').hidden = true;
  }

  setupToolbar();
  setupDeterrents();
  setupSearchBar();
  setupPresence();

  if (json.data.viewKind === 'grid') {
    try {
      await initGridView();
    } catch (error) {
      showError(error.message || '스프레드시트를 불러오지 못했습니다.');
    }
    return;
  }

  if (json.data.viewKind === 'svg') {
    try {
      await initSvgView();
    } catch (error) {
      showError(error.message || '도면을 불러오지 못했습니다.');
    }
    return;
  }

  setPageIndicator(1, pageCount);
  applyZoom();
  buildPages();
  trackCurrentPage();
  setupPageJump();
  setupZoomFit();
  setupRotate();
  setupFullscreen();
  setupDrawingTools();
  setupMarkupTool();
  setupTextTool();
  setupToolbarOverflow();
  loadMarkup();
}

init();
