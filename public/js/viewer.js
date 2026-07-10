'use strict';

const ZOOM_STEPS = [50, 67, 80, 100, 125, 150, 200, 250, 300, 400, 500];
const DEFAULT_ZOOM_INDEX = 3;

const token = new URLSearchParams(location.search).get('token');
const pagesContainer = document.getElementById('pages');
const docTitle = document.getElementById('doc-title');
const pageIndicator = document.getElementById('page-indicator');
const zoomLevelLabel = document.getElementById('zoom-level');
const printBtn = document.getElementById('print-btn');
const downloadBtn = document.getElementById('download-btn');

let zoomIndex = DEFAULT_ZOOM_INDEX;
let pageCount = 0;
let permissions = { canPrint: false, canDownload: false };
let docName = 'document';

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
  const label = button.querySelector('.btn-label') ?? button;
  const originalText = label.textContent;
  button.disabled = true;
  try {
    await task((received, total) => {
      label.textContent = `${progressPrefix} ${formatProgress(received, total)}`;
    });
    label.textContent = originalText;
  } catch (error) {
    label.textContent = '실패';
    setTimeout(() => {
      label.textContent = originalText;
    }, 2500);
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

const tokenQuery = `token=${encodeURIComponent(token)}`;

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
    const res = await fetch(`/api/view/pages/${pageNumber}?${tokenQuery}&scale=${scale}`);
    if (!res.ok) {
      canvas.closest('.page-frame').classList.add('page-error');
      return;
    }
    const bitmap = await createImageBitmap(await res.blob());
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
    const canvas = document.createElement('canvas');
    canvas.dataset.page = String(page);
    const label = document.createElement('span');
    label.className = 'page-number';
    label.textContent = String(page);
    frame.append(canvas, label);
    frames.push(frame);
    observer.observe(frame);
  }
  pagesContainer.replaceChildren(...frames);
}

function trackCurrentPage() {
  const observer = new IntersectionObserver((entries) => {
    const visible = entries.find((entry) => entry.isIntersecting);
    if (visible) {
      const page = visible.target.querySelector('canvas').dataset.page;
      pageIndicator.textContent = `${page} / ${pageCount}`;
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

function setupToolbar() {
  document.getElementById('zoom-in').addEventListener('click', () => {
    zoomIndex = Math.min(zoomIndex + 1, ZOOM_STEPS.length - 1);
    applyZoom();
  });
  document.getElementById('zoom-out').addEventListener('click', () => {
    zoomIndex = Math.max(zoomIndex - 1, 0);
    applyZoom();
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
      // 팝업 차단을 피하려면 사용자 클릭 스택 안에서 창을 먼저 열어둬야 한다
      const printWindow = window.open('', '_blank');
      runButtonTask(printBtn, '생성', async (onProgress) => {
        const chunks = await fetchWithProgress(`/api/view/print?${tokenQuery}`, onProgress);
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
      runButtonTask(downloadBtn, '저장', async (onProgress) => {
        const chunks = await fetchWithProgress(`/api/view/download?${tokenQuery}`, onProgress);
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
  });
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

  docTitle.textContent = json.data.name;
  document.title = `DocView — ${json.data.name}`;
  docName = json.data.name;
  pageCount = json.data.pageCount;
  permissions = json.data.permissions;
  pageIndicator.textContent = `1 / ${pageCount}`;

  applyZoom();
  setupToolbar();
  setupDeterrents();
  buildPages();
  trackCurrentPage();
}

init();
