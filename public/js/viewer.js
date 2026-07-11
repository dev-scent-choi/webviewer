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

/* ==================== 그리드 뷰 (스프레드시트) ==================== */

const GRID_ROW_H = 28;
const GRID_CHUNK = 200;
const GRID_OVERSCAN = 12;
const GRID_ROWNUM_W = 52;

const gridState = {
  sheets: [],
  styles: [],
  active: 0,
  chunks: new Map(), // `${sheet}:${chunkIdx}` -> rows | 'loading'
  mergeStarts: new Map(),
  mergeCovered: new Set(),
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
    if (style.col) parts.push(`color:${style.col}`);
    if (style.bg) parts.push(`background:${style.bg}`);
    if (style.a === 'c') parts.push('text-align:center');
    if (style.a === 'r') parts.push('text-align:right');
    if (style.w) parts.push('white-space:normal');
    return `.gs${i}{${parts.join(';')}}`;
  });
  const styleEl = document.createElement('style');
  styleEl.textContent = rules.join('\n');
  document.head.append(styleEl);
}

function buildMergeIndex(sheet) {
  gridState.mergeStarts = new Map();
  gridState.mergeCovered = new Set();
  for (const m of sheet.merges) {
    gridState.mergeStarts.set(`${m.r}:${m.c}`, m);
    for (let r = m.r; r < m.r + m.rs; r += 1) {
      for (let c = m.c; c < m.c + m.cs; c += 1) {
        if (r !== m.r || c !== m.c) gridState.mergeCovered.add(`${r}:${c}`);
      }
    }
  }
}

async function fetchChunk(sheetIdx, chunkIdx) {
  const key = `${sheetIdx}:${chunkIdx}`;
  if (gridState.chunks.has(key)) return;
  gridState.chunks.set(key, 'loading');
  try {
    const from = chunkIdx * GRID_CHUNK + 1;
    const data = await fetchJsonView(`/api/view/grid/rows?${tokenQuery}&sheet=${sheetIdx}&from=${from}&count=${GRID_CHUNK}`);
    gridState.chunks.set(key, data.rows);
    renderGridViewport();
  } catch {
    gridState.chunks.delete(key);
  }
}

async function fetchJsonView(url) {
  const res = await fetch(url);
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

function escapeHtml(text) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function renderGridViewport() {
  const sheet = gridState.sheets[gridState.active];
  if (!sheet) return;
  const scroll = document.getElementById('grid-scroll');
  const firstVisible = Math.max(1, Math.floor(scroll.scrollTop / GRID_ROW_H) + 1 - GRID_OVERSCAN);
  const visibleCount = Math.ceil(scroll.clientHeight / GRID_ROW_H) + GRID_OVERSCAN * 2;
  const lastVisible = Math.min(sheet.rowCount, firstVisible + visibleCount);

  // 필요한 청크를 비동기로 요청
  const firstChunk = Math.floor((firstVisible - 1) / GRID_CHUNK);
  const lastChunk = Math.floor((lastVisible - 1) / GRID_CHUNK);
  for (let c = firstChunk; c <= lastChunk; c += 1) fetchChunk(gridState.active, c);

  const cols = `<colgroup><col style="width:${GRID_ROWNUM_W}px">${sheet.colWidths
    .map((w) => `<col style="width:${w}px">`)
    .join('')}</colgroup>`;

  const rowsHtml = [];
  for (let r = firstVisible; r <= lastVisible; r += 1) {
    const cells = [`<td class="g-rownum">${r}</td>`];
    for (let c = 1; c <= sheet.colCount; c += 1) {
      if (gridState.mergeCovered.has(`${r}:${c}`)) continue;
      const merge = gridState.mergeStarts.get(`${r}:${c}`);
      const cell = cellAt(gridState.active, r, c - 1);
      const value = cell && cell !== 0 ? escapeHtml(cell.v) : '';
      const cls = cell && cell !== 0 && cell.s !== undefined ? ` class="gs${cell.s}"` : '';
      const span = merge ? ` colspan="${merge.cs}"${merge.rs > 1 ? ` rowspan="${merge.rs}"` : ''}` : '';
      cells.push(`<td${cls}${span}>${value}</td>`);
    }
    rowsHtml.push(`<tr style="height:${GRID_ROW_H}px">${cells.join('')}</tr>`);
  }

  const table = document.getElementById('grid-table');
  table.innerHTML = cols + rowsHtml.join('');
  table.style.transform = `translateY(${(firstVisible - 1) * GRID_ROW_H}px)`;

  pageIndicator.textContent = `${Math.min(sheet.rowCount, Math.floor(scroll.scrollTop / GRID_ROW_H) + 1).toLocaleString()} / ${sheet.rowCount.toLocaleString()}행`;
}

function renderGridHead() {
  const sheet = gridState.sheets[gridState.active];
  const head = document.getElementById('grid-head');
  const scroll = document.getElementById('grid-scroll');
  const cells = [`<div class="g-head-cell g-head-corner" style="width:${GRID_ROWNUM_W}px"></div>`];
  for (let c = 1; c <= sheet.colCount; c += 1) {
    cells.push(`<div class="g-head-cell" style="width:${sheet.colWidths[c - 1]}px">${colLetter(c)}</div>`);
  }
  head.innerHTML = `<div class="g-head-track">${cells.join('')}</div>`;
  head.firstChild.style.transform = `translateX(${-scroll.scrollLeft}px)`;
}

function activateSheet(index) {
  gridState.active = index;
  const sheet = gridState.sheets[index];
  buildMergeIndex(sheet);
  const inner = document.getElementById('grid-inner');
  inner.style.height = `${sheet.rowCount * GRID_ROW_H}px`;
  inner.style.width = `${GRID_ROWNUM_W + sheet.colWidths.reduce((a, b) => a + b, 0)}px`;
  document.getElementById('grid-scroll').scrollTop = 0;
  for (const [i, tab] of [...document.querySelectorAll('.sheet-tab')].entries()) {
    tab.classList.toggle('active', i === index);
  }
  renderGridHead();
  renderGridViewport();
}

async function initGridView() {
  document.getElementById('pages').hidden = true;
  document.getElementById('grid').hidden = false;
  document.getElementById('zoom-out').hidden = true;
  document.getElementById('zoom-in').hidden = true;
  document.getElementById('zoom-level').hidden = true;
  document.getElementById('invert-btn').hidden = true;

  const meta = await fetchJsonView(`/api/view/grid/meta?${tokenQuery}`);
  gridState.sheets = meta.sheets;
  injectGridStyles(meta.styles);

  const tabs = document.getElementById('sheet-tabs');
  tabs.replaceChildren(
    ...meta.sheets.map((sheet, i) => {
      const tab = document.createElement('button');
      tab.className = 'sheet-tab';
      tab.role = 'tab';
      tab.textContent = sheet.name;
      tab.addEventListener('click', () => activateSheet(i));
      return tab;
    }),
  );

  const scroll = document.getElementById('grid-scroll');
  let ticking = false;
  scroll.addEventListener('scroll', () => {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(() => {
      ticking = false;
      renderGridViewport();
      document.querySelector('.g-head-track').style.transform = `translateX(${-scroll.scrollLeft}px)`;
    });
  });

  // 복사 권한이 없으면 선택·복사·우클릭을 차단한다 (서버가 원본을 주지 않는 것이 근본 방어)
  if (!permissions.canCopy) {
    const grid = document.getElementById('grid');
    grid.classList.add('grid-noselect');
    grid.addEventListener('copy', (event) => event.preventDefault());
  }

  activateSheet(0);
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

  setupToolbar();
  setupDeterrents();

  if (json.data.viewKind === 'grid') {
    try {
      await initGridView();
    } catch (error) {
      showError(error.message || '스프레드시트를 불러오지 못했습니다.');
    }
    return;
  }

  pageIndicator.textContent = `1 / ${pageCount}`;
  applyZoom();
  buildPages();
  trackCurrentPage();
}

init();
