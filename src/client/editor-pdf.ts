/**
 * PDF 편집기(editor-pdf.html) 시작 — pdfcraft 같은 배치(설계 docs/superpowers/specs/2026-10-20-pdf-editor-design.md).
 * 1단계: 보기·확대·회전·쪽 이동·썸네일·책갈피·레이어·찾기, 협업 방 연결(참여자 표시), 내려받기·인쇄.
 * 편집 내용(주석·쪽 정리·양식)은 다음 단계에서 이 방의 맵(src/pdf/pdfModel.ts)에 쌓는다.
 */
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import { openPdf, openPdfData, type PDFDocumentProxy } from './pdf/pdfDocument.js';
import { PageView } from './pdf/pageView.js';
import { setupPanels, type Panels } from './pdf/panels.js';
import { AnnotLayer } from './pdf/annotLayer.js';
import { setupAnnotUi } from './pdf/annotUi.js';
import { cleanPageList, PDF_ANNOTS_MAP, PDF_FORM_MAP, PDF_META_MAP, PDF_PAGES_ARRAY, PDF_REDACT_MAP, PDF_TEXT_EDITS_MAP, type PdfRedaction, type PdfTextEdit, type PdfAnnot, type PdfPageEntry } from '../pdf/annotTypes.js';
import { FindController } from './pdf/findController.js';
import { PageOps } from './pdf/pageOps.js';
import { setupOrganize, type OrganizeView } from './pdf/organizeView.js';
import { FormLayer } from './pdf/formLayer.js';
import { EditLayer } from './pdf/editLayer.js';
import { askPassword, setupDocTools, type DocTools } from './pdf/docTools.js';
import { setupSignUi, type SignUi } from './pdf/signUi.js';
import { pickPresenceColorForName, renderPresenceAvatars } from './presence.js';

interface ViewInfo {
  readonly docId: string;
  readonly name: string;
  readonly viewer: string;
  readonly permissions: { canPrint: boolean; canDownload: boolean; canCopy: boolean; canEdit: boolean; pdfEditor?: boolean };
}

const token = new URLSearchParams(location.search).get('token') ?? '';
const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

function showFatalError(message: string): void {
  const el = $('pdf-error');
  el.textContent = message;
  el.hidden = false;
}

async function fetchViewInfo(): Promise<ViewInfo> {
  const res = await fetch(`/api/view/info?token=${encodeURIComponent(token)}`);
  const json = (await res.json()) as { success: boolean; data?: ViewInfo; error?: string };
  if (!res.ok || !json.success || !json.data) throw new Error(json.error ?? '문서 정보를 불러오지 못했습니다.');
  return json.data;
}

function connect(info: ViewInfo): { ydoc: Y.Doc; provider: WebsocketProvider } {
  const ydoc = new Y.Doc();
  const wsProtocol = location.protocol === 'https:' ? 'wss:' : 'ws:';
  const provider = new WebsocketProvider(`${wsProtocol}//${location.host}/ws/edit`, info.docId, ydoc, { params: { token } });
  provider.awareness.setLocalStateField('user', { name: info.viewer, color: pickPresenceColorForName(info.viewer) });
  const conn = $('conn-status');
  provider.on('status', ({ status }: { status: string }) => {
    conn.dataset.state = status;
    conn.textContent = status === 'connected' ? (info.permissions.canEdit ? '연결됨' : '읽기 전용') : '연결 중…';
  });
  // 서버가 방 자체를 거부하면(4400대 — 잘못된 토큰 등) 다시 붙지 않는다. 보기는 원본만으로 되므로 오류 화면 대신 표시만 바꾼다.
  provider.on('connection-close', (event: CloseEvent | null) => {
    if (event && event.code >= 4400) {
      provider.shouldConnect = false;
      conn.dataset.state = 'disconnected';
      conn.textContent = '보기 전용';
      conn.title = event.reason || '';
    }
  });
  const presence = $('presence');
  provider.awareness.on('change', () => renderPresenceAvatars(presence, provider.awareness));
  return { ydoc, provider };
}

let noticeTimer = 0;
/** 화면 위(정리 화면에선 아래) 잠깐 뜨는 안내 */
function showNotice(text: string): void {
  const notice = $('mode-notice');
  notice.textContent = text;
  notice.hidden = false;
  clearTimeout(noticeTimer);
  noticeTimer = window.setTimeout(() => (notice.hidden = true), 2600);
}

function setupChrome(view: PageView, find: FindController, info: ViewInfo, panels: Panels, organize: OrganizeView, sign: SignUi, docTools: DocTools): void {
  // 오른쪽 줄 — 쪽 번호·이동·배율
  const pageInput = $<HTMLInputElement>('page-input');
  const pageCount = $('page-count');
  const zoomLevel = $('zoom-level');
  pageCount.textContent = `/ ${view.pageCount}`;
  view.onPagesChange(() => {
    pageCount.textContent = `/ ${view.pageCount}`;
    pageInput.value = String(view.currentPage);
  });
  view.onPageChange((page) => {
    if (document.activeElement !== pageInput) pageInput.value = String(page);
  });
  const commitPage = () => {
    const n = Number(pageInput.value);
    if (Number.isFinite(n) && n >= 1) view.goTo(Math.min(view.pageCount, n));
    pageInput.value = String(view.currentPage);
  };
  pageInput.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      commitPage();
      pageInput.blur();
    }
  });
  pageInput.addEventListener('change', commitPage);
  $('page-prev').addEventListener('click', () => view.goTo(view.currentPage - 1));
  $('page-next').addEventListener('click', () => view.goTo(view.currentPage + 1));
  const showZoom = (scale: number) => {
    zoomLevel.textContent = `${Math.round(scale * 100)}%`;
  };
  view.onScaleChange(showZoom);
  showZoom(view.scale);
  $('zoom-in').addEventListener('click', () => view.zoomIn());
  $('zoom-out').addEventListener('click', () => view.zoomOut());
  // 맞춤: 폭에 맞춤과 쪽에 맞춤을 번갈아
  let fitWidthNext = false;
  $('fit-btn').addEventListener('click', () => {
    view.setScale(fitWidthNext ? 'fit-width' : 'fit-page');
    fitWidthNext = !fitWidthNext;
  });
  $('rotate-btn').addEventListener('click', () => view.rotate(90));

  // 모드 탭 — 모든 도구(왼쪽 목록) / 읽기(본문만) / 편집·서명(다음 단계 안내)
  const toolsPanel = $('tools-panel');
  const menuBtn = $('menu-btn');
  const setToolsOpen = (open: boolean) => {
    toolsPanel.hidden = !open;
    toolsPanel.classList.toggle('forced-open', open && window.matchMedia('(max-width: 1100px)').matches);
    menuBtn.setAttribute('aria-pressed', String(open));
  };
  if (window.matchMedia('(max-width: 1100px)').matches) setToolsOpen(false);
  menuBtn.addEventListener('click', () => setToolsOpen(toolsPanel.hidden));
  $('tools-close').addEventListener('click', () => setToolsOpen(false));
  const modes = [...document.querySelectorAll<HTMLButtonElement>('.pdf-mode')];
  if (!info.permissions.canEdit) for (const m of modes) if (m.dataset.editOnly !== undefined) m.hidden = true;
  const selectMode = (mode: string) => {
    for (const m of modes) m.setAttribute('aria-selected', String(m.dataset.mode === mode));
    document.body.classList.toggle('mode-read', mode === 'read');
    document.body.classList.toggle('mode-edit', mode === 'edit');
    document.body.classList.toggle('mode-sign', mode === 'sign');
    sign.setActive(mode === 'sign');
    // 서명 — 왼쪽 목록을 접고 오른쪽에 양식 칸 목록
    if (mode === 'sign') {
      setToolsOpen(false);
      panels.show('fields');
    }
    if (mode === 'tools') setToolsOpen(true);
    // 편집 — 왼쪽 목록을 접고 오른쪽에 메모 목록(pdfcraft의 주석 달기와 같다)
    if (mode === 'edit') {
      setToolsOpen(false);
      panels.show('comments');
    }
  };
  for (const m of modes) m.addEventListener('click', () => selectMode(m.dataset.mode ?? 'tools'));
  for (const item of document.querySelectorAll<HTMLButtonElement>('.pdf-tool-item')) {
    // 읽기 전용이면 고치는 도구는 숨기고, 추출·나누기는 내려받기 권한이 있을 때만
    const tool = item.dataset.tool;
    const allowed =
      tool === 'organize' || tool === 'combine' || tool === 'comment' || tool === 'fill' || tool === 'protect' || tool === 'compress' ? info.permissions.canEdit : tool === 'extract' || tool === 'download' ? info.permissions.canDownload : true;
    if (!allowed) item.closest('li')!.hidden = true;
    item.addEventListener('click', () => {
      const tool = item.dataset.tool;
      if (tool === 'download') download();
      else if (tool === 'comment') selectMode('edit');
      else if (tool === 'fill') selectMode('sign');
      else if (tool === 'protect') docTools.openProtect();
      else if (tool === 'compress') docTools.openCompress();
      else if (tool === 'props') docTools.openProps();
      else if (tool === 'organize') organize.open();
      else if (tool === 'combine') organize.pickFiles();
      else if (tool === 'extract') organize.open('추출할 쪽을 고른 뒤 ‘고른 쪽 추출’ 또는 ‘나누기’를 누르세요.');
      else showNotice(`'${item.querySelector('span')?.textContent ?? ''}'은(는) 다음 업데이트에서 열립니다.`);
    });
  }

  // 손 도구(annotUi가 고르면 .hand) — 끌어서 이동
  const scroller = $('pdf-scroll');
  let drag: { x: number; y: number; left: number; top: number } | null = null;
  scroller.addEventListener('pointerdown', (event) => {
    if (!scroller.classList.contains('hand') || event.button !== 0) return;
    drag = { x: event.clientX, y: event.clientY, left: scroller.scrollLeft, top: scroller.scrollTop };
    scroller.classList.add('dragging');
    scroller.setPointerCapture(event.pointerId);
    event.preventDefault();
  });
  scroller.addEventListener('pointermove', (event) => {
    if (!drag) return;
    scroller.scrollLeft = drag.left - (event.clientX - drag.x);
    scroller.scrollTop = drag.top - (event.clientY - drag.y);
  });
  const endDrag = () => {
    drag = null;
    scroller.classList.remove('dragging');
  };
  scroller.addEventListener('pointerup', endDrag);
  scroller.addEventListener('pointercancel', endDrag);
  // Ctrl+휠 = 확대·축소(브라우저 확대 대신)
  scroller.addEventListener(
    'wheel',
    (event) => {
      if (!event.ctrlKey) return;
      event.preventDefault();
      if (event.deltaY < 0) view.zoomIn();
      else view.zoomOut();
    },
    { passive: false },
  );

  // 위 줄 찾기 상자 → 찾기 줄
  const topFind = $<HTMLInputElement>('top-find');
  topFind.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      find.open(topFind.value);
    }
  });

  // 내려받기·인쇄·문서 정보
  function download(): void {
    if (!info.permissions.canDownload) return;
    location.href = `/api/view/download?token=${encodeURIComponent(token)}`;
  }
  $('download-btn').addEventListener('click', download);
  const printBtn = $('print-btn');
  printBtn.hidden = !info.permissions.canPrint;
  printBtn.addEventListener('click', () => window.open(`/api/view/print?token=${encodeURIComponent(token)}`, '_blank'));
  $('info-btn').addEventListener('click', () => docTools.openProps());

  // 단축키
  document.addEventListener('keydown', (event) => {
    const mod = event.ctrlKey || event.metaKey;
    const typing = (event.target as HTMLElement).closest('input, textarea, [contenteditable="true"]');
    if (mod && (event.key === 'f' || event.key === 'F')) {
      event.preventDefault();
      find.open();
    } else if (mod && (event.key === 'p' || event.key === 'P') && info.permissions.canPrint) {
      event.preventDefault();
      printBtn.click();
    } else if (mod && (event.key === '=' || event.key === '+')) {
      event.preventDefault();
      view.zoomIn();
    } else if (mod && event.key === '-') {
      event.preventDefault();
      view.zoomOut();
    } else if (mod && event.key === '0') {
      event.preventDefault();
      view.setScale('fit-width');
    } else if (!typing && (event.key === 'PageDown' || event.key === 'PageUp')) {
      event.preventDefault();
      view.goTo(view.currentPage + (event.key === 'PageDown' ? 1 : -1));
    } else if (!typing && mod && (event.key === 'Home' || event.key === 'End')) {
      event.preventDefault();
      view.goTo(event.key === 'Home' ? 1 : view.pageCount);
    }
  });
}


let currentDoc: Awaited<ReturnType<typeof openPdf>> | null = null;

async function main(): Promise<void> {
  if (!token) {
    showFatalError('열람 토큰이 없습니다.');
    return;
  }
  let info: ViewInfo;
  try {
    info = await fetchViewInfo();
  } catch (error) {
    showFatalError(error instanceof Error ? error.message : String(error));
    return;
  }
  $('doc-name').textContent = info.name;
  document.title = `${info.name} — PDF 편집`;
  const { ydoc, provider } = connect(info);

  let base = '';
  try {
    // 암호 파일 — 서버가 423(잠김)을 주면 비밀번호 창을 띄우고, 맞히면 같은 화면에서 다시 받는다
    // (새로 고치면 방이 닫혔다 다시 잠긴 채로 열릴 수 있다)
    for (;;) {
      const res = await fetch(`/api/view/pdf-source?token=${encodeURIComponent(token)}`);
      if (res.status === 423) {
        $('pdf-loading').hidden = true;
        await askPassword(token);
        $('pdf-loading').hidden = false;
        continue;
      }
      if (!res.ok) throw new Error(`pdf-source ${res.status}`);
      base = res.headers.get('x-pdf-base') ?? '';
      currentDoc = await openPdfData(new Uint8Array(await res.arrayBuffer()));
      break;
    }
  } catch (error) {
    showFatalError('PDF를 열지 못했습니다.');
    console.error(error);
    return;
  }
  // 넣은 PDF(부속) — 쪽 목록에 'p:' 쪽이 있으면 받아서 연다(한 번만)
  const parts = new Map<string, Promise<PDFDocumentProxy>>();
  const loadPart = (partId: string): Promise<PDFDocumentProxy> => {
    let p = parts.get(partId);
    if (!p) {
      p = openPdf(`/api/view/pdf-part/${encodeURIComponent(partId)}?token=${encodeURIComponent(token)}`);
      parts.set(partId, p);
      p.catch(() => parts.delete(partId));
    }
    return p;
  };
  const view = new PageView($('pdf-pages'), $('pdf-scroll'), currentDoc, loadPart);
  await view.init();
  // 쪽 목록(방의 pdfPages) — 처음엔 원본 순서로 바로 그리고, 방과 맞춰지면 목록대로
  const pagesArr = ydoc.getArray<PdfPageEntry>(PDF_PAGES_ARRAY);
  const applyPages = () => {
    const list = cleanPageList(pagesArr.toArray());
    if (list.length > 0) view.setPages(list);
  };
  pagesArr.observe(applyPages);
  applyPages();
  // 받은 원본이 방을 열 때의 원본과 다르면(그 사이 다른 사람이 저장하고 방이 새로 열린 경우) 다시 연다
  const checkBase = () => {
    const roomBase = ydoc.getMap<string>(PDF_META_MAP).get('base');
    if (!base || !roomBase || roomBase === base) return;
    let reloaded = false;
    try {
      reloaded = sessionStorage.getItem('pdf-base-reload') === roomBase;
      sessionStorage.setItem('pdf-base-reload', roomBase);
    } catch {
      // 저장소를 못 쓰면 한 번만 시도한다는 보장 없이 그냥 다시 연다
    }
    if (!reloaded) location.reload();
  };
  provider.on('sync', (synced: boolean) => {
    if (synced) checkBase();
  });
  const find = new FindController(view, {
    bar: $('find-bar'),
    input: $<HTMLInputElement>('find-input'),
    count: $('find-count'),
    prev: $<HTMLButtonElement>('find-prev'),
    next: $<HTMLButtonElement>('find-next'),
    close: $<HTMLButtonElement>('find-close'),
  });
  const panels = setupPanels({ doc: currentDoc, view });
  // 주석 — 협업 맵(방을 열 때 서버가 파일 주석을 채운다)을 쪽마다 그리고 도구로 고친다
  const annotMap = ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP);
  const layer = new AnnotLayer({ view, map: annotMap, ydoc, canEdit: info.permissions.canEdit, author: info.viewer, scroller: $('pdf-scroll') });
  const annotUi = setupAnnotUi({ layer, map: annotMap, pages: pagesArr, view, canEdit: info.permissions.canEdit, onShowComments: () => panels.show('comments') });
  view.onPagesChange(() => {
    const sel = layer.selectedAnnot;
    if (sel && view.pageOfKey(sel.pageKey) === null) layer.select(null);
    layer.renderAll();
  });
  // 썸네일에도 주석
  panels.setThumbOverlay((page, host, vp) => layer.drawOverlay(page, host, vp));
  layer.onPagesChanged((pages) => panels.refreshThumbOverlays(pages));
  const organize = setupOrganize({
    view,
    ops: new PageOps(ydoc, pagesArr, annotMap),
    undoManager: annotUi.undoManager,
    token,
    docName: info.name,
    canEdit: info.permissions.canEdit,
    canDownload: info.permissions.canDownload,
    drawOverlay: (page, host, vp) => layer.drawOverlay(page, host, vp),
    notice: showNotice,
  });
  // 양식 칸(값은 방의 pdfForm) — 실행 취소에 함께
  const formMap = ydoc.getMap<unknown>(PDF_FORM_MAP);
  annotUi.undoManager.addToScope(formMap);
  let sign: SignUi | null = null;
  const forms = new FormLayer({
    view,
    map: formMap,
    ydoc,
    canEdit: info.permissions.canEdit,
    onSignatureField: (page, rect) => sign?.signField(page, rect),
    // 서명 그림의 가운데가 칸 안에 있으면 서명된 칸
    isSigned: (pageKey, rect) =>
      layer.all().some((a) => {
        if (a.type !== 'stamp' || a.pageKey !== pageKey) return false;
        const cx = (a.rect[0] + a.rect[2]) / 2;
        const cy = (a.rect[1] + a.rect[3]) / 2;
        return cx >= Math.min(rect[0], rect[2]) && cx <= Math.max(rect[0], rect[2]) && cy >= Math.min(rect[1], rect[3]) && cy <= Math.max(rect[1], rect[3]);
      }),
  });
  layer.onChange(() => forms.refreshSignatureFields());
  // 기존 글자 고치기·가리기 — 고치기는 실행 취소 대상, 가리기 적용은 되돌릴 수 없다
  const editsMap = ydoc.getMap<PdfTextEdit>(PDF_TEXT_EDITS_MAP);
  const redactMap = ydoc.getMap<PdfRedaction>(PDF_REDACT_MAP);
  annotUi.undoManager.addToScope([editsMap, redactMap]);
  const edits = new EditLayer({
    view,
    ydoc,
    edits: editsMap,
    redactions: redactMap,
    canEdit: info.permissions.canEdit,
    author: info.viewer,
    scroller: $('pdf-scroll'),
    undoManager: annotUi.undoManager,
    notice: showNotice,
  });
  const applyBtn = $<HTMLButtonElement>('redact-apply');
  const syncApply = (n: number) => {
    applyBtn.hidden = n === 0 || !info.permissions.canEdit;
    $('redact-count').textContent = String(n);
  };
  edits.onPendingChange(syncApply);
  syncApply(edits.pendingCount());
  const redactDialog = $('redact-dialog');
  applyBtn.addEventListener('click', () => {
    $('redact-dialog-text').textContent = `표시한 ${edits.pendingCount()}곳을 가립니다.`;
    redactDialog.hidden = false;
    $<HTMLButtonElement>('redact-cancel').focus();
  });
  $('redact-cancel').addEventListener('click', () => (redactDialog.hidden = true));
  redactDialog.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') redactDialog.hidden = true;
  });
  $('redact-confirm').addEventListener('click', () => {
    redactDialog.hidden = true;
    const n = edits.applyRedactions();
    if (n) showNotice(`${n}곳을 가렸습니다 — 저장 파일에서 그 내용이 지워집니다.`);
  });
  sign = setupSignUi({ view, layer, forms, doc: currentDoc, canEdit: info.permissions.canEdit, showPanel: (name) => panels.show(name), notice: showNotice });
  const docTools = setupDocTools({ token, ydoc, canEdit: info.permissions.canEdit, docName: info.name, pageCount: () => view.pageCount, notice: showNotice });
  setupChrome(view, find, info, panels, organize, sign, docTools);
  // 처음 배율: 폭에 맞추되 너무 크게(작은 쪽을 넓은 화면에) 키우지 않는다
  view.setScale('fit-width');
  if (view.scale > 1.25) view.setScale(1.25);
  await view.firstPaint();
  $('pdf-loading').hidden = true;
  // 향기랩 등 부모 창에 "다 그렸다"를 알린다(다른 편집기의 docview:ready와 같다)
  requestAnimationFrame(() => requestAnimationFrame(() => window.parent.postMessage({ type: 'docview:ready' }, '*')));
}

void main();
