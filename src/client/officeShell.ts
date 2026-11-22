/**
 * 세 편집기(docx·xlsx·pptx)가 함께 쓰는 오피스식 화면 틀 — 앱 색 제목 표시줄(빠른 실행 도구 모음),
 * 리본 탭(파일·홈·삽입…·상황별 탭), 리본 축소(Ctrl+F1, 탭 두 번 클릭), "파일" 화면(정보·다운로드·
 * 인쇄·닫기). 리본 안의 실제 명령은 각 편집기가 그대로 맡는다.
 *
 * 마크업 약속: #toolbar 안에 .ribbon-tabs > button.ribbon-tab[data-tab], 탭마다
 * .ribbon-panel[data-panel] 하나. data-edit-only가 붙은 탭은 읽기 전용이면 숨긴다.
 */

import { mountBackToView } from './backToView.js';
import { officeAlert } from './officeDialog.js';
import { EDIT_DOC_KEY, installViewGuards, isContentUi, isViewUi, markReadOnly, viewControlAllowed } from './viewStrip.js';

export interface OfficeShellOptions {
  readonly token: string;
  /** /api/view/info 의 docId — 보기 꾸밈의 "편집 열기"가 같은 문서의 편집 토큰을 다시 받는다 */
  readonly docId: string;
  readonly docName: string;
  readonly canEdit: boolean;
  readonly canDownload: boolean;
  readonly canPrint: boolean;
  /** 보기 꾸밈(ui=view)에서만 쓴다 — false면 복사·선택·브라우저 우클릭 메뉴를 막는다(보기 화면 view.ts와 같은 규칙) */
  readonly canCopy: boolean;
  /** /api/view/info 의 viewerUrl — 보기 화면에서 들어왔으면(from=view) 제목 표시줄에 "보기로"(backToView.ts) */
  readonly viewerUrl?: string;
  /** 향기랩 옛 형식(.doc·.xls·.ppt)을 바꿔 연 문서 — 내려받기는 원래 형식(서버 /api/view/download), 매크로가 있으면 경고 한 줄 */
  readonly legacy?: { readonly ext: string; readonly hasMacros: boolean };
  readonly onUndo?: () => void;
  readonly onRedo?: () => void;
}

export interface OfficeShell {
  selectTab(tab: string): void;
  /** 상황별 탭(예: 표 안에 커서가 있을 때만 보이는 "표 레이아웃")을 보이거나 숨긴다. */
  setContextualTab(tab: string, visible: boolean): void;
}

const COLLAPSE_STORAGE_KEY = 'docview-ribbon-collapsed';

function readCollapsed(): boolean {
  try {
    return localStorage.getItem(COLLAPSE_STORAGE_KEY) === '1';
  } catch {
    return false;
  }
}

function writeCollapsed(collapsed: boolean): void {
  try {
    localStorage.setItem(COLLAPSE_STORAGE_KEY, collapsed ? '1' : '0');
  } catch {
    // 저장소를 못 쓰면 이번 화면에서만 유지된다
  }
}

/** 각 편집기는 Ctrl+S를 window keydown으로 처리한다(저장 중 표시 + 즉시 저장) — 같은 경로를 탄다. */
function triggerSave(): void {
  window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', ctrlKey: true, bubbles: true, cancelable: true }));
}

const EXT_LABELS: Readonly<Record<string, string>> = {
  docx: 'Word 문서',
  xlsx: 'Excel 통합 문서',
  pptx: 'PowerPoint 프레젠테이션',
  doc: 'Word 97-2003 문서',
  xls: 'Excel 97-2003 통합 문서',
  ppt: 'PowerPoint 97-2003 프레젠테이션',
};

/** 보기 꾸밈에서 흐리게 할 후보 — 리본 판·빠른 실행 안의 조작, 바꾸기 줄, 시트 추가. viewControlAllowed면 그대로 둔다 */
const VIEW_EDIT_CONTROLS = '.ribbon-panel :is(button, select, input, textarea), .quick-access button, #search-replace-input, #search-replace-one, #search-replace-all, #sheet-add';

function readEditDoc(): string | null {
  try {
    return sessionStorage.getItem(EDIT_DOC_KEY);
  } catch {
    return null;
  }
}

/** 메인 화면에서 연 업로더 본인만 — 같은 문서의 편집 토큰을 지금 받아 편집기로(API 키 규칙은 메인 화면과 같다) */
async function openEditFromView(docId: string): Promise<void> {
  const res = await fetch(`/api/documents/${encodeURIComponent(docId)}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '업로더', canPrint: true, canDownload: true, canCopy: true, canEdit: true }),
  }).catch(() => null);
  const json = (await res?.json().catch(() => null)) as { success?: boolean; data?: { editorUrl?: string }; error?: string } | null;
  if (!res?.ok || !json?.success || !json.data?.editorUrl) {
    await officeAlert(json?.error ?? '편집 화면을 열지 못했습니다. 잠시 뒤 다시 눌러 주세요.');
    return;
  }
  location.href = json.data.editorUrl;
}

export function initOfficeShell(options: OfficeShellOptions): OfficeShell {
  const toolbar = document.getElementById('toolbar') as HTMLElement;
  // 매크로는 고치지 않고 내려받으면 남고(같은 바이트), 고친 뒤 내려받으면 빠진다 — 고치기 전에 알린다(설계 A5)
  if (options.legacy?.hasMacros) {
    const note = document.createElement('div');
    note.className = 'unsupported-warning legacy-macro-note';
    note.setAttribute('role', 'status');
    note.textContent = '이 문서에는 매크로(VBA)가 있습니다. 고치지 않고 내려받으면 원본 그대로이지만, 고친 뒤 내려받으면 매크로가 빠집니다.';
    toolbar.after(note);
  }
  const titleRight = toolbar.querySelector('.titlebar-right');
  if (titleRight) {
    mountBackToView(titleRight, titleRight.querySelector('#presence-list'), { token: options.token, viewerUrl: options.viewerUrl, canEdit: options.canEdit });
  }
  // 앱 색(--app-accent)이 body에 붙어 있어야 body 밑에 따로 붙는 "파일" 화면·드롭다운도 같은 색을 쓴다.
  if (toolbar.dataset.app) document.body.dataset.app = toolbar.dataset.app;
  const tabs = [...toolbar.querySelectorAll<HTMLButtonElement>('.ribbon-tab')];
  const panels = [...toolbar.querySelectorAll<HTMLElement>('.ribbon-panel')];
  let activeTab = 'home';
  let collapsed = readCollapsed();
  let peeking = false;

  // 보기 꾸밈(ui=view)은 편집 탭도 보이게 두고 안의 단추만 흐리게 한다(아래) — 그 밖의 읽기 전용은 지금처럼 편집 탭을 숨긴다
  const viewUi = isViewUi(location.search, options.canEdit);
  if (!options.canEdit && !viewUi) {
    for (const tab of tabs) if (tab.dataset.editOnly !== undefined) tab.hidden = true;
  }

  function render(): void {
    for (const tab of tabs) {
      const selected = tab.dataset.tab === activeTab;
      tab.setAttribute('aria-selected', String(selected));
      tab.tabIndex = selected ? 0 : -1;
    }
    // 읽기 전용이라 보이는 편집 탭이 하나도 없으면 리본 판도 모두 숨긴다(파일 탭만 남는다).
    const activeVisible = tabs.some((tab) => tab.dataset.tab === activeTab && !tab.hidden);
    for (const panel of panels) panel.hidden = !activeVisible || panel.dataset.panel !== activeTab;
    toolbar.classList.toggle('ribbon-empty', !activeVisible);
    toolbar.classList.toggle('ribbon-collapsed', collapsed && !peeking);
    toolbar.classList.toggle('ribbon-peeking', collapsed && peeking);
    const collapseBtn = document.getElementById('ribbon-collapse');
    collapseBtn?.setAttribute('aria-pressed', String(collapsed));
    collapseBtn?.setAttribute('title', collapsed ? '리본 고정 (Ctrl+F1)' : '리본 축소 (Ctrl+F1)');
  }

  function firstVisibleTab(): string {
    return tabs.find((tab) => !tab.hidden && tab.dataset.tab !== 'file')?.dataset.tab ?? 'home';
  }

  function selectTab(tab: string): void {
    if (tab === 'file') {
      openBackstage();
      return;
    }
    const button = tabs.find((t) => t.dataset.tab === tab);
    if (!button || button.hidden) return;
    activeTab = tab;
    if (collapsed) peeking = true;
    render();
  }

  function toggleCollapsed(): void {
    collapsed = !collapsed;
    peeking = false;
    writeCollapsed(collapsed);
    render();
  }

  for (const tab of tabs) {
    tab.addEventListener('click', () => selectTab(tab.dataset.tab ?? 'home'));
    tab.addEventListener('dblclick', () => {
      if (tab.dataset.tab !== 'file') toggleCollapsed();
    });
    tab.addEventListener('keydown', (event) => {
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return;
      const visible = tabs.filter((t) => !t.hidden);
      const index = visible.indexOf(tab);
      const next = visible[(index + (event.key === 'ArrowRight' ? 1 : -1) + visible.length) % visible.length];
      next.focus();
      if (next.dataset.tab !== 'file') selectTab(next.dataset.tab ?? 'home');
      event.preventDefault();
    });
  }
  document.getElementById('ribbon-collapse')?.addEventListener('click', toggleCollapsed);
  window.addEventListener('keydown', (event) => {
    if ((event.ctrlKey || event.metaKey) && event.key === 'F1') {
      event.preventDefault();
      toggleCollapsed();
    }
  });
  // 축소된 리본을 탭으로 잠깐 펼쳤을 때(오피스와 같은 동작) — 리본 밖을 누르면 다시 접는다.
  document.addEventListener('mousedown', (event) => {
    if (peeking && !toolbar.contains(event.target as Node) && !(event.target as HTMLElement).closest('.dropdown-menu')) {
      peeking = false;
      render();
    }
  });

  // 다른 탭에 같은 명령이 또 있을 때(예: 삽입 > 새 슬라이드) — 원래 버튼을 대신 누른다.
  for (const proxy of toolbar.querySelectorAll<HTMLElement>('[data-proxy]')) {
    proxy.addEventListener('click', () => document.getElementById(proxy.dataset.proxy ?? '')?.click());
  }

  document.getElementById('qa-save')?.addEventListener('click', triggerSave);
  document.getElementById('qa-undo')?.addEventListener('click', () => options.onUndo?.());
  document.getElementById('qa-redo')?.addEventListener('click', () => options.onRedo?.());
  if (!options.canEdit && !viewUi) {
    for (const id of ['qa-save', 'qa-undo', 'qa-redo']) {
      const button = document.getElementById(id) as HTMLButtonElement | null;
      if (button) button.hidden = true;
    }
  }

  // ---------- "파일" 화면(백스테이지) ----------
  const ext = options.docName.split('.').pop()?.toLowerCase() ?? '';
  const homeHidden = document.body.classList.contains('home-hidden');
  const backstage = document.createElement('div');
  backstage.className = 'backstage';
  backstage.hidden = true;
  backstage.setAttribute('role', 'dialog');
  backstage.setAttribute('aria-label', '파일');
  backstage.innerHTML = `
    <nav class="backstage-nav">
      <button type="button" class="backstage-back" aria-label="문서로 돌아가기" title="문서로 돌아가기 (Esc)">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8l-4 4 4 4M8 12h8"/></svg>
      </button>
      <button type="button" class="backstage-item" data-pane="info">정보</button>
      <button type="button" class="backstage-item" data-pane="download">다운로드</button>
      <button type="button" class="backstage-item" data-pane="print">인쇄</button>
      <div class="backstage-nav-spacer"></div>
      ${homeHidden ? '' : '<button type="button" class="backstage-item" data-pane="close">닫기</button>'}
    </nav>
    <section class="backstage-content">
      <div class="backstage-pane" data-pane="info">
        <h2>정보</h2>
        <dl class="backstage-facts">
          <dt>이름</dt><dd class="backstage-doc-name"></dd>
          <dt>형식</dt><dd>${EXT_LABELS[ext] ?? ext.toUpperCase()}</dd>
          <dt>권한</dt><dd>${options.canEdit ? '편집 가능 — 변경 내용은 자동으로 저장됩니다' : '읽기 전용'}</dd>
        </dl>
      </div>
      <div class="backstage-pane" data-pane="download">
        <h2>다운로드</h2>
        ${
          options.canDownload
            ? options.legacy
              ? `<p>원래 형식(.${options.legacy.ext})으로 내려받습니다. 고치지 않았으면 올린 파일 그대로이고, 고쳤으면 LibreOffice로 다시 바꾸므로 모양이 조금 달라질 수 있습니다.</p><a class="backstage-action" data-action="download">${options.legacy.ext.toUpperCase()} 파일로 다운로드</a>`
              : '<p>지금 편집 중인 내용 그대로 원본 형식 파일을 내려받습니다.</p><a class="backstage-action" data-action="download">원본 형식으로 다운로드</a>'
            : '<p class="backstage-denied">이 문서는 다운로드가 허용되지 않았습니다.</p>'
        }
      </div>
      <div class="backstage-pane" data-pane="print">
        <h2>인쇄</h2>
        ${
          options.canPrint
            ? '<p>변경 내용을 저장한 뒤 인쇄용 PDF를 새 창에 엽니다.</p><button type="button" class="backstage-action" data-action="print">인쇄</button>'
            : '<p class="backstage-denied">이 문서는 인쇄가 허용되지 않았습니다.</p>'
        }
      </div>
    </section>`;
  (backstage.querySelector('.backstage-doc-name') as HTMLElement).textContent = options.docName;
  document.body.append(backstage);

  const downloadLink = backstage.querySelector<HTMLAnchorElement>('[data-action="download"]');
  if (downloadLink) downloadLink.href = `/api/view/download?token=${encodeURIComponent(options.token)}`;

  async function printDoc(): Promise<void> {
    // 새 창은 클릭과 같은 틱에 열어야 팝업 차단에 걸리지 않는다 — 저장이 끝나면 주소를 채운다.
    const printWindow = window.open('about:blank', '_blank');
    try {
      if (options.canEdit) await fetch(`/api/view/save?token=${encodeURIComponent(options.token)}`, { method: 'POST' });
    } catch {
      // 저장이 실패해도 마지막으로 저장된 상태를 인쇄한다
    }
    const url = `/api/view/print?token=${encodeURIComponent(options.token)}`;
    if (printWindow) printWindow.location.href = url;
    else window.open(url, '_blank');
  }
  backstage.querySelector('[data-action="print"]')?.addEventListener('click', async (event) => {
    const button = event.currentTarget as HTMLButtonElement;
    button.disabled = true;
    button.textContent = '저장하는 중…';
    await printDoc();
    button.disabled = false;
    button.textContent = '인쇄';
  });

  function showPane(pane: string): void {
    if (pane === 'close') {
      location.href = '/';
      return;
    }
    for (const item of backstage.querySelectorAll<HTMLElement>('.backstage-item')) {
      item.setAttribute('aria-current', String(item.dataset.pane === pane));
    }
    for (const el of backstage.querySelectorAll<HTMLElement>('.backstage-pane')) el.hidden = el.dataset.pane !== pane;
  }
  function openBackstage(): void {
    showPane('info');
    backstage.hidden = false;
    (backstage.querySelector('.backstage-item') as HTMLElement | null)?.focus();
  }
  function closeBackstage(): void {
    backstage.hidden = true;
    tabs.find((tab) => tab.dataset.tab === activeTab)?.focus();
  }
  backstage.querySelector('.backstage-back')?.addEventListener('click', closeBackstage);
  for (const item of backstage.querySelectorAll<HTMLElement>('.backstage-item')) {
    item.addEventListener('click', () => showPane(item.dataset.pane ?? 'info'));
  }
  backstage.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') closeBackstage();
  });

  // ---------- 보기 꾸밈(ui=view, 설계 2026-11-21-readonly-editor-look §3.1) ----------
  // 리본·빠른 실행·파일 화면은 그대로 보이고, 문서를 쓰는 조작만 흐리게(markReadOnly). 탭은 눌러 둘러볼 수 있다.
  if (viewUi) {
    document.body.classList.add('ui-view');
    document.body.classList.toggle('no-copy', !options.canCopy);
    installViewGuards(options.canCopy, options.canPrint);
    const greyOut = (): void => {
      for (const el of document.querySelectorAll<HTMLElement>(VIEW_EDIT_CONTROLS)) {
        // "자세히"·삽입 탭 대리 단추(data-proxy)는 대신 누르는 단추와 같은 판정
        if (!viewControlAllowed(el.dataset.proxy ?? el.id, options.canCopy)) markReadOnly(el);
      }
      // 편집기가 선택에 따라 disabled를 다시 풀어도 흐린 단추는 꺼진 채로
      for (const el of document.querySelectorAll<HTMLButtonElement>('.ro-disabled')) if ('disabled' in el && !el.disabled) el.disabled = true;
    };
    greyOut();
    // 나중에 다는 단추(PPT 디자인·전환 갤러리 등)까지 — 리본 안만 본다(격자·본문 변화는 보지 않는다)
    new MutationObserver(greyOut).observe(toolbar, { childList: true, subtree: true, attributes: true, attributeFilter: ['disabled'] });
    // 마지막 겹 — 흐린 조작에 닿은 클릭·바꿈은 편집기 처리기에 가기 전에 끊는다
    for (const type of ['click', 'mousedown', 'change', 'input'] as const) {
      document.addEventListener(
        type,
        (event) => {
          if (!(event.target instanceof Element) || !event.target.closest('.ro-disabled')) return;
          event.preventDefault();
          event.stopImmediatePropagation();
        },
        true,
      );
    }
    if (readEditDoc() === options.docId) titleRight?.insertBefore(editButton(), titleRight.querySelector('#presence-list'));
    const readOnlyItem = document.createElement('span');
    readOnlyItem.className = 'statusbar-item view-readonly';
    readOnlyItem.textContent = '읽기 전용';
    // 엑셀은 #statusbar 대신 격자 아래 .grid-status-row가 상태 줄이다
    document.querySelector('#statusbar, .grid-status-row')?.prepend(readOnlyItem);
  }

  // ---------- 콘텐츠만 보기(ui=content, 설계 2026-11-22-artcraft-app-screen §4) ----------
  // 도구는 CSS(body.ui-content)가 숨기고, 편집기는 이 화면에서 canEdit을 끄고 방에 readonly=1로 들어온다(편집기 입구). 여기서는 보기 규칙의 키·붙여넣기·우클릭 막기
  if (isContentUi(location.search)) {
    document.body.classList.toggle('no-copy', !options.canCopy);
    installViewGuards(options.canCopy, options.canPrint);
  }

  /** "편집 열기" — 메인 화면에서 연 경우만(공유 링크엔 없다) */
  function editButton(): HTMLElement {
    const edit = document.createElement('button');
    edit.type = 'button';
    edit.className = 'btn-pill back-to-view view-open-edit';
    edit.textContent = '편집 열기';
    edit.title = '편집할 수 있는 화면으로 다시 엽니다';
    edit.addEventListener('click', () => {
      edit.disabled = true;
      void openEditFromView(options.docId).finally(() => {
        edit.disabled = false;
      });
    });
    return edit;
  }

  if (!tabs.some((tab) => tab.dataset.tab === activeTab && !tab.hidden)) activeTab = firstVisibleTab();
  render();

  return {
    selectTab,
    setContextualTab(tab, visible) {
      const button = tabs.find((t) => t.dataset.tab === tab);
      if (!button || button.hidden === !visible) return;
      button.hidden = !visible;
      if (!visible && activeTab === tab) activeTab = 'home';
      render();
    },
  };
}
