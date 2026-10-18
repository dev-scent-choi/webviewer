/**
 * 세 편집기(docx·xlsx·pptx)가 함께 쓰는 오피스식 화면 틀 — 앱 색 제목 표시줄(빠른 실행 도구 모음),
 * 리본 탭(파일·홈·삽입…·상황별 탭), 리본 축소(Ctrl+F1, 탭 두 번 클릭), "파일" 화면(정보·다운로드·
 * 인쇄·닫기). 리본 안의 실제 명령은 각 편집기가 그대로 맡는다.
 *
 * 마크업 약속: #toolbar 안에 .ribbon-tabs > button.ribbon-tab[data-tab], 탭마다
 * .ribbon-panel[data-panel] 하나. data-edit-only가 붙은 탭은 읽기 전용이면 숨긴다.
 */

export interface OfficeShellOptions {
  readonly token: string;
  readonly docName: string;
  readonly canEdit: boolean;
  readonly canDownload: boolean;
  readonly canPrint: boolean;
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
};

export function initOfficeShell(options: OfficeShellOptions): OfficeShell {
  const toolbar = document.getElementById('toolbar') as HTMLElement;
  // 앱 색(--app-accent)이 body에 붙어 있어야 body 밑에 따로 붙는 "파일" 화면·드롭다운도 같은 색을 쓴다.
  if (toolbar.dataset.app) document.body.dataset.app = toolbar.dataset.app;
  const tabs = [...toolbar.querySelectorAll<HTMLButtonElement>('.ribbon-tab')];
  const panels = [...toolbar.querySelectorAll<HTMLElement>('.ribbon-panel')];
  let activeTab = 'home';
  let collapsed = readCollapsed();
  let peeking = false;

  if (!options.canEdit) {
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
  if (!options.canEdit) {
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
            ? '<p>지금 편집 중인 내용 그대로 원본 형식 파일을 내려받습니다.</p><a class="backstage-action" data-action="download">원본 형식으로 다운로드</a>'
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

  backstage.querySelector('[data-action="print"]')?.addEventListener('click', async (event) => {
    const button = event.currentTarget as HTMLButtonElement;
    button.disabled = true;
    button.textContent = '저장하는 중…';
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
