/**
 * 편집기의 "보기로" 단추(설계 V7) — 통합 보기 껍데기의 편집 단추로 들어온 경우(from=view)에만 넣는다. 향기랩처럼 편집기를
 * 바로 연 곳에는 생기지 않는다. 자동 저장 편집기(워드·엑셀·PPT·PDF)는 쓰기 권한이 있으면 서버에 지금 저장(/api/view/save —
 * 저장 뒤 보기용 PDF까지 다시 만든다)한 뒤 이동하고, 한글은 beforeLeave(저장 안 한 내용 확인)를 쓴다.
 */
import { officeAlert } from './officeDialog.js';
import { carryViewParams, offersBackToView } from './view/logic.js';

export interface BackToViewOptions {
  readonly token: string;
  /** /api/view/info 의 viewerUrl */
  readonly viewerUrl: string | undefined;
  readonly canEdit: boolean;
  /** 떠나도 되는지(한글) — 없으면 자동 저장 편집기 규칙 */
  readonly beforeLeave?: () => Promise<boolean>;
}

/** host 안 before 앞에 단추를 넣는다. from=view가 아니면 아무것도 하지 않고 null */
export function mountBackToView(host: Element, before: Element | null, opts: BackToViewOptions): HTMLButtonElement | null {
  const viewerUrl = opts.viewerUrl;
  if (!viewerUrl || !offersBackToView(location.search, viewerUrl)) return null;
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.id = 'back-to-view';
  btn.className = 'btn-pill back-to-view';
  btn.textContent = '보기로';
  btn.title = '보기 화면으로 돌아가기';

  const flush = async (): Promise<boolean> => {
    if (!opts.canEdit) return true;
    btn.textContent = '저장하는 중…';
    const res = await fetch(`/api/view/save?token=${encodeURIComponent(opts.token)}`, { method: 'POST' }).catch(() => null);
    btn.textContent = '보기로';
    if (res?.ok) return true;
    await officeAlert('편집한 내용을 저장하지 못했습니다. 잠시 뒤 다시 눌러 주세요.');
    return false;
  };

  btn.addEventListener('click', async () => {
    btn.disabled = true;
    try {
      const go = opts.beforeLeave ? await opts.beforeLeave() : await flush();
      if (go) location.assign(carryViewParams(viewerUrl, location.search));
    } finally {
      btn.disabled = false;
    }
  });
  host.insertBefore(btn, before);
  return btn;
}
