/** 편집 단추(설계 V6) — 세션 응답에 편집기 주소가 있을 때만. 같은 창(iframe)에서 편집기로, 돌아올 수 있게 from=view */
import type { ShellCtx } from './ctx.js';
import { editHref } from './logic.js';

export function setupEditButton(ctx: Pick<ShellCtx, 'info'>): void {
  const editorUrl = ctx.info.editorUrl;
  const btn = document.getElementById('edit-btn') as HTMLButtonElement;
  if (!editorUrl) return;
  btn.hidden = false;
  btn.addEventListener('click', () => location.assign(editHref(editorUrl, location.search)));
}
