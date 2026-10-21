/**
 * 브라우저 기본 창(alert·confirm·prompt) 대신 쓰는 앱 안 대화상자 — 오피스처럼 화면 안에 뜨고, 편집기의
 * .office-dialog 모양(Fluent 2: 제목·본문·주 단추)을 그대로 쓴다. 기본 창과 달리 기다리는 동안 화면이 멈추지 않으므로
 * 결과는 Promise로 받는다. Enter = 확인, Esc = 취소, 바깥(어두운 곳) 누르기 = 취소.
 */

interface DialogOptions {
  readonly title?: string;
  readonly okLabel?: string;
  readonly cancelLabel?: string;
  /** 지우기처럼 되돌리기 어려운 동작 — 확인 단추를 경고 색으로 */
  readonly danger?: boolean;
}

function open(
  message: string,
  kind: 'alert' | 'confirm' | 'prompt',
  opts: DialogOptions & { defaultValue?: string; maxLength?: number },
): Promise<{ ok: boolean; value: string }> {
  return new Promise((resolve) => {
    const previousFocus = document.activeElement as HTMLElement | null;
    const backdrop = document.createElement('div');
    backdrop.className = 'office-dialog office-dialog-runtime';
    backdrop.setAttribute('role', kind === 'alert' ? 'alertdialog' : 'dialog');
    backdrop.setAttribute('aria-modal', 'true');
    const card = document.createElement('div');
    card.className = 'office-dialog-card';
    const title = document.createElement('h2');
    title.textContent = opts.title ?? (kind === 'confirm' ? '확인' : kind === 'prompt' ? '입력' : '알림');
    title.id = `office-dialog-title-${Date.now()}`;
    backdrop.setAttribute('aria-labelledby', title.id);
    const body = document.createElement('p');
    body.className = 'office-dialog-message';
    body.textContent = message;
    card.append(title, body);
    let input: HTMLInputElement | null = null;
    if (kind === 'prompt') {
      input = document.createElement('input');
      input.type = 'text';
      input.className = 'office-dialog-input';
      input.value = opts.defaultValue ?? '';
      if (opts.maxLength) input.maxLength = opts.maxLength;
      card.append(input);
    }
    const actions = document.createElement('div');
    actions.className = 'office-dialog-actions';
    const okBtn = document.createElement('button');
    okBtn.type = 'button';
    okBtn.className = `btn-mini btn-mini-primary${opts.danger ? ' btn-mini-danger' : ''}`;
    okBtn.textContent = opts.okLabel ?? '확인';
    actions.append(okBtn);
    let cancelBtn: HTMLButtonElement | null = null;
    if (kind !== 'alert') {
      cancelBtn = document.createElement('button');
      cancelBtn.type = 'button';
      cancelBtn.className = 'btn-mini';
      cancelBtn.textContent = opts.cancelLabel ?? '취소';
      actions.append(cancelBtn);
    }
    card.append(actions);
    backdrop.append(card);

    const finish = (ok: boolean) => {
      document.removeEventListener('keydown', onKey, true);
      backdrop.remove();
      previousFocus?.focus?.({ preventScroll: true });
      resolve({ ok, value: input?.value ?? '' });
    };
    const onKey = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        finish(kind === 'alert');
      } else if (event.key === 'Enter' && !event.isComposing) {
        event.preventDefault();
        event.stopPropagation();
        finish(true);
      } else if (event.key === 'Tab') {
        // 창 안에서만 돈다
        const focusables = [input, okBtn, cancelBtn].filter((e): e is HTMLInputElement | HTMLButtonElement => e !== null);
        const index = focusables.indexOf(document.activeElement as HTMLInputElement | HTMLButtonElement);
        event.preventDefault();
        focusables[(index + (event.shiftKey ? -1 : 1) + focusables.length) % focusables.length].focus();
      }
    };
    okBtn.addEventListener('click', () => finish(true));
    cancelBtn?.addEventListener('click', () => finish(false));
    backdrop.addEventListener('mousedown', (event) => {
      if (event.target === backdrop) finish(kind === 'alert');
    });
    document.addEventListener('keydown', onKey, true);
    document.body.append(backdrop);
    (input ?? okBtn).focus();
    input?.select();
  });
}

/** 알림 — 확인을 누르면 끝난다 */
export async function officeAlert(message: string, opts: DialogOptions = {}): Promise<void> {
  await open(message, 'alert', opts);
}

/** 확인/취소 — 확인이면 true */
export async function officeConfirm(message: string, opts: DialogOptions = {}): Promise<boolean> {
  return (await open(message, 'confirm', opts)).ok;
}

/** 글자 입력 — 취소면 null */
export async function officePrompt(message: string, defaultValue = '', opts: DialogOptions & { maxLength?: number } = {}): Promise<string | null> {
  const r = await open(message, 'prompt', { ...opts, defaultValue });
  return r.ok ? r.value : null;
}
