/**
 * 인증서 전자서명 창(9단계) — 사용자가 고른 인증서(.pfx·.p12)와 비밀번호를 서버(/api/view/pdf-sign)로 보내
 * 지금 편집 상태를 서명한 PDF를 내려받는다. 인증서·비밀번호는 서버가 서명에만 쓰고 저장하지 않는다(이 창도 닫으면 비운다).
 */
const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

export interface CertSignUi {
  open(): void;
}

const BOX_W = 180;
const BOX_H = 56;
const MARGIN = 36;

/** 쪽 크기(pt)와 자리 이름 → 서명 칸 사각형(PDF 좌표, 왼쪽 아래 원점) */
export function boxRect(pos: string, size: { w: number; h: number }): [number, number, number, number] {
  const right = pos.endsWith('right');
  const top = pos.startsWith('top');
  const x = right ? size.w - MARGIN - BOX_W : MARGIN;
  const y = top ? size.h - MARGIN - BOX_H : MARGIN;
  return [Math.max(0, x), Math.max(0, y), Math.max(0, x) + BOX_W, Math.max(0, y) + BOX_H];
}

function toBase64(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

export function setupCertSign(opts: {
  token: string;
  docName: string;
  currentPage: () => number;
  pageKeyOf: (page: number) => string;
  sizeOf: (page: number) => { w: number; h: number };
  notice: (text: string) => void;
}): CertSignUi {
  const q = encodeURIComponent(opts.token);
  const dialog = $('cert-dialog');
  const file = $<HTMLInputElement>('cert-file');
  const password = $<HTMLInputElement>('cert-password');
  const reason = $<HTMLInputElement>('cert-reason');
  const location = $<HTMLInputElement>('cert-location');
  const position = $<HTMLSelectElement>('cert-position');
  const status = $('cert-status');
  const run = $<HTMLButtonElement>('cert-run');
  const stem = opts.docName.replace(/\.pdf$/i, '') || 'document';

  const close = () => {
    dialog.hidden = true;
    // 비밀번호·인증서는 창을 닫으면 지운다
    password.value = '';
    file.value = '';
    status.textContent = '';
  };
  for (const b of dialog.querySelectorAll<HTMLButtonElement>('[data-close]')) b.addEventListener('click', close);
  dialog.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') close();
  });

  run.addEventListener('click', async () => {
    const picked = file.files?.[0];
    if (!picked) {
      status.textContent = '인증서 파일(.pfx·.p12)을 고르세요.';
      return;
    }
    if (picked.size > 140_000) {
      status.textContent = '인증서 파일이 너무 큽니다.';
      return;
    }
    const page = opts.currentPage();
    const pos = position.value;
    const body: Record<string, unknown> = {
      p12: toBase64(new Uint8Array(await picked.arrayBuffer())),
      password: password.value,
      reason: reason.value.trim(),
      location: location.value.trim(),
    };
    if (pos !== 'none') {
      body.pageKey = opts.pageKeyOf(page);
      body.rect = boxRect(pos, opts.sizeOf(page));
    }
    run.disabled = true;
    status.textContent = '서명하는 중…';
    try {
      const res = await fetch(`/api/view/pdf-sign?token=${q}`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });
      if (!res.ok) {
        const json = (await res.json().catch(() => ({}))) as { error?: string };
        status.textContent = json.error ?? '서명하지 못했습니다.';
        return;
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${stem}-서명.pdf`;
      document.body.append(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30_000);
      close();
      opts.notice('서명한 PDF를 내려받았습니다 — 서명 뒤 파일을 고치면 서명이 깨집니다.');
    } catch {
      status.textContent = '서버에 연결하지 못했습니다.';
    } finally {
      run.disabled = false;
      body.password = '';
    }
  });

  return {
    open() {
      status.textContent = '';
      dialog.hidden = false;
      file.focus();
    },
  };
}
