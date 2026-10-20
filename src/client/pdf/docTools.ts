/**
 * PDF 편집기 6단계 화면 — 암호 파일 비밀번호 창, 문서 보호(암호 걸기·풀기·권한), 압축(미리 보기·적용), 문서 속성(고칠 수 있는 정보 +
 * 버전·쪽 크기·글꼴·보안). 비밀번호는 서버로만 보내고 화면·협업 문서에는 남기지 않는다.
 */
import type * as Y from 'yjs';
import { PDF_META_MAP, PDF_PROPS_MAP } from '../../pdf/annotTypes.js';
import { LOCAL_ORIGIN } from './annotLayer.js';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

async function postJson(url: string, body: unknown): Promise<{ ok: boolean; status: number; data?: unknown; error?: string }> {
  try {
    const res = await fetch(url, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });
    const json = (await res.json().catch(() => ({}))) as { success?: boolean; data?: unknown; error?: string };
    return { ok: res.ok && json.success === true, status: res.status, data: json.data, error: json.error };
  } catch {
    return { ok: false, status: 0, error: '서버에 연결하지 못했습니다.' };
  }
}

/** 비밀번호 창 — 맞히면 resolve. 서버가 아직 방을 못 열었으면(409) 잠시 뒤 다시 */
export function askPassword(token: string): Promise<void> {
  const dialog = $('pw-dialog');
  const input = $<HTMLInputElement>('pw-input');
  const error = $('pw-error');
  const ok = $<HTMLButtonElement>('pw-ok');
  dialog.hidden = false;
  input.value = '';
  error.hidden = true;
  input.focus();
  return new Promise((resolve) => {
    const submit = async () => {
      if (!input.value) {
        error.textContent = '비밀번호를 입력해 주세요.';
        error.hidden = false;
        return;
      }
      ok.disabled = true;
      let r = await postJson(`/api/view/pdf-unlock?token=${encodeURIComponent(token)}`, { password: input.value });
      for (let i = 0; r.status === 409 && i < 10; i += 1) {
        await new Promise((x) => setTimeout(x, 400));
        r = await postJson(`/api/view/pdf-unlock?token=${encodeURIComponent(token)}`, { password: input.value });
      }
      ok.disabled = false;
      if (r.ok) {
        dialog.hidden = true;
        input.value = '';
        ok.removeEventListener('click', onClick);
        input.removeEventListener('keydown', onKey);
        resolve();
        return;
      }
      error.textContent = r.error ?? '열지 못했습니다.';
      error.hidden = false;
      input.select();
    };
    const onClick = () => void submit();
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Enter') void submit();
    };
    ok.addEventListener('click', onClick);
    input.addEventListener('keydown', onKey);
  });
}

export interface DocToolsOptions {
  readonly token: string;
  readonly ydoc: Y.Doc;
  readonly canEdit: boolean;
  readonly docName: string;
  readonly pageCount: () => number;
  readonly notice: (text: string) => void;
}

export interface DocTools {
  openProtect(): void;
  openCompress(): void;
  openProps(): void;
}

interface ProtectionSummary {
  enabled: boolean;
  openPassword: boolean;
  permissions: Record<string, boolean>;
  fromFile: boolean;
}

const PERM_LABELS: Array<[string, string]> = [
  ['printing', '인쇄'],
  ['copying', '내용 복사'],
  ['modifying', '내용 수정·쪽 정리'],
  ['annotating', '주석 달기'],
  ['fillingForms', '양식 채우기'],
];
const LEVELS: Array<[string, string]> = [
  ['', '압축 안 함'],
  ['high', '선명하게 — 그림 200dpi'],
  ['medium', '보통 — 그림 150dpi'],
  ['low', '작게 — 그림 100dpi'],
];
const mb = (n: number) => (n >= 1048576 ? `${(n / 1048576).toFixed(1)}MB` : `${Math.max(1, Math.round(n / 1024))}KB`);

export function setupDocTools(opts: DocToolsOptions): DocTools {
  const { token, ydoc } = opts;
  const meta = ydoc.getMap<unknown>(PDF_META_MAP);
  const props = ydoc.getMap<string>(PDF_PROPS_MAP);
  const q = encodeURIComponent(token);
  const close = (id: string) => ($(id).hidden = true);
  for (const id of ['protect-dialog', 'compress-dialog', 'props-dialog']) {
    const d = $(id);
    d.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') d.hidden = true;
    });
    for (const b of d.querySelectorAll<HTMLButtonElement>('[data-close]')) b.addEventListener('click', () => (d.hidden = true));
  }

  // ---------- 문서 보호 ----------
  const permsHost = $('protect-perms');
  for (const [key, label] of PERM_LABELS) {
    const row = document.createElement('label');
    row.className = 'pdf-check-row';
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.dataset.perm = key;
    box.checked = true;
    row.append(box, document.createTextNode(` ${label}`));
    permsHost.append(row);
  }
  const protectStatus = () => {
    const s = meta.get('protection') as ProtectionSummary | undefined;
    const el = $('protect-status');
    if (!s?.enabled) {
      el.textContent = '지금은 암호가 없습니다.';
      $('protect-remove').hidden = true;
      return;
    }
    const denied = PERM_LABELS.filter(([k]) => s.permissions?.[k] === false).map(([, l]) => l);
    el.textContent = `${s.fromFile ? '올린 파일에 암호가 있습니다' : '저장할 때 암호를 겁니다'} — ${s.openPassword ? '열 때 비밀번호 필요' : '열기는 자유'}${denied.length ? `, 막은 것: ${denied.join('·')}` : ''}.`;
    $('protect-remove').hidden = false;
  };
  const openProtect = () => {
    protectStatus();
    $<HTMLInputElement>('protect-user').value = '';
    $<HTMLInputElement>('protect-owner').value = '';
    const s = meta.get('protection') as ProtectionSummary | undefined;
    for (const box of permsHost.querySelectorAll<HTMLInputElement>('input')) box.checked = s?.permissions?.[box.dataset.perm!] !== false;
    $('protect-error').hidden = true;
    $('protect-dialog').hidden = false;
    $<HTMLInputElement>('protect-user').focus();
  };
  const protectError = (text: string) => {
    $('protect-error').textContent = text;
    $('protect-error').hidden = false;
  };
  $('protect-apply').addEventListener('click', async () => {
    const userPassword = $<HTMLInputElement>('protect-user').value;
    const ownerPassword = $<HTMLInputElement>('protect-owner').value;
    if (!userPassword && !ownerPassword) return protectError('열기 비밀번호나 권한 비밀번호 중 하나는 넣어 주세요.');
    const permissions = Object.fromEntries([...permsHost.querySelectorAll<HTMLInputElement>('input')].map((b) => [b.dataset.perm!, b.checked]));
    const r = await postJson(`/api/view/pdf-protect?token=${q}`, { mode: 'set', userPassword, ownerPassword, permissions });
    if (!r.ok) return protectError(r.error ?? '설정하지 못했습니다.');
    close('protect-dialog');
    opts.notice(userPassword ? '저장할 때 열기 비밀번호를 겁니다.' : '저장할 때 권한을 제한합니다.');
  });
  $('protect-remove').addEventListener('click', async () => {
    const r = await postJson(`/api/view/pdf-protect?token=${q}`, { mode: 'remove' });
    if (!r.ok) return protectError(r.error ?? '풀지 못했습니다.');
    close('protect-dialog');
    opts.notice('암호를 풀었습니다 — 저장하면 암호 없는 파일이 됩니다.');
  });

  // ---------- 압축 ----------
  const levelsHost = $('compress-levels');
  for (const [value, label] of LEVELS) {
    const row = document.createElement('label');
    row.className = 'pdf-check-row';
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'compress-level';
    radio.value = value;
    row.append(radio, document.createTextNode(` ${label}`));
    levelsHost.append(row);
  }
  const chosenLevel = () => levelsHost.querySelector<HTMLInputElement>('input:checked')?.value ?? '';
  const openCompress = () => {
    const current = (meta.get('compress') as string | null | undefined) ?? '';
    for (const r of levelsHost.querySelectorAll<HTMLInputElement>('input')) r.checked = r.value === (current || 'medium');
    $('compress-result').textContent = '';
    $('compress-dialog').hidden = false;
  };
  $('compress-preview').addEventListener('click', async () => {
    const level = chosenLevel();
    const out = $('compress-result');
    if (!level) {
      out.textContent = '압축하지 않으면 지금 크기 그대로입니다.';
      return;
    }
    out.textContent = '계산하는 중…';
    const r = await postJson(`/api/view/pdf-compress?token=${q}`, { level, dryRun: true });
    if (!r.ok) {
      out.textContent = r.error ?? '계산하지 못했습니다.';
      return;
    }
    const d = r.data as { before: number; after: number; images: number };
    const pct = d.before ? Math.round((1 - d.after / d.before) * 100) : 0;
    out.textContent = d.images ? `지금 ${mb(d.before)} → 약 ${mb(d.after)} (${pct}% 줄어듦, 그림 ${d.images}개)` : `줄일 그림이 없습니다(지금 ${mb(d.before)}).`;
  });
  $('compress-apply').addEventListener('click', async () => {
    const level = chosenLevel();
    const r = await postJson(`/api/view/pdf-compress?token=${q}`, { level: level || null });
    if (!r.ok) {
      $('compress-result').textContent = r.error ?? '설정하지 못했습니다.';
      return;
    }
    close('compress-dialog');
    opts.notice(level ? '저장할 때 그림을 줄여 압축합니다.' : '압축하지 않습니다.');
  });

  // ---------- 문서 속성 ----------
  const fields: Array<[string, string]> = [
    ['title', '제목'],
    ['author', '작성자'],
    ['subject', '주제'],
    ['keywords', '키워드'],
  ];
  const propsHost = $('props-fields');
  const inputs = new Map<string, HTMLInputElement>();
  for (const [key, label] of fields) {
    const row = document.createElement('label');
    row.className = 'pdf-prop-row';
    const span = document.createElement('span');
    span.textContent = label;
    const input = document.createElement('input');
    input.type = 'text';
    input.maxLength = 500;
    input.dataset.prop = key;
    input.disabled = !opts.canEdit;
    if (key === 'keywords') input.placeholder = '쉼표로 나눠 적기';
    input.addEventListener('change', () => {
      if ((props.get(key) ?? '') !== input.value) ydoc.transact(() => props.set(key, input.value), LOCAL_ORIGIN);
    });
    row.append(span, input);
    propsHost.append(row);
    inputs.set(key, input);
  }
  props.observe(() => {
    for (const [key, input] of inputs) if (document.activeElement !== input) input.value = props.get(key) ?? '';
  });
  const openProps = async () => {
    for (const [key, input] of inputs) input.value = props.get(key) ?? '';
    const infoHost = $('props-info');
    infoHost.textContent = '읽는 중…';
    $('props-dialog').hidden = false;
    try {
      const res = await fetch(`/api/view/pdf-info?token=${q}`);
      const json = (await res.json()) as { success: boolean; data?: Record<string, unknown>; error?: string };
      if (!json.success || !json.data) throw new Error(json.error ?? '');
      const d = json.data as {
        version: string;
        pageCount: number;
        pageSizes: Array<{ w: number; h: number; count: number }>;
        fonts: Array<{ name: string; type: string; embedded: boolean }>;
        creator: string;
        producer: string;
        created: string | null;
        modified: string | null;
      };
      infoHost.replaceChildren();
      const dl = document.createElement('dl');
      dl.className = 'office-dialog-stats';
      const add = (k: string, v: string) => {
        const dt = document.createElement('dt');
        dt.textContent = k;
        const dd = document.createElement('dd');
        dd.textContent = v;
        dl.append(dt, dd);
      };
      const mm = (pt: number) => Math.round((pt / 72) * 25.4);
      const when = (iso: string | null) => (iso ? new Date(iso).toLocaleString('ko-KR') : '—');
      add('파일 이름', opts.docName);
      add('PDF 버전', d.version || '—');
      add('쪽 수', `${d.pageCount}쪽`);
      add('쪽 크기', d.pageSizes.map((s) => `${mm(s.w)}×${mm(s.h)}mm${d.pageSizes.length > 1 ? ` (${s.count}쪽)` : ''}`).join(', '));
      add('만든 프로그램', d.creator || '—');
      add('PDF 생성 도구', d.producer || '—');
      add('만든 날', when(d.created));
      add('고친 날', when(d.modified));
      const s = meta.get('protection') as ProtectionSummary | undefined;
      add('보안', s?.enabled ? `암호 ${s.openPassword ? '(열기 비밀번호)' : '(권한만)'}` : '없음');
      infoHost.append(dl);
      const h = document.createElement('h3');
      h.textContent = `글꼴 ${d.fonts.length}개`;
      infoHost.append(h);
      const ul = document.createElement('ul');
      ul.className = 'pdf-font-list';
      for (const f of d.fonts) {
        const li = document.createElement('li');
        const name = document.createElement('span');
        name.textContent = f.name;
        const tag = document.createElement('em');
        tag.textContent = `${f.type}${f.embedded ? ' · 포함' : ' · 포함 안 됨'}`;
        if (!f.embedded) tag.classList.add('warn');
        li.append(name, tag);
        ul.append(li);
      }
      if (d.fonts.length === 0) ul.innerHTML = '<li><span>글꼴이 없습니다(그림만 있는 문서).</span></li>';
      infoHost.append(ul);
    } catch {
      infoHost.textContent = '파일 정보를 읽지 못했습니다.';
    }
  };

  return { openProtect, openCompress, openProps: () => void openProps() };
}
