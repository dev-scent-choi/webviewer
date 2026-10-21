/**
 * PDF 편집기 서명 탭 — 서명 만들기 창(그리기·그림 올리기·글자로), 내 서명 목록(이 브라우저에만 보관), 날짜·글자 도장,
 * 양식 칸 목록(필드 창). 놓기·옮기기·크기는 주석 층(annotLayer.ts — 종류 'stamp')이 한다.
 *
 * 서명 그림은 PNG data URL — 긴 변 800px 이하로 줄이고, 그린 서명은 잉크가 있는 부분만 잘라 쓴다.
 */
import { isStampImage } from '../../pdf/annotTypes.js';
import type { AnnotLayer } from './annotLayer.js';
import { FormLayer, type FieldInfo, type FormValue } from './formLayer.js';
import type { PageView } from './pageView.js';
import { officePrompt } from '../officeDialog.js';

const STORE_KEY = 'docview-pdf-signatures';
const MAX_SAVED = 6;
const MAX_SIDE = 800;

interface SavedSignature {
  readonly id: string;
  readonly image: string;
  /** 높이/폭 */
  readonly aspect: number;
}

export interface SignUiOptions {
  readonly view: PageView;
  readonly layer: AnnotLayer;
  readonly forms: FormLayer;
  readonly doc: Parameters<typeof FormLayer.listFields>[0];
  readonly canEdit: boolean;
  readonly showPanel: (name: 'fields') => void;
  readonly notice: (text: string) => void;
}

export interface SignUi {
  /** 서명 탭에 들어가고 나갈 때 */
  setActive(active: boolean): void;
  /** 서명 칸("여기에 서명")을 눌렀을 때 */
  signField(page: number, rect: [number, number, number, number]): void;
}

function loadSaved(): SavedSignature[] {
  try {
    const raw = JSON.parse(localStorage.getItem(STORE_KEY) ?? '[]') as SavedSignature[];
    return Array.isArray(raw) ? raw.filter((s) => s && isStampImage(s.image) && s.aspect > 0) : [];
  } catch {
    return [];
  }
}
function storeSaved(list: readonly SavedSignature[]): void {
  try {
    localStorage.setItem(STORE_KEY, JSON.stringify(list));
  } catch {
    // 저장소를 못 쓰면 이번 화면에서만 쓴다
  }
}

/** canvas에서 잉크(불투명) 있는 곳만 잘라 PNG로 — 여백이 없어야 놓을 때 크기가 맞는다 */
function trimmed(canvas: HTMLCanvasElement, pad = 8): { image: string; aspect: number } | null {
  const ctx = canvas.getContext('2d')!;
  const { width, height } = canvas;
  const data = ctx.getImageData(0, 0, width, height).data;
  let minX = width;
  let minY = height;
  let maxX = -1;
  let maxY = -1;
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      if (data[(y * width + x) * 4 + 3] > 16) {
        if (x < minX) minX = x;
        if (x > maxX) maxX = x;
        if (y < minY) minY = y;
        if (y > maxY) maxY = y;
      }
    }
  }
  if (maxX < 0) return null;
  minX = Math.max(0, minX - pad);
  minY = Math.max(0, minY - pad);
  maxX = Math.min(width - 1, maxX + pad);
  maxY = Math.min(height - 1, maxY + pad);
  const w = maxX - minX + 1;
  const h = maxY - minY + 1;
  const scale = Math.min(1, MAX_SIDE / Math.max(w, h));
  const out = document.createElement('canvas');
  out.width = Math.max(1, Math.round(w * scale));
  out.height = Math.max(1, Math.round(h * scale));
  out.getContext('2d')!.drawImage(canvas, minX, minY, w, h, 0, 0, out.width, out.height);
  return { image: out.toDataURL('image/png'), aspect: out.height / out.width };
}

/** 글자 도장 — 빨간 테두리 안의 글 */
function sealImage(text: string): { image: string; aspect: number } {
  const c = document.createElement('canvas');
  const font = '700 64px Pretendard, "Malgun Gothic", sans-serif';
  const ctx = c.getContext('2d')!;
  ctx.font = font;
  const tw = Math.ceil(ctx.measureText(text).width);
  c.width = tw + 72;
  c.height = 120;
  ctx.font = font;
  ctx.strokeStyle = '#c62828';
  ctx.fillStyle = '#c62828';
  ctx.lineWidth = 7;
  ctx.beginPath();
  ctx.roundRect(6, 6, c.width - 12, c.height - 12, 14);
  ctx.stroke();
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(text, c.width / 2, c.height / 2 + 3);
  return { image: c.toDataURL('image/png'), aspect: c.height / c.width };
}

const TYPE_STYLES = [
  { label: '단정하게', font: '600 96px Pretendard, "Malgun Gothic", sans-serif' },
  { label: '명조', font: '600 96px "Nanum Myeongjo", Batang, "Times New Roman", serif' },
  { label: '흘림', font: 'italic 500 96px Gungsuh, "Segoe Script", cursive' },
];

export function setupSignUi(opts: SignUiOptions): SignUi {
  const { layer, view, forms } = opts;
  const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;
  const tools = $('sign-tools');
  const listEl = $('sign-list');
  const dialog = $('sign-dialog');
  const pad = $<HTMLCanvasElement>('sign-pad');
  const errorEl = $('sign-error');
  let saved = loadSaved();
  /** 서명 칸에서 열었으면 그 자리에 바로 놓는다 */
  let targetField: { page: number; rect: [number, number, number, number] } | null = null;

  // ---------- 내 서명 목록 ----------
  function place(sig: { image: string; aspect: number }): void {
    if (targetField) {
      const key = view.pageKeyOf(targetField.page);
      layer.placeStampInRect(key, targetField.rect, sig.image, sig.aspect);
      targetField = null;
      return;
    }
    layer.placeStamp(sig.image, sig.aspect);
    opts.notice('서명을 놓을 곳을 누르세요(끌면 그 폭으로).');
  }
  function renderList(): void {
    listEl.replaceChildren();
    for (const sig of saved) {
      const wrap = document.createElement('span');
      wrap.className = 'pdf-sign-chip';
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'pdf-sign-use';
      btn.title = '이 서명 놓기';
      btn.setAttribute('aria-label', '내 서명 놓기');
      const img = document.createElement('img');
      img.src = sig.image;
      img.alt = '';
      btn.append(img);
      btn.addEventListener('click', () => place(sig));
      const del = document.createElement('button');
      del.type = 'button';
      del.className = 'pdf-sign-del';
      del.title = '목록에서 지우기';
      del.setAttribute('aria-label', '서명 목록에서 지우기');
      del.textContent = '×';
      del.addEventListener('click', () => {
        saved = saved.filter((s) => s.id !== sig.id);
        storeSaved(saved);
        renderList();
      });
      wrap.append(btn, del);
      listEl.append(wrap);
    }
  }

  // ---------- 서명 만들기 창 ----------
  let tab: 'draw' | 'image' | 'type' = 'draw';
  let ink = '#1f1f23';
  let drawn = false;
  let imageResult: { image: string; aspect: number } | null = null;
  let typeStyle = 0;
  const padCtx = pad.getContext('2d')!;
  const clearPad = () => {
    padCtx.clearRect(0, 0, pad.width, pad.height);
    drawn = false;
  };
  const showError = (text: string) => {
    errorEl.textContent = text;
    errorEl.hidden = !text;
  };
  function selectTab(next: typeof tab): void {
    tab = next;
    for (const b of dialog.querySelectorAll<HTMLButtonElement>('[data-sign-tab]')) b.setAttribute('aria-selected', String(b.dataset.signTab === next));
    for (const p of dialog.querySelectorAll<HTMLElement>('[data-sign-pane]')) p.hidden = p.dataset.signPane !== next;
    showError('');
    if (next === 'type') renderTypeStyles();
  }
  for (const b of dialog.querySelectorAll<HTMLButtonElement>('[data-sign-tab]')) b.addEventListener('click', () => selectTab(b.dataset.signTab as typeof tab));
  function openDialog(): void {
    clearPad();
    imageResult = null;
    $<HTMLInputElement>('sign-file').value = '';
    $<HTMLImageElement>('sign-preview').hidden = true;
    selectTab('draw');
    dialog.hidden = false;
  }
  const closeDialog = () => {
    dialog.hidden = true;
    targetField = null;
  };

  // 그리기
  let stroke: { x: number; y: number } | null = null;
  const padPoint = (e: PointerEvent) => {
    const r = pad.getBoundingClientRect();
    return { x: ((e.clientX - r.left) / r.width) * pad.width, y: ((e.clientY - r.top) / r.height) * pad.height };
  };
  pad.addEventListener('pointerdown', (e) => {
    stroke = padPoint(e);
    pad.setPointerCapture(e.pointerId);
    padCtx.fillStyle = ink;
    padCtx.beginPath();
    padCtx.arc(stroke.x, stroke.y, 2.6, 0, Math.PI * 2);
    padCtx.fill();
    drawn = true;
  });
  pad.addEventListener('pointermove', (e) => {
    if (!stroke) return;
    const p = padPoint(e);
    padCtx.strokeStyle = ink;
    padCtx.lineWidth = 5.2;
    padCtx.lineCap = 'round';
    padCtx.lineJoin = 'round';
    padCtx.beginPath();
    padCtx.moveTo(stroke.x, stroke.y);
    padCtx.lineTo(p.x, p.y);
    padCtx.stroke();
    stroke = p;
  });
  const endStroke = () => (stroke = null);
  pad.addEventListener('pointerup', endStroke);
  pad.addEventListener('pointercancel', endStroke);
  $('sign-clear').addEventListener('click', clearPad);
  for (const b of dialog.querySelectorAll<HTMLButtonElement>('[data-ink]')) {
    b.addEventListener('click', () => {
      ink = b.dataset.ink ?? ink;
      for (const o of dialog.querySelectorAll<HTMLButtonElement>('[data-ink]')) o.setAttribute('aria-pressed', String(o === b));
    });
  }

  // 그림 올리기 — 긴 변 800px, 흰 바탕은 투명하게(선택)
  async function loadImage(file: File): Promise<void> {
    showError('');
    if (!/^image\/(png|jpeg)$/.test(file.type)) {
      showError('PNG·JPG 그림만 쓸 수 있습니다.');
      return;
    }
    const url = URL.createObjectURL(file);
    try {
      const img = new Image();
      img.src = url;
      await img.decode();
      const scale = Math.min(1, MAX_SIDE / Math.max(img.naturalWidth, img.naturalHeight));
      const c = document.createElement('canvas');
      c.width = Math.max(1, Math.round(img.naturalWidth * scale));
      c.height = Math.max(1, Math.round(img.naturalHeight * scale));
      const ctx = c.getContext('2d')!;
      ctx.drawImage(img, 0, 0, c.width, c.height);
      if ($<HTMLInputElement>('sign-whiten').checked) {
        const d = ctx.getImageData(0, 0, c.width, c.height);
        for (let i = 0; i < d.data.length; i += 4) {
          const lum = (d.data[i] + d.data[i + 1] + d.data[i + 2]) / 3;
          if (lum > 235) d.data[i + 3] = 0;
          else if (lum > 200) d.data[i + 3] = Math.round(d.data[i + 3] * ((235 - lum) / 35));
        }
        ctx.putImageData(d, 0, 0);
      }
      imageResult = trimmed(c, 2) ?? { image: c.toDataURL('image/png'), aspect: c.height / c.width };
      const preview = $<HTMLImageElement>('sign-preview');
      preview.src = imageResult.image;
      preview.hidden = false;
    } catch {
      showError('그림을 읽지 못했습니다.');
    } finally {
      URL.revokeObjectURL(url);
    }
  }
  $<HTMLInputElement>('sign-file').addEventListener('change', (e) => {
    const file = (e.target as HTMLInputElement).files?.[0];
    if (file) void loadImage(file);
  });
  $<HTMLInputElement>('sign-whiten').addEventListener('change', () => {
    const file = $<HTMLInputElement>('sign-file').files?.[0];
    if (file) void loadImage(file);
  });

  // 글자로
  const typeInput = $<HTMLInputElement>('sign-text');
  function typedImage(text: string, style: number): { image: string; aspect: number } | null {
    const c = document.createElement('canvas');
    const ctx = c.getContext('2d')!;
    ctx.font = TYPE_STYLES[style].font;
    c.width = Math.ceil(ctx.measureText(text).width) + 60;
    c.height = 150;
    ctx.font = TYPE_STYLES[style].font;
    ctx.fillStyle = ink;
    ctx.textBaseline = 'middle';
    ctx.fillText(text, 30, 78);
    return trimmed(c);
  }
  function renderTypeStyles(): void {
    const host = $('sign-styles');
    host.replaceChildren();
    const text = typeInput.value.trim() || '이름';
    TYPE_STYLES.forEach((style, i) => {
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'pdf-sign-style';
      b.setAttribute('aria-pressed', String(i === typeStyle));
      b.title = style.label;
      const img = document.createElement('img');
      img.alt = style.label;
      const out = typedImage(text, i);
      if (out) img.src = out.image;
      b.append(img);
      b.addEventListener('click', () => {
        typeStyle = i;
        renderTypeStyles();
      });
      host.append(b);
    });
  }
  typeInput.addEventListener('input', renderTypeStyles);

  $('sign-cancel').addEventListener('click', closeDialog);
  dialog.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeDialog();
  });
  $('sign-ok').addEventListener('click', () => {
    let result: { image: string; aspect: number } | null = null;
    if (tab === 'draw') {
      result = drawn ? trimmed(pad) : null;
      if (!result) return showError('서명을 그려 주세요.');
    } else if (tab === 'image') {
      result = imageResult;
      if (!result) return showError('그림을 골라 주세요.');
    } else {
      const text = typeInput.value.trim();
      if (!text) return showError('이름을 입력해 주세요.');
      result = typedImage(text, typeStyle);
    }
    if (!result || !isStampImage(result.image)) return showError('서명이 너무 큽니다. 더 작은 그림을 써 주세요.');
    const sig: SavedSignature = { id: `s${Date.now().toString(36)}`, ...result };
    saved = [sig, ...saved].slice(0, MAX_SAVED);
    storeSaved(saved);
    renderList();
    const field = targetField;
    dialog.hidden = true;
    targetField = field;
    place(sig);
  });

  // ---------- 날짜·글자 도장 ----------
  const today = () => {
    const d = new Date();
    const p = (n: number) => String(n).padStart(2, '0');
    return `${d.getFullYear()}. ${p(d.getMonth() + 1)}. ${p(d.getDate())}.`;
  };
  $('sign-date').addEventListener('click', () => {
    layer.placeText(today());
    opts.notice('날짜를 넣을 곳을 누르세요.');
  });
  const sealGroup = tools.querySelector<HTMLElement>('[data-group="seal"]')!;
  const sealMenu = sealGroup.querySelector<HTMLElement>('.pdf-float-menu')!;
  let sealText = '승인';
  const useSeal = (text: string) => {
    sealText = text;
    const seal = sealImage(text);
    layer.placeStamp(seal.image, seal.aspect);
    opts.notice(`‘${text}’ 도장을 놓을 곳을 누르세요.`);
  };
  $('sign-seal').addEventListener('click', () => useSeal(sealText));
  sealGroup.querySelector('.pdf-float-more')!.addEventListener('click', () => (sealMenu.hidden = !sealMenu.hidden));
  document.addEventListener('pointerdown', (event) => {
    if (!(event.target as HTMLElement).closest('[data-group="seal"]')) sealMenu.hidden = true;
  });
  for (const b of sealMenu.querySelectorAll<HTMLButtonElement>('[data-seal]')) {
    b.addEventListener('click', async () => {
      sealMenu.hidden = true;
      let text = b.dataset.seal ?? '';
      if (!text) text = ((await officePrompt('도장에 넣을 글자(10자 이내)', sealText, { title: '글자 도장', maxLength: 10 })) ?? '').trim().slice(0, 10);
      if (text) useSeal(text);
    });
  }
  $('sign-new').addEventListener('click', openDialog);
  $('sign-fields').addEventListener('click', () => opts.showPanel('fields'));

  // ---------- 필드 창 ----------
  const fieldsHost = $('fields');
  let fields: FieldInfo[] = [];
  const KIND_LABEL: Record<FieldInfo['kind'], string> = { text: '글자', checkbox: '체크', radio: '하나 고르기', choice: '목록', list: '목록', signature: '서명 칸' };
  const shown = (f: FieldInfo, v: FormValue | undefined): string => {
    const value = v ?? f.initial;
    if (f.kind === 'signature') return '';
    if (typeof value === 'boolean') return value ? '✓ 체크됨' : '체크 안 됨';
    if (Array.isArray(value)) return value.join(', ') || '—';
    return value || '—';
  };
  function renderFields(): void {
    $('fields-count').textContent = fields.length ? String(fields.length) : '';
    fieldsHost.replaceChildren();
    if (fields.length === 0) {
      fieldsHost.innerHTML = '<p class="pdf-empty">이 문서에는 양식 칸이 없습니다.</p>';
      return;
    }
    for (const f of fields) {
      const page = view.pageOfKey(`o:${f.origIndex}`);
      const row = document.createElement('button');
      row.type = 'button';
      row.className = 'pdf-field-row';
      row.dataset.field = f.name;
      row.disabled = page === null;
      const name = document.createElement('strong');
      name.textContent = f.label;
      const meta = document.createElement('span');
      meta.textContent = `${KIND_LABEL[f.kind]}${page === null ? ' · 지운 쪽' : ` · ${page}쪽`}${f.readOnly ? ' · 읽기 전용' : ''}`;
      const value = document.createElement('em');
      value.textContent = shown(f, forms.valueOf(f.name));
      row.append(name, meta, value);
      row.addEventListener('click', () => {
        if (page === null) return;
        view.goTo(page);
        // 쪽이 그려지면 그 칸으로
        let tries = 0;
        const go = () => {
          if (forms.focus(f.name) || tries++ > 20) return;
          setTimeout(go, 100);
        };
        go();
      });
      fieldsHost.append(row);
    }
  }
  void FormLayer.listFields(opts.doc)
    .then((list) => {
      fields = list;
      renderFields();
    })
    .catch(() => renderFields());
  forms.onChange(renderFields);
  view.onPagesChange(renderFields);

  renderList();
  return {
    setActive(active: boolean) {
      tools.hidden = !active || !opts.canEdit;
      if (!active) sealMenu.hidden = true;
    },
    signField(page, rect) {
      if (!opts.canEdit) return;
      targetField = { page, rect };
      if (saved[0]) place(saved[0]);
      else openDialog();
    },
  };
}
