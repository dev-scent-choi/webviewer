/**
 * PDF 양식 층 — 쪽마다 pdf.js가 알려 주는 양식 위젯(글자·체크·라디오·목록·서명 칸) 자리에 입력란을 놓는다.
 * 값은 방의 pdfForm(칸 이름 → 값)에 쓰고(LOCAL_ORIGIN — 실행 취소 대상), 같은 이름의 칸(여러 쪽에 걸친 칸 포함)과
 * 다른 참여자 화면에 바로 반영된다. 저장은 서버가 pdf-lib로 채운다(src/pdf/pdfForm.ts).
 *
 * 그림(canvas)은 ENABLE_FORMS로 그려 양식 칸 모양이 빠져 있으므로, 칸의 값은 늘 이 층이 보여 준다(읽기 전용이면 입력만 막는다).
 * 서명 칸(Sig)은 "여기에 서명" 단추 — 누르면 서명 놓기(onSignatureField)로 이어진다.
 */
import type * as Y from 'yjs';
import { LOCAL_ORIGIN } from './annotLayer.js';
import type { PageView, PageViewport } from './pageView.js';

type Widget = Record<string, unknown> & {
  fieldName?: string;
  fieldType?: string;
  rect?: number[];
  fieldValue?: unknown;
  readOnly?: boolean;
  hidden?: boolean;
};
export type FormValue = string | boolean | string[];

export interface FormLayerOptions {
  readonly view: PageView;
  readonly map: Y.Map<unknown>;
  readonly ydoc: Y.Doc;
  readonly canEdit: boolean;
  /** 서명 칸을 눌렀을 때 — 그 쪽과 칸의 PDF 사각형 */
  readonly onSignatureField?: (page: number, rect: [number, number, number, number]) => void;
  /** 그 칸 안에 이미 서명(그림)이 있는지 — 있으면 "여기에 서명" 안내를 감춘다 */
  readonly isSigned?: (pageKey: string, rect: readonly number[]) => boolean;
}

/** 필드 창에 보이는 칸 하나 */
export interface FieldInfo {
  readonly name: string;
  readonly label: string;
  readonly kind: 'text' | 'checkbox' | 'radio' | 'choice' | 'list' | 'signature';
  readonly origIndex: number;
  readonly initial: FormValue;
  readonly options?: ReadonlyArray<{ value: string; label: string }>;
  readonly readOnly: boolean;
}

const str = (v: unknown): string => (typeof v === 'string' ? v : '');

function initialOf(w: Widget): FormValue {
  const v = w.fieldValue;
  if (w.checkBox) return typeof v === 'string' && v !== 'Off' && v === w.exportValue;
  if (w.radioButton) return typeof v === 'string' && v !== 'Off' ? v : '';
  if (w.fieldType === 'Ch') {
    const list = Array.isArray(v) ? v.map(String) : typeof v === 'string' && v ? [v] : [];
    return w.multiSelect ? list : (list[0] ?? '');
  }
  return typeof v === 'string' ? v : '';
}

export class FormLayer {
  /** 칸 이름 → 그 이름의 입력란들(화면에 있는 것) */
  private readonly inputs = new Map<string, Set<HTMLElement>>();
  /** 칸 이름 → 파일의 처음 값(맵에 없으면 이것) */
  private readonly initial = new Map<string, FormValue>();
  private readonly changeListeners: Array<() => void> = [];
  private readonly timers = new Map<string, number>();

  constructor(private readonly opts: FormLayerOptions) {
    opts.view.onPageRendered((page, _layer, vp) => void this.renderPage(page, vp));
    opts.map.observe((event) => {
      for (const name of event.keysChanged) this.syncInputs(name);
      for (const cb of this.changeListeners) cb();
    });
  }

  onChange(cb: () => void): void {
    this.changeListeners.push(cb);
  }

  /** 서명이 놓이거나 지워졌을 때 — 화면의 서명 칸 안내를 다시 맞춘다 */
  refreshSignatureFields(): void {
    for (const btn of document.querySelectorAll<HTMLButtonElement>('.pdf-page .pdf-form-sig')) {
      const page = Number(btn.closest<HTMLElement>('.pdf-page')?.dataset.page);
      const rect = (btn.dataset.rect ?? '').split(',').map(Number);
      if (!Number.isFinite(page) || rect.length !== 4) continue;
      btn.classList.toggle('signed', this.opts.isSigned?.(this.opts.view.pageKeyOf(page), rect) ?? false);
    }
  }

  /** 지금 값 — 바꾼 값이 있으면 그것, 없으면 파일의 처음 값 */
  valueOf(name: string): FormValue | undefined {
    const v = this.opts.map.get(name);
    if (typeof v === 'string' || typeof v === 'boolean' || Array.isArray(v)) return v as FormValue;
    return this.initial.get(name);
  }

  setValue(name: string, value: FormValue): void {
    if (!this.opts.canEdit) return;
    const now = this.valueOf(name);
    if (JSON.stringify(now) === JSON.stringify(value)) return;
    this.opts.ydoc.transact(() => this.opts.map.set(name, value), LOCAL_ORIGIN);
  }

  /** 그 칸의 입력란으로 초점(필드 창에서 누를 때) */
  focus(name: string): boolean {
    const el = [...(this.inputs.get(name) ?? [])].find((e) => e.isConnected);
    if (!el) return false;
    el.scrollIntoView({ block: 'center', inline: 'nearest' });
    (el as HTMLInputElement).focus({ preventScroll: true });
    el.classList.add('flash');
    setTimeout(() => el.classList.remove('flash'), 900);
    return true;
  }

  private register(name: string, el: HTMLElement): void {
    let set = this.inputs.get(name);
    if (!set) {
      set = new Set();
      this.inputs.set(name, set);
    }
    for (const old of [...set]) if (!old.isConnected) set.delete(old);
    set.add(el);
  }

  /** 맵이 바뀌면(내 입력·실행 취소·다른 사람) 그 이름의 입력란들을 맞춘다 — 입력 중인 칸은 건드리지 않는다 */
  private syncInputs(name: string): void {
    const value = this.valueOf(name);
    for (const el of this.inputs.get(name) ?? []) {
      if (!el.isConnected || el === document.activeElement) continue;
      this.show(el, value);
    }
  }

  private show(el: HTMLElement, value: FormValue | undefined): void {
    if (el instanceof HTMLInputElement && el.type === 'checkbox') el.checked = value === true;
    else if (el instanceof HTMLInputElement && el.type === 'radio') el.checked = value === el.value && el.value !== '';
    else if (el instanceof HTMLSelectElement && el.multiple) {
      const set = new Set(Array.isArray(value) ? value : []);
      for (const o of el.options) o.selected = set.has(o.value);
    } else if (el instanceof HTMLSelectElement) el.value = typeof value === 'string' ? value : '';
    else if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) el.value = typeof value === 'string' ? value : '';
  }

  private async renderPage(page: number, vp: PageViewport): Promise<void> {
    const pageEl = this.opts.view.pageElement(page);
    const layer = pageEl?.querySelector<HTMLDivElement>('.pdf-form-layer');
    if (!layer) return;
    const key = this.opts.view.pageKeyOf(page);
    const widgets = (await this.opts.view.widgetsOf(page)) as Widget[];
    // 그리는 동안 쪽이 바뀌었으면(정리·다시 그림) 그만
    if (this.opts.view.pageKeyOf(page) !== key || this.opts.view.viewportOf(page) !== vp) return;
    layer.replaceChildren();
    for (const w of widgets) {
      const name = str(w.fieldName);
      // 읽기 전용 칸은 pdf.js가 그림에 그대로 그린다(겹쳐 그리면 글자가 두 번 보인다)
      if (!name || !w.rect || w.hidden || w.pushButton || (w.readOnly && w.fieldType !== 'Sig')) continue;
      if (!this.initial.has(name)) this.initial.set(name, initialOf(w));
      const el = this.makeInput(w, name, page, vp);
      if (!el) continue;
      const [x1, y1, x2, y2] = vp.convertToViewportRectangle(w.rect) as number[];
      el.style.left = `${Math.min(x1, x2)}px`;
      el.style.top = `${Math.min(y1, y2)}px`;
      el.style.width = `${Math.abs(x2 - x1)}px`;
      el.style.height = `${Math.abs(y2 - y1)}px`;
      layer.append(el);
    }
  }

  private makeInput(w: Widget, name: string, page: number, vp: PageViewport): HTMLElement | null {
    const readOnly = !this.opts.canEdit || w.readOnly === true;
    const box = vp.convertToViewportRectangle(w.rect!) as number[];
    const h = Math.abs(box[3] - box[1]);
    const daSize = Number((w.defaultAppearanceData as { fontSize?: number } | undefined)?.fontSize) || 0;
    const fontPx = daSize > 0 ? daSize * vp.scale : Math.max(8, Math.min(h * 0.62, 14 * vp.scale));
    const label = str(w.alternativeText) || name;
    let el: HTMLElement;
    if (w.fieldType === 'Sig') {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'pdf-form-sig';
      btn.textContent = '여기에 서명';
      btn.disabled = readOnly;
      btn.title = label;
      btn.addEventListener('click', () => this.opts.onSignatureField?.(page, w.rect as [number, number, number, number]));
      btn.dataset.rect = (w.rect as number[]).join(',');
      btn.classList.toggle('signed', this.opts.isSigned?.(this.opts.view.pageKeyOf(page), w.rect as number[]) ?? false);
      el = btn;
    } else if (w.checkBox || w.radioButton) {
      const input = document.createElement('input');
      input.type = w.checkBox ? 'checkbox' : 'radio';
      if (w.radioButton) {
        input.name = `pdf-radio-${page}-${name}`;
        input.value = str(w.buttonValue);
      }
      input.addEventListener('change', () => {
        if (w.checkBox) this.setValue(name, input.checked);
        else if (input.checked) this.setValue(name, input.value);
      });
      // 라디오는 다시 누르면 고르기 풀기(PDF 라디오의 NoToggleToOff가 아니면)
      if (w.radioButton) {
        input.addEventListener('click', (event) => {
          if (this.valueOf(name) === input.value && !w.noToggleToOff) {
            event.preventDefault();
            this.setValue(name, '');
          }
        });
      }
      el = input;
    } else if (w.fieldType === 'Ch') {
      const select = document.createElement('select');
      select.multiple = w.multiSelect === true;
      if (!w.combo) select.size = Math.max(2, Math.floor(h / (fontPx * 1.3)));
      if (w.combo && !select.multiple) select.append(new Option('', ''));
      for (const o of (w.options as Array<{ exportValue?: string; displayValue?: string }> | undefined) ?? []) {
        select.append(new Option(o.displayValue ?? o.exportValue ?? '', o.exportValue ?? o.displayValue ?? ''));
      }
      select.addEventListener('change', () => {
        if (select.multiple) this.setValue(name, [...select.selectedOptions].map((o) => o.value));
        else this.setValue(name, select.value);
      });
      el = select;
    } else if (w.fieldType === 'Tx') {
      const input = w.multiLine ? document.createElement('textarea') : document.createElement('input');
      if (input instanceof HTMLInputElement) input.type = w.password ? 'password' : 'text';
      const max = Number(w.maxLen) || 0;
      if (max > 0) input.maxLength = max;
      input.style.textAlign = w.textAlignment === 1 ? 'center' : w.textAlignment === 2 ? 'right' : 'left';
      // 입력하는 동안은 잠깐 모았다가 쓴다(글자마다 기록이 남지 않게) — 칸을 떠나면 바로
      const commit = () => {
        clearTimeout(this.timers.get(name));
        this.setValue(name, input.value);
      };
      input.addEventListener('input', () => {
        clearTimeout(this.timers.get(name));
        this.timers.set(name, window.setTimeout(commit, 400));
      });
      input.addEventListener('blur', () => {
        commit();
        this.syncInputs(name);
      });
      el = input;
    } else {
      return null;
    }
    el.classList.add('pdf-form-field');
    el.dataset.field = name;
    el.title = label;
    el.setAttribute('aria-label', label);
    el.style.fontSize = `${fontPx}px`;
    if (el instanceof HTMLInputElement || el instanceof HTMLSelectElement || el instanceof HTMLTextAreaElement) {
      el.disabled = readOnly;
      if (readOnly) el.classList.add('readonly');
    }
    this.show(el, this.valueOf(name));
    this.register(name, el);
    return el;
  }

  /** 문서 전체의 양식 칸(필드 창용) — pdf.js getFieldObjects */
  static async listFields(doc: { getFieldObjects(): Promise<Record<string, object[]> | null> }): Promise<FieldInfo[]> {
    const objects = ((await doc.getFieldObjects()) ?? {}) as Record<string, Array<Record<string, unknown>>>;
    const out: FieldInfo[] = [];
    for (const [name, list] of Object.entries(objects)) {
      const first = list.find((o) => typeof o.page === 'number' && (o.page as number) >= 0);
      if (!first) continue;
      const type = str(first.type);
      const kind: FieldInfo['kind'] | null =
        type === 'text' ? 'text' : type === 'checkbox' ? 'checkbox' : type === 'radiobutton' ? 'radio' : type === 'combobox' ? 'choice' : type === 'listbox' ? 'list' : type === 'signature' ? 'signature' : null;
      if (!kind) continue;
      const value = first.value;
      const initial: FormValue =
        kind === 'checkbox'
          ? list.some((o) => typeof o.value === 'string' && o.value !== 'Off' && o.value === o.exportValues)
          : kind === 'radio'
            ? (list.find((o) => typeof o.value === 'string' && o.value !== 'Off')?.value as string | undefined) ?? ''
            : kind === 'list' && Array.isArray(value)
              ? value.map(String)
              : typeof value === 'string'
                ? value
                : Array.isArray(value)
                  ? String(value[0] ?? '')
                  : '';
      out.push({ name, label: name, kind, origIndex: first.page as number, initial, readOnly: first.editable === false });
    }
    return out.sort((a, b) => a.origIndex - b.origIndex);
  }
}
