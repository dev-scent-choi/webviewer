/**
 * PDF 편집기 변환 창(7단계) — 지금 편집 상태의 PDF를 워드·파워포인트·엑셀·텍스트(서버, /api/view/pdf-convert)나
 * 그림(PNG·JPG — 이 브라우저가 pdf.js로 한 쪽씩 그린다, 여러 쪽이면 zip)으로 내려받고, 다른 파일(워드·엑셀·PPT·한글·그림)을
 * PDF로 바꿔 내려받는다(/api/view/file-to-pdf).
 */
import { zipSync } from 'fflate';
import { openPdfData } from './pdfDocument.js';

const $ = <T extends HTMLElement>(id: string) => document.getElementById(id) as T;

/** 넣기·바꾸기에 받는 파일(서버 변환기가 아는 형식) */
export const CONVERTIBLE_ACCEPT =
  'application/pdf,.pdf,.docx,.doc,.odt,.rtf,.txt,.xlsx,.xls,.ods,.csv,.pptx,.ppt,.odp,.hwp,.hwpx,.png,.jpg,.jpeg,.gif,.bmp,.webp,.tif,.tiff';

const TARGETS: Array<{ value: string; label: string; hint: string }> = [
  { value: 'docx-text', label: '워드(.docx) — 글만', hint: '문단으로 흘러 고쳐 쓰기 좋게. 그림·표 모양은 빠집니다.' },
  { value: 'docx', label: '워드(.docx) — 모양 그대로', hint: '쪽마다 글상자를 제자리에. 보기는 같지만 고치기는 불편합니다.' },
  { value: 'xlsx', label: '엑셀(.xlsx) — 표로', hint: '글자 위치로 행·열을 맞춥니다. 숫자는 숫자로 들어갑니다.' },
  { value: 'pptx', label: '파워포인트(.pptx)', hint: '쪽마다 슬라이드 한 장.' },
  { value: 'png', label: '그림(.png)', hint: '쪽마다 그림 한 장(여러 쪽이면 zip).' },
  { value: 'jpg', label: '그림(.jpg)', hint: '파일이 작습니다(여러 쪽이면 zip).' },
  { value: 'txt', label: '텍스트(.txt)', hint: '글자만, 줄 단위.' },
];

function saveBlob(blob: Blob, fileName: string): void {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = fileName;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

function fileNameOf(res: Response, fallback: string): string {
  const m = /filename\*=UTF-8''([^;]+)/i.exec(res.headers.get('content-disposition') ?? '');
  return m ? decodeURIComponent(m[1]) : fallback;
}

async function errorOf(res: Response, fallback: string): Promise<string> {
  const json = (await res.json().catch(() => ({}))) as { error?: string };
  return json.error ?? fallback;
}

export interface ConvertUi {
  open(): void;
}

export function setupConvert(opts: { token: string; docName: string; notice: (text: string) => void }): ConvertUi {
  const q = encodeURIComponent(opts.token);
  const dialog = $('convert-dialog');
  const list = $('convert-targets');
  const status = $('convert-status');
  const run = $<HTMLButtonElement>('convert-run');
  const dpiRow = $('convert-dpi-row');
  const stem = opts.docName.replace(/\.pdf$/i, '') || 'document';
  for (const t of TARGETS) {
    const row = document.createElement('label');
    row.className = 'pdf-convert-row';
    const radio = document.createElement('input');
    radio.type = 'radio';
    radio.name = 'convert-target';
    radio.value = t.value;
    const text = document.createElement('span');
    const strong = document.createElement('strong');
    strong.textContent = t.label;
    const em = document.createElement('em');
    em.textContent = t.hint;
    text.append(strong, em);
    row.append(radio, text);
    list.append(row);
    radio.addEventListener('change', () => (dpiRow.hidden = !(t.value === 'png' || t.value === 'jpg')));
  }
  const chosen = () => list.querySelector<HTMLInputElement>('input:checked')?.value ?? '';
  for (const b of dialog.querySelectorAll<HTMLButtonElement>('[data-close]')) b.addEventListener('click', () => (dialog.hidden = true));
  dialog.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') dialog.hidden = true;
  });

  let busy = false;
  const setBusy = (on: boolean, text = '') => {
    busy = on;
    run.disabled = on;
    $<HTMLButtonElement>('convert-file-btn').disabled = on;
    status.textContent = text;
  };

  async function toImages(kind: 'png' | 'jpg'): Promise<void> {
    const dpi = Number($<HTMLSelectElement>('convert-dpi').value) || 150;
    const res = await fetch(`/api/view/pdf-current?token=${q}`);
    if (!res.ok) throw new Error(await errorOf(res, 'PDF를 받지 못했습니다.'));
    const pdf = await openPdfData(new Uint8Array(await res.arrayBuffer()));
    const files: Record<string, Uint8Array> = {};
    const digits = String(pdf.numPages).length;
    let single: Blob | null = null;
    try {
      for (let p = 1; p <= pdf.numPages; p += 1) {
        status.textContent = `그리는 중… ${p}/${pdf.numPages}쪽`;
        const page = await pdf.getPage(p);
        const vp = page.getViewport({ scale: dpi / 72 });
        const canvas = document.createElement('canvas');
        canvas.width = Math.floor(vp.width);
        canvas.height = Math.floor(vp.height);
        const ctx = canvas.getContext('2d', { alpha: false })!;
        ctx.fillStyle = '#fff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        await page.render({ canvasContext: ctx, viewport: vp }).promise;
        const blob = await new Promise<Blob>((resolve, reject) =>
          canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('그림을 만들지 못했습니다.'))), kind === 'png' ? 'image/png' : 'image/jpeg', 0.9),
        );
        page.cleanup();
        canvas.width = 0;
        if (pdf.numPages === 1) single = blob;
        else files[`${stem}-${String(p).padStart(digits, '0')}.${kind}`] = new Uint8Array(await blob.arrayBuffer());
      }
    } finally {
      await pdf.destroy();
    }
    if (single) saveBlob(single, `${stem}.${kind}`);
    else saveBlob(new Blob([zipSync(files, { level: 0 })], { type: 'application/zip' }), `${stem}-그림.zip`);
  }

  run.addEventListener('click', async () => {
    const to = chosen();
    if (!to || busy) {
      if (!to) status.textContent = '바꿀 형식을 고르세요.';
      return;
    }
    setBusy(true, to === 'png' || to === 'jpg' ? '준비하는 중…' : '바꾸는 중… (쪽이 많으면 조금 걸립니다)');
    try {
      if (to === 'png' || to === 'jpg') await toImages(to);
      else {
        const res = await fetch(`/api/view/pdf-convert?token=${q}`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ to }) });
        if (!res.ok) throw new Error(await errorOf(res, '바꾸지 못했습니다.'));
        saveBlob(await res.blob(), fileNameOf(res, `${stem}.${to}`));
      }
      setBusy(false, '내려받았습니다.');
      opts.notice('변환한 파일을 내려받았습니다.');
    } catch (error) {
      setBusy(false, error instanceof Error ? error.message : String(error));
    }
  });

  // 다른 파일 → PDF
  const fileInput = $<HTMLInputElement>('convert-file');
  fileInput.accept = CONVERTIBLE_ACCEPT;
  $('convert-file-btn').addEventListener('click', () => {
    fileInput.value = '';
    fileInput.click();
  });
  fileInput.addEventListener('change', async () => {
    const file = fileInput.files?.[0];
    if (!file || busy) return;
    setBusy(true, `${file.name}을(를) PDF로 바꾸는 중…`);
    try {
      const res = await fetch(`/api/view/file-to-pdf?token=${q}&name=${encodeURIComponent(file.name)}`, {
        method: 'POST',
        headers: { 'content-type': 'application/octet-stream' },
        body: file,
      });
      if (!res.ok) throw new Error(await errorOf(res, 'PDF로 바꾸지 못했습니다.'));
      saveBlob(await res.blob(), fileNameOf(res, `${file.name.replace(/\.[^.]+$/, '')}.pdf`));
      setBusy(false, 'PDF를 내려받았습니다.');
    } catch (error) {
      setBusy(false, error instanceof Error ? error.message : String(error));
    }
  });

  return {
    open() {
      if (!chosen()) (list.querySelector('input') as HTMLInputElement).checked = true;
      dpiRow.hidden = !['png', 'jpg'].includes(chosen());
      status.textContent = '';
      dialog.hidden = false;
    },
  };
}
