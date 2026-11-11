/**
 * 인쇄·내려받기·다른 이름으로 저장(표시 포함/제외) — 지금 보기 화면(viewer.js setupToolbar)과 같은 서버 경로·진행률 표시.
 * 권한이 없는 단추는 흐리게 하고 이유를 툴팁으로(설계 §2).
 */
import type { ShellCtx } from './ctx.js';
import { disabledReason } from './logic.js';

const $ = <T extends HTMLElement = HTMLElement>(id: string): T => document.getElementById(id) as T;

/** content-length 기준 진행률과 함께 본문을 받는다 */
async function fetchWithProgress(url: string, onProgress: (received: number, total: number) => void): Promise<Uint8Array[]> {
  const res = await fetch(url);
  if (!res.ok || !res.body) {
    const json = (await res.json().catch(() => null)) as { error?: string } | null;
    throw new Error(json?.error || `요청 실패 (HTTP ${res.status})`);
  }
  const total = Number(res.headers.get('content-length')) || 0;
  const reader = res.body.getReader();
  const chunks: Uint8Array[] = [];
  let received = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    received += value.length;
    onProgress(received, total);
  }
  return chunks;
}

function formatProgress(received: number, total: number): string {
  return total > 0 ? `${Math.round((received / total) * 100)}%` : `${(received / 1048576).toFixed(1)}MB`;
}

/** 단추를 진행률 표시로 바꿨다가 끝나면 되돌린다(아이콘 단추는 툴팁에) */
async function runButtonTask(button: HTMLButtonElement, prefix: string, task: (onProgress: (r: number, t: number) => void) => Promise<void>): Promise<void> {
  const original = button.title;
  button.disabled = true;
  try {
    await task((received, total) => {
      button.title = `${prefix} ${formatProgress(received, total)}`;
    });
    button.title = original;
  } catch (error) {
    button.title = '실패';
    setTimeout(() => (button.title = original), 2500);
    throw error;
  } finally {
    button.disabled = false;
  }
}

function saveBlob(chunks: Uint8Array[], name: string): void {
  const url = URL.createObjectURL(new Blob(chunks as BlobPart[]));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

export function setupFileActions(ctx: ShellCtx): void {
  const { permissions, ext, name } = ctx.info;
  const includeMarkup = () => (($('markup-include-checkbox') as HTMLInputElement | null)?.checked ? '&includeMarkup=1' : '');
  const print = $<HTMLButtonElement>('print-btn');
  const download = $<HTMLButtonElement>('download-btn');
  const downloadAs = $<HTMLButtonElement>('download-as-btn');

  const printReason = disabledReason('print', permissions, ext);
  print.disabled = printReason !== null;
  print.title = printReason ?? '워터마크가 포함된 출력용 PDF';
  if (!printReason) {
    print.addEventListener('click', () => {
      // 팝업 차단을 피하려면 누른 그 순간에 창을 먼저 연다
      const printWindow = window.open('', '_blank');
      runButtonTask(print, '생성', async (onProgress) => {
        const chunks = await fetchWithProgress(`/api/view/print?${ctx.tokenQuery}${includeMarkup()}`, onProgress);
        const url = URL.createObjectURL(new Blob(chunks as BlobPart[], { type: 'application/pdf' }));
        if (printWindow && !printWindow.closed) printWindow.location.href = url;
        else window.open(url, '_blank', 'noopener');
        setTimeout(() => URL.revokeObjectURL(url), 60_000);
      }).catch(() => {
        if (printWindow && !printWindow.closed) printWindow.close();
      });
    });
    document.addEventListener('keydown', (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'p') print.click();
    });
  }

  const downloadReason = disabledReason('download', permissions, ext);
  for (const btn of [download, downloadAs]) {
    btn.disabled = downloadReason !== null;
    btn.title = downloadReason ?? (btn === download ? '원본 파일 다운로드' : '다른 이름으로 저장');
  }
  if (downloadReason) return;
  const run = (btn: HTMLButtonElement, fileName: string) =>
    runButtonTask(btn, '저장', async (onProgress) => {
      saveBlob(await fetchWithProgress(`/api/view/download?${ctx.tokenQuery}${includeMarkup()}`, onProgress), fileName);
    }).catch(() => undefined);
  download.addEventListener('click', () => void run(download, name));
  downloadAs.addEventListener('click', () => {
    const custom = window.prompt('저장할 파일 이름을 입력하세요', name);
    if (custom) void run(downloadAs, custom);
  });
}
