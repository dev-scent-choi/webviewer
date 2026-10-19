/**
 * 업로드 직후의 첫 PDF 변환을 기다리는 곳 — 워드(docx) 편집기는 원본 파일을 직접 읽으므로 변환이 끝나기 전에
 * 열 수 있다(DocumentMeta.editReady). 그러면 편집 → 자동저장의 PDF 재변환이 첫 변환보다 먼저 끝날 수 있는데,
 * 늦게 끝난 첫 변환(편집 전 원본으로 만든 PDF)이 새 PDF를 덮어쓰면 안 된다. 자동저장은 이걸 먼저 기다린다.
 */
import type { DocumentMeta } from './types.js';

const pending = new Map<string, Promise<unknown>>();

export function trackInitialConversion(docId: string, conversion: Promise<unknown>): void {
  const settled = conversion.catch(() => undefined).finally(() => {
    if (pending.get(docId) === settled) pending.delete(docId);
  });
  pending.set(docId, settled);
}

export async function waitForInitialConversion(docId: string): Promise<void> {
  await pending.get(docId);
}

/** 열람·편집 토큰을 내주고 접속을 받아도 되는 상태 — 변환이 끝났거나, 변환 중이어도 편집기가 원본만으로 열리는 문서 */
export function isOpenable(meta: DocumentMeta | null | undefined): meta is DocumentMeta {
  return Boolean(meta && (meta.status === 'ready' || (meta.editReady === true && meta.status === 'processing')));
}
