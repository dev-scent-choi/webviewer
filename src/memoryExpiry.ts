/**
 * 메모리 저장 모드(향기랩) 문서의 안전망 — 사이트는 "뒤로"와 페이지를 떠날 때 문서를 지우지만(최선 노력), 그 요청이
 * 실패하면(네트워크·브라우저 강제 종료 등) 메모리 저장소엔 만료가 없어 서버를 다시 켤 때까지 남았다(2026-10-19 실측:
 * 시험 문서 9개가 남아 있었다). "올린 문서를 저장하지 않는다"는 약속에 맞게, 아무도 열어 두지 않은 채 오래된 문서를
 * 주기적으로 지운다.
 */
import { config } from './config.js';
import { discardPrewarmedXlsx, forceCloseRoom, isRoomOpen } from './collab/room.js';
import { invalidateRenderCache } from './rendering.js';
import { store } from './storage.js';

export const MEMORY_DOC_MAX_AGE_MS = 2 * 60 * 60_000;
const CHECK_INTERVAL_MS = 10 * 60_000;

export async function expireIdleMemoryDocuments(now = Date.now()): Promise<number> {
  if (config.storageMode !== 'memory') return 0;
  let removed = 0;
  for (const meta of await store.listMeta()) {
    const age = now - Date.parse(meta.createdAt);
    if (!(age > MEMORY_DOC_MAX_AGE_MS) || isRoomOpen(meta.id)) continue;
    await store.deleteDocument(meta.id);
    discardPrewarmedXlsx(meta.id);
    forceCloseRoom(meta.id);
    invalidateRenderCache(meta.id);
    removed += 1;
  }
  return removed;
}

export function startMemoryExpiry(log: (removed: number) => void): void {
  if (config.storageMode !== 'memory') return;
  setInterval(() => {
    void expireIdleMemoryDocuments().then((n) => n > 0 && log(n)).catch(() => undefined);
  }, CHECK_INTERVAL_MS).unref();
}
