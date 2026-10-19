import { afterEach, describe, expect, test, vi } from 'vitest';

describe('메모리 저장 모드: 오래 남은 문서 정리', () => {
  afterEach(() => {
    vi.unstubAllEnvs();
    vi.resetModules();
  });

  test('2시간 넘은 문서는 지우고, 최근 문서는 남긴다', async () => {
    vi.resetModules();
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    const { store } = await import('../src/storage.js');
    const { expireIdleMemoryDocuments, MEMORY_DOC_MAX_AGE_MS } = await import('../src/memoryExpiry.js');
    const now = Date.now();
    const OLD = '11111111-1111-4111-8111-111111111111';
    const NEW = '22222222-2222-4222-8222-222222222222';
    const meta = (id: string, ageMs: number) => ({ id, name: "x.docx", ext: 'docx', status: 'ready' as const, pageCount: 1, createdAt: new Date(now - ageMs).toISOString() });
    await store.saveOriginal(OLD, 'docx', Buffer.from('x'));
    await store.writeMeta(meta(OLD, MEMORY_DOC_MAX_AGE_MS + 60_000));
    await store.saveOriginal(NEW, 'docx', Buffer.from('y'));
    await store.writeMeta(meta(NEW, 5 * 60_000));
    expect(await expireIdleMemoryDocuments(now)).toBe(1);
    expect(await store.readMeta(OLD)).toBeNull();
    expect(await store.readMeta(NEW)).not.toBeNull();
  });
});
