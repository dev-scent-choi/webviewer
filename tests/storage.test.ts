import { describe, expect, test } from 'vitest';
import { MemoryStore, newDocumentId } from '../src/storage.js';
import type { DocumentMeta } from '../src/types.js';

function makeMeta(id: string, createdAt: string): DocumentMeta {
  return { id, name: `${id}.pdf`, ext: 'pdf', status: 'ready', pageCount: 1, createdAt };
}

describe('MemoryStore', () => {
  test('stores and reads originals, converted files, and metadata in memory', async () => {
    const storeInstance = new MemoryStore(10 * 1024 * 1024);
    const id = newDocumentId();
    const original = Buffer.from('original-bytes');
    const converted = Buffer.from('%PDF-converted');

    await storeInstance.saveOriginal(id, 'pdf', original);
    await storeInstance.saveConverted(id, converted);
    await storeInstance.writeMeta(makeMeta(id, new Date().toISOString()));

    expect((await storeInstance.readOriginal(id, 'pdf')).equals(original)).toBe(true);
    expect((await storeInstance.readConverted(id)).equals(converted)).toBe(true);
    expect((await storeInstance.readMeta(id))?.status).toBe('ready');
    expect(await storeInstance.listMeta()).toHaveLength(1);
  });

  test('evicts the oldest document when the memory budget is exceeded', async () => {
    const storeInstance = new MemoryStore(100);
    const oldId = newDocumentId();
    const newId = newDocumentId();

    await storeInstance.saveOriginal(oldId, 'pdf', Buffer.alloc(60));
    await storeInstance.writeMeta(makeMeta(oldId, '2026-01-01T00:00:00.000Z'));
    await storeInstance.saveOriginal(newId, 'pdf', Buffer.alloc(60));
    await storeInstance.writeMeta(makeMeta(newId, '2026-06-01T00:00:00.000Z'));

    // 예산(100B) 초과 → 오래된 문서가 제거되고 새 문서는 남는다
    expect(await storeInstance.readMeta(oldId)).toBeNull();
    await expect(storeInstance.readOriginal(oldId, 'pdf')).rejects.toThrow('만료');
    expect((await storeInstance.readOriginal(newId, 'pdf')).length).toBe(60);
  });

  test('throws a clear error for expired documents', async () => {
    const storeInstance = new MemoryStore();
    await expect(storeInstance.readConverted(newDocumentId())).rejects.toThrow('다시 업로드');
  });

  test('deleteDocument removes every trace of a document', async () => {
    const storeInstance = new MemoryStore();
    const id = newDocumentId();
    await storeInstance.saveOriginal(id, 'pdf', Buffer.from('x'));
    await storeInstance.saveConverted(id, Buffer.from('y'));
    await storeInstance.writeMeta(makeMeta(id, new Date().toISOString()));

    await storeInstance.deleteDocument(id);

    expect(await storeInstance.readMeta(id)).toBeNull();
    expect(await storeInstance.listMeta()).toHaveLength(0);
  });
});
