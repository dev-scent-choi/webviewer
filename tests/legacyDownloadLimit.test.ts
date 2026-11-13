// 고친 옛 형식 내려받기가 LibreOffice를 띄우는 길은 무거운 요청 제한(분당 20)에 걸린다 — LibreOffice 없이 변환을 가짜로
import { afterAll, beforeAll, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';

const convert = vi.hoisted(() => vi.fn(async () => Buffer.from('옛 형식')));
vi.mock('../src/converters/libreOffice.js', async (orig) => ({ ...(await orig<object>()), convertOfficeFile: convert }));

let app: FastifyInstance;
beforeAll(async () => {
  vi.stubEnv('DOCVIEW_STORAGE', 'memory');
  const { buildApp } = await import('../src/app.js');
  app = await buildApp();
  await app.ready();
});
afterAll(async () => {
  await app.close();
  vi.unstubAllEnvs();
});

test('내용이 바뀔 때마다 다시 바꾸면 21번째는 429, 같은 내용은 캐시라 세지 않는다', async () => {
  const { store, newDocumentId } = await import('../src/storage.js');
  const { issueViewToken } = await import('../src/security.js');
  const id = newDocumentId();
  await store.saveDerived(id, 'legacy', Buffer.from('올린 원본'));
  await store.writeMeta({ id, name: '보고서.doc', ext: 'docx', status: 'ready', pageCount: 1, createdAt: '', legacy: { ext: 'doc', sha256: '다름', hasMacros: false } });
  const token = issueViewToken({ docId: id, viewer: '시험', perms: { canPrint: false, canDownload: true, canCopy: false, canEdit: true } }, 600);
  const get = () => app.inject({ method: 'GET', url: `/api/view/download?token=${encodeURIComponent(token)}` });
  for (let i = 0; i < 20; i++) {
    await store.saveOriginal(id, 'docx', Buffer.from(`고친 본 ${i}`));
    expect((await get()).statusCode).toBe(200);
    expect((await get()).statusCode).toBe(200); // 같은 내용 — 캐시
  }
  expect(convert).toHaveBeenCalledTimes(20);
  await store.saveOriginal(id, 'docx', Buffer.from('고친 본 21'));
  const res = await get();
  expect(res.statusCode).toBe(429);
  expect(res.json().error).toContain('요청이 너무 많습니다');
  expect(convert).toHaveBeenCalledTimes(20);
});
