import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';

/** /api/view/xlsx-cf-ranges — 편집 또는 내려받기 권한 토큰만, 요청 검증(검토 r4 C1) */
describe('편집 화면 조건부 서식 서버 답 경로', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let newDocumentId: (typeof import('../src/storage.js'))['newDocumentId'];
  let issueViewToken: (typeof import('../src/security.js'))['issueViewToken'];

  beforeAll(async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    ({ store, newDocumentId } = await import('../src/storage.js'));
    ({ issueViewToken } = await import('../src/security.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  async function xlsxDoc(): Promise<string> {
    const id = newDocumentId();
    await store.saveOriginal(id, 'xlsx', Buffer.from('x'));
    await store.writeMeta({ id, name: '시험.xlsx', ext: 'xlsx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), viewKind: 'grid' });
    return id;
  }
  const tok = (docId: string, perms: { canEdit?: boolean; canDownload?: boolean }) =>
    issueViewToken({ docId, viewer: '시험자', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...perms } }, 600);
  const post = (token: string, payload: unknown) =>
    app.inject({ method: 'POST', url: `/api/view/xlsx-cf-ranges?token=${encodeURIComponent(token)}`, payload: payload as object });
  const body = { sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 10, c1: 0, crit: ['a'] }] };

  test('편집·내려받기 권한이 모두 없는 토큰은 403', async () => {
    const id = await xlsxDoc();
    expect((await post(tok(id, {}), body)).statusCode).toBe(403);
  });

  test('내려받기만 되는 토큰(읽기 전용 편집기 화면)도 받는다 — 편집 방이 없으면 409', async () => {
    const id = await xlsxDoc();
    expect((await post(tok(id, { canDownload: true }), body)).statusCode).toBe(409);
  });

  test('편집 토큰: 모양이 틀린 요청은 400, 편집 방이 없으면 409', async () => {
    const id = await xlsxDoc();
    const t = tok(id, { canEdit: true, canDownload: true });
    expect((await post(t, { sheet: 0, ranges: Array.from({ length: 65 }, () => ({ r0: 0, c0: 0, r1: 1, c1: 0 })) })).statusCode).toBe(400);
    expect((await post(t, { sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 1, c1: 0, crit: Array.from({ length: 2_001 }, () => 'a') }] })).statusCode).toBe(400);
    expect((await post(t, body)).statusCode).toBe(409);
  });
});
