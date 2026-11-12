// 시험 자료: Apache POI test-data(Apache License 2.0), 커밋 a4e8640da7c3bb906efc6fa11914d1b80f3cc9dc
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';
import { sofficePath } from '../src/converters/libreOffice.js';

const fx = (name: string) => readFileSync(path.resolve('tests', 'fixtures', 'legacy', name));

describe.skipIf(!sofficePath)('향기랩 옛 형식 편집(legacyEdit)', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let issueViewToken!: (typeof import('../src/security.js'))['issueViewToken'];

  // 저장소는 메모리(향기랩과 같다) — 모듈을 읽기 전에 환경을 바꿔야 해서 동적 import
  beforeAll(async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    ({ store } = await import('../src/storage.js'));
    ({ issueViewToken } = await import('../src/security.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  const upload = async (name: string, legacyEdit: boolean) => {
    const res = await app.inject({
      method: 'POST',
      url: `/api/documents/stream?filename=${encodeURIComponent(name)}${legacyEdit ? '&legacyEdit=1' : ''}`,
      headers: { 'content-type': 'application/octet-stream' },
      payload: fx(name),
    });
    return { status: res.statusCode, body: res.json() as { success: boolean; data?: Record<string, any>; error?: string } };
  };
  const tokenFor = (docId: string) =>
    app.inject({ method: 'POST', url: `/api/documents/${docId}/tokens`, payload: { viewer: '시험', canEdit: true, canDownload: true } });

  test('doc: docx로 바꿔 편집기 주소, 이름은 그대로, 올린 바이트는 파생 legacy', async () => {
    const { status, body } = await upload('table-merges.doc', true);
    expect(status).toBe(201);
    const meta = body.data!;
    expect(meta).toMatchObject({ name: 'table-merges.doc', ext: 'docx', editReady: true, legacy: { ext: 'doc', hasMacros: false } });
    expect(meta.legacy.sha256).toMatch(/^[0-9a-f]{64}$/);
    expect((await store.readDerived(meta.id, 'legacy')).equals(fx('table-merges.doc'))).toBe(true);
    expect((await store.readOriginal(meta.id, 'docx')).subarray(0, 2).toString()).toBe('PK');
    const tok = (await tokenFor(meta.id)).json();
    expect(tok.data.editorUrl).toMatch(/^\/editor\.html\?token=/);
  }, 120_000);

  test('xls·ppt: xlsx·pptx로, 매크로 표시', async () => {
    const xls = (await upload('SimpleMacro.xls', true)).body.data!;
    expect(xls).toMatchObject({ name: 'SimpleMacro.xls', ext: 'xlsx', legacy: { ext: 'xls', hasMacros: true } });
    const ppt = (await upload('SimpleMacro.ppt', true)).body.data!;
    expect(ppt).toMatchObject({ name: 'SimpleMacro.ppt', ext: 'pptx', editReady: true, legacy: { ext: 'ppt', hasMacros: true } });
    expect((await tokenFor(ppt.id)).json().data.editorUrl).toMatch(/^\/editor-slides\.html\?token=/);
  }, 120_000);

  test('legacyEdit 없으면 지금 그대로: doc은 편집기 없음, xls는 xlsx로 이름까지', async () => {
    const doc = (await upload('table-merges.doc', false)).body.data!;
    expect(doc).toMatchObject({ name: 'table-merges.doc', ext: 'doc' });
    expect(doc.legacy).toBeUndefined();
    expect(doc.editReady).toBeUndefined();
    // 변환 중(processing)인 doc은 토큰을 아직 못 받으므로 주소 규칙으로 본다 — 편집기 없음(지금 그대로)
    const { editorUrlFor } = await import('../src/routes/documents.js');
    expect(editorUrlFor('doc', { effectiveCanEdit: true, canDownload: true, pdfEditor: false }, 't')).toBeUndefined();
    const xls = (await upload('chinese-provinces.xls', false)).body.data!;
    expect(xls).toMatchObject({ name: 'chinese-provinces.xlsx', ext: 'xlsx' });
    expect(xls.legacy).toBeUndefined();
  }, 120_000);

  test('바꾸지 못하는 옛 파일은 422(열람 화면으로 내려가지 않는다)', async () => {
    const res = await app.inject({
      method: 'POST',
      url: '/api/documents/stream?filename=broken.doc&legacyEdit=1',
      headers: { 'content-type': 'application/octet-stream' },
      payload: Buffer.from('this is not a word file'),
    });
    expect(res.statusCode).toBe(422);
    expect(res.json().error).toContain('편집할 수 있게 바꾸지 못했습니다');
  }, 120_000);

  test('legacyEdit 값은 1만', async () => {
    const res = await app.inject({ method: 'POST', url: '/api/documents/stream?filename=a.doc&legacyEdit=yes', headers: { 'content-type': 'application/octet-stream' }, payload: fx('table-merges.doc') });
    expect(res.statusCode).toBe(400);
  });

  // Task 3이 이 아래에 내려받기 시험을 더한다(issueViewToken을 쓴다)
  void issueViewToken;
});
