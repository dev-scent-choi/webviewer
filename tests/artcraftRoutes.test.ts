import { mkdirSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';

type Perms = { canPrint?: boolean; canDownload?: boolean; canCopy?: boolean; canEdit?: boolean };

describe('아트크래프트 엔진 — 토큰 발급·/api/view/info', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let newDocumentId: (typeof import('../src/storage.js'))['newDocumentId'];
  let issueViewToken: (typeof import('../src/security.js'))['issueViewToken'];
  let verifyViewToken: (typeof import('../src/security.js'))['verifyViewToken'];
  const installed = mkdtempSync(path.join(os.tmpdir(), 'artcraft-on-'));
  const missing = mkdtempSync(path.join(os.tmpdir(), 'artcraft-off-'));
  const dataDir = mkdtempSync(path.join(os.tmpdir(), 'artcraft-data-'));
  const M = {
    word: '/artcraft/artcraft_word-0123456789.wasm',
    deck: '/artcraft/artcraft_deck-0123456789.wasm',
    fonts: ['/artcraft/fonts/SCDream4-0123456789.otf', '/artcraft/fonts/SCDream6-0123456789.otf'],
  };

  beforeAll(async () => {
    writeFileSync(path.join(installed, 'manifest.json'), JSON.stringify(M));
    mkdirSync(path.join(dataDir, 'audit'));
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_DATA_DIR', dataDir);
    vi.stubEnv('DOCVIEW_PERSIST_LOGS', '1');
    vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', installed);
    ({ store, newDocumentId } = await import('../src/storage.js'));
    ({ issueViewToken, verifyViewToken } = await import('../src/security.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  async function addDoc(ext: string, status: 'ready' | 'processing' = 'ready') {
    const id = newDocumentId();
    await store.saveOriginal(id, ext, Buffer.from(`원본-${ext}`));
    await store.writeMeta({
      id,
      name: `시험.${ext}`,
      ext,
      status,
      pageCount: status === 'ready' ? 3 : 0,
      createdAt: new Date().toISOString(),
      ...(status === 'processing' ? { editReady: true } : {}),
    });
    return id;
  }
  const issue = (id: string, body: Record<string, unknown>) =>
    app.inject({ method: 'POST', url: `/api/documents/${id}/tokens`, payload: { viewer: '시험자', ...body } });
  const tok = (docId: string, perms: Perms, engine?: 'artcraft') =>
    issueViewToken({ docId, viewer: '시험자', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...perms }, ...(engine ? { engine } : {}) }, 600);
  const info = async (token: string) => (await app.inject({ method: 'GET', url: `/api/view/info?token=${encodeURIComponent(token)}` })).json().data;

  test('클레임: engine은 넣은 토큰에만 있고 그대로 돌아온다', () => {
    expect(verifyViewToken(tok('d', {}, 'artcraft')).engine).toBe('artcraft');
    expect('engine' in verifyViewToken(tok('d', {}))).toBe(false);
  });

  test('토큰: docx + 내려받기 + artcraft → engine=artcraft, 클레임에 engine, 주소는 지금 그대로', async () => {
    const id = await addDoc('docx');
    const res = await issue(id, { canDownload: true, engine: 'artcraft' });
    expect(res.statusCode).toBe(200);
    const data = res.json().data;
    expect(data.engine).toBe('artcraft');
    expect(data.engineReason).toBeUndefined();
    expect(data.viewerUrl).toMatch(/^\/view\.html\?token=/);
    expect(verifyViewToken(data.token).engine).toBe('artcraft');
  });

  test('토큰: 내려받기 권한이 없으면 libreoffice + 이유(거절하지 않음)', async () => {
    const id = await addDoc('pptx');
    const data = (await issue(id, { engine: 'artcraft' })).json().data;
    expect(data.engine).toBe('libreoffice');
    expect(data.engineReason).toBe('내려받기 권한이 없어 서버 변환으로 엽니다');
    expect(data.viewerUrl).toMatch(/^\/view\.html\?token=/);
  });

  test('토큰: 안 되는 형식(pdf)은 libreoffice + 이유', async () => {
    const id = await addDoc('pdf');
    const data = (await issue(id, { canDownload: true, engine: 'artcraft' })).json().data;
    expect(data).toMatchObject({ engine: 'libreoffice', engineReason: '이 형식은 브라우저 엔진이 지원하지 않습니다' });
  });

  test('토큰: engine을 안 보내면 예전과 같다(engine=libreoffice, 이유·클레임 없음)', async () => {
    const id = await addDoc('docx');
    const data = (await issue(id, { canDownload: true })).json().data;
    expect(data.engine).toBe('libreoffice');
    expect(data.engineReason).toBeUndefined();
    expect('engine' in verifyViewToken(data.token)).toBe(false);
  });

  test('토큰: 모르는 engine 값은 400', async () => {
    const id = await addDoc('docx');
    expect((await issue(id, { engine: 'chrome' })).statusCode).toBe(400);
  });

  test('토큰·info: manifest가 없으면 libreoffice + "설치되지 않았습니다"', async () => {
    const id = await addDoc('docx');
    vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', missing);
    try {
      const data = (await issue(id, { canDownload: true, engine: 'artcraft' })).json().data;
      expect(data).toMatchObject({ engine: 'libreoffice', engineReason: '브라우저 엔진이 설치되지 않았습니다' });
      const i = await info(data.token);
      expect(i.engine).toBe('libreoffice');
      expect(i.artcraft).toBeUndefined();
    } finally {
      vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', installed);
    }
  });

  test('info: artcraft 토큰이면 engine·artcraft 블록(원본 주소·해시 자산)', async () => {
    const id = await addDoc('docx');
    const t = tok(id, { canDownload: true }, 'artcraft');
    const i = await info(t);
    expect(i.engine).toBe('artcraft');
    expect(i.engineReason).toBeUndefined();
    expect(i.artcraft).toEqual({
      kind: 'word',
      ext: 'docx',
      experimental: false,
      originalUrl: `/api/view/engine-source?token=${encodeURIComponent(t)}`,
      assets: { wasm: M.word, fonts: M.fonts },
    });
  });

  test('info: pptx → deck, doc → word 실험', async () => {
    expect((await info(tok(await addDoc('pptx'), { canDownload: true }, 'artcraft'))).artcraft).toMatchObject({ kind: 'deck', experimental: false });
    expect((await info(tok(await addDoc('doc'), { canDownload: true }, 'artcraft'))).artcraft).toMatchObject({ kind: 'word', ext: 'doc', experimental: true });
  });

  test('info: 클레임에 engine이 있어도 내려받기 권한이 없으면 libreoffice + 이유, artcraft 블록 없음', async () => {
    const i = await info(tok(await addDoc('docx'), {}, 'artcraft'));
    expect(i.engine).toBe('libreoffice');
    expect(i.engineReason).toBe('내려받기 권한이 없어 서버 변환으로 엽니다');
    expect(i.artcraft).toBeUndefined();
  });

  test('info: 기본 토큰 → engine=libreoffice, 이유·artcraft 없음', async () => {
    const i = await info(tok(await addDoc('docx'), { canDownload: true }));
    expect(i.engine).toBe('libreoffice');
    expect(i.engineReason).toBeUndefined();
    expect(i.artcraft).toBeUndefined();
  });

  test('첫 변환 중(editReady) docx도 artcraft로 열리고 원본을 바로 받는다(서버 변환을 기다리지 않음)', async () => {
    const id = await addDoc('docx', 'processing');
    const i = await info(tok(id, { canDownload: true }, 'artcraft'));
    expect(i.engine).toBe('artcraft');
    expect(i.processing).toBe(true);
    const res = await app.inject({ method: 'GET', url: i.artcraft.originalUrl });
    expect(res.statusCode).toBe(200);
    expect(res.rawPayload.toString('utf8')).toBe('원본-docx');
  });

  const source = (token: string) => app.inject({ method: 'GET', url: `/api/view/engine-source?token=${encodeURIComponent(token)}` });
  const auditTypes = (docId: string) =>
    readdirSync(path.join(dataDir, 'audit'))
      .flatMap((f) => readFileSync(path.join(dataDir, 'audit', f), 'utf8').trim().split('\n'))
      .map((l) => JSON.parse(l) as { type: string; docId: string })
      .filter((e) => e.docId === docId)
      .map((e) => e.type);

  test('engine-source: 원본을 no-store로 내주고 감사는 engine_open(내려받기 아님)', async () => {
    const id = await addDoc('pptx');
    const res = await source(tok(id, { canDownload: true }, 'artcraft'));
    expect(res.statusCode).toBe(200);
    expect(res.rawPayload.toString('utf8')).toBe('원본-pptx');
    expect(res.headers['cache-control']).toBe('private, no-store');
    expect(res.headers['content-disposition']).toBeUndefined();
    const types = auditTypes(id);
    expect(types).toContain('engine_open');
    expect(types).not.toContain('downloaded');
  });

  test('engine-source: 내려받기 권한 없음·artcraft 요청 아님·안 되는 형식은 거절', async () => {
    const docx = await addDoc('docx');
    expect((await source(tok(docx, {}, 'artcraft'))).statusCode).toBe(403);
    expect((await source(tok(docx, { canDownload: true }))).statusCode).toBe(403);
    expect((await source(tok(await addDoc('xlsx'), { canDownload: true }, 'artcraft'))).statusCode).toBe(400);
    expect((await source('bad.token.x')).statusCode).toBe(401);
  });

  test('engine-source: 옛 형식(legacy .doc)은 고치지 않았으면 올린 바이트, 고쳤으면 409(LibreOffice를 띄우지 않음)', async () => {
    const { createHash } = await import('node:crypto');
    const id = newDocumentId();
    const ooxml = Buffer.from('바꾼-docx');
    await store.saveOriginal(id, 'docx', ooxml);
    await store.saveDerived(id, 'legacy', Buffer.from('올린-doc'));
    const meta = { id, name: '옛.doc', ext: 'docx', status: 'ready' as const, pageCount: 1, createdAt: new Date().toISOString(), legacy: { ext: 'doc' as const, sha256: createHash('sha256').update(ooxml).digest('hex'), hasMacros: false } };
    await store.writeMeta(meta);
    const t = tok(id, { canDownload: true }, 'artcraft');
    expect((await info(t)).artcraft).toMatchObject({ kind: 'word', ext: 'doc', experimental: true });
    const res = await source(t);
    expect(res.statusCode).toBe(200);
    expect(res.rawPayload.toString('utf8')).toBe('올린-doc');
    await store.saveOriginal(id, 'docx', Buffer.from('고친-docx'));
    expect((await source(t)).statusCode).toBe(409);
  });
});
