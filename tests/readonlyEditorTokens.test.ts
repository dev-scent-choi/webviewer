import { mkdirSync, mkdtempSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';

describe('읽기 전용 편집기 보기 꾸밈 — 토큰 engine:editor', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let newDocumentId: (typeof import('../src/storage.js'))['newDocumentId'];
  let verifyViewToken: (typeof import('../src/security.js'))['verifyViewToken'];
  const dataDir = mkdtempSync(path.join(os.tmpdir(), 'readonly-editor-data-'));

  beforeAll(async () => {
    mkdirSync(path.join(dataDir, 'audit'));
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_DATA_DIR', dataDir);
    vi.stubEnv('DOCVIEW_PERSIST_LOGS', '1');
    ({ store, newDocumentId } = await import('../src/storage.js'));
    ({ verifyViewToken } = await import('../src/security.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  async function addDoc(ext: string) {
    const id = newDocumentId();
    await store.saveOriginal(id, ext, Buffer.from(`원본-${ext}`));
    await store.writeMeta({ id, name: `시험.${ext}`, ext, status: 'ready', pageCount: 3, createdAt: new Date().toISOString() });
    return id;
  }
  const issue = (id: string, body: Record<string, unknown>) =>
    app.inject({ method: 'POST', url: `/api/documents/${id}/tokens`, payload: { viewer: '시험자', ...body } });
  const data = async (id: string, body: Record<string, unknown>) => (await issue(id, body)).json().data;

  test.each([
    ['docx', '/editor.html', '/view.html'],
    ['xlsx', '/editor-grid.html', '/viewer.html'],
    ['pptx', '/editor-slides.html', '/view.html'],
  ])('%s + 내려받기 + editor → %s?…&ui=view, fallbackUrl %s, canEdit 강제 false', async (ext, page, fallback) => {
    const id = await addDoc(ext);
    const d = await data(id, { canDownload: true, canEdit: true, engine: 'editor' });
    const t = encodeURIComponent(d.token);
    expect(d.engine).toBe('editor');
    expect(d.engineReason).toBeUndefined();
    expect(d.viewerUrl).toBe(`${page}?token=${t}&ui=view`);
    expect(d.fallbackUrl).toBe(`${fallback}?token=${t}`);
    const claims = verifyViewToken(d.token);
    expect(claims.perms.canEdit).toBe(false);
    expect(claims.engine).toBe('editor');
    const info = (await app.inject({ method: 'GET', url: `/api/view/info?token=${t}` })).json().data;
    expect(info.engine).toBe('editor');
    expect(info.artcraft).toBeUndefined();
    // 엔진 전용 원본 경로는 artcraft 토큰만 — editor 토큰은 거부
    expect((await app.inject({ method: 'GET', url: `/api/view/engine-source?token=${t}` })).statusCode).toBe(403);
  });

  test('옛 형식 편집 업로드(.doc 이름, 저장본 docx)도 editor', async () => {
    const id = newDocumentId();
    await store.saveOriginal(id, 'docx', Buffer.from('원본'));
    await store.writeMeta({ id, name: '시험.doc', ext: 'docx', status: 'ready', pageCount: 3, createdAt: new Date().toISOString() });
    const d = await data(id, { canDownload: true, engine: 'editor' });
    expect(d.engine).toBe('editor');
    expect(d.viewerUrl).toBe(`/editor.html?token=${encodeURIComponent(d.token)}&ui=view`);
  });

  test('내려받기 권한이 없으면 libreoffice + 이유, 보기 화면 주소, fallbackUrl 없음, canEdit false', async () => {
    const id = await addDoc('docx');
    const d = await data(id, { canEdit: true, engine: 'editor' });
    expect(d.engine).toBe('libreoffice');
    expect(d.engineReason).toBe('내려받기 권한이 없어 서버 변환으로 엽니다');
    expect(d.viewerUrl).toBe(`/view.html?token=${encodeURIComponent(d.token)}`);
    expect(d.fallbackUrl).toBeUndefined();
    expect(verifyViewToken(d.token).perms.canEdit).toBe(false);
  });

  test.each(['pdf', 'hwp'])('%s에 editor → libreoffice + 편집기 없음 이유, 지금 보기 화면, fallbackUrl 없음, canEdit 그대로', async (ext) => {
    const id = await addDoc(ext);
    const d = await data(id, { canDownload: true, canEdit: true, engine: 'editor' });
    expect(verifyViewToken(d.token).perms.canEdit).toBe(true);
    expect(d.engine).toBe('libreoffice');
    expect(d.engineReason).toBe('이 형식은 읽기 전용 편집기 화면이 없습니다');
    expect(d.viewerUrl).toMatch(/^\/view(er)?\.html\?token=[^&]+$/);
    expect(d.fallbackUrl).toBeUndefined();
  });

  test('pdf에 editor → 편집기 주소 없음, pdfEditor를 보내면 PDF 편집기 조건 그대로', async () => {
    const id = await addDoc('pdf');
    const d = await data(id, { canDownload: true, canEdit: true, engine: 'editor' });
    expect(d.viewerUrl).toBe(`/view.html?token=${encodeURIComponent(d.token)}`);
    expect(d.editorUrl).toBeUndefined();
    const e = await data(id, { canDownload: true, canEdit: true, pdfEditor: true, engine: 'editor' });
    expect(e.editorUrl).toBe(`/editor-pdf.html?token=${encodeURIComponent(e.token)}`);
    expect(verifyViewToken(e.token).perms).toMatchObject({ canEdit: true, pdfEditor: true });
  });

  test('engine 생략 → 지금과 같다(응답 필드 구성 그대로, editorUrl에 ui 없음, 편집 토큰은 편집 그대로)', async () => {
    const id = await addDoc('docx');
    const d = await data(id, { canDownload: true, canEdit: true });
    const t = encodeURIComponent(d.token);
    expect(Object.keys(d)).toEqual(['token', 'viewerUrl', 'editorUrl', 'expiresInSeconds', 'engine']);
    expect(d.engine).toBe('libreoffice');
    expect(d.viewerUrl).toBe(`/view.html?token=${t}`);
    expect(d.editorUrl).toBe(`/editor.html?token=${t}`);
    expect(verifyViewToken(d.token).perms.canEdit).toBe(true);
  });

  test.each([
    ['docx', '/view.html', '/editor.html'],
    ['xlsx', '/viewer.html', '/editor-grid.html'],
    ['pptx', '/view.html', '/editor-slides.html'],
    ['pdf', '/view.html', '/editor-pdf.html'],
  ])('향기랩 방식(canEdit:true, engine 없음) %s → 주소 그대로', async (ext, viewPage, editPage) => {
    const id = await addDoc(ext);
    const d = await data(id, { canDownload: true, canEdit: true, canPrint: true, canCopy: true, pdfEditor: true });
    const t = encodeURIComponent(d.token);
    expect(d.viewerUrl).toBe(`${viewPage}?token=${t}`);
    expect(d.editorUrl).toBe(`${editPage}?token=${t}`);
    expect(d.engine).toBe('libreoffice');
    expect(d.engineReason).toBeUndefined();
    expect(d.fallbackUrl).toBeUndefined();
  });

  test('모르는 engine 값은 400', async () => {
    const id = await addDoc('docx');
    expect((await issue(id, { engine: 'foo' })).statusCode).toBe(400);
  });
});
