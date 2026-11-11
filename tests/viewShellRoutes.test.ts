import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';
import { makeSamplePdf } from './helpers.js';

type Perms = { canPrint?: boolean; canDownload?: boolean; canCopy?: boolean; canEdit?: boolean; pdfEditor?: boolean };

describe('보기 껍데기 서버 경로(info·status·derived-pdf)', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let newDocumentId: (typeof import('../src/storage.js'))['newDocumentId'];
  let issueViewToken: (typeof import('../src/security.js'))['issueViewToken'];
  let viewPathFor: (typeof import('../src/routes/documents.js'))['viewPathFor'];
  let pdf: Buffer;

  // 저장소는 메모리 — 모듈을 읽기 전에 환경을 바꿔야 해서 모두 동적 import(정적 import하면 디스크 저장소가 data/에 쓴다)
  beforeAll(async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    ({ store, newDocumentId } = await import('../src/storage.js'));
    ({ issueViewToken } = await import('../src/security.js'));
    ({ viewPathFor } = await import('../src/routes/documents.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
    pdf = await makeSamplePdf(2);
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  async function addDoc(ext: string, opts: { status?: 'ready' | 'processing' | 'failed'; viewKind?: 'pages' | 'grid'; derived?: boolean; error?: string } = {}) {
    const id = newDocumentId();
    await store.saveOriginal(id, ext, ext === 'pdf' ? pdf : Buffer.from('원본'));
    if (opts.derived !== false) await store.saveDerived(id, 'pdf', pdf);
    await store.writeMeta({
      id,
      name: `시험.${ext}`,
      ext,
      status: opts.status ?? 'ready',
      pageCount: 2,
      createdAt: new Date().toISOString(),
      ...(opts.viewKind ? { viewKind: opts.viewKind } : {}),
      ...(opts.status === 'processing' ? { editReady: true } : {}),
      ...(opts.error ? { error: opts.error } : {}),
    });
    return id;
  }
  const tok = (docId: string, perms: Perms = {}) =>
    issueViewToken({ docId, viewer: '시험자', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...perms } }, 600);
  const get = (url: string) => app.inject({ method: 'GET', url });

  test('보기 화면 경로: 쪽 문서는 껍데기, 엑셀·도면·한글은 지금 보기 화면', () => {
    expect(viewPathFor('docx', 'pages')).toBe('/view.html');
    expect(viewPathFor('pptx', undefined)).toBe('/view.html');
    expect(viewPathFor('pdf', 'pages')).toBe('/view.html');
    expect(viewPathFor('xlsx', undefined)).toBe('/viewer.html');
    expect(viewPathFor('xlsx', 'grid')).toBe('/viewer.html');
    expect(viewPathFor('dwg', 'svg')).toBe('/viewer.html');
    expect(viewPathFor('hwp', 'pages')).toBe('/viewer.html');
    expect(viewPathFor('hwpx', 'pages')).toBe('/viewer.html');
  });

  test('derived-pdf: 내려받기 권한이 없으면 403이고 PDF를 주지 않는다', async () => {
    const docx = await addDoc('docx', { viewKind: 'pages' });
    const res = await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(docx, { canEdit: true, canCopy: true, canPrint: true }))}`);
    expect(res.statusCode).toBe(403);
    expect(res.json()).toEqual({ success: false, error: '저장(다운로드) 권한이 없습니다.' });
    const pdfDoc = await addDoc('pdf', { viewKind: 'pages' });
    expect((await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(pdfDoc))}`)).statusCode).toBe(403);
  });

  test('derived-pdf: 내려받기 권한이 있으면 변환한 PDF(보기용 — 저장 이름 없음, 캐시 안 함)', async () => {
    const docx = await addDoc('docx', { viewKind: 'pages' });
    const res = await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(docx, { canDownload: true }))}`);
    expect(res.statusCode).toBe(200);
    expect(res.headers['content-type']).toBe('application/pdf');
    expect(res.headers['cache-control']).toBe('no-store');
    expect(res.headers['content-disposition']).toBeUndefined();
    expect(res.rawPayload.subarray(0, 5).toString('latin1')).toBe('%PDF-');
  });

  test('derived-pdf: 원본이 PDF면 원본 바이트, 첫 변환 중이면 202, 격자 문서는 400', async () => {
    const pdfDoc = await addDoc('pdf', { viewKind: 'pages', derived: false });
    const res = await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(pdfDoc, { canDownload: true }))}`);
    expect(res.statusCode).toBe(200);
    expect(Buffer.compare(res.rawPayload, pdf)).toBe(0);
    const busy = await addDoc('pptx', { status: 'processing', derived: false });
    const r202 = await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(busy, { canDownload: true }))}`);
    expect(r202.statusCode).toBe(202);
    expect(r202.json()).toEqual({ success: true, data: { processing: true } });
    const grid = await addDoc('xlsx', { viewKind: 'grid', derived: false });
    expect((await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(grid, { canDownload: true }))}`)).statusCode).toBe(400);
  });

  test('info: 편집기·보기 주소와 첫 변환 여부', async () => {
    const docx = await addDoc('docx', { viewKind: 'pages' });
    const t = tok(docx, { canEdit: true });
    const data = (await get(`/api/view/info?token=${encodeURIComponent(t)}`)).json().data;
    expect(data.editorUrl).toBe(`/editor.html?token=${encodeURIComponent(t)}`);
    expect(data.viewerUrl).toBe(`/view.html?token=${encodeURIComponent(t)}`);
    expect(data.processing).toBe(false);
    // 열람만이면 편집기 주소 없음(토큰 응답 규칙 editorUrlFor 그대로)
    const viewOnly = (await get(`/api/view/info?token=${encodeURIComponent(tok(docx))}`)).json().data;
    expect(viewOnly.editorUrl).toBeUndefined();
    // pdfEditor 없는 PDF(8000 연동)는 편집기 없이 보기만
    const pdfDoc = await addDoc('pdf', { viewKind: 'pages' });
    expect((await get(`/api/view/info?token=${encodeURIComponent(tok(pdfDoc, { canEdit: true, canDownload: true }))}`)).json().data.editorUrl).toBeUndefined();
    const hwp = await addDoc('hwp', { viewKind: 'pages' });
    expect((await get(`/api/view/info?token=${encodeURIComponent(tok(hwp))}`)).json().data.viewerUrl.startsWith('/viewer.html?')).toBe(true);
    const busy = await addDoc('pptx', { status: 'processing', derived: false });
    expect((await get(`/api/view/info?token=${encodeURIComponent(tok(busy))}`)).json().data.processing).toBe(true);
  });

  test('status: 진행률, 실패 문구(열 수 없는 문서도), 잘못된 토큰 401', async () => {
    const busy = await addDoc('pptx', { status: 'processing', derived: false });
    const s1 = (await get(`/api/view/status?token=${encodeURIComponent(tok(busy))}`)).json();
    expect(s1).toEqual({ success: true, data: { status: 'processing', progress: 0, pageCount: 2 } });
    const failed = await addDoc('docx', { status: 'failed', derived: false, error: '문서를 변환하지 못했습니다. 파일이 손상되었습니다.' });
    const s2 = (await get(`/api/view/status?token=${encodeURIComponent(tok(failed))}`)).json();
    expect(s2.data.status).toBe('failed');
    expect(s2.data.error).toBe('문서를 변환하지 못했습니다. 파일이 손상되었습니다.');
    expect((await get('/api/view/status?token=bad')).statusCode).toBe(401);
  });

  test('토큰 응답 viewerUrl: 쪽 문서는 껍데기, 격자·한글은 지금 보기 화면', async () => {
    const issue = async (id: string) =>
      (await app.inject({ method: 'POST', url: `/api/documents/${id}/tokens`, payload: { viewer: '연동' } })).json().data;
    const docx = await addDoc('docx', { viewKind: 'pages' });
    expect((await issue(docx)).viewerUrl.startsWith('/view.html?token=')).toBe(true);
    const grid = await addDoc('xlsx', { viewKind: 'grid', derived: false });
    expect((await issue(grid)).viewerUrl.startsWith('/viewer.html?token=')).toBe(true);
    const hwp = await addDoc('hwp', { viewKind: 'pages' });
    expect((await issue(hwp)).viewerUrl.startsWith('/viewer.html?token=')).toBe(true);
  });
});
