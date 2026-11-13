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

  const uploadBytes = async (name: string, payload: Buffer) => {
    const res = await app.inject({
      method: 'POST',
      url: `/api/documents/stream?filename=${encodeURIComponent(name)}&legacyEdit=1`,
      headers: { 'content-type': 'application/octet-stream' },
      payload,
    });
    return { status: res.statusCode, body: res.json() as { success: boolean; data?: Record<string, any>; error?: string } };
  };

  test('이름만 .doc인 OOXML은 실제 종류의 편집기로, 바꾸지 않고, 원래 이름·바이트는 legacy', async () => {
    const { makeSamplePptx } = await import('./helpers.js');
    const pptx = makeSamplePptx(['이름만 doc']);
    const { status, body } = await uploadBytes('deck.doc', pptx);
    expect(status).toBe(201);
    expect(body.data).toMatchObject({ name: 'deck.doc', ext: 'pptx', editReady: true, legacy: { ext: 'doc', hasMacros: false } });
    expect((await store.readOriginal(body.data!.id, 'pptx')).equals(pptx)).toBe(true);
    expect((await tokenFor(body.data!.id)).json().data.editorUrl).toMatch(/^\/editor-slides\.html\?token=/);
    const res = await download(body.data!.id);
    expect(res.rawPayload.equals(pptx)).toBe(true);
    expect(fileName(res.headers['content-disposition'])).toBe('deck.doc');
  }, 120_000);

  test('OLE가 아닌 .doc(RTF·글자 파일)도 LibreOffice가 바꾸면 편집기로', async () => {
    for (const [name, bytes] of [
      ['memo.doc', Buffer.from('{\\rtf1\\ansi RTF로 저장한 문서}')],
      ['plain.doc', Buffer.from('this is not a word file')],
    ] as const) {
      const { status, body } = await uploadBytes(name, bytes);
      expect(status, name).toBe(201);
      expect(body.data).toMatchObject({ name, ext: 'docx', legacy: { ext: 'doc', hasMacros: false } });
    }
  }, 120_000);

  test('정말 바꾸지 못하는 옛 파일만 422(열람 화면으로 내려가지 않는다)', async () => {
    // 복합 파일 머리만 있고 속은 깨진 .ppt — LibreOffice가 아무것도 만들지 못한다
    const broken = Buffer.alloc(4096, 0x5a);
    broken.writeUInt32LE(0xe011cfd0, 0);
    broken.writeUInt32LE(0xe11ab1a1, 4);
    const { status, body } = await uploadBytes('broken.ppt', broken);
    expect(status).toBe(422);
    expect(body.error).toContain('편집할 수 있게 바꾸지 못했습니다');
  }, 120_000);

  test('legacyEdit 값은 1만', async () => {
    const res = await app.inject({ method: 'POST', url: '/api/documents/stream?filename=a.doc&legacyEdit=yes', headers: { 'content-type': 'application/octet-stream' }, payload: fx('table-merges.doc') });
    expect(res.statusCode).toBe(400);
  });

  const download = (docId: string) => {
    const t = issueViewToken({ docId, viewer: '시험', perms: { canPrint: false, canDownload: true, canCopy: false, canEdit: true } }, 600);
    return app.inject({ method: 'GET', url: `/api/view/download?token=${encodeURIComponent(t)}` });
  };
  const fileName = (cd: string | string[] | undefined) => decodeURIComponent(String(cd).split("''")[1] ?? '');

  test('안 고친 내려받기는 올린 바이트 그대로·원래 이름', async () => {
    for (const name of ['table-merges.doc', 'chinese-provinces.xls', 'SampleShow.ppt']) {
      const meta = (await upload(name, true)).body.data!;
      const res = await download(meta.id);
      expect(res.statusCode).toBe(200);
      expect(res.rawPayload.equals(fx(name))).toBe(true);
      expect(fileName(res.headers['content-disposition'])).toBe(name);
    }
  }, 180_000);

  test('고친(저장된 OOXML이 바뀐) 내려받기는 원래 형식으로 다시 바꾼다', async () => {
    const { makeSampleDocx } = await import('./helpers.js');
    const { convertOfficeFile } = await import('../src/converters/libreOffice.js');
    const meta = (await upload('table-merges.doc', true)).body.data!;
    await store.saveOriginal(meta.id, 'docx', makeSampleDocx(['고친글자LEGACY']));
    const res = await download(meta.id);
    expect(res.statusCode).toBe(200);
    expect(res.rawPayload.readUInt32LE(0)).toBe(0xe011cfd0);
    expect(fileName(res.headers['content-disposition'])).toBe('table-merges.doc');
    const back = await convertOfficeFile(res.rawPayload, 'doc', 'docx');
    const { unzipSync, strFromU8 } = await import('fflate');
    // LibreOffice가 글자를 여러 run으로 나누므로 태그를 떼고 본다
    expect(strFromU8(unzipSync(back, { filter: (f) => f.name === 'word/document.xml' })['word/document.xml']!).replace(/<[^>]+>/g, '')).toContain('고친글자LEGACY');
  }, 180_000);
});
