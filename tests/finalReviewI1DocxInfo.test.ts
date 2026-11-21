import fs from 'node:fs';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';

/**
 * FINAL-REVIEW I1 — /api/view/info가 docx마다(읽기 전용 보기 포함) 원본 전체를 풀고 파싱하던 문제.
 * 편집기 화면 정보(docxRender)는 편집기 요청(editor=1)에만, word/ 아래 XML·관계만 풀고, 문서 판(바이트)마다 캐시한다.
 */
const unzipCalls: Array<{ label: string; filter?: (f: { name: string; size: number; originalSize: number }) => boolean }> = [];
vi.mock('../src/zipGuard.js', async (importOriginal) => {
  const real = await importOriginal<typeof import('../src/zipGuard.js')>();
  return {
    ...real,
    guardedUnzipSync: (bytes: Uint8Array, label: string, opts?: { filter?: (f: { name: string; size: number; originalSize: number }) => boolean }) => {
      unzipCalls.push({ label, filter: opts?.filter });
      return real.guardedUnzipSync(bytes, label, opts as never);
    },
  };
});

describe('I1: /api/view/info의 docxRender', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let issueViewToken: (typeof import('../src/security.js'))['issueViewToken'];
  let docId: string;
  const docx = fs.readFileSync('audit/office/corpus/docx/d04-images.docx');

  beforeAll(async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    const storage = await import('../src/storage.js');
    store = storage.store;
    ({ issueViewToken } = await import('../src/security.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
    docId = storage.newDocumentId();
    await store.saveOriginal(docId, 'docx', docx);
    await store.writeMeta({ id: docId, name: 'x.docx', ext: 'docx', status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), viewKind: 'pages' });
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  const info = async (canEdit: boolean, editor: boolean) => {
    const token = issueViewToken({ docId, viewer: '시험자', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit } }, 600);
    const res = await app.inject({ method: 'GET', url: `/api/view/info?token=${encodeURIComponent(token)}${editor ? '&editor=1' : ''}` });
    expect(res.statusCode).toBe(200);
    return res.json().data as { docxRender: unknown; pageSetup: unknown };
  };
  const renderUnzips = () => unzipCalls.filter((c) => c.label === 'getDocxRenderExtras');

  test('읽기 전용 보기(editor 없음)는 편집기 화면 정보를 만들지 않는다', async () => {
    const data = await info(false, false);
    expect(data.docxRender).toBeNull();
    expect(data.pageSetup).not.toBeNull(); // 예전부터 있던 값은 그대로
    expect(renderUnzips()).toHaveLength(0);
  });

  test('편집기 요청은 word/ XML·관계만 풀어 만들고, 같은 판은 다시 풀지 않는다', async () => {
    const first = await info(true, true);
    expect(first.docxRender).toMatchObject({ sections: expect.any(Array) });
    expect(renderUnzips()).toHaveLength(1);
    const filter = renderUnzips()[0].filter!;
    expect(filter).toBeTypeOf('function');
    const f = (name: string) => filter({ name, size: 0, originalSize: 0 });
    expect(f('word/document.xml') && f('word/_rels/document.xml.rels') && f('word/theme/theme1.xml')).toBe(true);
    expect(f('word/media/image1.png') || f('word/embeddings/x.bin') || f('docProps/thumbnail.jpeg')).toBe(false);
    expect(await info(true, true)).toEqual(first);
    expect(renderUnzips()).toHaveLength(1); // 캐시
  });

  test('저장으로 원본이 바뀌면 다시 만든다', async () => {
    const { makeSampleDocx } = await import('./helpers.js');
    await store.saveOriginal(docId, 'docx', makeSampleDocx(['바뀐 문서']));
    await info(true, true);
    expect(renderUnzips()).toHaveLength(2);
  });
});
