import { createHash } from 'node:crypto';
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import type { FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';

// 편집 방 열림 흉내 — 저장 경로가 방이 열린 문서를 거절하는지(실제 방을 열지 않고)
const { openRooms } = vi.hoisted(() => ({ openRooms: new Set<string>() }));
vi.mock('../src/collab/room.js', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/collab/room.js')>()),
  isRoomOpen: (docId: string) => openRooms.has(docId),
}));
// 쪽 문서(docx) 저장 뒤 PDF 다시 만들기 — LibreOffice를 띄우지 않고 불렸는지만 본다
const { convertCalls } = vi.hoisted(() => ({ convertCalls: [] as string[] }));
vi.mock('../src/routes/documents.js', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/routes/documents.js')>()),
  convertToPdf: async (meta: { id: string }) => {
    convertCalls.push(meta.id);
    return { pdfBytes: Buffer.from('%PDF'), pageCount: 3 };
  },
}));

type Perms = { canPrint?: boolean; canDownload?: boolean; canCopy?: boolean; canEdit?: boolean };

describe('아트크래프트 앱 화면 — 토큰·info·engine-source·저장', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let newDocumentId: (typeof import('../src/storage.js'))['newDocumentId'];
  let issueViewToken: (typeof import('../src/security.js'))['issueViewToken'];
  let contentUrlFor: (typeof import('../src/routes/documents.js'))['contentUrlFor'];
  const engineDir = mkdtempSync(path.join(os.tmpdir(), 'artcraft-engine-'));
  const appDir = mkdtempSync(path.join(os.tmpdir(), 'artcraft-appdir-'));
  const missing = mkdtempSync(path.join(os.tmpdir(), 'artcraft-missing-'));
  const dataDir = mkdtempSync(path.join(os.tmpdir(), 'artcraft-appdata-'));
  const ENGINE = {
    word: '/artcraft/artcraft_word-0123456789.wasm',
    deck: '/artcraft/artcraft_deck-0123456789.wasm',
    fonts: ['/artcraft/fonts/SCDream4-0123456789.otf', '/artcraft/fonts/SCDream6-0123456789.otf'],
  };
  const APP = {
    word: '/artcraft-app/wordcraft-0123456789/wordcraft-web.js',
    deck: '/artcraft-app/deckcraft-0123456789/deckcraft-web.js',
    grid: '/artcraft-app/gridcraft-0123456789/gridcraft-web.js',
  };

  beforeAll(async () => {
    writeFileSync(path.join(engineDir, 'manifest.json'), JSON.stringify(ENGINE));
    writeFileSync(path.join(appDir, 'manifest.json'), JSON.stringify(APP));
    mkdirSync(path.join(dataDir, 'audit'));
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_DATA_DIR', dataDir);
    vi.stubEnv('DOCVIEW_PERSIST_LOGS', '1');
    vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', engineDir);
    vi.stubEnv('DOCVIEW_ARTCRAFT_APP_DIR', appDir);
    vi.stubEnv('DOCVIEW_MAX_UPLOAD_MB', '1'); // 본문 상한 413 시험용
    // viewer.ts를 먼저 — room.ts가 viewer.ts를 import하는 순환이라, room.js를 먼저 부르면 viewer.ts가 흉내 전 원래 isRoomOpen에 묶인다
    await import('../src/routes/viewer.js');
    ({ store, newDocumentId } = await import('../src/storage.js'));
    ({ issueViewToken } = await import('../src/security.js'));
    ({ contentUrlFor } = await import('../src/routes/documents.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  async function addDoc(ext: string, bytes: Buffer = Buffer.from(`원본-${ext}`), viewKind?: 'pages' | 'grid') {
    const id = newDocumentId();
    await store.saveOriginal(id, ext, bytes);
    await store.writeMeta({ id, name: `시험.${ext}`, ext, status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), ...(viewKind ? { viewKind } : {}) });
    return id;
  }
  async function xlsxBytes(text: string): Promise<Buffer> {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('가').getCell('A1').value = text;
    return Buffer.from(await wb.xlsx.writeBuffer());
  }
  const issue = (id: string, body: Record<string, unknown>) =>
    app.inject({ method: 'POST', url: `/api/documents/${id}/tokens`, payload: { viewer: '시험자', ...body } });
  const tok = (docId: string, perms: Perms, engine?: 'artcraft') =>
    issueViewToken({ docId, viewer: '시험자', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...perms }, ...(engine ? { engine } : {}) }, 600);
  const info = async (token: string) => (await app.inject({ method: 'GET', url: `/api/view/info?token=${encodeURIComponent(token)}` })).json().data;
  const sha = (b: Uint8Array) => createHash('sha256').update(b).digest('hex');
  /** engine-source가 주는 ETag와 같은 값 — 지금 저장된 원본의 sha256 */
  const etag = async (id: string, ext: string) => `"${sha(await store.readOriginal(id, ext))}"`;
  const save = (token: string, body: Buffer, ifMatch?: string) =>
    app.inject({
      method: 'PUT',
      url: `/api/view/artcraft-save?token=${encodeURIComponent(token)}`,
      headers: { 'content-type': 'application/octet-stream', ...(ifMatch ? { 'if-match': ifMatch } : {}) },
      payload: body,
    });
  const editTok = (id: string) => tok(id, { canDownload: true, canEdit: true }, 'artcraft');
  const audits = (docId: string) =>
    readdirSync(path.join(dataDir, 'audit'))
      .flatMap((f) => readFileSync(path.join(dataDir, 'audit', f), 'utf8').trim().split('\n'))
      .map((l) => JSON.parse(l) as { type: string; docId: string; detail?: string })
      .filter((e) => e.docId === docId);

  describe('contentUrlFor — 같은 토큰 + ui=content', () => {
    const ro = { canEdit: false, canDownload: true };
    const none = { canEdit: false, canDownload: false };
    test.each([
      ['docx', 'pages', 'editor', ro, '/editor.html?token=t&ui=content'],
      ['xlsx', 'grid', 'editor', ro, '/editor-grid.html?token=t&ui=content'],
      ['pptx', 'pages', 'editor', ro, '/editor-slides.html?token=t&ui=content'],
      ['docx', 'pages', 'libreoffice', none, '/view.html?token=t&ui=content'],
      ['docx', 'pages', 'artcraft', ro, '/view.html?token=t&ui=content'],
      ['pdf', 'pages', 'libreoffice', none, '/view.html?token=t&ui=content'],
      ['xlsx', 'grid', 'artcraft', ro, '/editor-grid.html?token=t&ui=content'],
      ['xlsx', 'grid', 'libreoffice', ro, '/editor-grid.html?token=t&ui=content'],
      ['xlsx', 'grid', 'libreoffice', none, undefined],
      ['hwp', 'pages', 'libreoffice', ro, undefined],
      ['dwg', 'svg', 'libreoffice', ro, undefined],
    ] as const)('%s · %s · %s → %s', (ext, viewKind, engine, perms, want) => {
      expect(contentUrlFor(ext, viewKind, engine, perms, 't')).toBe(want);
    });
    test('토큰은 주소 인코딩', () => {
      expect(contentUrlFor('pdf', 'pages', 'libreoffice', none, 'a+b')).toBe('/view.html?token=a%2Bb&ui=content');
    });
  });

  test('토큰: docx + 내려받기 + artcraft → viewerUrl은 앱 화면, contentUrl은 엔진 쪽 그림 보기', async () => {
    const data = (await issue(await addDoc('docx'), { canDownload: true, engine: 'artcraft' })).json().data;
    expect(data.engine).toBe('artcraft');
    expect(data.viewerUrl).toBe(`/artcraft-app.html?token=${encodeURIComponent(data.token)}`);
    expect(data.contentUrl).toBe(`/view.html?token=${encodeURIComponent(data.token)}&ui=content`);
  });

  test('토큰: xlsx + 내려받기 + artcraft → 앱 화면(엔진은 엑셀을 못 그려도), contentUrl은 격자 편집기', async () => {
    const data = (await issue(await addDoc('xlsx', Buffer.from('x'), 'grid'), { canDownload: true, engine: 'artcraft' })).json().data;
    expect(data.engine).toBe('artcraft');
    expect(data.viewerUrl).toMatch(/^\/artcraft-app\.html\?token=/);
    expect(data.contentUrl).toMatch(/^\/editor-grid\.html\?token=[^&]+&ui=content$/);
  });

  test('토큰: 앱 manifest가 없으면 지금 그대로(docx → 엔진 보기, xlsx → 지원하지 않음)', async () => {
    vi.stubEnv('DOCVIEW_ARTCRAFT_APP_DIR', missing);
    try {
      const docx = (await issue(await addDoc('docx'), { canDownload: true, engine: 'artcraft' })).json().data;
      expect(docx.viewerUrl).toMatch(/^\/view\.html\?token=/);
      const xlsx = (await issue(await addDoc('xlsx', Buffer.from('x'), 'grid'), { canDownload: true, engine: 'artcraft' })).json().data;
      expect(xlsx).toMatchObject({ engine: 'libreoffice', engineReason: '이 형식은 브라우저 엔진이 지원하지 않습니다' });
    } finally {
      vi.stubEnv('DOCVIEW_ARTCRAFT_APP_DIR', appDir);
    }
  });

  test('토큰: 읽기 전용 편집기 화면(engine editor)의 contentUrl은 편집기 + ui=content', async () => {
    const data = (await issue(await addDoc('pptx'), { canDownload: true, engine: 'editor' })).json().data;
    expect(data.contentUrl).toBe(`/editor-slides.html?token=${encodeURIComponent(data.token)}&ui=content`);
  });

  test('info: artcraft 토큰이면 artcraftApp 블록(워드는 글꼴, 엑셀은 글꼴 없음), viewerUrl은 그대로 보기 화면', async () => {
    const d = await addDoc('docx');
    const t = tok(d, { canDownload: true }, 'artcraft');
    const i = await info(t);
    expect(i.artcraftApp).toEqual({ kind: 'word', js: APP.word, fonts: ENGINE.fonts, originalUrl: `/api/view/engine-source?token=${encodeURIComponent(t)}` });
    expect(i.viewerUrl).toMatch(/^\/view\.html\?token=/);
    const x = await info(tok(await addDoc('xlsx', Buffer.from('x'), 'grid'), { canDownload: true }, 'artcraft'));
    expect(x.artcraftApp).toMatchObject({ kind: 'grid', js: APP.grid, fonts: [] });
  });

  test('info: 내려받기 권한 없음·artcraft 요청 아님·엔진 글꼴 없음(워드)이면 artcraftApp 없음', async () => {
    expect((await info(tok(await addDoc('docx'), {}, 'artcraft'))).artcraftApp).toBeUndefined();
    expect((await info(tok(await addDoc('docx'), { canDownload: true }))).artcraftApp).toBeUndefined();
    vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', missing);
    try {
      expect((await info(tok(await addDoc('docx'), { canDownload: true }, 'artcraft'))).artcraftApp).toBeUndefined();
    } finally {
      vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', engineDir);
    }
  });

  test('engine-source: 앱이 있으면 엑셀 원본도 내준다(no-store, ETag = sha256)', async () => {
    const id = await addDoc('xlsx', Buffer.from('엑셀-원본'), 'grid');
    const res = await app.inject({ method: 'GET', url: `/api/view/engine-source?token=${encodeURIComponent(tok(id, { canDownload: true }, 'artcraft'))}` });
    expect(res.statusCode).toBe(200);
    expect(res.rawPayload.toString('utf8')).toBe('엑셀-원본');
    expect(res.headers['cache-control']).toBe('private, no-store');
    expect(res.headers.etag).toBe(`"${sha(Buffer.from('엑셀-원본'))}"`);
  });

  test('저장: 토큰 없음·잘못된 토큰 401', async () => {
    expect((await save('bad.token.x', Buffer.from('x'))).statusCode).toBe(401);
    const none = await app.inject({ method: 'PUT', url: '/api/view/artcraft-save', headers: { 'content-type': 'application/octet-stream' }, payload: Buffer.from('x') });
    expect(none.statusCode).toBe(401);
  });

  test('저장: artcraft 토큰 아님·편집 권한 없음·내려받기 권한 없음이면 403(edit_save_denied permission), 원본 그대로', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const before = await store.readOriginal(id, 'xlsx');
    const bytes = await xlsxBytes('바뀜');
    const base = await etag(id, 'xlsx');
    expect((await save(tok(id, { canDownload: true, canEdit: true }), bytes, base)).statusCode).toBe(403);
    expect((await save(tok(id, { canDownload: true }, 'artcraft'), bytes, base)).statusCode).toBe(403);
    expect((await save(tok(id, { canEdit: true }, 'artcraft'), bytes, base)).statusCode).toBe(403);
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), before)).toBe(0);
    expect(audits(id).filter((e) => e.type === 'edit_save_denied' && e.detail === 'permission')).toHaveLength(3);
  });

  test('저장: If-Match 없거나 틀린 꼴이면 428(no-base)', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    expect((await save(editTok(id), await xlsxBytes('바뀜'))).statusCode).toBe(428);
    expect((await save(editTok(id), await xlsxBytes('바뀜'), '"abc"')).statusCode).toBe(428);
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_save_denied', detail: 'no-base' }));
  });

  test('저장: 형식이 다른 바이트 400(format), 저장 안 되는 형식 400(unsupported), 빈 본문 400(empty) — 모두 감사', async () => {
    const docx = await addDoc('docx');
    expect((await save(editTok(docx), await xlsxBytes('x'), await etag(docx, 'docx'))).statusCode).toBe(400);
    expect(audits(docx)).toContainEqual(expect.objectContaining({ type: 'edit_save_denied', detail: 'format' }));
    const pdf = await addDoc('pdf');
    expect((await save(editTok(pdf), Buffer.from('%PDF'), await etag(pdf, 'pdf'))).statusCode).toBe(400);
    expect(audits(pdf)).toContainEqual(expect.objectContaining({ type: 'edit_save_denied', detail: 'unsupported' }));
    const xlsx = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    expect((await save(editTok(xlsx), Buffer.alloc(0), await etag(xlsx, 'xlsx'))).statusCode).toBe(400);
    expect(audits(xlsx)).toContainEqual(expect.objectContaining({ type: 'edit_save_denied', detail: 'empty' }));
  });

  test('저장: 본문 상한 초과 413, 풀면 너무 큰 zip(선언 크기 위조) 413 — 원본 그대로', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const before = await store.readOriginal(id, 'xlsx');
    const base = await etag(id, 'xlsx');
    expect((await save(editTok(id), Buffer.alloc(1024 * 1024 + 1, 1), base)).statusCode).toBe(413);
    const bomb = await xlsxBytes('폭탄');
    const view = new DataView(bomb.buffer, bomb.byteOffset, bomb.byteLength);
    for (let i = 0; i + 4 <= bomb.length; i += 1) if (view.getUint32(i, true) === 0x02014b50) view.setUint32(i + 24, 0xfffffff0, true);
    expect((await save(editTok(id), bomb, base)).statusCode).toBe(413);
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), before)).toBe(0);
  });

  test('저장: 편집 방이 열려 있으면 409 + 안내, 원본 그대로', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const before = await store.readOriginal(id, 'xlsx');
    openRooms.add(id);
    try {
      const res = await save(editTok(id), await xlsxBytes('바뀜'), await etag(id, 'xlsx'));
      expect(res.statusCode).toBe(409);
      expect(res.json().error).toContain('편집기');
    } finally {
      openRooms.delete(id);
    }
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), before)).toBe(0);
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_save_denied', detail: 'room-open' }));
  });

  test('저장: 같은 base로 두 번 저장하면 두 번째는 409(stale) — 먼저 저장한 내용 그대로', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const base = await etag(id, 'xlsx');
    const first = await xlsxBytes('먼저');
    expect((await save(editTok(id), first, base)).statusCode).toBe(200);
    const res = await save(editTok(id), await xlsxBytes('나중'), base);
    expect(res.statusCode).toBe(409);
    expect(res.json().error).toContain('먼저 저장');
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), first)).toBe(0);
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_save_denied', detail: 'stale' }));
    // 새 원본의 해시로는 다시 저장된다
    expect((await save(editTok(id), await xlsxBytes('다시'), await etag(id, 'xlsx'))).statusCode).toBe(200);
  });

  test('저장: 엑셀 성공 → 원본 바뀜, 보기 격자 다시 만듦(시트 수), edit_saved(artcraft-app)', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const bytes = await xlsxBytes('새 값');
    const res = await save(editTok(id), bytes, await etag(id, 'xlsx'));
    expect(res.statusCode).toBe(200);
    expect(res.json().data.pageCount).toBe(1);
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), bytes)).toBe(0);
    expect((await store.readMeta(id))?.viewKind).toBe('grid');
    expect(await store.hasDerived(id, 'grid')).toBe(true);
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_saved', detail: 'artcraft-app' }));
  });

  test('저장: 워드 성공 → 원본 바뀜, PDF 다시 만듦(쪽 수·meta pages), edit_saved', async () => {
    const { zipSync, strToU8 } = await import('fflate');
    const docx = (text: string) =>
      Buffer.from(
        zipSync({
          '[Content_Types].xml': strToU8(
            '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
          ),
          'word/document.xml': strToU8(`<w:document><w:body><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:body></w:document>`),
        }),
      );
    const id = await addDoc('docx', docx('원래'), 'pages');
    const bytes = docx('새 글');
    const res = await save(editTok(id), bytes, await etag(id, 'docx'));
    expect(res.statusCode).toBe(200);
    expect(res.json().data.pageCount).toBe(3);
    expect(convertCalls).toContain(id);
    expect(Buffer.compare(await store.readOriginal(id, 'docx'), bytes)).toBe(0);
    expect(await store.readMeta(id)).toMatchObject({ viewKind: 'pages', pageCount: 3, status: 'ready' });
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_saved', detail: 'artcraft-app' }));
  });
});
