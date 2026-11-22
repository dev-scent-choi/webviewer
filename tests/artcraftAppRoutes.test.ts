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
  const save = (token: string, body: Buffer) =>
    app.inject({ method: 'PUT', url: `/api/view/artcraft-save?token=${encodeURIComponent(token)}`, headers: { 'content-type': 'application/octet-stream' }, payload: body });
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

  test('engine-source: 앱이 있으면 엑셀 원본도 내준다(no-store, engine_open)', async () => {
    const id = await addDoc('xlsx', Buffer.from('엑셀-원본'), 'grid');
    const res = await app.inject({ method: 'GET', url: `/api/view/engine-source?token=${encodeURIComponent(tok(id, { canDownload: true }, 'artcraft'))}` });
    expect(res.statusCode).toBe(200);
    expect(res.rawPayload.toString('utf8')).toBe('엑셀-원본');
    expect(res.headers['cache-control']).toBe('private, no-store');
  });

  test('저장: 토큰 없음·잘못된 토큰 401', async () => {
    expect((await save('bad.token.x', Buffer.from('x'))).statusCode).toBe(401);
  });

  test('저장: artcraft 토큰이 아니거나 편집 권한이 없으면 403(edit_save_denied), 원본 그대로', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const before = await store.readOriginal(id, 'xlsx');
    const bytes = await xlsxBytes('바뀜');
    expect((await save(tok(id, { canDownload: true, canEdit: true }), bytes)).statusCode).toBe(403);
    expect((await save(tok(id, { canDownload: true }, 'artcraft'), bytes)).statusCode).toBe(403);
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), before)).toBe(0);
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_save_denied', detail: 'permission' }));
  });

  test('저장: 형식이 다른 바이트(docx 문서에 xlsx)는 400, 저장 안 되는 형식(pdf)도 400', async () => {
    const docx = await addDoc('docx');
    expect((await save(tok(docx, { canDownload: true, canEdit: true }, 'artcraft'), await xlsxBytes('x'))).statusCode).toBe(400);
    const pdf = await addDoc('pdf');
    expect((await save(tok(pdf, { canDownload: true, canEdit: true }, 'artcraft'), Buffer.from('%PDF'))).statusCode).toBe(400);
  });

  test('저장: 편집 방이 열려 있으면 409 + 안내, 원본 그대로', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const before = await store.readOriginal(id, 'xlsx');
    openRooms.add(id);
    try {
      const res = await save(tok(id, { canDownload: true, canEdit: true }, 'artcraft'), await xlsxBytes('바뀜'));
      expect(res.statusCode).toBe(409);
      expect(res.json().error).toContain('편집기');
    } finally {
      openRooms.delete(id);
    }
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), before)).toBe(0);
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_save_denied', detail: 'room-open' }));
  });

  test('저장: 엑셀 성공 → 원본 바뀜, 보기 격자 다시 만듦(시트 수), edit_saved(artcraft-app)', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const bytes = await xlsxBytes('새 값');
    const res = await save(tok(id, { canDownload: true, canEdit: true }, 'artcraft'), bytes);
    expect(res.statusCode).toBe(200);
    expect(res.json().data.pageCount).toBe(1);
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), bytes)).toBe(0);
    expect((await store.readMeta(id))?.viewKind).toBe('grid');
    expect(await store.hasDerived(id, 'grid')).toBe(true);
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_saved', detail: 'artcraft-app' }));
  });
});
