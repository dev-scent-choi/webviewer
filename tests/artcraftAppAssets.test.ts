import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { brotliCompressSync } from 'node:zlib';
import Fastify, { type FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import { appKindFor, parseAppManifest, readArtcraftAppManifest } from '../src/artcraftApp.js';
import { APP_ASSET_PATH, registerArtcraftStatic } from '../src/artcraftStatic.js';

const APP = {
  word: '/artcraft-app/wordcraft-0123456789/wordcraft-web.js',
  deck: '/artcraft-app/deckcraft-abcdef0123/deckcraft-web.js',
  grid: '/artcraft-app/gridcraft-0123456789/gridcraft-web.js',
  commits: { wordcraft: 'f8ec3cc0' },
};
const ENGINE = {
  word: '/artcraft/artcraft_word-0123456789.wasm',
  deck: '/artcraft/artcraft_deck-0123456789.wasm',
  fonts: ['/artcraft/fonts/SCDream4-0123456789.otf', '/artcraft/fonts/SCDream6-0123456789.otf'],
};
const tmp = (tag: string) => mkdtempSync(path.join(os.tmpdir(), `artcraft-app-${tag}-`));

describe('앱 manifest 읽기', () => {
  test('맞는 manifest는 word·deck·grid 주소만 돌려준다', () => {
    expect(parseAppManifest(APP)).toEqual({ word: APP.word, deck: APP.deck, grid: APP.grid });
  });

  test('모양이 틀리거나 앱과 폴더가 다르거나 밖 주소·..이면 null', () => {
    expect(parseAppManifest(null)).toBeNull();
    expect(parseAppManifest({ ...APP, grid: undefined })).toBeNull();
    expect(parseAppManifest({ ...APP, word: '/artcraft-app/deckcraft-0123456789/deckcraft-web.js' })).toBeNull();
    expect(parseAppManifest({ ...APP, deck: 'https://evil.example/deckcraft-web.js' })).toBeNull();
    expect(parseAppManifest({ ...APP, word: '/artcraft-app/wordcraft-0123456789/../wordcraft-web.js' })).toBeNull();
  });

  test('파일이 없으면 null(앱 화면 꺼짐), 있으면 읽는다, 깨진 JSON은 null', () => {
    expect(readArtcraftAppManifest(tmp('none'))).toBeNull();
    const ok = tmp('ok');
    writeFileSync(path.join(ok, 'manifest.json'), JSON.stringify(APP));
    expect(readArtcraftAppManifest(ok)?.grid).toBe(APP.grid);
    const broken = tmp('broken');
    writeFileSync(path.join(broken, 'manifest.json'), '{ 깨짐');
    expect(readArtcraftAppManifest(broken)).toBeNull();
  });
});

describe('appKindFor — 앱 화면으로 열 형식', () => {
  const app = parseAppManifest(APP);
  test('manifest가 없으면 어떤 형식도 앱 화면이 아니다(옵션 꺼짐)', () => {
    for (const ext of ['docx', 'doc', 'pptx', 'xlsx']) expect(appKindFor(ext, null, ENGINE)).toBeNull();
  });
  test('docx·doc → word, pptx → deck, xlsx → grid, ppt·xls·hwp·pdf·constructor → null', () => {
    expect(appKindFor('docx', app, ENGINE)).toBe('word');
    expect(appKindFor('doc', app, ENGINE)).toBe('word');
    expect(appKindFor('pptx', app, ENGINE)).toBe('deck');
    expect(appKindFor('xlsx', app, ENGINE)).toBe('grid');
    for (const ext of ['ppt', 'xls', 'hwp', 'pdf', 'constructor']) expect(appKindFor(ext, app, ENGINE)).toBeNull();
  });
  test('엔진 manifest(한글 글꼴)가 없으면 워드·PPT는 끄고 엑셀만', () => {
    expect(appKindFor('docx', app, null)).toBeNull();
    expect(appKindFor('pptx', app, null)).toBeNull();
    expect(appKindFor('xlsx', app, null)).toBe('grid');
  });
});

describe('/artcraft-app/* 정적 서빙', () => {
  let app: FastifyInstance;
  const dir = tmp('static');
  const js = Buffer.from('export default async function init() {} // 아트크래프트 '.repeat(50));
  const wasm = Buffer.from('\0asm 앱 '.repeat(300));
  beforeAll(async () => {
    const d = path.join(dir, 'wordcraft-0123456789');
    mkdirSync(path.join(d, 'snippets', 'eframe-d31a8450ea12dfc1'), { recursive: true });
    writeFileSync(path.join(d, 'wordcraft-web.js'), js);
    writeFileSync(path.join(d, 'wordcraft-web.js.br'), brotliCompressSync(js));
    writeFileSync(path.join(d, 'wordcraft-web_bg.wasm'), wasm);
    writeFileSync(path.join(d, 'wordcraft-web_bg.wasm.br'), brotliCompressSync(wasm));
    writeFileSync(path.join(d, 'snippets', 'eframe-d31a8450ea12dfc1', 'inline0.js'), 'export {};');
    writeFileSync(path.join(dir, 'manifest.json'), JSON.stringify(APP));
    app = Fastify();
    await registerArtcraftStatic(app, dir, '/artcraft-app/', APP_ASSET_PATH);
    await app.ready();
  });
  afterAll(() => app.close());

  test('JS·wasm: br을 받으면 .br(content-encoding: br), immutable, wasm은 application/wasm, JS는 javascript', async () => {
    const w = await app.inject({ method: 'GET', url: '/artcraft-app/wordcraft-0123456789/wordcraft-web_bg.wasm', headers: { 'accept-encoding': 'br' } });
    expect(w.statusCode).toBe(200);
    expect(w.headers['content-encoding']).toBe('br');
    expect(w.headers['content-type']).toBe('application/wasm');
    expect(w.headers['cache-control']).toBe('public, max-age=31536000, immutable');
    const j = await app.inject({ method: 'GET', url: '/artcraft-app/wordcraft-0123456789/wordcraft-web.js', headers: { 'accept-encoding': 'identity' } });
    expect(j.statusCode).toBe(200);
    expect(String(j.headers['content-type'])).toContain('javascript');
    expect(Buffer.compare(j.rawPayload, js)).toBe(0);
  });

  test('접착 JS가 상대 경로로 부르는 snippets도 내준다', async () => {
    expect((await app.inject({ method: 'GET', url: '/artcraft-app/wordcraft-0123456789/snippets/eframe-d31a8450ea12dfc1/inline0.js' })).statusCode).toBe(200);
  });

  test('manifest·.br 직접·다른 이름·.. 은 404', async () => {
    for (const url of [
      '/artcraft-app/manifest.json',
      '/artcraft-app//manifest.json',
      '/artcraft-app/Manifest.json',
      '/artcraft-app/wordcraft-0123456789/wordcraft-web.js.br',
      '/artcraft-app/wordcraft-0123456789/../manifest.json',
      '/artcraft-app/wordcraft-0123456789/other.js',
      '/artcraft-app/wordcraft-XYZ/wordcraft-web.js',
    ]) {
      expect((await app.inject({ method: 'GET', url, headers: { 'accept-encoding': 'br' } })).statusCode, url).toBe(404);
    }
  });
});

describe('buildApp — public 전체 서빙으로 /artcraft-app/ manifest가 새지 않는다', () => {
  let app: FastifyInstance;
  beforeAll(async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_ARTCRAFT_APP_DIR', tmp('buildapp'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  test('대소문자·// 를 바꿔도 404', async () => {
    for (const url of ['/artcraft-app/manifest.json', '/ARTCRAFT-APP/manifest.json', '/Artcraft-App/manifest.json', '/artcraft-app//manifest.json', '/artcraft-app']) {
      expect((await app.inject({ method: 'GET', url })).statusCode, url).toBe(404);
    }
  });
});
