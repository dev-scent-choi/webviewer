import { mkdtempSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { brotliCompressSync } from 'node:zlib';
import Fastify, { type FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { parseManifest, readArtcraftManifest } from '../src/artcraftManifest.js';
import { registerArtcraftStatic } from '../src/artcraftStatic.js';

const GOOD = {
  word: '/artcraft/artcraft_word-0123456789.wasm',
  deck: '/artcraft/artcraft_deck-abcdef0123.wasm',
  fonts: ['/artcraft/fonts/SCDream4-0123456789.otf', '/artcraft/fonts/SCDream6-0123456789.otf'],
  commits: { wordcraft: 'f8ec3cc0', deckcraft: 'b547f108' },
};
const tmp = (tag: string) => mkdtempSync(path.join(os.tmpdir(), `artcraft-${tag}-`));

describe('manifest.json 읽기', () => {
  test('맞는 manifest는 word·deck·fonts만 돌려준다', () => {
    expect(parseManifest(GOOD)).toEqual({ word: GOOD.word, deck: GOOD.deck, fonts: GOOD.fonts });
  });

  test('모양이 틀리거나 /artcraft/ 밖 주소·..이 있으면 null', () => {
    expect(parseManifest(null)).toBeNull();
    expect(parseManifest({ ...GOOD, word: 42 })).toBeNull();
    expect(parseManifest({ ...GOOD, fonts: [] })).toBeNull();
    expect(parseManifest({ ...GOOD, deck: 'https://evil.example/x.wasm' })).toBeNull();
    expect(parseManifest({ ...GOOD, fonts: ['/artcraft/../secret.otf'] })).toBeNull();
  });

  test('파일이 없으면 null(아트크래프트 옵션이 꺼진 채 서버는 뜬다), 있으면 읽는다, 깨진 JSON은 null', () => {
    expect(readArtcraftManifest(tmp('none'))).toBeNull();
    const ok = tmp('ok');
    writeFileSync(path.join(ok, 'manifest.json'), JSON.stringify(GOOD));
    expect(readArtcraftManifest(ok)?.word).toBe(GOOD.word);
    const broken = tmp('broken');
    writeFileSync(path.join(broken, 'manifest.json'), '{ 깨짐');
    expect(readArtcraftManifest(broken)).toBeNull();
  });
});

describe('/artcraft/* 정적 서빙', () => {
  let app: FastifyInstance;
  const dir = tmp('static');
  const raw = Buffer.from('\0asm 아트크래프트 시험 '.repeat(200));
  const br = brotliCompressSync(raw);
  beforeAll(async () => {
    writeFileSync(path.join(dir, 'artcraft_word-0123456789.wasm'), raw);
    writeFileSync(path.join(dir, 'artcraft_word-0123456789.wasm.br'), br);
    writeFileSync(path.join(dir, 'manifest.json'), JSON.stringify(GOOD));
    app = Fastify();
    await registerArtcraftStatic(app, dir);
    await app.ready();
  });
  afterAll(() => app.close());

  test('br을 받는 브라우저에는 .br(content-encoding: br, application/wasm, immutable)', async () => {
    const res = await app.inject({ method: 'GET', url: '/artcraft/artcraft_word-0123456789.wasm', headers: { 'accept-encoding': 'gzip, deflate, br' } });
    expect(res.statusCode).toBe(200);
    expect(res.headers['content-encoding']).toBe('br');
    expect(res.headers['content-type']).toBe('application/wasm');
    expect(res.headers['cache-control']).toBe('public, max-age=31536000, immutable');
    expect(String(res.headers.vary)).toContain('accept-encoding');
    expect(Buffer.compare(res.rawPayload, br)).toBe(0);
  });

  test('br을 안 받으면 원본 그대로', async () => {
    const res = await app.inject({ method: 'GET', url: '/artcraft/artcraft_word-0123456789.wasm', headers: { 'accept-encoding': 'identity' } });
    expect(res.statusCode).toBe(200);
    expect(res.headers['content-encoding']).toBeUndefined();
    expect(res.headers['cache-control']).toBe('public, max-age=31536000, immutable');
    expect(Buffer.compare(res.rawPayload, raw)).toBe(0);
  });

  test('manifest.json은 내주지 않는다(서버만 읽는다) — //·대소문자·.br 직접 요청도 404(해시 이름 wasm·otf만 허용)', async () => {
    for (const url of ['/artcraft/manifest.json', '/artcraft//manifest.json', '/artcraft/Manifest.json', '/artcraft/MANIFEST.JSON', '/artcraft/artcraft_word-0123456789.wasm.br']) {
      expect((await app.inject({ method: 'GET', url, headers: { 'accept-encoding': 'br' } })).statusCode, url).toBe(404);
    }
  });
});
