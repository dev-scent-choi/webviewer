import { mkdtempSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';

// 실제 buildApp() — /artcraft/* 등록 순서와 public 전체 서빙의 cache-control(send 기본값에 덮이지 않는지)
describe('buildApp 정적 서빙', () => {
  let app: FastifyInstance;
  beforeAll(async () => {
    const dir = mkdtempSync(path.join(os.tmpdir(), 'artcraft-app-'));
    writeFileSync(path.join(dir, 'manifest.json'), '{"word":"/artcraft/artcraft_word-0123456789.wasm"}');
    writeFileSync(path.join(dir, 'artcraft_word-0123456789.wasm'), '\0asm');
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', dir);
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  test('/artcraft/manifest.json은 public 전체 서빙으로도 새지 않는다(404), 해시 이름 wasm은 immutable', async () => {
    for (const url of ['/artcraft/manifest.json', '/artcraft//manifest.json', '/artcraft/Manifest.json']) {
      expect((await app.inject({ method: 'GET', url })).statusCode, url).toBe(404);
    }
    const wasm = await app.inject({ method: 'GET', url: '/artcraft/artcraft_word-0123456789.wasm' });
    expect(wasm.statusCode).toBe(200);
    expect(wasm.headers['cache-control']).toBe('public, max-age=31536000, immutable');
  });

  test('public: ?v= 해시 주소는 1년 immutable, 맨 주소는 no-cache, 다른 정적 응답도 cache-control이 있다', async () => {
    const hashed = await app.inject({ method: 'GET', url: '/js/main.js?v=abc123' });
    expect(hashed.statusCode).toBe(200);
    expect(hashed.headers['cache-control']).toBe('public, max-age=31536000, immutable');
    const plain = await app.inject({ method: 'GET', url: '/js/main.js' });
    expect(plain.statusCode).toBe(200);
    expect(plain.headers['cache-control']).toBe('no-cache');
    for (const url of ['/', '/index.html', '/css/style.css', '/fonts/s-core-dream/SCDream4.otf']) {
      const res = await app.inject({ method: 'GET', url });
      expect(res.statusCode, url).toBe(200);
      expect(res.headers['cache-control'], url).toBe('no-cache');
    }
    // HTML 은 조건부 요청으로 304 가 나가면 안 된다(옛 ?v= 주소가 immutable 로 남는다)
    for (const url of ['/', '/index.html', '/view.html']) {
      const res = await app.inject({ method: 'GET', url });
      expect(res.headers['last-modified'], url).toBeUndefined();
      expect(res.headers['etag'], url).toBeUndefined();
      const again = await app.inject({ method: 'GET', url, headers: { 'if-modified-since': new Date().toUTCString() } });
      expect(again.statusCode, url).toBe(200);
    }
  });
});
