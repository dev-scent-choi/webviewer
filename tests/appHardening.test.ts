import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';

describe('서버 기본 방어(본문 상한·프록시 IP·로그의 토큰)', () => {
  let app: FastifyInstance;
  beforeAll(async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    app.get('/__ip', async (request) => ({ ip: request.ip }));
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  test('파일을 받지 않는 라우트는 큰 본문을 읽기 전에 413', async () => {
    const res = await app.inject({
      method: 'POST',
      url: '/api/view/presence?token=x',
      headers: { 'content-type': 'application/json' },
      payload: JSON.stringify({ sessionId: 'a'.repeat(5 * 1024 * 1024) }),
    });
    expect(res.statusCode).toBe(413);
  });

  test('파일을 받는 라우트(pdf-part)는 큰 본문도 받아서 토큰을 확인한다', async () => {
    const res = await app.inject({
      method: 'POST',
      url: '/api/view/pdf-part?token=bad',
      headers: { 'content-type': 'application/octet-stream' },
      payload: Buffer.alloc(6 * 1024 * 1024),
    });
    expect(res.statusCode).not.toBe(413);
  });

  test('이 컴퓨터의 프록시가 붙인 X-Forwarded-For는 믿고, 그 밖의 주소에서 온 것은 믿지 않는다', async () => {
    const proxied = await app.inject({ method: 'GET', url: '/__ip', remoteAddress: '127.0.0.1', headers: { 'x-forwarded-for': '203.0.113.7' } });
    expect(proxied.json().ip).toBe('203.0.113.7');
    const direct = await app.inject({ method: 'GET', url: '/__ip', remoteAddress: '198.51.100.2', headers: { 'x-forwarded-for': '127.0.0.1' } });
    expect(direct.json().ip).toBe('198.51.100.2');
  });

  test('로그용 주소에서 토큰(쿼리)을 지운다', async () => {
    const { redactQuery } = await import('../src/app.js');
    expect(redactQuery('/editor-pdf.html?token=eyJabc.def')).toBe('/editor-pdf.html?…');
    expect(redactQuery('/api/documents')).toBe('/api/documents');
  });
});

describe('오류 응답', () => {
  let app: FastifyInstance;
  beforeAll(async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    app.get('/__boom', async () => {
      throw new Error('ENOENT: C:\secret\path\file.pdf');
    });
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  test('500은 내부 메시지를 숨긴다', async () => {
    const res = await app.inject({ method: 'GET', url: '/__boom' });
    expect(res.statusCode).toBe(500);
    expect(res.body).not.toContain('secret');
    expect(res.json()).toEqual({ success: false, error: '서버 오류가 발생했습니다.' });
  });

  test('검증 오류(400)는 예전 모양 그대로', async () => {
    const res = await app.inject({ method: 'POST', url: '/api/view/presence?token=x', headers: { 'content-type': 'application/json' }, payload: '{}' });
    expect(res.statusCode).toBe(400);
    expect(res.json().message).toBeTruthy();
  });

  test('무거운 작업(암호 풀기)은 접속 IP마다 분당 20번 — 넘으면 429(같은 봉투)', async () => {
    const hit = (ip: string) => app.inject({ method: 'POST', url: '/api/view/pdf-unlock?token=bad', remoteAddress: ip, headers: { 'content-type': 'application/json' }, payload: '{"password":"x"}' });
    for (let i = 0; i < 20; i += 1) expect((await hit('198.51.100.9')).statusCode).not.toBe(429);
    const over = await hit('198.51.100.9');
    expect(over.statusCode).toBe(429);
    expect(over.json()).toMatchObject({ success: false });
    expect((await hit('198.51.100.10')).statusCode).not.toBe(429);
  });
});
