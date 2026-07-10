import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { makeSamplePdf, multipartPayload } from './helpers.js';

let app: FastifyInstance;
let dataDir: string;

beforeAll(async () => {
  dataDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-api-'));
  process.env.DOCVIEW_DATA_DIR = dataDir;
  process.env.LOG_LEVEL = 'silent';
  const { ensureDataDirs } = await import('../src/storage.js');
  const { buildApp } = await import('../src/app.js');
  await ensureDataDirs();
  app = await buildApp();
});

afterAll(async () => {
  await app.close();
  await fs.rm(dataDir, { recursive: true, force: true });
});

interface DocMeta {
  readonly id: string;
  readonly status: string;
  readonly pageCount: number;
  readonly error?: string;
}

/** 변환은 백그라운드로 진행되므로 상태가 확정될 때까지 폴링한다. */
async function waitForStatus(id: string, timeoutMs = 30_000): Promise<DocMeta> {
  const startedAt = Date.now();
  for (;;) {
    const res = await app.inject({ method: 'GET', url: `/api/documents/${id}` });
    expect(res.statusCode).toBe(200);
    const meta = res.json().data as DocMeta;
    if (meta.status !== 'processing') return meta;
    if (Date.now() - startedAt > timeoutMs) throw new Error('conversion polling timed out');
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
}

async function uploadSamplePdf(pages = 2): Promise<DocMeta> {
  const { payload, headers } = multipartPayload('계약서.pdf', 'application/pdf', await makeSamplePdf(pages));
  const res = await app.inject({ method: 'POST', url: '/api/documents', payload, headers });
  expect(res.statusCode).toBe(201);
  const meta = await waitForStatus(res.json().data.id);
  expect(meta.status).toBe('ready');
  return meta;
}

async function issueToken(docId: string, canPrint: boolean, canDownload: boolean): Promise<string> {
  const res = await app.inject({
    method: 'POST',
    url: `/api/documents/${docId}/tokens`,
    payload: { viewer: '테스터', canPrint, canDownload },
  });
  expect(res.statusCode).toBe(200);
  return res.json().data.token;
}

describe('document upload', () => {
  test('uploads a PDF and reports page count', async () => {
    const doc = await uploadSamplePdf(3);
    expect(doc.pageCount).toBe(3);
    expect(doc.id).toMatch(/^[0-9a-f-]{36}$/);
  });

  test('rejects unsupported file types with 415', async () => {
    const { payload, headers } = multipartPayload('virus.exe', 'application/octet-stream', Buffer.from('MZ'));
    const res = await app.inject({ method: 'POST', url: '/api/documents', payload, headers });
    expect(res.statusCode).toBe(415);
    expect(res.json().success).toBe(false);
  });

  test('marks corrupt PDF content as failed after background conversion', async () => {
    const { payload, headers } = multipartPayload('broken.pdf', 'application/pdf', Buffer.from('not a pdf'));
    const res = await app.inject({ method: 'POST', url: '/api/documents', payload, headers });
    expect(res.statusCode).toBe(201);
    const meta = await waitForStatus(res.json().data.id);
    expect(meta.status).toBe('failed');
    expect(meta.error).toBeTruthy();
  });

  test('returns 404 for an unknown document id lookup', async () => {
    const res = await app.inject({
      method: 'GET',
      url: '/api/documents/00000000-0000-0000-0000-000000000000',
    });
    expect(res.statusCode).toBe(404);
  });

  test('lists uploaded documents', async () => {
    const res = await app.inject({ method: 'GET', url: '/api/documents' });
    expect(res.statusCode).toBe(200);
    expect(res.json().data.length).toBeGreaterThan(0);
  });

  test('exposes the supported format list', async () => {
    const res = await app.inject({ method: 'GET', url: '/api/formats' });
    expect(res.statusCode).toBe(200);
    expect(res.json().data).toEqual(expect.arrayContaining(['pdf', 'png', 'jpg', 'jpeg']));
  });
});

describe('token issuance', () => {
  test('returns 404 for an unknown document id', async () => {
    const res = await app.inject({
      method: 'POST',
      url: '/api/documents/00000000-0000-0000-0000-000000000000/tokens',
      payload: { viewer: '테스터' },
    });
    expect(res.statusCode).toBe(404);
  });

  test('rejects a token request without viewer name', async () => {
    const doc = await uploadSamplePdf(1);
    const res = await app.inject({
      method: 'POST',
      url: `/api/documents/${doc.id}/tokens`,
      payload: { canPrint: true },
    });
    expect(res.statusCode).toBe(400);
  });
});

describe('viewer access control', () => {
  test('serves info and page images with a view-only token, but denies print/download', async () => {
    const doc = await uploadSamplePdf(2);
    const token = await issueToken(doc.id, false, false);

    const info = await app.inject({ method: 'GET', url: `/api/view/info?token=${token}` });
    expect(info.statusCode).toBe(200);
    expect(info.json().data.permissions).toEqual({ canPrint: false, canDownload: false });

    const page = await app.inject({ method: 'GET', url: `/api/view/pages/1?token=${token}` });
    expect(page.statusCode).toBe(200);
    expect(page.headers['content-type']).toBe('image/png');
    expect(page.rawPayload.subarray(0, 4)).toEqual(Buffer.from([0x89, 0x50, 0x4e, 0x47]));

    const print = await app.inject({ method: 'GET', url: `/api/view/print?token=${token}` });
    expect(print.statusCode).toBe(403);

    const download = await app.inject({ method: 'GET', url: `/api/view/download?token=${token}` });
    expect(download.statusCode).toBe(403);
  });

  test('allows print and download when the token grants them', async () => {
    const doc = await uploadSamplePdf(1);
    const token = await issueToken(doc.id, true, true);

    const print = await app.inject({ method: 'GET', url: `/api/view/print?token=${token}` });
    expect(print.statusCode).toBe(200);
    expect(print.headers['content-type']).toBe('application/pdf');
    expect(print.rawPayload.subarray(0, 5).toString()).toBe('%PDF-');

    const download = await app.inject({ method: 'GET', url: `/api/view/download?token=${token}` });
    expect(download.statusCode).toBe(200);
    expect(download.headers['content-disposition']).toContain('attachment');
  });

  test('renders higher-resolution pages when a larger scale is requested', async () => {
    const doc = await uploadSamplePdf(1);
    const token = await issueToken(doc.id, false, false);

    const pngWidth = (payload: Buffer) => payload.readUInt32BE(16);
    const base = await app.inject({ method: 'GET', url: `/api/view/pages/1?token=${token}&scale=1` });
    const high = await app.inject({ method: 'GET', url: `/api/view/pages/1?token=${token}&scale=4` });
    expect(base.statusCode).toBe(200);
    expect(high.statusCode).toBe(200);
    expect(pngWidth(high.rawPayload)).toBeGreaterThan(pngWidth(base.rawPayload) * 3);
  });

  test('rejects an out-of-range scale with 400', async () => {
    const doc = await uploadSamplePdf(1);
    const token = await issueToken(doc.id, false, false);
    const res = await app.inject({ method: 'GET', url: `/api/view/pages/1?token=${token}&scale=99` });
    expect(res.statusCode).toBe(400);
  });

  test('returns 404 for a page outside the document range', async () => {
    const doc = await uploadSamplePdf(1);
    const token = await issueToken(doc.id, false, false);
    const res = await app.inject({ method: 'GET', url: `/api/view/pages/99?token=${token}` });
    expect(res.statusCode).toBe(404);
  });

  test('rejects an invalid token with 401', async () => {
    const res = await app.inject({ method: 'GET', url: '/api/view/info?token=bogus-token' });
    expect(res.statusCode).toBe(401);
  });

  test('rejects a missing token with 401', async () => {
    const res = await app.inject({ method: 'GET', url: '/api/view/info' });
    expect(res.statusCode).toBe(401);
  });
});
