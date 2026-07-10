import path from 'node:path';
import { fileURLToPath } from 'node:url';
import multipart from '@fastify/multipart';
import fastifyStatic from '@fastify/static';
import Fastify, { type FastifyInstance } from 'fastify';
import { config } from './config.js';
import { documentRoutes } from './routes/documents.js';
import { viewerRoutes } from './routes/viewer.js';

const publicDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'public');

export async function buildApp(): Promise<FastifyInstance> {
  const app = Fastify({
    logger: { level: process.env.LOG_LEVEL ?? 'info' },
    bodyLimit: config.maxUploadBytes,
  });

  await app.register(multipart, {
    limits: { fileSize: config.maxUploadBytes, files: 1 },
  });
  await app.register(fastifyStatic, {
    root: publicDir,
    // 뷰어/관리 화면 JS가 브라우저에 캐시되어 구버전이 돌지 않도록 항상 재검증시킨다.
    setHeaders(res) {
      res.setHeader('cache-control', 'no-cache');
    },
  });
  await app.register(documentRoutes);
  await app.register(viewerRoutes);

  return app;
}
