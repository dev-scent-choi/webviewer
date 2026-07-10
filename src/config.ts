import path from 'node:path';

const dataRoot = process.env.DOCVIEW_DATA_DIR ?? path.resolve(process.cwd(), 'data');

const DEFAULT_PORT = 8000;
const DEFAULT_TOKEN_TTL_SECONDS = 60 * 60;
const MAX_UPLOAD_BYTES = 100 * 1024 * 1024;
const RENDER_SCALE = 2;

export const config = {
  port: Number(process.env.DOCVIEW_PORT ?? DEFAULT_PORT),
  host: process.env.DOCVIEW_HOST ?? '127.0.0.1',
  dataRoot,
  originalsDir: path.join(dataRoot, 'originals'),
  convertedDir: path.join(dataRoot, 'converted'),
  metaDir: path.join(dataRoot, 'meta'),
  auditDir: path.join(dataRoot, 'audit'),
  tokenTtlSeconds: Number(process.env.DOCVIEW_TOKEN_TTL ?? DEFAULT_TOKEN_TTL_SECONDS),
  maxUploadBytes: MAX_UPLOAD_BYTES,
  renderScale: RENDER_SCALE,
} as const;
