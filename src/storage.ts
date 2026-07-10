import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
import { config } from './config.js';
import type { DocumentMeta } from './types.js';

const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/;
const EXT_PATTERN = /^[a-z0-9]{1,10}$/;

export async function ensureDataDirs(): Promise<void> {
  const dirs = [config.originalsDir, config.convertedDir, config.metaDir, config.auditDir];
  await Promise.all(dirs.map((dir) => fs.mkdir(dir, { recursive: true })));
}

export function newDocumentId(): string {
  return randomUUID();
}

function assertValidId(id: string): void {
  if (!UUID_PATTERN.test(id)) {
    throw new Error('잘못된 문서 ID 형식입니다.');
  }
}

function assertValidExt(ext: string): void {
  if (!EXT_PATTERN.test(ext)) {
    throw new Error('잘못된 확장자 형식입니다.');
  }
}

export function originalPath(id: string, ext: string): string {
  assertValidId(id);
  assertValidExt(ext);
  return path.join(config.originalsDir, `${id}.${ext}`);
}

export function convertedPath(id: string): string {
  assertValidId(id);
  return path.join(config.convertedDir, `${id}.pdf`);
}

export async function saveOriginal(id: string, ext: string, data: Buffer): Promise<void> {
  await fs.writeFile(originalPath(id, ext), data);
}

export async function writeMeta(meta: DocumentMeta): Promise<void> {
  assertValidId(meta.id);
  const target = path.join(config.metaDir, `${meta.id}.json`);
  await fs.writeFile(target, JSON.stringify(meta, null, 2), 'utf8');
}

export async function readMeta(id: string): Promise<DocumentMeta | null> {
  if (!UUID_PATTERN.test(id)) {
    return null;
  }
  try {
    const raw = await fs.readFile(path.join(config.metaDir, `${id}.json`), 'utf8');
    return JSON.parse(raw) as DocumentMeta;
  } catch {
    return null;
  }
}

export async function listMeta(): Promise<readonly DocumentMeta[]> {
  let entries: string[];
  try {
    entries = await fs.readdir(config.metaDir);
  } catch {
    return [];
  }
  const ids = entries
    .filter((name) => name.endsWith('.json'))
    .map((name) => name.slice(0, -'.json'.length));
  const metas = await Promise.all(ids.map((id) => readMeta(id)));
  return metas
    .filter((meta): meta is DocumentMeta => meta !== null)
    .sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}
