/**
 * 조각 업로드 — 요청 하나에 100MB까지만 받는 앞단(Cloudflare 등)을 거치는 곳(향기랩)에서도 큰 파일을 올릴 수 있게,
 * 브라우저가 파일을 여러 조각으로 나눠 보내면 여기서 임시 폴더에 모았다가 마지막 조각이 오면 하나로 합친다.
 * 조각은 업로드 ID(브라우저가 만든 UUID) 폴더에 순번 파일로 둔다. 1시간 넘게 마무리되지 않은 업로드는 지운다.
 */
import { randomUUID } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';

const ROOT = path.join(os.tmpdir(), 'docview-chunks');
const UPLOAD_ID_RE = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
/** 조각 수 상한 — 64MB 조각이면 4GB, 실제 상한은 maxBytes가 정한다 */
export const MAX_CHUNKS = 64;
const STALE_MS = 60 * 60 * 1000;

export class ChunkUploadError extends Error {
  constructor(
    readonly status: number,
    message: string,
  ) {
    super(message);
  }
}

const assembling = new Set<string>();
let lastSweep = 0;

/** 오래된(마무리되지 않은) 업로드 폴더 정리 — 10분에 한 번만 */
async function sweepStale(): Promise<void> {
  if (Date.now() - lastSweep < 10 * 60 * 1000) return;
  lastSweep = Date.now();
  const names = await fs.readdir(ROOT).catch(() => [] as string[]);
  for (const name of names) {
    const dir = path.join(ROOT, name);
    const stat = await fs.stat(dir).catch(() => null);
    if (stat && Date.now() - stat.mtimeMs > STALE_MS) await fs.rm(dir, { recursive: true, force: true }).catch(() => undefined);
  }
}

/**
 * 조각 하나를 받는다. 모든 조각이 모이면 합친 바이트를 돌려주고 임시 폴더를 지운다.
 * 조각은 순서대로 오지 않아도 된다(같은 순번이 다시 오면 덮어쓴다).
 */
export async function acceptChunk(
  uploadId: string,
  index: number,
  total: number,
  data: Buffer,
  maxBytes: number,
): Promise<{ done: false; received: number } | { done: true; bytes: Buffer }> {
  if (!UPLOAD_ID_RE.test(uploadId)) throw new ChunkUploadError(400, '업로드 ID가 올바르지 않습니다.');
  if (!Number.isInteger(total) || total < 1 || total > MAX_CHUNKS) throw new ChunkUploadError(400, '조각 수가 올바르지 않습니다.');
  if (!Number.isInteger(index) || index < 0 || index >= total) throw new ChunkUploadError(400, '조각 순번이 올바르지 않습니다.');
  if (assembling.has(uploadId)) throw new ChunkUploadError(409, '이미 합치는 중인 업로드입니다.');
  void sweepStale();
  const dir = path.join(ROOT, uploadId.toLowerCase());
  await fs.mkdir(dir, { recursive: true });
  // 같은 순번이 동시에 와도 반쪽 파일이 남지 않게 임시 이름으로 쓴 뒤 바꾼다
  const tmp = path.join(dir, `${index}.${randomUUID()}.part`);
  await fs.writeFile(tmp, data);
  await fs.rename(tmp, path.join(dir, String(index)));

  const names = (await fs.readdir(dir)).filter((n) => /^\d+$/.test(n));
  let size = 0;
  for (const name of names) size += (await fs.stat(path.join(dir, name))).size;
  if (size > maxBytes) {
    await fs.rm(dir, { recursive: true, force: true });
    throw new ChunkUploadError(413, '파일이 업로드 용량 제한을 초과했습니다.');
  }
  if (names.length < total) return { done: false, received: names.length };

  assembling.add(uploadId);
  try {
    const parts: Buffer[] = [];
    for (let i = 0; i < total; i += 1) parts.push(await fs.readFile(path.join(dir, String(i))));
    return { done: true, bytes: Buffer.concat(parts) };
  } finally {
    assembling.delete(uploadId);
    await fs.rm(dir, { recursive: true, force: true }).catch(() => undefined);
  }
}
