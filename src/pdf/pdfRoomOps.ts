/**
 * PDF 편집 방의 서버 쪽 작업 — viewer.ts의 경로가 쓴다.
 *  - 부속 PDF(파일에서 넣기·파일 합치기)와 추출·나누기(설계 §2.2·§2.3-6). 부속은 방 메모리에만 두고 방마다 합계 크기를 제한한다.
 *  - 6단계: 암호 걸린 파일 풀기(비밀번호를 맞힌 토큰만 풀린 원본·부속·추출을 쓴다), 보호(암호 걸기·풀기), 압축, 문서 정보.
 */
import { createHash, randomBytes } from 'node:crypto';
import { zipSync } from 'fflate';
import { PDFDocument } from 'pdf-lib';
import { pdfRoomState } from '../collab/room.js';
import { appliedRedactions, buildPdfFromRoom, extractPdf, pdfBaseFingerprint, PDF_META_MAP, PDF_SEED_ORIGIN, preparePdfPart, redactedBase } from './pdfModel.js';
import { compressPdf, isCompressLevel, COMPRESS_LEVELS, type CompressLevel } from './pdfCompress.js';
import { decryptPdf, isEncryptedPdf, readPermissions, summarize, ALL_PERMISSIONS, type PdfPermissions, type PdfProtection } from './pdfSecurity.js';
import { pdfInfo, type PdfInfo } from './pdfProps.js';
import type { PdfAnnot } from './annotTypes.js';

export class PdfRoomError extends Error {
  constructor(
    readonly status: number,
    message: string,
  ) {
    super(message);
  }
}

/** 토큰 → 방에 기억할 열쇠(토큰 자체는 두지 않는다) */
export function tokenKey(token: string): string {
  return createHash('sha256').update(token).digest('hex').slice(0, 32);
}

const LOCKED = () => new PdfRoomError(423, '암호가 걸린 PDF입니다. 비밀번호를 입력해 주세요.');

/** 암호 파일이면 그 토큰이 비밀번호를 맞혔는지 */
async function assertUnlocked(state: NonNullable<Awaited<ReturnType<typeof pdfRoomState>>>, key: string): Promise<void> {
  if (state.encryptedFile && (state.locked || !state.unlocked.has(key))) throw LOCKED();
}

export interface PdfPartInfo {
  readonly partId: string;
  readonly pageCount: number;
  readonly sizes: ReadonlyArray<{ w: number; h: number; rotate: number }>;
  readonly annots: readonly PdfAnnot[];
}

/** 부속 PDF를 방에 넣는다. maxBytes는 한 파일, 방 합계는 그 4배까지 */
export async function addPdfPart(docId: string, bytes: Uint8Array, maxBytes: number, key: string): Promise<PdfPartInfo> {
  const state = await pdfRoomState(docId);
  if (!state) throw new PdfRoomError(409, '편집 연결이 없습니다. 편집기를 다시 열어 주세요.');
  await assertUnlocked(state, key);
  if (bytes.length === 0) throw new PdfRoomError(400, '빈 파일입니다.');
  if (bytes.length > maxBytes) throw new PdfRoomError(413, '파일이 너무 큽니다.');
  let total = 0;
  for (const b of state.parts.values()) total += b.length;
  if (total + bytes.length > maxBytes * 4) throw new PdfRoomError(413, '이 문서에 넣은 파일이 너무 많습니다. 저장한 뒤 다시 열어 주세요.');
  const partId = randomBytes(6).toString('base64url');
  let prepared: Awaited<ReturnType<typeof preparePdfPart>>;
  try {
    prepared = await preparePdfPart(bytes, partId);
  } catch (error) {
    const name = (error as { name?: string })?.name ?? '';
    throw new PdfRoomError(400, /Encrypted/i.test(name) || /encrypt/i.test(String(error)) ? '암호가 걸린 PDF는 넣을 수 없습니다.' : 'PDF 파일을 읽지 못했습니다.');
  }
  state.parts.set(partId, prepared.bytes);
  return { partId, pageCount: prepared.pageCount, sizes: prepared.sizes, annots: prepared.annots };
}

export async function readPdfPart(docId: string, partId: string, key: string): Promise<Uint8Array | null> {
  const state = await pdfRoomState(docId);
  if (!state) return null;
  await assertUnlocked(state, key);
  return state.parts.get(partId) ?? null;
}

const redactedCache = new Map<string, Promise<Uint8Array>>();

/**
 * 편집기가 그릴 원본 — 방이 있으면 방을 열 때의 원본(쪽 목록 열쇠의 기준), 없으면 저장본. 지문은 원래 원본의 것.
 * 방에서 가리기를 적용했으면 그 내용을 지운 원본을 준다(나중에 들어온 참여자에게 가린 글자가 가지 않게).
 * 암호 파일은 비밀번호를 맞힌 토큰에게만(아니면 423 — 편집기가 비밀번호 창을 띄운다).
 */
export async function editorSourceOf(docId: string, readStored: () => Promise<Uint8Array>, key: string): Promise<{ bytes: Uint8Array; base: string }> {
  const state = await pdfRoomState(docId);
  if (!state) {
    const bytes = await readStored();
    if (await isEncryptedPdf(bytes)) throw LOCKED();
    return { bytes, base: pdfBaseFingerprint(bytes) };
  }
  await assertUnlocked(state, key);
  const base = pdfBaseFingerprint(state.original);
  const applied = appliedRedactions(state.ydoc);
  if (applied.length === 0) return { bytes: state.original, base };
  const cacheKey = `${docId}:${base}:${applied.map((r) => r.id).sort().join(',')}`;
  let p = redactedCache.get(cacheKey);
  if (!p) {
    p = redactedBase(state.original, state.ydoc);
    redactedCache.set(cacheKey, p);
    p.catch(() => redactedCache.delete(cacheKey));
    while (redactedCache.size > 8) redactedCache.delete(redactedCache.keys().next().value as string);
  }
  return { bytes: await p, base };
}

/**
 * 쪽 열쇠 묶음마다 새 PDF. 하나면 pdf, 여럿이면 zip(이름-1.pdf …). 방이 있으면 지금 편집 상태, 없으면 저장본(열쇠 'o:<순번>').
 */
export async function extractPdfGroups(
  docId: string,
  groups: readonly (readonly string[])[],
  baseName: string,
  readStored: () => Promise<Uint8Array>,
  key: string,
): Promise<{ bytes: Uint8Array; type: 'pdf' | 'zip'; fileName: string }> {
  const state = await pdfRoomState(docId);
  let full: Uint8Array;
  let keys: string[];
  if (state) {
    await assertUnlocked(state, key);
    ({ bytes: full, keys } = await buildPdfFromRoom(state.original, state.ydoc, (id) => state.parts.get(id)));
  } else {
    full = await readStored();
    if (await isEncryptedPdf(full)) throw LOCKED();
    keys = [];
  }
  const indexOf = (k: string): number => {
    if (state) return keys.indexOf(k);
    const m = /^o:(\d+)$/.exec(k);
    return m ? Number(m[1]) : -1;
  };
  const indexGroups = groups.map((g) => g.map(indexOf).filter((i) => i >= 0)).filter((g) => g.length > 0);
  if (indexGroups.length === 0) throw new PdfRoomError(400, '고른 쪽이 없습니다.');
  const files = await extractPdf(full, indexGroups);
  if (files.length === 0) throw new PdfRoomError(400, '고른 쪽이 없습니다.');
  const stem = baseName.replace(/\.pdf$/i, '') || 'document';
  if (files.length === 1) return { bytes: files[0], type: 'pdf', fileName: `${stem}-추출.pdf` };
  const width = String(files.length).length;
  const entries: Record<string, Uint8Array> = {};
  files.forEach((f, i) => (entries[`${stem}-${String(i + 1).padStart(width, '0')}.pdf`] = f));
  return { bytes: zipSync(entries, { level: 0 }), type: 'zip', fileName: `${stem}-나누기.zip` };
}

// ---------- 6단계: 암호 풀기·걸기, 압축, 정보 ----------

const MAX_UNLOCK_FAILS = 10;

/** 비밀번호로 풀기 — 맞으면 그 토큰을 기억하고, 처음 푼 사람이면 방의 원본을 풀린 것으로 바꾼다(저장 때 같은 비밀번호로 다시 잠근다) */
export async function unlockPdf(docId: string, password: string, key: string): Promise<void> {
  const state = await pdfRoomState(docId);
  if (!state) throw new PdfRoomError(409, '편집 연결이 없습니다. 잠시 뒤 다시 시도해 주세요.');
  if (!state.encryptedFile) {
    state.unlocked.add(key);
    return;
  }
  if (state.unlocked.has(key)) return;
  const decrypted = await decryptPdf(state.encryptedFile, password);
  if (!decrypted) {
    if (state.failUnlock() > MAX_UNLOCK_FAILS) throw new PdfRoomError(429, '비밀번호를 너무 여러 번 틀렸습니다. 편집기를 닫았다가 다시 열어 주세요.');
    throw new PdfRoomError(403, '비밀번호가 맞지 않습니다.');
  }
  if (state.locked) {
    const permissions = await readPermissions(state.encryptedFile);
    await state.unlock(decrypted, { userPassword: password, ownerPassword: password, permissions });
    writeSummary(state, { userPassword: password, ownerPassword: password, permissions }, true);
  }
  state.unlocked.add(key);
}

function writeSummary(state: NonNullable<Awaited<ReturnType<typeof pdfRoomState>>>, protection: PdfProtection | null, fromFile: boolean): void {
  state.ydoc.transact(() => state.ydoc.getMap<unknown>(PDF_META_MAP).set('protection', summarize(protection, fromFile)), PDF_SEED_ORIGIN);
}

export interface ProtectRequest {
  readonly mode: 'set' | 'remove';
  readonly userPassword?: string;
  readonly ownerPassword?: string;
  readonly permissions?: Partial<PdfPermissions>;
}

const pw = (v: unknown) => (typeof v === 'string' ? v : '');

export async function protectPdf(docId: string, req: ProtectRequest, key: string, viewer: string): Promise<void> {
  const state = await pdfRoomState(docId);
  if (!state) throw new PdfRoomError(409, '편집 연결이 없습니다. 편집기를 다시 열어 주세요.');
  await assertUnlocked(state, key);
  let protection: PdfProtection | null = null;
  if (req.mode === 'set') {
    const user = pw(req.userPassword);
    const owner = pw(req.ownerPassword);
    if (user.length > 128 || owner.length > 128) throw new PdfRoomError(400, '비밀번호는 128자까지입니다.');
    const permissions: PdfPermissions = { ...ALL_PERMISSIONS, ...Object.fromEntries(Object.entries(req.permissions ?? {}).filter(([, v]) => typeof v === 'boolean')) };
    const restricted = Object.values(permissions).some((v) => !v);
    if (!user && !owner) throw new PdfRoomError(400, '열기 비밀번호나 권한 비밀번호 중 하나는 넣어 주세요.');
    if (restricted && !owner && !user) throw new PdfRoomError(400, '권한을 제한하려면 권한 비밀번호가 필요합니다.');
    // 권한 비밀번호가 없으면 열기 비밀번호로(권한을 바꾸려면 그 비밀번호가 필요하다)
    protection = { userPassword: user, ownerPassword: owner || user, permissions };
  }
  state.setProtection(protection);
  writeSummary(state, protection, false);
  state.tracker?.record('protection', {
    kind: 'protection_changed',
    label: '문서 보호',
    page: null,
    text: protection ? (protection.userPassword ? '열기 비밀번호·권한 설정' : '권한 설정') : '암호 풀기',
    viewer,
    at: new Date().toISOString(),
  });
}

export interface CompressPreview {
  readonly before: number;
  readonly after: number;
  readonly images: number;
}

/** 압축 — 미리 보기(dryRun)면 지금 상태 그대로의 크기와 압축한 크기, 아니면 저장 설정으로 둔다(level null이면 끔) */
export async function compressRoomPdf(docId: string, level: CompressLevel | null, dryRun: boolean, key: string, viewer: string): Promise<CompressPreview | null> {
  const state = await pdfRoomState(docId);
  if (!state) throw new PdfRoomError(409, '편집 연결이 없습니다. 편집기를 다시 열어 주세요.');
  await assertUnlocked(state, key);
  if (level !== null && !isCompressLevel(level)) throw new PdfRoomError(400, '압축 단계가 올바르지 않습니다.');
  if (dryRun) {
    if (!level) throw new PdfRoomError(400, '압축 단계를 고르세요.');
    const { bytes } = await buildPdfFromRoom(state.original, state.ydoc, (id) => state.parts.get(id));
    const result = await compressPdf(bytes, level);
    return { before: bytes.length, after: Math.min(bytes.length, result.bytes.length), images: result.images };
  }
  state.setCompress(level);
  state.ydoc.transact(() => state.ydoc.getMap<unknown>(PDF_META_MAP).set('compress', level), PDF_SEED_ORIGIN);
  state.tracker?.record('compress', {
    kind: 'compression_changed',
    label: '압축',
    page: null,
    text: level ? COMPRESS_LEVELS[level].label : '압축 안 함',
    viewer,
    at: new Date().toISOString(),
  });
  return null;
}

/** 문서 정보(버전·쪽 크기·글꼴…) — 지금 편집 상태 기준, 방이 없으면 저장본 */
export async function pdfInfoOf(docId: string, readStored: () => Promise<Uint8Array>, key: string): Promise<PdfInfo> {
  const state = await pdfRoomState(docId);
  let bytes: Uint8Array;
  if (state) {
    await assertUnlocked(state, key);
    bytes = (await buildPdfFromRoom(state.original, state.ydoc, (id) => state.parts.get(id))).bytes;
  } else {
    bytes = await readStored();
    if (await isEncryptedPdf(bytes)) throw LOCKED();
  }
  // 버전은 원래 파일 머리글에서(다시 저장한 것은 pdf-lib의 1.7)
  const head = state ? state.original : bytes;
  return { ...pdfInfo(await PDFDocument.load(bytes, { updateMetadata: false }), head) };
}
