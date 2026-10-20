/**
 * PDF 편집 방의 부속 PDF(파일에서 넣기·파일 합치기)와 추출·나누기 — viewer.ts의 경로가 쓴다(설계 §2.2·§2.3-6).
 * 부속은 방 메모리에만 두고(방이 닫히며 저장하면 원본에 들어간다), 방마다 합계 크기를 제한한다.
 */
import { randomBytes } from 'node:crypto';
import { zipSync } from 'fflate';
import { pdfRoomState } from '../collab/room.js';
import { buildPdfFromRoom, extractPdf, pdfBaseFingerprint, preparePdfPart } from './pdfModel.js';
import type { PdfAnnot } from './annotTypes.js';

export class PdfRoomError extends Error {
  constructor(
    readonly status: number,
    message: string,
  ) {
    super(message);
  }
}

export interface PdfPartInfo {
  readonly partId: string;
  readonly pageCount: number;
  readonly sizes: ReadonlyArray<{ w: number; h: number; rotate: number }>;
  readonly annots: readonly PdfAnnot[];
}

/** 부속 PDF를 방에 넣는다. maxBytes는 한 파일, 방 합계는 그 4배까지 */
export async function addPdfPart(docId: string, bytes: Uint8Array, maxBytes: number): Promise<PdfPartInfo> {
  const state = await pdfRoomState(docId);
  if (!state) throw new PdfRoomError(409, '편집 연결이 없습니다. 편집기를 다시 열어 주세요.');
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

export async function readPdfPart(docId: string, partId: string): Promise<Uint8Array | null> {
  const state = await pdfRoomState(docId);
  return state?.parts.get(partId) ?? null;
}

/** 편집기가 그릴 원본 — 방이 있으면 방을 열 때의 원본(쪽 목록 열쇠의 기준), 없으면 저장본. 지문을 함께 */
export async function editorSourceOf(docId: string, readStored: () => Promise<Uint8Array>): Promise<{ bytes: Uint8Array; base: string }> {
  const state = await pdfRoomState(docId);
  const bytes = state ? state.original : await readStored();
  return { bytes, base: pdfBaseFingerprint(bytes) };
}

/**
 * 쪽 열쇠 묶음마다 새 PDF. 하나면 pdf, 여럿이면 zip(이름-1.pdf …). 방이 있으면 지금 편집 상태, 없으면 저장본(열쇠 'o:<순번>').
 */
export async function extractPdfGroups(
  docId: string,
  groups: readonly (readonly string[])[],
  baseName: string,
  readStored: () => Promise<Uint8Array>,
): Promise<{ bytes: Uint8Array; type: 'pdf' | 'zip'; fileName: string }> {
  const state = await pdfRoomState(docId);
  let full: Uint8Array;
  let keys: string[];
  if (state) {
    ({ bytes: full, keys } = await buildPdfFromRoom(state.original, state.ydoc, (id) => state.parts.get(id)));
  } else {
    full = await readStored();
    keys = [];
  }
  const indexOf = (key: string): number => {
    if (state) return keys.indexOf(key);
    const m = /^o:(\d+)$/.exec(key);
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
