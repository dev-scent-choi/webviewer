/**
 * 향기랩 옛 오피스 형식 편집(설계 2026-11-12-hyanggilab-always-edit §2): .doc·.xls·.ppt를 업로드 때 OOXML로 바꿔 편집기로 열고,
 * 내려받기는 원래 형식(안 고쳤으면 올린 바이트 그대로). 업로드에 legacyEdit=1이 있을 때만(향기랩) — 다른 호출자는 지금 그대로.
 */
import { createHash } from 'node:crypto';
import { getCurrentBytes, hasUnsavedEdits } from './collab/room.js';
import { convertOfficeFile } from './converters/libreOffice.js';
import { store } from './storage.js';
import type { DocumentMeta, LegacyExt } from './types.js';

export const LEGACY_EDIT_TARGET = { doc: 'docx', xls: 'xlsx', ppt: 'pptx' } as const;

export function isLegacyExt(ext: string): ext is LegacyExt {
  return Object.hasOwn(LEGACY_EDIT_TARGET, ext);
}

export const sha256 = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');

const VBA_STORAGES = new Set(['Macros', '_VBA_PROJECT_CUR', 'VBA']);
/** VBAInfoContainer(0x03FF) 머리의 recType·길이 20 — 바로 뒤에 VBAInfoAtom(0x0400, 길이 12)이 와야 한다 */
const PPT_VBA = Buffer.from([0xff, 0x03, 0x14, 0x00, 0x00, 0x00]);

/**
 * 매크로(VBA)가 든 옛 오피스 파일인가(설계 §2.3). 워드·엑셀은 복합 파일 디렉터리 항목(128바이트: 이름 UTF-16LE, +64 이름 길이,
 * +66 종류 1 = 저장소) 중 VBA 저장소 이름, PPT는 PowerPoint Document 흐름의 VBAInfo 레코드. 틀려도 경고 한 줄이 더 뜨거나 안 뜰 뿐이다.
 */
export function hasVbaMacros(b: Buffer, ext: LegacyExt): boolean {
  if (b.length < 512 || b.readUInt32LE(0) !== 0xe011cfd0) return false;
  if (ext === 'ppt') {
    for (let i = b.indexOf(PPT_VBA, 2); i >= 2; i = b.indexOf(PPT_VBA, i + 1)) {
      const h = i - 2;
      if ((b[h]! & 0xf) === 0xf && h + 16 <= b.length && b.readUInt16LE(h + 10) === 0x0400 && b.readUInt32LE(h + 12) === 12) return true;
    }
    return false;
  }
  for (let o = 512; o + 128 <= b.length; o += 128) {
    if (b[o + 66] !== 1) continue;
    const len = b.readUInt16LE(o + 64);
    if (len < 4 || len > 64 || len % 2 !== 0) continue;
    if (b.readUInt16LE(o + len - 2) === 0 && VBA_STORAGES.has(b.toString('utf16le', o, o + len - 2))) return true;
  }
  return false;
}

export interface LegacyDownload {
  readonly bytes: Buffer;
  readonly name: string;
  readonly how: 'original' | 'converted' | 'fallback';
}

/**
 * 옛 형식 문서 내려받기(설계 §2.4). 안 고쳤으면(방에 저장 안 된 편집 없음 + 저장된 OOXML 해시가 바꾼 직후와 같음) 올린 바이트 그대로,
 * 고쳤으면 지금 내용을 LibreOffice로 원래 형식으로. 그게 실패하면 고친 것을 잃지 않게 편집본을 OOXML 이름으로(A4).
 */
export async function legacyDownload(meta: DocumentMeta, convert = convertOfficeFile): Promise<LegacyDownload> {
  const legacy = meta.legacy!;
  const stored = await store.readOriginal(meta.id, meta.ext);
  const edited = (await hasUnsavedEdits(meta.id)) || sha256(stored) !== legacy.sha256;
  if (!edited) return { bytes: await store.readDerived(meta.id, 'legacy'), name: meta.name, how: 'original' };
  const current = Buffer.from((await getCurrentBytes(meta.id)) ?? stored);
  try {
    return { bytes: await convert(current, meta.ext, legacy.ext), name: meta.name, how: 'converted' };
  } catch (error) {
    console.error(`옛 형식으로 다시 바꾸지 못해 편집본을 .${meta.ext}로 내려준다(문서 ${meta.id}):`, error);
    return { bytes: current, name: `${meta.name.replace(/.[^./]+$/, '')}.${meta.ext}`, how: 'fallback' };
  }
}
