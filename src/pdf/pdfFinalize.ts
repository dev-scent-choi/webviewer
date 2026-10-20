/**
 * 저장 마지막 단계 — 압축(설정했으면) 다음 암호(설정했으면). 방의 저장·내려받기가 같은 순서로 쓴다(room.ts).
 */
import { compressPdf, type CompressLevel } from './pdfCompress.js';
import { encryptPdf, type PdfProtection } from './pdfSecurity.js';

export async function finalizePdf(bytes: Uint8Array, compress: CompressLevel | null | undefined, protection: PdfProtection | null | undefined): Promise<Uint8Array> {
  let out = bytes;
  if (compress) out = (await compressPdf(out, compress)).bytes;
  if (protection) out = await encryptPdf(out, protection);
  return out;
}
