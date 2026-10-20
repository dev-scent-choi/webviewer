/**
 * PDF 암호 — 풀기(비밀번호로 연 사본)·걸기(AES-256, 사용자/소유자 비밀번호, 권한). 설계 §3 "암호 걸기·풀기".
 * 암호 처리만 @cantoo/pdf-lib(pdf-lib 포크, MIT)에 맡기고, 나머지 편집은 지금의 pdf-lib 그대로 — 둘은 바이트로만 주고받는다.
 * 비밀번호는 방(서버 메모리)에만 두고 Y.Doc(참여자 모두에게 가는 문서)에는 넣지 않는다.
 */
import { PDFDocument as CantooDocument } from '@cantoo/pdf-lib';
import { PDFDict, PDFDocument, PDFName, PDFNumber, PDFRef } from 'pdf-lib';

export interface PdfPermissions {
  readonly printing: boolean;
  readonly copying: boolean;
  readonly modifying: boolean;
  readonly annotating: boolean;
  readonly fillingForms: boolean;
}

export interface PdfProtection {
  readonly userPassword: string;
  readonly ownerPassword: string;
  readonly permissions: PdfPermissions;
}

export const ALL_PERMISSIONS: PdfPermissions = { printing: true, copying: true, modifying: true, annotating: true, fillingForms: true };

/** 암호가 걸린 PDF인지(pdf-lib가 그대로 열지 못하면) */
export async function isEncryptedPdf(bytes: Uint8Array): Promise<boolean> {
  try {
    const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
    return doc.isEncrypted;
  } catch {
    return false;
  }
}

/** 암호 사전의 /P → 권한(ISO 32000 표 22의 비트) */
export async function readPermissions(bytes: Uint8Array): Promise<PdfPermissions> {
  try {
    const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
    const enc = doc.context.trailerInfo.Encrypt;
    const dict = enc instanceof PDFRef ? doc.context.lookup(enc) : enc;
    const p = dict instanceof PDFDict ? dict.get(PDFName.of('P')) : undefined;
    if (!(p instanceof PDFNumber)) return ALL_PERMISSIONS;
    const bits = p.asNumber() >>> 0;
    const bit = (n: number) => (bits & (1 << (n - 1))) !== 0;
    return { printing: bit(3), modifying: bit(4), copying: bit(5), annotating: bit(6), fillingForms: bit(9) || bit(6) };
  } catch {
    return ALL_PERMISSIONS;
  }
}

/** 비밀번호로 푼 사본 — 틀리면 null */
export async function decryptPdf(bytes: Uint8Array, password: string): Promise<Uint8Array | null> {
  let plain: Uint8Array;
  try {
    const doc = await CantooDocument.load(bytes, { password, updateMetadata: false });
    plain = await doc.save();
  } catch (error) {
    if (/password/i.test(String((error as Error)?.message ?? error))) return null;
    throw error;
  }
  return restoreInfo(plain, bytes, password);
}

/**
 * @cantoo/pdf-lib는 푼 사본을 저장할 때 문서 정보 사전(제목·작성자…)을 잃는다 — pdf.js로 원래 파일의 정보를 읽어 다시 적는다.
 */
async function restoreInfo(plain: Uint8Array, encrypted: Uint8Array, password: string): Promise<Uint8Array> {
  let info: Record<string, unknown> = {};
  try {
    const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const pdf = await getDocument({ data: new Uint8Array(encrypted), password, isEvalSupported: false, verbosity: 0 }).promise;
    info = ((await pdf.getMetadata()).info ?? {}) as Record<string, unknown>;
    await pdf.destroy();
  } catch {
    return plain;
  }
  const doc = await PDFDocument.load(plain, { updateMetadata: false });
  const str = (k: string) => (typeof info[k] === 'string' && info[k] ? (info[k] as string) : undefined);
  const date = (k: string) => {
    const m = /^D:(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?/.exec(str(k) ?? '');
    if (!m) return undefined;
    const d = new Date(Date.UTC(Number(m[1]), Number(m[2] ?? 1) - 1, Number(m[3] ?? 1), Number(m[4] ?? 0), Number(m[5] ?? 0), Number(m[6] ?? 0)));
    return Number.isNaN(d.getTime()) ? undefined : d;
  };
  const title = str('Title');
  if (title) doc.setTitle(title);
  const author = str('Author');
  if (author) doc.setAuthor(author);
  const subject = str('Subject');
  if (subject) doc.setSubject(subject);
  const keywords = str('Keywords');
  if (keywords) doc.setKeywords(keywords.split(/[,;]\s*/).filter(Boolean));
  const creator = str('Creator');
  if (creator) doc.setCreator(creator);
  const producer = str('Producer');
  if (producer) doc.setProducer(producer);
  const created = date('CreationDate');
  if (created) doc.setCreationDate(created);
  const modified = date('ModDate');
  if (modified) doc.setModificationDate(modified);
  return doc.save();
}

/** 저장 마지막 단계 — AES-256으로 잠근다 */
export async function encryptPdf(bytes: Uint8Array, protection: PdfProtection): Promise<Uint8Array> {
  const doc = await CantooDocument.load(bytes, { updateMetadata: false });
  const p = protection.permissions;
  doc.encrypt({
    userPassword: protection.userPassword,
    ownerPassword: protection.ownerPassword || protection.userPassword,
    permissions: {
      printing: p.printing ? 'highResolution' : undefined,
      copying: p.copying,
      modifying: p.modifying,
      annotating: p.annotating,
      fillingForms: p.fillingForms,
      contentAccessibility: true,
      documentAssembly: p.modifying,
    },
  });
  return doc.save();
}

/** 화면에 보이는 보호 상태(비밀번호는 빼고) — Y.Doc의 pdfMeta 'protection'에 둔다 */
export interface ProtectionSummary {
  readonly enabled: boolean;
  /** 열 때 비밀번호가 필요한지 */
  readonly openPassword: boolean;
  readonly permissions: PdfPermissions;
  /** 처음 올린 파일에 암호가 있었는지 */
  readonly fromFile: boolean;
}

export function summarize(protection: PdfProtection | null, fromFile: boolean): ProtectionSummary {
  return {
    enabled: !!protection,
    openPassword: !!protection?.userPassword,
    permissions: protection?.permissions ?? ALL_PERMISSIONS,
    fromFile,
  };
}
