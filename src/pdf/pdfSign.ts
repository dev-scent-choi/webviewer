/**
 * 인증서 전자서명(9단계, 설계 §3) — 사용자가 올린 인증서(.pfx/.p12 + 비밀번호)로 PAdES(ETSI.CAdES.detached) 서명을 한다.
 * @signpdf(MIT) — 서명 자리(pdflibAddPlaceholder)·서명(P12Signer, node-forge). 인증서와 비밀번호는 이 함수 안에서만 쓰고
 * 저장·기록하지 않는다. 서명한 파일은 내려받기로만 준다(문서는 그대로 — 서명 뒤 고치면 서명이 깨지므로).
 *
 * 참고: npm audit은 node-forge의 RSA 서명 **검증** 취약점(GHSA-86w9-cpqp-85rv)을 표시한다. 여기서는 서명을 **만들기만** 하고
 * 남의 서명을 검증하지 않으므로 그 경로를 쓰지 않는다.
 */
import forge from 'node-forge';
import { P12Signer } from '@signpdf/signer-p12';
import { SignPdf } from '@signpdf/signpdf';
import { pdflibAddPlaceholder } from '@signpdf/placeholder-pdf-lib';
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFRef, type PDFPage } from 'pdf-lib';
import { fontsFor } from './annotWrite.js';

export class PdfSignError extends Error {
  constructor(
    readonly status: number,
    message: string,
  ) {
    super(message);
  }
}

export interface SignOptions {
  readonly p12: Uint8Array;
  readonly password: string;
  readonly reason?: string;
  readonly location?: string;
  /** 보이는 서명 — 쪽 순번(0부터)과 PDF 좌표 사각형. 없으면 보이지 않는 서명 */
  readonly visible?: { readonly pageIndex: number; readonly rect: readonly [number, number, number, number] };
}

export interface CertInfo {
  readonly name: string;
  readonly issuer: string;
  readonly notBefore: Date;
  readonly notAfter: Date;
}

/** node-forge가 주는 이름 값은 UTF-8 바이트 문자열 — 한글 이름을 풀어 준다 */
function decodeName(v: string): string {
  try {
    return forge.util.decodeUtf8(v);
  } catch {
    return v;
  }
}

interface ForgeCert {
  subject: { getField(name: string): { value: string } | null };
  issuer: { getField(name: string): { value: string } | null };
  validity: { notBefore: Date; notAfter: Date };
}

/** 인증서를 열어 이름·기간을 본다 — 비밀번호가 틀리거나 서명용 키가 없으면 오류 */
export function readCertificate(p12: Uint8Array, password: string): CertInfo {
  let parsed: { getBags(q: object): Record<string, Array<{ cert?: ForgeCert; key?: unknown }>> };
  try {
    const asn1 = forge.asn1.fromDer(forge.util.createBuffer(Buffer.from(p12).toString('binary')));
    parsed = forge.pkcs12.pkcs12FromAsn1(asn1, password);
  } catch (error) {
    const msg = String((error as Error)?.message ?? error);
    throw new PdfSignError(400, /password|MAC/i.test(msg) ? '인증서 비밀번호가 맞지 않습니다.' : '인증서 파일을 읽지 못했습니다(.pfx·.p12만 됩니다).');
  }
  const keyBags = [
    ...(parsed.getBags({ bagType: forge.pki.oids.pkcs8ShroudedKeyBag })[forge.pki.oids.pkcs8ShroudedKeyBag] ?? []),
    ...(parsed.getBags({ bagType: forge.pki.oids.keyBag })[forge.pki.oids.keyBag] ?? []),
  ];
  const cert = (parsed.getBags({ bagType: forge.pki.oids.certBag })[forge.pki.oids.certBag] ?? [])[0]?.cert;
  if (!cert || keyBags.length === 0) throw new PdfSignError(400, '인증서 파일에 서명용 키나 인증서가 없습니다.');
  const field = (dn: ForgeCert['subject'], key: string) => decodeName(dn.getField(key)?.value ?? '');
  return {
    name: field(cert.subject, 'CN') || field(cert.subject, 'O') || '(이름 없음)',
    issuer: field(cert.issuer, 'CN') || field(cert.issuer, 'O'),
    notBefore: cert.validity.notBefore,
    notAfter: cert.validity.notAfter,
  };
}

const pad = (n: number) => String(n).padStart(2, '0');
const stamp = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;

/** 서명 칸의 보이는 모양 — 테두리와 "전자서명: 이름 / 일시 / 사유" */
async function drawAppearance(doc: PDFDocument, page: PDFPage, rect: readonly number[], lines: string[]): Promise<void> {
  const fonts = fontsFor(doc);
  const font = await fonts.korean();
  const w = Math.abs(rect[2] - rect[0]);
  const h = Math.abs(rect[3] - rect[1]);
  const size = Math.max(6, Math.min(12, (h - 8) / (lines.length * 1.25)));
  const ops = ['q', '0.1 0.3 0.6 RG 1 w', `0.5 0.5 ${w - 1} ${h - 1} re S`, 'BT', `/F0 ${size} Tf 0.1 0.1 0.15 rg`];
  lines.forEach((line, i) => {
    let text = line;
    while (text.length > 1 && font.widthOfTextAtSize(text, size) > w - 8) text = text.slice(0, -1);
    ops.push(`1 0 0 1 4 ${h - 4 - size * (i + 1) * 1.15} Tm ${font.encodeText(text).toString()} Tj`);
  });
  ops.push('ET', 'Q');
  const ctx = doc.context;
  const ap = ctx.register(ctx.flateStream(ops.join('\n'), { Type: 'XObject', Subtype: 'Form', BBox: [0, 0, w, h], Resources: ctx.obj({ Font: ctx.obj({ F0: font.ref }) }) }));
  // 방금 서명 자리가 만든 위젯(쪽의 마지막 주석)의 모양을 바꾼다
  const annots = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray);
  const last = annots?.get(annots.size() - 1);
  const widget = last instanceof PDFRef ? ctx.lookup(last) : last;
  if (widget instanceof PDFDict) widget.set(PDFName.of('AP'), ctx.obj({ N: ap }));
  await fonts.finish();
}

/** 서명한 PDF 바이트 */
export async function signPdf(input: Uint8Array, opts: SignOptions): Promise<{ bytes: Uint8Array; cert: CertInfo }> {
  const cert = readCertificate(opts.p12, opts.password);
  const now = new Date();
  if (now < cert.notBefore || now > cert.notAfter) {
    throw new PdfSignError(400, `인증서 사용 기간이 아닙니다(${stamp(cert.notBefore)} ~ ${stamp(cert.notAfter)}).`);
  }
  let doc: PDFDocument;
  try {
    doc = await PDFDocument.load(input, { updateMetadata: false });
  } catch {
    throw new PdfSignError(400, 'PDF를 읽지 못해 서명할 수 없습니다.');
  }
  const pages = doc.getPages();
  const visible = opts.visible && pages[opts.visible.pageIndex] ? opts.visible : undefined;
  const page = visible ? pages[visible.pageIndex] : pages[0];
  const reason = (opts.reason ?? '').slice(0, 200);
  const location = (opts.location ?? '').slice(0, 200);
  pdflibAddPlaceholder({
    pdfDoc: doc,
    pdfPage: page,
    reason,
    contactInfo: '',
    name: cert.name,
    location,
    signingTime: now,
    signatureLength: 16384,
    widgetRect: visible ? [...visible.rect] : [0, 0, 0, 0],
    appName: 'Hyanggilab Docview',
  });
  if (visible) {
    const lines = [`전자서명: ${cert.name}`, `일시: ${stamp(now)}`];
    if (reason) lines.push(`사유: ${reason}`);
    await drawAppearance(doc, page, visible.rect, lines);
  }
  const prepared = Buffer.from(await doc.save({ useObjectStreams: false }));
  const signer = new P12Signer(Buffer.from(opts.p12), { passphrase: opts.password });
  const signpdf = new SignPdf();
  try {
    const signed = await signpdf.sign(prepared, signer);
    return { bytes: new Uint8Array(signed), cert };
  } catch (error) {
    console.error('PDF 서명 실패:', (error as Error)?.message);
    throw new PdfSignError(422, '서명하지 못했습니다.');
  }
}
