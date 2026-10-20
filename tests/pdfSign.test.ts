import { createHash, verify as cryptoVerify } from 'node:crypto';
import forge from 'node-forge';
import { PDFDict, PDFDocument, PDFName, StandardFonts } from 'pdf-lib';
import { describe, expect, test } from 'vitest';
import { readCertificate, signPdf } from '../src/pdf/pdfSign.js';

/** 시험용 자체 서명 인증서(.p12) — 이름은 UTF-8(한글 가능), 기간은 days(음수면 이미 지난 것) */
function makeP12(name: string, password: string, days = 365): { p12: Uint8Array; certPem: string } {
  const keys = forge.pki.rsa.generateKeyPair(2048);
  const cert = forge.pki.createCertificate();
  cert.publicKey = keys.publicKey;
  cert.serialNumber = '0a';
  const now = Date.now();
  cert.validity.notBefore = new Date(now - 2 * 86400000 + (days < 0 ? days * 86400000 : 0));
  cert.validity.notAfter = new Date(now + days * 86400000);
  const attrs = [{ name: 'commonName', value: name, valueTagClass: forge.asn1.Type.UTF8 }];
  cert.setSubject(attrs);
  cert.setIssuer(attrs);
  cert.sign(keys.privateKey, forge.md.sha256.create());
  const der = forge.asn1.toDer(forge.pkcs12.toPkcs12Asn1(keys.privateKey, [cert], password, { algorithm: 'aes256' })).getBytes();
  return { p12: new Uint8Array(Buffer.from(der, 'binary')), certPem: forge.pki.certificateToPem(cert) };
}

async function samplePdf(): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  doc.addPage([400, 300]).drawText('Contract body', { x: 20, y: 250, size: 16, font });
  doc.addPage([400, 300]);
  return doc.save();
}

/** 서명 검증(node-forge는 구조를 풀기만, 서명 확인은 Node crypto) */
function verifySigned(bytes: Uint8Array, certPem: string): { covers: boolean; digestOk: boolean; signatureOk: boolean } {
  const buf = Buffer.from(bytes);
  const s = buf.toString('latin1');
  const br = /\/ByteRange\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]/.exec(s)!.slice(1).map(Number);
  const signedContent = Buffer.concat([buf.subarray(br[0], br[0] + br[1]), buf.subarray(br[2], br[2] + br[3])]);
  const hex = s.slice(br[0] + br[1] + 1, br[2] - 1);
  const cms = Buffer.from(hex, 'hex');
  // 서명 값 자리는 뒤를 0으로 채운다 — 남는 바이트는 무시하고 푼다
  const msg = forge.pkcs7.messageFromAsn1(forge.asn1.fromDer(forge.util.createBuffer(cms.toString('binary')), { parseAllBytes: false }));
  const attrs = msg.rawCapture.authenticatedAttributes as unknown[];
  const set = forge.asn1.create(forge.asn1.Class.UNIVERSAL, forge.asn1.Type.SET, true, attrs);
  const signedAttrsDer = Buffer.from(forge.asn1.toDer(set).getBytes(), 'binary');
  const digest = createHash('sha256').update(signedContent).digest();
  const digestOk = signedAttrsDer.includes(digest);
  const signature = Buffer.from(msg.rawCapture.signature, 'binary');
  const signatureOk = cryptoVerify('sha256', signedAttrsDer, certPem, signature);
  return { covers: br[0] === 0 && br[2] + br[3] === buf.length, digestOk, signatureOk };
}

describe('인증서 전자서명(9단계)', () => {
  test('보이지 않는 서명 — 파일 전체를 덮고, 해시·서명 값이 인증서로 맞는다', async () => {
    const { p12, certPem } = makeP12('Hong Gildong', 'pw-1');
    const { bytes, cert } = await signPdf(await samplePdf(), { p12, password: 'pw-1', reason: '승인' });
    expect(cert.name).toBe('Hong Gildong');
    expect(verifySigned(bytes, certPem)).toEqual({ covers: true, digestOk: true, signatureOk: true });
    // 서명 뒤 한 바이트라도 바꾸면 해시가 맞지 않는다
    const tampered = Buffer.from(bytes);
    const at = tampered.indexOf('%PDF-') + 7; // 파일 머리의 버전 숫자(서명 범위 안)
    tampered[at] = tampered[at] === 0x37 ? 0x36 : 0x37;
    expect(verifySigned(tampered, certPem).digestOk).toBe(false);
  }, 60_000);

  test('보이는 서명 — 지정한 쪽에 서명 칸(모양 포함), 한글 이름', async () => {
    const { p12, certPem } = makeP12('홍길동', 'pw-2');
    expect(readCertificate(p12, 'pw-2').name).toBe('홍길동');
    const { bytes } = await signPdf(await samplePdf(), { p12, password: 'pw-2', visible: { pageIndex: 1, rect: [200, 20, 380, 80] } });
    expect(verifySigned(bytes, certPem).signatureOk).toBe(true);
    const doc = await PDFDocument.load(bytes);
    const fields = doc.getForm().getFields();
    expect(fields.map((f) => f.constructor.name)).toContain('PDFSignature');
    const annots = doc.getPage(1).node.Annots()!;
    const widget = doc.context.lookup(annots.get(annots.size() - 1)) as PDFDict;
    expect(widget.get(PDFName.of('FT'))).toBe(PDFName.of('Sig'));
    expect(widget.has(PDFName.of('AP'))).toBe(true);
    expect(doc.getPage(0).node.Annots()?.size() ?? 0).toBe(0);
  }, 60_000);

  test('비밀번호가 틀리거나 기간이 지난 인증서는 서명하지 않는다', async () => {
    const { p12 } = makeP12('Hong', 'right');
    await expect(signPdf(await samplePdf(), { p12, password: 'wrong' })).rejects.toMatchObject({ status: 400, message: '인증서 비밀번호가 맞지 않습니다.' });
    const expired = makeP12('Old', 'pw', -10);
    await expect(signPdf(await samplePdf(), { p12: expired.p12, password: 'pw' })).rejects.toMatchObject({ status: 400 });
    await expect(signPdf(await samplePdf(), { p12: new Uint8Array([1, 2, 3]), password: 'x' })).rejects.toMatchObject({ status: 400 });
  }, 60_000);
});
