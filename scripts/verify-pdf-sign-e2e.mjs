// PDF 편집기 9단계(인증서 전자서명) 브라우저 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-pdf-sign-e2e.mjs [BASE] [스크린샷 폴더]
// 시험용 인증서(.p12, 한글 이름)를 만들어 서명 탭에서 틀린 비밀번호 → 안내, 맞는 비밀번호 → 내려받은 PDF의 서명이
// 파일 전체를 덮고 인증서로 맞는지(Node crypto), 보이는 서명 칸·편집한 주석이 들어갔는지 본다. 읽기 전용은 버튼·API 모두 막힘.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createHash, verify as cryptoVerify } from 'node:crypto';
import fontkit from '@pdf-lib/fontkit';
import forge from 'node-forge';
import { PDFDocument, rgb } from 'pdf-lib';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-sign-'));
fs.mkdirSync(shotDir, { recursive: true });
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

function makeP12(name, password) {
  const keys = forge.pki.rsa.generateKeyPair(2048);
  const cert = forge.pki.createCertificate();
  cert.publicKey = keys.publicKey;
  cert.serialNumber = '0b';
  cert.validity.notBefore = new Date(Date.now() - 86400000);
  cert.validity.notAfter = new Date(Date.now() + 30 * 86400000);
  const attrs = [{ name: 'commonName', value: name, valueTagClass: forge.asn1.Type.UTF8 }];
  cert.setSubject(attrs);
  cert.setIssuer(attrs);
  cert.sign(keys.privateKey, forge.md.sha256.create());
  const der = forge.asn1.toDer(forge.pkcs12.toPkcs12Asn1(keys.privateKey, [cert], password, { algorithm: 'aes256' })).getBytes();
  return { p12: Buffer.from(der, 'binary'), certPem: forge.pki.certificateToPem(cert) };
}

function verifySigned(bytes, certPem) {
  const buf = Buffer.from(bytes);
  const s = buf.toString('latin1');
  const m = /\/ByteRange\s*\[\s*(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s*\]/.exec(s);
  if (!m) return { covers: false, digestOk: false, signatureOk: false };
  const br = m.slice(1).map(Number);
  const signed = Buffer.concat([buf.subarray(br[0], br[0] + br[1]), buf.subarray(br[2], br[2] + br[3])]);
  const cms = Buffer.from(s.slice(br[0] + br[1] + 1, br[2] - 1), 'hex');
  const msg = forge.pkcs7.messageFromAsn1(forge.asn1.fromDer(forge.util.createBuffer(cms.toString('binary')), { parseAllBytes: false }));
  const set = forge.asn1.create(forge.asn1.Class.UNIVERSAL, forge.asn1.Type.SET, true, msg.rawCapture.authenticatedAttributes);
  const attrsDer = Buffer.from(forge.asn1.toDer(set).getBytes(), 'binary');
  return {
    covers: br[0] === 0 && br[2] + br[3] === buf.length,
    digestOk: attrsDer.includes(createHash('sha256').update(signed).digest()),
    signatureOk: cryptoVerify('sha256', attrsDer, certPem, Buffer.from(msg.rawCapture.signature, 'binary')),
  };
}

async function koreanPdf() {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const font = await doc.embedFont(fs.readFileSync('assets/fonts/Pretendard-Regular.ttf'), { subset: false });
  for (const t of ['공급 계약서', '2쪽 — 서명란']) doc.addPage([595, 842]).drawText(t, { x: 60, y: 760, size: 24, font, color: rgb(0.1, 0.2, 0.4) });
  return doc.save();
}

async function upload(bytes, name) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패: ${JSON.stringify(r)}`);
  for (let i = 0; i < 200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready' || st.editReady) return r.data.id;
    if (st.status === 'error') throw new Error(`변환 실패: ${JSON.stringify(st)}`);
    await new Promise((res) => setTimeout(res, 150));
  }
  throw new Error('대기 시간 초과');
}
const token = (id, body = {}) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '시험 편집자', canPrint: true, canDownload: true, canCopy: true, canEdit: true, pdfEditor: true, ...body }),
  }).then((r) => r.json());

const { p12, certPem } = makeP12('김서명', 'cert-pw-1');
const certPath = path.join(shotDir, 'test-cert.p12');
fs.writeFileSync(certPath, p12);
const docId = await upload(await koreanPdf(), '공급계약서.pdf');
const t1 = await token(docId);

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1500, height: 950 }, acceptDownloads: true });
const errors = [];
const page = await ctx.newPage();
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => {
  if (m.type() === 'error' && !/400/.test(m.text())) errors.push(m.text());
});
await page.goto(`${BASE}${t1.data.editorUrl}`);
await page.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });

// 편집 — 1쪽에 사각형 주석
await page.click('.pdf-mode[data-mode="edit"]');
await page.click('[data-group="shape"] .pdf-float-more');
await page.click('[data-group="shape"] [data-pick="rect"]');
const pb = await page.locator('.pdf-page[data-page="1"]').boundingBox();
await page.mouse.move(pb.x + 80, pb.y + 200);
await page.mouse.down();
await page.mouse.move(pb.x + 240, pb.y + 280, { steps: 4 });
await page.mouse.up();
await page.waitForTimeout(300);

// 모든 도구 목록에서 열기 → 닫기
await page.click('.pdf-mode[data-mode="tools"]');
await page.click('.pdf-tool-item[data-tool="certsign"]');
await page.waitForSelector('#cert-dialog:not([hidden])');
check(true, '모든 도구 → 인증서 서명 창');
await page.click('#cert-dialog [data-close]');

// 서명 탭 — 2쪽으로 가서 서명
await page.click('.pdf-mode[data-mode="sign"]');
await page.fill('#page-input', '2');
await page.press('#page-input', 'Enter');
await page.waitForTimeout(500);
await page.click('#sign-cert');
await page.waitForSelector('#cert-dialog:not([hidden])');
await page.click('#cert-run');
check((await page.textContent('#cert-status')).includes('고르세요'), '인증서 없이 누르면 안내');
await page.setInputFiles('#cert-file', certPath);
await page.fill('#cert-password', 'wrong');
await page.fill('#cert-reason', '계약 승인');
await page.fill('#cert-location', '서울');
await page.click('#cert-run');
await page.waitForFunction(() => /맞지 않/.test(document.getElementById('cert-status')?.textContent ?? ''), null, { timeout: 30000 });
check(true, '틀린 비밀번호 안내', await page.textContent('#cert-status'));
await page.fill('#cert-password', 'cert-pw-1');
await page.screenshot({ path: path.join(shotDir, 'cert-dialog.png') });
const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 60000 }), page.click('#cert-run')]);
const signedPath = path.join(shotDir, 'signed.pdf');
await dl.saveAs(signedPath);
const signed = fs.readFileSync(signedPath);
check(dl.suggestedFilename() === '공급계약서-서명.pdf', '내려받은 파일 이름', dl.suggestedFilename());
const v = verifySigned(signed, certPem);
check(v.covers && v.digestOk && v.signatureOk, '서명이 파일 전체를 덮고 인증서로 맞음', JSON.stringify(v));
await page.waitForSelector('#cert-dialog', { state: 'hidden' });
check((await page.inputValue('#cert-password')) === '' && (await page.inputValue('#cert-file')) === '', '서명 뒤 창을 닫고 비밀번호·인증서 비움');

const pdf = await getDocument({ data: new Uint8Array(signed), isEvalSupported: false }).promise;
const a1 = (await (await pdf.getPage(1)).getAnnotations()).map((a) => a.subtype);
const a2 = await (await pdf.getPage(2)).getAnnotations();
const sig = a2.find((a) => a.fieldType === 'Sig');
await pdf.destroy();
check(a1.includes('Square'), '편집한 주석이 서명 파일에 들어감', a1.join(','));
check(!!sig && sig.rect[0] > 300 && sig.rect[1] < 120, '2쪽 오른쪽 아래에 보이는 서명 칸', sig ? sig.rect.map(Math.round).join(',') : '없음');
check(signed.toString('latin1').includes('/Reason') && /\/M \(D:\d{14}/.test(signed.toString('latin1')), '사유·서명 시각 기록');

// 원본 문서는 그대로(서명 없음)
const plain = Buffer.from(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(t1.data.token)}`)).arrayBuffer());
check(!plain.toString('latin1').includes('/ByteRange'), '문서 자체는 서명되지 않음(내려받기만)');

// 보이지 않는 서명(API 직접)
const inv = await fetch(`${BASE}/api/view/pdf-sign?token=${encodeURIComponent(t1.data.token)}`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({ p12: p12.toString('base64'), password: 'cert-pw-1' }),
});
const invBytes = Buffer.from(await inv.arrayBuffer());
check(inv.status === 200 && verifySigned(invBytes, certPem).signatureOk, '보이지 않는 서명', String(inv.status));

// 읽기 전용 — 버튼 없음, API 거부
const ro = await ctx.newPage();
await ro.goto(`${BASE}${(await token(docId, { viewer: '읽기만', canEdit: false })).data.editorUrl}`);
await ro.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
const roTools = await ro.$$eval('.pdf-tool-item', (els) => els.filter((e) => e.offsetParent !== null).map((e) => e.dataset.tool));
check(!roTools.includes('certsign'), '읽기 전용: 인증서 서명 없음', roTools.join(','));
const roSign = await fetch(`${BASE}/api/view/pdf-sign?token=${encodeURIComponent((await token(docId, { canEdit: false })).data.token)}`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({ p12: p12.toString('base64'), password: 'cert-pw-1' }),
});
check(roSign.status === 403, '읽기 전용 토큰으로 서명 거부', String(roSign.status));
const noDl = await fetch(`${BASE}/api/view/pdf-sign?token=${encodeURIComponent((await token(docId, { canDownload: false })).data.token)}`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({ p12: p12.toString('base64'), password: 'cert-pw-1' }),
});
check(noDl.status === 403, '내려받기 권한 없으면 서명 거부', String(noDl.status));
await ro.close();

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 3).join(' | '));
await browser.close();
await fetch(`${BASE}/api/documents/${docId}`, { method: 'DELETE' });
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
