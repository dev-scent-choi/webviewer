// PDF 편집기 6단계(암호·압축·문서 속성) 브라우저 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-pdf-security-e2e.mjs [BASE] [스크린샷 폴더]
// 암호 걸린 한글 PDF: 비밀번호 창(틀림·맞음), 다른 참여자도 비밀번호 필요, 편집 후 저장 파일은 같은 비밀번호로 잠김, 암호 풀기·권한만 걸기.
// 큰 그림 PDF: 압축 미리 보기·적용 → 내려받은 파일이 작아짐. 문서 속성: 제목 고치기·글꼴 목록. 읽기 전용: 보호·압축 없음.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { randomFillSync } from 'node:crypto';
import fontkit from '@pdf-lib/fontkit';
import sharp from 'sharp';
import { PDFDocument as Cantoo } from '@cantoo/pdf-lib';
import { PDFDocument, rgb } from 'pdf-lib';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-security-'));
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

async function koreanPdf(title, withImage = false) {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const font = await doc.embedFont(fs.readFileSync('assets/fonts/Pretendard-Regular.ttf'), { subset: false });
  const p = doc.addPage([595, 842]);
  p.drawText(title, { x: 60, y: 760, size: 24, font, color: rgb(0.1, 0.2, 0.4) });
  p.drawText('대외비 — 열람 주의', { x: 60, y: 720, size: 14, font });
  if (withImage) {
    const noise = Buffer.alloc(1600 * 1200 * 3);
    randomFillSync(noise);
    const png = await sharp(noise, { raw: { width: 1600, height: 1200, channels: 3 } }).png().toBuffer();
    p.drawImage(await doc.embedPng(png), { x: 60, y: 300, width: 300, height: 225 });
  }
  return doc.save();
}
async function lock(bytes, password) {
  const d = await Cantoo.load(bytes);
  d.encrypt({ userPassword: password, ownerPassword: `${password}!owner`, permissions: { printing: 'highResolution', copying: true, modifying: true } });
  return d.save();
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
const download = async (tok) => new Uint8Array(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(tok)}`)).arrayBuffer());
async function openWith(bytes, password) {
  const pdf = await getDocument({ data: new Uint8Array(bytes), password, isEvalSupported: false }).promise;
  const text = (await (await pdf.getPage(1)).getTextContent()).items.map((i) => i.str).join('');
  const annots = (await (await pdf.getPage(1)).getAnnotations()).map((a) => a.subtype);
  const perms = await pdf.getPermissions();
  const info = (await pdf.getMetadata()).info;
  await pdf.destroy();
  return { text, annots, perms, info };
}

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1500, height: 950 } });
const errors = [];
const watch = (p) => {
  p.on('pageerror', (e) => errors.push(String(e)));
  p.on('console', (m) => {
    if (m.type() === 'error' && !/423|403/.test(m.text())) errors.push(m.text());
  });
};

// ---------- 암호 파일 ----------
const lockedId = await upload(await lock(await koreanPdf('비밀 계약서'), 'open1234'), '비밀계약서.pdf');
const t1 = await token(lockedId);
const page = await ctx.newPage();
watch(page);
await page.goto(`${BASE}${t1.data.editorUrl}`);
await page.waitForSelector('#pw-dialog:not([hidden])', { timeout: 30000 });
check(true, '암호 파일 — 비밀번호 창');
await page.fill('#pw-input', 'wrong');
await page.click('#pw-ok');
await page.waitForSelector('#pw-error:not([hidden])');
check((await page.textContent('#pw-error')).includes('맞지 않'), '틀린 비밀번호 안내', await page.textContent('#pw-error'));
await page.fill('#pw-input', 'open1234');
await page.press('#pw-input', 'Enter');
await page.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await page.waitForFunction(() => (document.querySelector('.pdf-page[data-page="1"] .pdf-text-layer')?.textContent ?? '').length > 0, null, { timeout: 15000 });
const text1 = await page.locator('.pdf-page[data-page="1"] .pdf-text-layer').textContent();
check(text1.includes('비밀 계약서') && (await page.isHidden('#pw-dialog')), '맞히면 같은 화면에서 열림(한글)', text1.slice(0, 40));

// 다른 참여자도 비밀번호가 있어야
const other = await ctx.newPage();
watch(other);
await other.goto(`${BASE}${(await token(lockedId, { viewer: '다른 사람' })).data.editorUrl}`);
await other.waitForSelector('#pw-dialog:not([hidden])', { timeout: 30000 });
const srcStatus = (await fetch(`${BASE}/api/view/pdf-source?token=${encodeURIComponent((await token(lockedId, { viewer: '셋째' })).data.token)}`)).status;
check(srcStatus === 423, '다른 참여자(비밀번호 안 넣음)는 원본을 못 받음', String(srcStatus));
await other.fill('#pw-input', 'open1234');
await other.press('#pw-input', 'Enter');
await other.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
check(true, '다른 참여자도 비밀번호로 열림');
await other.close();

// 편집 → 저장 파일은 같은 비밀번호로 잠김
await page.click('.pdf-mode[data-mode="edit"]');
await page.click('[data-group="shape"] .pdf-float-more');
await page.click('[data-group="shape"] [data-pick="rect"]');
const pb = await page.locator('.pdf-page[data-page="1"]').boundingBox();
await page.mouse.move(pb.x + 80, pb.y + 300);
await page.mouse.down();
await page.mouse.move(pb.x + 260, pb.y + 380, { steps: 4 });
await page.mouse.up();
await page.waitForTimeout(400);
const saved1 = await download(t1.data.token);
let noPw = 'opened';
try {
  await openWith(saved1);
} catch (e) {
  noPw = e.name;
}
const withPw = await openWith(saved1, 'open1234');
check(noPw === 'PasswordException' && withPw.text.includes('비밀 계약서') && withPw.annots.includes('Square'), '저장 파일: 같은 비밀번호로 잠김(주석 포함)', `${noPw} / ${withPw.annots.join(',')}`);

// 암호·권한 창 — 올린 파일에 암호가 있다는 표시, 암호 풀기
await page.click('.pdf-mode[data-mode="tools"]');
await page.click('.pdf-tool-item[data-tool="protect"]');
await page.waitForSelector('#protect-dialog:not([hidden])');
check((await page.textContent('#protect-status')).includes('올린 파일에 암호') && (await page.isVisible('#protect-remove')), '보호 창: 지금 상태 표시', await page.textContent('#protect-status'));
await page.click('#protect-remove');
await page.waitForTimeout(400);
const saved2 = await download(t1.data.token);
const plain = await openWith(saved2);
check(plain.text.includes('비밀 계약서') && plain.annots.includes('Square'), '암호 풀기 → 비밀번호 없이 열림');

// 권한만(권한 비밀번호, 복사 금지)
await page.click('.pdf-tool-item[data-tool="protect"]');
await page.click('#protect-apply');
check((await page.textContent('#protect-error')).includes('하나는'), '비밀번호 없이 걸기 → 안내');
await page.fill('#protect-owner', 'owner-pw');
await page.locator('#protect-perms input[data-perm="copying"]').uncheck();
await page.click('#protect-apply');
await page.waitForTimeout(400);
const saved3 = await download(t1.data.token);
const permOnly = await openWith(saved3);
check(permOnly.text.includes('비밀 계약서') && Array.isArray(permOnly.perms) && !permOnly.perms.includes(16) && permOnly.perms.includes(4), '권한만 걸기(열기는 자유, 복사 금지)', JSON.stringify(permOnly.perms));

// 문서 속성 — 제목 고치기, 글꼴 목록
await page.click('#info-btn');
await page.waitForSelector('#props-dialog:not([hidden])');
await page.waitForSelector('.pdf-font-list li');
const fontsText = await page.textContent('.pdf-font-list');
check(/Pretendard/.test(fontsText) && fontsText.includes('포함'), '문서 속성: 글꼴 목록(포함 여부)', fontsText.slice(0, 60));
await page.fill('#props-fields input[data-prop="title"]', '수정된 계약서 제목');
await page.press('#props-fields input[data-prop="title"]', 'Tab');
await page.waitForTimeout(400);
await page.screenshot({ path: path.join(shotDir, '1-props.png') });
await page.click('#props-dialog [data-close]');
const saved4 = await download(t1.data.token);
check((await openWith(saved4)).info.Title === '수정된 계약서 제목', '저장 파일 제목(한글)');
const changes = (await (await fetch(`${BASE}/api/documents/${lockedId}`)).json()).data.changes ?? [];
check(['protection_changed', 'props_changed'].every((k) => changes.some((c) => c.kind === k)), '변경 기록(보호·속성)', changes.map((c) => c.kind).join(','));
await page.close();

// ---------- 압축 ----------
const bigBytes = await koreanPdf('그림 많은 보고서', true);
const bigId = await upload(bigBytes, '큰보고서.pdf');
const t2 = await token(bigId);
const p2 = await ctx.newPage();
watch(p2);
await p2.goto(`${BASE}${t2.data.editorUrl}`);
await p2.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await p2.click('.pdf-tool-item[data-tool="compress"]');
await p2.waitForSelector('#compress-dialog:not([hidden])');
await p2.click('#compress-preview');
await p2.waitForFunction(() => /→/.test(document.getElementById('compress-result')?.textContent ?? ''), null, { timeout: 30000 });
const previewText = await p2.textContent('#compress-result');
check(/줄어듦, 그림 1개/.test(previewText), '압축 미리 보기', previewText);
await p2.screenshot({ path: path.join(shotDir, '2-compress.png') });
await p2.click('#compress-apply');
await p2.waitForTimeout(500);
const compressed = await download(t2.data.token);
check(compressed.length < bigBytes.length / 4 && (await openWith(compressed)).text.includes('그림 많은 보고서'), '압축 적용 → 내려받은 파일이 작아짐(글자 그대로)', `${bigBytes.length} → ${compressed.length}`);
await p2.close();

// ---------- 읽기 전용 ----------
const ro = await ctx.newPage();
watch(ro);
await ro.goto(`${BASE}${(await token(bigId, { viewer: '읽기만', canEdit: false })).data.editorUrl}`);
await ro.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
const roTools = await ro.$$eval('.pdf-tool-item', (els) => els.filter((e) => e.offsetParent !== null).map((e) => e.dataset.tool));
await ro.click('#info-btn');
await ro.waitForSelector('#props-dialog:not([hidden])');
check(!roTools.includes('protect') && !roTools.includes('compress') && roTools.includes('props') && (await ro.isDisabled('#props-fields input[data-prop="title"]')), '읽기 전용: 보호·압축 없음, 속성은 보기만', roTools.join(','));
const roProtect = await fetch(`${BASE}/api/view/pdf-protect?token=${encodeURIComponent((await token(bigId, { canEdit: false })).data.token)}`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ mode: 'remove' }) });
check(roProtect.status === 403, '읽기 전용 토큰으로 보호 바꾸기 거부', String(roProtect.status));
await ro.close();

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 3).join(' | '));
await browser.close();
for (const id of [lockedId, bigId]) await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' });
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
