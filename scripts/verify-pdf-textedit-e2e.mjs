// PDF 편집기 5단계(기존 글자 고치기·가리기) 브라우저 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-pdf-textedit-e2e.mjs [BASE] [스크린샷 폴더]
// 한글 PDF(글·그림·링크)로: 글 조각 눌러 고치기(원래 글자 재배열 → 원래 글꼴, 새 한글 → Pretendard), 원래대로·실행 취소,
// 가리기 표시 → 적용(확인 창, 되돌릴 수 없음) → 검은 상자·글자 숨김, 저장 파일(글자 실제 삭제·그림 검게·링크 제거),
// 나중에 들어온 참여자에게 가린 글이 가지 않는지, 다시 열기, 변경 기록, 읽기 전용을 확인한다.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import fontkit from '@pdf-lib/fontkit';
import sharp from 'sharp';
import { PDFDocument, PDFName, PDFRawStream, PDFString, rgb } from 'pdf-lib';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-textedit-'));
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

async function fixture() {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const font = await doc.embedFont(fs.readFileSync('assets/fonts/Pretendard-Regular.ttf'), { subset: false });
  const p = doc.addPage([595, 842]);
  p.drawText('견적서', { x: 60, y: 770, size: 26, font, color: rgb(0.1, 0.2, 0.4) });
  p.drawText('합계 1234원', { x: 60, y: 700, size: 16, font });
  p.drawText('담당 홍길동', { x: 60, y: 660, size: 16, font });
  p.drawText('주민번호 900101-1234567 확인', { x: 60, y: 600, size: 16, font });
  const img = await doc.embedPng(await sharp({ create: { width: 120, height: 60, channels: 3, background: '#3a7bd5' } }).png().toBuffer());
  p.drawImage(img, { x: 60, y: 480, width: 240, height: 100 });
  const ctx = doc.context;
  p.node.addAnnot(ctx.register(ctx.obj({ Type: 'Annot', Subtype: 'Link', Rect: [60, 480, 160, 520], A: ctx.obj({ S: 'URI', URI: PDFString.of('https://example.com/secret') }) })));
  return doc.save();
}

async function upload(bytes, name) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
  const { data } = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  for (let i = 0; i < 200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${data.id}`)).json()).data;
    if (st.status === 'ready' || st.editReady) return data.id;
    await new Promise((r) => setTimeout(r, 150));
  }
  throw new Error('대기 시간 초과');
}
const token = (id, body = {}) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '시험 편집자', canPrint: true, canDownload: true, canCopy: true, canEdit: true, pdfEditor: true, ...body }),
  }).then((r) => r.json());

async function describeSaved(bytes) {
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise;
  const page = await pdf.getPage(1);
  const text = (await page.getTextContent()).items.map((i) => i.str).join('|');
  const links = (await page.getAnnotations()).filter((a) => a.subtype === 'Link').length;
  await pdf.destroy();
  return { text, links };
}

const id = await upload(await fixture(), '견적서.pdf');
const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1500, height: 950 } });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => {
  if (m.type() === 'error') errors.push(m.text());
});
const t = await token(id);
await page.goto(`${BASE}${t.data.editorUrl}`);
await page.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await page.waitForTimeout(1000);
const span = (text) => page.locator('.pdf-page[data-page="1"] .pdf-text-layer span', { hasText: text }).first();

// 1) 편집 탭 → 글자 고치기 → 글 조각 누르기 → 입력란
await page.click('.pdf-mode[data-mode="edit"]');
await page.click('#annot-tools [data-tool="textedit"]');
check((await page.getAttribute('#pdf-scroll', 'data-tool')) === 'textedit', '글자 고치기 도구');
await span('합계').click();
const input = page.locator('.pdf-edit-input input');
await input.waitFor();
check((await input.inputValue()) === '합계 1234원', '누른 글 조각이 입력란에', await input.inputValue());
await input.fill('합계 4321원');
await input.press('Enter');
await page.waitForTimeout(400);
const shown = await page.locator('.pdf-page[data-page="1"] .pdf-edit-text').first().textContent();
check(shown === '합계 4321원' && (await span('합계').evaluate((e) => e.classList.contains('pdf-text-hidden'))), '고친 글 미리 보기(원래 글은 숨김)', shown);

// 2) 다른 조각 — 새 한글
await span('담당').click();
await input.waitFor();
await input.fill('담당 김미정');
await input.press('Enter');
await page.waitForTimeout(400);
check((await page.$$('.pdf-page[data-page="1"] .pdf-edit-text')).length === 2, '두 번째 고치기');

// 3) 실행 취소·다시 실행
await page.locator('#pdf-scroll').click({ position: { x: 5, y: 5 } });
await page.keyboard.press('Control+z');
await page.waitForTimeout(300);
const afterUndo = (await page.$$('.pdf-page[data-page="1"] .pdf-edit-text')).length;
await page.keyboard.press('Control+y');
await page.waitForTimeout(300);
check(afterUndo === 1 && (await page.$$('.pdf-page[data-page="1"] .pdf-edit-text')).length === 2, '고치기도 Ctrl+Z / Ctrl+Y', `${afterUndo} → 2`);

// 4) 원래대로 — 고친 글을 눌러 다시 열고 되돌렸다가 다시 고침
await page.locator('.pdf-page[data-page="1"] .pdf-edit-text', { hasText: '김미정' }).click();
await input.waitFor();
check((await input.inputValue()) === '담당 김미정' && (await page.isVisible('.pdf-edit-actions button:has-text("원래대로")')), '고친 글을 누르면 다시 고치기(원래대로 단추)');
await page.click('.pdf-edit-actions button:has-text("원래대로")');
await page.waitForTimeout(300);
const afterRevert = (await page.$$('.pdf-page[data-page="1"] .pdf-edit-text')).length;
await span('담당').click();
await input.waitFor();
await input.fill('담당 김미정');
await input.press('Enter');
await page.waitForTimeout(300);
check(afterRevert === 1 && (await page.$$('.pdf-page[data-page="1"] .pdf-edit-text')).length === 2, '원래대로 → 다시 고치기');

// 5) 가리기 — 주민번호 부분 + 그림 왼쪽 절반(링크 포함) 표시
await page.click('#annot-tools [data-tool="redact"]');
const target = await span('주민번호').boundingBox();
const pb = await page.locator('.pdf-page[data-page="1"]').boundingBox();
const scale = pb.width / 595;
const startX = target.x + 16 * 4.2 * scale; // "주민번호 " 뒤부터
await page.mouse.move(startX, target.y - 2);
await page.mouse.down();
await page.mouse.move(startX + 16 * 7.5 * scale, target.y + target.height + 2, { steps: 5 });
await page.mouse.up();
await page.waitForTimeout(200);
const imgTop = pb.y + (842 - 580) * scale;
await page.mouse.move(pb.x + 55 * scale, imgTop - 3);
await page.mouse.down();
await page.mouse.move(pb.x + 180 * scale, imgTop + 105 * scale, { steps: 5 });
await page.mouse.up();
await page.waitForTimeout(300);
check((await page.$$('.pdf-page[data-page="1"] .pdf-redact:not(.applied)')).length === 2 && (await page.textContent('#redact-count')) === '2' && (await page.isVisible('#redact-apply')), '가리기 표시 2곳·적용 단추');

// 6) 적용 — 확인 창 취소 → 그대로, 적용 → 검은 상자·글자 숨김·되돌릴 수 없음
await page.click('#redact-apply');
await page.waitForSelector('#redact-dialog:not([hidden])');
check((await page.textContent('#redact-dialog-text')).includes('2곳'), '확인 창', await page.textContent('#redact-dialog-text'));
await page.click('#redact-cancel');
check((await page.$$('.pdf-redact.applied')).length === 0, '취소하면 적용 안 함');
await page.click('#redact-apply');
await page.click('#redact-confirm');
await page.waitForTimeout(600);
const appliedCount = (await page.$$('.pdf-page[data-page="1"] .pdf-redact.applied')).length;
const idHidden = await span('주민번호').evaluate((e) => e.classList.contains('pdf-text-hidden'));
await page.locator('#pdf-scroll').click({ position: { x: 5, y: 5 } });
await page.keyboard.press('Control+z');
await page.waitForTimeout(300);
check(appliedCount === 2 && idHidden && (await page.$$('.pdf-page[data-page="1"] .pdf-redact.applied')).length === 2 && (await page.isHidden('#redact-apply')), '적용: 검은 상자 2개·글자 숨김·Ctrl+Z로도 안 돌아감');
await page.screenshot({ path: path.join(shotDir, '1-edited.png') });

// 7) 저장 파일 — 고친 글, 가린 글자는 실제로 없음, 그림 검게, 링크 제거
const dl = new Uint8Array(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(t.data.token)}`)).arrayBuffer());
const saved = await describeSaved(dl);
check(saved.text.includes('합계 4321원') && saved.text.includes('담당 김미정') && !saved.text.includes('1234원') && !saved.text.includes('홍길동'), '저장 파일: 고친 글', saved.text);
check(saved.text.includes('주민번호') && saved.text.includes('확인') && !/900101|1234567/.test(saved.text), '저장 파일: 가린 글자는 실제로 지워짐', saved.text);
check(saved.links === 0, '저장 파일: 가린 곳의 링크 제거');
const savedDoc = await PDFDocument.load(dl);
const images = savedDoc.context
  .enumerateIndirectObjects()
  .map(([, o]) => o)
  .filter((o) => o instanceof PDFRawStream && o.dict.get(PDFName.of('Subtype')) === PDFName.of('Image'));
check(images.length >= 1, '저장 파일: 그림은 남음(가린 부분만 검게)', String(images.length));
fs.writeFileSync(path.join(shotDir, 'saved.pdf'), dl);

// 8) 나중에 들어온 참여자 — 가린 글자가 화면 원본에도 없다
const late = await ctx.newPage();
await late.goto(`${BASE}${(await token(id, { viewer: '나중' })).data.editorUrl}`);
await late.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await late.waitForTimeout(1500);
const lateText = await late.locator('.pdf-page[data-page="1"] .pdf-text-layer').textContent();
check(!/900101|1234567/.test(lateText) && (await late.$$('.pdf-page[data-page="1"] .pdf-edit-text')).length === 2, '나중에 들어온 참여자: 가린 글자가 원본에도 없음·고친 글 보임', lateText.slice(0, 80));
await late.close();

// 9) 변경 기록
const changes = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data.changes ?? [];
check(changes.filter((c) => c.kind === 'text_edited').length === 2 && changes.filter((c) => c.kind === 'redaction_applied').length === 2, '변경 기록(고치기 2·가리기 2)', changes.map((c) => c.kind).join(','));

// 10) 저장 뒤 다시 열기 — 고친 글이 파일에 들어가 있다
await page.goto('about:blank');
await new Promise((r) => setTimeout(r, 2500));
const page2 = await ctx.newPage();
await page2.goto(`${BASE}${(await token(id)).data.editorUrl}`);
await page2.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await page2.waitForTimeout(1200);
const reText = await page2.locator('.pdf-page[data-page="1"] .pdf-text-layer').textContent();
check(reText.includes('4321') && reText.includes('김미정') && !/900101/.test(reText) && (await page2.$$('.pdf-edit-text')).length === 0, '다시 열면 고친 글이 원본에(덮개 없음)', reText.slice(0, 80));
await page2.screenshot({ path: path.join(shotDir, '2-reopened.png') });
await page2.close();

// 11) 읽기 전용 — 도구 없음
const ro = await ctx.newPage();
await ro.goto(`${BASE}${(await token(id, { viewer: '읽기만', canEdit: false })).data.editorUrl}`);
await ro.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
check(!(await ro.isVisible('[data-tool="textedit"]')) && !(await ro.isVisible('[data-tool="redact"]')), '읽기 전용: 고치기·가리기 도구 없음');
await ro.close();

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 3).join(' | '));
await browser.close();
await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' });
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
