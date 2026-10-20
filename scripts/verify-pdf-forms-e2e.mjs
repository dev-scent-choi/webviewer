// PDF 편집기 4단계(양식 채우기·서명) 브라우저 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-pdf-forms-e2e.mjs [BASE] [스크린샷 폴더]
// 양식(글자·체크·라디오·목록·여러 고르기·읽기 전용·서명 칸)이 있는 한글 PDF로: 본문에 칸이 보이는지, 값 넣기(한글),
// 다른 참여자 화면, 실행 취소, 필드 창, 서명 만들기(그리기·그림·글자)·서명 칸에 놓기·비율 유지 크기, 날짜·도장,
// 내려받은 파일(양식 값·Stamp), 저장 뒤 다시 열기, 읽기 전용, 변경 기록을 확인한다.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import fontkit from '@pdf-lib/fontkit';
import { PDFDocument, PDFString, rgb } from 'pdf-lib';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-forms-'));
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
  const label = (t, y) => p.drawText(t, { x: 60, y, size: 12, font });
  p.drawText('참가 신청서', { x: 60, y: 780, size: 24, font, color: rgb(0.1, 0.2, 0.4) });
  label('이름', 720);
  label('동의', 680);
  label('크기', 640);
  label('지역', 600);
  label('관심', 540);
  label('번호', 480);
  label('서명', 410);
  const form = doc.getForm();
  form.createTextField('name').addToPage(p, { x: 120, y: 712, width: 220, height: 24 });
  form.createCheckBox('agree').addToPage(p, { x: 120, y: 676, width: 16, height: 16 });
  const size = form.createRadioGroup('size');
  size.addOptionToPage('S', p, { x: 120, y: 636, width: 14, height: 14 });
  size.addOptionToPage('L', p, { x: 150, y: 636, width: 14, height: 14 });
  const city = form.createDropdown('city');
  city.addOptions(['Seoul', 'Busan']);
  city.addToPage(p, { x: 120, y: 592, width: 160, height: 22 });
  const tags = form.createOptionList('tags');
  tags.addOptions(['a', 'b', 'c']);
  tags.enableMultiselect();
  tags.addToPage(p, { x: 120, y: 520, width: 120, height: 54 });
  const fixed = form.createTextField('fixed');
  fixed.setText('A-100');
  fixed.enableReadOnly();
  fixed.addToPage(p, { x: 120, y: 472, width: 120, height: 22 });
  const ctx = doc.context;
  const sigRef = ctx.register(ctx.obj({ FT: 'Sig', T: PDFString.of('sign1'), Type: 'Annot', Subtype: 'Widget', Rect: [120, 380, 320, 440], F: 4, P: p.ref }));
  p.node.addAnnot(sigRef);
  form.acroForm.addField(sigRef);
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

/** 서버 쪽 pdf.js로 위젯 값과 Stamp */
async function describeSaved(bytes) {
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise;
  const annots = await (await pdf.getPage(1)).getAnnotations();
  const fields = {};
  for (const a of annots) if (a.fieldName) fields[a.fieldName] = a.radioButton ? (a.fieldValue === a.buttonValue ? a.buttonValue : fields[a.fieldName]) : a.fieldValue;
  const stamps = annots.filter((a) => a.subtype === 'Stamp');
  const freetexts = annots.filter((a) => a.subtype === 'FreeText').map((a) => a.contentsObj?.str ?? '');
  await pdf.destroy();
  return { fields, stamps, freetexts };
}

const id = await upload(await fixture(), '신청서.pdf');
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
const field = (name) => page.locator(`.pdf-page[data-page="1"] .pdf-form-field[data-field="${name}"]`);

// 1) 본문에 양식 칸(예전엔 안 보였다)
const kinds = await page.$$eval('.pdf-page[data-page="1"] .pdf-form-layer > *', (els) => els.map((e) => `${e.dataset.field}:${e.tagName.toLowerCase()}${e.type ? `/${e.type}` : ''}`));
check(kinds.length === 7 && kinds.includes('sign1:button/button') && kinds.includes('tags:select/select-multiple'), '본문에 양식 칸 7개(라디오 둘·서명 칸 포함, 읽기 전용 칸은 그림에)', kinds.join(','));
check((await field('fixed').count()) === 0 && (await page.locator('.pdf-page[data-page="1"] .pdf-text-layer').textContent()).length >= 0, '읽기 전용 칸은 입력란 없이 그림으로만');

// 2) 서명 탭 → 필드 창, 값 넣기
await page.click('.pdf-mode[data-mode="sign"]');
await page.waitForSelector('#panel-fields:not([hidden])');
check((await page.textContent('#fields-count')) === '7' && (await page.isVisible('#sign-tools')) && !(await page.isVisible('#annot-tools')), '서명 탭: 서명 도구 줄·필드 창(7칸)', await page.textContent('#fields-count'));
await field('name').fill('홍길동');
await field('name').press('Tab');
await field('agree').check();
await page.locator('.pdf-page[data-page="1"] .pdf-form-field[data-field="size"]').nth(1).check();
await field('city').selectOption('Busan');
await field('tags').selectOption(['a', 'c']);
await page.waitForTimeout(600);
const panelText = await page.textContent('#fields');
check(panelText.includes('홍길동') && panelText.includes('✓ 체크됨') && panelText.includes('Busan') && panelText.includes('a, c'), '필드 창에 넣은 값', panelText.replace(/\s+/g, ' ').slice(0, 160));

// 3) 다른 참여자 화면에 같은 값
const other = await ctx.newPage();
await other.goto(`${BASE}${(await token(id, { viewer: '다른 사람' })).data.editorUrl}`);
await other.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await other.waitForTimeout(1500);
const ov = await other.evaluate(() => {
  const q = (n) => document.querySelector(`.pdf-page[data-page="1"] .pdf-form-field[data-field="${n}"]`);
  return { name: q('name').value, agree: q('agree').checked, city: q('city').value, L: document.querySelectorAll('.pdf-form-field[data-field="size"]')[1].checked };
});
check(ov.name === '홍길동' && ov.agree && ov.city === 'Busan' && ov.L, '다른 참여자 화면에도 같은 값', JSON.stringify(ov));
// 다른 사람이 바꾸면 내 화면에도
await other.locator('.pdf-form-field[data-field="city"]').selectOption('Seoul');
await page.waitForTimeout(800);
check((await field('city').inputValue()) === 'Seoul', '다른 사람이 바꾼 값이 내 화면에도');
await other.close();

// 4) 실행 취소(내 편집만) — 마지막 내 편집(tags) 되돌리고 다시
await page.locator('#pdf-scroll').click({ position: { x: 5, y: 5 } });
await page.keyboard.press('Control+z');
await page.waitForTimeout(300);
const undoneTags = await field('tags').evaluate((s) => [...s.selectedOptions].map((o) => o.value).join(','));
await page.keyboard.press('Control+y');
await page.waitForTimeout(300);
const redoneTags = await field('tags').evaluate((s) => [...s.selectedOptions].map((o) => o.value).join(','));
check(undoneTags === '' && redoneTags === 'a,c', '양식 값도 Ctrl+Z / Ctrl+Y', `${undoneTags || '(없음)'} → ${redoneTags}`);

// 5) 서명 칸 → 서명 만들기(그리기) → 칸 안에
await page.click('.pdf-page[data-page="1"] .pdf-form-sig');
await page.waitForSelector('#sign-dialog:not([hidden])');
await page.click('#sign-ok');
check(await page.isVisible('#sign-error'), '빈 서명은 안내', await page.textContent('#sign-error'));
const padBox = await page.locator('#sign-pad').boundingBox();
await page.mouse.move(padBox.x + 30, padBox.y + 90);
await page.mouse.down();
for (let i = 0; i < 28; i += 1) await page.mouse.move(padBox.x + 30 + i * 15, padBox.y + 90 + Math.sin(i / 3) * 35);
await page.mouse.up();
await page.click('#sign-ok');
await page.waitForTimeout(500);
const sigBox = await page.locator('.pdf-page[data-page="1"] .pdf-form-sig').boundingBox();
const stamp1 = await page.locator('.pdf-page[data-page="1"] .pdf-annot-layer .pdf-stamp').first().boundingBox();
const inside = stamp1 && stamp1.x >= sigBox.x - 1 && stamp1.y >= sigBox.y - 1 && stamp1.x + stamp1.width <= sigBox.x + sigBox.width + 1 && stamp1.y + stamp1.height <= sigBox.y + sigBox.height + 1;
check(inside && (await page.locator('.pdf-form-sig').getAttribute('class')).includes('signed'), '그린 서명이 서명 칸 안에, 안내 글씨는 감춤', JSON.stringify(stamp1));
check((await page.$$('#sign-list .pdf-sign-chip')).length === 1, '내 서명 목록에 저장');

// 6) 크기 바꾸기 — 비율 유지
await page.click('.pdf-page[data-page="1"] .pdf-annot-layer .pdf-stamp');
const handle = page.locator('.pdf-page[data-page="1"] .pdf-annot-handle');
const hb = await handle.boundingBox();
await page.mouse.move(hb.x + hb.width / 2, hb.y + hb.height / 2);
await page.mouse.down();
await page.mouse.move(hb.x + 120, hb.y + 10, { steps: 5 });
await page.mouse.up();
await page.waitForTimeout(300);
const stamp1b = await page.locator('.pdf-page[data-page="1"] .pdf-annot-layer .pdf-stamp').first().boundingBox();
const ratioA = stamp1.height / stamp1.width;
const ratioB = stamp1b.height / stamp1b.width;
check(stamp1b.width > stamp1.width + 80 && Math.abs(ratioA - ratioB) < 0.03, '서명 크기 바꿔도 비율 유지', `${ratioA.toFixed(3)} → ${ratioB.toFixed(3)}`);

// 7) 그림 올리기 / 글자로 서명 → 눌러 놓기
const pngFile = path.join(shotDir, 'seal.png');
fs.writeFileSync(pngFile, Buffer.from('iVBORw0KGgoAAAANSUhEUgAAABAAAAAICAYAAADwdn+XAAAAJklEQVR4nGP4z8DwnxLMMGrAqAGjBjBQxQBKMCgfGDWASgYMfQUAzq0r8a5x8nkAAAAASUVORK5CYII=', 'base64'));
await page.click('#sign-new');
await page.click('[data-sign-tab="image"]');
await page.setInputFiles('#sign-file', pngFile);
await page.waitForSelector('#sign-preview:not([hidden])');
await page.click('#sign-ok');
const pb = await page.locator('.pdf-page[data-page="1"]').boundingBox();
await page.mouse.click(pb.x + pb.width * 0.62, pb.y + pb.height * 0.62);
await page.waitForTimeout(300);
await page.click('#sign-new');
await page.click('[data-sign-tab="type"]');
await page.fill('#sign-text', '김철수');
await page.waitForTimeout(200);
await page.locator('#sign-styles .pdf-sign-style').nth(1).click();
await page.click('#sign-ok');
await page.mouse.click(pb.x + pb.width * 0.62, pb.y + pb.height * 0.72);
await page.waitForTimeout(300);
check((await page.$$('.pdf-page[data-page="1"] .pdf-annot-layer .pdf-stamp')).length === 3 && (await page.$$('#sign-list .pdf-sign-chip')).length === 3, '그림·글자 서명도 만들어 놓기(목록 3개)');

// 8) 날짜·도장
await page.click('#sign-date');
await page.mouse.click(pb.x + pb.width * 0.62, pb.y + pb.height * 0.5);
await page.waitForTimeout(300);
await page.click('[data-group="seal"] .pdf-float-more');
await page.click('[data-seal="검토"]');
await page.mouse.click(pb.x + pb.width * 0.3, pb.y + pb.height * 0.8);
await page.waitForTimeout(400);
const dateText = await page.locator('.pdf-page[data-page="1"] .pdf-annot-layer .pdf-freetext').first().textContent();
check(/^\d{4}\. \d{2}\. \d{2}\.$/.test(dateText) && (await page.$$('.pdf-page[data-page="1"] .pdf-annot-layer .pdf-stamp')).length === 4, '날짜 글상자·검토 도장', dateText);
await page.screenshot({ path: path.join(shotDir, '1-signed.png') });

// 9) 내려받은 파일 — 양식 값·Stamp 4개·날짜
const dl = new Uint8Array(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(t.data.token)}`)).arrayBuffer());
const saved = await describeSaved(dl);
check(saved.fields.name === '홍길동' && saved.fields.agree === 'Yes' && saved.fields.city?.[0] === 'Seoul' && JSON.stringify(saved.fields.tags) === '["a","c"]' && saved.fields.fixed === 'A-100', '저장 파일 양식 값(한글 포함)', JSON.stringify(saved.fields));
check(saved.stamps.length === 4 && saved.stamps.every((s) => s.hasAppearance) && saved.freetexts.some((x) => /^\d{4}\./.test(x)), '저장 파일에 Stamp 4개(모양 포함)·날짜', `${saved.stamps.length}`);
const savedDoc = await PDFDocument.load(dl);
check(savedDoc.getForm().getRadioGroup('size').getSelected() === 'L', '저장 파일 라디오 값');

// 10) 변경 기록
const changes = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data.changes ?? [];
const formRows = changes.filter((c) => c.kind === 'form_changed').map((c) => c.text);
check(formRows.includes('name: 홍길동') && changes.filter((c) => c.kind === 'annot_added' && c.annotType === 'stamp').length === 4, '변경 기록(양식·서명)', formRows.join(' / '));

// 11) 저장 뒤 다시 열기 — 값은 파일에서, 서명은 다시 고칠 수 있게(편집기 층)
await page.goto('about:blank');
await new Promise((r) => setTimeout(r, 2500));
const page2 = await ctx.newPage();
await page2.goto(`${BASE}${(await token(id)).data.editorUrl}`);
await page2.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await page2.waitForTimeout(1200);
const nameAgain = await page2.locator('.pdf-form-field[data-field="name"]').inputValue();
const stampsAgain = await page2.$$('.pdf-page[data-page="1"] .pdf-annot-layer .pdf-stamp');
check(nameAgain === '홍길동' && stampsAgain.length === 4, '다시 열면 값·서명 그대로(서명은 편집기 층 — 고칠 수 있음)', `${nameAgain} / ${stampsAgain.length}`);
await page2.screenshot({ path: path.join(shotDir, '2-reopened.png') });
await page2.close();

// 12) 읽기 전용 — 입력 불가, 서명 탭 없음
const ro = await ctx.newPage();
await ro.goto(`${BASE}${(await token(id, { viewer: '읽기만', canEdit: false })).data.editorUrl}`);
await ro.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await ro.waitForTimeout(1000);
const roName = ro.locator('.pdf-form-field[data-field="name"]');
check((await roName.isDisabled()) && (await roName.inputValue()) === '홍길동' && !(await ro.isVisible('.pdf-mode[data-mode="sign"]')) && !(await ro.isVisible('.pdf-form-sig')), '읽기 전용: 값만 보이고 입력·서명 불가');
await ro.close();

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 3).join(' | '));
await browser.close();
await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' });
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
