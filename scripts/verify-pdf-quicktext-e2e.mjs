// PDF 편집기 글 쓰기 — 누른 곳에 바로 입력(마우스를 대면 자리 미리 보기, 글에 맞춰 자람, 도구 유지·Esc 로 끝),
// 끌어 그린 상자(폭 안에서 줄 바꿈), 쓴 글 다시 누르면 고치기, 다음 글 크기, 빈 글 지우기, 저장된 PDF 의 글 상자.
// 사용: node scripts/verify-pdf-quicktext-e2e.mjs [http://127.0.0.1:8124]
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-quicktext-'));
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

async function fixture() {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const page = doc.addPage([595, 842]);
  page.drawText('Name: ____________________', { x: 60, y: 760, size: 14, font });
  page.drawText('Date: ____________________', { x: 60, y: 720, size: 14, font });
  return doc.save();
}
const token = (id, body = {}) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '글쓰기 시험', canPrint: true, canDownload: true, canCopy: true, canEdit: true, pdfEditor: true, ...body }),
  }).then((r) => r.json());

const form = new FormData();
form.append('file', new Blob([await fixture()]), '서식.pdf');
const { data } = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
for (let i = 0; i < 200; i += 1) {
  const st = (await (await fetch(`${BASE}/api/documents/${data.id}`)).json()).data;
  if (st.status === 'ready' || st.editReady) break;
  await new Promise((r) => setTimeout(r, 150));
}
const t = await token(data.id);
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1400, height: 950 } });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
await page.goto(BASE + t.data.editorUrl);
await page.waitForSelector('.pdf-page[data-page="1"] canvas', { timeout: 60000 });
await page.waitForTimeout(1200);
let box = await page.locator('.pdf-page[data-page="1"]').boundingBox();
const at = (fx, fy) => ({ x: box.x + box.width * fx, y: box.y + box.height * fy });
const annots = () => page.evaluate(() => [...document.querySelectorAll('.pdf-page .pdf-annot-layer .pdf-freetext')].map((d) => ({ text: d.textContent, w: d.offsetWidth, h: d.offsetHeight })));
const pressed = () => page.getAttribute('#annot-tools [data-tool="freetext"]', 'aria-pressed');
const hintShown = () => page.evaluate(() => !document.querySelector('.pdf-tool-hint')?.hidden);

// 1) 글 도구 고르기 → 안내 문구
await page.click('#annot-tools [data-tool="freetext"]');
check(await hintShown(), '글 도구를 고르면 아래쪽에 쓰는 법 안내');

// 2) 마우스를 대면 자리 미리 보기
const p1 = at(0.36, 0.088); // Name 밑줄 위
await page.mouse.move(p1.x, p1.y);
await page.waitForTimeout(150);
const ghost = await page.evaluate(() => {
  const g = document.querySelector('.pdf-text-ghost');
  const r = g?.getBoundingClientRect();
  return g && !g.hidden ? { left: r.left, anchor: r.top + r.height * 0.78, cursor: getComputedStyle(document.querySelector('.pdf-annot-layer')).cursor } : null;
});
await page.screenshot({ path: path.join(shotDir, 'hover.png') });
check(!!ghost && Math.abs(ghost.left - p1.x) < 2 && Math.abs(ghost.anchor - p1.y) < 2 && ghost.cursor === 'text', '마우스를 대면 그 자리에 입력 막대 미리 보기·글 입력 커서', JSON.stringify(ghost));

// 3) 누르면 바로 입력 — 글에 맞춰 자람
await page.mouse.click(p1.x, p1.y);
await page.waitForSelector('.pdf-freetext-editor.fit');
const w0 = await page.evaluate(() => document.querySelector('.pdf-freetext-editor').offsetWidth);
await page.keyboard.insertText('홍길동');
const w1 = await page.evaluate(() => document.querySelector('.pdf-freetext-editor').offsetWidth);
await page.keyboard.insertText(' (영업팀 과장)');
const e2 = await page.evaluate(() => ({ w: document.querySelector('.pdf-freetext-editor').offsetWidth, h: document.querySelector('.pdf-freetext-editor').offsetHeight }));
await page.keyboard.press('Enter');
await page.keyboard.insertText('두 번째 줄');
const h2 = await page.evaluate(() => document.querySelector('.pdf-freetext-editor').offsetHeight);
check(w1 > w0 && e2.w > w1 && h2 > e2.h * 1.6, '누른 곳에서 바로 입력 — 칠수록 가로로, Enter 면 세로로 자람', `${w0}→${w1}→${e2.w}, h ${e2.h}→${h2}`);

// 4) 다른 곳을 누르면 앞 글을 마치고 거기서 새 글 — 글 도구 유지
const p2 = at(0.36, 0.135); // Date 밑줄 위
await page.mouse.click(p2.x, p2.y);
await page.waitForSelector('.pdf-freetext-editor.fit');
await page.keyboard.insertText('2026-11-05');
let list = await annots();
check(list.length === 2 && list[0].text === '홍길동 (영업팀 과장)\n두 번째 줄' && list[1].text === '' && (await pressed()) === 'true', '다른 곳을 누르면 앞 글이 남고 새 글을 바로 시작(글 도구 유지)', JSON.stringify(list));

// 5) Esc → 그 글만 마침(도구 유지), 한 번 더 Esc → 도구 끝
await page.keyboard.press('Escape');
await page.waitForTimeout(250);
list = await annots();
check(list.length === 2 && (await pressed()) === 'true' && (await hintShown()), '첫 Esc 는 쓰던 글만 마침 — 글 도구·안내는 그대로', JSON.stringify(list.map((a) => a.text)));
// 바로 쓴 글 상자는 글에 꼭 맞고 줄이 바뀌지 않는다
const one = list.find((a) => a.text === '2026-11-05');
const lineH = await page.evaluate(() => parseFloat(getComputedStyle(document.querySelector('.pdf-page .pdf-annot-layer .pdf-freetext')).fontSize) * 1.2 + 4);
check(!!one && one.h <= lineH + 2, '바로 쓴 한 줄 글은 한 줄 높이 그대로(줄 바꿈 없음)', `${one?.h} <= ${lineH}`);
await page.keyboard.press('Escape');
await page.waitForTimeout(200);
check((await pressed()) === 'false' && !(await hintShown()) && (await page.getAttribute('#annot-select', 'aria-pressed').catch(() => null)) !== 'false', '두 번째 Esc 로 글 도구 끝 — 안내도 사라짐');

// 6) 다음 글 크기 — 글 도구에서 글자 크기를 바꾸면 새 글에 적용
await page.click('#annot-tools [data-tool="freetext"]');
await page.click('#annot-style-btn');
await page.locator('#annot-fontsize').fill('24');
await page.locator('#annot-fontsize').dispatchEvent('change');
await page.click('#annot-style-btn');
const p3 = at(0.2, 0.3);
await page.mouse.click(p3.x, p3.y);
await page.waitForSelector('.pdf-freetext-editor.fit');
await page.keyboard.insertText('큰 글씨');
await page.keyboard.press('Escape');
await page.waitForTimeout(250);
const sizes = await page.evaluate(() => [...document.querySelectorAll('.pdf-page .pdf-annot-layer .pdf-freetext')].map((d) => parseFloat(d.style.fontSize)));
check(sizes.length === 3 && Math.max(...sizes) > Math.min(...sizes) * 1.6, '글자 크기를 24pt 로 바꾸면 다음 글이 크게', JSON.stringify(sizes));

// 7) 쓴 글을 글 도구로 다시 누르면 새로 만들지 않고 고친다
const target = page.locator('.pdf-page .pdf-annot-layer .pdf-freetext', { hasText: '2026-11-05' });
await target.click();
await page.waitForSelector('.pdf-freetext-editor');
await page.keyboard.press('End');
await page.keyboard.insertText(' 오전');
await page.keyboard.press('Escape');
await page.waitForTimeout(250);
list = await annots();
const edited = list.find((a) => a.text.startsWith('2026-11-05'));
check(list.length === 3 && edited?.text === '2026-11-05 오전' && edited.h <= lineH + 2, '쓴 글을 누르면 고치기 — 늘어난 글에 맞춰 상자도 넓어짐(줄 바꿈 없음)', JSON.stringify(edited));

// 8) 아무것도 안 치고 다른 곳을 누르면 빈 글은 남지 않는다
await page.mouse.click(at(0.5, 0.5).x, at(0.5, 0.5).y);
await page.waitForSelector('.pdf-freetext-editor.fit');
await page.mouse.click(at(0.5, 0.6).x, at(0.5, 0.6).y);
await page.waitForTimeout(200);
await page.keyboard.press('Escape');
await page.waitForTimeout(250);
check((await annots()).length === 3, '빈 글은 지워짐(누르기만 하고 안 친 곳)', String((await annots()).length));

// 9) 끌어 그린 상자 — 그 폭 안에서 줄 바꿈, 넘치면 아래로 자람
const d1 = at(0.15, 0.45);
const d2 = at(0.45, 0.48);
await page.mouse.move(d1.x, d1.y);
await page.mouse.down();
await page.mouse.move(d2.x, d2.y, { steps: 6 });
await page.mouse.up();
await page.waitForSelector('.pdf-freetext-editor:not(.fit)');
const drawnW = await page.evaluate(() => document.querySelector('.pdf-freetext-editor').offsetWidth);
await page.keyboard.insertText('끌어서 그린 상자에는 긴 글을 넣으면 상자 폭 안에서 줄이 바뀌고 아래로 자랍니다.');
await page.keyboard.press('Escape');
await page.waitForTimeout(250);
const drawn = (await annots()).find((a) => a.text.startsWith('끌어서'));
check(!!drawn && Math.abs(drawn.w - drawnW) <= 2 && drawn.h > lineH * 2, '끌어 그린 상자 — 폭은 그대로, 긴 글은 줄을 바꿔 아래로 자람', JSON.stringify({ drawnW, w: drawn?.w, h: drawn?.h }));
await page.keyboard.press('Escape');
await page.screenshot({ path: path.join(shotDir, 'quicktext.png') });

// 10) 저장된 PDF — 글 상자 주석으로 남고, 바로 쓴 한 줄 글은 PDF 에서도 한 줄
await page.waitForTimeout(9000); // 자동 저장(8초)
const t2 = await token(data.id, { canEdit: false });
const bytes = new Uint8Array(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(t2.data.token)}`)).arrayBuffer());
const saved = await PDFDocument.load(bytes);
const pdf = await getDocument({ data: bytes.slice(), isEvalSupported: false }).promise;
const free = (await (await pdf.getPage(1)).getAnnotations()).filter((a) => a.subtype === 'FreeText');
await pdf.destroy();
// 저장된 외형(AP) 글줄 수 — 한 줄짜리가 서버 글꼴 폭 때문에 두 줄로 나뉘지 않았는지
const lines = await (async () => {
  const { PDFName, PDFRawStream, decodePDFRawStream } = await import('pdf-lib');
  const out = {};
  for (const ref of saved.getPage(0).node.Annots()?.asArray() ?? []) {
    const dict = saved.context.lookup(ref);
    if (dict.get(PDFName.of('Subtype'))?.toString() !== '/FreeText') continue;
    const contents = dict.get(PDFName.of('Contents'))?.decodeText?.() ?? '';
    const ap = dict.lookup(PDFName.of('AP'))?.lookup(PDFName.of('N'));
    const raw = ap instanceof PDFRawStream ? Buffer.from(decodePDFRawStream(ap).decode()).toString('latin1') : '';
    out[contents] = (raw.match(/\bTj\b/g) ?? []).length;
  }
  return out;
})();
check(free.length === 4, '저장된 PDF 에 글 상자 4개(빈 글은 남지 않음)', String(free.length));
check(lines['2026-11-05 오전'] === 1 && lines['큰 글씨'] === 1 && lines['홍길동 (영업팀 과장)\n두 번째 줄'] === 2, '저장된 PDF 에서도 바로 쓴 글은 친 그대로의 줄 수', JSON.stringify(lines));

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 2).join(' | '));
await browser.close();
const ok = results.every(Boolean);
console.log(ok ? `ALL CHECKS PASSED (${results.length}/${results.length}) — 스크린샷: ${shotDir}` : `SOME CHECKS FAILED (${results.filter(Boolean).length}/${results.length})`);
process.exit(ok ? 0 : 1);
