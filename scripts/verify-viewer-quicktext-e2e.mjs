// 8000 docview 보기 화면(viewer.html)의 '텍스트 추가' — 누른 곳에 바로 입력(마우스를 대면 자리 미리 보기),
// 다른 곳을 누르면 앞 글 저장 + 새 글, 첫 Esc 는 쓰던 글 저장(버리지 않음), 두 번째 Esc 는 도구 끝, 도구를 고른 동안 안내.
// 사용: node scripts/verify-viewer-quicktext-e2e.mjs [http://127.0.0.1:8124]
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = fs.mkdtempSync(path.join(os.tmpdir(), 'viewer-quicktext-'));
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

const doc = await PDFDocument.create();
const font = await doc.embedFont(StandardFonts.Helvetica);
const pdfPage = doc.addPage([595, 842]);
pdfPage.drawText('Name: ____________________', { x: 60, y: 760, size: 14, font });
pdfPage.drawText('Date: ____________________', { x: 60, y: 720, size: 14, font });
const form = new FormData();
form.append('file', new Blob([await doc.save()]), '서식.pdf');
const { data } = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
for (let i = 0; i < 200; i += 1) {
  const st = (await (await fetch(`${BASE}/api/documents/${data.id}`)).json()).data;
  if (st.status === 'ready') break;
  await new Promise((r) => setTimeout(r, 200));
}
// 8000 연동처럼 pdfEditor 없이 — PDF 는 보기 화면(마크업)으로 열린다
const t = (await (await fetch(`${BASE}/api/documents/${data.id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: '검토자', canEdit: true, canDownload: true, canPrint: true, canCopy: true }) })).json()).data;
check(!t.editorUrl, 'pdfEditor 없는 토큰은 보기 화면으로(8000 docview 경로)', String(t.editorUrl));

const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1400, height: 950 } });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
await page.goto(BASE + t.viewerUrl);
await page.waitForSelector('.page-frame img, .page-frame canvas', { timeout: 60000 });
await page.waitForTimeout(1500);
const frameBox = await page.locator('.page-frame').first().boundingBox();
const at = (fx, fy) => ({ x: frameBox.x + frameBox.width * fx, y: frameBox.y + frameBox.height * fy });
const items = () => page.evaluate(() => [...document.querySelectorAll('.text-overlay-item')].map((e) => e.textContent.replace(/[−+×]/g, '').trim()));
const state = () =>
  page.evaluate(() => ({
    pressed: document.getElementById('text-tool-btn').getAttribute('aria-pressed'),
    hint: !!document.querySelector('.text-tool-hint') && !document.querySelector('.text-tool-hint').hidden,
    editing: !!document.querySelector('.text-insert-input'),
  }));

// 1) 도구 고르기 → 안내
await page.click('#text-tool-btn');
let s = await state();
check(s.pressed === 'true' && s.hint, '텍스트 추가 도구를 고르면 아래쪽에 쓰는 법 안내');

// 2) 마우스를 대면 자리 미리 보기(누른 점 위에 글이 앉음)
const p1 = at(0.45, 0.095);
await page.mouse.move(p1.x, p1.y);
await page.waitForTimeout(150);
const ghost = await page.evaluate(() => {
  const g = document.querySelector('.text-tool-ghost');
  const r = g?.getBoundingClientRect();
  return g && !g.hidden ? { left: r.left, bottom: r.bottom } : null;
});
check(!!ghost && Math.abs(ghost.left - p1.x) < 2 && Math.abs(ghost.bottom - p1.y) < 2, '마우스를 대면 그 자리에 입력 막대 미리 보기', JSON.stringify(ghost));
await page.screenshot({ path: path.join(shotDir, 'hover.png') });

// 3) 누르면 바로 입력 → 다른 곳을 누르면 앞 글 저장 + 새 글
await page.mouse.click(p1.x, p1.y);
await page.waitForSelector('.text-insert-input');
await page.keyboard.insertText('홍길동');
const p2 = at(0.45, 0.143);
await page.mouse.click(p2.x, p2.y);
await page.waitForSelector('.text-insert-input');
await page.waitForTimeout(500);
check((await items()).includes('홍길동') && (await state()).editing, '다른 곳을 누르면 앞 글이 저장되고 그 자리에서 새 글 시작', JSON.stringify(await items()));

// 4) 첫 Esc — 쓰던 글을 저장(예전엔 버렸다), 도구는 그대로
await page.keyboard.insertText('2026-11-06');
await page.keyboard.press('Escape');
await page.waitForTimeout(600);
s = await state();
check((await items()).includes('2026-11-06') && s.pressed === 'true' && s.hint && !s.editing, '첫 Esc 는 쓰던 글을 저장하고 마침 — 도구·안내는 그대로', JSON.stringify({ items: await items(), ...s }));

// 5) 두 번째 Esc — 도구 끝
await page.keyboard.press('Escape');
await page.waitForTimeout(200);
s = await state();
check(s.pressed === 'false' && !s.hint, '두 번째 Esc 로 도구 끝 — 안내도 사라짐', JSON.stringify(s));

// 6) Enter 로 마침도 그대로
await page.click('#text-tool-btn');
const p3 = at(0.3, 0.3);
await page.mouse.click(p3.x, p3.y);
await page.waitForSelector('.text-insert-input');
await page.keyboard.insertText('검토 완료');
await page.keyboard.press('Enter');
await page.waitForTimeout(600);
check((await items()).includes('검토 완료') && (await state()).pressed === 'true', 'Enter 로 마침 — 도구는 남아 다음 글을 바로 넣을 수 있음');
await page.screenshot({ path: path.join(shotDir, 'after.png') });

// 7) 서버에 저장됨(새로 고쳐도 남음)
const saved = (await (await fetch(`${BASE}/api/view/markup?token=${encodeURIComponent(t.token)}`)).json()).data.filter((m) => m.kind === 'text').map((m) => m.text);
check(['홍길동', '2026-11-06', '검토 완료'].every((x) => saved.includes(x)) && saved.length === 3, '넣은 글 3개가 서버에 저장(빈 글 없음)', JSON.stringify(saved));

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 2).join(' | '));
await browser.close();
const ok = results.every(Boolean);
console.log(ok ? `ALL CHECKS PASSED (${results.length}/${results.length}) — 스크린샷: ${shotDir}` : `SOME CHECKS FAILED (${results.filter(Boolean).length}/${results.length})`);
process.exit(ok ? 0 : 1);
