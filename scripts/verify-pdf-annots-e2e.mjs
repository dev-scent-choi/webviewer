// PDF 편집기 2단계(주석) 브라우저 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-pdf-annots-e2e.mjs [BASE] [스크린샷 폴더]
// 파일에 원래 있던 주석(형광펜·메모+Popup·링크)을 가진 한글 PDF를 올려: 원래 주석이 편집기 층에 보이는지, 각 도구로 만들기,
// 고치기·지우기, 실행 취소, 다른 참여자 화면, 읽기 전용, 그리고 내려받은 파일을 pdf.js(서버 쪽)로 열어 표준 주석인지 확인한다.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import fontkit from '@pdf-lib/fontkit';
import { PDFDocument, PDFHexString, PDFName, PDFString, rgb } from 'pdf-lib';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-annots-'));
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

async function fixture() {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const font = await doc.embedFont(fs.readFileSync('assets/fonts/Pretendard-Regular.ttf'), { subset: false });
  const p1 = doc.addPage([595, 842]);
  doc.addPage([595, 842]);
  p1.drawText('분기 보고서', { x: 60, y: 760, size: 26, font, color: rgb(0.1, 0.25, 0.45) });
  p1.drawText('가동률은 97.2%로 목표를 넘었습니다. 다음 분기 계획을 검토합니다.', { x: 60, y: 700, size: 13, font });
  p1.drawText('세 번째 줄 — 형광펜 시험용 문장입니다.', { x: 60, y: 660, size: 13, font });
  const ctx = doc.context;
  const add = (dict) => {
    const ref = ctx.register(ctx.obj({ Type: 'Annot', ...dict }));
    p1.node.addAnnot(ref);
    return ref;
  };
  add({ Subtype: 'Highlight', Rect: [58, 755, 200, 790], QuadPoints: [58, 790, 200, 790, 58, 755, 200, 755], C: [1, 1, 0], CA: 0.5, NM: PDFString.of('orig-hl'), T: PDFHexString.fromText('원작성자') });
  const note = add({ Subtype: 'Text', Rect: [500, 760, 524, 784], Contents: PDFHexString.fromText('원래 메모'), C: [1, 0.85, 0.3], NM: PDFString.of('orig-note') });
  const popup = add({ Subtype: 'Popup', Rect: [300, 600, 500, 700], Parent: note });
  ctx.lookup(note).set(PDFName.of('Popup'), popup);
  add({ Subtype: 'Link', Rect: [60, 60, 200, 80], A: ctx.obj({ S: 'URI', URI: PDFString.of('https://hyanggilab.com') }) });
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

async function savedAnnots(id) {
  const t = await token(id, { canEdit: false });
  const bytes = new Uint8Array(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(t.data.token)}`)).arrayBuffer());
  const pdf = await getDocument({ data: bytes, isEvalSupported: false }).promise;
  const out = [];
  for (let p = 1; p <= pdf.numPages; p += 1) {
    for (const a of await (await pdf.getPage(p)).getAnnotations()) out.push({ page: p, subtype: a.subtype, hasAppearance: a.hasAppearance, contents: a.contentsObj?.str ?? '', title: a.titleObj?.str ?? '' });
  }
  await pdf.destroy();
  return out;
}

const id = await upload(await fixture(), '주석시험.pdf');
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => {
  if (m.type() === 'error') errors.push(m.text());
});
const t = await token(id);
await page.goto(`${BASE}${t.data.editorUrl}`);
await page.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await page.waitForTimeout(1200);
const p1 = page.locator('.pdf-page[data-page="1"]');
// 누를 때마다 쪽 위치를 다시 잰다(입력란에 초점이 가면 화면이 스크롤될 수 있다)
let box = await p1.boundingBox();
const at = (fx, fy) => ({ x: box.x + box.width * fx, y: box.y + box.height * fy });
const remeasure = async () => {
  box = await p1.boundingBox();
};
const annotCount = () => page.evaluate(() => document.querySelectorAll('.pdf-page[data-page="1"] [data-id]:not(.pdf-annot-selection)').length);

// 1) 원래 주석이 편집기 층에 보이고 canvas에는 두 번 그려지지 않는다
const ids = await page.$$eval('.pdf-page[data-page="1"] [data-id]', (els) => [...new Set(els.map((e) => e.dataset.id))]);
check(ids.includes('orig-hl') && ids.includes('orig-note'), '파일의 형광펜·메모가 편집기 층에', ids.join(','));
check((await page.textContent('#comments-count')) === '2', '메모 목록에 2개');

// 2) 편집 탭 → 도구 줄
await page.click('.pdf-mode[data-mode="edit"]');
check(await page.isVisible('#annot-tools') && await page.isVisible('#panel-comments'), '편집 탭 → 주석 도구와 메모 목록');

// 3) 사각형
await page.click('#annot-tools [data-tool="rect"]');
await page.mouse.move(at(0.55, 0.3).x, at(0.55, 0.3).y);
await page.mouse.down();
await page.mouse.move(at(0.75, 0.38).x, at(0.75, 0.38).y, { steps: 6 });
await page.mouse.up();
await page.waitForTimeout(200);
// 4) 펜
await page.click('#annot-tools [data-tool="ink"]');
await page.mouse.move(at(0.2, 0.5).x, at(0.2, 0.5).y);
await page.mouse.down();
for (let i = 1; i <= 10; i += 1) await page.mouse.move(at(0.2 + i * 0.02, 0.5 + (i % 2) * 0.02).x, at(0.2 + i * 0.02, 0.5 + (i % 2) * 0.02).y);
await page.mouse.up();
// 5) 화살표(도형 묶음 메뉴)
await page.click('[data-group="shape"] .pdf-float-more');
await page.click('[data-group="shape"] [data-pick="arrow"]');
await page.mouse.move(at(0.6, 0.55).x, at(0.6, 0.55).y);
await page.mouse.down();
await page.mouse.move(at(0.8, 0.6).x, at(0.8, 0.6).y, { steps: 5 });
await page.mouse.up();
await page.waitForTimeout(200);
check((await annotCount()) >= 5, '사각형·펜·화살표 만들기', String(await annotCount()));

// 6) 글상자(한글)
await page.click('#annot-tools [data-tool="freetext"]');
await remeasure();
await page.mouse.click(at(0.15, 0.7).x, at(0.15, 0.7).y);
await page.waitForSelector('.pdf-freetext-editor');
await page.keyboard.insertText('한글 글상자입니다');
await page.keyboard.press('Control+Enter');
await page.waitForTimeout(300);
check((await page.textContent('.pdf-freetext')).includes('한글 글상자입니다'), '글상자에 한글');

// 7) 메모(한글)
await page.click('#annot-tools [data-tool="note"]');
await remeasure();
await page.mouse.click(at(0.85, 0.75).x, at(0.85, 0.75).y);
await page.waitForSelector('.pdf-note-popup textarea');
await page.keyboard.insertText('새 메모 — 확인 부탁');
await page.click('.pdf-note-popup .primary');
await page.waitForTimeout(300);
check((await page.textContent('#comments')).includes('새 메모 — 확인 부탁'), '메모 목록에 새 메모 내용');

// 8) 형광펜 — 글자 끌어 고르기
await page.click('[data-group="markup"] [data-tool]');
const line = page.locator('.pdf-page[data-page="1"] .pdf-text-layer span', { hasText: '가동률' }).first();
await line.scrollIntoViewIfNeeded();
const lb = await line.boundingBox();
await page.mouse.move(lb.x + 2, lb.y + lb.height / 2);
await page.mouse.down();
await page.mouse.move(lb.x + lb.width - 2, lb.y + lb.height / 2, { steps: 8 });
await page.mouse.up();
await page.waitForTimeout(400);
const types = await page.$$eval('.pdf-page[data-page="1"] g.pdf-annot', (els) => els.map((e) => e.getAttribute('class')));
check(types.filter((c) => c.includes('pdf-annot-highlight')).length === 2, '글자를 끌어 형광펜(원래 것 + 새 것)', types.join('|'));
await page.screenshot({ path: path.join(shotDir, '1-annots.png') });

// 9) 색 바꾸기(고른 주석) — 선택 도구로 사각형 고르고 파랑
await page.click('#tool-select');
const rectEl = page.locator('.pdf-page[data-page="1"] g.pdf-annot-rect').first();
await rectEl.scrollIntoViewIfNeeded();
const rb = await rectEl.boundingBox();
await page.mouse.click(rb.x + 2, rb.y + rb.height / 2);
await page.waitForTimeout(200);
await page.click('#annot-style-btn');
await page.click('.pdf-style-swatch[data-color="#1A73E8"]');
await page.waitForTimeout(200);
const stroke = await page.$eval('.pdf-page[data-page="1"] g.pdf-annot-rect polygon', (e) => e.getAttribute('stroke'));
check(stroke === '#1A73E8', '고른 사각형 색을 파랑으로', stroke);

// 10) 끌어 옮기기
const before = await rectEl.boundingBox();
await page.mouse.move(before.x + 2, before.y + before.height / 2);
await page.mouse.down();
await page.mouse.move(before.x + 82, before.y + before.height / 2 + 40, { steps: 6 });
await page.mouse.up();
await page.waitForTimeout(300);
const after = await page.locator('.pdf-page[data-page="1"] g.pdf-annot-rect').first().boundingBox();
check(Math.abs(after.x - before.x - 80) < 4 && Math.abs(after.y - before.y - 40) < 4, '사각형 끌어 옮기기', `${Math.round(after.x - before.x)},${Math.round(after.y - before.y)}`);

// 11) 원래 메모 고치기
await page.click('.pdf-note-icon[data-id="orig-note"]');
await page.waitForSelector('.pdf-note-popup textarea');
await page.fill('.pdf-note-popup textarea', '원래 메모 — 고침');
await page.click('.pdf-note-popup .primary');
// 12) 원래 형광펜 지우기(지우개)
await page.click('#annot-tools [data-tool="eraser"]');
const hl = page.locator('.pdf-page[data-page="1"] g[data-id="orig-hl"] polygon').first();
const hb = await hl.boundingBox();
await page.mouse.click(hb.x + hb.width / 2, hb.y + hb.height / 2);
await page.waitForTimeout(300);
check((await page.$$('.pdf-page[data-page="1"] g[data-id="orig-hl"]')).length === 0, '지우개로 원래 형광펜 지우기');

// 13) 실행 취소 → 다시 생김, 다시 실행 → 사라짐
await page.click('#tool-select');
await page.keyboard.press('Control+z');
await page.waitForTimeout(300);
const undone = (await page.$$('.pdf-page[data-page="1"] g[data-id="orig-hl"]')).length === 1;
await page.keyboard.press('Control+y');
await page.waitForTimeout(300);
const redone = (await page.$$('.pdf-page[data-page="1"] g[data-id="orig-hl"]')).length === 0;
check(undone && redone, 'Ctrl+Z로 되살리고 Ctrl+Y로 다시 지움');

// 14) 다른 참여자 화면에 같은 주석
const other = await browser.newPage({ viewport: { width: 1300, height: 900 } });
const t2 = await token(id, { viewer: '다른 사람' });
await other.goto(`${BASE}${t2.data.editorUrl}`);
await other.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await other.waitForTimeout(1200);
const otherIds = await other.$$eval('.pdf-page[data-page="1"] [data-id]', (els) => [...new Set(els.map((e) => e.dataset.id))].length);
const mineIds = await page.$$eval('.pdf-page[data-page="1"] [data-id]', (els) => [...new Set(els.filter((e) => !e.classList.contains('pdf-annot-selection')).map((e) => e.dataset.id))].length);
check(otherIds === mineIds && otherIds >= 6, '다른 참여자 화면에도 같은 주석', `${otherIds}/${mineIds}`);
await other.close();

// 15) 읽기 전용 — 도구 없음, 주석은 보임
const ro = await browser.newPage({ viewport: { width: 1300, height: 900 } });
const t3 = await token(id, { viewer: '읽기만', canEdit: false });
await ro.goto(`${BASE}${t3.data.editorUrl}`);
await ro.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await ro.waitForTimeout(1000);
check(!(await ro.isVisible('#annot-tools')) && (await ro.$$('.pdf-page[data-page="1"] [data-id]')).length > 0, '읽기 전용 — 도구 없이 주석만 보임');
await ro.close();

// 15-1) 보기 회전 중 그린 주석이 회전을 되돌려도 같은 자리(쪽 내용 기준) — 글자 "세 번째 줄" 바로 위에 사각형
const sel3 = page.locator('.pdf-page[data-page="1"] .pdf-text-layer span', { hasText: '세 번째 줄' }).first();
await page.click('#rotate-btn');
await page.waitForTimeout(600);
await sel3.scrollIntoViewIfNeeded();
const r3 = await sel3.boundingBox();
await page.click('#annot-tools [data-tool="rect"], [data-group="shape"] [data-tool]');
await page.click('[data-group="shape"] .pdf-float-more');
await page.click('[data-group="shape"] [data-pick="rect"]');
await page.mouse.move(r3.x - 4, r3.y - 4);
await page.mouse.down();
await page.mouse.move(r3.x + r3.width + 4, r3.y + r3.height + 4, { steps: 5 });
await page.mouse.up();
await page.waitForTimeout(300);
for (let i = 0; i < 3; i += 1) await page.click('#rotate-btn');
await page.waitForTimeout(600);
const textBox = await sel3.boundingBox();
const rects = await page.$$eval('.pdf-page[data-page="1"] g.pdf-annot-rect polygon', (els) => els.map((e) => e.getBoundingClientRect()).map((r) => ({ x: r.x, y: r.y, w: r.width, h: r.height })));
const around = rects.some((r) => r.x <= textBox.x + 2 && r.y <= textBox.y + 2 && r.x + r.w >= textBox.x + textBox.width - 2 && r.y + r.h >= textBox.y + textBox.height - 2);
check(around, '보기 회전 중 그린 사각형이 되돌린 뒤에도 같은 글자를 감싼다', JSON.stringify(rects.map((r) => [Math.round(r.x), Math.round(r.y), Math.round(r.w), Math.round(r.h)])));
await page.click('#tool-select');
const mineAfter = await page.$$eval('.pdf-page[data-page="1"] [data-id]', (els) => [...new Set(els.filter((e) => !e.classList.contains('pdf-annot-selection')).map((e) => e.dataset.id))].length);

// 16) 내려받은 파일 — 표준 주석(모양 포함), 링크 그대로, 한글 내용
const saved = await savedAnnots(id);
const summary = saved.map((a) => `${a.page}:${a.subtype}:${a.hasAppearance ? 'AP' : '-'}`).join(' ');
check(saved.every((a) => a.subtype === 'Link' || a.hasAppearance), '모든 주석에 모양(다른 뷰어에서 보임)', summary);
check(saved.some((a) => a.subtype === 'Link'), '링크는 손대지 않음');
check(!saved.some((a) => a.subtype === 'Popup') && saved.filter((a) => a.subtype === 'Highlight').length === 1, '지운 원래 형광펜·Popup은 빠지고 새 형광펜 하나', summary);
check(saved.some((a) => a.subtype === 'FreeText' && a.contents === '한글 글상자입니다'), '저장 파일의 글상자(한글)');
check(saved.some((a) => a.subtype === 'Text' && a.contents === '원래 메모 — 고침') && saved.some((a) => a.subtype === 'Text' && a.contents === '새 메모 — 확인 부탁' && a.title === '시험 편집자'), '저장 파일의 메모(고친 것·새 것, 작성자)');
check(['Square', 'Ink', 'Line'].every((s) => saved.some((a) => a.subtype === s)), '사각형·펜·화살표도 저장');

// 17) 변경 기록
const changes = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data.changes ?? [];
check(changes.some((c) => c.kind === 'annot_removed' && c.annotType === 'highlight') && changes.some((c) => c.kind === 'annot_added' && c.viewer === '시험 편집자'), '변경 기록(누가 무엇을)', changes.map((c) => `${c.kind}:${c.annotType}`).join(','));

// 18) 다시 열면(저장 뒤 새 방) 같은 주석
await page.goto('about:blank');
await new Promise((r) => setTimeout(r, 1500));
const page2 = await browser.newPage({ viewport: { width: 1500, height: 950 } });
await page2.goto(`${BASE}${(await token(id)).data.editorUrl}`);
await page2.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await page2.waitForTimeout(1200);
const reopened = await page2.$$eval('.pdf-page[data-page="1"] [data-id]', (els) => [...new Set(els.map((e) => e.dataset.id))].length);
check(reopened === mineAfter && mineAfter === mineIds + 1, '저장 뒤 다시 열어도 같은 주석', `${reopened}/${mineAfter}`);
await page2.screenshot({ path: path.join(shotDir, '2-reopened.png') });

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 3).join(' | '));
await browser.close();
await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' });
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
