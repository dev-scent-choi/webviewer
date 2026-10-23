// PDF 편집기 찾기 패널 — 3단 검색(이 페이지·전체 문서·결과 안에서), 제목만, 결과 목록·이동, 문단 테두리, 미리보기, 스캔 안내.
// 사용: node scripts/verify-pdf-search-e2e.mjs [서버주소] [스크린샷 폴더]
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import fontkit from '@pdf-lib/fontkit';
import { PDFDocument, rgb } from 'pdf-lib';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-search-'));
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

// 설비 지침서 6쪽: 쪽마다 큰 제목 + 번호 제목 + 본문. '펌프'는 1·2·4·6쪽에 8번(제목 3번 포함), '배관'은 2·3·4쪽(2·4쪽은 펌프와 같은 문단)
const PAGES = [
  ['설비 운전 지침', '1. 개요', ['이 지침은 냉각수 펌프와 공조 설비의 운전 방법을 정합니다.', '운전원은 매일 아침 점검표를 작성합니다.']],
  ['펌프 점검', '2. 원심펌프 P101', ['원심펌프는 소음과 진동을 매일 확인합니다.', '배관 연결부 누수가 보이면 즉시 정지합니다.']],
  ['배관 관리', '3. 배관 압력', ['배관 압력은 0.3MPa 이하로 유지합니다.', '압력계는 분기마다 교정합니다.']],
  ['예비 설비', '4. 예비 펌프 전환', ['주 펌프가 멈추면 예비 펌프로 전환하고 배관 밸브를 엽니다.', '전환 기록은 운전 일지에 남깁니다.']],
  ['전기 설비', '5. 수전 설비', ['수전반은 월 1회 열화상 점검을 합니다.', '차단기 동작 시험은 반기마다 합니다.']],
  ['부록', '6. 점검 주기표', ['펌프 베어링 그리스는 3개월마다 보충합니다.', '필터는 오염도에 따라 교체합니다.']],
];

async function manualPdf() {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const font = await doc.embedFont(fs.readFileSync('assets/fonts/Pretendard-Regular.ttf'), { subset: false });
  for (const [title, sub, body] of PAGES) {
    const page = doc.addPage([595, 842]);
    page.drawText(title, { x: 60, y: 770, size: 22, font, color: rgb(0.1, 0.2, 0.4) });
    page.drawText(sub, { x: 60, y: 720, size: 14, font });
    body.forEach((line, i) => page.drawText(line, { x: 60, y: 690 - i * 17, size: 11, font }));
  }
  return doc.save();
}

async function scannedPdf() {
  const doc = await PDFDocument.create();
  const page = doc.addPage([400, 300]);
  page.drawRectangle({ x: 40, y: 40, width: 320, height: 220, color: rgb(0.85, 0.85, 0.85) });
  return doc.save();
}

async function upload(bytes, name) {
  const form = new FormData();
  form.append('file', new Blob([bytes], { type: 'application/pdf' }), name);
  const up = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  const id = up.data.id;
  for (let i = 0; i < 200; i++) {
    const st = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data;
    if (st.status === 'ready' || st.editReady) break;
    await new Promise((r) => setTimeout(r, 200));
  }
  const t = (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: '검색 시험', canEdit: true, canDownload: true, canCopy: true, canPrint: true, pdfEditor: true }) })).json()).data;
  return { id, url: BASE + t.editorUrl };
}

const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1440, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
const doc = await upload(await manualPdf(), '설비지침.pdf');
await page.goto(doc.url);
await page.waitForSelector('.pdf-page[data-page="1"] .pdf-text-layer span', { timeout: 60000 });

const state = () => page.evaluate(() => ({
  open: !document.getElementById('panel-search').hidden,
  count: document.getElementById('find-count').textContent,
  pages: document.getElementById('find-count').dataset.pages,
  items: document.querySelectorAll('.pdf-sx-item').length,
  groups: [...document.querySelectorAll('.pdf-sx-page')].map((e) => e.firstChild.textContent.trim()),
  selected: document.querySelector('.pdf-sx-item[aria-selected="true"] .pdf-sx-snip')?.textContent ?? '',
  heads: [...document.querySelectorAll('.pdf-sx-item')].map((e) => e.querySelector('.pdf-sx-head')?.textContent ?? ''),
  total: document.getElementById('sx-total').textContent,
  hits: document.querySelectorAll('.pdf-find-hit').length,
  current: document.querySelectorAll('.pdf-find-hit.current').length,
  block: document.querySelectorAll('.pdf-find-block').length,
  pageInput: document.getElementById('page-input').value,
  previewHeading: document.getElementById('sx-preview-heading').textContent,
  previewBody: document.getElementById('sx-preview-body').textContent,
  empty: document.getElementById('sx-empty').hidden ? '' : document.getElementById('sx-empty').textContent,
  trail: document.getElementById('sx-trail').hidden ? '' : document.getElementById('sx-trail').textContent,
}));
const settle = async () => {
  await page.waitForFunction(() => !/찾는 중|…/.test(document.getElementById('find-count').textContent), null, { timeout: 15000 }).catch(() => {});
  await page.waitForTimeout(400);
};

// ② 전체 문서
await page.keyboard.press('Control+f');
let s = await state();
check(s.open && (await page.getAttribute('#rail-search', 'aria-pressed')) === 'true', 'Ctrl+F → 찾기 패널 열림·오른쪽 단추 눌림');
check(await page.evaluate(() => document.activeElement?.id === 'find-input'), '입력칸에 초점');
await page.type('#find-input', '펌프');
await settle();
s = await state();
check(s.count === '1/8' && s.pages === '4쪽' && s.total === '8', '"펌프" 전체 문서 → 8건·4쪽(제목 옆 개수 8)', `${s.count} ${s.pages} ${s.total}`);
check(s.groups.join(',') === '1쪽,2쪽,4쪽,6쪽', '결과를 쪽별로 묶음', s.groups.join(','));
check(s.heads[0].includes('개요') || s.heads[0].includes('설비 운전 지침'), '결과마다 속한 제목', s.heads.join(' | '));
check(s.hits >= 1 && s.current === 1 && s.block === 1, '본문 노란 표시·현재 결과·문단 테두리', `hits ${s.hits} current ${s.current} block ${s.block}`);
check(/1쪽/.test(s.previewHeading) && s.previewBody.includes('펌프'), '미리보기에 그 문단', `${s.previewHeading} / ${s.previewBody.slice(0, 40)}`);
await page.screenshot({ path: path.join(shotDir, '1-doc.png') });

await page.keyboard.press('Enter');
await page.waitForTimeout(400);
s = await state();
check(s.count === '2/8', 'Enter → 다음 결과', s.count);
await page.locator('.pdf-sx-item').nth(4).click(); // 4쪽 두 번째(예비 펌프)
await page.waitForTimeout(700);
s = await state();
check(s.count === '5/8' && s.pageInput === '4', '목록에서 누르면 그 쪽으로 이동', `${s.count} 쪽 ${s.pageInput}`);
await page.keyboard.press('Shift+Enter').catch(() => {});

// 제목만
await page.check('#sx-headings');
await settle();
s = await state();
check(s.count === '1/3' && s.heads.every((h) => h.includes('제목')), '제목만 → 제목 3건(펌프 점검, 2. 원심펌프, 4. 예비 펌프 전환)', `${s.count} ${s.heads.join('|')}`);
await page.uncheck('#sx-headings');
await settle();

// ① 이 페이지 — 지금 4쪽
await page.fill('#page-input', '4');
await page.press('#page-input', 'Enter');
await page.waitForTimeout(500);
await page.click('[data-scope="page"]');
await settle();
s = await state();
check(s.groups.join(',') === '4쪽' && s.count.endsWith('/3'), '이 페이지 → 4쪽 결과만(제목 1 + 본문 2)', `${s.groups} ${s.count}`);
await page.screenshot({ path: path.join(shotDir, '2-page.png') });

// ③ 결과 안에서 — 전체 '펌프' 결과 안에서 '배관'
await page.click('[data-scope="doc"]');
await settle();
check(!(await page.isDisabled('[data-scope="within"]')), '결과가 있으면 "결과 안에서" 켜짐');
await page.click('[data-scope="within"]');
check((await page.inputValue('#find-input')) === '' && (await page.getAttribute('[data-scope="within"]', 'aria-checked')) === 'true', '결과 안에서 → 입력칸 비우고 단계 시작');
await page.type('#find-input', '배관');
await settle();
s = await state();
check(s.groups.join(',') === '2쪽,4쪽', '"펌프" 결과 문단 안의 "배관"만 → 2·4쪽 (3쪽은 빠짐)', `${s.groups} ${s.count}`);
check(s.trail.includes('펌프') && s.trail.includes('배관'), '단계 표시 "펌프 8 › 배관"', s.trail);
await page.screenshot({ path: path.join(shotDir, '3-within.png') });
await page.click('#sx-trail [data-stage="0"]');
await settle();
s = await state();
check(s.count.endsWith('/8') && (await page.inputValue('#find-input')) === '펌프' && !s.trail, '단계를 누르면 앞 단계로 돌아감', `${s.count} trail "${s.trail}"`);

// 미리보기: 쪽 전체 글 · 쪽 넘기기 · 복사
await page.click('[data-view="page"]');
await page.waitForTimeout(400);
s = await state();
check(/쪽 전체 글/.test(s.previewHeading) && s.previewBody.includes('설비 운전 지침'), '쪽 전체 글 미리보기', s.previewHeading);
await page.click('#sx-page-next');
await page.waitForTimeout(600);
s = await state();
check(/2쪽 전체 글/.test(s.previewHeading) && s.previewBody.includes('원심펌프') && s.pageInput === '2', '다음 쪽 → 미리보기·본문 함께 넘어감', `${s.previewHeading} 쪽 ${s.pageInput}`);
await page.click('#sx-copy');
await page.waitForTimeout(300);
const clip = await page.evaluate(() => navigator.clipboard.readText()).catch(() => '');
check(clip.includes('원심펌프'), '글 복사', clip.slice(0, 30));
await page.screenshot({ path: path.join(shotDir, '4-preview.png') });

// Esc → 닫고 전에 보던 패널로, 강조 지움
await page.focus('#find-input');
await page.keyboard.press('Escape');
await page.waitForTimeout(300);
s = await state();
check(!s.open && s.hits === 0 && s.block === 0 && !(await page.isHidden('#panel-thumbs')), 'Esc → 닫고 강조 지움·페이지 패널로 돌아감');

// 위 줄 찾기 상자 → 패널
await page.fill('#top-find', '압력');
await page.press('#top-find', 'Enter');
await settle();
s = await state();
check(s.open && s.count === '1/3', '위쪽 "문서에서 찾기" Enter → 패널로 찾기', s.count);

// 휴대폰: 아래 시트
const m = await browser.newPage({ viewport: { width: 390, height: 844 }, isMobile: true, hasTouch: true });
m.on('pageerror', (e) => errors.push('m: ' + String(e)));
await m.goto(doc.url);
await m.waitForSelector('.pdf-page[data-page="1"] .pdf-text-layer span', { timeout: 60000 });
await m.click('#rail-search');
await m.fill('#find-input', '펌프');
await m.waitForTimeout(1200);
const sheet = await m.evaluate(() => {
  const r = document.getElementById('inspector').getBoundingClientRect();
  return { top: Math.round(r.top), h: Math.round(r.height), w: Math.round(r.width), items: document.querySelectorAll('.pdf-sx-item').length, overflow: document.documentElement.scrollWidth - innerWidth };
});
check(sheet.w === 390 && sheet.top > 200 && sheet.items === 8 && sheet.overflow <= 0, '휴대폰: 아래에서 올라오는 시트로 찾기', JSON.stringify(sheet));
await m.screenshot({ path: path.join(shotDir, '5-mobile.png') });
await m.close();

// 스캔 문서 안내
const scan = await upload(await scannedPdf(), '스캔.pdf');
await page.goto(scan.url);
await page.waitForSelector('.pdf-page[data-page="1"] canvas', { timeout: 60000 });
await page.waitForTimeout(800);
await page.keyboard.press('Control+f');
await page.type('#find-input', '펌프');
await settle();
s = await state();
check(/글자 정보가 없습니다/.test(s.empty) && (await page.isVisible('[data-ocr]')), '스캔 문서 → 문자 인식 안내와 단추', s.empty.slice(0, 40));
await page.screenshot({ path: path.join(shotDir, '6-scanned.png') });

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 2).join(' | '));
await browser.close();
for (const d of [doc, scan]) await fetch(`${BASE}/api/documents/${d.id}`, { method: 'DELETE' }).catch(() => {});
const failed = results.filter((r) => !r).length;
console.log(failed ? `FAILED ${failed}/${results.length} — 스크린샷: ${shotDir}` : `ALL CHECKS PASSED (${results.length}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(failed ? 1 : 0);
