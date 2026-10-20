// PDF 편집기 8단계(OCR) 브라우저 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-pdf-ocr-e2e.mjs [BASE] [스크린샷 폴더]
// 1쪽은 글자가 그림으로만 있는 한글 "스캔본", 2쪽은 진짜 글자가 있는 쪽. 글자 인식(글자 없는 쪽만) → 1쪽만 읽음,
// 바로 찾기·글자 고르기, 실행 취소, 저장 파일에 보이지 않는 글자(위치가 그림의 글자와 맞음), 다른 참여자, 읽기 전용을 확인한다.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import fontkit from '@pdf-lib/fontkit';
import sharp from 'sharp';
import { PDFDocument } from 'pdf-lib';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-ocr-'));
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

/** 595×842pt 쪽에 200dpi 그림(1654×2339) — 글자는 그림 속에만 */
async function scannedPdf() {
  const lines = [
    ['분기 보고서', 64, 300],
    ['이번 분기 가동률은 97.2%로 목표를 넘었습니다.', 44, 460],
    ['담당자 홍길동 연락처 010-1234-5678', 44, 560],
    ['Revenue grew 12.5% year over year.', 44, 660],
  ];
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="1654" height="2339"><rect width="100%" height="100%" fill="#fbfaf6"/>${lines.map(([t, s, y]) => `<text x="160" y="${y}" font-family="Malgun Gothic" font-size="${s}" fill="#222">${t}</text>`).join('')}</svg>`;
  const jpg = await sharp(Buffer.from(svg)).rotate(0.4, { background: '#fbfaf6' }).blur(0.6).jpeg({ quality: 70 }).toBuffer();
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const p1 = doc.addPage([595, 842]);
  p1.drawImage(await doc.embedJpg(jpg), { x: 0, y: 0, width: 595, height: 842 });
  const font = await doc.embedFont(fs.readFileSync('assets/fonts/Pretendard-Regular.ttf'), { subset: false });
  doc.addPage([595, 842]).drawText('둘째 쪽은 진짜 글자', { x: 60, y: 760, size: 14, font });
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

const id = await upload(await scannedPdf(), '스캔본.pdf');
const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1500, height: 950 } });
const page = await ctx.newPage();
const errors = [];
const outside = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => {
  if (m.type() === 'error') errors.push(m.text());
});
// 외부로 나가는 요청이 없어야 한다(문서 그림·학습 데이터 모두 이 서버에서만)
page.on('request', (r) => {
  if (!r.url().startsWith(BASE) && !r.url().startsWith('data:') && !r.url().startsWith('blob:')) outside.push(r.url());
});
const t = await token(id);
await page.goto(`${BASE}${t.data.editorUrl}`);
await page.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await page.waitForTimeout(800);
check((await page.locator('.pdf-page[data-page="1"] .pdf-text-layer').textContent()).trim() === '', '1쪽은 글자가 없음(그림뿐)');

// 1) 글자 인식 — 글자 없는 쪽만
await page.click('.pdf-tool-item[data-tool="ocr"]');
await page.waitForSelector('#ocr-dialog:not([hidden])');
const t0 = Date.now();
await page.click('#ocr-run');
await page.waitForFunction(() => /낱말 \d+개를 읽었습니다/.test(document.getElementById('ocr-status')?.textContent ?? ''), null, { timeout: 180000 });
const status = await page.textContent('#ocr-status');
check(/^1쪽에서 낱말 \d+개/.test(status), '글자 없는 쪽(1쪽)만 읽음', `${status} (${Math.round((Date.now() - t0) / 1000)}초)`);
await page.screenshot({ path: path.join(shotDir, '1-ocr-dialog.png') });
await page.click('#ocr-dialog [data-close]');

// 2) 바로 찾기 — 인식한 낱말에서
await page.waitForFunction(() => document.querySelectorAll('.pdf-page[data-page="1"] .pdf-ocr-layer span').length > 5, null, { timeout: 10000 });
const ocrText = await page.locator('.pdf-page[data-page="1"] .pdf-ocr-layer').textContent();
check(ocrText.includes('가동률은') && ocrText.includes('홍길동') && ocrText.includes('Revenue'), '인식 결과(한글·영문)', ocrText.replace(/\s+/g, ' ').slice(0, 80));
await page.locator('#pdf-scroll').click({ position: { x: 5, y: 5 } });
await page.keyboard.press('Control+f');
await page.fill('#find-input', '가동률은');
await page.waitForFunction(() => /1\/1/.test(document.getElementById('find-count')?.textContent ?? ''), null, { timeout: 10000 });
const hit = await page.locator('.pdf-page[data-page="1"] .pdf-find-hit').first().boundingBox();
// 그림 속 "가동률은"(200dpi 그림의 x≈160+, y≈460 기준선) 근처인지 — 쪽 상자 기준 비율로
const pb = await page.locator('.pdf-page[data-page="1"]').boundingBox();
const rx = (hit.x - pb.x) / pb.width;
const ry = (hit.y + hit.height - pb.y) / pb.height;
check(hit && rx > 0.05 && rx < 0.45 && ry > 0.17 && ry < 0.22, '찾기 표시가 그림의 글자 자리에', `x ${rx.toFixed(3)}, 아래 ${ry.toFixed(3)}`);
await page.keyboard.press('Escape');
await page.screenshot({ path: path.join(shotDir, '2-find.png') });

// 3) 실행 취소·다시 실행
await page.locator('#pdf-scroll').click({ position: { x: 5, y: 5 } });
await page.keyboard.press('Control+z');
await page.waitForTimeout(500);
const afterUndo = await page.$$eval('.pdf-page[data-page="1"] .pdf-ocr-layer span', (els) => els.length);
await page.keyboard.press('Control+y');
await page.waitForTimeout(500);
const afterRedo = await page.$$eval('.pdf-page[data-page="1"] .pdf-ocr-layer span', (els) => els.length);
check(afterUndo === 0 && afterRedo > 5, '인식 결과도 Ctrl+Z / Ctrl+Y', `${afterUndo} → ${afterRedo}`);

// 4) 저장 파일 — 보이지 않는 글자, 2쪽은 그대로
const dl = new Uint8Array(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(t.data.token)}`)).arrayBuffer());
const pdf = await getDocument({ data: dl, isEvalSupported: false }).promise;
const p1items = (await (await pdf.getPage(1)).getTextContent()).items.filter((i) => i.str.trim());
const p2text = (await (await pdf.getPage(2)).getTextContent()).items.map((i) => i.str).join('');
await pdf.destroy();
const joined = p1items.map((i) => i.str).join(' ');
const gado = p1items.find((i) => i.str.includes('가동률은'));
check(joined.includes('가동률은') && joined.includes('Revenue') && p2text.includes('둘째 쪽은 진짜 글자') && !p2text.includes('가동률'), '저장 파일: 1쪽에 인식한 글자, 2쪽은 그대로', joined.slice(0, 60));
// 그림 속 "이번 분기 가동률은" 줄의 기준선 ≈ 842 - 460×(842/2339) ≈ 676pt
check(gado && Math.abs(gado.transform[5] - 676) < 8, '저장한 글자의 위치가 그림의 글자와 맞음', gado ? `y ${gado.transform[5].toFixed(1)}` : '없음');
fs.writeFileSync(path.join(shotDir, 'saved.pdf'), dl);

// 5) 변경 기록
const changes = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data.changes ?? [];
check(changes.some((c) => c.kind === 'ocr_added' && c.page === 1), '변경 기록(글자 인식)', changes.map((c) => `${c.kind}:${c.text ?? ''}`).join(','));

// 6) 다른 참여자 — 인식 결과가 바로 보이고 찾아진다
const other = await ctx.newPage();
await other.goto(`${BASE}${(await token(id, { viewer: '다른 사람' })).data.editorUrl}`);
await other.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await other.waitForFunction(() => document.querySelectorAll('.pdf-page[data-page="1"] .pdf-ocr-layer span').length > 5, null, { timeout: 15000 });
check(true, '다른 참여자 화면에도 인식 결과');
await other.close();

// 7) 읽기 전용 — OCR 도구 없음
const ro = await ctx.newPage();
await ro.goto(`${BASE}${(await token(id, { viewer: '읽기만', canEdit: false })).data.editorUrl}`);
await ro.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
check(!(await ro.isVisible('.pdf-tool-item[data-tool="ocr"]')), '읽기 전용: 글자 인식 없음');
await ro.close();

check(outside.length === 0, '외부로 나간 요청 없음(학습 데이터·worker 자체 호스팅)', outside.slice(0, 3).join(' | '));
check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 3).join(' | '));
await browser.close();
await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' });
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
