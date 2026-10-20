// PDF 편집기 7단계(변환) 브라우저 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-pdf-convert-e2e.mjs [BASE] [스크린샷 폴더] [워드 파일(선택)]
// 2쪽 한글 PDF(제목·문단·표)에 주석을 하나 더한 뒤 변환 창에서 워드(글만·모양 그대로)·엑셀·PPT·PNG(zip)·JPG·텍스트로 내려받아
// 내용을 확인하고, 다른 파일(워드·그림) → PDF 내려받기와 페이지 정리의 "파일에서 넣기"(워드·그림을 쪽으로)를 확인한다.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import fontkit from '@pdf-lib/fontkit';
import ExcelJS from 'exceljs';
import sharp from 'sharp';
import { strFromU8, unzipSync } from 'fflate';
import { PDFDocument, rgb } from 'pdf-lib';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-convert-'));
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

async function reportPdf() {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const font = await doc.embedFont(fs.readFileSync('assets/fonts/Pretendard-Regular.ttf'), { subset: false });
  const p = doc.addPage([595, 842]);
  p.drawText('분기 보고서', { x: 60, y: 770, size: 24, font, color: rgb(0.1, 0.2, 0.4) });
  p.drawText('이번 분기 가동률은 목표를 넘었습니다.', { x: 60, y: 720, size: 11, font });
  const rows = [
    ['항목', '값', '비고'],
    ['가동률', '97.2%', '정상'],
    ['생산량', '12,500', '증가'],
  ];
  rows.forEach((r, i) => r.forEach((cell, c) => p.drawText(cell, { x: 60 + c * 120, y: 640 - i * 20, size: 11, font })));
  doc.addPage([595, 842]).drawText('둘째 쪽 내용', { x: 60, y: 770, size: 11, font });
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
const pdfText = async (bytes) => {
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise;
  const out = [];
  for (let p = 1; p <= pdf.numPages; p += 1) out.push((await (await pdf.getPage(p)).getTextContent()).items.map((i) => i.str).join(''));
  const n = pdf.numPages;
  await pdf.destroy();
  return { pages: n, text: out.join('\n') };
};

const id = await upload(await reportPdf(), '분기보고서.pdf');
const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1500, height: 950 }, acceptDownloads: true });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => {
  if (m.type() === 'error') errors.push(m.text());
});
const t = await token(id);
await page.goto(`${BASE}${t.data.editorUrl}`);
await page.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await page.waitForTimeout(800);

// 지금 상태가 변환에 들어가는지 보려고 사각형 주석 하나
await page.click('.pdf-mode[data-mode="edit"]');
await page.click('[data-group="shape"] .pdf-float-more');
await page.click('[data-group="shape"] [data-pick="rect"]');
const pb = await page.locator('.pdf-page[data-page="1"]').boundingBox();
await page.mouse.move(pb.x + 300, pb.y + 500);
await page.mouse.down();
await page.mouse.move(pb.x + 420, pb.y + 560, { steps: 4 });
await page.mouse.up();
await page.waitForTimeout(300);
await page.click('.pdf-mode[data-mode="tools"]');

const convertTo = async (value, dpi) => {
  await page.click('.pdf-tool-item[data-tool="convert"]');
  await page.waitForSelector('#convert-dialog:not([hidden])');
  await page.check(`#convert-targets input[value="${value}"]`);
  if (dpi) await page.selectOption('#convert-dpi', String(dpi));
  const [dl] = await Promise.all([page.waitForEvent('download', { timeout: 120000 }), page.click('#convert-run')]);
  const file = path.join(shotDir, dl.suggestedFilename());
  await dl.saveAs(file);
  await page.waitForFunction(() => (document.getElementById('convert-status')?.textContent ?? '').includes('내려받았습니다'));
  await page.click('#convert-dialog [data-close]');
  return { name: dl.suggestedFilename(), bytes: fs.readFileSync(file) };
};
const docXml = (bytes) => strFromU8(unzipSync(new Uint8Array(bytes))['word/document.xml']);
const textsOf = (xml) => [...xml.matchAll(/<(?:w|a):t[^>]*>([^<]*)</g)].map((m) => m[1]).join('');

// 1) 워드 — 글만
const d1 = await convertTo('docx-text');
check(d1.name === '분기보고서-글만.docx' && textsOf(docXml(d1.bytes)).includes('이번 분기 가동률은 목표를 넘었습니다.') && docXml(d1.bytes).includes('w:type="page"'), '워드(글만)', d1.name);
// 2) 워드 — 모양 그대로
const d2 = await convertTo('docx');
check(d2.name === '분기보고서.docx' && textsOf(docXml(d2.bytes)).includes('분기 보고서'), '워드(모양 그대로)', d2.name);
// 3) 엑셀
const x = await convertTo('xlsx');
const wb = new ExcelJS.Workbook();
await wb.xlsx.load(x.bytes);
let pct = null;
wb.worksheets[0].eachRow((row) => {
  if (row.getCell(1).value === '가동률') pct = row.getCell(2).value;
});
check(x.name === '분기보고서.xlsx' && wb.worksheets.length === 2 && Math.abs(pct - 0.972) < 1e-9, '엑셀(표 · 숫자)', `${wb.worksheets.map((w) => w.name).join(',')} / ${pct}`);
// 4) PPT
const pp = await convertTo('pptx');
const slides = Object.keys(unzipSync(new Uint8Array(pp.bytes))).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k));
check(pp.name === '분기보고서.pptx' && slides.length === 2, '파워포인트(쪽마다 슬라이드)', `${slides.length}장`);
// 5) PNG — 2쪽이면 zip, 150dpi
const png = await convertTo('png', 150);
const zip = unzipSync(new Uint8Array(png.bytes));
const names = Object.keys(zip);
const meta1 = await sharp(Buffer.from(zip[names[0]])).metadata();
check(png.name === '분기보고서-그림.zip' && names.length === 2 && meta1.format === 'png' && Math.abs(meta1.width - 1240) <= 2, 'PNG(zip, 150dpi)', `${names.join(',')} ${meta1.width}×${meta1.height}`);
// 그림에 지금 상태(사각형 주석)가 들어갔는지 — 빨간 화소
const { data, info } = await sharp(Buffer.from(zip[names[0]])).raw().toBuffer({ resolveWithObject: true });
let red = 0;
for (let i = 0; i < data.length; i += info.channels * 7) if (data[i] > 180 && data[i + 1] < 90 && data[i + 2] < 90) red += 1;
check(red > 50, '그림에 지금 편집 상태(주석)가 들어감', `빨간 점 ${red}`);
// 6) JPG 72dpi
const jpg = await convertTo('jpg', 72);
const jmeta = await sharp(Buffer.from(unzipSync(new Uint8Array(jpg.bytes))[Object.keys(unzipSync(new Uint8Array(jpg.bytes)))[0]])).metadata();
check(jmeta.format === 'jpeg' && Math.abs(jmeta.width - 595) <= 2, 'JPG(72dpi)', `${jmeta.width}×${jmeta.height}`);
// 7) 텍스트
const txt = await convertTo('txt');
const tx = txt.bytes.toString('utf8');
check(txt.name === '분기보고서.txt' && tx.includes('가동률 97.2% 정상') && tx.includes('둘째 쪽 내용'), '텍스트', tx.replace(/\s+/g, ' ').slice(1, 60));
await page.click('.pdf-tool-item[data-tool="convert"]');
await page.check('#convert-targets input[value="png"]');
await page.screenshot({ path: path.join(shotDir, '1-convert.png') });

// 8) 다른 파일 → PDF(워드·그림)
const docxPath = process.argv[4] ?? path.join(shotDir, '분기보고서-글만.docx');
const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('#convert-file-btn')]);
const [dl8] = await Promise.all([page.waitForEvent('download', { timeout: 120000 }), chooser.setFiles(docxPath)]);
const pdf8 = fs.readFileSync(await dl8.path());
const t8 = await pdfText(pdf8);
check(dl8.suggestedFilename().endsWith('.pdf') && t8.pages >= 1 && t8.text.length > 0, '워드 파일 → PDF 내려받기', `${dl8.suggestedFilename()} ${t8.pages}쪽`);
await page.click('#convert-dialog [data-close]');

// 9) 페이지 정리 — 그림·워드를 쪽으로 넣기
const imgPath = path.join(shotDir, '도장.png');
fs.writeFileSync(imgPath, await sharp({ create: { width: 400, height: 300, channels: 3, background: '#d93025' } }).png().toBuffer());
await page.click('.pdf-tool-item[data-tool="organize"]');
await page.waitForSelector('#organize:not([hidden])');
const [ch2] = await Promise.all([page.waitForEvent('filechooser'), page.click('#org-insert-file')]);
await ch2.setFiles([imgPath, docxPath]);
await page.waitForFunction(() => document.querySelectorAll('#org-grid .pdf-org-card').length >= 4, null, { timeout: 120000 });
const cards = await page.$$eval('#org-grid .pdf-org-card', (els) => els.map((e) => e.dataset.key));
check(cards.length >= 4 && cards.filter((k) => k.startsWith('p:')).length >= 2, '그림·워드 파일을 쪽으로 넣기(PDF로 바꿔서)', cards.join(','));
// 넣은 워드 쪽의 카드에 실제로 글이 그려졌는지(어두운 점)
await page.waitForFunction((n) => document.querySelectorAll('#org-grid .pdf-org-card canvas').length >= n, cards.length, { timeout: 30000 });
const wordKey = cards.filter((k) => k.startsWith('p:'))[1];
const ink = await page.locator(`#org-grid .pdf-org-card[data-key="${wordKey}"] canvas`).evaluate((c) => {
  const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
  let n = 0;
  for (let i = 0; i < d.length; i += 4) if (d[i] + d[i + 1] + d[i + 2] < 600) n += 1;
  return n;
});
check(ink > 20, '넣은 워드 쪽에 글이 그려짐', `어두운 점 ${ink}`);
await page.screenshot({ path: path.join(shotDir, '2-organize.png') });
await page.keyboard.press('Escape');

// 10) 내려받기 권한 없는 토큰 — 변환 없음
const nd = await ctx.newPage();
await nd.goto(`${BASE}${(await token(id, { viewer: '보기만', canEdit: false, canDownload: false })).data.editorUrl}`).catch(() => undefined);
const ndConvert = await fetch(`${BASE}/api/view/pdf-convert?token=${encodeURIComponent((await token(id, { canDownload: false })).data.token)}`, {
  method: 'POST',
  headers: { 'content-type': 'application/json' },
  body: JSON.stringify({ to: 'txt' }),
});
check(ndConvert.status === 403, '내려받기 권한이 없으면 변환 거부', String(ndConvert.status));
await nd.close();

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 3).join(' | '));
await browser.close();
await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' });
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
