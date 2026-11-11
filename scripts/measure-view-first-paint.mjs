// 휴대폰(CPU×4) 첫 쪽 시간 — 쪽 틀 하나가 그려질 때까지(.page-frame.page-ready). 시험 서버에서만.
//   node scripts/measure-view-first-paint.mjs [http://127.0.0.1:8124] [/view.html | /viewer.html] [횟수=5]
// 시험 문서는 매번 만든다(20쪽 글자 PDF). 열람만(서버 PNG)·내려받기(새 화면은 pdf.js) 토큰 각각의 중앙값을 찍는다.
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const PATH = process.argv[3] ?? '/view.html';
const RUNS = Number(process.argv[4] ?? 5);

const doc = await PDFDocument.create();
const font = await doc.embedFont(StandardFonts.Helvetica);
for (let i = 1; i <= 20; i += 1) {
  const page = doc.addPage([595, 842]);
  for (let y = 780; y > 60; y -= 18) page.drawText(`First paint page ${i} line ${y}`, { x: 50, y, size: 11, font });
}
const form = new FormData();
form.append('file', new Blob([await doc.save()]), '첫화면.pdf');
const up = (await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json()).data;
for (let i = 0; i < 200; i += 1) {
  if ((await (await fetch(`${BASE}/api/documents/${up.id}`)).json()).data.status === 'ready') break;
  await new Promise((r) => setTimeout(r, 250));
}
const token = async (body) =>
  (await (await fetch(`${BASE}/api/documents/${up.id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: '측정', canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...body }) })).json()).data.token;

const browser = await chromium.launch();
const median = (xs) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];
for (const [label, body] of [['열람만', {}], ['내려받기', { canDownload: true, canCopy: true }]]) {
  const t = await token(body);
  const times = [];
  for (let run = 0; run < RUNS; run += 1) {
    const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 3, isMobile: true, hasTouch: true });
    const page = await ctx.newPage();
    const cdp = await ctx.newCDPSession(page);
    await cdp.send('Emulation.setCPUThrottlingRate', { rate: 4 });
    const t0 = Date.now();
    await page.goto(`${BASE}${PATH}?token=${encodeURIComponent(t)}`);
    await page.waitForSelector('.page-frame.page-ready', { timeout: 120000 });
    times.push(Date.now() - t0);
    await ctx.close();
  }
  console.log(`${PATH} ${label}: 중앙값 ${median(times)}ms (${times.join(', ')})`);
}
await browser.close();
await fetch(`${BASE}/api/documents/${up.id}`, { method: 'DELETE' });
