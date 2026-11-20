// 사용자가 보는 화면 찍기: 말뭉치(corpus)의 docx·xlsx·pptx를 서버에 올리고, 보기 화면과 편집 화면을 실제 브라우저(Playwright Chromium)로 찍는다.
// 보기(쪽) 화면이 그리는 PDF(LibreOffice 변환본)도 받아 둔다 — compare.py가 정답(MS Office PDF)과 나란히 놓는다.
// 사용: node audit/office/capture.mjs <서버 주소> <corpus 폴더> <출력 폴더> [이름 거르기]
// 서버는 따로 띄워 둔다(키 없이 업로드·토큰 API가 열린 로컬 시험 서버). 고객 문서를 쓰지 않는다.
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';
import { runPptxDomChecks } from './pptx_dom.mjs';

const [BASE = 'http://127.0.0.1:8141', CORPUS = 'audit/office/corpus', OUT = 'docs/quality-audit/out', FILTER = ''] = process.argv.slice(2);
for (const d of ['view', 'edit', 'lo']) fs.mkdirSync(path.join(OUT, d), { recursive: true });

async function upload(file) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
  const res = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!res.success) throw new Error(`업로드 실패 ${file}: ${res.error}`);
  const id = res.data.id;
  for (let i = 0; ; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data;
    if (st.status === 'ready') break;
    if (st.status === 'failed' || i > 600) throw new Error(`변환 실패 ${file}: ${st.error ?? 'timeout'}`);
    await new Promise((r) => setTimeout(r, 300));
  }
  const body = { viewer: 'quality-audit', canPrint: true, canDownload: true, canCopy: true, canEdit: true };
  const token = (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) })).json()).data;
  return { id, token };
}

const settle = (page, ms = 1500) => page.waitForLoadState('networkidle', { timeout: 20000 }).catch(() => {}).then(() => page.waitForTimeout(ms));

async function shotPagesView(page, base, name) {
  await page.goto(BASE + base.token.viewerUrl);
  await page.waitForSelector('.page-frame.page-ready', { timeout: 120000 });
  await settle(page);
  await page.screenshot({ path: path.join(OUT, 'view', `${name}.png`) });
  // 보기 화면이 그리는 바로 그 PDF(LibreOffice 변환본)
  const pdf = await fetch(`${BASE}/api/view/derived-pdf?token=${encodeURIComponent(base.token.token)}`);
  fs.writeFileSync(path.join(OUT, 'lo', `${name}.pdf`), Buffer.from(await pdf.arrayBuffer()));
}

async function shotGrid(page, url, ready, name, mode) {
  await page.goto(BASE + url);
  await page.waitForSelector(ready, { timeout: 120000 });
  await settle(page, 2000);
  const n = await page.locator('.grid-sheet-tab').count();
  const shots = [];
  for (let i = 0; i < Math.max(1, n); i += 1) {
    if (n > 1) { await page.locator('.grid-sheet-tab').nth(i).click(); await settle(page, 1500); }
    const file = path.join(OUT, mode, `${name}-s${i + 1}.png`);
    await page.screenshot({ path: file });
    const text = await page.locator(mode === 'view' ? '#grid-table' : '#grid-edit-table').innerText().catch(() => '');
    shots.push({ file, tab: n ? (await page.locator('.grid-sheet-tab').nth(i).innerText()).trim() : '', text });
  }
  return shots;
}

async function shotDocxEdit(page, base, name) {
  await page.goto(`${BASE}${base.token.editorUrl}&home=hidden`);
  await page.waitForSelector('.ProseMirror', { timeout: 120000 });
  await settle(page, 2500);
  const host = page.locator('#editor');
  const box = await host.boundingBox();
  await page.setViewportSize({ width: 1400, height: Math.min(16000, Math.ceil((box?.y ?? 0) + (box?.height ?? 1000) + 60)) });
  await settle(page, 800);
  await host.screenshot({ path: path.join(OUT, 'edit', `${name}.png`) });
  const text = await page.locator('.ProseMirror').innerText();
  const warn = await page.locator('body').innerText();
  const images = await page.locator('.ProseMirror img').count();
  return { text, images, unsupportedWarning: /이미지가 사라|지원하지 않는|보존되지/.test(warn) };
}

async function shotSlidesEdit(page, base, name) {
  await page.goto(`${BASE}${base.token.editorUrl}&home=hidden`);
  await page.waitForSelector('#slide-canvas .slide-textbox, #slide-canvas .slide-visual', { timeout: 120000 });
  await settle(page, 3000);
  const n = await page.locator('.slide-card').count();
  const texts = [];
  const dom = []; // pptx_checks.json DOM 점검(pptx_dom.mjs)
  for (let i = 0; i < n; i += 1) {
    await page.locator('.slide-card').nth(i).click();
    await settle(page, 2500);
    await page.locator('#slide-canvas').screenshot({ path: path.join(OUT, 'edit', `${name}-s${i + 1}.png`) });
    texts.push(await page.locator('#slide-canvas').innerText());
    dom.push(...(await runPptxDomChecks(page, name, i + 1)));
  }
  return { slides: n, texts, dom };
}

const browser = await chromium.launch();
const manifest = [];
try {
  for (const ext of ['docx', 'xlsx', 'pptx']) {
    const dir = path.join(CORPUS, ext);
    if (!fs.existsSync(dir)) continue;
    for (const f of fs.readdirSync(dir).filter((n) => n.endsWith(`.${ext}`) && !n.startsWith('~$') && n.includes(FILTER)).sort()) {
      const name = f.replace(/\.[^.]+$/, '');
      const entry = { name, ext, errors: [] };
      const t0 = Date.now();
      try {
        const base = await upload(path.join(dir, f));
        entry.convertMs = Date.now() - t0;
        const ctx = await browser.newContext({ viewport: { width: 1400, height: 1000 }, deviceScaleFactor: 1, locale: 'ko-KR' });
        const page = await ctx.newPage();
        page.on('pageerror', (e) => entry.errors.push(String(e).slice(0, 200)));
        try {
          if (ext === 'xlsx') {
            entry.view = await shotGrid(page, base.token.viewerUrl, '#grid-table td[data-r]', name, 'view');
            entry.edit = await shotGrid(page, `${base.token.editorUrl}&home=hidden`, '#grid-edit-table td[data-row]', name, 'edit');
          } else {
            await shotPagesView(page, base, name);
            entry.edit = ext === 'docx' ? await shotDocxEdit(page, base, name) : await shotSlidesEdit(page, base, name);
          }
        } finally { await ctx.close(); }
        await fetch(`${BASE}/api/documents/${base.id}`, { method: 'DELETE' }).catch(() => {});
      } catch (e) { entry.errors.push(String(e).slice(0, 300)); }
      console.log(`${entry.errors.length ? 'ERR ' : 'ok  '} ${f} ${entry.convertMs ?? '-'}ms ${entry.errors.join(' | ')}`);
      manifest.push(entry);
    }
  }
} finally { await browser.close(); }
fs.writeFileSync(path.join(OUT, `manifest${FILTER ? '-' + FILTER : ''}.json`), JSON.stringify(manifest, null, 2));
