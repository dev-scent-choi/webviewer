// 콘텐츠만 보기(ui=content, 설계 2026-11-22-artcraft-app-screen §4) 브라우저 검사 — 세 편집기(편집 토큰으로 열어도)와 보기 화면(view.html)에서
//  도구(제목 줄·탭·리본·상태 줄·섬네일 레일·수식 줄·시트 추가) 0개 보임, 본문 보임, 타자를 쳐도 내용이 바뀌지 않음, 메인 화면 "콘텐츠만 보기" 링크.
// 사용법(서버를 먼저 띄워 둘 것 — 업로드·토큰 API가 키 없이 열려 있는 로컬 시험 서버):
//   npm run build:editor && node scripts/verify-content-only-e2e.mjs [http://127.0.0.1:8169]
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8169';
const FILES = { docx: 'audit/office/corpus/docx/a1-korean-report.docx', xlsx: 'audit/office/corpus/xlsx/x07-multisheet-print.xlsx', pptx: 'audit/office/corpus/pptx/a2-korean-deck.pptx' };
const BODY = { docx: '.ProseMirror', xlsx: '#grid-edit-table td[data-row]', pptx: '#slide-canvas-wrap' };
const TOOLS = ['#toolbar', '.office-titlebar', '.ribbon-tabs', '.ribbon-panel', '#statusbar', '.formula-bar', '.grid-status-row', '#slides-rail', '#sheet-add', '#thumbs-panel', '.toolbar-row-tools'];
const MARK = 'ZQXJ콘텐츠';
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function upload(file) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
  const id = (await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json()).data.id;
  for (let i = 0; i < 1200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data;
    if (st.status === 'ready') return id;
    if (st.status === 'failed') throw new Error(`변환 실패 ${file}`);
    await sleep(250);
  }
  throw new Error(`변환 대기 초과 ${file}`);
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: 'content-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...body }) })
    .then((r) => r.json())
    .then((j) => j.data);
const withContent = (url) => {
  const u = new URL(url, BASE);
  u.searchParams.set('ui', 'content');
  return u.href;
};
const visibleTools = (page) => page.evaluate((sels) => sels.flatMap((s) => [...document.querySelectorAll(s)].filter((el) => el.getClientRects().length > 0).map(() => s)), TOOLS);
const bodyText = (page) => page.evaluate(() => document.body.innerText);

const browser = await chromium.launch();
try {
  const ids = {};
  for (const [ext, file] of Object.entries(FILES)) ids[ext] = await upload(file);

  for (const ext of ['docx', 'xlsx', 'pptx']) {
    const ctx = await browser.newContext({ viewport: { width: 1280, height: 860 } });
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(String(e)));
    // 편집 토큰으로 — 콘텐츠만 보기는 권한이 있어도 편집하지 않는다
    await page.goto(withContent((await token(ids[ext], { canEdit: true })).editorUrl));
    await page.waitForSelector(BODY[ext], { timeout: 120000 });
    await sleep(1500);
    const tools = await visibleTools(page);
    const box = await page.$eval(BODY[ext], (el) => { const r = el.getBoundingClientRect(); return { w: Math.round(r.width), h: Math.round(r.height), top: Math.round(r.top) }; });
    rec(`${ext} 도구 0개 보임·본문 보임`, tools.length === 0 && box.w > 0 && box.h > 0 && (await page.evaluate(() => document.body.classList.contains('ui-content'))), JSON.stringify({ tools, box }));
    const before = await bodyText(page);
    if (ext === 'docx') await page.click('.ProseMirror p >> nth=0');
    if (ext === 'xlsx') await page.click('#grid-edit-table td[data-row="1"] >> nth=0');
    if (ext === 'pptx') await page.dblclick('#slide-canvas-wrap', { position: { x: 200, y: 150 } }).catch(() => undefined);
    await page.keyboard.type(MARK);
    await page.keyboard.press('Enter');
    await page.keyboard.press('Control+v');
    await sleep(800);
    const after = await bodyText(page);
    rec(`${ext} 타자·붙여넣기 무효(내용 그대로)`, !after.includes(MARK) && after.length === before.length, `${before.length} → ${after.length}`);
    if (ext === 'docx') rec('docx 본문 편집 불가(contenteditable 아님)', (await page.getAttribute('.ProseMirror', 'contenteditable')) !== 'true');
    rec(`${ext} 화면 오류 없음`, errors.length === 0, errors.slice(0, 2).join(' | '));
    await ctx.close();
  }

  // 보기 화면(view.html) — 리브레오피스로 그린 docx
  {
    const ctx = await browser.newContext({ viewport: { width: 1280, height: 860 } });
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(String(e)));
    const tok = await token(ids.docx, {});
    rec('보기 docx — /view.html', new URL(tok.viewerUrl, BASE).pathname === '/view.html', tok.viewerUrl.split('?')[0]);
    await page.goto(withContent(tok.viewerUrl));
    await page.waitForSelector('.page-frame.page-ready', { timeout: 120000 });
    await sleep(800);
    const tools = await visibleTools(page);
    const top = await page.$eval('.page-frame', (el) => Math.round(el.getBoundingClientRect().top));
    rec('보기 docx 도구 0개 보임·첫 쪽이 맨 위(막대 자리 없음)', tools.length === 0 && top >= 0 && top < 40, JSON.stringify({ tools, top }));
    const before = await bodyText(page);
    await page.click('.page-frame >> nth=0');
    await page.keyboard.type(MARK);
    await sleep(500);
    rec('보기 docx 타자 무효', !(await bodyText(page)).includes(MARK) && (await bodyText(page)).length === before.length);
    rec('보기 docx 화면 오류 없음', errors.length === 0, errors.slice(0, 2).join(' | '));
    await ctx.close();
  }

  // 아트크래프트 토큰의 contentUrl — 앱 화면이 아니라 보기 화면의 엔진 쪽 그림(ArtcraftSource, 도구 없음, 설계 §4)
  {
    const ctx = await browser.newContext({ viewport: { width: 1280, height: 860 } });
    const page = await ctx.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(String(e)));
    const tok = await token(ids.docx, { engine: 'artcraft' });
    const cu = tok.contentUrl ? new URL(tok.contentUrl, BASE) : null;
    rec('아트크래프트 docx contentUrl → /view.html?…&ui=content', tok.engine === 'artcraft' && cu?.pathname === '/view.html' && cu.searchParams.get('ui') === 'content', tok.contentUrl?.split('?')[0] ?? String(tok.engineReason));
    if (cu) {
      await page.goto(cu.href);
      await page.waitForSelector('.page-frame.page-ready', { timeout: 120000 });
      await sleep(800);
      const engine = await page.$eval('#engine-badge', (b) => !b.hidden && !!b.dataset.workerTiming);
      const tools = await visibleTools(page);
      rec('아트크래프트 docx ui=content — 엔진이 그림·도구 0개 보임', engine && tools.length === 0 && (await page.evaluate(() => document.body.classList.contains('ui-content'))), JSON.stringify({ engine, tools }));
      rec('아트크래프트 docx 화면 오류 없음', errors.length === 0, errors.slice(0, 2).join(' | '));
    }
    await ctx.close();
  }

  // 메인 화면 "콘텐츠만 보기" 링크 → ui=content
  {
    const ctx = await browser.newContext({ viewport: { width: 1280, height: 860 } });
    const page = await ctx.newPage();
    await page.goto(`${BASE}/`);
    const row = page.locator('li.doc-item', { hasText: path.basename(FILES.docx) }).first();
    await row.waitFor({ timeout: 30000 });
    await row.locator('a', { hasText: '콘텐츠만 보기' }).click();
    await page.waitForURL(/ui=content/, { timeout: 30000 }).catch(() => undefined);
    const u = new URL(page.url());
    rec('메인 "콘텐츠만 보기" docx → 편집기 ui=content(ui=view 아님)', u.pathname === '/editor.html' && u.searchParams.getAll('ui').join() === 'content', u.pathname + u.search.replace(/token=[^&]+/, 'token=…'));
    // 단추 이름은 실제 엔진(2026-11-22 판정) — 엑셀은 우리 격자 화면이라 "기본 편집기로 열기"
    await page.goto(`${BASE}/`);
    const xrow = page.locator('li.doc-item', { hasText: path.basename(FILES.xlsx) }).first();
    await xrow.waitFor({ timeout: 30000 });
    const xlabels = (await xrow.locator('.doc-info > button, .doc-info > span > button').allTextContents()).map((t) => t.trim());
    rec('메인 xlsx 줄: "기본 편집기로 열기"·"아트크래프트로 열기"(리브레오피스 이름 없음)', xlabels.includes('기본 편집기로 열기') && xlabels.includes('아트크래프트로 열기') && !xlabels.includes('리브레오피스로 열기'), xlabels.join(','));
    await ctx.close();
  }
} catch (error) {
  rec('스크립트 오류', false, String(error?.stack ?? error));
} finally {
  await browser.close();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed ? 1 : 0);
