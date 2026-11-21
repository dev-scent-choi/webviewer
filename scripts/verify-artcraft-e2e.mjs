// 아트크래프트(브라우저 엔진) 보기 — 설계 2026-11-21 §7의 브라우저 검사 일곱 가지 + 메인 화면 단추·열린 뒤 Worker 죽음, 쪽 수 비교 기록.
//  (1) 아트크래프트로 열려 쪽 수·첫 쪽 그림(잉크 비율·LibreOffice 첫 쪽과의 유사도·한글 글자 층)
//  (2) 확대·회전·쪽 이동·섬네일·찾기  (3) 서버 변환 요청 0건  (4) wasm을 막으면 리브레오피스로 자동 전환 + 안내
//  (5) 내려받기 권한 없는 토큰은 서버 변환으로  (6) CPU×4 첫 쪽 시간(로컬 서버라 받기 시간은 사실상 0 — 목표 2초, 고객 pptx 3초)
//  (7) 두 번째 열기에서 wasm·글꼴을 서버에서 다시 받지 않음(서버 요청 기록으로 센다 — 304 재검증도 요청으로 잡힌다)
// 이 스크립트가 시험 서버를 직접 띄운다(npm run build:editor, npm run build:artcraft 뒤, LibreOffice 필요):
//   ARTCRAFT_E2E_PRIVATE_DIR=<고객형 원본 폴더> node scripts/verify-artcraft-e2e.mjs [포트=8150, 8150~8159]
// 고객형 문서는 그 폴더에서 .audit-run/artcraft-e2e-<포트>/로 복사해 쓴다(저장소에 넣지 않는다). 폴더가 없으면 그 둘은 건너뛴다(SKIP 줄).
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import sharp from 'sharp';
import { chromium } from 'playwright';

const PORT = Number(process.argv[2] ?? 8150);
// 운영 포트(8000·8010·8020·8100·8110·8130·8148)와 겹치지 않는 시험 전용 범위만
if (!(PORT >= 8150 && PORT <= 8159)) throw new Error(`시험 포트는 8150~8159만 쓴다(받은 값 ${PORT})`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `artcraft-e2e-${PORT}`);
const ORIGINALS =
  process.env.ARTCRAFT_E2E_PRIVATE_DIR ?? process.env.ARTCRAFT_E2E_ORIGINALS ?? 'C:/Users/OpenMinds/projects/webviewer-q-all/.audit-run/test-server/originals';
const CUSTOMER = [
  { src: '62c3881f-247e-48ad-aacb-977010a23ef1.docx', name: '고객형.docx', limitMs: 2000 },
  { src: '8c76f45a-e15a-40ca-b6be-d9ca5ed63137.pptx', name: '고객형.pptx', limitMs: 3000 },
];
const REPORT = path.resolve('docs/quality-audit/out/artcraft-pages.md');
const SERVER_CONVERT = /^\/api\/view\/(pages|derived-pdf|status|search|pdf-)/;
const results = [];
const rows = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

for (const [file, how] of [['public/artcraft/manifest.json', 'npm run build:artcraft'], ['public/js/artcraftWorker.bundle.js', 'npm run build:editor']]) {
  if (!fs.existsSync(file)) {
    console.error(`${file} 없음 — ${how} 먼저`);
    process.exit(1);
  }
}

let server = null;
async function startServer() {
  fs.rmSync(RUN, { recursive: true, force: true });
  fs.mkdirSync(path.join(RUN, 'tmp'), { recursive: true });
  const log = fs.openSync(path.join(RUN, 'server.log'), 'a');
  server = spawn(process.execPath, ['node_modules/tsx/dist/cli.mjs', 'src/server.ts'], {
    env: { ...process.env, DOCVIEW_PORT: String(PORT), DOCVIEW_HOST: '127.0.0.1', DOCVIEW_SECRET: randomBytes(24).toString('hex'), DOCVIEW_DATA_DIR: path.join(RUN, 'data'), TMP: path.join(RUN, 'tmp'), TEMP: path.join(RUN, 'tmp'), DOCVIEW_LO_PROFILE_ROOT: path.join(RUN, 'lo') },
    stdio: ['ignore', log, log],
  });
  for (let i = 0; i < 120; i += 1) {
    if (await fetch(`${BASE}/api/formats`).then((r) => r.ok, () => false)) return;
    await sleep(500);
  }
  throw new Error('시험 서버가 뜨지 않음');
}
function stopServer() {
  // 내가 띄운 서버만(PID, 자식 포함) — 이름으로 node를 끄지 않는다
  if (server?.pid) execFileSync('taskkill', ['/PID', String(server.pid), '/T', '/F'], { stdio: 'ignore' });
  server = null;
}
/** 서버 기록에서 /artcraft/ 요청 줄 수(브라우저 캐시에서 오면 서버에 요청이 오지 않는다) */
async function artcraftHits() {
  await sleep(500);
  return fs.readFileSync(path.join(RUN, 'server.log'), 'utf8').split('\n').filter((line) => /"url":"\/artcraft\//.test(line)).length;
}

async function upload(file, name) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), name);
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패 ${name}: ${JSON.stringify(r)}`);
  for (let i = 0; i < 1200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready') return st;
    if (st.status === 'failed') throw new Error(`변환 실패 ${name}: ${st.error}`);
    await sleep(250);
  }
  throw new Error(`대기 시간 초과 ${name}`);
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: 'artcraft-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);

async function gray(png) {
  const { data, info } = await sharp(png).flatten({ background: '#ffffff' }).resize({ width: 128 }).grayscale().raw().toBuffer({ resolveWithObject: true });
  return { data, w: info.width, h: info.height };
}
/** 1 − 평균 밝기 차/255(겹치는 높이만) */
function similarity(a, b) {
  const h = Math.min(a.h, b.h);
  let sum = 0;
  for (let i = 0; i < h * a.w; i += 1) sum += Math.abs(a.data[i] - b.data[i]);
  return 1 - sum / (h * a.w * 255);
}
/** 흰 쪽과 LibreOffice 쪽의 유사도 — 흰 바탕이 많아 유사도가 원래 높으므로, 아트크래프트 그림은 이보다 LibreOffice에 가까워야 한다 */
const blankSimilarity = (g) => similarity({ data: new Uint8Array(g.data.length).fill(255), w: g.w, h: g.h }, g);
const inkRatio = (g) => g.data.filter((v) => v < 200).length / g.data.length;
const pct = (v) => `${(v * 100).toFixed(1)}%`;

async function openPage(context, url, { throttle = false, block = false } = {}) {
  const page = await context.newPage();
  const errors = [];
  const paths = [];
  const blocked = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  page.on('request', (r) => paths.push(new URL(r.url()).pathname));
  // Worker 안의 fetch도 이 경로로 잡히는지 blocked로 확인한다(0건이면 (4)가 실패로 알린다)
  if (block)
    await context.route('**/artcraft/**', (r) => {
      blocked.push(new URL(r.request().url()).pathname);
      return r.abort();
    });
  let cdp = null;
  if (throttle) {
    cdp = await context.newCDPSession(page);
    await cdp.send('Emulation.setCPUThrottlingRate', { rate: 4 });
  }
  const t0 = Date.now();
  await page.goto(`${BASE}${url}`);
  await page.waitForSelector('.page-frame[data-page="1"].page-ready', { timeout: 90000 });
  const firstMs = Date.now() - t0;
  if (cdp) await cdp.send('Emulation.setCPUThrottlingRate', { rate: 1 });
  return { page, errors, paths, blocked, firstMs };
}

const cases = [
  { file: 'audit/office/corpus/docx/a1-korean-report.docx', name: 'a1-korean-report.docx', limitMs: 2000, hangul: true },
  { file: 'audit/office/corpus/pptx/a2-korean-deck.pptx', name: 'a2-korean-deck.pptx', limitMs: 2000, hangul: true },
  { file: 'tests/fixtures/legacy/SimpleMacro.doc', name: 'SimpleMacro.doc', limitMs: 2000, hangul: false },
];

await startServer();
const browser = await chromium.launch();
try {
  if (fs.existsSync(ORIGINALS)) {
    fs.mkdirSync(path.join(RUN, 'cust'), { recursive: true });
    for (const c of CUSTOMER) {
      const to = path.join(RUN, 'cust', c.src);
      fs.copyFileSync(path.join(ORIGINALS, c.src), to);
      cases.push({ file: to, name: c.name, limitMs: c.limitMs, hangul: true });
    }
  } else {
    console.log(`SKIP 고객형 docx·pptx — 원본 폴더 없음(ARTCRAFT_E2E_PRIVATE_DIR): ${ORIGINALS}`);
  }

  const ids = {};
  for (const c of cases) {
    const meta = await upload(c.file, c.name);
    ids[c.name] = meta.id;
    const tok = await token(meta.id, { engine: 'artcraft' });
    rec(`${c.name}: 토큰 engine=artcraft`, tok.engine === 'artcraft', tok.engineReason ?? '');
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const before = await artcraftHits();
    const o = await openPage(context, tok.viewerUrl, { throttle: true });

    // (6) CPU×4 첫 쪽
    rec(`${c.name}: (6) CPU×4 첫 쪽 ${o.firstMs} ms ≤ ${c.limitMs}`, o.firstMs <= c.limitMs);

    // (1) 아트크래프트 표지·쪽 수·첫 쪽 그림
    const badge = await o.page.locator('#engine-badge').isVisible();
    const total = Number((await o.page.textContent('#page-jump-total')).replace(/\D/g, ''));
    rec(`${c.name}: (1) 아트크래프트로 열림(표지) · 쪽 수 ${total} (LibreOffice ${meta.pageCount})`, badge && total > 0);
    await sleep(3000); // 두 단계 그리기의 선명한 그림까지
    const dataUrl = await o.page.$eval('.page-frame[data-page="1"] canvas', (cv) => cv.toDataURL('image/png'));
    const mine = await gray(Buffer.from(dataUrl.split(',')[1], 'base64'));
    const loTok = await token(meta.id, {});
    const loPng = Buffer.from(await (await fetch(`${BASE}/api/view/pages/1?token=${encodeURIComponent(loTok.token)}&scale=2`)).arrayBuffer());
    const lo = await gray(loPng);
    const sim = similarity(mine, lo);
    const blank = blankSimilarity(lo);
    const ink = inkRatio(mine);
    const loInk = inkRatio(lo);
    const hangul = await o.page.$$eval('.pdf-text-layer span', (ss) => ss.some((s) => /[가-힣]/.test(s.textContent ?? '')));
    // 판정: 화소 유사도는 흰 바탕이 지배해(흰 쪽도 LibreOffice와 0.84~0.999) 글꼴·줄 높이가 다른 문서에서는 흰 쪽보다 낮게도 나온다 —
    // 그래서 유사도 ≥ 0.8은 "통째로 틀린 그림(검은·뒤집힌·색 깨진 쪽)"만 거르고, 빈·반쯤 그린 쪽은 잉크 비율이 LibreOffice의 0.5~2배인지로 거른다.
    // 한글이 실제로 쪽에 있는지는 엔진 글자 층(.pdf-text-layer)에 한글이 있는지로 본다.
    const inkOk = ink > 0 && ink >= loInk * 0.5 && ink <= loInk * 2;
    rec(
      `${c.name}: (1) 첫 쪽 그림 — 잉크 ${pct(ink)}(LibreOffice ${pct(loInk)}) · 유사도 ${sim.toFixed(3)}(흰 쪽 ${blank.toFixed(3)}) · 한글 글자 층 ${hangul}`,
      inkOk && sim >= 0.8 && (!c.hangul || hangul),
    );

    // (2) 확대·회전·쪽 이동·섬네일·찾기
    // 확대 = 쪽 틀(CSS 폭)이 넓어지고 캔버스 화소 폭이 새 틀 폭×dpr 이상으로 다시 그려짐
    //  (처음 그림은 자리 잡기 전 더 넓은 틀 폭으로 그려질 수 있어 "이전 캔버스보다 n배"로는 재지 않는다)
    const frameW = () => o.page.$eval('.page-frame[data-page="1"]', (f) => f.getBoundingClientRect().width);
    const css0 = await frameW();
    await o.page.click('#zoom-in');
    const zoomed = await o.page
      .waitForFunction(
        (w) => {
          const f = document.querySelector('.page-frame[data-page="1"]');
          const cssW = f?.getBoundingClientRect().width ?? 0;
          return cssW > w * 1.2 && (f?.querySelector('canvas')?.width ?? 0) >= Math.floor(cssW * devicePixelRatio);
        },
        css0,
        { timeout: 30000 },
      )
      .then(() => true, () => false);
    const zoomDetail = await o.page.$eval('.page-frame[data-page="1"]', (f) => `틀 ${Math.round(f.getBoundingClientRect().width)}px · 캔버스 ${f.querySelector('canvas')?.width}px`);
    rec(`${c.name}: (2) 확대 ${await o.page.textContent('#zoom-level')} — 쪽 그림을 새 배율로(틀 ${Math.round(css0)}px → ${zoomDetail})`, zoomed);
    await o.page.click('#rotate-btn');
    rec(`${c.name}: (2) 회전`, await o.page.$eval('#pages', (el) => el.classList.contains('pages-rotated')));
    for (let i = 0; i < 3; i += 1) await o.page.click('#rotate-btn');
    if (total > 1) {
      await o.page.click('#page-next');
      const moved = await o.page
        .waitForFunction(() => document.querySelector('#page-jump-input')?.value === '2' && document.querySelector('.page-frame[data-page="2"]')?.classList.contains('page-ready'), null, { timeout: 30000 })
        .then(() => true, () => false);
      rec(`${c.name}: (2) 쪽 이동 → 2쪽`, moved);
    }
    const thumbs = await o.page
      .waitForFunction(() => [...document.querySelectorAll('#thumbs-panel canvas')].some((cv) => cv.width > 0), null, { timeout: 30000 })
      .then(() => true, () => false);
    rec(`${c.name}: (2) 섬네일`, thumbs);
    const word = await o.page.$$eval('.pdf-text-layer span', (ss) => {
      for (const s of ss) {
        const m = (s.textContent ?? '').match(/[가-힣A-Za-z]{2,}/);
        if (m) return m[0];
      }
      return null;
    });
    if (word) {
      await o.page.click('#search-btn');
      await o.page.fill('#search-input', word);
      await o.page.press('#search-input', 'Enter');
      const found = await o.page
        .waitForFunction(() => /\d+ \/ /.test(document.querySelector('#search-count')?.textContent ?? '') && !!document.querySelector('.search-hit-overlay'), null, { timeout: 30000 })
        .then(() => true, () => false);
      rec(`${c.name}: (2) 찾기 "${word}" — ${await o.page.textContent('#search-count')}`, found);
    } else {
      rec(`${c.name}: (2) 찾기 — 글자 층이 비어 건너뜀`, true);
    }

    // (3) 서버 변환 요청 0건
    const converts = o.paths.filter((p) => SERVER_CONVERT.test(p));
    rec(`${c.name}: (3) 서버 변환 요청 0건`, converts.length === 0, converts.join(', '));
    rec(`${c.name}: 페이지 오류 없음`, o.errors.length === 0, o.errors.join(' | ').slice(0, 300));

    // (7) 두 번째 열기 — 같은 브라우저(캐시)
    await o.page.close();
    const mid = await artcraftHits();
    const again = await openPage(context, tok.viewerUrl);
    const after = await artcraftHits();
    rec(`${c.name}: (7) 두 번째 열기 — wasm·글꼴을 캐시에서(처음 ${mid - before}건, 두 번째 ${after - mid}건)`, mid > before && after === mid);
    await again.page.close();
    await context.close();
    rows.push({ name: c.name, lo: meta.pageCount, art: total, sim, firstMs: o.firstMs });
  }

  // (4) wasm을 막으면 리브레오피스로 자동 전환 + 안내
  {
    const tok = await token(ids['a1-korean-report.docx'], { engine: 'artcraft' });
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const o = await openPage(context, tok.viewerUrl, { block: true });
    const text = await o.page.textContent('#view-notice');
    const badge = await o.page.locator('#engine-badge').isVisible();
    const wasmBlocked = o.blocked.filter((p) => p.endsWith('.wasm'));
    rec(`(4) 경로 막기가 Worker의 wasm 요청을 잡음(${wasmBlocked.length}건)`, wasmBlocked.length > 0, o.blocked.join(', '));
    rec('(4) wasm을 막으면 리브레오피스로 전환 + 안내', /브라우저 엔진이 이 문서를 열지 못해/.test(text ?? '') && !badge, text ?? '');
    await context.close();
  }

  // 열린 뒤 Worker가 죽으면(잡히지 않은 오류) 다음 그리기에서 리브레오피스로 + 안내
  {
    const tok = await token(ids['a1-korean-report.docx'], { engine: 'artcraft' });
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const o = await openPage(context, tok.viewerUrl);
    const worker = o.page.workers().find((w) => w.url().includes('artcraftWorker'));
    if (!worker) {
      rec('열린 뒤 Worker 죽음 — Worker를 찾음', false, o.page.workers().map((w) => w.url()).join(', '));
    } else {
      await worker.evaluate(() =>
        setTimeout(() => {
          throw new Error('e2e: Worker를 일부러 멈춤');
        }),
      );
      await sleep(500);
      const n = o.paths.length;
      await o.page.click('#zoom-in');
      const switched = await o.page
        .waitForFunction(() => /브라우저 엔진이 이 문서를 열지 못해/.test(document.querySelector('#view-notice')?.textContent ?? '') && document.querySelector('#engine-badge')?.hidden, null, { timeout: 30000 })
        .then(() => true, () => false);
      const redrawn = await o.page
        .waitForSelector('.page-frame[data-page="1"].page-ready', { timeout: 60000 })
        .then(() => true, () => false);
      await sleep(1000);
      const served = o.paths.slice(n).filter((p) => SERVER_CONVERT.test(p) || p === '/api/view/download');
      rec('열린 뒤 Worker 죽음 → 리브레오피스로 전환 + 안내 + 1쪽 다시 그림', switched && redrawn && served.length > 0, served.slice(0, 3).join(', '));
    }
    await context.close();
  }

  // (5) 내려받기 권한 없는 토큰은 서버 변환
  {
    const tok = await token(ids['a1-korean-report.docx'], { engine: 'artcraft', canDownload: false });
    rec('(5) 내려받기 권한 없음 → 토큰 engine=libreoffice + 이유', tok.engine === 'libreoffice' && tok.engineReason === '내려받기 권한이 없어 서버 변환으로 엽니다', tok.engineReason ?? '');
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const o = await openPage(context, tok.viewerUrl);
    const asked = o.paths.filter((p) => p.startsWith('/artcraft/') || p.includes('artcraftWorker'));
    const text = await o.page.textContent('#view-notice');
    rec('(5) 서버 변환으로 열림 — 엔진 자산 요청 없음 + 이유 안내', asked.length === 0 && /내려받기 권한이 없어/.test(text ?? ''), asked.join(', '));
    await context.close();
  }

  // 메인 화면 — 단추·공유 폼 라디오
  {
    await upload('tests/fixtures/legacy/SimpleMacro.ppt', 'SimpleMacro.ppt');
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const page = await context.newPage();
    await page.goto(`${BASE}/`);
    const docRow = page.locator('li.doc-item', { hasText: 'a1-korean-report.docx' });
    const pptRow = page.locator('li.doc-item', { hasText: 'SimpleMacro.ppt' });
    await docRow.waitFor({ timeout: 30000 });
    const pptButton = pptRow.getByRole('button', { name: '아트크래프트로 열기' });
    const pptTitle = await pptButton.locator('xpath=..').getAttribute('title');
    rec('메인: .ppt 줄의 "아트크래프트로 열기"는 흐림 + 이유', (await pptButton.isDisabled()) && pptTitle === '이 형식은 브라우저 엔진이 지원하지 않습니다', pptTitle ?? '');
    rec('메인: docx 줄에 "리브레오피스로 열기"', await docRow.getByRole('button', { name: '리브레오피스로 열기' }).isVisible());
    const form = docRow.locator('form.token-form');
    await form.locator('input[name=viewer]').fill('공유받는 사람');
    await form.locator('input[name=canDownload]').check();
    await form.locator('input[name=engine][value=artcraft]').check();
    await form.getByRole('button', { name: '열람 링크 발급' }).click();
    const link = docRow.locator('.token-result a', { hasText: '아트크래프트로 열기' });
    await link.waitFor({ timeout: 15000 });
    const shared = await openPage(context, await link.getAttribute('href'));
    rec('메인: 공유 폼 아트크래프트 링크 → 표지 보임', await shared.page.locator('#engine-badge').isVisible());
    await shared.page.close();
    await docRow.getByRole('button', { name: '아트크래프트로 열기' }).click();
    await page.waitForURL(/\/view\.html\?token=/, { timeout: 15000 });
    await page.waitForSelector('.page-frame[data-page="1"].page-ready', { timeout: 90000 });
    rec('메인: "아트크래프트로 열기" 단추 → 표지 보임', await page.locator('#engine-badge').isVisible());
    await context.close();
  }

  fs.mkdirSync(path.dirname(REPORT), { recursive: true });
  fs.writeFileSync(
    REPORT,
    [
      '# 아트크래프트 쪽 수 비교',
      '',
      `만든 때 ${new Date().toISOString()} · \`scripts/verify-artcraft-e2e.mjs\` · 고객형 문서는 이름 대신 "고객형"으로 적음`,
      '',
      '| 문서 | LibreOffice 쪽 | 아트크래프트 쪽 | 일치 | 첫 쪽 유사도 | 첫 쪽 ms(CPU×4) |',
      '|---|---|---|---|---|---|',
      ...rows.map((r) => `| ${r.name} | ${r.lo} | ${r.art} | ${r.art === r.lo ? '같음' : `다름(${r.art > r.lo ? '+' : ''}${r.art - r.lo})`} | ${r.sim.toFixed(3)} | ${r.firstMs} |`),
      '',
    ].join('\n'),
  );
  console.log(`쪽 수 기록 → ${REPORT}`);
} finally {
  await browser.close();
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed ? 1 : 0);
