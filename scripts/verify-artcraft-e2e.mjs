// 아트크래프트(브라우저 엔진) 보기 — 설계 2026-11-21 §7의 브라우저 검사 일곱 가지 + 메인 화면 단추·열린 뒤 Worker 죽음, 쪽 수 비교 기록.
//  (1) 아트크래프트로 열려 쪽 수·첫 쪽 그림(잉크 비율·LibreOffice 첫 쪽과 줄 잉크 분포 상관·한글 글자 층·한글 span 80 % 이상이 LibreOffice 같은 자리 잉크의 0.3배 이상)
//  (2) 확대·회전·쪽 이동·섬네일·찾기  (3) 서버 변환 요청 0건  (4) wasm을 막으면 리브레오피스로 자동 전환 + 안내
//  (5) 내려받기 권한 없는 토큰은 서버 변환으로  (6) CPU×4 첫 쪽 시간(환산 — CDP 스로틀은 메인 스레드만 늦추고 Worker는 늦추지 않으므로
//      스로틀한 전체 시간 − Worker 시간 + 4 × Worker 시간(Worker가 잰 initMs·openMs·firstRenderMs). 로컬 서버라 받기 시간은 사실상 0 — 목표 docx 2.5초·pptx·doc 2초, 고객 pptx 3초)
//  (7) 두 번째 열기에서 wasm·글꼴을 서버에서 다시 받지 않음(서버 요청 기록으로 센다 — 304 재검증도 요청으로 잡힌다)
// 이 스크립트가 시험 서버를 직접 띄운다(npm run build:editor, npm run build:artcraft 뒤, LibreOffice 필요):
//   ARTCRAFT_E2E_PRIVATE_DIR=<고객형 원본 폴더> node scripts/verify-artcraft-e2e.mjs [포트=8150, 8150~8159]
// 고객형 문서는 그 폴더에서 .audit-run/artcraft-e2e-<포트>/로 복사해 쓴다(저장소에 넣지 않는다). 변수가 없거나 폴더가 없으면 그 둘은 건너뛴다(SKIP 줄).
// 다 통과하면 .audit-run/artcraft-e2e-<포트>/를 지운다(고객형 사본·서버 데이터). 실패하면 살펴보도록 남긴다.
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
const ORIGINALS = process.env.ARTCRAFT_E2E_PRIVATE_DIR ?? '';
const CUSTOMER = [
  { src: '62c3881f-247e-48ad-aacb-977010a23ef1.docx', name: '고객형.docx', limitMs: 2500 },
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
  const pid = server?.pid;
  server = null;
  if (pid) {
    try {
      execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' });
    } catch {
      // 이미 끝났다
    }
  }
}
// Ctrl+C·예외로 끝나도 시험 서버를 남기지 않는다
process.on('SIGINT', () => {
  stopServer();
  process.exit(130);
});
process.on('exit', stopServer);
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
/** 줄이지 않은 회색조 — 글자 상자(span) 단위로 견줄 때 */
async function grayFull(png) {
  const { data, info } = await sharp(png).flatten({ background: '#ffffff' }).grayscale().raw().toBuffer({ resolveWithObject: true });
  return { data, w: info.width, h: info.height };
}
/** 쪽 비율 상자(fx·fy·fw·fh) 안 어두운(< 128) 화소 비율 */
function darkRatioIn(g, b) {
  const x0 = Math.max(0, Math.floor(b.fx * g.w));
  const y0 = Math.max(0, Math.floor(b.fy * g.h));
  const x1 = Math.min(g.w, Math.ceil((b.fx + b.fw) * g.w));
  const y1 = Math.min(g.h, Math.ceil((b.fy + b.fh) * g.h));
  let dark = 0;
  for (let y = y0; y < y1; y += 1) for (let x = x0; x < x1; x += 1) if (g.data[y * g.w + x] < 128) dark += 1;
  return (x1 - x0) * (y1 - y0) > 0 ? dark / ((x1 - x0) * (y1 - y0)) : 0;
}
// 자체 점검 — 왼쪽 절반이 검은 2×1 그림
if (darkRatioIn({ data: Uint8Array.from([0, 255]), w: 2, h: 1 }, { fx: 0, fy: 0, fw: 0.5, fh: 1 }) !== 1) throw new Error('darkRatioIn 자체 점검 실패');
/** 화소 줄마다 잉크 비율(밝기 < 200) — 흰 바탕과 상관없이 줄·문단 배치를 본다 */
const rowInk = (g) => Array.from({ length: g.h }, (_, y) => g.data.subarray(y * g.w, (y + 1) * g.w).filter((v) => v < 200).length / g.w);
/** 쪽 높이를 n띠로 — 띠 b는 LibreOffice 쪽의 [b/n, (b+1)/n], 아트크래프트 쪽은 그 위치를 stretch로 나눈 곳 */
function bands(rows, n, stretch = 1) {
  return Array.from({ length: n }, (_, b) => {
    const y0 = Math.floor((b / n / stretch) * rows.length);
    const y1 = Math.max(y0 + 1, Math.floor(((b + 1) / n / stretch) * rows.length));
    let sum = 0;
    for (let y = y0; y < y1; y += 1) sum += rows[y] ?? 0;
    return sum / (y1 - y0);
  });
}
/** 피어슨 상관(한쪽이 고르면 0) */
function pearson(a, b) {
  const n = a.length;
  const ma = a.reduce((x, y) => x + y, 0) / n;
  const mb = b.reduce((x, y) => x + y, 0) / n;
  let sab = 0;
  let saa = 0;
  let sbb = 0;
  for (let i = 0; i < n; i += 1) {
    sab += (a[i] - ma) * (b[i] - mb);
    saa += (a[i] - ma) ** 2;
    sbb += (b[i] - mb) ** 2;
  }
  return saa && sbb ? sab / Math.sqrt(saa * sbb) : 0;
}
/** 줄 잉크 분포 상관 — 줄 간격이 엔진마다 달라(LibreOffice가 1.1~1.2배 넓다) 아래로 갈수록 줄이 밀리므로 세로 늘이기 0.7~1.4배 중 가장 높은 값.
 *  재 본 값(2026-11-21): 같은 문서 0.665~1.0, 다른 문서끼리 최대 0.587 → 기준 0.6 */
function layoutCorr(mine, lo, n = 64) {
  const m = rowInk(mine);
  const l = bands(rowInk(lo), n);
  let best = -1;
  for (let s = 0.7; s <= 1.4001; s += 0.02) best = Math.max(best, pearson(bands(m, n, s), l));
  return best;
}
// 자체 점검 — 같은 배치면 1, 뒤집힌 배치면 음수
if (!(pearson([0, 1, 0, 1], [0, 2, 0, 2]) > 0.99 && pearson([0, 1, 0, 1], [1, 0, 1, 0]) < 0)) throw new Error('pearson 자체 점검 실패');
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
  // Worker 안에서 잰 시간(view.ts가 표지에 적는다) — 아트크래프트로 열리지 않았으면 null
  const timing = await page.$eval('#engine-badge', (b) => (b.dataset.workerTiming ? JSON.parse(b.dataset.workerTiming) : null)).catch(() => null);
  return { page, errors, paths, blocked, firstMs, timing };
}

const cases = [
  { file: 'audit/office/corpus/docx/a1-korean-report.docx', name: 'a1-korean-report.docx', limitMs: 2500, hangul: true },
  { file: 'audit/office/corpus/pptx/a2-korean-deck.pptx', name: 'a2-korean-deck.pptx', limitMs: 2000, hangul: true },
  { file: 'tests/fixtures/legacy/SimpleMacro.doc', name: 'SimpleMacro.doc', limitMs: 2000, hangul: false },
];

let browser = null;
let finished = false;
try {
  await startServer();
  browser = await chromium.launch();
  if (ORIGINALS && fs.existsSync(ORIGINALS)) {
    fs.mkdirSync(path.join(RUN, 'cust'), { recursive: true });
    for (const c of CUSTOMER) {
      const to = path.join(RUN, 'cust', c.src);
      fs.copyFileSync(path.join(ORIGINALS, c.src), to);
      cases.push({ file: to, name: c.name, limitMs: c.limitMs, hangul: true });
    }
  } else {
    console.log(`SKIP 고객형 docx·pptx — ARTCRAFT_E2E_PRIVATE_DIR ${ORIGINALS ? `폴더 없음: ${ORIGINALS}` : '변수 없음'}`);
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

    // (6) CPU×4(환산) 첫 쪽 — 메인 스레드 몫(스로틀됨) + 4 × Worker 몫
    const workerMs = o.timing ? o.timing.initMs + o.timing.openMs + o.timing.firstRenderMs : null;
    const mainMs = workerMs === null ? null : Math.max(0, o.firstMs - workerMs);
    const cpu4Ms = workerMs === null ? null : mainMs + 4 * workerMs;
    rec(
      `${c.name}: (6) CPU×4(환산) 첫 쪽 ${cpu4Ms} ms ≤ ${c.limitMs} (메인 ${mainMs} + 4 × Worker ${workerMs}${o.timing ? ` = init ${o.timing.initMs}·open ${o.timing.openMs}·첫 그림 ${o.timing.firstRenderMs}` : ''}, 스로틀 실측 ${o.firstMs})`,
      cpu4Ms !== null && cpu4Ms <= c.limitMs,
    );

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
    // 실패하면 살펴보도록 시험 폴더에 남긴다(통과하면 폴더째 지운다)
    fs.writeFileSync(path.join(RUN, `${c.name}-artcraft.png`), Buffer.from(dataUrl.split(',')[1], 'base64'));
    fs.writeFileSync(path.join(RUN, `${c.name}-libreoffice.png`), loPng);
    const sim = layoutCorr(mine, lo);
    const ink = inkRatio(mine);
    const loInk = inkRatio(lo);
    // 한글 글자 층 span 상자마다 캔버스 화소 중 어두운 것(휘도 < 128)의 비율 — 상자는 쪽 비율로 돌려주고 LibreOffice 같은 자리와 견준다
    const glyph = await o.page.evaluate(() => {
      const frame = document.querySelector('.page-frame[data-page="1"]');
      const cv = frame?.querySelector('canvas');
      const spans = [...(frame?.querySelectorAll('.pdf-text-layer span') ?? [])].filter((s) => /[가-힣]/.test(s.textContent ?? ''));
      if (!cv || !spans.length) return { runs: spans.length, boxes: [] };
      const ctx = cv.getContext('2d');
      const r = cv.getBoundingClientRect();
      const k = cv.width / r.width;
      const boxes = [];
      for (const s of spans) {
        const b = s.getBoundingClientRect();
        const x = Math.max(0, Math.floor((b.left - r.left) * k));
        const y = Math.max(0, Math.floor((b.top - r.top) * k));
        const w = Math.min(cv.width - x, Math.ceil(b.width * k));
        const h = Math.min(cv.height - y, Math.ceil(b.height * k));
        if (w <= 0 || h <= 0) continue;
        const d = ctx.getImageData(x, y, w, h).data;
        let dark = 0;
        for (let i = 0; i < d.length; i += 4) if (0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2] < 128) dark += 1;
        boxes.push({ fx: x / cv.width, fy: y / cv.height, fw: w / cv.width, fh: h / cv.height, dark: dark / (w * h) });
      }
      return { runs: spans.length, boxes };
    });
    const hangul = glyph.runs > 0;
    // span 단위(최종 재검토 N3): 상자 넓이 대비 어두운 비율이 LibreOffice 같은 자리의 0.3배 이상인 span이 80 % 이상 — 두부 상자·밑줄·표 선만으로는 못 넘는다
    const loFull = await grayFull(loPng);
    const spanOk = glyph.boxes.filter((b) => b.dark >= 0.3 * darkRatioIn(loFull, b)).length;
    const spanShare = glyph.boxes.length ? spanOk / glyph.boxes.length : 0;
    // 판정: 줄 잉크 분포(64띠, 세로 늘이기 허용) 상관 ≥ 0.6 — 흰 바탕에 좌우되지 않고 줄·문단 배치가 LibreOffice와 닮았는지(빈·뒤섞인·다른 쪽은 낮다).
    // 잉크 비율은 LibreOffice의 0.5~2배(빈·반쯤 그린·검은 쪽 거르기). 한글은 글자 층에 있고(엔진이 꺼냄) 그 자리에 잉크가 LibreOffice만큼 있어야(실제로 그림) 한다.
    const inkOk = ink > 0 && ink >= loInk * 0.5 && ink <= loInk * 2;
    rec(
      `${c.name}: (1) 첫 쪽 그림 — 잉크 ${pct(ink)}(LibreOffice ${pct(loInk)}) · 줄 분포 상관 ${sim.toFixed(3)} · 한글 글자 층 ${glyph.runs}개 · LibreOffice 자리 잉크의 0.3배 이상 span ${spanOk}/${glyph.boxes.length}(${pct(spanShare)})`,
      inkOk && sim >= 0.6 && (!c.hangul || (hangul && spanShare >= 0.8)),
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
    rows.push({ name: c.name, lo: meta.pageCount, art: total, sim, firstMs: o.firstMs, workerMs, cpu4Ms });
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
      // 요청 세기는 죽이기 전부터 — 선명하게 그리기가 진행 중이면 확대를 누르기 전에 이미 전환한다(onDead)
      const n = o.paths.length;
      await worker.evaluate(() =>
        setTimeout(() => {
          throw new Error('e2e: Worker를 일부러 멈춤');
        }),
      );
      await sleep(500);
      await o.page.click('#zoom-in');
      const switched = await o.page
        .waitForFunction(() => /브라우저 엔진이 이 문서를 열지 못해/.test(document.querySelector('#view-notice')?.textContent ?? '') && document.querySelector('#engine-badge')?.hidden, null, { timeout: 30000 })
        .then(() => true, () => false);
      const redrawn = await o.page
        .waitForSelector('.page-frame[data-page="1"].page-ready', { timeout: 60000 })
        .then(() => true, () => false);
      await sleep(1000);
      const served = o.paths.slice(n).filter((p) => SERVER_CONVERT.test(p) || p === '/api/view/download');
      rec('열린 뒤 Worker 죽음 → 리브레오피스로 전환 + 안내 + 1쪽 다시 그림', switched && redrawn && served.length > 0, `전환 ${switched} · 다시 그림 ${redrawn} · ${served.slice(0, 3).join(', ')}`);
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
      '| 문서 | LibreOffice 쪽 | 아트크래프트 쪽 | 일치 | 첫 쪽 줄 분포 상관 | 첫 쪽 ms(메인 스레드만 ×4 실측) | Worker ms | 첫 쪽 ms(CPU×4 환산) |',
      '|---|---|---|---|---|---|---|---|',
      ...rows.map((r) => `| ${r.name} | ${r.lo} | ${r.art} | ${r.art === r.lo ? '같음' : `다름(${r.art > r.lo ? '+' : ''}${r.art - r.lo})`} | ${r.sim.toFixed(3)} | ${r.firstMs} | ${r.workerMs} | ${r.cpu4Ms} |`),
      '',
    ].join('\n'),
  );
  console.log(`쪽 수 기록 → ${REPORT}`);
  finished = true;
} finally {
  await browser?.close().catch(() => undefined);
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
if (finished && !failed) {
  await sleep(500); // 끈 서버가 파일을 놓을 때까지
  fs.rmSync(RUN, { recursive: true, force: true });
} else {
  console.log(`시험 폴더를 살펴보도록 남김(고객형 사본·서버 데이터 포함 — 다 보면 지울 것): ${RUN}`);
}
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed ? 1 : 0);
