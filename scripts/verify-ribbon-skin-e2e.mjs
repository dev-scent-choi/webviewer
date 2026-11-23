#!/usr/bin/env node
// 아트크래프트 리본 꾸밈 검사(설계 2026-11-22-artcraft-ribbon-skin §4) — 세 편집기(편집·읽기 전용 ui=view)와 보기 화면(view.html):
//  (1) 제목 줄 24·탭 줄 22·리본 62·상태 줄 18px(±2, 캡처 실측)  (2) 탭·묶음·단추 이름·말풍선이 한글(영문자만 있는 라벨 0)
//  (3) 큰 단추 = 아이콘 위·이름 아래  (4) 모든 탭의 리본 내용이 62px 안(잘림 없음)  (5) "자세히" 수는 목록 기준, 준비 중은 꺼짐·나머지는 있는 단추를 가리킴
//  (6) 읽기 전용 흐림이 고르다(0.45), "자세히" 흐림 = 가리키는 단추 흐림  (7) 600px 가로 스크롤 없음·제목 줄 한 줄
//  (8) 보기 화면 같은 규격·id 그대로·첫 쪽이 막대 밑에 깔리지 않음  (9) 아이콘 큰 30px·작은 15px 이상, 보기 화면 단추 이름, 상태 표시(읽기 전용·연결됨·편집 열기/편집) 안 보임
//  (10) 워드 묶음 접기 — 축소 리본 엿보기·다시 고정 때 접힘, 같은 폭 다시 맞춤은 열린 묶음 판 유지
//  + 아트크래프트 캡처와 나란히 놓은 보고서
// 쓰는 법: npm run build:editor && node scripts/verify-ribbon-skin-e2e.mjs [포트 8160~8169, 기본 8166]
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { chromium } from 'playwright';
import { SKIN_PX, SKIN_TOLERANCE_PX, bigButtonLayoutOk, heightProblems, latinOnlyLabels } from './ribbonSkinLib.mjs';
import { APP_OF_EXT, LISTED_EXTS as LISTED, launcherCount, loadInventory } from './ribbonGenLib.mjs';

const PORT = Number(process.argv[2] ?? 8166);
// 운영 포트(8000·8010·8020·8100·8110·8130·8148)와 다른 시험(8150~8159) 범위와 겹치지 않게
if (!((PORT >= 8160 && PORT <= 8169) || (PORT >= 8180 && PORT <= 8189))) throw new Error(`시험 포트는 8160~8169·8180~8189만 쓴다(받은 값 ${PORT})`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `ribbon-skin-e2e-${PORT}`);
const SHOTS = path.resolve('.audit-run', 'ribbon-skin-shots');
const REF_DIR = process.env.ARTCRAFT_SHOTS_DIR ?? 'C:\\Users\\OpenMinds\\projects\\artcraft-eval\\spike-webapp\\shots';
const FILES = {
  docx: 'audit/office/corpus/docx/a1-korean-report.docx',
  xlsx: 'audit/office/corpus/xlsx/x07-multisheet-print.xlsx',
  pptx: 'audit/office/corpus/pptx/a2-korean-deck.pptx',
};
const READY = { docx: '.ProseMirror', xlsx: '#grid-edit-table td[data-row]', pptx: '#slides-rail .slide-card' };
const REF = { docx: 'word-report.png', xlsx: 'grid-plain.png', pptx: 'deck-korean.png' };
// "자세히" 수는 오피스 리본 목록에서(상황별 탭 판 포함, 준비 중 포함) — 아직 목록으로 만들지 않은 앱(LISTED_EXTS 밖)은 옛 수
const INV = loadInventory();
const LAUNCHERS = Object.fromEntries(Object.entries(APP_OF_EXT).map(([ext, name]) => [ext, LISTED.has(ext) ? launcherCount(INV.apps[name]) : { xlsx: 1, pptx: 0 }[ext]]));
const EDITOR_SEL = { title: '.office-titlebar', tabs: '.ribbon-tabs', ribbon: '.ribbon-panel:not([hidden])', status: '#statusbar, .grid-status-row' };
const VIEW_SEL = { title: '.toolbar-row-title', tabs: '.ribbon-tabs', ribbon: '.toolbar-row-tools', status: null };
const VIEW_PAGE_IDS = ['back-btn', 'doc-title', 'edit-btn', 'thumbs-btn', 'measure-btn', 'invert-btn', 'markup-add-btn', 'text-tool-btn', 'markup-include-wrap', 'zoom-out', 'zoom-in', 'zoom-fit-btn', 'zoom-level', 'page-prev', 'page-jump-input', 'page-next', 'rotate-btn', 'search-btn', 'print-btn', 'download-btn', 'download-as-btn', 'fullscreen-btn', 'tools-btn', 'tools-menu'];
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

if (!fs.existsSync('public/js/editor.bundle.js')) {
  console.error('public/js/editor.bundle.js 없음 — npm run build:editor 먼저');
  process.exit(1);
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
process.on('SIGINT', () => {
  stopServer();
  process.exit(130);
});
process.on('exit', stopServer);

async function upload(bytes, name) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
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
    body: JSON.stringify({ viewer: 'ribbon-skin-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);
async function smallPdf() {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  doc.addPage([400, 300]).drawText('ribbon skin e2e pdf', { x: 40, y: 150, size: 20, font });
  return Buffer.from(await doc.save());
}

async function open(context, url, ready) {
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.goto(`${BASE}${url}`);
  await page.waitForSelector(ready, { timeout: 90000 });
  await page.waitForFunction(() => !document.querySelector('#conn-status') || document.querySelector('#conn-status').getAttribute('data-status') === 'synced', null, { timeout: 30000 }).catch(() => {});
  await sleep(1000);
  return { page, errors };
}
const heights = (page, sel) =>
  page.evaluate((s) => {
    const h = (q) => {
      const el = q && document.querySelector(q);
      return el && el.getClientRects().length > 0 ? Math.round(el.getBoundingClientRect().height) : null;
    };
    return { title: h(s.title), tabs: h(s.tabs), ribbon: h(s.ribbon), status: h(s.status) };
  }, sel);
/** 숨은 탭·판의 것까지 — 글자는 보이는지와 상관없이 한글이어야 한다 */
const labels = (page) =>
  page.evaluate(() => {
    const text = [...document.querySelectorAll('.ribbon-tab, .ribbon-group-label, .icon-btn-big > span, .ribbon-btn-label, .statusbar-item')].map((el) => el.textContent ?? '');
    const tips = [...document.querySelectorAll('#toolbar button, #toolbar select, #toolbar input, #toolbar a')].flatMap((el) => [el.getAttribute('title') ?? '', el.getAttribute('aria-label') ?? '']);
    return [...text, ...tips];
  });
const bigButtons = (page) =>
  page.evaluate(() => {
    const r = (el) => {
      const x = el.getBoundingClientRect();
      return { left: x.left, top: x.top, width: x.width, height: x.height };
    };
    return [...document.querySelectorAll('.ribbon-panel:not([hidden]) .icon-btn-big')]
      .filter((b) => b.getClientRects().length > 0)
      .map((b) => {
        const svg = b.querySelector(':scope > svg');
        const span = b.querySelector(':scope > span');
        return svg && span && span.getClientRects().length > 0 ? { name: b.id || b.title, icon: r(svg), label: r(span) } : null;
      })
      .filter(Boolean);
  });
/** 아이콘 크기(아트크래프트 기준 큰 32·작은 16px, 다듬기 2026-11-22)와 숨겨야 하는 상태 표시 — 보이는 것만 잰다.
 *  큰: 편집기 큰 단추·보기 화면 단추의 아이콘, 작은: 편집기 작은 단추 첫 아이콘(펼침 화살표 단추 제외) */
const iconChrome = (page) =>
  page.evaluate(() => {
    const vis = (el) => el.getClientRects().length > 0;
    const side = (el) => Math.round(Math.min(el.getBoundingClientRect().width, el.getBoundingClientRect().height));
    const min = (sel) => {
      const list = [...document.querySelectorAll(sel)].filter(vis).map(side);
      return list.length ? Math.min(...list) : null;
    };
    return {
      big: min('.ribbon-panel:not([hidden]) .icon-btn-big > svg, .toolbar-row-tools .ribbon-group-buttons > .icon-btn > svg:not([hidden]):not(.icon-btn-badge):not(.icon-btn-caret-glyph)'),
      small: min('.ribbon-panel:not([hidden]) .icon-btn:not(.icon-btn-big):not(.dropdown-caret) > svg:first-child'),
      chrome: [...document.querySelectorAll('#readonly-badge, .readonly-badge, .conn-status, .view-open-edit, #edit-btn, .view-readonly')].filter(vis).map((el) => el.id || el.className),
      names: [...document.querySelectorAll('.toolbar-row-tools .ribbon-group-buttons > .icon-btn')].filter(vis).filter((b) => !b.querySelector(':scope > .tool-name')?.textContent?.trim() || !vis(b.querySelector(':scope > .tool-name'))).map((b) => b.id),
    };
  });
/** 보이는 탭을 하나씩 눌러 판 높이(62±2)·잘림(scrollHeight > clientHeight)·판 안 스크롤바·묶음 이름 줄과 겹친 조작을 본다 */
async function tabPanels(page) {
  const tabs = await page.$$eval('.ribbon-tab:not(.ribbon-tab-file)', (ts) => ts.filter((t) => !t.hidden && t.dataset.tab).map((t) => t.dataset.tab));
  const bad = [];
  for (const tab of tabs) {
    await page.click(`.ribbon-tab[data-tab="${tab}"]`);
    const m = await page.evaluate(panelState, '.ribbon-panel:not([hidden])');
    if (!m || Math.abs(m.h - SKIN_PX.ribbon) > SKIN_TOLERANCE_PX || m.sh > m.ch + 1 || m.bar > 0 || m.clash.length > 0) bad.push(`${tab} ${JSON.stringify(m)}`);
  }
  await page.click('.ribbon-tab[data-tab="home"]');
  return { tabs, bad };
}
/** 판 하나의 상태 — 높이, 가로 스크롤바(offsetHeight − clientHeight), 묶음 이름 줄 위 변보다 1px 넘게 내려온 조작(자세히·펼침 메뉴 제외).
 *  스타일 갤러리는 안 항목 대신 갤러리 상자 자체와 그 스크롤바를 본다(넘친 항목은 갤러리 안에서 잘려 보이지 않음) */
function panelState(sel) {
  const p = document.querySelector(sel);
  if (!p) return null;
  const clash = [...p.querySelectorAll('.ribbon-group')].flatMap((g) => {
    const lab = g.querySelector(':scope > .ribbon-group-label');
    if (!lab?.getClientRects().length) return [];
    const top = lab.getBoundingClientRect().top;
    return [...g.querySelectorAll('button:not(.ribbon-launcher), select, input, .style-gallery')]
      .filter((c) => c.getClientRects().length && !c.closest('.dropdown-menu') && !c.parentElement.closest('.style-gallery') && c.getBoundingClientRect().bottom > top + 1)
      .map((c) => c.id || c.title || c.className);
  });
  const gallery = [...p.querySelectorAll('.style-gallery')].filter((g) => g.getClientRects().length && g.offsetHeight - g.clientHeight > 0).map((g) => g.id || 'style-gallery');
  return { h: Math.round(p.getBoundingClientRect().height), sh: p.scrollHeight, ch: p.clientHeight, bar: p.offsetHeight - p.clientHeight, clash: [...clash, ...gallery.map((g) => `${g} 스크롤바`)] };
}
/** 600px — 문서·리본 가로 넘침, 제목 줄 높이(한 줄) */
const narrowState = (page) =>
  page.evaluate(() => {
    const p = document.querySelector('.ribbon-panel:not([hidden]), .toolbar-row-tools');
    const t = document.querySelector('.office-titlebar, .toolbar-row-title');
    return { doc: document.documentElement.scrollWidth, body: document.body.scrollWidth, panel: p ? p.scrollWidth - p.clientWidth : 0, title: Math.round(t.getBoundingClientRect().height) };
  });
const narrowOk = (n) => n.doc <= 600 && n.body <= 600 && n.panel <= 1 && Math.abs(n.title - SKIN_PX.title) <= SKIN_TOLERANCE_PX;
const underToolbar = (page) => page.evaluate(() => document.querySelector('.page-frame').getBoundingClientRect().top < document.getElementById('toolbar').getBoundingClientRect().bottom);
async function shot(page, name) {
  await page.screenshot({ path: path.join(SHOTS, `${name}.png`) });
}
function writeReport(pairs) {
  const rows = pairs.map(([name, ref]) => {
    const src = path.join(REF_DIR, ref);
    const has = fs.existsSync(src);
    if (has) fs.copyFileSync(src, path.join(SHOTS, `ref-${ref}`));
    return `<tr><td><img src="${name}.png" alt="${name}"><p>${name}</p></td><td>${has ? `<img src="ref-${ref}" alt="${ref}">` : '(참고 캡처 없음)'}<p>${ref}</p></td></tr>`;
  });
  fs.writeFileSync(
    path.join(SHOTS, 'report.html'),
    `<!doctype html><html lang="ko"><meta charset="utf-8"><title>리본 꾸밈 비교</title><style>body{font-family:"맑은 고딕",sans-serif}img{width:680px;border:1px solid #ccc}td{vertical-align:top;padding:8px}</style><h1>리본 꾸밈 — 우리 화면과 아트크래프트 나란히</h1><table><tr><th>우리 화면(1400×900)</th><th>아트크래프트</th></tr>${rows.join('')}</table></html>`,
  );
}

await startServer();
fs.rmSync(SHOTS, { recursive: true, force: true });
fs.mkdirSync(SHOTS, { recursive: true });
const pairs = [];
let browser;
try {
  const ids = {};
  for (const [ext, file] of Object.entries(FILES)) ids[ext] = (await upload(fs.readFileSync(file), path.basename(file))).id;
  ids.pdf = (await upload(await smallPdf(), 'ribbon-skin.pdf')).id;
  // 실제 Windows Chrome처럼 스크롤바를 보이게 — 숨기면 62px 판 안을 덮는 스크롤바를 못 잡는다
  browser = await chromium.launch({ ignoreDefaultArgs: ['--hide-scrollbars'] });

  for (const ext of ['docx', 'xlsx', 'pptx']) {
    // ---- 편집 화면 ----
    const wide = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const edit = await open(wide, (await token(ids[ext], { canEdit: true })).editorUrl, READY[ext]);
    const h = await heights(edit.page, EDITOR_SEL);
    rec(`(1) ${ext} 편집 — 높이 24/22/62/18(±2)`, heightProblems(h).length === 0 && h.status !== null, `${JSON.stringify(h)} ${heightProblems(h).join(', ')}`);
    const latin = latinOnlyLabels(await labels(edit.page));
    rec(`(2) ${ext} 탭·묶음·단추 이름·말풍선이 한글`, latin.length === 0, latin.slice(0, 5).join(' | '));
    const big = await bigButtons(edit.page);
    const bigBad = big.filter((b) => !bigButtonLayoutOk(b.icon, b.label)).map((b) => b.name);
    rec(`(3) ${ext} 홈 탭 큰 단추 ${big.length}개 — 아이콘 위·이름 아래`, big.length > 0 && bigBad.length === 0, bigBad.join(', '));
    const ic = await iconChrome(edit.page);
    rec(`(9) ${ext} 홈 탭 아이콘 — 큰 단추 30px 이상·작은 단추 15px 이상(아트크래프트 32·16)`, ic.big >= 30 && ic.small >= 15, JSON.stringify(ic));
    const panels = await tabPanels(edit.page);
    rec(`(4) ${ext} 탭 ${panels.tabs.length}개 모두 리본 62px 안(잘림 없음)`, panels.tabs.length >= 5 && panels.bad.length === 0, panels.bad.join(' | '));
    const launchers = await edit.page.$$eval('.ribbon-launcher', (ls) => ls.map((l) => ({ proxy: l.dataset.proxy ?? '', pending: l.hasAttribute('data-pending'), ok: !l.id && (l.hasAttribute('data-pending') ? l.disabled : !!document.getElementById(l.dataset.proxy ?? '')) })));
    rec(`(5) ${ext} "자세히" ${LAUNCHERS[ext]}개(목록 기준) — 준비 중은 꺼짐, 나머지는 기존 단추를 가리킴, id 없음`, launchers.length === LAUNCHERS[ext] && launchers.every((l) => l.ok), JSON.stringify(launchers.filter((l) => !l.ok)));
    rec(`${ext} 편집 — 화면 오류 없음`, edit.errors.length === 0, edit.errors.slice(0, 2).join(' | '));
    await shot(edit.page, `editor-${ext}`);
    pairs.push([`editor-${ext}`, REF[ext]]);
    await wide.close();

    if (ext === 'docx') {
      // ---- 묶음 접기(오피스 리본 구조 Task 2 고침 I1·M2): 홈 판 묶음에 data-collapse·접힘 단추를 달고 오른쪽 묶음을 넓혀 넘치게 한다 ----
      const fold = await browser.newContext({ viewport: { width: 1000, height: 900 } });
      const f = await open(fold, (await token(ids[ext], { canEdit: true })).editorUrl, READY[ext]);
      await f.page.evaluate(() => {
        const groups = [...document.querySelectorAll('.ribbon-panel[data-panel="home"] > .ribbon-group')];
        for (const g of groups) {
          g.dataset.collapse = '1';
          g.insertAdjacentHTML('afterbegin', '<button type="button" class="icon-btn icon-btn-big rb-group-toggle" aria-expanded="false"><span>묶음</span></button>');
        }
        groups.at(-1).querySelector('.ribbon-group-buttons').insertAdjacentHTML('beforeend', '<span style="display:inline-block;width:1600px"></span>');
      });
      const foldState = () => f.page.evaluate(() => {
        const p = document.querySelector('.ribbon-panel[data-panel="home"]');
        return { shown: p.getClientRects().length > 0, folded: p.querySelectorAll(':scope > .rb-collapsed').length, over: p.scrollWidth - p.clientWidth, open: p.querySelectorAll('.rb-open').length };
      });
      const unfold = () => f.page.evaluate(() => document.querySelectorAll('.rb-collapsed').forEach((g) => g.classList.remove('rb-collapsed')));
      const foldOk = (s) => s.shown && s.folded > 0 && s.over <= 1;
      await f.page.click('#ribbon-collapse'); // 리본 축소
      await unfold();
      await f.page.click('.ribbon-tab[data-tab="home"]'); // 탭을 눌러 엿보기
      const peek = await foldState();
      rec('(10) docx 축소한 리본을 탭으로 펼쳐 보면 넘치는 묶음이 접힘', foldOk(peek), JSON.stringify(peek));
      await f.page.mouse.click(700, 600); // 리본 밖 — 다시 축소
      await unfold();
      await f.page.click('#ribbon-collapse'); // 다시 고정
      const pin = await foldState();
      rec('(10) docx 리본을 다시 고정하면 넘치는 묶음이 접힘', foldOk(pin), JSON.stringify(pin));
      await f.page.click('.ribbon-panel[data-panel="home"] > .rb-collapsed > .rb-group-toggle');
      await f.page.evaluate(() => window.dispatchEvent(new Event('resize')));
      await sleep(300);
      const same = await foldState();
      await f.page.setViewportSize({ width: 950, height: 900 });
      await sleep(300);
      const changed = await foldState();
      rec('(10) docx 접힌 묶음 판 — 같은 폭으로 다시 맞춰도 열림 유지, 폭이 바뀌면 닫힘', same.open === 1 && changed.open === 0, JSON.stringify({ same, changed }));
      rec('docx 묶음 접기 — 화면 오류 없음', f.errors.length === 0, f.errors.slice(0, 2).join(' | '));
      await fold.close();
    }

    // ---- 좁은 창(900·800px): 판 안 스크롤바·묶음 이름 줄 겹침 ----
    for (const width of [900, 800]) {
      const mid = await browser.newContext({ viewport: { width, height: 900 } });
      const m = await open(mid, (await token(ids[ext], { canEdit: true })).editorUrl, READY[ext]);
      const mp = await tabPanels(m.page);
      rec(`(4) ${ext} ${width}px — 탭 ${mp.tabs.length}개 모두 62px·스크롤바 없음·이름 줄 겹침 없음`, mp.tabs.length >= 5 && mp.bad.length === 0, mp.bad.join(' | '));
      await mid.close();
    }

    // ---- 읽기 전용(ui=view) ----
    const roCtx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    // 메인 화면에서 연 것처럼 — "편집 열기" 단추가 DOM에 생겨야 숨김을 잴 수 있다
    await roCtx.addInitScript((id) => sessionStorage.setItem('docview-edit-doc', id), ids[ext]);
    const ro = await open(roCtx, (await token(ids[ext], { engine: 'editor' })).viewerUrl, READY[ext]);
    const rh = await heights(ro.page, EDITOR_SEL);
    rec(`(1) ${ext} 읽기 전용 — 같은 높이 규격`, heightProblems(rh).length === 0, `${JSON.stringify(rh)} ${heightProblems(rh).join(', ')}`);
    const grey = await ro.page.evaluate(() => {
      const vis = [...document.querySelectorAll('.ribbon-panel:not([hidden]) button.ro-disabled, .quick-access button.ro-disabled')].filter((b) => b.getClientRects().length > 0);
      const odd = vis.filter((b) => getComputedStyle(b).opacity !== '0.45').map((b) => b.id || b.title);
      const launch = [...document.querySelectorAll('.ribbon-launcher')]
        .filter((l) => l.classList.contains('ro-disabled') !== !!document.getElementById(l.dataset.proxy ?? '')?.classList.contains('ro-disabled'))
        .map((l) => l.dataset.proxy);
      return { count: vis.length, odd, launch, body: document.body.classList.contains('ui-view') };
    });
    rec(`(6) ${ext} 읽기 전용 — 흐린 단추 ${grey.count}개 모두 0.45, "자세히" 흐림 = 대상 흐림`, grey.body && grey.count > 0 && grey.odd.length === 0 && grey.launch.length === 0, JSON.stringify(grey));
    const ric = await iconChrome(ro.page);
    const roEdit = await ro.page.evaluate(() => document.querySelectorAll('.view-open-edit').length);
    rec(`(9) ${ext} 읽기 전용 — 읽기 전용 표지·연결됨·편집 열기·상태 줄 읽기 전용 안 보임(편집 열기는 DOM에 있음), 아이콘 32·16`, ric.chrome.length === 0 && roEdit === 1 && ric.big >= 30 && ric.small >= 15, JSON.stringify({ ...ric, roEdit }));
    await shot(ro.page, `readonly-${ext}`);
    pairs.push([`readonly-${ext}`, REF[ext]]);
    await roCtx.close();

    // ---- 600px (편집·읽기 전용) ----
    for (const [mode, body, key] of [
      ['편집', { canEdit: true }, 'editorUrl'],
      ['읽기 전용', { engine: 'editor' }, 'viewerUrl'],
    ]) {
      const narrow = await browser.newContext({ viewport: { width: 600, height: 900 } });
      const n = await open(narrow, (await token(ids[ext], body))[key], READY[ext]);
      const tabs = await n.page.$$eval('.ribbon-tab:not(.ribbon-tab-file)', (ts) => ts.filter((t) => !t.hidden && t.dataset.tab).map((t) => t.dataset.tab));
      const bad = [];
      const toggleBad = [];
      let toggleCount = 0;
      for (const tab of tabs) {
        await n.page.click(`.ribbon-tab[data-tab="${tab}"]`);
        const s = await narrowState(n.page);
        if (!narrowOk(s)) bad.push(`${tab} ${JSON.stringify(s)}`);
        // 접힌 묶음 단추마다 보이는 한글 이름(단추 안)이 있고 아이콘은 16px 이하
        const tg = await n.page.$$eval('.ribbon-panel:not([hidden]) .rb-group-toggle', (bs) => bs.filter((b) => b.getClientRects().length > 0).map((b) => {
          const r = b.getBoundingClientRect();
          const sp = b.querySelector(':scope > span');
          const sr = sp && sp.getClientRects().length > 0 ? sp.getBoundingClientRect() : null;
          const ico = b.querySelector('.rb-ico');
          const ir = ico && ico.getClientRects().length > 0 ? ico.getBoundingClientRect() : null;
          const named = !!sr && sr.width > 0 && (sp.textContent ?? '').trim().length > 0 && sr.top >= r.top - 1 && sr.bottom <= r.bottom + 1 && sr.right <= r.right + 1;
          return { name: b.title, named, ico: ir ? Math.round(Math.max(ir.width, ir.height)) : 0 };
        }));
        toggleCount += tg.length;
        for (const t of tg) if (!t.named || t.ico > 16) toggleBad.push(`${tab}/${t.name} ${JSON.stringify(t)}`);
      }
      rec(`(7) ${ext} ${mode} 600px — 탭 ${tabs.length}개 모두 가로 넘침 없음·제목 줄 한 줄`, tabs.length >= 5 && bad.length === 0, bad.join(' | '));
      // 목록으로 만든 앱만(PPT는 Task 5 전엔 접힌 묶음이 없다)
      if (LISTED.has(ext)) rec(`(7) ${ext} ${mode} 600px — 접힌 묶음 단추 ${toggleCount}개마다 보이는 이름이 단추 안에 있고 아이콘 ≤ 16px`, toggleCount > 0 && toggleBad.length === 0, toggleBad.slice(0, 3).join(' | '));
      await narrow.close();
    }
  }

  // ---- 보기 화면(view.html): 리브레오피스로 그린 docx, pdf ----
  for (const ext of ['docx', 'pdf']) {
    const tok = await token(ids[ext], {});
    const viewPath = new URL(tok.viewerUrl, BASE).pathname;
    const wide = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const v = await open(wide, tok.viewerUrl, '.page-frame.page-ready');
    const vh = await heights(v.page, VIEW_SEL);
    rec(`(8) 보기 ${ext} — ${viewPath}, 높이 24/22/62(상태 줄 없음)`, viewPath === '/view.html' && heightProblems(vh).length === 0 && vh.title !== null && vh.tabs !== null && vh.ribbon !== null, `${JSON.stringify(vh)} ${heightProblems(vh).join(', ')}`);
    const vl = latinOnlyLabels(await labels(v.page));
    const groups = await v.page.$$eval('#toolbar .ribbon-group-label', (ls) => ls.map((l) => l.textContent));
    rec(`(8) 보기 ${ext} — 묶음 표시·보기·쪽·찾기·인쇄·내려받기, 라벨 한글`, vl.length === 0 && groups.join(',') === '표시,보기,쪽,찾기,인쇄,내려받기', `${groups.join(',')} ${vl.slice(0, 3).join(' | ')}`);
    const vs = await v.page.evaluate((list) => ({
      missing: list.filter((id) => !document.getElementById(id)),
      printRight: !!document.getElementById('print-btn')?.closest('.toolbar-tools-right'),
    }), VIEW_PAGE_IDS);
    const under = await underToolbar(v.page);
    rec(`(8) 보기 ${ext} — id 그대로·인쇄는 오른쪽 칸·첫 쪽이 막대 밑에 깔리지 않음`, vs.missing.length === 0 && vs.printRight && !under, JSON.stringify({ ...vs, under }));
    for (const width of [1400, 900, 800]) {
      await v.page.setViewportSize({ width, height: 900 });
      const vp = await v.page.evaluate(panelState, '.toolbar-row-tools');
      rec(`(8) 보기 ${ext} ${width}px — 리본 스크롤바 없음·이름 줄 겹침 없음`, !!vp && vp.bar <= 0 && vp.clash.length === 0, JSON.stringify(vp));
    }
    await v.page.setViewportSize({ width: 1400, height: 900 });
    const vic = await iconChrome(v.page);
    rec(`(9) 보기 ${ext} — 단추마다 아이콘(30px 이상) 아래 한글 이름, 편집 단추·상태 표시 안 보임`, vic.big >= 30 && vic.names.length === 0 && vic.chrome.length === 0, JSON.stringify(vic));
    if (ext === 'docx') {
      // 편집 권한 토큰 — #edit-btn이 켜져도(hidden 아님) 화면에는 보이지 않는다
      const ectx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
      const e = await open(ectx, (await token(ids[ext], { canEdit: true })).viewerUrl, '.page-frame.page-ready');
      const eb = await e.page.$eval('#edit-btn', (b) => ({ on: !b.hidden, shown: b.getClientRects().length > 0 }));
      rec('(9) 보기 docx 편집 토큰 — 편집 단추 켜짐·화면에는 숨김', eb.on && !eb.shown, JSON.stringify(eb));
      await ectx.close();
    }
    rec(`보기 ${ext} — 화면 오류 없음`, v.errors.length === 0, v.errors.slice(0, 2).join(' | '));
    await shot(v.page, `view-${ext}`);
    pairs.push([`view-${ext}`, 'word-report.png']);
    await wide.close();

    const narrow = await browser.newContext({ viewport: { width: 600, height: 900 } });
    const n = await open(narrow, tok.viewerUrl, '.page-frame.page-ready');
    const ns = await narrowState(n.page);
    const nUnder = await underToolbar(n.page);
    rec(`(7) 보기 ${ext} 600px — 가로 넘침 없음·제목 줄 한 줄·첫 쪽이 막대 밑에 깔리지 않음`, narrowOk(ns) && !nUnder, JSON.stringify({ ...ns, under: nUnder }));
    await narrow.close();
  }
  writeReport(pairs);
  console.log(`나란히 비교 보고서: ${path.join(SHOTS, 'report.html')}`);
} catch (error) {
  rec('스크립트 오류', false, String(error?.stack ?? error));
} finally {
  await browser?.close().catch(() => {});
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${results.length - failed}/${results.length} 통과`);
if (!failed) fs.rmSync(RUN, { recursive: true, force: true, maxRetries: 5, retryDelay: 500 });
else console.log(`실패 — 서버 기록·데이터는 ${RUN}에 남겨 둔다`);
process.exit(failed ? 1 : 0);
