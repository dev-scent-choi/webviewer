#!/usr/bin/env node
// 오피스 리본 목록 대조(설계 2026-11-22-office-ribbon-structure §4) — 세 편집기의 탭·묶음·단추가 src/client/ribbon/inventory.json과 1:1:
//  (1) 탭 이름·순서·편집 전용·상황별  (2) 묶음 이름·순서·접힘 순위·자세히  (3) 단추 data-rb·이름·크기·종류·준비 중 여부
//  (4) 준비 중은 꺼짐·말풍선 "준비 중 — 이름"  (5) 기존 기능 선택자가 하나씩, 그 단추 안에 있고 가려지지 않음(접힌 묶음은 펼쳐서)
//  (6) 연 탭마다 아이콘이 붙고 스프라이트에 모양이 있음  (7) 1920·1400·1280·1200·900·800px 모든 탭 가로 넘침 없음·접힘은 순위대로,
//      1200px에서 기존 시험이 누르는 홈 묶음은 펼친 채  (8) 600px 모든 묶음 접힘·가로 넘침 없음·펼친 판이 화면 안
//  (9) ui=view — 준비 중·접힌 묶음 단추에 .ro-disabled 없음  (10) gen-ribbon --check  + M365 캡처와 나란히 보고서
//  (11) 1400px 홈 — 기존 기능이 있는 묶음은 크게 또는 작은 단추 3줄(rb-compact)로 보이고 접히지 않음
//  (12) Ctrl+F1 축소→재고정 때 열린 묶음 판이 닫힘  (13) PPT 상자를 누르면(pointerdown) 도형 서식 탭
// (7)의 접힘은 "순위 순서"가 아니라 "더 펼 수 없음" — 접힌 묶음을 크게든 작게든 펴면 판이 넘쳐야 한다(fitPanel이 다시 펴 보기 때문)
// 쓰는 법: npm run build:editor && node scripts/verify-ribbon-inventory-e2e.mjs [포트 8180~8189, 기본 8180]
import { execFileSync, spawn, spawnSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';
import { APP_OF_EXT, controlKey, expectedStructure, loadInventory, refsOf } from './ribbonGenLib.mjs';

const PORT = Number(process.argv[2] ?? 8180);
// 운영 포트(8000·8010·8020·8100·8110·8130·8148)와 다른 시험 범위를 피해 이 계획 범위만
if (!(PORT >= 8180 && PORT <= 8189)) throw new Error(`시험 포트는 8180~8189만 쓴다(받은 값 ${PORT})`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `ribbon-inventory-e2e-${PORT}`);
const SHOTS = path.resolve('.audit-run', 'ribbon-inventory-shots');
const M365 = path.resolve('.audit-run', 'm365-ribbon');
const FILES = {
  docx: 'audit/office/corpus/docx/a1-korean-report.docx',
  xlsx: 'audit/office/corpus/xlsx/x07-multisheet-print.xlsx',
  pptx: 'audit/office/corpus/pptx/a2-korean-deck.pptx',
};
const READY = { docx: '.ProseMirror', xlsx: '#grid-edit-table td[data-row]', pptx: '#slides-rail .slide-card' };
// 기존 e2e(verify-xlsx-view 1200px, browser-check 1280px, readonly 강제 클릭)가 누르는 홈 묶음 — 1200px에서 접히면 그 시험이 깨진다
const LEGACY_OPEN = { docx: ['clipboard', 'font', 'paragraph'], xlsx: ['clipboard', 'font', 'alignment', 'cells'], pptx: ['slides', 'font'] };
const WIDTHS = [1920, 1400, 1366, 1280, 1200, 900, 800];
const INV = loadInventory();
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

async function upload(file) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패 ${file}: ${JSON.stringify(r)}`);
  for (let i = 0; i < 1200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready') return st.id;
    if (st.status === 'failed') throw new Error(`변환 실패 ${file}: ${st.error}`);
    await sleep(250);
  }
  throw new Error(`대기 시간 초과 ${file}`);
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: 'ribbon-inventory-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);
async function open(context, url, ready) {
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.goto(`${BASE}${url}`);
  await page.waitForSelector(ready, { timeout: 90000 });
  await page.waitForFunction(() => !document.querySelector('#conn-status') || document.querySelector('#conn-status').getAttribute('data-status') === 'synced', null, { timeout: 30000 }).catch(() => {});
  await sleep(800);
  return { page, errors };
}
const withCtx = (url) => `${url}${url.includes('?') ? '&' : '?'}rbctx=1`;

/** 브라우저 안 — DOM을 목록과 같은 모양으로 읽는다(이름·종류·크기 규칙은 계획 Task 7에 적힌 대로) */
function readRibbon(refs) {
  const text = (el) => (el?.textContent ?? '').replace(/\s+/g, ' ').trim();
  const clean = (s) => s.replace(/\s*▾\s*$/, '').trim();
  const nameOf = (unit, ref) => {
    if (unit.matches('label')) {
      const s = [...unit.querySelectorAll(':scope > span')].map(text).find(Boolean);
      if (s) return clean(s);
    }
    const lab = ref.querySelector('.ribbon-btn-label');
    if (lab) return clean(text(lab));
    if (ref.classList.contains('icon-btn-big') || ref.classList.contains('rb-gallery')) {
      const s = ref.querySelector(':scope > span');
      if (s && text(s)) return clean(text(s));
    }
    const aria = ref.getAttribute('aria-label');
    if (aria) return clean(aria);
    return clean((ref.getAttribute('title') ?? '').split(/ \(| — /)[0]);
  };
  const kindOf = (unit, ref) => {
    if (unit.dataset.kind) return unit.dataset.kind;
    if (unit.matches('.style-gallery, .effect-gallery, .slide-bg-swatches')) return 'gallery';
    if (ref.matches('select')) return 'select';
    if (ref.matches('input[type="checkbox"]')) return 'toggle';
    if (ref.matches('input[type="color"]')) return 'split';
    if (ref.matches('input')) return 'input';
    if (unit.matches('.dropdown-caret') || unit.querySelector('.dropdown-caret')) return ref.matches('.dropdown-caret') ? 'dropdown' : 'split';
    if (!ref.matches('button')) return 'info';
    return 'button';
  };
  const tabs = [...document.querySelectorAll('.ribbon-tabs .ribbon-tab[data-tab]')].map((t) => ({ id: t.dataset.tab, name: text(t), editOnly: t.hasAttribute('data-edit-only'), contextual: t.classList.contains('ribbon-tab-contextual') }));
  const panels = {};
  for (const p of document.querySelectorAll('#toolbar .ribbon-panel[data-panel]')) {
    panels[p.dataset.panel] = [...p.querySelectorAll(':scope > .ribbon-group')].map((g) => {
      const l = g.querySelector(':scope > .ribbon-launcher');
      return {
        id: g.dataset.group ?? '',
        name: text(g.querySelector(':scope > .ribbon-group-label')),
        collapse: Number(g.dataset.collapse),
        launcher: l ? (l.hasAttribute('data-pending') ? 'pending' : `proxy:${l.dataset.proxy}`) : null,
        controls: [...g.querySelectorAll('[data-rb]')].map((unit) => {
          const key = unit.dataset.rb;
          const ref = (refs[key] && document.querySelector(refs[key])) || unit;
          const kind = kindOf(unit, ref);
          return { key, name: nameOf(unit, ref), size: kind === 'gallery' || ref.classList.contains('icon-btn-big') ? 'big' : 'small', kind, pending: unit.hasAttribute('data-pending') };
        }),
      };
    });
  }
  return { tabs, panels };
}
/** 같은 모양 둘의 차이를 경로와 함께(처음 20개) */
function diff(want, got, at = '', out = []) {
  if (out.length >= 20) return out;
  if (Array.isArray(want)) {
    if (!Array.isArray(got) || got.length !== want.length) out.push(`${at} 길이 목록 ${want.length} · 화면 ${Array.isArray(got) ? got.length : typeof got}${Array.isArray(got) ? ` (${JSON.stringify(got.map((x) => x.key ?? x.id ?? x)).slice(0, 300)})` : ''}`);
    want.forEach((w, i) => diff(w, got?.[i], `${at}[${i}${w?.key ? ` ${w.key}` : w?.id ? ` ${w.id}` : ''}]`, out));
  } else if (want && typeof want === 'object') {
    for (const k of Object.keys(want)) diff(want[k], got?.[k], `${at}.${k}`, out);
  } else if (want !== got) out.push(`${at} 목록 ${JSON.stringify(want)} · 화면 ${JSON.stringify(got)}`);
  return out;
}
/** 판 하나의 접힘 상태 — 넘침, 접힌 묶음마다 크게·작게 펴 보면 넘치는지(넘치지 않으면 덜 편 것) */
const fitState = (page, tabId) =>
  page.evaluate((t) => {
    const p = document.querySelector(`.ribbon-panel[data-panel="${t}"]`);
    const groups = [...p.querySelectorAll(':scope > .ribbon-group[data-collapse]')];
    const over = () => p.scrollWidth - p.clientWidth;
    const couldOpen = [];
    for (const g of groups.filter((x) => x.classList.contains('rb-collapsed'))) {
      g.classList.remove('rb-collapsed');
      const full = over() <= 1;
      g.classList.add('rb-compact');
      const compact = over() <= 1;
      g.classList.remove('rb-compact');
      g.classList.add('rb-collapsed');
      if (full || compact) couldOpen.push(`${g.dataset.group}${full ? '(크게)' : '(작게)'}`);
    }
    const state = (g) => (g.classList.contains('rb-collapsed') ? 'collapsed' : g.classList.contains('rb-compact') ? 'compact' : 'full');
    return {
      over: over(),
      collapsed: groups.filter((g) => g.classList.contains('rb-collapsed')).map((g) => g.dataset.group),
      compact: groups.filter((g) => g.classList.contains('rb-compact')).map((g) => g.dataset.group),
      existingCollapsed: groups.filter((g) => state(g) === 'collapsed' && g.querySelector('[data-rb]:not([data-pending]), [data-proxy]:not(.ribbon-launcher)')).map((g) => g.dataset.group),
      couldOpen,
      // 리본(#toolbar)이 창을 넘는지 — 문서 본문(워드 쪽 폭 확대 상태 등)은 이 시험 대상이 아니다
      doc: Math.max(document.getElementById('toolbar').scrollWidth, Math.round(document.getElementById('toolbar').getBoundingClientRect().right)),
    };
  }, tabId);

function writeReport(rows) {
  const body = rows
    .map(([ours, ref]) => {
      const has = ref && fs.existsSync(path.join(M365, ref));
      if (has) fs.copyFileSync(path.join(M365, ref), path.join(SHOTS, `m365-${ref}`));
      return `<tr><td><img src="${ours}" alt="${ours}"><p>${ours}</p></td><td>${has ? `<img src="m365-${ref}" alt="${ref}">` : '(M365 캡처 없음 — Task 1 덤프)'}<p>${ref ?? ''}</p></td></tr>`;
    })
    .join('');
  fs.writeFileSync(path.join(SHOTS, 'report.html'), `<!doctype html><html lang="ko"><meta charset="utf-8"><title>리본 구조 비교</title><style>body{font-family:"맑은 고딕",sans-serif}img{width:900px;border:1px solid #ccc}td{vertical-align:top;padding:8px}</style><h1>우리 리본(1920px, ?rbctx=1)과 M365 나란히</h1><table><tr><th>우리</th><th>M365</th></tr>${body}</table></html>`);
}

const check = spawnSync(process.execPath, ['scripts/gen-ribbon.mjs', '--check'], { encoding: 'utf8' });
rec('(10) gen-ribbon --check — 세 페이지가 목록과 같음', check.status === 0, (check.stderr || '').trim().split('\n').slice(0, 3).join(' | '));

await startServer();
fs.rmSync(SHOTS, { recursive: true, force: true });
fs.mkdirSync(SHOTS, { recursive: true });
const report = [];
let browser;
try {
  browser = await chromium.launch({ ignoreDefaultArgs: ['--hide-scrollbars'] });
  for (const [ext, file] of Object.entries(FILES)) {
    const app = INV.apps[APP_OF_EXT[ext]];
    const refs = refsOf(app);
    const id = await upload(file);
    const editorUrl = (await token(id, { canEdit: true })).editorUrl;
    const ctx = await browser.newContext({ viewport: { width: 1920, height: 1000 } });
    const { page, errors } = await open(ctx, withCtx(editorUrl), READY[ext]);

    // (1)~(3) 구조 1:1
    const want = expectedStructure(app);
    const got = await page.evaluate(readRibbon, refs);
    const tabDiff = diff(want.tabs, got.tabs, 'tabs');
    rec(`(1) ${ext} 탭 ${want.tabs.length}개 — 이름·순서·편집 전용·상황별`, tabDiff.length === 0, tabDiff.join(' | '));
    const panelDiff = diff(want.panels, got.panels, 'panels');
    rec(`(2)(3) ${ext} 묶음·단추 — 이름·순서·접힘 순위·자세히·크기·종류·준비 중`, panelDiff.length === 0 && Object.keys(got.panels).length === Object.keys(want.panels).length, panelDiff.join(' | '));

    // (4)(5) 준비 중 꺼짐, 기존 선택자
    const units = await page.evaluate((r) => {
      const badPending = [...document.querySelectorAll('#toolbar [data-pending]')]
        .filter((u) => !(u.matches(':disabled') || u.querySelector(':disabled') || u.classList.contains('rb-gallery')) || !(u.getAttribute('title') ?? '').startsWith('준비 중 — '))
        .map((u) => u.dataset.rb ?? u.getAttribute('aria-label'));
      const badRefs = Object.entries(r)
        .filter(([key, sel]) => {
          const all = document.querySelectorAll(sel);
          const unit = document.querySelector(`[data-rb="${key}"]`);
          return all.length !== 1 || !unit || !(unit === all[0] || unit.contains(all[0])) || unit.hasAttribute('data-pending');
        })
        .map(([key, sel]) => `${key} ${sel}`);
      return { pending: document.querySelectorAll('#toolbar [data-pending]').length, badPending, badRefs };
    }, refs);
    rec(`(4) ${ext} 준비 중 ${units.pending}개 모두 꺼짐·말풍선 "준비 중 — 이름"`, units.pending > 0 && units.badPending.length === 0, units.badPending.slice(0, 5).join(', '));
    rec(`(5) ${ext} 기존 기능 ${Object.keys(refs).length}개 — 선택자가 하나씩, 그 단추(data-rb) 안`, units.badRefs.length === 0, units.badRefs.slice(0, 5).join(' | '));

    // (5) 눌림 — 탭을 하나씩 고르고, 접힌 묶음은 펼쳐서 가운데가 그 단추에 닿는지 + (6) 아이콘
    const covered = [];
    for (const tab of [...app.tabs, ...app.contextual]) {
      await page.click(`.ribbon-tab[data-tab="${tab.id}"]`);
      await sleep(150);
      await page.locator('#toolbar').screenshot({ path: path.join(SHOTS, `${ext}-${tab.id}.png`) });
      report.push([`${ext}-${tab.id}.png`, `${APP_OF_EXT[ext]}-${tab.name}.png`.replace(/[\\/:*?"<>|]/g, '_')]);
      for (const g of tab.groups) {
        for (const c of g.controls) {
          if (c.pending || c.kind === 'info') continue;
          const groupSel = `.ribbon-panel[data-panel="${tab.id}"] .ribbon-group[data-group="${g.id}"]`;
          const folded = await page.$eval(groupSel, (el) => el.classList.contains('rb-collapsed'));
          if (folded) await page.click(`${groupSel} > .rb-group-toggle`);
          const hit = await page.$eval(refs[controlKey(tab.id, g.id, c)], (el) => {
            const r = el.getBoundingClientRect();
            if (!r.width || !r.height) return 'hidden';
            const at = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
            return at && (at === el || el.contains(at) || at.contains(el)) ? 'ok' : `가림 ${at?.className ?? at?.tagName}`;
          });
          if (hit !== 'ok') covered.push(`${controlKey(tab.id, g.id, c)} ${hit}`);
          if (folded) await page.keyboard.press('Escape');
        }
      }
    }
    await page.click('.ribbon-tab[data-tab="home"]');
    rec(`(5) ${ext} 기존 기능이 모두 보이고 가려지지 않음(접힌 묶음은 펼쳐서)`, covered.length === 0, covered.slice(0, 5).join(' | '));
    const icons = await page.evaluate(() => {
      const els = [...document.querySelectorAll('#toolbar .ribbon-panel [data-icon]')];
      const missing = els.filter((el) => !el.querySelector(':scope > svg.rb-ico > use')?.getAttribute('href')?.endsWith(`#rbi-${el.dataset.icon}`)).map((el) => el.dataset.rb ?? el.dataset.icon);
      return { count: els.length, missing, used: [...new Set(els.map((el) => el.dataset.icon))] };
    });
    const sprite = await (await fetch(`${BASE}/ribbon/icons.svg`)).text();
    const noShape = icons.used.filter((k) => !sprite.includes(`id="rbi-${k}"`));
    rec(`(6) ${ext} 연 탭의 아이콘 ${icons.count}개가 붙고 스프라이트에 모양이 있음`, icons.missing.length === 0 && noShape.length === 0, JSON.stringify({ missing: icons.missing.slice(0, 5), noShape: noShape.slice(0, 5) }));
    rec(`${ext} 화면 오류 없음`, errors.length === 0, errors.slice(0, 2).join(' | '));

    // (7) 폭별 접힘
    for (const width of WIDTHS) {
      await page.setViewportSize({ width, height: 1000 });
      await sleep(300);
      const bad = [];
      for (const tab of app.tabs) {
        await page.click(`.ribbon-tab[data-tab="${tab.id}"]`);
        await sleep(80);
        const s = await fitState(page, tab.id);
        if (s.over > 1 || s.couldOpen.length || s.doc > width) bad.push(`${tab.id} ${JSON.stringify(s)}`);
        if (width === 1400 && tab.id === 'home') {
          rec(`(11) ${ext} 1400px 홈 — 기존 기능 묶음 접힘 없음(작게 ${s.compact.join(',') || '없음'})`, s.existingCollapsed.length === 0 && s.over <= 1, JSON.stringify(s));
        }
        if (width === 1200 && tab.id === 'home') {
          const closed = LEGACY_OPEN[ext].filter((g) => s.collapsed.includes(g));
          if (closed.length) bad.push(`1200px 홈에서 기존 시험이 누르는 묶음이 접힘: ${closed.join(',')} — 목록 collapse를 높일 것`);
          console.log(`     ${ext} 1200px 홈 접힘: ${s.collapsed.join(',') || '없음'}`);
        }
      }
      await page.click('.ribbon-tab[data-tab="home"]');
      rec(`(7) ${ext} ${width}px — 기본 탭 ${app.tabs.length}개 가로 넘침 없음·더 펼 수 있는 접힌 묶음 없음`, bad.length === 0, bad.join(' | '));
    }
    // (12) 900px에서 접힌 묶음 판을 연 채 Ctrl+F1 두 번(축소→재고정) — 판이 다시 열리면 안 된다
    await page.setViewportSize({ width: 900, height: 1000 });
    await sleep(300);
    const toggle = `.ribbon-panel[data-panel="home"] > .ribbon-group.rb-collapsed > .rb-group-toggle`;
    const hasFold = (await page.$(toggle)) !== null;
    if (hasFold) await page.click(toggle);
    const opened = await page.$('.ribbon-group.rb-open');
    await page.keyboard.press('Control+F1');
    await sleep(150);
    await page.keyboard.press('Control+F1');
    await sleep(200);
    const reopened = await page.$('.ribbon-group.rb-open');
    rec(`(12) ${ext} Ctrl+F1 축소→재고정 — 열린 묶음 판이 닫힌 채`, hasFold && !!opened && !reopened, JSON.stringify({ hasFold, opened: !!opened, reopened: !!reopened }));
    await ctx.close();

    // (13) PPT 상자 테두리 쪽을 누르면 도형 서식 탭(상황별 탭을 모두 보이게 하는 ?rbctx=1 없이)
    if (ext === 'pptx') {
      const pc = await browser.newContext({ viewport: { width: 1400, height: 900 } });
      const sp = await open(pc, editorUrl, READY[ext]);
      const before = await sp.page.$eval('.ribbon-tab[data-tab="shape-format"]', (t) => t.hidden);
      // 글 영역 포커스(이미 탭을 여는 길)와 떼어 pointerdown만 보낸다
      await sp.page.$eval('.slide-textbox', (b) => b.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true })));
      await sleep(200);
      const after = await sp.page.$eval('.ribbon-tab[data-tab="shape-format"]', (t) => t.hidden);
      rec('(13) pptx 상자를 누르면 도형 서식 탭이 나타남', before === true && after === false, JSON.stringify({ before, after }));
      await pc.close();
    }

    // (8) 600px
    const narrow = await browser.newContext({ viewport: { width: 600, height: 900 } });
    const n = await open(narrow, editorUrl, READY[ext]);
    const ns = await n.page.evaluate(() => {
      const groups = [...document.querySelectorAll('.ribbon-panel:not([hidden]) > .ribbon-group[data-collapse]')];
      return { all: groups.length, folded: groups.filter((g) => g.classList.contains('rb-collapsed')).length, doc: document.documentElement.scrollWidth };
    });
    await n.page.click('.ribbon-panel:not([hidden]) > .ribbon-group[data-collapse] > .rb-group-toggle');
    const pop = await n.page.evaluate(() => {
      const b = document.querySelector('.ribbon-group.rb-open > .ribbon-group-buttons');
      if (!b) return null;
      const r = b.getBoundingClientRect();
      return { left: Math.round(r.left), right: Math.round(r.right), w: Math.round(r.width) };
    });
    rec(`(8) ${ext} 600px — 홈 묶음 ${ns.all}개 모두 접힘·가로 넘침 없음·펼친 판이 화면 안`, ns.all > 0 && ns.folded === ns.all && ns.doc <= 600 && !!pop && pop.left >= 0 && pop.right <= 600 && pop.w > 0, JSON.stringify({ ...ns, pop }));
    await narrow.close();

    // (9) ui=view
    const ro = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const viewUrl = (await token(id, { engine: 'editor' })).viewerUrl;
    const v = await open(ro, viewUrl, READY[ext]);
    const vs = await v.page.evaluate(() => ({
      body: document.body.classList.contains('ui-view'),
      pendingGrey: [...document.querySelectorAll('#toolbar [data-pending], #toolbar [data-pending] *')].filter((el) => el.classList.contains('ro-disabled')).length,
      toggleGrey: document.querySelectorAll('#toolbar .rb-group-toggle.ro-disabled').length,
      greyed: document.querySelectorAll('#toolbar .ro-disabled').length,
    }));
    rec(`(9) ${ext} ui=view — 쓰는 단추 흐림 ${vs.greyed}개, 준비 중·접힌 묶음 단추엔 흐림 없음`, vs.body && vs.greyed > 0 && vs.pendingGrey === 0 && vs.toggleGrey === 0, JSON.stringify(vs));
    if (ext === 'pptx') {
      await v.page.$eval('.slide-textbox', (b) => b.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true })));
      await sleep(200);
      const hidden = await v.page.$eval('.ribbon-tab[data-tab="shape-format"]', (t) => t.hidden);
      rec('(13) pptx ui=view — 상자 pointerdown이 도형 서식 탭을 열지 않음', hidden === true, JSON.stringify({ hidden }));
    }
    await ro.close();
    // (9) 좁은 폭 ui=view — 열린 접힌 묶음 판 안의 조작은 꺼진 채
    const rn = await browser.newContext({ viewport: { width: 600, height: 900 } });
    const vn = await open(rn, viewUrl, READY[ext]);
    await vn.page.click('.ribbon-panel:not([hidden]) > .ribbon-group[data-collapse] > .rb-group-toggle');
    const pn = await vn.page.evaluate(() => {
      const ctl = [...document.querySelectorAll('.ribbon-group.rb-open > .ribbon-group-buttons button, .ribbon-group.rb-open > .ribbon-group-buttons select, .ribbon-group.rb-open > .ribbon-group-buttons input')];
      const live = ctl.filter((el) => !el.closest('[data-pending]') && el.id !== 'btn-copy' && // 복사는 읽기 전용에서도 허용(canCopy)
         !el.classList.contains('ro-disabled') && !el.disabled && el.getAttribute('aria-disabled') !== 'true');
      return { total: ctl.length, live: live.length };
    });
    rec(`(9) ${ext} ui=view 접힌 묶음 판 — 안의 조작 ${pn.total}개 모두 꺼짐(복사 제외)`, pn.total > 0 && pn.live === 0, JSON.stringify(pn));
    await rn.close();
  }
  writeReport(report);
  console.log(`\n나란히 보기: ${path.join(SHOTS, 'report.html')}`);
} catch (error) {
  rec('스크립트 오류', false, String(error?.stack ?? error));
} finally {
  await browser?.close();
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed);
