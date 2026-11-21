// 재연결 회귀 검사(워드·엑셀·PPT·PDF) — 실제 브라우저로 2·3차 검토 RX-C1·R3-I1·R3-I2를 확인한다.
//  1) 짧은 끊김(방 닫기 유예 안): 새로 고침 없이 이어지고 끊긴 동안의 입력도 그대로, 내용은 한 벌
//  2) 반쯤 열린 연결(보낸 프레임이 사라짐) 뒤 재접속: 그 사이 입력이 사라지지 않는다
//  3) 열어 둔 채 서버 재시작(방이 새 세대): 안내가 뜨고, 입력이 있었으면 '내 입력 내용'을 건네며 저절로 새로 고치지 않는다.
//     입력이 없었으면 초읽기 뒤 저절로 새로 고친다. 어느 쪽이든 내용·저장 파일은 한 벌
//  4) 세대 값을 모르는 옛 번들(주소에 epoch 없음, 메시지 7·8 모름): 유예 안 재접속은 받아 주고, 서버 재시작 뒤엔 4409로 거절해
//     옛 상태를 합치지 않는다(PDF 포함)
//
// 이 스크립트가 시험 서버를 직접 띄우고 다시 띄운다(재시작 시험) — 작업 폴더에서 `npm run build:editor` 뒤:
//   node scripts/verify-reconnect-e2e.mjs [포트=8145]
// 서버 비밀 값은 이 실행에서만 쓰는 임의 값(재시작해도 같은 값이라 토큰이 유효 — 실제 배포 재시작과 같다).
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import JSZip from 'jszip';
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';

const PORT = Number(process.argv[2] ?? 8145);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run');
const SECRET = randomBytes(24).toString('hex');
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

let server = null;
async function startServer() {
  fs.mkdirSync(path.join(RUN, 'tmp'), { recursive: true });
  const log = fs.openSync(path.join(RUN, `server-${PORT}-reconnect.log`), 'a');
  server = spawn(process.execPath, ['node_modules/tsx/dist/cli.mjs', 'src/server.ts'], {
    env: { ...process.env, DOCVIEW_PORT: String(PORT), DOCVIEW_HOST: '127.0.0.1', DOCVIEW_SECRET: SECRET, DOCVIEW_DATA_DIR: path.join(RUN, 'data'), TMP: path.join(RUN, 'tmp'), TEMP: path.join(RUN, 'tmp'), DOCVIEW_LO_PROFILE_ROOT: path.join(RUN, 'lo') },
    stdio: ['ignore', log, log],
  });
  for (let i = 0; i < 120; i += 1) {
    if (await fetch(`${BASE}/api/formats`).then((r) => r.ok, () => false)) return;
    await sleep(500);
  }
  throw new Error('시험 서버가 뜨지 않음');
}
function stopServer() {
  // 내가 띄운 서버만(PID, 자식 tsx 프로세스 포함) — 이름으로 node를 끄지 않는다
  if (server?.pid) execFileSync('taskkill', ['/PID', String(server.pid), '/T', '/F'], { stdio: 'ignore' });
  server = null;
}
async function restartServer() {
  stopServer();
  for (let i = 0; i < 40 && (await fetch(`${BASE}/api/formats`).then(() => true, () => false)); i += 1) await sleep(250);
  await startServer();
}

async function open(file, extra = {}) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
  const id = (await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json()).data.id;
  for (let i = 0; ; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data;
    if (st.status === 'ready') break;
    if (st.status === 'failed' || i > 600) throw new Error(`변환 실패 ${file}`);
    await sleep(400);
  }
  const token = (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: 'reconnect', canPrint: true, canDownload: true, canCopy: true, canEdit: true, ...extra }) })).json()).data;
  return { id, token };
}
const save = (t) => fetch(`${BASE}/api/view/save?token=${encodeURIComponent(t.token)}`, { method: 'POST' });
const download = async (t) => JSZip.loadAsync(Buffer.from(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(t.token)}`)).arrayBuffer()));

/** 웹소켓 가로채기 — 끊기·반쯤 열림(보내기 버림)·옛 번들 흉내(epoch 빼기, 메시지 7·8 버리기)·닫기 코드 기록 */
function wsHook(oldBundle) {
  const W = window.WebSocket;
  window.__ws = [];
  window.__codes = [];
  window.__mute = false;
  window.WebSocket = class extends W {
    constructor(url, ...rest) {
      super(oldBundle ? String(url).replace(/([?&])epoch=[^&]*&?/, '$1').replace(/[?&]$/, '') : url, ...rest);
      window.__ws.push(this);
      this.addEventListener('close', (e) => window.__codes.push(e.code));
    }
    send(data) {
      if (!window.__mute) super.send(data);
    }
    set onmessage(fn) {
      super.onmessage = !oldBundle ? fn : (e) => {
        const first = new Uint8Array(e.data)[0];
        if (first !== 7 && first !== 8) fn(e);
      };
    }
    get onmessage() {
      return super.onmessage;
    }
  };
}

async function newPage(browser, token, { oldBundle = false } = {}) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  page.errors = [];
  page.on('pageerror', (e) => page.errors.push(String(e)));
  await page.addInitScript(wsHook, oldBundle);
  await page.goto(`${BASE}${token.editorUrl}&home=hidden`);
  return page;
}
const openSockets = (page) => page.evaluate(() => window.__ws.filter((w) => w.readyState === 1).length);
async function waitReconnected(page, before) {
  for (let i = 0; i < 60; i += 1) {
    if ((await page.evaluate(() => window.__ws.length)) > before && (await openSockets(page)) > 0) return true;
    await sleep(500);
  }
  return false;
}
const dropSockets = (page) => page.evaluate(() => window.__ws.forEach((w) => w.close()));
const count = (s, marker) => (s.match(new RegExp(marker, 'g')) ?? []).length;

const KINDS = {
  PPT: {
    file: 'audit/office/corpus/pptx/p01-placeholders-bullets.pptx',
    ready: '#slide-canvas .slide-textbox .ProseMirror',
    count: (page) => page.locator('.slide-card').count(),
    edit: async (page, text) => {
      await page.locator('#slide-canvas .slide-textbox .ProseMirror').first().click();
      await page.keyboard.press('End');
      await page.keyboard.type(text);
    },
    saved: async (zip, marker) => {
      const pres = await zip.file('ppt/presentation.xml').async('string');
      const ids = [...pres.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]);
      let hits = 0;
      for (const n of Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))) hits += count(await zip.file(n).async('string'), marker);
      return { sldId: ids.length, unique: new Set(ids).size, marker: hits };
    },
  },
  워드: {
    file: 'audit/office/corpus/docx/d01-styles-fonts.docx',
    ready: '.ProseMirror',
    count: (page) => page.locator('.ProseMirror > *').count(),
    edit: async (page, text) => {
      await page.locator('.ProseMirror p').first().click();
      await page.keyboard.press('End');
      await page.keyboard.type(text);
    },
    saved: async (zip, marker) => {
      const xml = await zip.file('word/document.xml').async('string');
      return { paragraphs: count(xml, '<w:p[\\s>]'), marker: count(xml, marker) };
    },
  },
  엑셀: {
    file: 'audit/office/corpus/xlsx/x08-korean-quotation.xlsx',
    ready: '#grid-edit-table td[data-row]',
    count: (page) => page.evaluate(() => document.querySelectorAll('#grid-edit-table tr').length),
    edit: async (page, text, n = 40) => {
      await page.locator('#grid-edit-table td[data-row]').nth(n).dblclick();
      await page.keyboard.type(text);
      await page.keyboard.press('Enter');
    },
    saved: async (zip, marker) => {
      const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
      const shared = (await zip.file('xl/sharedStrings.xml')?.async('string')) ?? '';
      return { rows: count(sheet, '<row\\b'), marker: count(shared, marker) + count(sheet, marker) };
    },
  },
};

/** 1) 유예 안 짧은 끊김 + 2) 반쯤 열린 연결 */
async function graceAndHalfOpen(browser, label, k) {
  const { id, token } = await open(k.file);
  const page = await newPage(browser, token);
  await page.waitForSelector(k.ready, { timeout: 120000 });
  await sleep(2500);
  const before = await k.count(page);
  await k.edit(page, 'GA');
  await sleep(1200);
  let n = await page.evaluate(() => window.__ws.length);
  await dropSockets(page);
  const back = await waitReconnected(page, n);
  await sleep(1500);
  rec(`${label} 유예: 새로 고침 없이 다시 붙음`, back && (await page.locator('#room-stale-notice').count()) === 0);
  // 반쯤 열림: 브라우저는 연결됐다고 믿지만 보낸 프레임이 서버에 닿지 않는다
  await page.evaluate(() => { window.__mute = true; });
  await k.edit(page, 'HB', 41);
  await sleep(1000);
  n = await page.evaluate(() => window.__ws.length);
  await page.evaluate(() => { window.__mute = false; window.__ws.forEach((w) => w.close()); });
  const back2 = await waitReconnected(page, n);
  await sleep(2500);
  rec(`${label} 반쯤 열림: 안내 없이 다시 붙음`, back2 && (await page.locator('#room-stale-notice').count()) === 0);
  const after = await k.count(page);
  rec(`${label} 유예: 화면 내용 한 벌`, after === before, `${before} → ${after}`);
  await save(token);
  const zip = await download(token);
  const ga = await k.saved(zip, 'GA');
  const hb = await k.saved(zip, 'HB');
  rec(`${label} 유예·반쯤 열림: 끊긴 동안 입력도 저장(각 한 번)`, ga.marker === 1 && hb.marker === 1, JSON.stringify({ ga, hb }));
  rec(`${label} 유예: 페이지 오류 없음`, page.errors.length === 0, page.errors.slice(0, 2).join(' | '));
  await page.close();
  await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' }).catch(() => {});
}

/** 3) 열어 둔 채 서버 재시작 — 입력이 있었던 탭(내 입력 내용, 저절로 새로 고침 없음)과 입력이 없던 탭(초읽기 뒤 새로 고침) */
async function restartScenario(browser) {
  const tabs = [];
  for (const [label, k] of Object.entries(KINDS)) {
    const { id, token } = await open(k.file);
    const page = await newPage(browser, token);
    await page.waitForSelector(k.ready, { timeout: 120000 });
    await sleep(2500);
    await k.edit(page, 'RS');
    await sleep(1500);
    await save(token);
    tabs.push({ label, k, id, token, page, before: await k.count(page), savedBefore: await k.saved(await download(token), 'RS') });
  }
  const idle = await open(KINDS.워드.file);
  const idlePage = await newPage(browser, idle.token);
  await idlePage.waitForSelector(KINDS.워드.ready, { timeout: 120000 });
  await sleep(2000);
  const idleOrigin = await idlePage.evaluate(() => performance.timeOrigin);

  await restartServer();

  for (const t of tabs) {
    const notice = await t.page.waitForSelector('#room-stale-notice', { timeout: 40000 }).then(() => true, () => false);
    const recovery = await t.page.locator('#room-stale-recovery').inputValue().catch(() => '');
    const buttons = (await t.page.getByRole('button', { name: '내 입력 내용 복사' }).count()) + (await t.page.getByRole('button', { name: '내려받기' }).count());
    rec(`${t.label} 재시작: 안내와 '내 입력 내용'(복사·내려받기)`, notice && recovery.includes('RS') && buttons === 2, JSON.stringify(recovery.slice(0, 60)));
  }
  const origins = await Promise.all(tabs.map((t) => t.page.evaluate(() => performance.timeOrigin)));
  await sleep(13000);
  for (const [i, t] of tabs.entries()) {
    rec(`${t.label} 재시작: 입력이 있었으면 저절로 새로 고치지 않음`, (await t.page.evaluate(() => performance.timeOrigin)) === origins[i] && (await t.page.locator('#room-stale-notice').count()) === 1);
  }
  const idleNow = await idlePage.evaluate(() => performance.timeOrigin).catch(() => 0);
  rec('워드(입력 없음) 재시작: 초읽기 뒤 저절로 다시 불러옴', idleNow !== idleOrigin && idleNow !== 0);
  await idlePage.close();
  await fetch(`${BASE}/api/documents/${idle.id}`, { method: 'DELETE' }).catch(() => {});

  for (const t of tabs) {
    const nav = t.page.waitForEvent('framenavigated', { timeout: 30000 });
    await t.page.getByRole('button', { name: '다시 불러오기' }).click();
    await nav;
    await t.page.waitForSelector(t.k.ready, { timeout: 120000 });
    await sleep(2500);
    const after = await t.k.count(t.page);
    rec(`${t.label} 재시작: 다시 불러온 뒤 내용 한 벌`, after === t.before, `${t.before} → ${after}`);
    await save(t.token);
    const savedAfter = await t.k.saved(await download(t.token), 'RS');
    rec(`${t.label} 재시작: 다시 저장한 파일도 한 벌`, JSON.stringify(savedAfter) === JSON.stringify(t.savedBefore), `${JSON.stringify(t.savedBefore)} → ${JSON.stringify(savedAfter)}`);
    rec(`${t.label} 재시작: 페이지 오류 없음`, t.page.errors.length === 0, t.page.errors.slice(0, 2).join(' | '));
    await t.page.close();
    await fetch(`${BASE}/api/documents/${t.id}`, { method: 'DELETE' }).catch(() => {});
  }
}

/** 4) 옛 번들 흉내 — 유예 안 재접속은 받아 주고(워드), 서버 재시작 뒤엔 4409로 거절(넷 다) */
async function oldBundleScenario(browser) {
  const pdfFile = 'docs/quality-audit/out/ref/p01-placeholders-bullets.pdf';
  const tabs = [];
  for (const [label, k] of Object.entries(KINDS)) {
    const { id, token } = await open(k.file);
    const page = await newPage(browser, token, { oldBundle: true });
    await page.waitForSelector(k.ready, { timeout: 120000 });
    await sleep(2500);
    await k.edit(page, 'OB');
    await sleep(1500);
    if (label === '워드') {
      const n = await page.evaluate(() => window.__ws.length);
      await dropSockets(page);
      const back = await waitReconnected(page, n);
      await sleep(1000);
      await k.edit(page, 'OC');
      await sleep(1500);
      await save(token);
      const oc = await k.saved(await download(token), 'OC');
      rec('워드 옛 번들: 유예 안 재접속은 받아 주고 이어서 입력도 저장', back && !(await page.evaluate(() => window.__codes.includes(4409))) && oc.marker === 1, JSON.stringify(oc));
    }
    await save(token);
    tabs.push({ label, k, id, token, page, before: await k.count(page), savedBefore: await k.saved(await download(token), 'OB') });
  }
  if (fs.existsSync(pdfFile)) {
    const { id, token } = await open(pdfFile, { pdfEditor: true });
    const page = await newPage(browser, token, { oldBundle: true });
    await page.waitForSelector('#conn-status[data-state="connected"]', { timeout: 120000 });
    await sleep(2000);
    tabs.push({ label: 'PDF', id, token, page });
  } else rec('PDF 옛 번들: 시험 파일 있음', false, pdfFile);

  await restartServer();
  await sleep(8000);
  for (const t of tabs) {
    const codes = await t.page.evaluate(() => window.__codes);
    const sockets = await t.page.evaluate(() => window.__ws.length);
    await sleep(3000);
    const stopped = (await t.page.evaluate(() => window.__ws.length)) === sockets && (await openSockets(t.page)) === 0;
    const shown = t.label === 'PDF' ? await t.page.locator('#conn-status').innerText() : (await t.page.locator('body').innerText()).includes('다시 열렸습니다');
    rec(`${t.label} 옛 번들: 재시작 뒤 4409로 거절·다시 붙지 않음`, codes.includes(4409) && stopped, `codes=${codes.join(',')} 표시=${JSON.stringify(shown)}`);
    await t.page.close();
    if (t.k) {
      const fresh = await newPage(browser, t.token);
      await fresh.waitForSelector(t.k.ready, { timeout: 120000 });
      await sleep(2500);
      const after = await t.k.count(fresh);
      await save(t.token);
      const savedAfter = await t.k.saved(await download(t.token), 'OB');
      rec(`${t.label} 옛 번들: 새 탭 내용·저장 파일 한 벌`, after === t.before && JSON.stringify(savedAfter) === JSON.stringify(t.savedBefore), `${t.before} → ${after}, ${JSON.stringify(t.savedBefore)} → ${JSON.stringify(savedAfter)}`);
      await fresh.close();
    }
    await fetch(`${BASE}/api/documents/${t.id}`, { method: 'DELETE' }).catch(() => {});
  }
}

await startServer();
const browser = await chromium.launch();
try {
  for (const [label, k] of Object.entries(KINDS)) await graceAndHalfOpen(browser, label, k);
  await restartScenario(browser);
  await oldBundleScenario(browser);
} finally {
  await browser.close();
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${failed === 0 ? 'ALL CHECKS PASSED' : `${failed} FAILED`} (${results.length - failed}/${results.length})`);
process.exit(failed === 0 ? 0 : 1);
