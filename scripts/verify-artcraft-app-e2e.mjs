// 아트크래프트 앱 화면 — 설계 2026-11-22-artcraft-app-screen §5의 브라우저 검사(docx·xlsx·pptx 각각 읽기 전용·편집 토큰):
//  (1) 앱 화면이 뜨고 문서가 보임(워드 document.text·엑셀 시트 수·PPT 슬라이드 수)
//  (2) 읽기 전용: 워드·엑셀 타자 → 변화 없음(엑셀 시트 모두 보호, Ctrl+Z·edit.undo 뒤 1.5초 안에 다시 보호), PPT는 "편집은 저장되지 않습니다" 띠
//  (3) 편집 토큰 Ctrl+S/file.save → 서버 저장 → 받은 파일·다시 연 앱에 반영, 편집 방이 열려 있으면 거절 안내, 옛 기준(다른 창이 먼저 저장)이면 409 안내
//  (4) wasm 막기 → 안내 + 엔진 보기(워드)·서버 격자(엑셀)  (5) 서버 변환 요청 0건  (6) 두 번째 열기에서 앱 자산을 서버에서 다시 받지 않음
//  (7) 우리 제목 줄 24px·한글, 읽기 전용 표지·"편집 열기" 없음(사용자 규칙 2026-11-22), 600px 가로 넘침 없음
// 이 스크립트가 시험 서버를 직접 띄운다(npm run build:editor·build:artcraft·build:artcraft-app 뒤, LibreOffice 필요 — 업로드 변환):
//   node scripts/verify-artcraft-app-e2e.mjs [포트=8172, 8170~8179]
// 합성 문서(audit/office/corpus)만 쓴다. 다 통과하면 .audit-run/artcraft-app-e2e-<포트>/를 지운다. 실패하면 살펴보도록 남긴다.
// 실제 API 모양은 앱 화면 Task 1 보고(document.text는 워드만, 격자 docInspect.sheets[].protected·undo[], 덱 docInspect.slides[]).
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { chromium } from 'playwright';
import { SKIN_PX, heightProblems, latinOnlyLabels } from './ribbonSkinLib.mjs';

const PORT = Number(process.argv[2] ?? 8172);
// 운영 포트(8000·8010·8020·8100·8110·8130·8148)·다른 시험 범위와 겹치지 않게
if (!(PORT >= 8170 && PORT <= 8179)) throw new Error(`시험 포트는 8170~8179만 쓴다(받은 값 ${PORT})`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `artcraft-app-e2e-${PORT}`);
const FILES = {
  docx: 'audit/office/corpus/docx/a1-korean-report.docx',
  xlsx: 'audit/office/corpus/xlsx/x07-multisheet-print.xlsx',
  pptx: 'audit/office/corpus/pptx/a2-korean-deck.pptx',
};
const API = { docx: 'wordcraft', xlsx: 'gridcraft', pptx: 'deckcraft' };
const SERVER_CONVERT = /^\/api\/view\/(pages|derived-pdf|status|search|pdf-)/;
const FAILED_NOTICE = '아트크래프트 앱 화면을 열지 못해 다른 화면으로 보여 줍니다';
const STALE_NOTICE = '문서를 연 뒤에 다른 사람이 먼저 저장했습니다';
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail !== '' ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

for (const [file, how] of [
  ['public/artcraft-app/manifest.json', 'npm run build:artcraft-app'],
  ['public/artcraft/manifest.json', 'npm run build:artcraft'],
  ['public/js/artcraftApp.bundle.js', 'npm run build:editor'],
]) {
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
    // 방 닫기 유예 0 — (3) 편집기를 닫으면 방이 바로 닫혀 저장이 다시 된다
    env: { ...process.env, DOCVIEW_PORT: String(PORT), DOCVIEW_HOST: '127.0.0.1', DOCVIEW_SECRET: randomBytes(24).toString('hex'), DOCVIEW_DATA_DIR: path.join(RUN, 'data'), TMP: path.join(RUN, 'tmp'), TEMP: path.join(RUN, 'tmp'), DOCVIEW_LO_PROFILE_ROOT: path.join(RUN, 'lo'), DOCVIEW_ROOM_CLOSE_GRACE_MS: '0' },
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
async function appHits() {
  await sleep(500);
  return fs.readFileSync(path.join(RUN, 'server.log'), 'utf8').split('\n').filter((line) => /"url":"\/artcraft-app\//.test(line)).length;
}

async function upload(file) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패 ${file}: ${JSON.stringify(r)}`);
  for (let i = 0; i < 1200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready') return st;
    if (st.status === 'failed') throw new Error(`변환 실패 ${file}: ${st.error}`);
    await sleep(250);
  }
  throw new Error(`대기 시간 초과 ${file}`);
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: 'app-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, engine: 'artcraft', ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);

/** 앱 화면을 열어 앱이 문서를 연 상태(data-app-state=open)까지 — 요청 경로·화면 오류를 모은다 */
async function openApp(context, url, { waitOpen = true } = {}) {
  const page = await context.newPage();
  const paths = [];
  page.on('request', (r) => paths.push(new URL(r.url()).pathname));
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.goto(`${BASE}${url}`);
  if (waitOpen) await page.waitForFunction(() => document.body.dataset.appState === 'open', null, { timeout: 120000 });
  return { page, paths, errors };
}
const frameOf = (page) => page.frames().find((f) => f.url().includes('/artcraft-frame.html'));
const exec = (page, api, id, params = {}) => frameOf(page).evaluate(([a, c, p]) => window[a].execute(c, p), [api, id, params]);
const inspect = (page, api) => frameOf(page).evaluate((a) => window[a].docInspect(), api);
const wordText = async (page) => (await exec(page, 'wordcraft', 'document.text')).text;
async function clickCanvas(page, fx = 0.45, fy = 0.55) {
  const box = await page.locator('#app-frame').boundingBox();
  await page.mouse.click(box.x + box.width * fx, box.y + box.height * fy);
}
/** 저장 상태가 끝(saved·failed)에 닿을 때까지 — 그 값을 돌려준다 */
async function saveSettled(page, timeout = 60000) {
  await page.waitForFunction(() => ['saved', 'failed'].includes(document.body.dataset.saveState), null, { timeout }).catch(() => undefined);
  return page.getAttribute('body', 'data-save-state');
}
/** 받은 파일(지금 서버 원본) 안 XML 글자 */
async function savedXmlText(tokenStr) {
  const bytes = Buffer.from(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(tokenStr)}`)).arrayBuffer());
  const zip = await JSZip.loadAsync(bytes);
  const parts = await Promise.all(Object.values(zip.files).filter((f) => f.name.endsWith('.xml')).map((f) => f.async('string')));
  return parts.join('\n');
}
/** 격자 모든 시트가 보호될 때까지(최대 ms) — 걸린 시간, 못 닿으면 null */
async function allProtectedWithin(page, ms) {
  const t0 = Date.now();
  do {
    const d = await inspect(page, 'gridcraft');
    if (d.sheets.every((s) => s.protected === true)) return Date.now() - t0;
    await sleep(50);
  } while (Date.now() - t0 <= ms);
  return null;
}

let browser = null;
let finished = false;
try {
  await startServer();
  // headless는 WebGPU 어댑터가 없어 앱이 WebGL2(SwiftShader)로 떨어진다 — 그것도 동작(spike §2)
  browser = await chromium.launch({ args: ['--enable-unsafe-webgpu'] });
  const ids = {};
  for (const [ext, file] of Object.entries(FILES)) ids[ext] = (await upload(file)).id;

  // (1)(2)(5)(6)(7) 읽기 전용 토큰
  for (const ext of Object.keys(FILES)) {
    const api = API[ext];
    const tok = await token(ids[ext], {});
    rec(`(1) ${ext}: 토큰 viewerUrl이 앱 화면`, /^\/artcraft-app\.html\?token=/.test(tok.viewerUrl ?? ''), tok.viewerUrl);
    // 디스크 캐시가 있는 문맥(실제 브라우저처럼) — 보통 문맥은 메모리 캐시(최대 50 MB, 한 항목 ≤ 1/8)라 워드 wasm(br 8.3 MB)을 담지 못한다
    const context = await chromium.launchPersistentContext(path.join(RUN, `profile-${ext}`), { viewport: { width: 1400, height: 900 }, args: ['--enable-unsafe-webgpu'] });
    const before = await appHits();
    const o = await openApp(context, tok.viewerUrl);
    if (ext === 'docx') {
      const text = await wordText(o.page);
      rec('(1) docx: 문서 글자가 앱에 열림', text.includes('한글 보고서 시험 문서'), text.slice(0, 40));
      await clickCanvas(o.page);
      await o.page.keyboard.type('QQQ');
      await o.page.keyboard.press('Enter');
      await sleep(800);
      const after = await wordText(o.page);
      rec('(2) docx 읽기 전용: 타자해도 글자 그대로', after === text, `${after.length - text.length}자 차이`);
    } else if (ext === 'xlsx') {
      const d = await inspect(o.page, api);
      rec('(1) xlsx: 시트 4개', d.sheets.length === 4, d.sheets.length);
      rec('(2) xlsx 읽기 전용: 모든 시트 보호', d.sheets.every((s) => s.protected === true), JSON.stringify(d.sheets.map((s) => s.protected)));
      const undoBefore = d.undo.length;
      await clickCanvas(o.page, 0.3, 0.4);
      await o.page.keyboard.type('QQQ', { delay: 100 });
      await o.page.keyboard.press('Enter');
      await sleep(800);
      const after = await inspect(o.page, api);
      rec('(2) xlsx 읽기 전용: 타자해도 되돌리기 목록 그대로', after.undo.length === undoBefore, `${undoBefore} → ${after.undo.length}`);
      // 시트 보호는 되돌리기로 풀린다(Task 1) — 키는 막고, 메뉴·명령으로 풀리면 1.5초 안에 다시 건다(Task 3)
      await clickCanvas(o.page, 0.3, 0.4);
      await o.page.keyboard.press('Control+z');
      await sleep(300);
      const z = await allProtectedWithin(o.page, 1500);
      rec('(2) xlsx 읽기 전용: Ctrl+Z 뒤에도 모든 시트 보호', z !== null, z === null ? '풀림' : `${z} ms`);
      const undone = await exec(o.page, api, 'edit.undo').catch((e) => String(e));
      const t = await allProtectedWithin(o.page, 1500);
      rec('(2) xlsx 읽기 전용: edit.undo 뒤 1.5초 안에 다시 보호', t !== null, `${JSON.stringify(undone)} → ${t === null ? '1.5초 넘음' : `${t} ms`}`);
      await clickCanvas(o.page, 0.3, 0.4);
      await o.page.keyboard.type('WWW', { delay: 100 });
      await o.page.keyboard.press('Enter');
      await sleep(800);
      const cells = (await inspect(o.page, api)).sheets.map((s) => s.cells);
      rec('(2) xlsx 읽기 전용: 다시 보호된 뒤 타자해도 칸 수 그대로', JSON.stringify(cells) === JSON.stringify(d.sheets.map((s) => s.cells)), JSON.stringify(cells));
    } else {
      const d = await inspect(o.page, api);
      rec('(1) pptx: 슬라이드 3장', d.slides.length === 3, d.slides.length);
      rec('(2) pptx 읽기 전용: "이 화면의 편집은 저장되지 않습니다" 띠', (await o.page.textContent('#app-banner')) === '이 화면의 편집은 저장되지 않습니다' && (await o.page.isVisible('#app-banner')));
    }
    rec(`(2) ${ext}: 읽기 전용 저장 시도 → 서버에 보내지 않음`, await (async () => {
      const n = o.paths.length;
      await exec(o.page, api, 'file.save').catch(() => undefined);
      await sleep(800);
      return !o.paths.slice(n).some((p) => p === '/api/view/artcraft-save') && (await o.page.getAttribute('body', 'data-save-state')) === 'clean';
    })());
    const converted = o.paths.filter((p) => SERVER_CONVERT.test(p));
    rec(`(5) ${ext}: 서버 변환 요청 0건`, converted.length === 0, converted.join(', '));
    rec(`(1) ${ext}: 화면 오류 없음`, o.errors.length === 0, o.errors.join(' | '));
    // (7) 제목 줄 — 24px, 한글, 읽기 전용 표지·"편집 열기" 없음(사용자 규칙)
    const title = await o.page.$eval('#toolbar .toolbar-row-title', (e) => e.getBoundingClientRect().height);
    const strip = await o.page.evaluate(() => ({
      // 문서 이름(#doc-title)은 파일 이름이라 빼고 우리 띠 글자만
      labels: [document.querySelector('.doclist-link')?.getAttribute('aria-label') ?? '', document.querySelector('.doclist-link')?.getAttribute('title') ?? ''],
      text: document.querySelector('#toolbar')?.textContent ?? '',
      marks: document.querySelectorAll('#readonly-badge, #edit-open').length,
    }));
    rec(`(7) ${ext}: 제목 줄 ${SKIN_PX.title}px`, heightProblems({ title }).length === 0, `${title}px`);
    rec(`(7) ${ext}: 띠 글자 한글`, latinOnlyLabels(strip.labels).length === 0 && strip.labels.every((l) => /[가-힣]/.test(l)), strip.labels.join(' / '));
    rec(`(7) ${ext}: 읽기 전용 표지·"편집 열기" 없음`, strip.marks === 0 && !/읽기 전용|편집 열기/.test(strip.text), strip.text.trim());
    if (ext === 'docx') {
      await o.page.setViewportSize({ width: 600, height: 900 });
      await sleep(300);
      const m = await o.page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, title: document.querySelector('#toolbar .toolbar-row-title').getBoundingClientRect().height }));
      rec('(7) 600px: 가로 넘침 없음, 제목 줄 한 줄', m.scroll <= 600 && heightProblems({ title: m.title }).length === 0, JSON.stringify(m));
      await o.page.screenshot({ path: path.resolve('.audit-run', 'app-screen', 'e2e-docx-600.png') }).catch(() => undefined);
    }
    await o.page.screenshot({ path: path.resolve('.audit-run', 'app-screen', `e2e-${ext}-readonly.png`) }).catch(() => undefined);
    const first = (await appHits()) - before;
    await o.page.close();
    // (6) 같은 문맥에서 다시 열기 — 앱 자산(immutable)을 서버에서 다시 받지 않는다
    const mid = await appHits();
    const again = await openApp(context, tok.viewerUrl);
    const second = (await appHits()) - mid;
    rec(`(6) ${ext}: 두 번째 열기 앱 자산 요청 0건(첫 열기 ${first}건)`, first > 0 && second === 0, `${first} → ${second}`);
    await again.page.close();
    await context.close();
  }

  // (3) 편집 토큰 저장
  {
    const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const tok = await token(ids.docx, { canEdit: true });
    const o = await openApp(context, tok.viewerUrl);
    rec('(3) docx 편집 토큰: 띠 없음', !(await o.page.isVisible('#app-banner')));
    await clickCanvas(o.page);
    await o.page.keyboard.type('ZZQA');
    await sleep(500);
    rec('(3) docx 편집 토큰: 친 글자가 문서에', (await wordText(o.page)).includes('ZZQA'));
    await o.page.keyboard.press('Control+s');
    const st = await saveSettled(o.page);
    rec('(3) docx: Ctrl+S → 저장됨', st === 'saved', `${st} ${await o.page.textContent('#save-status')}`);
    rec('(3) docx: 서버 원본에 친 글자', (await savedXmlText(tok.token)).includes('ZZQA'));
    await o.page.screenshot({ path: path.resolve('.audit-run', 'app-screen', 'e2e-docx-edit-saved.png') }).catch(() => undefined);
    // 다시 열면 저장한 내용
    const re = await openApp(context, tok.viewerUrl);
    rec('(3) docx: 다시 연 앱에 저장한 글자', (await wordText(re.page)).includes('ZZQA'));
    await re.page.close();

    // 옛 기준(409) — 둘째 창(B)이 연 뒤 첫 창(A)이 저장하면 B의 저장은 거절 + 다시 열기 안내
    const b = await openApp(context, tok.viewerUrl);
    await clickCanvas(o.page);
    await o.page.keyboard.type('AAA1');
    await exec(o.page, 'wordcraft', 'file.save');
    const a1 = await saveSettled(o.page);
    await o.page.evaluate(() => (document.body.dataset.saveState = 'pending'));
    await clickCanvas(b.page);
    await b.page.keyboard.type('BBB2');
    await exec(b.page, 'wordcraft', 'file.save');
    const b1 = await saveSettled(b.page);
    const bNote = (await b.page.textContent('#app-notice')) ?? '';
    rec('(3) 옛 기준 저장 → 409 거절 + "먼저 저장" 안내 + 다시 열기 단추', a1 === 'saved' && b1 === 'failed' && bNote.includes(STALE_NOTICE) && (await b.page.getByRole('button', { name: '다시 열기' }).isVisible()), `A ${a1} · B ${b1} · ${bNote}`);
    const xml = await savedXmlText(tok.token);
    rec('(3) 409 뒤 서버 원본 = 먼저 저장한 내용(덮어쓰지 않음)', xml.includes('AAA1') && !xml.includes('BBB2'));
    await b.page.screenshot({ path: path.resolve('.audit-run', 'app-screen', 'e2e-docx-stale-409.png') }).catch(() => undefined);
    await b.page.close();

    // 편집 방이 열려 있으면 거절 + 안내, 닫으면 다시 저장됨
    const editorTok = await fetch(`${BASE}/api/documents/${ids.docx}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: 'app-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: true }) }).then((r) => r.json()).then((j) => j.data);
    const editor = await context.newPage();
    await editor.goto(`${BASE}${editorTok.editorUrl}`);
    await editor.waitForSelector('.ProseMirror', { timeout: 60000 });
    await sleep(1000);
    await o.page.evaluate(() => (document.body.dataset.saveState = 'pending'));
    await exec(o.page, 'wordcraft', 'file.save');
    const r1 = await saveSettled(o.page, 30000);
    const r1Text = (await o.page.textContent('#save-status')) ?? '';
    rec('(3) 편집 방 열림 → 저장 거절 + "편집기" 안내', r1 === 'failed' && /편집기/.test(r1Text), r1Text);
    await editor.close();
    await sleep(1500);
    await o.page.evaluate(() => (document.body.dataset.saveState = 'pending'));
    await exec(o.page, 'wordcraft', 'file.save');
    rec('(3) 편집기를 닫으면 다시 저장됨', (await saveSettled(o.page, 30000)) === 'saved', await o.page.textContent('#save-status'));
    await o.page.close();

    const xt = await token(ids.xlsx, { canEdit: true });
    const x = await openApp(context, xt.viewerUrl);
    rec('(3) xlsx 편집 토큰: 시트 보호 없음', (await inspect(x.page, 'gridcraft')).sheets.every((s) => s.protected !== true));
    await clickCanvas(x.page, 0.3, 0.4);
    await x.page.keyboard.type('ZZQA', { delay: 100 });
    await x.page.keyboard.press('Enter');
    await sleep(500);
    await x.page.keyboard.press('Control+s');
    const xs = await saveSettled(x.page);
    rec('(3) xlsx: Ctrl+S → 저장됨 + 서버 원본에 친 글자', xs === 'saved' && (await savedXmlText(xt.token)).includes('ZZQA'), `${xs} ${await x.page.textContent('#save-status')}`);
    await x.page.screenshot({ path: path.resolve('.audit-run', 'app-screen', 'e2e-xlsx-edit-saved.png') }).catch(() => undefined);
    await x.page.close();

    const pt = await token(ids.pptx, { canEdit: true });
    const p = await openApp(context, pt.viewerUrl);
    rec('(3) pptx 편집 토큰: 띠 없음', !(await p.page.isVisible('#app-banner')));
    await exec(p.page, 'deckcraft', 'file.save');
    const ps = await saveSettled(p.page);
    rec('(3) pptx: file.save → 저장됨', ps === 'saved', `${ps} ${await p.page.textContent('#save-status')}`);
    await p.page.close();
    const reopened = await openApp(context, pt.viewerUrl);
    rec('(3) pptx: 저장 뒤 다시 열어도 슬라이드 3장', (await inspect(reopened.page, 'deckcraft')).slides.length === 3);
    await reopened.page.close();
    await context.close();
  }

  // (4) wasm 막기 → 안내 + 엔진 보기(docx)·서버 격자(xlsx)
  for (const [ext, re] of [['docx', /\/view\.html\?token=/], ['xlsx', /\/viewer\.html\?token=/]]) {
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    await context.route('**/artcraft-app/**/*.wasm', (route) => route.abort());
    const tok = await token(ids[ext], {});
    const o = await openApp(context, tok.viewerUrl, { waitOpen: false });
    const note = await o.page.waitForSelector('.view-fallback-notice', { timeout: 30000 }).then((h) => h.textContent(), () => '');
    await o.page.waitForURL(re, { timeout: 30000 }).catch(() => undefined);
    rec(`(4) ${ext}: wasm 막기 → 안내`, note === FAILED_NOTICE, note);
    rec(`(4) ${ext}: 다른 화면으로 넘어감`, re.test(o.page.url()), o.page.url().replace(/token=[^&]+/, 'token=…'));
    if (ext === 'docx') {
      await o.page.waitForSelector('.page-frame[data-page="1"].page-ready', { timeout: 90000 }).catch(() => undefined);
      rec('(4) docx: 엔진 보기 표지', await o.page.isVisible('#engine-badge'));
    } else {
      rec('(4) xlsx: 서버 격자가 그려짐', await o.page.waitForSelector('#grid-table td[data-r]', { timeout: 90000 }).then(() => true, () => false));
    }
    await context.close();
  }
  finished = true;
} catch (error) {
  rec('스크립트 오류', false, String(error?.stack ?? error));
} finally {
  await browser?.close().catch(() => undefined);
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
if (finished && !failed) {
  await sleep(500);
  fs.rmSync(RUN, { recursive: true, force: true });
} else {
  console.log(`시험 폴더를 살펴보도록 남김: ${RUN}`);
}
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed ? 1 : 0);
