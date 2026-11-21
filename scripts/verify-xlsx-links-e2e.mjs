// 엑셀 하이퍼링크(바로가기) — 보기 화면(viewer.html)·편집 화면(grid-editor) 둘 다 실제 브라우저로 누른다.
//   - 문서 안 위치: 시트 이름(따옴표·띄어쓰기·#), 칸 범위(B10:D12), 정의된 이름, 없는 이름(안내), HYPERLINK 수식, 병합 칸의 링크
//     → 그 시트로 바뀌고(탭이 보이게), 칸·범위가 골라지고(이름 상자), 그 칸이 화면 안에 보여야 한다
//   - 바깥 주소(https·HYPERLINK 수식)는 새 창, 파일·네트워크 경로는 열지 않고 안내
//   - 큰 실제 문서(시트 71개, 이 폴더의 감사용 원본이 있으면): 제목!A1 → 목차!B3, 먼 행·먼 탭·범위 링크
// 이 스크립트가 시험 서버를 직접 띄운다(npm run build:editor 뒤):
//   node scripts/verify-xlsx-links-e2e.mjs [포트=8152] [큰 실제 문서 xlsx 경로]
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';

const PORT = Number(process.argv[2] ?? 8152);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `links-e2e-${PORT}`);
const REAL = process.argv[3] ?? '';
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

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

async function upload(bytes, name) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패: ${JSON.stringify(r)}`);
  for (let i = 0; i < 1200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready') return r.data.id;
    if (st.status === 'failed') throw new Error(`변환 실패: ${JSON.stringify(st)}`);
    await sleep(250);
  }
  throw new Error('대기 시간 초과');
}
const issue = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '링크시험', canPrint: true, canDownload: true, canCopy: true, canEdit: true, ...body }),
  }).then((r) => r.json()).then((j) => j.data);

const esc = (s) => s.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&apos;').replace(/</g, '&lt;');

/** 시험용 통합 문서 — 링크 XML은 직접 쓴다(ExcelJS는 UNC 경로의 \를 지우고 내부 링크에도 r:id를 단다) */
async function makeXlsx() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('요약');
  const texts = ['규약 C40', '#자료 B2500', '자료 범위', '정의된 이름', '없는 이름', null, '외부 사이트', '공유 폴더', null, '병합 링크'];
  texts.forEach((t, i) => t && (ws.getCell(i + 1, 1).value = t));
  ws.getCell('A6').value = { formula: 'HYPERLINK("#\'통신 규약\'!A5","수식 링크")', result: '수식 링크' };
  ws.getCell('A9').value = { formula: 'HYPERLINK("https://example.org/x","외부 수식")', result: '외부 수식' };
  ws.mergeCells('A10:B11');
  // 주소를 칸 값으로 만드는 수식 링크(expr) — 열람만 토큰에는 수식 글을 주지 않으므로 링크가 아니다
  ws.getCell('B12').value = '통신 규약';
  ws.getCell('A12').value = { formula: `HYPERLINK("#'"&B12&"'!A5","동적 링크")`, result: '동적 링크' };
  // 숨긴 행의 링크 — 열람만 토큰에는 오지 않는다
  ws.getCell('A13').value = '숨긴 행 링크';
  ws.getRow(13).hidden = true;
  ws.getColumn(1).width = 16;
  const rule = wb.addWorksheet('통신 규약');
  rule.getCell('A5').value = '수식 목적지';
  rule.getCell('C40').value = '규약 목적지';
  const data = wb.addWorksheet('자료');
  for (let r = 1; r <= 3000; r += 1) data.getCell(r, 1).value = `행 ${r}`;
  data.getCell('B2500').value = '먼 목적지';
  data.getCell('E3000').value = '이름 목적지';
  // 탭이 많아 마지막 탭이 탭 줄 밖에 있게
  for (let i = 1; i <= 30; i += 1) wb.addWorksheet(`채움 시트 ${i}`).getCell('A1').value = i;
  wb.addWorksheet('맨 끝').getCell('Z200').value = '끝 목적지';
  wb.definedNames.add("'자료'!$E$3000", '목표');
  const bytes = new Uint8Array(await wb.xlsx.writeBuffer());
  const entries = unzipSync(bytes);
  const sheetPath = 'xl/worksheets/sheet1.xml';
  const links = [
    ['A1', "'통신 규약'!C40"],
    ['A2', '#자료!B2500'],
    ['A3', '자료!B10:D12'],
    ['A4', '목표'],
    ['A5', '없는이름'],
    ['A10', "'맨 끝'!Z200"],
    ['A13', '자료!A1'],
  ].map(([ref, loc]) => `<hyperlink ref="${ref}" location="${esc(loc)}" tooltip="${ref === 'A1' ? '규약 설명' : ''}"/>`.replace(' tooltip=""', ''));
  links.push('<hyperlink ref="A7" r:id="rIdL1"/>', '<hyperlink ref="A8" r:id="rIdL2"/>');
  let xml = strFromU8(entries[sheetPath]);
  if (!xml.includes('xmlns:r=')) xml = xml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
  xml = xml.replace('<pageMargins', `<hyperlinks>${links.join('')}</hyperlinks><pageMargins`);
  entries[sheetPath] = strToU8(xml);
  const relsPath = 'xl/worksheets/_rels/sheet1.xml.rels';
  const rels = '<Relationship Id="rIdL1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>' +
    '<Relationship Id="rIdL2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="\\\\server\\share\\a.xlsx" TargetMode="External"/>';
  entries[relsPath] = strToU8(entries[relsPath]
    ? strFromU8(entries[relsPath]).replace('</Relationships>', `${rels}</Relationships>`)
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
  return zipSync(entries);
}

const colLetters = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

/** 두 화면의 차이를 여기서만 — 칸 좌표는 0부터 */
const UI = {
  view: {
    table: '#grid-table', scroll: '#grid-scroll', ref: '#grid-formula-ref', value: '#grid-formula-value', tabs: '#sheet-tabs',
    td: (r, c) => `#grid-table td[data-r="${r + 1}"][data-c="${c + 1}"]`,
    selected: '#grid-table td.g-selected', inRange: '#grid-table td.g-range-selected',
  },
  edit: {
    table: '#grid-edit-table', scroll: '.grid-editor-scroll', ref: '#cell-ref-box', value: '#formula-input', tabs: '#sheet-tabs',
    td: (r, c) => `#grid-edit-table td[data-row="${r}"][data-col="${c}"]`,
    selected: '#grid-edit-table td:focus', inRange: '#grid-edit-table td.g-range-selected',
  },
};

async function openPage(context, url, ui) {
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.goto(BASE + url);
  await page.waitForSelector(`${ui.table} td`, { timeout: 60000 });
  await sleep(800);
  return { page, errors };
}

async function gotoSheet(page, ui, name) {
  const tab = page.locator(`${ui.tabs} .grid-sheet-tab`, { hasText: name }).filter({ hasText: new RegExp(`^${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`) }).first();
  await tab.scrollIntoViewIfNeeded();
  await tab.click();
  // 링크가 있는 칸이 왼쪽 위 근처라 맨 위로 — 앞 시험이 남긴 스크롤 위치와 무관하게
  await page.evaluate((sel) => { const el = document.querySelector(sel); el.scrollTop = 0; el.scrollLeft = 0; }, ui.scroll);
  await sleep(900);
}

/** 링크 칸을 눌러 목적지 확인 */
async function clickLink(page, ui, label, src, want) {
  const td = page.locator(ui.td(src.r, src.c)).first();
  await td.scrollIntoViewIfNeeded().catch(() => undefined);
  const look = await td.evaluate((el) => ({ cursor: getComputedStyle(el).cursor, title: el.title, cls: el.className })).catch((e) => ({ err: e.message }));
  await td.click();
  // 먼 행은 페이지를 받아 와야 한다
  let state;
  for (let i = 0; i < 30; i += 1) {
    await sleep(200);
    state = await page.evaluate(({ ui, want, tdSel }) => {
      const tab = document.querySelector(`${ui.tabs} .grid-sheet-tab.active`);
      const bar = document.querySelector(ui.tabs);
      const tr = tab?.getBoundingClientRect();
      const br = bar?.getBoundingClientRect();
      const target = document.querySelector(tdSel);
      const sr = document.querySelector(ui.scroll).getBoundingClientRect();
      const r = target?.getBoundingClientRect();
      const visible = !!r && r.width > 0 && r.top >= sr.top - 1 && r.top < sr.bottom - 10 && r.left >= sr.left - 1 && r.left < sr.right - 10 &&
        document.elementFromPoint(r.left + Math.min(4, r.width / 2), r.top + Math.min(4, r.height / 2))?.closest('td') === target;
      const hit = r ? document.elementFromPoint(r.left + Math.min(4, r.width / 2), r.top + Math.min(4, r.height / 2)) : null;
      return {
        sheet: tab?.textContent?.trim(),
        tabVisible: !!tr && !!br && tr.left >= br.left - 1 && tr.right <= br.right + 1,
        ref: document.querySelector(ui.ref)?.textContent?.trim(),
        value: (document.querySelector(ui.value)?.value ?? document.querySelector(ui.value)?.textContent ?? '').trim(),
        visible,
        selected: !!target && (target.classList.contains('g-selected') || document.activeElement === target),
        inRange: document.querySelectorAll(ui.inRange).length,
        at: visible ? undefined : { top: r && Math.round(r.top - sr.top), left: r && Math.round(r.left - sr.left), h: Math.round(sr.height), w: Math.round(sr.width), hit: hit ? hit.tagName + '.' + hit.className : null },
      };
    }, { ui: { ...ui, td: undefined }, want, tdSel: ui.td(want.r, want.c) });
    if (state.sheet === want.sheet && state.visible && state.selected) break;
  }
  const ref = `${colLetters(want.c)}${want.r + 1}`;
  const ok = state.sheet === want.sheet && state.visible && state.selected && state.ref === ref && state.tabVisible &&
    (want.value === undefined || state.value === want.value) && (want.rangeCells === undefined || state.inRange >= want.rangeCells) &&
    look.cursor === 'pointer' && !!look.title;
  rec(label, ok, JSON.stringify({ ...state, cursor: look.cursor, title: look.title?.slice(0, 40) }));
  if (!ok) await page.screenshot({ path: path.join(RUN, `fail-${results.length}.png`) });
}

async function clickExternal(page, context, ui, label, src, wantUrl) {
  const td = page.locator(ui.td(src.r, src.c)).first();
  const popup = context.waitForEvent('page', { timeout: 5000 }).catch(() => null);
  await td.click();
  const p = await popup;
  const url = p ? p.url() : '(새 창 없음)';
  rec(label, !!p && url.startsWith(wantUrl), url);
  await p?.close();
}

async function clickBlocked(page, ui, label, src, wantText) {
  await page.locator(ui.td(src.r, src.c)).first().click();
  const dlg = page.locator('.office-dialog-runtime');
  const ok = await dlg.waitFor({ timeout: 3000 }).then(() => true, () => false);
  const text = ok ? await dlg.innerText() : '';
  rec(label, ok && text.includes(wantText), text.replace(/\s+/g, ' ').slice(0, 90));
  if (ok) await dlg.locator('button').first().click();
  await sleep(200);
}

await startServer();
const browser = await chromium.launch();
try {
  const synthId = await upload(await makeXlsx(), '링크시험.xlsx');
  const realId = REAL && fs.existsSync(REAL) ? await upload(fs.readFileSync(REAL), '실제문서.xlsx') : null;
  for (const mode of ['view', 'edit']) {
    const ui = UI[mode];
    // 보기 화면은 열람만 토큰(편집·내려받기 없음) — 링크 정보도 그 토큰으로 받아야 한다
    const tokenBody = mode === 'view' ? { canEdit: false, canDownload: false } : {};
    const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    await context.route(/^https:\/\/example\.(com|org)\//, (route) => route.fulfill({ status: 200, contentType: 'text/html', body: '<p>바깥</p>' }));
    const t = await issue(synthId, tokenBody);
    const { page, errors } = await openPage(context, mode === 'view' ? t.viewerUrl : t.editorUrl, ui);
    const back = () => gotoSheet(page, ui, '요약');
    await clickLink(page, ui, `[${mode}] 따옴표·띄어쓰기 시트 '통신 규약'!C40`, { r: 0, c: 0 }, { sheet: '통신 규약', r: 39, c: 2, value: '규약 목적지' });
    await back();
    await clickLink(page, ui, `[${mode}] #자료!B2500 (먼 행 — 스크롤)`, { r: 1, c: 0 }, { sheet: '자료', r: 2499, c: 1, value: '먼 목적지' });
    await back();
    await clickLink(page, ui, `[${mode}] 범위 자료!B10:D12 — 범위 전체 선택`, { r: 2, c: 0 }, { sheet: '자료', r: 9, c: 1, rangeCells: 9 });
    await back();
    await clickLink(page, ui, `[${mode}] 정의된 이름 목표 → 자료!E3000`, { r: 3, c: 0 }, { sheet: '자료', r: 2999, c: 4, value: '이름 목적지' });
    await back();
    await clickBlocked(page, ui, `[${mode}] 없는 이름 — 안내`, { r: 4, c: 0 }, '없는이름');
    await clickLink(page, ui, `[${mode}] HYPERLINK("#'통신 규약'!A5")`, { r: 5, c: 0 }, { sheet: '통신 규약', r: 4, c: 0, value: '수식 목적지' });
    await back();
    await clickLink(page, ui, `[${mode}] 병합 칸 링크 → 맨 끝 탭 'Z200 (탭 줄 스크롤)`, { r: 9, c: 0 }, { sheet: '맨 끝', r: 199, c: 25, value: '끝 목적지' });
    await back();
    await clickExternal(page, context, ui, `[${mode}] 바깥 주소 https — 새 창`, { r: 6, c: 0 }, 'https://example.com/');
    await clickExternal(page, context, ui, `[${mode}] HYPERLINK("https://…") — 새 창`, { r: 8, c: 0 }, 'https://example.org/x');
    await clickBlocked(page, ui, `[${mode}] 네트워크 경로 — 열지 않고 안내`, { r: 7, c: 0 }, '웹 화면에서 열 수 없습니다');
    // Ctrl+클릭은 이동하지 않고 칸만 고른다
    await page.locator(ui.td(0, 0)).first().click({ modifiers: ['Control'] });
    await sleep(400);
    const still = await page.locator(`${ui.tabs} .grid-sheet-tab.active`).innerText();
    rec(`[${mode}] Ctrl+클릭은 이동하지 않음`, still.trim() === '요약', still);
    if (mode === 'view') {
      const isLink = await page.locator(ui.td(11, 0)).first().evaluate((el) => el.classList.contains('g-link'));
      const got = (await (await fetch(`${BASE}/api/view/xlsx-links?token=${encodeURIComponent(t.token)}`)).json()).data.links;
      rec('[view] 열람만 토큰: 수식 링크(expr)·숨긴 행 링크는 받지 않음', !isLink && !got.some((l) => l.expr || l.r1 === 12) && got.length > 0, `링크 ${got.length}개`);
      // 내려받기 권한이 있는 토큰의 보기 화면은 수식 링크도 누른 칸 값으로 풀어 이동한다
      const dt = await issue(synthId, { canEdit: false, canDownload: true });
      const dl = await openPage(context, dt.viewerUrl, ui);
      await clickLink(dl.page, ui, `[view·내려받기] 동적 수식 링크 "#'"&B12&"'!A5"`, { r: 11, c: 0 }, { sheet: '통신 규약', r: 4, c: 0, value: '수식 목적지' });
      await dl.page.close();
    } else {
      await clickLink(page, ui, `[edit] 동적 수식 링크 "#'"&B12&"'!A5"`, { r: 11, c: 0 }, { sheet: '통신 규약', r: 4, c: 0, value: '수식 목적지' });
    }
    rec(`[${mode}] 페이지 오류 없음`, errors.length === 0, errors.join(' | ').slice(0, 200));
    await page.close();

    if (realId) {
      const rt = await issue(realId, tokenBody);
      const real = await openPage(context, mode === 'view' ? rt.viewerUrl : rt.editorUrl, ui);
      await clickLink(real.page, ui, `[${mode}] 실제 문서 제목!A1 → 목차!B3`, { r: 0, c: 0 }, { sheet: '목차', r: 2, c: 1 });
      // 목차의 먼 링크 하나·범위 링크 하나를 원본 링크 표에서 골라 누른다(편집 토큰으로 링크 표를 읽는다)
      const et = await issue(realId, {});
      const ex = (await (await fetch(`${BASE}/api/view/xlsx-extras?token=${encodeURIComponent(et.token)}`)).json()).data;
      const parse = (loc) => {
        const m = /^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)/.exec(loc);
        if (!m) return null;
        let c = 0;
        for (const ch of m[3]) c = c * 26 + ch.charCodeAt(0) - 64;
        return { sheet: (m[1] ?? m[2]).replace(/''/g, "'"), r: Number(m[4]) - 1, c: c - 1 };
      };
      const far = ex.links.find((l) => l.sheetName === '목차' && l.location && (parse(l.location)?.r ?? 0) > 150);
      const range = ex.links.find((l) => l.location && /:[A-Z]+\d+$/.test(l.location));
      for (const [label, link] of [['먼 행', far], ['범위', range]]) {
        if (!link) { rec(`[${mode}] 실제 문서 ${label} 링크 찾기`, false); continue; }
        await gotoSheet(real.page, ui, link.sheetName);
        const want = parse(link.location);
        await clickLink(real.page, ui, `[${mode}] 실제 문서 ${link.sheetName}!${colLetters(link.c1)}${link.r1 + 1} → ${link.location}`, { r: link.r1, c: link.c1 }, want);
      }
      await real.page.screenshot({ path: path.join(RUN, `real-${mode}.png`) });
      rec(`[${mode}] 실제 문서 페이지 오류 없음`, real.errors.length === 0, real.errors.join(' | ').slice(0, 200));
      await real.page.close();
    }
    await context.close();
  }
} finally {
  await browser.close();
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed ? 1 : 0);
