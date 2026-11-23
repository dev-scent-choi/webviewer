// 기본 열기 = 읽기 전용 편집기 화면(ui=view) — 설계 2026-11-21-readonly-editor-look §4의 브라우저 검사 (1)~(8). (9) 기존 e2e는 따로 돌린다.
//  (1) 리본(모든 탭)·빠른 실행·수식 입력줄이 보이고, 문서를 쓰는 조작은 흐림(disabled·aria-disabled·.ro-disabled), 보기 조작(확대·찾기·메모 표시…)은 그대로
//  (2) 편집 시도(흐린 단추 강제 클릭·타자·Delete·F2·붙여넣기·Ctrl+S/H/B/D/M/I/U·시트 탭·슬라이드 끌기)가 화면·서버(새로 연 화면)를 바꾸지 않음
//  (3) 찾기·확대(보기 탭)·메모 창·시트 탭·섬네일·파일 화면(정보·인쇄·내려받기)  (4) canCopy=false면 복사·우클릭·선택 막힘
//  (5) 다른 탭의 편집이 실시간 반영  (5b) 큰 엑셀(지연 불러오기) 뒤쪽 행의 조건부 서식이 서버 답으로 칠해짐
//  (6) 메인 화면 기본 "열기"(docx·xlsx·pptx → ui=view, pdf는 그대로)·한 줄 단추(리브레오피스·아트크래프트·동시편집·삭제)·공유 폼 기본값·편집 열기  (7) 600px 폭
//  (8) 정원 초과(4429) → 안내 + 자동 이동 — 큰 엑셀로 정원을 실제로 채우는 대신 Playwright routeWebSocket으로 서버가 4429로 닫은 것과 같은 상황을 만든다.
// 이 스크립트가 시험 서버를 직접 띄운다(npm run build:editor 뒤, LibreOffice 필요):
//   [ARTCRAFT_E2E_PRIVATE_DIR=<고객형 원본 폴더>] node scripts/verify-readonly-editor-e2e.mjs [포트=8162, 8160~8169]
// 고객형 문서(verify:artcraft와 같은 변수·파일)는 .audit-run/readonly-editor-e2e-<포트>/로 복사해 (1)·화면 오류만 본다 — 변수가 없으면 SKIP.
// 다 통과하면 .audit-run/readonly-editor-e2e-<포트>/를 지운다. 실패하면 살펴보도록 남긴다.
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { chromium } from 'playwright';

const PORT = Number(process.argv[2] ?? 8162);
// 운영 포트(8000·8010·8020·8100·8110·8130·8148)와 다른 시험(8150~8159) 범위와 겹치지 않게
if (!(PORT >= 8160 && PORT <= 8169)) throw new Error(`시험 포트는 8160~8169만 쓴다(받은 값 ${PORT})`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `readonly-editor-e2e-${PORT}`);
const FILES = {
  docx: 'audit/office/corpus/docx/a1-korean-report.docx',
  xlsx: 'audit/office/corpus/xlsx/x07-multisheet-print.xlsx',
  pptx: 'audit/office/corpus/pptx/a2-korean-deck.pptx',
};
const ORIGINALS = process.env.ARTCRAFT_E2E_PRIVATE_DIR ?? '';
const CUSTOMER = [
  { src: '62c3881f-247e-48ad-aacb-977010a23ef1.docx', name: '고객형.docx', ext: 'docx' },
  { src: '8c76f45a-e15a-40ca-b6be-d9ca5ed63137.pptx', name: '고객형.pptx', ext: 'pptx' },
];
const EDITOR_PAGE = { docx: '/editor.html', xlsx: '/editor-grid.html', pptx: '/editor-slides.html' };
const CF_ROWS = 3000;
const PINK = 'rgb(255, 199, 206)';
const CAPACITY_REASON = '사용자가 많습니다. 잠시 후 다시 시도해주세요.';
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
if (!fs.existsSync('public/artcraft-app/manifest.json')) {
  console.error('public/artcraft-app/manifest.json 없음 — npm run build:artcraft-app 먼저');
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
// Ctrl+C·예외로 끝나도 시험 서버를 남기지 않는다
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
    body: JSON.stringify({ viewer: 'readonly-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);

/** 모두 다른 값, 맨 끝 행만 첫 행과 같다 — 끝으로 건너뛰면 첫 페이지는 구독이 풀려 서버 답(/api/view/xlsx-cf-ranges)이 있어야 칠한다 */
async function cfLazyXlsx(rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('s');
  for (let r = 1; r <= rows; r += 1) ws.getCell(r, 1).value = r === rows ? 'K1' : `K${r}`;
  ws.addConditionalFormatting({ ref: `A1:A${rows}`, rules: [{ type: 'expression', priority: 1, formulae: ['COUNTIF($A:$A,$A1)>1'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
  return Buffer.from(await wb.xlsx.writeBuffer());
}
async function smallPdf() {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  doc.addPage([400, 300]).drawText('readonly e2e pdf', { x: 40, y: 150, size: 20, font });
  return Buffer.from(await doc.save());
}

/**
 * 편집기마다: 준비 표시, 누를 곳, 문서 지문(글자 + 굵게·기울임·밑줄·들여쓰기 — Ctrl+B/I/U/M도 잡히게, 선택 표시는 빼고).
 * 페이지 안에서 돈다.
 */
const fmtOf = (root) =>
  root
    ? [...root.querySelectorAll('*')]
        .filter((el) => el.textContent.trim() && (el.childElementCount === 0 || /^(P|LI|H\d)$/.test(el.tagName)))
        .map((el) => {
          const cs = getComputedStyle(el);
          return `${el.childElementCount === 0 ? el.textContent : ''}/${cs.fontWeight}/${cs.fontStyle}/${cs.textDecorationLine}/${cs.marginLeft}/${cs.paddingLeft}/${cs.textIndent}`;
        })
        .join('|')
    : '';
const ED = {
  docx: {
    ready: '.ProseMirror',
    surface: '.ProseMirror',
    root: '.ProseMirror',
    snap: () => document.querySelector('.ProseMirror')?.textContent ?? '',
  },
  xlsx: {
    ready: '#grid-edit-table td[data-row]',
    surface: '#grid-edit-table td[data-row]',
    root: '#grid-edit-table',
    snap: () =>
      [...document.querySelectorAll('.grid-sheet-tab')].map((t) => t.textContent).join('|') +
      '#' +
      [...document.querySelectorAll('#grid-edit-table td[data-row]')].map((td) => td.textContent).join('|'),
  },
  pptx: {
    ready: '#slides-rail .slide-card',
    surface: '#slide-canvas',
    root: '#slide-canvas',
    snap: () =>
      [...document.querySelectorAll('#slides-rail .slide-card')].map((c) => c.dataset.id).join('|') + '#' + (document.querySelector('#slide-canvas')?.textContent ?? ''),
  },
};
const snapshot = (page, ext) =>
  page.evaluate(([snapSrc, fmtSrc, root]) => `${(0, eval)(`(${snapSrc})`)()}\n${(0, eval)(`(${fmtSrc})`)(document.querySelector(root))}`, [String(ED[ext].snap), String(fmtOf), ED[ext].root]);
/** 보기 꾸밈에서 보이되 흐려야 하는 조작(문서를 쓴다)과 그대로 눌려야 하는 조작(보기만) */
const WRITE_IDS = {
  docx: ['qa-save', 'qa-undo', 'btn-paste', 'font-family', 'btn-bold', 'btn-ribbon-replace', 'btn-table', 'btn-page-break', 'indent-left', 'btn-new-comment', 'btn-delete-comment'],
  xlsx: ['qa-save', 'qa-undo', 'btn-paste', 'cell-font-family', 'btn-cell-bold', 'btn-merge-center', 'btn-add-row', 'btn-find-replace', 'btn-sort-asc', 'btn-freeze-panes', 'sheet-add'],
  pptx: ['qa-save', 'qa-undo', 'btn-add-blank', 'btn-dup-slide', 'btn-del-slide', 'font-family', 'btn-bold', 'btn-add-textbox', 'btn-shape-rect', 'btn-bg-apply-all', 'btn-transition-apply-all'],
};
const VIEW_IDS = {
  docx: ['btn-copy', 'btn-ribbon-find', 'btn-show-comments', 'btn-zoom-in-ribbon', 'btn-zoom-out-ribbon'],
  xlsx: ['btn-copy', 'btn-zoom-in-ribbon', 'btn-zoom-out-ribbon'],
  pptx: ['btn-transition-preview', 'btn-animation-preview'],
};
/** 흐린 단추를 강제로 눌러 본다 — 문서가 바뀌면 안 된다(엑셀은 칸을, 워드는 전체를 고른 채로) */
const FORCE_CLICK = {
  docx: ['btn-bold', 'btn-italic', 'btn-align-center', 'btn-bullet-list', 'btn-clear-format', 'btn-indent', 'qa-undo'],
  xlsx: ['btn-cell-bold', 'btn-merge-center', 'btn-add-row', 'btn-del-col', 'btn-wrap-text', 'sheet-add', 'qa-undo'],
  pptx: ['btn-add-blank', 'btn-dup-slide', 'btn-del-slide', 'btn-add-textbox', 'btn-shape-rect', 'qa-undo'],
};
const MIN_TABS = { docx: 6, xlsx: 6, pptx: 6 };

async function openView(context, url, ext) {
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.goto(`${BASE}${url}`);
  await page.waitForSelector(ED[ext].ready, { timeout: 90000 });
  await page.waitForFunction(() => document.querySelector('#conn-status')?.getAttribute('data-status') === 'synced', null, { timeout: 30000 }).catch(() => {});
  await sleep(1500);
  return { page, errors };
}
const contextMenuPrevented = (page, sel) =>
  page.evaluate((s) => {
    const ev = new MouseEvent('contextmenu', { bubbles: true, cancelable: true });
    document.querySelector(s).dispatchEvent(ev);
    return ev.defaultPrevented;
  }, sel);
const chromeState = (page, ext) =>
  page.evaluate(
    ([write, view]) => {
      const shown = (sel) => [...document.querySelectorAll(sel)].some((el) => el.getClientRects().length > 0 && getComputedStyle(el).visibility !== 'hidden');
      const greyed = (id) => {
        const el = document.getElementById(id);
        return !!el && el.classList.contains('ro-disabled') && el.getAttribute('aria-disabled') === 'true' && (!('disabled' in el) || el.disabled);
      };
      const active = (id) => {
        const el = document.getElementById(id);
        return !!el && !el.classList.contains('ro-disabled') && !el.disabled;
      };
      return {
        body: document.body.classList.contains('ui-view'),
        tabs: [...document.querySelectorAll('.ribbon-tab')].filter((t) => t.getClientRects().length > 0).length,
        panel: shown('.ribbon-panel'),
        qa: shown('.quick-access'),
        formula: document.getElementById('formula-input') ? shown('#formula-input') : true,
        strip: document.querySelectorAll('.view-strip, .view-strip-tools').length,
        notGreyed: write.filter((id) => !greyed(id)),
        notActive: view.filter((id) => !active(id)),
        edit: document.querySelectorAll('.view-open-edit').length,
      };
    },
    [WRITE_IDS[ext], VIEW_IDS[ext]],
  );
// 흐림은 숨은 탭의 판에 있는 조작까지 본다(보이는지는 탭을 눌러 보는 검사가 맡는다)
const chromeOk = (s, ext) => s.body && s.tabs >= MIN_TABS[ext] && s.panel && s.qa && s.formula && s.strip === 0 && s.notGreyed.length === 0;
/** 파일 화면(백스테이지)을 열고 한 칸을 고른다 */
async function backstagePane(page, pane) {
  if (await page.locator('.backstage').isHidden()) await page.click('.ribbon-tab[data-tab="file"]');
  await page.click(`.backstage-item[data-pane="${pane}"]`);
}

await startServer();
let browser;
try {
  const ids = {};
  for (const [ext, file] of Object.entries(FILES)) ids[ext] = (await upload(fs.readFileSync(file), path.basename(file))).id;
  browser = await chromium.launch();

  for (const ext of ['docx', 'xlsx', 'pptx']) {
    const id = ids[ext];
    const tok = await token(id, { engine: 'editor' });
    const viewPath = new URL(tok.viewerUrl, BASE).pathname;
    rec(
      `${ext} 토큰: engine=editor · ${EDITOR_PAGE[ext]}?…&ui=view(보기 화면 아님) · canEdit 강제 끔 · fallbackUrl`,
      tok.engine === 'editor' && viewPath === EDITOR_PAGE[ext] && /&ui=view$/.test(tok.viewerUrl) && typeof tok.fallbackUrl === 'string',
      `${tok.viewerUrl} · fallback ${tok.fallbackUrl}`,
    );
    const forced = await token(id, { engine: 'editor', canEdit: true });
    const forcedPerms = (await (await fetch(`${BASE}/api/view/info?token=${encodeURIComponent(forced.token)}`)).json()).data?.permissions;
    rec(`${ext} 토큰: engine=editor에 canEdit=true를 보내도 ui=view·canEdit=false`, forced.engine === 'editor' && /&ui=view$/.test(forced.viewerUrl) && forcedPerms?.canEdit === false, JSON.stringify(forcedPerms));
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] });
    const { page, errors } = await openView(context, tok.viewerUrl, ext);

    // (1) 리본은 보이고 쓰는 조작만 흐림
    const shown = await chromeState(page, ext);
    rec(`(1) ${ext} 리본(탭 ${MIN_TABS[ext]}개 이상)·빠른 실행·수식 입력줄 보임, 얇은 줄 없음, 쓰는 조작 흐림(${WRITE_IDS[ext].length}개)`, chromeOk(shown, ext), JSON.stringify(shown));
    rec(`(1) ${ext} 보기 조작은 그대로(${VIEW_IDS[ext].join('·')}) · API로 연 화면엔 편집 열기 없음`, shown.notActive.length === 0 && shown.edit === 0, JSON.stringify({ notActive: shown.notActive, edit: shown.edit }));
    {
      // 편집 탭도 눌러 둘러볼 수 있다 — 눌러 보면 판이 바뀌고 그 안 단추는 흐림
      const tab = ext === 'pptx' ? 'transitions' : 'insert';
      await page.click(`.ribbon-tab[data-tab="${tab}"]`);
      const panel = await page.evaluate((t) => {
        const el = document.querySelector(`.ribbon-panel[data-panel="${t}"]`);
        const controls = [...el.querySelectorAll('button, select, input')].filter((c) => c.getClientRects().length > 0);
        return { visible: !el.hidden && el.getClientRects().length > 0, controls: controls.length, greyed: controls.filter((c) => c.classList.contains('ro-disabled')).length };
      }, tab);
      const wantActive = ext === 'pptx' ? 1 : 0; // 전환 미리 보기
      rec(`(1) ${ext} "${tab}" 탭 눌러 보기 — 판이 열리고 쓰는 단추는 모두 흐림`, panel.visible && panel.controls > 0 && panel.greyed === panel.controls - wantActive, JSON.stringify(panel));
      await page.click('.ribbon-tab[data-tab="home"]');
    }

    // (2) 편집 시도
    const before = await snapshot(page, ext);
    await page.click(ED[ext].surface);
    await page.keyboard.type('ZZQ');
    for (const key of ['Delete', 'Backspace', 'F2', 'Control+v', 'Control+s', 'Control+h']) await page.keyboard.press(key);
    rec(`(2) ${ext} Ctrl+H — 바꾸기 줄 안 보임`, !(await page.locator('#search-replace-row').isVisible()));
    await page.keyboard.press('Escape');
    // 서식 단축키 — 엑셀은 Ctrl+D(아래로 채우기)가 뜻이 있게 둘째 행 칸에서
    if (ext === 'xlsx') await page.locator('#grid-edit-table td[data-row="1"][data-col="0"]').click();
    else await page.click(ED[ext].surface);
    if (ext === 'docx') await page.keyboard.press('Control+a');
    // 흐린 단추를 강제로(마우스가 닿지 않게 막혀 있어도 Playwright force로) 눌러 본다
    let pressed = 0;
    for (const id of FORCE_CLICK[ext]) {
      const tab = await page.evaluate((i) => document.getElementById(i)?.closest('.ribbon-panel')?.dataset.panel ?? null, id);
      if (tab) await page.click(`.ribbon-tab[data-tab="${tab}"]`);
      if (await page.locator(`#${id}`).click({ force: true, timeout: 3000 }).then(() => true, () => false)) pressed += 1;
    }
    await page.click('.ribbon-tab[data-tab="home"]');
    await sleep(800);
    const afterForce = await snapshot(page, ext);
    rec(`(2) ${ext} 흐린 단추 강제 클릭(${FORCE_CLICK[ext].join('·')}) — 화면 그대로`, afterForce === before && pressed > 0, `눌린 ${pressed}/${FORCE_CLICK[ext].length} · ${afterForce === before ? '' : '달라짐'}`);
    for (const key of ['Control+b', 'Control+d', 'Control+m', 'Control+i', 'Control+u']) {
      await page.keyboard.press(key);
      await page.keyboard.press('Escape');
    }
    await sleep(800);
    const afterKeys = await snapshot(page, ext);
    rec(`(2) ${ext} Ctrl+B/D/M/I/U — 화면(글자·굵게·기울임·밑줄·들여쓰기) 그대로`, afterKeys === before, afterKeys === before ? '' : `달라짐(전 ${before.length}자 → 뒤 ${afterKeys.length}자)`);
    if (ext === 'xlsx') {
      const tab = page.locator('.grid-sheet-tab').first();
      await tab.click({ button: 'right' });
      rec('(2) xlsx 시트 탭 우클릭 메뉴 없음', (await page.locator('.sheet-context-menu').count()) === 0);
      await tab.dblclick();
      rec('(2) xlsx 시트 탭 두 번 클릭 — 이름 바꾸기 입력창 없음', (await page.locator('.grid-sheet-tab-rename').count()) === 0);
      rec('(2) xlsx 시트 추가(+) 단추는 보이되 흐림', (await page.locator('#sheet-add').isVisible()) && (await page.locator('#sheet-add').isDisabled()));
      await page.locator(ED.xlsx.surface).first().click({ button: 'right' });
      rec('(2) xlsx 칸 우클릭 메뉴 없음', (await page.locator('.grid-context-menu').count()) === 0);
    }
    if (ext === 'pptx') {
      const cards = page.locator('#slides-rail .slide-card');
      if ((await cards.count()) >= 2) await cards.nth(0).dragTo(cards.nth(1));
      rec('(2) pptx 섬네일 끌기 꺼짐·복제/삭제 단추 없음', (await page.locator('.slide-card-actions:visible').count()) === 0 && (await cards.first().getAttribute('draggable')) !== 'true');
      await cards.first().click();
    }
    await sleep(1500);
    const after = await snapshot(page, ext);
    const fresh = await openView(context, (await token(id, { engine: 'editor' })).viewerUrl, ext);
    const onServer = await snapshot(fresh.page, ext);
    await fresh.page.close();
    rec(`(2) ${ext} 편집 시도 뒤 화면·서버(새로 연 화면) 그대로`, after === before && onServer === before && !after.includes('ZZQ'), `화면 ${after === before} · 서버 ${onServer === before} · ${before.length}자`);

    // (3) 찾기·확대·이동·인쇄·내려받기
    if (ext !== 'pptx') {
      const text = before.split('\n')[0];
      const body = ext === 'xlsx' ? text.split('#').slice(1).join('#') : text;
      const term = body.split(/[\s|#]+/).find((w) => w.length >= 2) ?? '';
      await page.click('#search-btn');
      await page.fill('#search-input', term);
      await page.keyboard.press('Enter');
      await sleep(800);
      const count = await page.textContent('#search-count');
      rec(`(3) ${ext} 찾기`, (await page.isVisible('#search-bar')) && /[1-9]/.test(count ?? ''), `${term} → ${count}`);
      rec(`(3) ${ext} 찾기 상자의 바꾸기는 흐리거나 숨김`, !(await page.locator('#search-replace-row').isVisible()) || (await page.locator('#search-replace-all').isDisabled()));
      await page.keyboard.press('Escape');
      const z0 = await page.textContent('#zoom-level');
      await page.click('.ribbon-tab[data-tab="view"]');
      await page.click('#btn-zoom-in-ribbon');
      await sleep(300);
      const z1 = await page.textContent('#zoom-level');
      rec(`(3) ${ext} 보기 탭 확대`, z0 !== z1, `${z0} → ${z1}`);
      await page.click('#btn-zoom-out-ribbon');
      await page.click('.ribbon-tab[data-tab="home"]');
    }
    if (ext === 'docx') {
      // 메모 창은 검토 탭 "메모 표시"로(읽기 전용 — 새 메모는 흐림)
      await page.click('.ribbon-tab[data-tab="review"]');
      await page.click('#btn-show-comments');
      await sleep(300);
      rec('(3) docx 검토 → 메모 표시 — 읽기 전용 메모 창 열림(새 메모 흐림)', (await page.isVisible('#review-pane')) && (await page.locator('#btn-new-comment').isDisabled()));
      await page.click('#btn-show-comments');
      await page.click('.ribbon-tab[data-tab="home"]');
    }
    if (ext === 'xlsx') {
      const t0 = await page.textContent('.grid-sheet-tab.active');
      await page.locator('.grid-sheet-tab').nth(1).click();
      await sleep(800);
      const t1 = await page.textContent('.grid-sheet-tab.active');
      rec('(3) xlsx 시트 탭 눌러 시트 바꾸기', t0 !== t1, `${t0} → ${t1}`);
      await page.locator('.grid-sheet-tab').first().click();
      await sleep(800);
    }
    if (ext === 'pptx') {
      const num = () => page.textContent('#slides-rail .slide-card.active .slide-card-number');
      const n0 = await num();
      await page.locator('#slides-rail .slide-card').nth(1).click();
      await sleep(500);
      const n1 = await num();
      rec('(3) pptx 둘째 섬네일 → 슬라이드 바뀜', n0 !== n1, `${n0} → ${n1}`);
      await page.locator('#slides-rail .slide-card').first().click();
      await sleep(500);
      rec('(3) pptx 섬네일로 이동', (await num())?.trim() === '1', await num());
    }
    // 새 창은 about:blank로 먼저 열고 인쇄 주소로 옮긴다 — 헤드리스는 PDF를 화면에 띄우지 않아 창 주소 대신 그 창의 인쇄 요청 응답을 본다
    await backstagePane(page, 'info');
    rec(`(3) ${ext} 파일 화면 — 보이고 정보에 "읽기 전용"`, (await page.isVisible('.backstage')) && (await page.locator('.backstage-pane[data-pane="info"]').innerText()).includes('읽기 전용'));
    await backstagePane(page, 'print');
    const printResp = context.waitForEvent('response', { predicate: (r) => /\/api\/view\/print\?token=/.test(r.url()), timeout: 60000 }).catch(() => null);
    const [popup] = await Promise.all([context.waitForEvent('page', { timeout: 15000 }), page.click('.backstage-action[data-action="print"]')]);
    const pr = await printResp;
    rec(`(3) ${ext} 인쇄 — 새 창이 인쇄 주소를 받음(200)`, pr !== null && pr.status() === 200 && pr.frame()?.page() === popup, pr ? `${pr.status()} ${pr.headers()['content-type'] ?? ''}` : '요청 없음');
    await popup.close();
    await backstagePane(page, 'download');
    const [download] = await Promise.all([page.waitForEvent('download', { timeout: 60000 }), page.click('.backstage-action[data-action="download"]')]);
    rec(`(3) ${ext} 파일 화면 내려받기`, download.suggestedFilename().length > 0, download.suggestedFilename());
    await page.click('.backstage-back');

    // (4) 복사 권한
    rec(`(4) ${ext} canCopy=true — 브라우저 우클릭 메뉴는 그대로`, !(await contextMenuPrevented(page, ED[ext].surface)));
    {
      const ctx2 = await browser.newContext({ viewport: { width: 1280, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] });
      const noCopy = await openView(ctx2, (await token(id, { engine: 'editor', canCopy: false })).viewerUrl, ext);
      const p = noCopy.page;
      await p.evaluate(() => navigator.clipboard.writeText('SENTINEL'));
      await p.click(ED[ext].surface);
      await p.keyboard.press('Control+a');
      await p.keyboard.press('Control+c');
      await sleep(300);
      const clip = await p.evaluate(() => navigator.clipboard.readText());
      const prevented = await contextMenuPrevented(p, ED[ext].surface);
      const userSelect = await p.evaluate((s) => getComputedStyle(document.querySelector(s)).userSelect, ED[ext].surface);
      rec(`(4) ${ext} canCopy=false — 복사·우클릭·선택 막힘`, clip === 'SENTINEL' && prevented && userSelect === 'none', `${clip.slice(0, 20)} · ${prevented} · ${userSelect}`);
      await ctx2.close();
    }

    // (5) 실시간 반영 — 편집 토큰으로 연 다른 탭에서 고친다
    {
      const edit = await openView(context, (await token(id, { canEdit: true })).editorUrl, ext);
      const mark = `LIVE${Date.now() % 100000}`;
      if (ext === 'docx') {
        await edit.page.click('.ProseMirror');
        await edit.page.keyboard.press('Control+End');
        // 한 번에 넣는다 — 글자마다 치면 첫 글자 뒤 커서가 한 칸 앞으로 돌아가 표지가 갈라지는 일이 있었다(2026-11-21 재작업 중 관찰, 편집 화면 쪽)
        await edit.page.keyboard.insertText(mark);
      } else if (ext === 'xlsx') {
        await edit.page.locator(ED.xlsx.surface).first().click();
        await edit.page.keyboard.type(mark);
        await edit.page.keyboard.press('Enter');
      } else {
        await edit.page.locator('#slide-canvas .slide-textbox').first().click();
        await edit.page.keyboard.press('End');
        await edit.page.keyboard.type(mark);
      }
      // textContent — 워드는 화면에서 먼 본문 블록의 그리기를 건너뛰어(content-visibility: hidden, docxLayout.ts data-far) 맨 아래 글이 innerText에 없을 수 있다
      const seen = await page.waitForFunction((m) => document.body.textContent.includes(m), mark, { timeout: 20000 }).then(() => true, () => false);
      const inEdit = await edit.page.evaluate((m) => document.body.textContent.includes(m), mark);
      rec(`(5) ${ext} 다른 탭의 편집이 보기 꾸밈 화면에 실시간`, seen, `${mark} · 편집 탭에 ${inEdit} · 보기 ${await page.getAttribute('#conn-status', 'data-status')}`);
      await edit.page.close();
    }

    rec(`${ext} 화면 오류 없음`, errors.length === 0, errors.slice(0, 2).join(' | '));
    await context.close();
  }

  // (5b) 큰 엑셀(지연 불러오기) — 뒤쪽 행의 조건부 서식(COUNTIF 중복)은 서버 답이 있어야 칠한다. 칠하거나 "일부 생략"을 알려야 한다
  {
    const cfId = (await upload(await cfLazyXlsx(CF_ROWS), `cf-lazy-${CF_ROWS}.xlsx`)).id;
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const { page } = await openView(context, (await token(cfId, { engine: 'editor' })).viewerUrl, 'xlsx');
    const cfCalls = [];
    page.on('response', (r) => {
      if (r.url().includes('/api/view/xlsx-cf-ranges')) cfCalls.push(r.status());
    });
    const last = CF_ROWS - 1;
    for (let i = 0; i < 30; i += 1) {
      const there = await page.evaluate((r) => {
        const host = document.querySelector('.grid-editor-scroll');
        host.scrollTop = host.scrollHeight;
        const td = document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="0"]`);
        return td?.textContent === 'K1';
      }, last);
      if (there) break;
      await sleep(500);
    }
    const state = await page
      .waitForFunction(
        (r) => {
          const td = document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="0"]`);
          const bg = td ? getComputedStyle(td).backgroundColor : null;
          const notice = !document.getElementById('cf-skip-notice')?.hidden;
          return bg === 'rgb(255, 199, 206)' ? { bg, notice } : false;
        },
        last,
        { timeout: 15000, polling: 300 },
      )
      .then((h) => h.jsonValue(), () => null);
    const fallbackState = state ?? (await page.evaluate((r) => {
      const td = document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="0"]`);
      return { bg: td ? getComputedStyle(td).backgroundColor : null, text: td?.textContent ?? null, notice: !document.getElementById('cf-skip-notice')?.hidden };
    }, last));
    rec(
      `(5b) xlsx ${CF_ROWS}행(지연 불러오기) ${CF_ROWS}행 조건부 서식을 서버 답으로 칠함(읽기 전용 토큰도 403 아님)`,
      state !== null && state.bg === PINK && !cfCalls.includes(403),
      `${JSON.stringify(fallbackState)} · xlsx-cf-ranges 응답 ${cfCalls.length}건 [${[...new Set(cfCalls)].join(',')}]`,
    );
    await context.close();
  }

  // (6) 메인 화면
  {
    const pdfName = 'readonly-e2e.pdf';
    await upload(await smallPdf(), pdfName);
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const page = await context.newPage();
    const rowOf = (name) => page.locator('li.doc-item', { hasText: name });
    const row = () => rowOf(path.basename(FILES.docx));
    const home = async (name = path.basename(FILES.docx)) => {
      await page.goto(`${BASE}/`);
      await rowOf(name).waitFor({ timeout: 30000 });
    };
    for (const ext of ['xlsx', 'pptx']) {
      await home(path.basename(FILES[ext]));
      await rowOf(path.basename(FILES[ext])).getByRole('button', { name: '열기', exact: true }).click();
      const re = new RegExp(`${EDITOR_PAGE[ext].replace('.', '\\.')}\\?token=[^&]+&ui=view$`);
      await page.waitForURL(re, { timeout: 30000 }).catch(() => {});
      const ok = re.test(page.url()) && (await page.waitForSelector(ED[ext].ready, { timeout: 60000 }).then(() => true, () => false));
      rec(`(6) 메인 기본 "열기" ${ext} → 읽기 전용 편집기 화면(ui=view)`, ok && (await page.evaluate(() => document.body.classList.contains('ui-view'))), page.url());
    }
    await home(pdfName);
    await rowOf(pdfName).getByRole('button', { name: '열기', exact: true }).click();
    await page.waitForURL(/\/editor-pdf\.html\?token=[^&]+$/, { timeout: 30000 }).catch(() => {});
    rec('(6) 메인 기본 "열기" pdf → 지금처럼 PDF 편집기(ui=view 없음)', /\/editor-pdf\.html\?token=[^&]+$/.test(page.url()), page.url());
    await home();
    await row().getByRole('button', { name: '열기', exact: true }).click();
    await page.waitForURL(/\/editor\.html\?token=[^&]+&ui=view$/, { timeout: 30000 });
    // 2026-11-22 사용자 지시 — "편집 열기" 단추는 DOM에만 있고 화면에는 보이지 않는다
    const editBtn = await page.waitForSelector('.office-titlebar .view-open-edit', { state: 'attached', timeout: 60000 }).then(() => true, () => false);
    const editHidden = editBtn && (await page.$eval('.view-open-edit', (b) => b.getClientRects().length === 0));
    rec('(6) 메인 기본 "열기" docx → 읽기 전용 편집기 화면 + 편집 열기 단추(DOM에 있고 화면에는 숨김)', editBtn && editHidden && (await page.evaluate(() => document.body.classList.contains('ui-view'))), page.url());
    await page.$eval('.view-open-edit', (b) => b.click());
    await page.waitForURL((u) => u.pathname === '/editor.html' && !u.search.includes('ui=view'), { timeout: 30000 });
    await page.waitForSelector('.ProseMirror', { timeout: 60000 });
    rec('(6) 편집 열기 → 편집 권한 편집기(리본 보임)', !(await page.evaluate(() => document.body.classList.contains('ui-view'))) && (await page.isVisible('.ribbon-tabs')));
    await home();
    const labels = (await row().locator('.doc-info > button, .doc-info > span > button').allTextContents()).map((t) => t.trim());
    const want = ['열기', '리브레오피스로 열기', '아트크래프트로 열기', '동시편집 열기', '삭제'];
    rec('(6) 메인 한 줄 단추(펼침 메뉴 없음): 열기·리브레오피스·아트크래프트·동시편집·삭제', JSON.stringify(labels) === JSON.stringify(want) && (await row().locator('details').count()) === 0, labels.join(','));
    for (const [label, re] of [
      ['리브레오피스로 열기', /\/view\.html\?token=/],
      // 배포 기본 = 앱 화면(설계 2026-11-22-artcraft-app-screen §3, manifest 준비 검사는 맨 위)
      ['아트크래프트로 열기', /\/artcraft-app\.html\?token=/],
      ['동시편집 열기', /\/editor\.html\?token=[^&]+$/],
    ]) {
      await home();
      await row().getByRole('button', { name: label }).click();
      await page.waitForURL(re, { timeout: 30000 }).catch(() => {});
      rec(`(6) 한 줄 단추 → ${label}`, re.test(page.url()), page.url());
    }
    await home(pdfName);
    const pdfLabels = (await rowOf(pdfName).locator('.doc-info > button, .doc-info > span > button').allTextContents()).map((t) => t.trim());
    rec('(6) pdf 줄: 열기·…·PDF 편집·삭제(동시편집 없음)', pdfLabels[0] === '열기' && pdfLabels.includes('PDF 편집') && pdfLabels.at(-1) === '삭제' && !pdfLabels.includes('동시편집 열기'), pdfLabels.join(','));
    await home();
    const form = row().locator('form.token-form');
    rec('(6) 공유 폼 — "읽기 전용 편집기 화면(기본)"이 기본 선택', await form.locator('input[name=engine][value=editor]').isChecked());
    await form.locator('input[name=canEdit]').check();
    rec('(6) 공유 폼 — 동시편집을 켜면 엔진 라디오 꺼짐', await form.locator('input[name=engine][value=editor]').isDisabled());
    await form.locator('input[name=canEdit]').uncheck();
    await form.locator('input[name=viewer]').fill('공유받는 사람');
    await form.locator('input[name=canDownload]').check();
    await form.getByRole('button', { name: '열람 링크 발급' }).click();
    const link = row().locator('.token-result a', { hasText: '읽기 전용 편집기 화면 열기' });
    await link.waitFor({ timeout: 10000 });
    const href = (await link.getAttribute('href')) ?? '';
    rec('(6) 공유 링크 = 보기 꾸밈 주소', /&ui=view$/.test(href), href);
    const other = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const shared = await openView(other, href, 'docx');
    rec('(6) 공유 링크로 연 사람에게는 편집 열기 단추 없음(리본은 보임)', (await shared.page.locator('.view-open-edit').count()) === 0 && (await shared.page.isVisible('.ribbon-tabs')));
    await other.close();
    await context.close();
  }

  // (7) 600px 폭 — 메인 화면에서 연 경우(편집 열기 단추까지 있는 가장 붐비는 제목 줄)도 가로 넘침·겹침 없음
  for (const ext of ['docx', 'xlsx', 'pptx']) {
    const context = await browser.newContext({ viewport: { width: 600, height: 900 } });
    await context.addInitScript((id) => sessionStorage.setItem('docview-edit-doc', id), ids[ext]);
    const { page } = await openView(context, (await token(ids[ext], { engine: 'editor' })).viewerUrl, ext);
    const m = await page.evaluate(() => {
      const right = document.querySelector('.titlebar-right').getBoundingClientRect();
      const title = document.getElementById('doc-title').getBoundingClientRect();
      const items = [...document.querySelectorAll('.office-titlebar :is(button, a, .conn-status, .readonly-badge)')].filter((b) => b.getClientRects().length > 0);
      return {
        scroll: document.documentElement.scrollWidth,
        body: document.body.scrollWidth,
        edit: document.querySelectorAll('.view-open-edit').length,
        inside: items.every((b) => {
          const r = b.getBoundingClientRect();
          return r.left >= -0.5 && r.right <= window.innerWidth + 0.5 && r.height < 40;
        }),
        overlap: title.right > right.left + 0.5,
      };
    });
    rec(`(7) ${ext} 600px — 가로 넘침 없음, 제목 줄 단추 모두 화면 안·한 줄, 문서 이름과 겹치지 않음`, m.scroll <= 600 && m.body <= 600 && m.edit === 1 && m.inside && !m.overlap, JSON.stringify(m));
    await context.close();
  }

  // (8) 정원 초과(4429)
  {
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const page = await context.newPage();
    await page.routeWebSocket(/\/ws\/edit/, (ws) => ws.close({ code: 4429, reason: CAPACITY_REASON }));
    const tok = await token(ids.xlsx, { engine: 'editor' });
    await page.goto(`${BASE}${tok.viewerUrl}`);
    const notice = await page.waitForSelector('.view-fallback-notice', { timeout: 30000 }).then((el) => el.textContent(), () => null);
    await page.waitForURL(/\/viewer\.html\?token=/, { timeout: 15000 }).catch(() => {});
    rec('(8) 정원 초과 → 안내 + 같은 토큰의 보기 화면으로 자동 이동', notice === '지금은 보는 사람이 많습니다. 서버 변환으로 보여 줍니다' && page.url().endsWith(tok.fallbackUrl), `${notice} · ${page.url()}`);
    const page2 = await context.newPage();
    await page2.routeWebSocket(/\/ws\/edit/, (ws) => ws.close({ code: 4429, reason: CAPACITY_REASON }));
    await page2.goto(`${BASE}${(await token(ids.xlsx, { canEdit: true })).editorUrl}`);
    const shownError = await page2.waitForSelector('#error-screen:not([hidden])', { timeout: 30000 }).then(() => true, () => false);
    await sleep(2000);
    rec('(8) 편집 화면은 4429면 지금처럼 오류 화면(자동 이동 없음)', shownError && page2.url().includes('/editor-grid.html'), page2.url());
    await context.close();
  }

  // 고객형 문서 — (1) 리본 보임·쓰는 조작 흐림·화면 오류만(원본은 저장소 밖, 사본은 시험 폴더에)
  for (const c of CUSTOMER) {
    const src = ORIGINALS && path.join(ORIGINALS, c.src);
    if (!src || !fs.existsSync(src)) {
      console.log(`SKIP 고객형 ${c.name} — ARTCRAFT_E2E_PRIVATE_DIR 없음`);
      continue;
    }
    const copy = path.join(RUN, c.name);
    fs.copyFileSync(src, copy);
    const id = (await upload(fs.readFileSync(copy), c.name)).id;
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const { page, errors } = await openView(context, (await token(id, { engine: 'editor' })).viewerUrl, c.ext);
    const s = await chromeState(page, c.ext);
    rec(`고객형 ${c.name} — 보기 꾸밈(리본 보임·쓰는 조작 흐림)·화면 오류 없음`, chromeOk(s, c.ext) && s.notActive.length === 0 && errors.length === 0, `${JSON.stringify(s)} ${errors.slice(0, 1).join('')}`);
    await context.close();
  }
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
