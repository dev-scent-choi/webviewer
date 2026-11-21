// 엑셀 보기·편집 화면 회귀 검사(검토 r2·r3에서 고친 것) — 실제 브라우저로 연다.
//   1) 악의적 차트(공유 맵에 스크립트를 심은 차트)가 실행되지 않고 <img>로 그려진다
//   2) 엑셀 표 덧칠: 정렬(값 섞기) 뒤에도 줄무늬가 행 위치 그대로, 행 삽입 뒤에는 표 범위가 함께 늘어난다
//   3) 무거운 조건부 서식(SUMPRODUCT 40개 × 2만 행): 보기·편집 화면 "일부 생략" 알림, 편집 화면이 멈추지 않는다
//      (첫 칸 1초 안, 300ms 넘는 긴 작업 없음 — 검토 r3 E1)
//   4) 편집 화면 COUNTIF($A:$A,$A1)>1(1,500행·2만 행): 받지 않은 페이지의 중복도 칠하고, 편집으로 중복이 생기고
//      없어지면 따라 바뀐다. 보기 화면과 같은 결과(검토 r3 E2)
//   5) x06 차트(축 숨김 c:delete·tickLblPos none 포함) 보기·편집 화면이 그려진다(스크린샷)
//   6) 409·429·예산 초과를 받은 범위도 잠시 뒤 다시 물어 결국 칠한다(재검토 3)
//   6) 서버 답 경로(검토 r4 C1): 내려받기만 되는 토큰 403, 규칙에 없는 범위 skipped, 64범위 × 동시 4요청 뒤에도 서버가 산다
//   7) 엑셀 표 경계(검토 r4 I2): 머리 줄 삭제는 막고, 데이터 행을 모두 지우면 첫 데이터 행을 비워 남기며(실행 취소·다시 실행
//      버튼), 요약 행을 지우면 요약 행이 꺼진다
//   8) SUMPRODUCT((범위=값)*1)는 칠하고, 엔진이 못 하는 규칙은 "일부 생략"을 알린다(검토 r4)
//   9) SUMPRODUCT((범위=값)*1)는 정확히 같은 값만 — a*·<5·문의?는 다른 값과 맞지 않는다(검토 r5 F1)
//  10) 20만 칸 범위에 2만 칸 붙여넣기 뒤 서버 답(표 다시 만들기) 동안 다른 요청이 100ms 넘게 기다리지 않는다(검토 r5 F2)
//
// 사용법(서버를 먼저 띄워 둘 것 — 업로드·토큰 API가 키 없이 열려 있는 로컬 서버):
//   node scripts/verify-xlsx-view.mjs [http://127.0.0.1:8144] [스크린샷 폴더]
// 시험 파일은 매번 ExcelJS로 만든다(x06은 저장소의 감사용 합성 파일).
import ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { chromium } from 'playwright';
import WebSocket from 'ws';
import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8144';
const OUT = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-view-'));
fs.mkdirSync(OUT, { recursive: true });
const results = [];
const check = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const PINK = 'rgb(255, 199, 206)';

async function upload(name, bytes) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
  const id = (await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json()).data.id;
  for (let i = 0; ; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data;
    if (st.status === 'ready') break;
    if (st.status === 'failed' || i > 600) throw new Error(`변환 실패: ${name}`);
    await wait(300);
  }
  const tok = (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: 'q', canPrint: true, canDownload: true, canCopy: true, canEdit: true }) })).json()).data;
  return { id, tok, token: new URL(BASE + tok.editorUrl).searchParams.get('token') };
}

/** 편집 방에 붙는 Node 쪽 협업 참여자(다른 사람이 고친 것처럼 칸을 바꾼다) */
function connect(id, token) {
  const ydoc = new Y.Doc();
  const provider = new WebsocketProvider(BASE.replace(/^http/, 'ws') + '/ws/edit', id, ydoc, { params: { token }, WebSocketPolyfill: WebSocket });
  return new Promise((resolve) => provider.on('sync', (s) => s && resolve({ ydoc, provider })));
}

async function openEditor(page, tok) {
  await page.goto(BASE + tok.editorUrl + '&home=hidden');
  await page.waitForSelector('#grid-edit-table td[data-row]', { timeout: 60_000 });
}
const bgOf = (page, sel) => page.evaluate((s) => [...document.querySelectorAll(s)].map((td) => getComputedStyle(td).backgroundColor), sel);
const editBg = (page, row, col = 0) => page.evaluate(([r, c]) => {
  const td = document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="${c}"]`);
  return td ? getComputedStyle(td).backgroundColor : null;
}, [row, col]);

/** 수식 엔진으로만 되는 MATCH 규칙 count개. span: 규칙이 읽는 범위 행 수 — 2만이면 너무 무거워 건너뛰고(알림), 500이면 나눠 칠한다.
 * formula를 주면 그 수식(i는 규칙 번호)으로 */
async function heavyCfXlsx(span = 20000, count = 40, formula) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('s');
  for (let r = 1; r <= 20000; r += 1) ws.getCell(r, 1).value = `코드${r % 997}`;
  const b = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
  const rules = Array.from({ length: count }, (_, i) => `<conditionalFormatting sqref="A1:A20000"><cfRule type="expression" dxfId="0" priority="${i + 1}"><formula>${formula ? formula(i) : `MATCH($A1,$A$1:$A$${span},0)&gt;${i}`}</formula></cfRule></conditionalFormatting>`).join('');
  b['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(b['xl/worksheets/sheet1.xml']).replace('</sheetData>', `</sheetData>${rules}`));
  b['xl/styles.xml'] = strToU8(strFromU8(b['xl/styles.xml']).replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'));
  return Buffer.from(zipSync(b));
}

async function countifXlsx(rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('s');
  // 모두 다른 값, 맨 끝 행만 첫 행과 같다(편집 화면이 처음엔 받지 않는 페이지)
  for (let r = 1; r <= rows; r += 1) ws.getCell(r, 1).value = r === rows ? 'K1' : `K${r}`;
  ws.addConditionalFormatting({ ref: `A1:A${rows}`, rules: [{ type: 'expression', priority: 1, formulae: ['COUNTIF($A:$A,$A1)>1'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

const browser = await chromium.launch();
try {
  // 1) 악의적 차트
  {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('S').addRows([['a', 1], ['b', 2]]);
    const { id, tok, token } = await upload('evil.xlsx', Buffer.from(await wb.xlsx.writeBuffer()));
    const { ydoc, provider } = await connect(id, token);
    const EVIL = '"/><img src=x onerror="window.__xss=1">';
    ydoc.getMap('gridCharts').set('evil', {
      id: 'evil', sheetName: 'S', kind: 'column', title: EVIL, from: { row: 3, col: 3, dx: 0, dy: 0 }, size: { width: 300, height: 200 }, to: null,
      categories: { ref: null, cache: [EVIL, 'b'] }, series: [{ name: EVIL, nameRef: null, values: { ref: null, cache: [1, 2] }, look: { color: EVIL, pointColors: [EVIL] } }],
      source: 'new', look: { palette: [EVIL], legend: 'r' },
    });
    await wait(1500);
    const stored = ydoc.getMap('gridCharts').get('evil');
    check('차트: 서버가 공유 맵의 차트 모양(look)을 지운다', stored && stored.look === undefined && stored.series[0].look === undefined, JSON.stringify(stored?.look ?? null));
    const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
    const dialogs = [];
    page.on('dialog', (d) => {
      dialogs.push(d.message());
      void d.dismiss();
    });
    await openEditor(page, tok);
    await wait(2500);
    const xss = await page.evaluate(() => window.__xss);
    const imgs = await page.evaluate(() => [...document.querySelectorAll('.grid-chart img')].map((i) => (i.getAttribute('onerror') ? 'ONERROR' : i.className)));
    check('차트: 악의적 차트가 스크립트로 실행되지 않는다', xss === undefined && dialogs.length === 0, `imgs=${imgs.join(',')}`);
    check('차트: <img>(data: SVG)로 그린다', imgs.length > 0 && imgs.every((c) => c === 'grid-chart-img'));
    await page.screenshot({ path: `${OUT}/1-evil-chart-edit.png` });
    await page.close();
    provider.destroy();
  }

  // 2) 엑셀 표 덧칠 — 정렬 뒤·행 삽입 뒤
  {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('T');
    ws.addTable({ name: 'T1', ref: 'A1', headerRow: true, style: { theme: 'TableStyleMedium2', showRowStripes: true }, columns: [{ name: '이름' }, { name: '값' }], rows: [['c', 3], ['a', 1], ['b', 2], ['d', 4]] });
    const { id, tok, token } = await upload('table.xlsx', Buffer.from(await wb.xlsx.writeBuffer()));
    const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
    await page.goto(BASE + tok.viewerUrl);
    await page.waitForSelector('#grid-table td[data-r]');
    await wait(800);
    const viewBg = await bgOf(page, '#grid-table td[data-c="1"][data-r]');
    check('표: 보기 화면 머리 줄·줄무늬', viewBg[0] !== 'rgba(0, 0, 0, 0)' && viewBg[1] !== viewBg[2], viewBg.slice(0, 5).join(' | '));
    await page.screenshot({ path: `${OUT}/2-table-view.png` });
    await openEditor(page, tok);
    await wait(2500);
    const before = await bgOf(page, '#grid-edit-table td[data-col="0"][data-row]');
    check('표: 편집 화면 머리 줄·줄무늬(덧칠)', before[0] !== 'rgba(0, 0, 0, 0)' && before[1] !== before[2], before.slice(0, 6).join(' | '));
    const { ydoc, provider } = await connect(id, token);
    const cells = ydoc.getMap('cells');
    ydoc.transact(() => {
      [['a', '1'], ['b', '2'], ['c', '3'], ['d', '4']].forEach(([n, v], i) => {
        cells.set(`0:${i + 1}:0`, n);
        cells.set(`0:${i + 1}:1`, v);
      });
    });
    await wait(1500);
    const after = await bgOf(page, '#grid-edit-table td[data-col="0"][data-row]');
    const names = await page.evaluate(() => [1, 2, 3, 4].map((r) => document.querySelector(`#grid-edit-table td[data-row="${r}"][data-col="0"]`)?.textContent));
    check('표: 정렬 뒤에도 줄무늬가 행 위치 그대로', JSON.stringify(after.slice(0, 6)) === JSON.stringify(before.slice(0, 6)) && names.join('') === 'abcd', `${names.join('')} ${after.slice(0, 6).join(' | ')}`);
    const leaked = [...ydoc.getMap('styles').entries()].filter(([, s]) => s && s.bg);
    check('표: 표 색이 칸 서식(공유 맵)에 들어가지 않는다', leaked.length === 0, JSON.stringify(leaked.slice(0, 2)));
    await page.screenshot({ path: `${OUT}/2-table-edit-sorted.png` });
    // 행 삽입: 표 안(3행 위)에 한 줄 → 표가 A1:B6으로 늘어 6행(0부터 5)까지 칠하고, 그 아래(6)는 칠하지 않는다
    await page.click('#grid-edit-table td[data-row="2"][data-col="0"]');
    await page.click('#btn-add-row');
    await wait(2500);
    const inserted = await bgOf(page, '#grid-edit-table td[data-col="0"][data-row]');
    const transparent = (c) => c === 'rgba(0, 0, 0, 0)' || c === 'rgb(255, 255, 255)';
    check('표: 행 삽입 뒤 표 범위가 함께 늘어난다(마지막 줄까지 덧칠, 그 아래는 없음)', !transparent(inserted[0]) && !transparent(inserted[5]) && transparent(inserted[6]), inserted.slice(0, 8).join(' | '));
    const extras = (await (await fetch(`${BASE}/api/view/xlsx-extras?token=${encodeURIComponent(token)}`)).json()).data;
    check('표: 서버가 알려 주는 표 범위도 옮겨졌다', extras.tables?.[0]?.r1 === 5, JSON.stringify(extras.tables?.[0] ?? null));
    await page.screenshot({ path: `${OUT}/2-table-edit-inserted.png` });
    await page.close();
    provider.destroy();
  }

  // 3) 무거운 조건부 서식 — 생략 알림, 편집 화면이 멈추지 않는다(E1)
  {
    const { id, tok, token } = await upload('heavy.xlsx', await heavyCfXlsx());
    const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
    await page.goto(BASE + tok.viewerUrl);
    await page.waitForSelector('#grid-table td[data-r]');
    await wait(800);
    const shown = await page.evaluate(() => {
      const el = document.getElementById('grid-cf-notice');
      return el && !el.hidden ? el.textContent : null;
    });
    check('무거운 조건부 서식: 보기 화면 "일부 생략" 알림', Boolean(shown), shown ?? '');
    await page.screenshot({ path: `${OUT}/3-cf-notice-view.png` });
    // 편집 방을 미리 열어 둔다(2만 행 읽기는 서버 몫 — 화면이 멈추는지만 잰다)
    const warm = await connect(id, token);
    await page.addInitScript(() => {
      window.__long = [];
      new PerformanceObserver((list) => list.getEntries().forEach((e) => window.__long.push(Math.round(e.duration)))).observe({ type: 'longtask', buffered: true });
      new MutationObserver((_, obs) => {
        if (document.querySelector('#grid-edit-table td[data-row]')) {
          window.__firstCell = performance.now();
          obs.disconnect();
        }
      }).observe(document, { childList: true, subtree: true });
    });
    await openEditor(page, tok);
    const sync = await page.evaluate(() => performance.getEntriesByType('navigation')[0]?.domContentLoadedEventEnd ?? 0);
    await wait(6000);
    const m = await page.evaluate(() => ({ first: window.__firstCell, long: window.__long, notice: !document.getElementById('cf-skip-notice')?.hidden }));
    const maxLong = Math.max(0, ...m.long);
    check('무거운 조건부 서식: 편집 화면 "일부 생략" 알림', m.notice);
    check('무거운 조건부 서식: 편집 화면 첫 칸 1초 안', m.first !== undefined && m.first < 1000, `첫 칸 ${Math.round(m.first)}ms(DOMContentLoaded ${Math.round(sync)}ms)`);
    check('무거운 조건부 서식: 300ms 넘는 긴 작업 없음', maxLong < 300, `긴 작업 ${m.long.length}개, 가장 긴 것 ${maxLong}ms`);
    // 스크롤해도 멈추지 않는다
    const t0 = Date.now();
    await page.mouse.move(600, 400);
    for (let i = 0; i < 10; i += 1) await page.mouse.wheel(0, 2000);
    await wait(1500);
    const m2 = await page.evaluate(() => window.__long);
    const maxLong2 = Math.max(0, ...m2);
    check('무거운 조건부 서식: 스크롤 중에도 300ms 넘는 긴 작업 없음', maxLong2 < 300, `가장 긴 것 ${maxLong2}ms (${Date.now() - t0}ms 동안)`);
    await page.screenshot({ path: `${OUT}/3-cf-notice-edit.png` });
    await page.close();
    warm.provider.destroy();
  }

  // 3b) 칠할 수 있는 무거운 규칙(MATCH 500칸 범위 × 30개 — 칸마다 수식 엔진) — 나눠 칠하므로 멈추지 않고 칠한다(E1)
  {
    const { id, tok, token } = await upload('heavy-500.xlsx', await heavyCfXlsx(500, 30));
    const warm = await connect(id, token);
    const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
    await page.addInitScript(() => {
      window.__long = [];
      new PerformanceObserver((list) => list.getEntries().forEach((e) => window.__long.push(Math.round(e.duration)))).observe({ type: 'longtask', buffered: true });
    });
    await openEditor(page, tok);
    await wait(5000);
    const m = await page.evaluate(() => window.__long);
    const painted = await editBg(page, 3);
    check('나눠 칠하기: MATCH 500칸 × 30개 규칙을 칠한다', painted === PINK, String(painted));
    check('나눠 칠하기: 300ms 넘는 긴 작업 없음', Math.max(0, ...m) < 300, `긴 작업 ${m.length}개, 가장 긴 것 ${Math.max(0, ...m)}ms`);
    await page.screenshot({ path: `${OUT}/3b-cf-sliced-edit.png` });
    await page.close();
    warm.provider.destroy();
  }

  // 3c) SUMPRODUCT((범위=값)*1) 30개 × 2만 행 — COUNTIF로 바꿔 계산하므로 칠하고, 멈추지 않는다(검토 r4)
  {
    const { id, tok, token } = await upload('sumproduct.xlsx', await heavyCfXlsx(20000, 30, (i) => `SUMPRODUCT(($A$1:$A$20000=$A1)*1)&gt;${i + 1}`));
    const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
    await page.goto(BASE + tok.viewerUrl);
    await page.waitForSelector('#grid-table td[data-r]');
    await wait(1000);
    const view = await bgOf(page, '#grid-table td[data-c="1"][data-r="1"]');
    check('SUMPRODUCT 배열 비교: 보기 화면이 칠한다(코드1은 20번 나옴)', view[0] === PINK, view.join(''));
    const warm = await connect(id, token);
    await page.addInitScript(() => {
      window.__long = [];
      new PerformanceObserver((list) => list.getEntries().forEach((e) => window.__long.push(Math.round(e.duration)))).observe({ type: 'longtask', buffered: true });
    });
    await openEditor(page, tok);
    await wait(4000);
    const m = await page.evaluate(() => window.__long);
    const edit = await editBg(page, 0);
    check('SUMPRODUCT 배열 비교: 편집 화면도 칠한다(받지 않은 행은 서버 답)', edit === PINK, String(edit));
    check('SUMPRODUCT 배열 비교: 300ms 넘는 긴 작업 없음', Math.max(0, ...m) < 300, `가장 긴 것 ${Math.max(0, ...m)}ms`);
    await page.screenshot({ path: `${OUT}/3c-sumproduct-edit.png` });
    await page.close();
    warm.provider.destroy();
  }

  // 8) 엔진이 못 하는 규칙 — 칠하지 않고 알린다
  {
    const b = unzipSync(new Uint8Array(await (async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('s').addRows([['a', 1], ['b', 2], ['a', 3]]);
      return wb.xlsx.writeBuffer();
    })()));
    b['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(b['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1:A3"><cfRule type="expression" dxfId="0" priority="1"><formula>SUMPRODUCT(($A$1:$A$3=$A1)*($B$1:$B$3&gt;1))&gt;0</formula></cfRule></conditionalFormatting>'));
    b['xl/styles.xml'] = strToU8(strFromU8(b['xl/styles.xml']).replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'));
    const { tok } = await upload('unsupported.xlsx', Buffer.from(zipSync(b)));
    const page = await browser.newPage({ viewport: { width: 1000, height: 500 } });
    await page.goto(BASE + tok.viewerUrl);
    await page.waitForSelector('#grid-table td[data-r]');
    await wait(800);
    const viewNotice = await page.evaluate(() => {
      const el = document.getElementById('grid-cf-notice');
      return Boolean(el && !el.hidden);
    });
    await openEditor(page, tok);
    await wait(2500);
    const editNotice = await page.evaluate(() => !document.getElementById('cf-skip-notice')?.hidden);
    check('엔진이 못 하는 규칙: 보기·편집 화면 "일부 생략" 알림', viewNotice && editNotice, `보기 ${viewNotice}, 편집 ${editNotice}`);
    await page.close();
  }

  // 9) F1 — 정확히 같은 값만
  {
    const b = unzipSync(new Uint8Array(await (async () => {
      const wb = new ExcelJS.Workbook();
      wb.addWorksheet('s').addRows([['a*'], ['abc'], ['<5'], [3], [1], ['문의?'], ['문의1'], ['x'], ['ABC'], ['abc']]);
      return wb.xlsx.writeBuffer();
    })()));
    b['xl/worksheets/sheet1.xml'] = strToU8(strFromU8(b['xl/worksheets/sheet1.xml']).replace('</sheetData>', '</sheetData><conditionalFormatting sqref="A1:A10"><cfRule type="expression" dxfId="0" priority="1"><formula>SUMPRODUCT(($A$1:$A$10=$A1)*1)&gt;1</formula></cfRule></conditionalFormatting>'));
    b['xl/styles.xml'] = strToU8(strFromU8(b['xl/styles.xml']).replace('<dxfs count="0"/>', '<dxfs count="1"><dxf><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>'));
    const { tok } = await upload('exact.xlsx', Buffer.from(zipSync(b)));
    const page = await browser.newPage({ viewport: { width: 1000, height: 600 } });
    await page.goto(BASE + tok.viewerUrl);
    await page.waitForSelector('#grid-table td[data-r]');
    await wait(800);
    const view = await bgOf(page, Array.from({ length: 10 }, (_, r) => `#grid-table td[data-c="1"][data-r="${r + 1}"]`).join(', '));
    await openEditor(page, tok);
    await wait(2500);
    const edit = await Promise.all(Array.from({ length: 10 }, (_, r) => editBg(page, r)));
    // 칠해야 하는 것: abc·ABC·abc(2·9·10행)만
    const want = [false, true, false, false, false, false, false, false, true, true];
    const ok = (list) => list.every((c, r) => (c === PINK) === want[r]);
    check('F1: SUMPRODUCT 정확 비교 — a*·<5·문의?는 칠하지 않고 abc·ABC만(보기·편집)', ok(view) && ok(edit), `보기 ${view.map((c) => (c === PINK ? 1 : 0)).join('')} 편집 ${edit.map((c) => (c === PINK ? 1 : 0)).join('')}`);
    await page.close();
  }

  // 10) F2 — 큰 붙여넣기 뒤 서버 답
  {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('s');
    for (let r = 1; r <= 20_000; r += 1) ws.addRow(Array.from({ length: 10 }, (_, c) => `v${r}_${c}`));
    ws.addConditionalFormatting({ ref: 'A1:J20000', rules: [{ type: 'expression', priority: 1, formulae: ['COUNTIF($A:$J,A1)>1'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
    const { id, token } = await upload('paste.xlsx', Buffer.from(await wb.xlsx.writeBuffer()));
    const { ydoc, provider } = await connect(id, token);
    const ask = (crit) => fetch(`${BASE}/api/view/xlsx-cf-ranges?token=${encodeURIComponent(token)}`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ sheet: 0, ranges: [{ r0: 0, c0: 0, r1: 19_999, c1: 9, crit }] }) }).then((r) => r.json());
    await ask(['v1_0']); // 처음 센다
    ydoc.transact(() => {
      const cells = ydoc.getMap('cells');
      for (let r = 0; r < 20_000; r += 1) cells.set(`0:${r}:1`, String(500_000 + r));
    });
    await wait(4000); // 붙여넣기 반영(웹소켓 업데이트 풀기 — 조건부 서식과 무관한 기존 비용)이 끝나기를 기다린다
    let worst = 0;
    let pinging = true;
    const pings = (async () => {
      while (pinging) {
        const t0 = Date.now();
        await fetch(`${BASE}/api/formats`);
        worst = Math.max(worst, Date.now() - t0);
        await wait(10);
      }
    })();
    const a = await ask(['500000', 'v1_0']);
    pinging = false;
    await pings;
    const r = a.data?.ranges?.[0];
    check('F2: 2만 칸 붙여넣기 뒤 서버 답(표 다시 만들기) 동안 다른 요청 지연 100ms 미만, 답이 맞다', worst < 100 && r?.agg?.count === 20_000 && r?.agg?.min === 500_000, `다른 요청 최대 ${worst}ms, count ${r?.agg?.count}`);
    provider.destroy();
  }

  // 6) 서버 답 경로(C1)
  {
    const { id, tok, token } = await upload('c1.xlsx', await countifXlsx(20_000));
    const warm = await connect(id, token);
    const dl = (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: 'd', canPrint: false, canDownload: true, canCopy: false, canEdit: false }) })).json()).data;
    const dlToken = new URL(BASE + dl.viewerUrl).searchParams.get('token');
    const ask = (t, body) => fetch(`${BASE}/api/view/xlsx-cf-ranges?token=${encodeURIComponent(t)}`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) });
    const ranges = Array.from({ length: 64 }, (_, i) => ({ r0: i, c0: 0, r1: 19_999, c1: 0, crit: ['K1'] }));
    check('서버 답: 내려받기만 되는 토큰은 403', (await ask(dlToken, { sheet: 0, ranges })).status === 403);
    const t0 = Date.now();
    // 문서마다 답 만들기는 동시에 하나(하나는 뒤에서 기다림) — 나머지는 409(편집 화면은 "일부 생략"으로 두고 다음에 다시 묻는다)
    const replies = await Promise.all(Array.from({ length: 4 }, () => ask(token, { sheet: 0, ranges }).then(async (r) => ({ status: r.status, body: await r.json() }))));
    const answers = replies.filter((r) => r.status === 200).map((r) => r.body);
    const okShape = answers.length >= 1 && answers.every((a) => a.data && 'cells' in a.data.ranges[0] && a.data.ranges.slice(1).every((r) => r.skipped)) && replies.every((r) => r.status === 200 || r.status === 409);
    const bytes = Math.max(...answers.map((a) => JSON.stringify(a).length));
    check('서버 답: 64범위 × 동시 4요청 — 규칙 범위만 답하고 나머지는 skipped, 답이 작다, 겹친 요청은 409', okShape && bytes < 20_000, `${Date.now() - t0}ms, 답 ${answers.length}개·409 ${replies.length - answers.length}개, 답 최대 ${bytes}바이트`);
    check('서버 답: 그 뒤에도 서버가 응답한다', (await fetch(`${BASE}/api/formats`)).ok);
    warm.provider.destroy();
    void tok;
  }

  // 7) 엑셀 표 경계(I2) — 편집 화면 단추로
  {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('T');
    ws.addTable({ name: 'T1', ref: 'A1', headerRow: true, totalsRow: true, style: { theme: 'TableStyleMedium2', showRowStripes: true }, columns: [{ name: '이름', totalsRowLabel: '합계' }, { name: '값', totalsRowFunction: 'sum' }], rows: [['a', 1], ['b', 2], ['c', 3]] });
    const { tok, token } = await upload('table-edge.xlsx', Buffer.from(await wb.xlsx.writeBuffer()));
    const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
    const extras = async () => (await (await fetch(`${BASE}/api/view/xlsx-extras?token=${encodeURIComponent(token)}`)).json()).data.tables[0] ?? null;
    const text = (r, c = 0) => page.evaluate(([row, col]) => document.querySelector(`#grid-edit-table td[data-row="${row}"][data-col="${col}"]`)?.textContent ?? null, [r, c]);
    const dialog = async (ok = true) => {
      const card = await page.waitForSelector('.office-dialog-runtime .office-dialog-message', { timeout: 5000 }).catch(() => null);
      const msg = card ? await card.textContent() : null;
      if (card) await page.click(ok ? '.office-dialog-runtime .btn-mini-primary' : '.office-dialog-runtime .btn-mini:not(.btn-mini-primary)');
      return msg;
    };
    await openEditor(page, tok);
    await wait(2500);
    // 머리 줄(1행) 삭제 → 막는다
    await page.click('#grid-edit-table td[data-row="0"][data-col="0"]');
    await page.click('#btn-del-row');
    const refused = await dialog();
    await wait(800);
    check('표 경계: 머리 줄 삭제는 막고 알린다', Boolean(refused?.includes('머리글 행은 삭제할 수 없습니다')) && (await extras())?.r1 === 4, String(refused));
    // 데이터 행 전부(2~4행) 삭제 → 첫 데이터 행을 비워 남긴다
    await page.click('#grid-edit-table td[data-row="1"][data-col="0"]');
    await page.click('#grid-edit-table td[data-row="3"][data-col="1"]', { modifiers: ['Shift'] });
    await page.click('#btn-del-row');
    await dialog(); // 값이 있어 "삭제할까요?" 확인
    await wait(2500);
    const afterDel = await extras();
    check('표 경계: 데이터 행을 모두 지우면 빈 데이터 행 하나를 남긴다', afterDel?.r1 === 2 && afterDel?.totalsRows === 1 && (await text(1)) === '' && (await text(2)) === '합계', `${JSON.stringify(afterDel)} 2행="${await text(1)}" 3행="${await text(2)}"`);
    // 실행 취소·다시 실행 단추 — 행 삭제는 되돌릴 수 없다(삭제 확인 창이 알린다). 구조가 바뀌면 실행 취소 기록을 버리므로
    // 단추를 눌러도 다른 칸에 옛 값이 써지지 않아야 한다
    const tableBefore = JSON.stringify(await extras());
    await page.click('#qa-undo');
    await wait(1000);
    await page.click('#qa-redo');
    await wait(1000);
    const cellsAfter = [await text(0), await text(1), await text(2)];
    check('표 경계: 행 삭제 뒤 실행 취소·다시 실행 단추가 아무 칸도 바꾸지 않는다', JSON.stringify(cellsAfter) === JSON.stringify(['이름', '', '합계']) && JSON.stringify(await extras()) === tableBefore, cellsAfter.join('|'));
    // 표 안 칸 편집은 실행 취소·다시 실행 단추로 되돌리고 다시 한다(표 덧칠도 그대로)
    await page.dblclick('#grid-edit-table td[data-row="1"][data-col="0"]');
    await page.keyboard.type('새값');
    await page.keyboard.press('Enter');
    await wait(800);
    const typed = await text(1);
    const bgTyped = await editBg(page, 1, 1); // 옆 칸(선택 표시가 없는 칸)의 표 덧칠
    await page.click('#qa-undo');
    await wait(800);
    const undone = await text(1);
    await page.click('#qa-redo');
    await wait(800);
    const redone = await text(1);
    check('표 경계: 표 칸 편집을 실행 취소·다시 실행 단추로', typed === '새값' && undone === '' && redone === '새값' && (await editBg(page, 1, 1)) === bgTyped && bgTyped !== null, `덧칠 ${bgTyped} 입력 "${typed}" 취소 "${undone}" 다시 "${redone}"`);
    await page.screenshot({ path: `${OUT}/7-table-edge-edit.png` });
    // 요약 행(3행) 삭제 → 요약 행이 꺼진다
    await page.click('#grid-edit-table td[data-row="2"][data-col="0"]');
    await page.click('#btn-del-row');
    await dialog();
    await wait(2500);
    const noTotals = await extras();
    check('표 경계: 요약 행을 지우면 요약 행이 꺼진다', noTotals?.r1 === 1 && noTotals?.totalsRows === 0, JSON.stringify(noTotals));
    await page.close();
  }

  // 4) 편집 화면 COUNTIF($A:$A,$A1)>1 — 받지 않은 페이지의 중복, 편집으로 생기고 없어지는 중복(E2)
  for (const rows of [1_500, 20_000]) {
    const { id, tok, token } = await upload(`countif-${rows}.xlsx`, await countifXlsx(rows));
    const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
    await page.goto(BASE + tok.viewerUrl);
    await page.waitForSelector('#grid-table td[data-r]');
    await wait(800);
    const view = await bgOf(page, '#grid-table td[data-c="1"][data-r="1"], #grid-table td[data-c="1"][data-r="2"]'); // 보기 화면 칸 번호는 1부터
    check(`COUNTIF ${rows}행: 보기 화면 A1 중복 칠함, A2 아님`, view[0] === PINK && view[1] !== PINK, view.join(' | '));
    await openEditor(page, tok);
    await wait(3000);
    const a1 = await editBg(page, 0);
    const a2 = await editBg(page, 1);
    check(`COUNTIF ${rows}행: 편집 화면도 A1 중복 칠함(맨 끝 행은 받지 않은 페이지), A2 아님`, a1 === PINK && a2 !== PINK, `${a1} | ${a2}`);
    const { ydoc, provider } = await connect(id, token);
    ydoc.getMap('cells').set('0:1:0', 'K3'); // A2 = K3 → A2·A3 중복
    await wait(2500);
    const made = [await editBg(page, 1), await editBg(page, 2)];
    check(`COUNTIF ${rows}행: 편집으로 중복이 생기면 칠한다`, made.every((c) => c === PINK), made.join(' | '));
    ydoc.getMap('cells').set(`0:${rows - 1}:0`, 'Z'); // 맨 끝 행(받지 않은 페이지)을 바꾼다 → A1은 중복 아님
    await wait(2500);
    const gone = await editBg(page, 0);
    check(`COUNTIF ${rows}행: 받지 않은 행을 고쳐 중복이 없어지면 지운다`, gone !== PINK, String(gone));
    await page.screenshot({ path: `${OUT}/4-countif-${rows}-edit.png` });
    await page.close();
    provider.destroy();
  }

  // 5) x06 차트(축 숨김 포함)
  {
    const { tok } = await upload('x06-charts.xlsx', fs.readFileSync('audit/office/corpus/xlsx/x06-charts.xlsx'));
    const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
    await page.goto(BASE + tok.viewerUrl);
    await page.waitForSelector('#grid-table td[data-r]');
    await wait(1500);
    const viewCharts = await page.evaluate(() => document.querySelectorAll('#grid-image-layer img').length);
    await page.screenshot({ path: `${OUT}/5-x06-view.png` });
    await openEditor(page, tok);
    await wait(3000);
    const editCharts = await page.evaluate(() => document.querySelectorAll('.grid-chart img').length);
    check('x06 차트: 보기·편집 화면에 차트가 그려진다', viewCharts > 0 && editCharts > 0, `보기 ${viewCharts}, 편집 ${editCharts}`);
    await page.screenshot({ path: `${OUT}/5-x06-edit.png` });
    await page.close();
  }

  // 6) 재검토 3 — 409(바쁨)·429(요청 제한)·예산 초과를 받은 범위도 "잠시 뒤 다시" 물어 결국 칠한다(예전엔 시트를 바꿀 때까지 영영 안 칠했다)
  {
    const tokenFor = async (id, viewer) => {
      const t = (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer, canPrint: true, canDownload: true, canCopy: true, canEdit: true }) })).json()).data;
      return { tok: t, token: new URL(BASE + t.editorUrl).searchParams.get('token') };
    };
    /** A열 rows행: [코드]+채움, 코드는 r % half(그래서 각 코드가 A에 두 번), B열은 그 행의 [코드] — 규칙 COUNTIF($A:$A,"*"&$B1&"*")>1 */
    const codeBook = async (rows, fill, uniqueFrom = Infinity) => {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('s');
      const half = rows / 2;
      for (let r = 1; r <= rows; r += 1) {
        const code = r >= uniqueFrom ? `u${r}` : String(r % half).padStart(3, '0');
        ws.getCell(r, 1).value = `[${code}]${'가'.repeat(fill)}`;
        ws.getCell(r, 2).value = `[${r % half === 0 ? '000' : String(r % half).padStart(3, '0')}]`;
      }
      ws.addConditionalFormatting({ ref: `B1:B${rows}`, rules: [{ type: 'expression', priority: 1, formulae: ['COUNTIF($A:$A,"*"&$B1&"*")>1'], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }] });
      return Buffer.from(await wb.xlsx.writeBuffer());
    };
    const statusCounter = (page, counts) => page.on('response', (res) => {
      if (res.url().includes('/api/view/xlsx-cf-ranges')) counts[res.status()] = (counts[res.status()] ?? 0) + 1;
    });
    const allPink = async (page, rows) => (await Promise.all(rows.map((r) => editBg(page, r, 1)))).every((bg) => bg === PINK);
    const waitPink = async (pages, rows, ms) => {
      const t0 = Date.now();
      while (Date.now() - t0 < ms) {
        if ((await Promise.all(pages.map((p) => allPink(p, rows)))).every(Boolean)) return Date.now() - t0;
        await wait(1000);
      }
      return -1;
    };
    const visible = Array.from({ length: 15 }, (_, i) => i);

    // 6a) 예산 초과: 칸 1,000개(2천 자) × 화면 조건 — 열 A를 1분에 8번 고쳐도(그때마다 표가 바뀌어 다시 센다) 결국 모두 칠한다
    {
      const { id, tok, token } = await upload('budget.xlsx', await codeBook(1000, 2000));
      const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
      const counts = {};
      let overBudget = 0;
      page.on('response', async (res) => {
        if (!res.url().includes('/api/view/xlsx-cf-ranges') || res.status() !== 200) return;
        if ((await res.json().catch(() => ({}))).data?.overBudget) overBudget += 1;
      });
      statusCounter(page, counts);
      await openEditor(page, tok);
      const { ydoc, provider } = await connect(id, token);
      const t0 = Date.now();
      for (let i = 0; i < 8; i += 1) {
        ydoc.getMap('cells').set('0:999:0', `[000]${'나'.repeat(2000 + i)}`); // 코드는 그대로 — 결과는 같고 표만 바뀐다
        await wait(7_000);
      }
      // 예산을 넘긴 뒤 아래로 굴려 새 조건(처음 보는 행)을 묻게 한다 — 예전엔 그 범위를 영영 묻지 않아 새 행이 칠해지지 않았다
      await page.evaluate(() => { document.querySelector('.grid-editor-scroll').scrollTop = 4_000; });
      await wait(1_500);
      const shown = await page.evaluate(() => {
        const box = document.querySelector('.grid-editor-scroll').getBoundingClientRect();
        return [...document.querySelectorAll('#grid-edit-table td[data-col="1"][data-row]')]
          .filter((td) => { const r = td.getBoundingClientRect(); return r.top >= box.top + 30 && r.bottom <= box.bottom; })
          .map((td) => Number(td.dataset.row)).filter((r) => r < 499).slice(0, 12);
      });
      const took = await waitPink([page], shown, 150_000);
      check('잠시 뒤 다시(예산 초과): 열을 1분에 8번 고친 뒤 새로 보이는 행도 결국 칠한다', took >= 0 && shown.length >= 5, `행 ${shown[0]}~${shown.at(-1)}, ${took}ms, 예산 초과 답 ${overBudget}회, 응답 ${JSON.stringify(counts)}, 전체 ${Date.now() - t0}ms`);
      provider.destroy();
      await page.close();
    }

    // 6b) 409: 같은 문서 탭 3개(토큰 셋) — 편집 하나에 셋이 한꺼번에 물어 하나는 409를 받아도 결국 셋 다 새 결과로 칠한다
    // 6c) 429: 한 토큰을 세 사람이 같이 쓰고 빠르게 20번 고쳐 분당 40회를 넘겨도 결국 셋 다 칠한다
    for (const [label, shared] of [['409(탭 3개)', false], ['429(한 토큰 3명)', true]]) {
      const rows = 1000;
      // 모두 다른 코드 — 처음엔 칠하지 않음. 탭 3개는 칸을 길게(600자) 해 세 화면의 물음이 서버에서 겹치게 한다(편집마다 표가 바뀌어 다시 센다)
      const { id, tok, token } = await upload(`retry-${shared ? 'shared' : 'tabs'}.xlsx`, await codeBook(rows, shared ? 20 : 600, 1));
      const pages = [];
      const counts = {};
      for (let i = 0; i < 3; i += 1) {
        const page = await browser.newPage({ viewport: { width: 1200, height: 700 } });
        statusCounter(page, counts);
        await openEditor(page, shared ? tok : (await tokenFor(id, `t${i}`)).tok);
        pages.push(page);
      }
      await wait(2_000);
      const { ydoc, provider } = await connect(id, token);
      const edits = shared ? 20 : 3;
      // A의 뒤쪽 행(받지 않은 페이지) 두 곳에 B1~B15의 코드를 넣는다 → 그 칸들이 중복(칠함)이 된다
      const put = (i) => {
        const code = `[${String((i % 15) + 1).padStart(3, '0')}]`;
        ydoc.getMap('cells').set(`0:${900 + (i % 15)}:0`, `${code}a${i}`);
        ydoc.getMap('cells').set(`0:${950 + (i % 15)}:0`, `${code}b${i}`);
      };
      for (let i = 0; i < edits; i += 1) {
        put(i);
        await wait(shared ? 1_000 : 300);
      }
      for (let i = edits; i < 15; i += 1) put(i);
      const took = await waitPink(pages, visible, 150_000); // 토큰의 1분 예산을 다 쓰면 다음 창까지 기다린다
      check(`잠시 뒤 다시 ${label}: 결국 세 화면 모두 새 결과로 칠한다`, took >= 0, `${took}ms, 응답 ${JSON.stringify(counts)}`);
      provider.destroy();
      for (const p of pages) await p.close();
    }
  }
} finally {
  await browser.close();
}
console.log(`스크린샷: ${OUT}`);
const ok = results.every(Boolean);
console.log(ok ? `ALL OK (${results.length})` : `FAILED ${results.filter((r) => !r).length}/${results.length}`);
process.exit(ok ? 0 : 1);
