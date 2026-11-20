// 엑셀 보기·편집 화면 회귀 검사(검토 r2·r3에서 고친 것) — 실제 브라우저로 연다.
//   1) 악의적 차트(공유 맵에 스크립트를 심은 차트)가 실행되지 않고 <img>로 그려진다
//   2) 엑셀 표 덧칠: 정렬(값 섞기) 뒤에도 줄무늬가 행 위치 그대로, 행 삽입 뒤에는 표 범위가 함께 늘어난다
//   3) 무거운 조건부 서식(SUMPRODUCT 40개 × 2만 행): 보기·편집 화면 "일부 생략" 알림, 편집 화면이 멈추지 않는다
//      (첫 칸 1초 안, 300ms 넘는 긴 작업 없음 — 검토 r3 E1)
//   4) 편집 화면 COUNTIF($A:$A,$A1)>1(1,500행·2만 행): 받지 않은 페이지의 중복도 칠하고, 편집으로 중복이 생기고
//      없어지면 따라 바뀐다. 보기 화면과 같은 결과(검토 r3 E2)
//   5) x06 차트(축 숨김 c:delete·tickLblPos none 포함) 보기·편집 화면이 그려진다(스크린샷)
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

async function heavyCfXlsx() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('s');
  for (let r = 1; r <= 20000; r += 1) ws.getCell(r, 1).value = `코드${r % 997}`;
  const b = unzipSync(new Uint8Array(await wb.xlsx.writeBuffer()));
  const rules = Array.from({ length: 40 }, (_, i) => `<conditionalFormatting sqref="A1:A20000"><cfRule type="expression" dxfId="0" priority="${i + 1}"><formula>SUMPRODUCT(($A$1:$A$20000=$A1)*1)&gt;${i + 1}</formula></cfRule></conditionalFormatting>`).join('');
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
} finally {
  await browser.close();
}
console.log(`스크린샷: ${OUT}`);
const ok = results.every(Boolean);
console.log(ok ? `ALL OK (${results.length})` : `FAILED ${results.filter((r) => !r).length}/${results.length}`);
process.exit(ok ? 0 : 1);
