// 엑셀 표시 충실도 회귀 검사 — 보기(viewer.html)·편집(editor-grid.html) 두 화면을 실제 브라우저로 열어
// 열 너비·행 높이(자동 맞춤 포함)·병합·숨긴 행/열·줄 바꿈·세로 맞춤·테두리·격자선 숨김·그림·둘째 시트를 잰다.
// 한 번 고친 것이 다른 경로(예: 보기 → 편집 기본값 변경)에서 다시 깨지지 않게, 운영에 반영하기 전에 돌린다.
//
// 사용법(서버를 먼저 띄워 둘 것 — 업로드·토큰 API가 키 없이 열려 있는 로컬 서버):
//   node scripts/verify-xlsx-fidelity-e2e.mjs [http://127.0.0.1:8124] [스크린샷 폴더]
// 시험 파일은 매번 ExcelJS로 만든다(고객 문서를 쓰지 않는다).
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-fidelity-'));
fs.mkdirSync(shotDir, { recursive: true });

const results = [];
const rec = (name, ok, detail = '') => {
  results.push({ name, ok, detail });
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const near = (a, b, tol = 2) => typeof a === 'number' && Math.abs(a - b) <= tol;

// 1x1 빨간 PNG
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');

async function buildFixture() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('점검표', { views: [{ showGridLines: false }] });
  const thin = { style: 'thin', color: { argb: 'FF000000' } };
  const box = { top: thin, left: thin, bottom: thin, right: thin };
  [5, 22, 9, 40, 12, 15].forEach((w, i) => (ws.getColumn(i + 1).width = w));
  ws.mergeCells('A1:F1');
  ws.getCell('A1').value = '설비 정기 점검표';
  ws.getCell('A1').font = { size: 16, bold: true };
  ws.getCell('A1').alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(1).height = 36;
  ['No', '설비명', '구분', '점검 내용', '결과', '비고'].forEach((h, i) => {
    const c = ws.getCell(3, i + 1);
    c.value = h;
    c.border = box;
    c.alignment = { horizontal: 'center' };
  });
  const rows = [
    [1, '공조기 AHU-1', '기계', '필터 차압 확인 및 벨트 장력 점검, 이상 소음 여부 확인', '양호', ''],
    [2, null, '전기', '모터 절연저항 측정', '양호', '1.2 GΩ'],
    [3, null, '제어', 'DDC 설정값 확인', '불량', '설정 오류 수정 필요'],
  ];
  rows.forEach((row, r) =>
    row.forEach((v, c) => {
      const cell = ws.getCell(4 + r, c + 1);
      cell.value = v;
      cell.border = box;
      cell.alignment = { wrapText: true, vertical: 'middle' };
    }),
  );
  ws.mergeCells('B4:B6');
  ws.getRow(8).hidden = true;
  ws.getCell('A8').value = '숨김';
  ws.getColumn(7).hidden = true;
  ws.getCell('G4').value = '숨긴 열';
  ws.mergeCells('C10:F12');
  ws.getCell('C10').value = '종합 의견은 위에 맞춘다';
  ws.getCell('C10').alignment = { vertical: 'top', wrapText: true };
  ws.getCell('H4').value = '끝';
  const img = wb.addImage({ buffer: PNG, extension: 'png' });
  ws.addImage(img, { tl: { col: 7, row: 0 }, ext: { width: 120, height: 40 } });
  // 바로가기(문서 안 링크)·목록 상자 — 편집 화면에서 눌러 이동·골라 넣기
  ws.getCell('H6').value = { text: '둘째로', hyperlink: "#'둘째'!B4" };
  ws.getCell('I4').dataValidation = { type: 'list', allowBlank: true, formulae: ['"가,나,다"'] };
  const ws2 = wb.addWorksheet('둘째');
  ws2.properties.defaultColWidth = 18;
  ws2.getCell('A1').value = '둘째 시트';
  // 표(표 서식) — 필터 단추는 시트가 아니라 표 정의에 있다
  ws2.addTable({ name: '표1', ref: 'A3', headerRow: true, columns: [{ name: '이름', filterButton: true }, { name: '값', filterButton: true }], rows: [['a', 1], ['b', 2]] });
  // 여러 페이지(500행 단위)에 걸친 필터 — 편집기가 화면 근처만 받아 오므로, 뒤쪽 행에만 있는 값이 목록에서 빠지던 문제
  const ws3 = wb.addWorksheet('긴목록');
  ws3.addRow(['번호', '상태']);
  for (let i = 1; i <= 1200; i += 1) ws3.addRow([i, i > 1150 ? '끝쪽값' : ['양호', '불량'][i % 2]]);
  ws3.autoFilter = 'A1:B1201';
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function upload(bytes) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), 'fidelity.xlsx');
  const id = (await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json()).data.id;
  for (;;) {
    const st = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data;
    if (st.status === 'ready') break;
    if (st.status === 'failed') throw new Error('변환 실패');
    await new Promise((r) => setTimeout(r, 400));
  }
  const token = (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: 'fidelity', canPrint: true, canDownload: true, canCopy: true, canEdit: true }) })).json()).data;
  return { id, token };
}

// Calibri 11(기본 글꼴) 기준 열 너비(px) — OOXML 변환식: 5→35, 22→154, 9→63, 40→280, 12→84, 15→105
const EXPECT_COLS = [35, 154, 63, 280, 84, 105];

const { id, token } = await upload(await buildFixture());
const browser = await chromium.launch();
const errors = [];
for (const mode of ['viewer', 'editor']) {
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  page.on('pageerror', (e) => errors.push(`${mode}: ${e}`));
  await page.goto(BASE + (mode === 'viewer' ? token.viewerUrl : token.editorUrl));
  const sel = mode === 'viewer' ? (r, c) => `#grid-table td[data-r="${r}"][data-c="${c}"]` : (r, c) => `#grid-edit-table td[data-row="${r - 1}"][data-col="${c - 1}"]`;
  await page.waitForSelector(mode === 'viewer' ? '#grid-table td[data-r]' : '#grid-edit-table td[data-row]', { timeout: 30000 });
  await page.waitForTimeout(2000);
  const box = async (r, c) => {
    const l = page.locator(sel(r, c));
    return (await l.count()) ? l.boundingBox() : null;
  };
  const cols = [];
  for (let c = 1; c <= 6; c += 1) cols.push(Math.round((await box(14, c))?.width ?? -1));
  rec(`[${mode}] 열 너비(엑셀 변환식)`, cols.every((w, i) => near(w, EXPECT_COLS[i])), JSON.stringify(cols));
  rec(`[${mode}] 숨긴 열(G)은 폭 0`, ((await box(14, 7))?.width ?? 0) <= 1);
  const rowH = await page.evaluate((m) => {
    const out = {};
    const trs = m === 'viewer' ? document.querySelectorAll('#grid-table tr') : document.querySelectorAll('#grid-edit-table tbody tr:not([aria-hidden])');
    for (const tr of trs) {
      const n = m === 'viewer' ? Number(tr.querySelector('td[data-r]')?.dataset.r) : Number(tr.querySelector('.g-edit-rownum')?.dataset.row) + 1;
      if (n) out[n] = Math.round(tr.getBoundingClientRect().height);
    }
    return out;
  }, mode);
  rec(`[${mode}] 행 높이: 1행 36pt=48px, 4행 줄 바꿈 자동 맞춤(>20), 기본 20`, near(rowH[1], 48) && rowH[4] > 30 && near(rowH[5], 20), JSON.stringify({ 1: rowH[1], 4: rowH[4], 5: rowH[5] }));
  rec(`[${mode}] 숨긴 8행은 높이 0`, (rowH[8] ?? 0) === 0, String(rowH[8]));
  const merged = await box(1, 1);
  rec(`[${mode}] 병합 A1:F1 폭 = A~F 합`, near(merged?.width, EXPECT_COLS.reduce((a, b) => a + b, 0), 4), String(Math.round(merged?.width ?? 0)));
  const vmerge = await box(4, 2);
  rec(`[${mode}] 세로 병합 B4:B6 높이 = 4~6행 합`, near(vmerge?.height, rowH[4] + rowH[5] + rowH[6], 4), `${Math.round(vmerge?.height ?? 0)} vs ${rowH[4] + rowH[5] + rowH[6]}`);
  const info = await page.evaluate(({ m, wrapSel, topSel, borderSel }) => {
    const wrap = document.querySelector(wrapSel);
    const el = wrap?.querySelector('.g-cell-spill') ?? wrap;
    const table = document.querySelector(m === 'viewer' ? '#grid-table' : '#grid-edit-table');
    return {
      wrapLines: el ? Math.round(el.getBoundingClientRect().height / 18) : 0,
      vtop: document.querySelector(topSel) ? getComputedStyle(document.querySelector(topSel)).verticalAlign : null,
      border: document.querySelector(borderSel) ? getComputedStyle(document.querySelector(borderSel)).borderBottomStyle : null,
      noGrid: table?.classList.contains('g-no-gridlines'),
      images: document.querySelectorAll(m === 'viewer' ? '#grid-image-layer img' : '.grid-chart-layer img.grid-image').length,
    };
  }, { m: mode, wrapSel: sel(4, 4), topSel: sel(10, 3), borderSel: sel(5, 3) });
  rec(`[${mode}] 줄 바꿈 칸이 여러 줄로`, info.wrapLines >= 2, String(info.wrapLines));
  rec(`[${mode}] 세로 맞춤 위`, info.vtop === 'top', info.vtop);
  rec(`[${mode}] 칸 테두리(실선)`, info.border === 'solid', info.border);
  rec(`[${mode}] 격자선 숨김 시트`, info.noGrid === true);
  rec(`[${mode}] 그림 표시`, info.images === 1, String(info.images));
  // 맞닿은 두 칸의 테두리가 겹쳐 두 배로 굵어지지 않는다 — 4행/5행(둘 다 테두리) 경계선과 C/D열 경계선을 화면 픽셀로 잰다
  const darkPixels = async (x, y, w, h) => {
    const png = await page.screenshot({ clip: { x: Math.round(x), y: Math.round(y), width: w, height: h } });
    return page.evaluate(async (b64) => {
      const img = new Image();
      img.src = `data:image/png;base64,${b64}`;
      await img.decode();
      const canvas = document.createElement('canvas');
      canvas.width = img.width;
      canvas.height = img.height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);
      const data = ctx.getImageData(0, 0, img.width, img.height).data;
      let n = 0;
      for (let i = 0; i < data.length; i += 4) if ((data[i] + data[i + 1] + data[i + 2]) / 3 < 120) n += 1;
      return n;
    }, png.toString('base64'));
  };
  const c5 = await box(5, 3);
  const hLine = await darkPixels(c5.x + 4, c5.y - 6, 1, 12);
  const vLine = await darkPixels(c5.x + c5.width - 6, c5.y + c5.height / 2, 12, 1);
  rec(`[${mode}] 맞닿은 테두리 1px(두 배로 굵지 않음)`, hLine === 1 && vLine === 1, `가로 ${hLine}px, 세로 ${vLine}px`);
  if (mode === 'editor') {
    await page.locator(sel(4, 9)).click();
    await page.waitForTimeout(300);
    rec('[editor] 목록 상자 칸을 고르면 ▼ 단추', await page.locator('.grid-list-btn').isVisible());
    await page.locator('.grid-list-btn').click();
    const items = await page.locator('.grid-list-menu button').allInnerTexts();
    rec('[editor] 목록 항목', items.join(',') === '가,나,다', items.join(','));
    await page.keyboard.press('Escape');
  }
  await page.screenshot({ path: path.join(shotDir, `${mode}.png`) });
  const tabSel = mode === 'viewer' ? '.grid-sheet-tab' : '#sheet-tabs .grid-sheet-tab';
  rec(`[${mode}] 둘째 시트 탭이 있다`, (await page.locator(tabSel).count()) >= 2);
  if (mode === 'editor') {
    await page.locator(sel(6, 8)).click();
    await page.waitForTimeout(800);
    const where = await page.evaluate(() => [document.querySelector('#sheet-tabs .grid-sheet-tab.active')?.textContent, document.activeElement?.dataset?.row, document.activeElement?.dataset?.col].join(':'));
    rec('[editor] 바로가기 클릭 → 둘째 시트 B4로 이동', where === '둘째:3:1', where);
    rec('[editor] 표(표 서식)의 필터 단추', (await page.locator('.grid-filter-btn').count()) === 2, String(await page.locator('.grid-filter-btn').count()));
    // 긴 목록 시트: 로딩 표시가 사라지고, 필터 목록에 뒤쪽 행(1151~1200)의 값까지 나온다
    await page.locator('#sheet-tabs .grid-sheet-tab', { hasText: '긴목록' }).click();
    await page.waitForFunction(() => document.querySelector('.grid-loading')?.hidden === true, null, { timeout: 30000 }).catch(() => {});
    rec('[editor] 데이터가 오면 로딩 표시가 사라진다', await page.locator('.grid-loading').isHidden());
    await page.locator('.grid-filter-btn').nth(1).click();
    await page.waitForFunction(() => document.querySelectorAll('.filter-menu .filter-values label').length > 0, null, { timeout: 30000 }).catch(() => {});
    const longValues = await page.locator('.filter-menu .filter-values label').allInnerTexts();
    rec('[editor] 여러 페이지에 걸친 필터 목록에 뒤쪽 값까지(빈 값 없이)', longValues.includes('끝쪽값') && !longValues.some((v) => v.includes('값 없음')), longValues.join(','));
    await page.keyboard.press('Escape');
  }
  await page.close();
}
rec('페이지 오류 없음', errors.length === 0, errors.slice(0, 2).join(' | '));
await browser.close();
await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' }).catch(() => {});
const fails = results.filter((r) => !r.ok);
console.log(`\n${fails.length === 0 ? 'ALL CHECKS PASSED' : `${fails.length} FAILED`} (${results.length - fails.length}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(fails.length === 0 ? 0 : 1);
