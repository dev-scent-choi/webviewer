// '열람만' 토큰(편집·내려받기 둘 다 없음) 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-view-only-e2e.mjs [BASE] [스크린샷 폴더]
// 열람만 토큰은 편집기 주소를 받지 않고 서버가 그린 열람 화면으로 열린다. 편집 연결(WebSocket)·원본·편집기 전용 데이터는 거부,
// 엑셀 표의 숨긴 행·열 값은 보내지 않는다. 같은 문서를 내려받기 권한이 있는 읽기 전용 토큰으로 열면 예전처럼 편집기.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { strToU8, zipSync } from 'fflate';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'view-only-'));
fs.mkdirSync(shotDir, { recursive: true });
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

async function makeXlsx() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('보이는 시트');
  ws.getCell('A1').value = '품목';
  ws.getCell('B1').value = '단가';
  ws.getCell('C1').value = '원가(숨김열)';
  ws.getCell('A2').value = '사과';
  ws.getCell('B2').value = 1200;
  ws.getCell('C2').value = '비밀원가900';
  ws.getCell('A3').value = '숨긴행비밀';
  ws.getCell('B3').value = { formula: 'B2*2', result: 2400 };
  ws.getCell('A4').value = '합계';
  ws.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 3600 };
  ws.getColumn(3).hidden = true;
  ws.getRow(3).hidden = true;
  const hiddenSheet = wb.addWorksheet('숨긴 시트');
  hiddenSheet.state = 'hidden';
  hiddenSheet.getCell('A1').value = '숨긴시트비밀';
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

function makeDocx() {
  const files = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    'word/document.xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>열람 전용 시험 문서</w:t></w:r></w:p><w:p><w:r><w:t>두 번째 문단</w:t></w:r></w:p></w:body></w:document>',
    ),
  };
  return zipSync(files);
}

async function upload(bytes, name) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패: ${JSON.stringify(r)}`);
  for (let i = 0; i < 400; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready') return r.data.id;
    if (st.status === 'failed' || st.status === 'error') throw new Error(`변환 실패: ${JSON.stringify(st)}`);
    await new Promise((res) => setTimeout(res, 250));
  }
  throw new Error('대기 시간 초과');
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '열람자', canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);
const status = async (url) => (await fetch(`${BASE}${url}`)).status;
function wsClose(docId, tok) {
  return new Promise((resolve) => {
    const ws = new WebSocket(`${BASE.replace(/^http/, 'ws')}/ws/edit/${docId}?token=${encodeURIComponent(tok)}`);
    let opened = false;
    ws.binaryType = 'arraybuffer';
    ws.onopen = () => (opened = true);
    // 방에 들어가면 서버가 문서 상태를 보낸다(첫 메시지) — 그걸 받으면 들어간 것
    ws.onmessage = () => {
      resolve({ joined: true, code: 0 });
      ws.close();
    };
    ws.onclose = (e) => resolve({ joined: false, code: e.code, opened });
    setTimeout(() => resolve({ joined: false, code: -1, opened }), 15000);
  });
}

const xlsxId = await upload(await makeXlsx(), '단가표.xlsx');
const docxId = await upload(makeDocx(), '열람시험.docx');

// ---------- 토큰 발급: 열람만은 편집기 주소 없음 ----------
const voX = await token(xlsxId, {});
const voD = await token(docxId, {});
check(!voX.editorUrl && !voD.editorUrl && voX.viewerUrl && voD.viewerUrl, '열람만 토큰: 편집기 주소 없이 열람 화면 주소만', `${voX.editorUrl ?? '-'} / ${voD.editorUrl ?? '-'}`);
const roX = await token(xlsxId, { canDownload: true });
check(!!roX.editorUrl, '읽기 전용+내려받기 토큰은 예전처럼 편집기 주소', roX.editorUrl?.split('?')[0]);

// ---------- 서버가 막는 것 ----------
const q = (t) => `token=${encodeURIComponent(t)}`;
const blocked = {};
for (const route of ['hwp-source', 'slide-media', 'xlsx-charts', 'xlsx-extras', 'changes', 'download']) {
  blocked[route] = await status(`/api/view/${route}?${q(voX.token)}${route === 'slide-media' ? '&path=ppt/media/image1.png' : ''}`);
}
check(Object.values(blocked).every((s) => s === 403), '열람만: 원본·편집기 전용 데이터 거부', JSON.stringify(blocked));
check((await status(`/api/view/xlsx-charts?${q(roX.token)}`)) === 200, '읽기 전용+내려받기: 편집기 데이터는 그대로');
const wsVo = await wsClose(xlsxId, voX.token);
check(!wsVo.joined && wsVo.code === 4403, '열람만: 편집 연결(WebSocket) 거부', JSON.stringify(wsVo));
const fileBefore = Buffer.from(await (await fetch(`${BASE}/api/view/download?${q(roX.token)}`)).arrayBuffer());
const wsRo = await wsClose(xlsxId, roX.token);
check(wsRo.joined, '읽기 전용+내려받기: 편집 연결은 들어감(쓰기는 서버가 막음)', JSON.stringify(wsRo));
// 마지막 참여자가 나간 뒤 저장 여부가 정해진다 — 편집이 없었으면 파일이 그대로여야 한다(숨긴 시트·행·열 포함)
await new Promise((r) => setTimeout(r, 4000));
const fileAfter = Buffer.from(await (await fetch(`${BASE}/api/view/download?${q(roX.token)}`)).arrayBuffer());
check(fileBefore.equals(fileAfter), '읽기 전용으로 열었다 닫아도 파일이 바뀌지 않음', `${fileBefore.length} → ${fileAfter.length}`);

// ---------- 엑셀 표: 숨긴 행·열·시트, 수식 ----------
const meta = (await (await fetch(`${BASE}/api/view/grid/meta?${q(voX.token)}`)).json()).data;
const rows = (await (await fetch(`${BASE}/api/view/grid/rows?${q(voX.token)}&sheet=0&from=1&count=10`)).json()).data.rows;
const flat = JSON.stringify(rows);
check(meta.sheets.length === 1 && !JSON.stringify(meta).includes('숨긴 시트'), '숨긴 시트는 목록에 없음', meta.sheets.map((s) => s.name).join(','));
check(flat.includes('사과') && (flat.includes('3,600') || flat.includes('3600')), '보이는 칸 값은 그대로', flat.slice(0, 120));
check(!flat.includes('비밀원가900') && !flat.includes('숨긴행비밀'), '숨긴 행·열 값은 보내지 않음');
check(!flat.includes('SUM(') && !flat.includes('B2*2'), '수식은 보내지 않고 계산값만');
const sHidden = (await (await fetch(`${BASE}/api/view/search?${q(voX.token)}&q=${encodeURIComponent('비밀')}`)).json()).data;
check(sHidden.total === 0, '찾기로도 숨긴 값이 나오지 않음', JSON.stringify(sHidden));
const roRows = JSON.stringify((await (await fetch(`${BASE}/api/view/grid/rows?${q(roX.token)}&sheet=0&from=1&count=10`)).json()).data.rows);
check(roRows.includes('비밀원가900'), '내려받기 권한이 있으면 예전처럼(숨긴 칸 값 포함)');

// ---------- 브라우저: 열람 화면으로 열림 ----------
const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
const errors = [];
const page = await ctx.newPage();
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => {
  if (m.type() === 'error') errors.push(m.text());
});
await page.goto(`${BASE}${voX.viewerUrl}`);
await page.waitForFunction(() => (document.getElementById('grid-table')?.textContent ?? '').includes('사과'), null, { timeout: 30000 });
const gridText = await page.textContent('#grid-table');
await page.screenshot({ path: path.join(shotDir, 'view-only-xlsx.png') });
check(gridText.includes('품목') && gridText.includes('사과') && !gridText.includes('비밀'), '엑셀: 열람 화면에 보이는 값만', gridText.replace(/\s+/g, ' ').slice(0, 80));
check((await page.$$('.grid-sheet-tab')).length <= 1, '엑셀: 숨긴 시트 탭 없음');

await page.goto(`${BASE}${voD.viewerUrl}`);
await page.waitForSelector('.page-frame.page-ready', { timeout: 60000 });
await page.screenshot({ path: path.join(shotDir, 'view-only-docx.png') });
const dlHidden = await page.evaluate(() => {
  const b = document.getElementById('download-btn');
  return !b || b.hidden || b.disabled || b.offsetParent === null;
});
check(true, '워드: 열람 화면(서버가 그린 쪽 그림)으로 열림');
check(dlHidden, '워드: 내려받기 단추 없음(꺼짐)');
check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 3).join(' | '));
await browser.close();

for (const id of [xlsxId, docxId]) await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' });
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
