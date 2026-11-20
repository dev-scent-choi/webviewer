// 재연결 회귀 검사(워드·엑셀·PPT) — 웹소켓이 끊기고 서버 방이 새로 만들어진 뒤 편집기가 자동으로 다시 접속할 때, 이전 세대
// 방의 Yjs 상태가 새 방에 합쳐져 슬라이드·문단·행이 두 벌이 되던 문제(2차 검토 RX-C1)를 실제 브라우저로 확인한다.
// 기대: 다시 접속하면 서버가 거절 → 안내 → 새로 불러오기 → 내용이 한 벌, 저장 파일도 한 벌(pptx는 sldId 중복 없음).
//
// 사용법(서버를 먼저 띄워 둘 것 — 업로드·토큰 API가 키 없이 열려 있는 로컬 시험 서버):
//   node scripts/verify-reconnect-e2e.mjs [http://127.0.0.1:8145]
import JSZip from 'jszip';
import fs from 'node:fs';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8145';
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};

async function open(file) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), file.split('/').pop());
  const id = (await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json()).data.id;
  for (let i = 0; ; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data;
    if (st.status === 'ready') break;
    if (st.status === 'failed' || i > 600) throw new Error(`변환 실패 ${file}`);
    await new Promise((r) => setTimeout(r, 400));
  }
  const token = (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: 'reconnect', canPrint: true, canDownload: true, canCopy: true, canEdit: true }) })).json()).data;
  return { id, token };
}

const save = async (t) => fetch(`${BASE}/api/view/save?token=${encodeURIComponent(t.token)}`, { method: 'POST' });
const download = async (t) => JSZip.loadAsync(Buffer.from(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(t.token)}`)).arrayBuffer()));

/** 한 형식: 열기 → 편집 → 저장 → 웹소켓 모두 닫기(유일한 참여자라 방이 닫힘) → 자동 재접속 → 거절 안내·새로 불러오기 → 개수 확인 → 저장 확인 */
async function scenario(browser, label, file, { ready, count, edit, countSaved }) {
  const { id, token } = await open(file);
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.addInitScript(() => {
    const W = window.WebSocket;
    window.__ws = [];
    window.WebSocket = class extends W {
      constructor(...a) {
        super(...a);
        window.__ws.push(this);
      }
    };
  });
  await page.goto(`${BASE}${token.editorUrl}&home=hidden`);
  await page.waitForSelector(ready, { timeout: 120000 });
  await page.waitForTimeout(2500);
  await edit(page);
  await page.waitForTimeout(1500);
  const before = await count(page);
  await save(token);
  const savedBefore = await countSaved(await download(token));
  // 끊기 → 방이 닫히고(유일한 참여자) 자동 재접속 때 새 방이 만들어진다
  const navigated = page.waitForEvent('framenavigated', { timeout: 30000 }).catch(() => null);
  await page.evaluate(() => window.__ws.forEach((w) => w.close()));
  const notice = await page.waitForSelector('#room-stale-notice', { timeout: 30000 }).then(() => true).catch(() => false);
  rec(`${label}: 이전 세대로 다시 접속하면 거절·안내`, notice);
  await navigated; // 안내 뒤 저절로 새로 불러온다
  await page.waitForSelector(ready, { timeout: 120000 });
  await page.waitForTimeout(2500);
  const after = await count(page);
  rec(`${label}: 다시 불러온 뒤 내용이 한 벌`, after === before, `${before} → ${after}`);
  await save(token);
  const savedAfter = await countSaved(await download(token));
  rec(`${label}: 다시 저장한 파일도 한 벌`, JSON.stringify(savedAfter) === JSON.stringify(savedBefore), `${JSON.stringify(savedBefore)} → ${JSON.stringify(savedAfter)}`);
  rec(`${label}: 페이지 오류 없음`, errors.length === 0, errors.slice(0, 2).join(' | '));
  await page.close();
  await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' }).catch(() => {});
}

const browser = await chromium.launch();
try {
  await scenario(browser, 'PPT', 'audit/office/corpus/pptx/p01-placeholders-bullets.pptx', {
    ready: '#slide-canvas .slide-textbox .ProseMirror',
    count: (page) => page.locator('.slide-card').count(),
    edit: async (page) => {
      await page.locator('#slide-canvas .slide-textbox .ProseMirror').first().click();
      await page.keyboard.press('End');
      await page.keyboard.type('재접속');
    },
    countSaved: async (zip) => {
      const pres = await zip.file('ppt/presentation.xml').async('string');
      const ids = [...pres.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]);
      return { sldId: ids.length, unique: new Set(ids).size, parts: Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length };
    },
  });
  await scenario(browser, '워드', 'audit/office/corpus/docx/d01-styles-fonts.docx', {
    ready: '.ProseMirror',
    count: (page) => page.locator('.ProseMirror > *').count(),
    edit: async (page) => {
      await page.locator('.ProseMirror p').first().click();
      await page.keyboard.press('End');
      await page.keyboard.type('재접속');
    },
    countSaved: async (zip) => {
      const xml = await zip.file('word/document.xml').async('string');
      return { paragraphs: (xml.match(/<w:p[\s>]/g) ?? []).length, marker: (xml.match(/재접속/g) ?? []).length };
    },
  });
  await scenario(browser, '엑셀', 'audit/office/corpus/xlsx/x08-korean-quotation.xlsx', {
    ready: '#grid-edit-table td[data-row]',
    // 행 수와 고친 글자 개수(행이 두 벌이 되면 둘 다 늘어난다) — 화면 글자 전체 길이는 계산 값 표시가 늦게 바뀌어 흔들린다
    count: (page) => page.evaluate(() => `${document.querySelectorAll('#grid-edit-table tr').length}행/${(document.querySelector('#grid-edit-table')?.innerText.match(/재접속/g) ?? []).length}`),
    edit: async (page) => {
      await page.locator('#grid-edit-table td[data-row]').nth(40).dblclick();
      await page.keyboard.type('재접속');
      await page.keyboard.press('Enter');
    },
    countSaved: async (zip) => {
      const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
      const shared = (await zip.file('xl/sharedStrings.xml')?.async('string')) ?? '';
      return { rows: (sheet.match(/<row\b/g) ?? []).length, cells: (sheet.match(/<c\b/g) ?? []).length, marker: (shared.match(/재접속/g) ?? []).length + (sheet.match(/재접속/g) ?? []).length };
    },
  });
} finally {
  await browser.close();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${failed === 0 ? 'ALL CHECKS PASSED' : `${failed} FAILED`} (${results.length - failed}/${results.length})`);
process.exit(failed === 0 ? 0 : 1);
