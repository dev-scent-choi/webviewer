// 향기랩 옛 형식 편집(legacyEdit) 확인 — 시험 서버(기본 http://127.0.0.1:8141, 메모리 저장 권장).
//   node scripts/verify-legacy-edit-e2e.mjs [BASE] [스크린샷 폴더]
// doc·xls·ppt를 legacyEdit=1로 올리면 편집기 주소, 안 고친 내려받기는 같은 바이트, 워드는 화면에서 고친 뒤 .doc로 받아 다시 올리면
// 고친 글자가 남는다. legacyEdit 없이 올린 doc은 지금처럼 편집기가 없다(다른 호출자 회귀). 시험 자료: tests/fixtures/legacy(Apache POI).
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8141';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'legacy-edit-'));
fs.mkdirSync(shotDir, { recursive: true });
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};
const fx = (name) => fs.readFileSync(path.resolve('tests', 'fixtures', 'legacy', name));
const hash = (b) => crypto.createHash('sha256').update(b).digest('hex');

async function upload(name, legacyEdit = true, bytes = fx(name)) {
  const q = `filename=${encodeURIComponent(name)}${legacyEdit ? '&legacyEdit=1' : ''}`;
  const r = await fetch(`${BASE}/api/documents/stream?${q}`, { method: 'POST', headers: { 'content-type': 'application/octet-stream' }, body: bytes });
  return { status: r.status, json: await r.json() };
}
/** 변환 중(processing)이면 토큰이 404 — 열릴 때까지 다시 묻는다 */
async function token(id) {
  for (let i = 0; i < 240; i++) {
    const r = await fetch(`${BASE}/api/documents/${id}/tokens`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ viewer: '옛형식시험', canEdit: true, canDownload: true, canPrint: true, canCopy: true }),
    });
    if (r.status === 200) return (await r.json()).data;
    await new Promise((res) => setTimeout(res, 500));
  }
  throw new Error(`토큰을 받지 못함: ${id}`);
}
async function download(tok) {
  const r = await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(tok)}`);
  const name = decodeURIComponent((r.headers.get('content-disposition') ?? '').split("''")[1] ?? '');
  return { status: r.status, name, bytes: Buffer.from(await r.arrayBuffer()) };
}

// 1) 안 고친 내려받기 = 같은 바이트·원래 이름, 편집기 주소, 매크로 표시
for (const [name, editor, macros] of [
  ['table-merges.doc', '/editor.html', false],
  ['chinese-provinces.xls', '/editor-grid.html', false],
  ['SampleShow.ppt', '/editor-slides.html', false],
  ['SimpleMacro.xls', '/editor-grid.html', true],
]) {
  const up = await upload(name);
  check(up.status === 201 && up.json.data.legacy?.ext === name.split('.').pop(), `${name}: 올림(legacy 표시)`, `status ${up.status}`);
  check(up.json.data.legacy?.hasMacros === macros, `${name}: 매크로 표시 ${macros}`);
  const t = await token(up.json.data.id);
  check(t.editorUrl?.startsWith(editor), `${name}: 편집기 주소 ${editor}`, t.editorUrl);
  const d = await download(t.token);
  check(d.status === 200 && hash(d.bytes) === hash(fx(name)), `${name}: 안 고친 내려받기는 같은 바이트`);
  check(d.name === name, `${name}: 내려받기 이름`, d.name);
}

// 2) 다른 호출자 회귀: legacyEdit 없는 doc은 편집기 없음
{
  const up = await upload('table-merges.doc', false);
  const t = await token(up.json.data.id);
  check(up.json.data.ext === 'doc' && !up.json.data.legacy && !t.editorUrl, 'legacyEdit 없는 doc: 지금처럼 열람(편집기 주소 없음)');
}

// 3) 워드: 화면에서 고친 뒤 .doc로 받고, 그 .doc를 다시 올려 편집기에서 고친 글자를 본다 + 매크로 경고 줄
const browser = await chromium.launch();
try {
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  const up = await upload('SimpleMacro.doc');
  const t = await token(up.json.data.id);
  await page.goto(`${BASE}${t.editorUrl}&home=hidden&back=parent`);
  await page.waitForSelector('.ProseMirror', { timeout: 60000 });
  check(await page.locator('.legacy-macro-note').isVisible(), '매크로 경고 줄이 보인다');
  await page.click('.ProseMirror');
  await page.keyboard.press('Control+End');
  await page.keyboard.type(' 옛형식편집LEGACY');
  await page.waitForTimeout(1500);
  await fetch(`${BASE}/api/view/save?token=${encodeURIComponent(t.token)}`, { method: 'POST' });
  await page.screenshot({ path: path.join(shotDir, 'doc-edited.png') });
  const d = await download(t.token);
  check(d.bytes.readUInt32LE(0) === 0xe011cfd0 && d.name === 'SimpleMacro.doc', '고친 뒤 내려받기: .doc(OLE)', d.name);
  check(hash(d.bytes) !== hash(fx('SimpleMacro.doc')), '고친 뒤 내려받기: 올린 바이트와 다름');
  const again = await upload('SimpleMacro.doc', true, d.bytes);
  const t2 = await token(again.json.data.id);
  await page.goto(`${BASE}${t2.editorUrl}&home=hidden&back=parent`);
  await page.waitForSelector('.ProseMirror', { timeout: 60000 });
  check((await page.locator('.ProseMirror').innerText()).includes('옛형식편집LEGACY'), '다시 올린 .doc에 고친 글자가 남았다');
  check(errors.length === 0, '편집기 페이지 오류 없음', errors.join(' | '));
} finally {
  await browser.close();
}

console.log(`\n${results.filter(Boolean).length}/${results.length} 통과 — 스크린샷 ${shotDir}`);
process.exit(results.every(Boolean) ? 0 : 1);
