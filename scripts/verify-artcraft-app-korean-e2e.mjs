// 앱 화면 6 (8) 격자 앱 한글 칸 — ???·네모가 아니라 한글 글리프로 그려지는지, 글꼴이 서빙되는 wasm에 실제로 들어 있는지.
// 시험 서버를 직접 띄운다(포트 8177 고정, 끝나면 PID로 끔):  node scripts/verify-artcraft-app-korean-e2e.mjs
// 캡처는 .audit-run/app-screen/korean-*.png. 한글 시스템 글꼴이 있는 PC는 글꼴 대체로도 한글이 나오므로 wasm 안 글꼴 바이트 단언을 따로 둔다.
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import { chromium } from 'playwright';
import sharp from 'sharp';

const PORT = 8177;
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', 'app-screen', 'korean-run');
const SHOTS = path.resolve('.audit-run', 'app-screen');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};

fs.rmSync(RUN, { recursive: true, force: true });
fs.mkdirSync(path.join(RUN, 'tmp'), { recursive: true });
const log = fs.openSync(path.join(RUN, 'server.log'), 'a');
const server = spawn(process.execPath, ['node_modules/tsx/dist/cli.mjs', 'src/server.ts'], {
  env: { ...process.env, DOCVIEW_PORT: String(PORT), DOCVIEW_HOST: '127.0.0.1', DOCVIEW_SECRET: randomBytes(24).toString('hex'), DOCVIEW_DATA_DIR: path.join(RUN, 'data'), TMP: path.join(RUN, 'tmp'), TEMP: path.join(RUN, 'tmp'), DOCVIEW_LO_PROFILE_ROOT: path.join(RUN, 'lo') },
  stdio: ['ignore', log, log],
});
const stop = () => { try { execFileSync('taskkill', ['/PID', String(server.pid), '/T', '/F'], { stdio: 'ignore' }); } catch {} };
process.on('exit', stop);

let browser = null;
try {
  for (let i = 0; i < 120 && !(await fetch(`${BASE}/api/formats`).then((r) => r.ok, () => false)); i += 1) await sleep(500);

  // (8a) 서빙되는 grid wasm에 한글 글꼴 조각이 들어 있다(build.ts 검사와 같은 조각: SCDream4.otf 4096~4160바이트)
  const manifest = JSON.parse(fs.readFileSync('public/artcraft-app/manifest.json', 'utf8'));
  const wasmUrl = manifest.grid.replace(/gridcraft-web\.js$/, 'gridcraft-web_bg.wasm');
  const wasm = Buffer.from(await (await fetch(`${BASE}${wasmUrl}`)).arrayBuffer());
  const sig = fs.readFileSync('public/fonts/s-core-dream/SCDream4.otf').subarray(4096, 4160);
  rec('(8) grid wasm에 한글 글꼴(S-Core Dream 4) 바이트가 있음', wasm.indexOf(sig) >= 0, `${(wasm.length / 1048576).toFixed(2)} MB`);

  const upload = async (file) => {
    const form = new FormData();
    form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
    const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
    for (let i = 0; i < 1200; i += 1) {
      const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
      if (st.status === 'ready') return st.id;
      if (st.status === 'failed') throw new Error(st.error);
      await sleep(250);
    }
    throw new Error('업로드 대기 초과');
  };
  const make = async (dir, text) => {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('시험');
    ws.getColumn(1).width = 60;
    ws.getRow(1).height = 60;
    ws.getCell('A1').value = text;
    ws.getCell('A1').font = { size: 36 };
    fs.mkdirSync(path.join(RUN, dir), { recursive: true });
    const file = path.join(RUN, dir, 'glyph.xlsx');
    fs.writeFileSync(file, Buffer.from(await wb.xlsx.writeBuffer()));
    return upload(file);
  };
  browser = await chromium.launch();
  const shot = async (id, name) => {
    const t = (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ viewer: 'korean', canPrint: true, canDownload: true, canCopy: true, canEdit: false, engine: 'artcraft' }) })).json()).data;
    const context = await browser.newContext({ viewport: { width: 1280, height: 800 } });
    const page = await context.newPage();
    await page.goto(`${BASE}${t.viewerUrl}`);
    await page.waitForFunction(() => document.body.dataset.appState === 'open', null, { timeout: 45_000 });
    await sleep(2500);
    const png = await page.locator('#app-frame').screenshot();
    fs.writeFileSync(path.join(SHOTS, `korean-${name}.png`), png);
    await context.close();
    return (await sharp(png).grayscale().raw().toBuffer({ resolveWithObject: true })).data;
  };
  const diff = (a, b) => {
    let n = 0;
    for (let i = 0; i < Math.min(a.length, b.length); i += 1) if (Math.abs(a[i] - b[i]) > 64) n += 1;
    return n;
  };
  const hangul = await shot(await make('g-hangul', '가나다라마바사아자차'), 'hangul');
  const question = await shot(await make('g-question', '??????????'), 'question');
  const tofu = await shot(await make('g-tofu', '□□□□□□□□□□'), 'tofu');
  const dq = diff(hangul, question);
  const dt = diff(hangul, tofu);
  rec('(8) xlsx 한글 칸: 물음표로 그린 그림과 다름', dq > 400, `${dq}px`);
  rec('(8) xlsx 한글 칸: 네모로 그린 그림과 다름', dt > 400, `${dt}px`);
  rec('(8) 비교가 유효함(물음표·네모 그림끼리도 다름)', diff(question, tofu) > 400, `${diff(question, tofu)}px`);
} finally {
  await browser?.close();
  stop();
}
const bad = results.filter((x) => !x).length;
console.log(bad === 0 ? `통과 ${results.length}항목` : `실패 ${bad}/${results.length}`);
process.exit(bad === 0 ? 0 : 1);
