// npm run audit:office — 오피스 표시 품질 점수 한 번에 재기.
// 1) 정답 PDF(MS Office)가 없으면 ref_office.ps1로 만든다(이 PC에 Office가 있을 때만)
// 2) 이 작업 폴더의 서버를 시험용 포트·전용 데이터/임시/LibreOffice 프로필로 따로 띄운다(운영 8000/8010과 겹치지 않게)
// 3) capture.mjs로 보기·편집 화면을 찍고 4) compare.py로 나란히 그림과 점수(scores.md)를 낸다 5) 띄운 서버만 끈다
// 사용: npm run audit:office [-- --port 8141 --only docx --no-office]
//   --no-office: 정답 PDF가 없어도 MS Office를 절대 띄우지 않는다(없으면 멈춘다) — 다른 작업이 쓰는 PC에서 Office 창을 건드리지 않게
import { spawn, spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

const arg = (name, def) => { const i = process.argv.indexOf(`--${name}`); return i > 0 ? process.argv[i + 1] : def; };
const PORT = Number(arg('port', '8141'));
const ONLY = arg('only', '');
const NO_OFFICE = process.argv.includes('--no-office');
if ([8000, 8010, 8020, 8100, 8110, 8130].includes(PORT)) throw new Error(`운영 포트 ${PORT}는 쓰지 않는다`);
const root = process.cwd();
const OUT = path.join(root, 'docs/quality-audit/out');
const CORPUS = path.join(root, 'audit/office/corpus');
const run = path.join(root, '.audit-run');
for (const d of ['data', 'tmp', 'lo']) fs.mkdirSync(path.join(run, d), { recursive: true });

const refDir = path.join(OUT, 'ref');
const want = (ONLY ? [ONLY] : ['docx', 'xlsx', 'pptx']).flatMap((k) => fs.readdirSync(path.join(CORPUS, k)).filter((f) => !f.startsWith('~$')).map((f) => f.replace(/\.[^.]+$/, '')));
const missing = want.filter((n) => !fs.existsSync(path.join(refDir, `${n}.pdf`)));
if (missing.length && NO_OFFICE) throw new Error(`정답 PDF가 없다(--no-office라 Office를 띄우지 않음): ${missing.join(', ')}`);
if (missing.length && process.platform === 'win32') {
  // 정답은 Office로 다시 저장하기 전 원본(raw)에서 만든다 — git엔 raw가 없으니 먼저 다시 만든다(a1~a3은 ArtCraft 시험 파일이라 corpus 것을 쓴다)
  if (!fs.existsSync(path.join(CORPUS, 'raw'))) {
    spawnSync('python', ['audit/office/gen_corpus.py', 'audit/office/corpus'], { stdio: 'inherit' });
    for (const k of ['docx', 'xlsx', 'pptx']) for (const f of fs.readdirSync(path.join(CORPUS, k))) if (/^a\d/.test(f)) fs.copyFileSync(path.join(CORPUS, k, f), path.join(CORPUS, 'raw', k, f));
  }
  console.log(`정답 PDF ${missing.length}개 만들기(MS Office)…`);
  spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', 'audit/office/ref_office.ps1', '-Raw', 'audit/office/corpus/raw', '-Corpus', 'audit/office/corpus', '-RefDir', 'docs/quality-audit/out/ref'], { stdio: 'inherit' });
}

const env = { ...process.env, DOCVIEW_PORT: String(PORT), DOCVIEW_HOST: '127.0.0.1', DOCVIEW_DATA_DIR: path.join(run, 'data'), TMP: path.join(run, 'tmp'), TEMP: path.join(run, 'tmp'), DOCVIEW_LO_PROFILE_ROOT: path.join(run, 'lo') };
const server = spawn(process.execPath, ['node_modules/tsx/dist/cli.mjs', 'src/server.ts'], { env, stdio: ['ignore', 'pipe', 'pipe'] });
const log = fs.createWriteStream(path.join(run, 'server-audit.log'));
server.stdout.pipe(log); server.stderr.pipe(log);
try {
  const base = `http://127.0.0.1:${PORT}`;
  for (let i = 0; ; i += 1) {
    try { if ((await fetch(`${base}/api/formats`)).ok) break; } catch { /* 아직 안 뜸 */ }
    if (i > 120 || server.exitCode !== null) throw new Error('시험 서버가 뜨지 않았다 — .audit-run/server-audit.log 참고');
    await new Promise((r) => setTimeout(r, 500));
  }
  for (const ext of ONLY ? [ONLY] : ['docx', 'xlsx', 'pptx']) {
    const r = spawnSync(process.execPath, ['audit/office/capture.mjs', base, 'audit/office/corpus', 'docs/quality-audit/out', `.${ext}`], { stdio: 'inherit' });
    if (r.status !== 0) throw new Error(`capture ${ext} 실패`);
  }
} finally {
  server.kill(); // 이 스크립트가 띄운 서버(자식 PID)만 끈다 — 그 서버의 상주 LibreOffice는 표준 입력이 닫히면 스스로 끝난다
}
spawnSync('python', ['audit/office/compare.py', 'docs/quality-audit/out'], { stdio: 'inherit' });
console.log('점수: docs/quality-audit/out/scores.md · 나란히 그림: docs/quality-audit/out/sbs/');
