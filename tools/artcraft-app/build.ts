/**
 * npm run build:artcraft-app — 아트크래프트 앱 화면(설계 2026-11-22-artcraft-app-screen §3, 시험 SPIKE-WEBAPP 부록 build-z.sh와 같은 방식).
 *  1. ARTCRAFT_DIR(기본 C:/Users/OpenMinds/projects/artcraft-eval)의 wordcraft·deckcraft·gridcraft가 tools/artcraft-shell/ARTCRAFT.lock 커밋
 *     그대로인지, rustc가 lock 판인지, wasm-bindgen이 0.2.129인지 본다.
 *  2. 세 앱(<앱>craft-web)을 wasm32-unknown-unknown으로 크기 최적화 빌드한다 — 저장소는 고치지 않고 환경 변수로만(opt z·fat LTO·CGU 1·strip).
 *  3. wasm-bindgen --target web --weak-refs --remove-name-section --remove-producers-section.
 *  4. 크기 상한(brotli)을 넘으면 아무것도 쓰지 않고 멈춘다. 통과하면 public/artcraft-app/를 비우고 <앱>craft-<내용 해시>/ 폴더(파일마다 .br) + manifest.json.
 * 첫 빌드는 앱마다 5~9분(세 앱 약 18분). shortcut: wasm-opt 없음(binaryen 미설치) — 10% 안팎 더 줄 여지, 설치하면 3 뒤에 넣는다.
 */
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { appSizeProblems, brotli, checkLock, containsBytes, craftFontsManifest, treeHash } from '../../scripts/artcraftBuildLib.js';

const ROOT = process.cwd();
const BUILD = path.join(ROOT, 'tools', 'artcraft-app', '.build');
const OUT = path.join(ROOT, 'public', 'artcraft-app');
const LOCK = path.join(ROOT, 'tools', 'artcraft-shell', 'ARTCRAFT.lock');
const ARTCRAFT_DIR = path.resolve(process.env.ARTCRAFT_DIR ?? 'C:/Users/OpenMinds/projects/artcraft-eval');
const WASM_BINDGEN = process.env.WASM_BINDGEN ?? path.join(ROOT, 'tools', 'wasm-bindgen', 'bin', 'wasm-bindgen.exe');
// wasm32 대상은 GNU 도구 묶음에만 설치돼 있다(build-artcraft.ts와 같다)
const TOOLCHAIN = '+stable-x86_64-pc-windows-gnu';
const APPS = [
  { id: 'word', repo: 'wordcraft' },
  { id: 'deck', repo: 'deckcraft' },
  { id: 'grid', repo: 'gridcraft' },
] as const;

function fail(message: string): never {
  console.error(`build:artcraft-app 실패 — ${message}`);
  process.exit(1);
}
// --no-optional-locks: status가 저장소의 .git/index를 새로 쓰지 않게
const git = (dir: string, ...args: string[]) => execFileSync('git', ['--no-optional-locks', '-C', dir, ...args], { encoding: 'utf8' });

const lock = JSON.parse(fs.readFileSync(LOCK, 'utf8')) as { rustc?: string; repos: Record<string, { url: string; commit: string }> };
const rustc = execFileSync('rustc', [TOOLCHAIN, '--version'], { encoding: 'utf8' }).trim();
if (!lock.rustc || !rustc.startsWith(`rustc ${lock.rustc} `)) fail(`rustc 판이 ARTCRAFT.lock(${lock.rustc ?? '없음'})과 다릅니다: ${rustc}`);
for (const { repo } of APPS) {
  const want = lock.repos[repo]?.commit;
  if (!want) fail(`ARTCRAFT.lock에 ${repo} 커밋이 없습니다`);
  const dir = path.join(ARTCRAFT_DIR, repo);
  if (!fs.existsSync(dir)) fail(`${dir} 없음 — ${lock.repos[repo].url}을(를) ${want}로 받아 두세요`);
  const problem = checkLock(repo, want, git(dir, 'rev-parse', 'HEAD'), git(dir, 'status', '--porcelain', '--untracked-files=normal'));
  if (problem) fail(problem);
}
if (!fs.existsSync(WASM_BINDGEN)) fail(`wasm-bindgen 없음(${WASM_BINDGEN}) — 앱 Cargo.lock과 같은 0.2.129를 받아 두세요`);
const wbVersion = execFileSync(WASM_BINDGEN, ['--version'], { encoding: 'utf8' }).trim();
if (wbVersion !== 'wasm-bindgen 0.2.129') fail(`${wbVersion} ≠ wasm-bindgen 0.2.129(앱 Cargo.lock 판과 같아야 한다)`);

// 사용자 환경의 RUSTFLAGS·글꼴·번역 폴더 변수가 빌드 결과를 바꾸지 않게 뺀다(spike: craft-i18n 없이 빌드 → 메뉴 영어)
const buildEnv: NodeJS.ProcessEnv = { ...process.env };
for (const k of ['RUSTFLAGS', 'CARGO_ENCODED_RUSTFLAGS', 'CRAFT_FONTS_DIR', 'CRAFT_I18N_DIR']) delete buildEnv[k];
Object.assign(buildEnv, {
  CARGO_PROFILE_RELEASE_OPT_LEVEL: 'z',
  CARGO_PROFILE_RELEASE_LTO: 'fat',
  CARGO_PROFILE_RELEASE_CODEGEN_UNITS: '1',
  CARGO_PROFILE_RELEASE_STRIP: 'true',
});

/** 격자 앱 칸 한글(계획 결정 12, Task 6을 Task 1에서 앞당김) — S-Core Dream 4를 craft-fonts 꼴 폴더로 만들어 CRAFT_FONTS_DIR로 준다(저장소는 고치지 않는다) */
const GRID_FONT = fs.readFileSync(path.join(ROOT, 'public', 'fonts', 's-core-dream', 'SCDream4.otf'));
function gridFontsDir(): string {
  const dir = path.join(BUILD, 'craft-fonts-ko');
  fs.mkdirSync(path.join(dir, 'fonts'), { recursive: true });
  fs.writeFileSync(path.join(dir, 'fonts', 'SCDream4.otf'), GRID_FONT);
  fs.writeFileSync(path.join(dir, 'fonts', 'LICENSE'), 'S-Core Dream (에스코어 드림) — public/fonts/s-core-dream 과 같은 조건으로 쓴다\n');
  fs.writeFileSync(path.join(dir, 'fonts', 'manifest.txt'), craftFontsManifest('fonts/SCDream4.otf', createHash('sha256').update(GRID_FONT).digest('hex')));
  return dir;
}

interface Built {
  readonly id: (typeof APPS)[number]['id'];
  readonly repo: string;
  readonly files: Array<{ rel: string; bytes: Buffer; br: Buffer }>;
}
const built: Built[] = [];
for (const a of APPS) {
  const target = path.join(BUILD, `target-${a.id}`);
  console.log(`[${a.id}] cargo build (5~9분)…`);
  // --locked: 저장소의 Cargo.lock을 고쳐 쓰지 않는다(다음 빌드의 잠금 검사가 "고친 파일"로 멈추지 않게)
  execFileSync('cargo', [TOOLCHAIN, 'build', '--release', '--locked', '--target', 'wasm32-unknown-unknown', '-p', `${a.repo}-web`], {
    cwd: path.join(ARTCRAFT_DIR, a.repo),
    stdio: 'inherit',
    env: { ...buildEnv, CARGO_TARGET_DIR: target, ...(a.id === 'grid' ? { CRAFT_FONTS_DIR: gridFontsDir() } : {}) },
  });
  const wbg = path.join(BUILD, `wbg-${a.id}`);
  fs.rmSync(wbg, { recursive: true, force: true });
  execFileSync(
    WASM_BINDGEN,
    ['--target', 'web', '--no-typescript', '--weak-refs', '--remove-name-section', '--remove-producers-section', '--out-dir', wbg, path.join(target, 'wasm32-unknown-unknown', 'release', `${a.repo}-web.wasm`)],
    { stdio: 'inherit' },
  );
  const rels = (fs.readdirSync(wbg, { recursive: true }) as string[])
    .map((r) => r.replaceAll('\\', '/'))
    .filter((r) => fs.statSync(path.join(wbg, r)).isFile());
  const files = rels.map((rel) => {
    const bytes = fs.readFileSync(path.join(wbg, rel));
    return { rel, bytes, br: brotli(bytes) };
  });
  if (!files.some((f) => f.rel === `${a.repo}-web.js`) || !files.some((f) => f.rel === `${a.repo}-web_bg.wasm`)) fail(`${a.id}: wasm-bindgen 결과에 접착 JS·wasm이 없습니다(${rels.join(', ')})`);
  // 글꼴이 정말 들어갔는지 — 빌드 입력이 조용히 무시되면(목록 거부·경로 틀림) 칸이 다시 ???가 된다
  if (a.id === 'grid' && !containsBytes(files.find((f) => f.rel.endsWith('_bg.wasm'))!.bytes, GRID_FONT.subarray(4096, 4160))) {
    fail('grid 앱 wasm에 한글 글꼴(S-Core Dream 4)이 들어가지 않았습니다 — CRAFT_FONTS_DIR manifest를 확인하세요');
  }
  built.push({ id: a.id, repo: a.repo, files });
}

const wasmBr = (b: Built) => b.files.find((f) => f.rel.endsWith('_bg.wasm'))!.br.length;
const byId = Object.fromEntries(built.map((b) => [b.id, wasmBr(b)])) as { word: number; deck: number; grid: number };
const problems = appSizeProblems(byId);
if (problems.length > 0) fail(`크기 상한 초과:\n  ${problems.join('\n  ')}`);

fs.rmSync(OUT, { recursive: true, force: true });
const manifest: Record<string, unknown> = {};
for (const b of built) {
  const dir = `${b.repo}-${treeHash(b.files)}`;
  for (const f of b.files) {
    const file = path.join(OUT, dir, f.rel);
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, f.bytes);
    fs.writeFileSync(`${file}.br`, f.br);
  }
  manifest[b.id] = `/artcraft-app/${dir}/${b.repo}-web.js`;
  console.log(`  ${dir}/  wasm br ${(wasmBr(b) / 1048576).toFixed(2)} MB`);
}
manifest.commits = Object.fromEntries(APPS.map((a) => [a.repo, lock.repos[a.repo].commit]));
fs.writeFileSync(path.join(OUT, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
console.log('build:artcraft-app 완료 → public/artcraft-app/manifest.json');
