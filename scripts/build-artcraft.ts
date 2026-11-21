/**
 * npm run build:artcraft — 아트크래프트 브라우저 엔진(설계 §5).
 *  1. ARTCRAFT_DIR(기본 C:/Users/OpenMinds/projects/artcraft-eval)의 wordcraft·deckcraft가 ARTCRAFT.lock 커밋 그대로인지 본다.
 *  2. tools/artcraft-shell/{word,deck}을 wasm32-unknown-unknown(GNU 도구 묶음, wasm-bindgen 없이)으로 빌드한다.
 *  3. wasm-bindgen CLI(WASM_BINDGEN 또는 tools/wasm-bindgen/bin/wasm-bindgen.exe)가 있으면 크기를 줄인다(없으면 건너뜀).
 *  4. 크기 상한을 넘으면 아무것도 쓰지 않고 멈춘다. 통과하면 public/artcraft/를 비우고 해시 이름 파일 + .br + manifest.json을 쓴다.
 */
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { brotli, checkLock, hashedName, sizeProblems } from './artcraftBuildLib.js';

const ROOT = process.cwd();
const SHELL = path.join(ROOT, 'tools', 'artcraft-shell');
const BUILD = path.join(SHELL, '.build');
const OUT = path.join(ROOT, 'public', 'artcraft');
const ARTCRAFT_DIR = path.resolve(process.env.ARTCRAFT_DIR ?? 'C:/Users/OpenMinds/projects/artcraft-eval');
const WASM_BINDGEN = process.env.WASM_BINDGEN ?? path.join(ROOT, 'tools', 'wasm-bindgen', 'bin', 'wasm-bindgen.exe');
/** wasm 최대 메모리(설계 §6 — 조작된 문서가 탭 메모리를 다 쓰지 못하게). 고객 문서 실측 35 MB, 별칭 8개 115 MB */
const MAX_MEMORY = 512 * 1024 * 1024;
const FONT_DIR = path.join(ROOT, 'public', 'fonts', 's-core-dream');
const CRATES = [
  { id: 'word', repo: 'wordcraft', lib: 'artcraft_word' },
  { id: 'deck', repo: 'deckcraft', lib: 'artcraft_deck' },
] as const;

const fwd = (p: string) => p.replace(/\\/g, '/');
function fail(message: string): never {
  console.error(`build:artcraft 실패 — ${message}`);
  process.exit(1);
}
const git = (dir: string, ...args: string[]) => execFileSync('git', ['-C', dir, ...args], { encoding: 'utf8' });

const lock = JSON.parse(fs.readFileSync(path.join(SHELL, 'ARTCRAFT.lock'), 'utf8')) as { repos: Record<string, { url: string; commit: string }> };
for (const { repo } of CRATES) {
  const want = lock.repos[repo]?.commit;
  if (!want) fail(`ARTCRAFT.lock에 ${repo} 커밋이 없습니다`);
  const dir = path.join(ARTCRAFT_DIR, repo);
  if (!fs.existsSync(dir)) fail(`${dir} 없음 — ${lock.repos[repo].url}을(를) ${want}로 받아 두세요`);
  const problem = checkLock(repo, want, git(dir, 'rev-parse', 'HEAD'), git(dir, 'status', '--porcelain', '--untracked-files=no'));
  if (problem) fail(problem);
}

const wasmBytes: Record<string, Buffer> = {};
for (const c of CRATES) {
  const dir = path.join(BUILD, c.id);
  fs.mkdirSync(dir, { recursive: true });
  const toml = fs
    .readFileSync(path.join(SHELL, c.id, 'Cargo.toml.in'), 'utf8')
    .replace(/@ARTCRAFT_DIR@/g, fwd(ARTCRAFT_DIR))
    .replace(/@SHELL@/g, fwd(SHELL));
  fs.writeFileSync(path.join(dir, 'Cargo.toml'), toml);
  fs.copyFileSync(path.join(SHELL, c.id, 'Cargo.lock'), path.join(dir, 'Cargo.lock'));
  console.log(`[${c.id}] cargo build (몇 분 걸린다)…`);
  // wasm32 대상은 GNU 도구 묶음에만 설치돼 있다(MSVC로 하면 "can't find crate for core" — spike §4)
  execFileSync('cargo', ['+stable-x86_64-pc-windows-gnu', 'build', '--release', '--locked', '--target', 'wasm32-unknown-unknown'], {
    cwd: dir,
    stdio: 'inherit',
    env: {
      ...process.env,
      CARGO_TARGET_DIR: path.join(BUILD, 'target'),
      CARGO_TARGET_WASM32_UNKNOWN_UNKNOWN_RUSTFLAGS: `-C link-arg=--max-memory=${MAX_MEMORY}`,
    },
  });
  let wasm = path.join(BUILD, 'target', 'wasm32-unknown-unknown', 'release', `${c.lib}.wasm`);
  if (fs.existsSync(WASM_BINDGEN)) {
    const out = path.join(BUILD, 'wbg');
    execFileSync(WASM_BINDGEN, ['--target', 'web', '--no-typescript', '--out-dir', out, wasm], { stdio: 'inherit' });
    wasm = path.join(out, `${c.lib}_bg.wasm`);
  } else {
    console.log(`[${c.id}] wasm-bindgen 없음(${WASM_BINDGEN}) — 크기 줄이기 건너뜀(약 10% 큼)`);
  }
  wasmBytes[c.id] = fs.readFileSync(wasm);
}

const fontBytes = ['SCDream4', 'SCDream6'].map((name) => ({ name, bytes: fs.readFileSync(path.join(FONT_DIR, `${name}.otf`)) }));
const br = {
  word: brotli(wasmBytes.word),
  deck: brotli(wasmBytes.deck),
  fonts: fontBytes.map((f) => brotli(f.bytes)),
};
const problems = sizeProblems({ word: br.word.length, deck: br.deck.length, fonts: br.fonts.map((b) => b.length) });
if (problems.length > 0) fail(`크기 상한 초과:\n  ${problems.join('\n  ')}`);

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(path.join(OUT, 'fonts'), { recursive: true });
function emit(sub: string, base: string, ext: string, raw: Buffer, compressed: Buffer): string {
  const name = hashedName(base, ext, raw);
  const file = path.join(OUT, sub, name);
  fs.writeFileSync(file, raw);
  fs.writeFileSync(`${file}.br`, compressed);
  console.log(`  ${sub ? `${sub}/` : ''}${name}  원본 ${(raw.length / 1048576).toFixed(2)} MB · br ${(compressed.length / 1048576).toFixed(2)} MB`);
  return `/artcraft/${sub ? `${sub}/` : ''}${name}`;
}
const manifest = {
  word: emit('', 'artcraft_word', '.wasm', wasmBytes.word, br.word),
  deck: emit('', 'artcraft_deck', '.wasm', wasmBytes.deck, br.deck),
  fonts: fontBytes.map((f, i) => emit('fonts', f.name, '.otf', f.bytes, br.fonts[i])),
  commits: Object.fromEntries(CRATES.map((c) => [c.repo, lock.repos[c.repo].commit])),
};
fs.writeFileSync(path.join(OUT, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
console.log('build:artcraft 완료 → public/artcraft/manifest.json');
