/**
 * build:artcraft의 순수 도우미 — 시험(tests/artcraftBuild.test.ts)이 직접 부른다.
 * 크기 상한은 설계 §5(word ≤ 5 MB, deck ≤ 2.5 MB, 글꼴 2개 합 ≤ 0.6 MB, 모두 brotli). MB = 1024×1024.
 */
import { createHash } from 'node:crypto';
import { brotliCompressSync, constants } from 'node:zlib';
import { APP_ASSET_PATH } from '../src/artcraftStatic.js';

export const MiB = 1024 * 1024;
export const SIZE_LIMITS = { word: 5 * MiB, deck: 2.5 * MiB, fonts: 0.6 * MiB } as const;

const mb = (n: number) => (n / MiB).toFixed(2);

/** `<base>-<내용 sha256 앞 10자리><ext>` — 이름이 바뀌어야 immutable 캐시가 새 파일을 받는다 */
export function hashedName(base: string, ext: string, bytes: Uint8Array): string {
  return `${base}-${createHash('sha256').update(bytes).digest('hex').slice(0, 10)}${ext}`;
}

/** 상한을 넘은 것마다 한 줄 — 빈 배열이면 통과 */
export function sizeProblems(br: { readonly word: number; readonly deck: number; readonly fonts: readonly number[] }): string[] {
  const out: string[] = [];
  if (br.word > SIZE_LIMITS.word) out.push(`word wasm brotli ${mb(br.word)} MB > ${mb(SIZE_LIMITS.word)} MB`);
  if (br.deck > SIZE_LIMITS.deck) out.push(`deck wasm brotli ${mb(br.deck)} MB > ${mb(SIZE_LIMITS.deck)} MB`);
  const fonts = br.fonts.reduce((a, b) => a + b, 0);
  if (fonts > SIZE_LIMITS.fonts) out.push(`글꼴 brotli 합 ${mb(fonts)} MB > ${mb(SIZE_LIMITS.fonts)} MB`);
  return out;
}

/** ArtCraft 저장소가 고정 커밋 그대로인지 — 문제면 이유, 아니면 null */
export function checkLock(repo: string, wantCommit: string, head: string, porcelain: string): string | null {
  if (head.trim() !== wantCommit) return `${repo}: HEAD ${head.trim().slice(0, 12)} ≠ ARTCRAFT.lock ${wantCommit.slice(0, 12)}`;
  if (porcelain.trim() !== '') return `${repo}: 고친 파일이 있습니다(git status) — 저장소는 고치지 않는다`;
  return null;
}

export function brotli(bytes: Uint8Array): Buffer {
  return brotliCompressSync(bytes, { params: { [constants.BROTLI_PARAM_QUALITY]: 11, [constants.BROTLI_PARAM_SIZE_HINT]: bytes.length } });
}

/** wasm 메모리 절(5)의 최대 크기(바이트) — 상한이 없거나 메모리 절이 없으면 null. 빌드 뒤 --max-memory가 정말 들어갔는지 본다 */
export function wasmMaxMemory(bytes: Uint8Array): number | null {
  let i = 8;
  const leb = (): number => {
    let v = 0;
    for (let shift = 0; ; shift += 7) {
      const b = bytes[i++];
      if (b === undefined) throw new Error('wasm이 잘렸습니다');
      v += (b & 0x7f) * 2 ** shift;
      if ((b & 0x80) === 0) return v;
    }
  };
  while (i < bytes.length) {
    const id = bytes[i++];
    const size = leb();
    if (id === 5) {
      if (leb() < 1) return null;
      const flags = leb();
      leb(); // min
      // 0x01 = max 있음(0x04 memory64는 쪽 수가 같은 단위라 그대로 읽힌다)
      return flags & 0x01 ? leb() * 65536 : null;
    }
    i += size;
  }
  return null;
}

/** 앱 화면(설계 2026-11-22-artcraft-app-screen §3) wasm brotli 상한 — 시험(SPIKE-WEBAPP) 실측 word 7.9 · deck 4.9 · grid 4.5 MB에 여유(격자는 한글 글꼴 약 0.3 MB 포함) */
export const APP_SIZE_LIMITS = { word: 8.5 * MiB, deck: 5.5 * MiB, grid: 5.5 * MiB } as const;

export function appSizeProblems(br: { readonly word: number; readonly deck: number; readonly grid: number }): string[] {
  return (['word', 'deck', 'grid'] as const)
    .filter((k) => br[k] > APP_SIZE_LIMITS[k])
    .map((k) => `${k} 앱 wasm brotli ${mb(br[k])} MB > ${mb(APP_SIZE_LIMITS[k])} MB`);
}

/** 앱 폴더 이름의 해시 — 파일 이름과 내용 모두(순서 무관). 접착 JS가 snippets를 상대 경로로 부르므로 폴더째 이름을 바꾼다 */
export function treeHash(files: ReadonlyArray<{ readonly rel: string; readonly bytes: Uint8Array }>): string {
  const h = createHash('sha256');
  for (const f of [...files].sort((a, b) => (a.rel < b.rel ? -1 : a.rel > b.rel ? 1 : 0))) {
    h.update(f.rel).update('\0').update(f.bytes).update('\0');
  }
  return h.digest('hex').slice(0, 10);
}

/**
 * gridcraft 웹 빌드의 글꼴 입력(craft-fonts fonts/manifest.txt) — 격자 앱은 실행 중 글꼴 추가(font.add)가 없고 이 허용 목록 글꼴만 wasm에 넣는다
 * (gridcraft crates/ui-egui/src/craft_fonts_policy.rs WEB_ALLOWED). S-Core Dream 4를 한글 자리(NanumGothic Regular, Hang,Kore)에 넣는다.
 * shortcut: 이름만 NanumGothic(허용 목록이 이름으로 고른다) — 앱 쪽에 글꼴 이름 설정이 생기면 제 이름으로 바꾼다
 */
export function craftFontsManifest(file: string, sha256hex: string): string {
  return `# docview: S-Core Dream 4 under the NanumGothic slot (gridcraft web allowlist)\nNanumGothic | Regular | ${file} | Hang,Kore | S-Core-Dream | fonts/LICENSE | ${sha256hex} | local\n`;
}

export function containsBytes(hay: Uint8Array, needle: Uint8Array): boolean {
  return Buffer.from(hay.buffer, hay.byteOffset, hay.byteLength).indexOf(Buffer.from(needle.buffer, needle.byteOffset, needle.byteLength)) >= 0;
}

/** 서버 허용 목록(APP_ASSET_PATH)에 걸리지 않는 산출 파일 — 있으면 그 파일이 404라 앱이 조용히 엔진 보기로 넘어간다(wasm-bindgen snippet 모양이 바뀐 경우 등) */
export function unservedAppAssets(dir: string, rels: readonly string[]): string[] {
  return rels.filter((rel) => !APP_ASSET_PATH.test(`/${dir}/${rel}`));
}
