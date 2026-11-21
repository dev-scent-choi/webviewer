/**
 * build:artcraft의 순수 도우미 — 시험(tests/artcraftBuild.test.ts)이 직접 부른다.
 * 크기 상한은 설계 §5(word ≤ 5 MB, deck ≤ 2.5 MB, 글꼴 2개 합 ≤ 0.6 MB, 모두 brotli). MB = 1024×1024.
 */
import { createHash } from 'node:crypto';
import { brotliCompressSync, constants } from 'node:zlib';

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
