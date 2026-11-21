/**
 * 아트크래프트 브라우저 엔진 자산 목록(public/artcraft/manifest.json, npm run build:artcraft가 만든다) — /api/view/info가
 * 해시 붙은 wasm·글꼴 주소를 여기서 읽는다. 파일이 없거나 틀리면 null → 아트크래프트 옵션을 끈 채 서버는 그대로 돈다(설계 §5).
 * 빌드를 서버가 뜬 뒤에 돌려도 되게 파일 수정 시각이 바뀌면 다시 읽는다(assetVersion.ts와 같은 방식).
 */
import { readFileSync, statSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export interface ArtcraftManifest {
  readonly word: string;
  readonly deck: string;
  readonly fonts: readonly string[];
}

const DEFAULT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'public', 'artcraft');
/** 같은 출처 정적 주소만 — manifest가 바뀌어도 Worker가 밖으로 요청하지 않게(설계 §6) */
const ASSET_URL = /^\/artcraft\/[A-Za-z0-9_./-]+$/;

/** 자산 폴더 — 시험은 DOCVIEW_ARTCRAFT_DIR로 바꾼다(부를 때마다 읽는다) */
export function artcraftDir(): string {
  return process.env.DOCVIEW_ARTCRAFT_DIR ?? DEFAULT_DIR;
}

export function parseManifest(raw: unknown): ArtcraftManifest | null {
  if (typeof raw !== 'object' || raw === null) return null;
  const { word, deck, fonts } = raw as Record<string, unknown>;
  const ok = (v: unknown): v is string => typeof v === 'string' && ASSET_URL.test(v) && !v.includes('..');
  if (!ok(word) || !ok(deck) || !Array.isArray(fonts) || fonts.length === 0 || !fonts.every(ok)) return null;
  return { word, deck, fonts: [...fonts] };
}

let cached: { readonly file: string; readonly mtimeMs: number; readonly manifest: ArtcraftManifest | null } | null = null;

export function readArtcraftManifest(dir: string = artcraftDir()): ArtcraftManifest | null {
  const file = path.join(dir, 'manifest.json');
  let mtimeMs: number;
  try {
    mtimeMs = statSync(file).mtimeMs;
  } catch {
    return null;
  }
  if (cached && cached.file === file && cached.mtimeMs === mtimeMs) return cached.manifest;
  let manifest: ArtcraftManifest | null;
  try {
    manifest = parseManifest(JSON.parse(readFileSync(file, 'utf8')));
  } catch {
    manifest = null;
  }
  cached = { file, mtimeMs, manifest };
  return manifest;
}
