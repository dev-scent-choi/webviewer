/**
 * 아트크래프트 앱 화면 자산(public/artcraft-app/manifest.json, npm run build:artcraft-app가 만든다 — 설계 2026-11-22-artcraft-app-screen §3).
 * 없거나 틀리면 null → 앱 화면은 꺼지고 지금의 엔진 보기·서버 변환으로 연다. 워드·PPT 본문 한글은 엔진 manifest의 S-Core Dream 글꼴을
 * font.add로 넣어야 나오므로(spike §4) 엔진 manifest도 있어야 켠다. 엑셀은 화면 글자를 브라우저 글꼴로 그려 글꼴이 필요 없다.
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { cachedJsonReader, type ArtcraftManifest } from './artcraftManifest.js';

export type AppKind = 'word' | 'deck' | 'grid';

export interface ArtcraftAppManifest {
  readonly word: string;
  readonly deck: string;
  readonly grid: string;
}

/** /api/view/info의 artcraftApp 블록 — 앱 화면(public/artcraft-app.html)이 쓴다 */
export interface ArtcraftAppInfo {
  readonly kind: AppKind;
  /** 앱 접착 JS(같은 출처, 폴더 이름에 내용 해시) */
  readonly js: string;
  /** font.add로 넣을 한글 글꼴(엔진 manifest) — 엑셀은 빈 배열 */
  readonly fonts: readonly string[];
  /** 엔진 전용 원본 경로(토큰 포함, canDownload + artcraft 검사 그대로) */
  readonly originalUrl: string;
}

/** 저장된 바이트 형식(meta.ext) → 앱. .doc는 워드 앱이 못 열면 :failed로 엔진 보기·서버 변환으로 넘어간다. .ppt·.xls 바이너리는 앱이 못 연다 */
export const APP_FORMATS: ReadonlyMap<string, AppKind> = new Map([
  ['docx', 'word'],
  ['doc', 'word'],
  ['pptx', 'deck'],
  ['xlsx', 'grid'],
] as const);

const KINDS = ['word', 'deck', 'grid'] as const;
const DEFAULT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'public', 'artcraft-app');

/** 자산 폴더 — 시험·앱 화면 끄기는 DOCVIEW_ARTCRAFT_APP_DIR로 바꾼다(부를 때마다 읽는다) */
export function artcraftAppDir(): string {
  return process.env.DOCVIEW_ARTCRAFT_APP_DIR ?? DEFAULT_DIR;
}

export function parseAppManifest(raw: unknown): ArtcraftAppManifest | null {
  if (typeof raw !== 'object' || raw === null) return null;
  const r = raw as Record<string, unknown>;
  // 같은 출처 해시 폴더의 그 앱 JS만 — manifest가 바뀌어도 화면이 밖이나 다른 앱 스크립트를 부르지 않게
  const ok = (k: AppKind) => typeof r[k] === 'string' && new RegExp(`^/artcraft-app/${k}craft-[0-9a-f]{10}/${k}craft-web\\.js$`).test(r[k] as string);
  if (!KINDS.every(ok)) return null;
  return { word: r.word as string, deck: r.deck as string, grid: r.grid as string };
}

const readAppManifestFile = cachedJsonReader(parseAppManifest);

export function readArtcraftAppManifest(dir: string = artcraftAppDir()): ArtcraftAppManifest | null {
  return readAppManifestFile(path.join(dir, 'manifest.json'));
}

export function appKindFor(ext: string, app: ArtcraftAppManifest | null, engine: ArtcraftManifest | null): AppKind | null {
  const kind = APP_FORMATS.get(ext);
  if (!kind || !app) return null;
  return kind === 'grid' || engine ? kind : null;
}

export function appInfo(ext: string, token: string, app: ArtcraftAppManifest, engine: ArtcraftManifest | null): ArtcraftAppInfo | undefined {
  const kind = appKindFor(ext, app, engine);
  if (!kind) return undefined;
  return {
    kind,
    js: app[kind],
    fonts: kind === 'grid' || !engine ? [] : engine.fonts,
    originalUrl: `/api/view/engine-source?token=${encodeURIComponent(token)}`,
  };
}
