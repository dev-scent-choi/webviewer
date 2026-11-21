/**
 * 여는 엔진 고르기(설계 §2·§3.1) — 토큰 발급과 /api/view/info가 같은 판정을 쓴다. 토큰 클레임에는 "요청한 엔진"만 넣고
 * (위조 방지 — 주소 쿼리는 보지 않는다) 판정은 부를 때마다 다시 한다(그사이 manifest가 사라져도 안전하게).
 * 원본 바이트가 브라우저로 내려가므로 내려받기 권한(canDownload)과 같은 뜻이다.
 */
import type { ArtcraftManifest } from './artcraftManifest.js';

export type Engine = 'libreoffice' | 'artcraft';

export const ENGINE_REASONS = {
  unsupported: '이 형식은 브라우저 엔진이 지원하지 않습니다',
  notInstalled: '브라우저 엔진이 설치되지 않았습니다',
  noDownload: '내려받기 권한이 없어 서버 변환으로 엽니다',
} as const;

/** 브라우저 엔진이 그리는 형식 — .doc는 표가 깨지는 파일이 있어 "실험"(설계 §2). .ppt·엑셀·한글은 엔진 없음 */
export const ARTCRAFT_FORMATS: ReadonlyMap<string, { readonly kind: 'word' | 'deck'; readonly experimental: boolean }> = new Map([
  ['docx', { kind: 'word', experimental: false }],
  ['pptx', { kind: 'deck', experimental: false }],
  ['doc', { kind: 'word', experimental: true }],
] as const);

export interface ArtcraftInfo {
  readonly kind: 'word' | 'deck';
  /** Worker가 엔진에 넘길 파일 이름의 확장자(엔진은 확장자로 형식을 고른다) */
  readonly ext: string;
  readonly experimental: boolean;
  /** 엔진 전용 원본 경로(토큰 포함, canDownload 검사 그대로) — 사용자 내려받기로 감사하지 않는다(/api/view/engine-source) */
  readonly originalUrl: string;
  readonly assets: { readonly wasm: string; readonly fonts: readonly string[] };
}

/** ext는 meta.ext(저장된 바이트의 실제 종류) — 옛 형식 편집 업로드도 engine-source가 내주는 바이트와 형식이 맞는다 */
export function engineFor(
  ext: string,
  perms: { readonly canDownload: boolean },
  requested: Engine | undefined,
  installed: boolean,
): { engine: Engine; reason?: string } {
  if (requested !== 'artcraft') return { engine: 'libreoffice' };
  if (!installed) return { engine: 'libreoffice', reason: ENGINE_REASONS.notInstalled };
  if (!ARTCRAFT_FORMATS.has(ext)) return { engine: 'libreoffice', reason: ENGINE_REASONS.unsupported };
  if (!perms.canDownload) return { engine: 'libreoffice', reason: ENGINE_REASONS.noDownload };
  return { engine: 'artcraft' };
}

export function artcraftInfo(ext: string, token: string, manifest: ArtcraftManifest): ArtcraftInfo | undefined {
  const format = ARTCRAFT_FORMATS.get(ext);
  if (!format) return undefined;
  return {
    kind: format.kind,
    ext,
    experimental: format.experimental,
    originalUrl: `/api/view/engine-source?token=${encodeURIComponent(token)}`,
    assets: { wasm: manifest[format.kind], fonts: manifest.fonts },
  };
}
