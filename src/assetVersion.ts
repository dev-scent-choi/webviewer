/**
 * 화면 HTML이 부르는 JS·CSS 주소에 그 파일의 내용 해시(?v=…)를 붙인다 — 서버는 정적 파일을 no-cache로
 * 내보내지만 앞단 Cloudflare가 브라우저 캐시를 4시간으로 덮어써서, 고친 뒤에도 사용자 브라우저가 예전
 * grid-editor.bundle.js·viewer.js·style.css로 그리는 일이 있었다("고친 게 다시 안 된다"). HTML은 캐시되지
 * 않으므로 파일 내용이 바뀌면 주소가 바뀌어 브라우저가 반드시 새 파일을 받는다.
 *
 * HTML 원본은 건드리지 않고 내보낼 때만 고친다 — 빌드마다 HTML 파일을 다시 쓰면 운영 저장소가 늘
 * "수정됨" 상태가 되어 병합이 꼬인다. 해시는 파일 수정 시각이 바뀔 때만 다시 계산한다.
 */
import { createHash } from 'node:crypto';
import { readFileSync, statSync } from 'node:fs';
import path from 'node:path';

const cache = new Map<string, { mtimeMs: number; version: string }>();

function versionOf(publicDir: string, urlPath: string): string | null {
  const file = path.join(publicDir, urlPath);
  if (!file.startsWith(publicDir)) return null;
  try {
    const { mtimeMs } = statSync(file);
    const hit = cache.get(file);
    if (hit && hit.mtimeMs === mtimeMs) return hit.version;
    const version = createHash('sha256').update(readFileSync(file)).digest('hex').slice(0, 10);
    cache.set(file, { mtimeMs, version });
    return version;
  } catch {
    return null;
  }
}

export function versionAssetUrls(html: string, publicDir: string): string {
  return html.replace(/((?:src|href)=")(\/(?:js|css)\/[^"?#]+)"/g, (whole, attr: string, url: string) => {
    const version = versionOf(publicDir, url);
    return version ? `${attr}${url}?v=${version}"` : whole;
  });
}
