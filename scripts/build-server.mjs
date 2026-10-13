import { build } from 'esbuild';
import { readdir } from 'node:fs/promises';
import { join } from 'node:path';

/**
 * npm으로 배포하는 패키지에는 서버측 TypeScript 원본을 그대로 담지 않는다(요청: 소스가 그대로
 * 보이지 않게) — 이 스크립트가 src/**\/*.ts 전체를 dist/ 밑에 컴파일+압축(minify)한다.
 *
 * src/client/(브라우저 전용, build-editor.mjs가 따로 번들)도 제외하지 않고 함께 컴파일한다 —
 * src/xlsx/formulaParserPool.ts처럼 서버측 파일이 src/client/formulaOverrides.ts(수식 함수 재정의,
 * 클라이언트 미리보기와 서버 계산이 같은 동작을 하도록 공유)를 직접 import하는 경우가 있어서,
 * "클라이언트 전용"과 "서버가 실제로 쓰는 것"의 경계가 디렉터리 기준으로 깔끔하게 안 나뉜다.
 * 어떤 파일이 서버에서 쓰이는지 일일이 추적하는 대신 전부 컴파일해 두면(쓰이지 않는 파일은 그냥
 * dist/에 남아있을 뿐 실행에 영향 없음) 이런 의존 관계가 나중에 하나 더 생겨도 조용히 깨지지 않는다.
 *
 * bundle:false(파일을 서로 합치지 않고 하나씩 변환)를 쓴다 — bundle:true로 여러 파일을 하나로
 * 합치면 그 안의 모든 import.meta.url이 (원래 파일이 어디 있었든) 합쳐진 산출물 하나의 위치를
 * 가리키게 된다. 이 프로젝트는 import.meta.url로 자기 파일 기준 상대 경로를 찾는 곳이 여럿이라
 * (app.ts의 publicDir, rendering.ts의 FONTS_DIR, dwgConverter.ts의 외부 도구 경로 등 — 전부 서로
 * 다른 깊이에 있다), 합치면 그중 일부가 조용히 엉뚱한 경로를 가리키게 된다. 파일 단위로 변환하면
 * dist/아래 구조가 src/와 정확히 같은 상대 깊이를 유지해 이 경로들이 그대로 맞는다.
 *
 * minify:true가 이름 압축·공백/주석 제거를 한다 — 완전한 난독화(제어 흐름 평탄화, 문자열 암호화
 * 등)는 아니지만 새 의존성 없이(esbuild는 이미 이 프로젝트의 의존성) 원본을 바로 읽지 못하게 하는
 * 수준은 된다.
 */

async function collectTsFiles(dir) {
  const entries = await readdir(dir, { withFileTypes: true });
  const files = [];
  for (const entry of entries) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      files.push(...(await collectTsFiles(full)));
    } else if (entry.name.endsWith('.ts')) {
      files.push(full);
    }
  }
  return files;
}

const entryPoints = await collectTsFiles('src');

await build({
  entryPoints,
  outdir: 'dist',
  outbase: 'src',
  bundle: false,
  minify: true,
  platform: 'node',
  target: 'node20',
  format: 'esm',
  sourcemap: false,
  logLevel: 'info',
});

console.log(`서버 ${entryPoints.length}개 파일을 dist/로 컴파일했습니다.`);
