import { closeSync, cpSync, mkdirSync, openSync, rmSync, statSync } from 'node:fs';
import { build } from 'esbuild';

/**
 * 같은 폴더에서 두 서버(8000 run-webviewer.bat의 npm run dev, 8010 향기랩 bat의 npm start)가 로그온 때 동시에
 * 시작하면 이 빌드가 같은 public/js 파일을 동시에 쓰다 윈도 파일 잠금으로 한쪽이 실패했고, 실패한 쪽 서버는 뜨지
 * 않은 채 멈췄다(10-21 두 번 — 8010만 502). 잠금 파일로 한 번에 하나만 빌드하고, 나중 것은 앞 빌드가 끝나길 기다린다.
 * 잠금이 2분 넘게 남아 있으면(빌드 도중 죽은 경우) 지우고 진행한다.
 */
mkdirSync('public/js', { recursive: true });
const LOCK = 'public/js/.build.lock';
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
for (let waited = 0; ; waited += 250) {
  try {
    closeSync(openSync(LOCK, 'wx'));
    break;
  } catch {
    let stale = false;
    try {
      stale = Date.now() - statSync(LOCK).mtimeMs > 120_000;
    } catch {
      continue; // 그 사이 풀렸다 — 바로 다시 시도
    }
    if (stale || waited > 180_000) rmSync(LOCK, { force: true });
    else await sleep(250);
  }
}
process.on('exit', () => rmSync(LOCK, { force: true }));

/** 잠금이 있어도 백신 검사 등으로 파일이 잠깐 잠길 수 있다 — 한 번 쉬었다 다시 */
async function retry(fn) {
  try {
    return await fn();
  } catch (error) {
    console.warn('빌드 실패 — 2초 뒤 한 번 더:', error?.message ?? error);
    await sleep(2000);
    return fn();
  }
}

/**
 * 편집기 4개(docx: editor.ts, xlsx: grid-editor.ts, pptx: editor-slides.ts, hwp: editor-hwp.ts)를
 * 한 번의 splitting 빌드로 묶는다. 이 프로젝트의 나머지 프런트엔드(main.js, viewer.js)는 지금까지처럼
 * 번들 없이 <script> 태그로 직접 로드된다 — ProseMirror/Yjs만 npm 패키지 형태로 배포되어 번들이
 * 필요하다. 외부 CDN 없이 오프라인 번들로 묶어 기존 "원본 비전송·자체 호스팅" 설계 기조를 유지한다.
 *
 * editor.ts/grid-editor.ts/editor-slides.ts 셋 다 prosemirror-state/view/keymap과 yjs/y-websocket을
 * 겹쳐서 쓴다(editor-hwp.ts는 이 라이브러리들을 쓰지 않아 겹치지 않는다). 예전에는 4개를 각자
 * 독립된 outfile로 따로 빌드해(build() 4번 호출) 이 공유 라이브러리들이 번들마다 통째로 다시
 * 들어갔다 — 같은 세션에서 docx 에디터를 열었다가 pptx 에디터를 열면 겹치는 ~600~700KB를
 * 브라우저 캐시가 전혀 재사용하지 못하고 매번 새로 받았다(파일명이 다른 별개 리소스라 캐시
 * 식별자가 다름). splitting:true + 여러 entryPoints를 한 outdir에 함께 빌드하면 esbuild가 겹치는
 * 코드를 공유 chunk 파일로 뽑아내고 각 엔트리가 import로 그 chunk를 참조하게 만들어, 두 편집기를
 * 같은 세션에서 열어도 공유 chunk는 한 번만 받는다(파일명이 같은 하나의 리소스이므로 캐시가
 * 재사용됨).
 *
 * entryNames를 '[name].bundle'로 지정해 기존 outfile 이름 그대로(editor.bundle.js 등) 유지한다 —
 * public/*.html이 그 경로를 하드코딩하고 있어 바꾸면 전부 고쳐야 한다.
 */
await retry(() => build({
  bundle: true,
  splitting: true,
  format: 'esm',
  target: 'es2020',
  // pdf.js(PDF 편집기)가 모듈 최상위에서 await를 쓴다 — ESM 출력이라 브라우저가 그대로 실행한다(다른 편집기 코드는 안 씀)
  supported: { 'top-level-await': true },
  platform: 'browser',
  sourcemap: true,
  // 줄여서 내보낸다 — 받는 양과 브라우저가 읽는 시간이 준다. 함수·클래스 이름은 남긴다(이름에 기대는 코드·오류 추적용)
  minify: true,
  keepNames: true,
  logLevel: 'info',
  entryPoints: [
    'src/client/editor.ts',
    'src/client/grid-editor.ts',
    'src/client/editor-slides.ts',
    'src/client/editor-hwp.ts',
    'src/client/editor-pdf.ts',
  ],
  outdir: 'public/js',
  entryNames: '[name].bundle',
  chunkNames: 'chunks/[name]-[hash]',
}));

// PDF 편집기(editor-pdf.ts)가 쓰는 pdf.js 부속 — worker(별도 스레드에서 PDF 해석), cMap(글꼴이 안 들어 있는 한글·중국어·
// 일본어 PDF), 표준 글꼴. 외부 CDN 없이 자체 호스팅한다.
mkdirSync('public/pdfjs', { recursive: true });
cpSync('node_modules/pdfjs-dist/build/pdf.worker.min.mjs', 'public/js/pdf.worker.mjs');
cpSync('node_modules/pdfjs-dist/cmaps', 'public/pdfjs/cmaps', { recursive: true });
cpSync('node_modules/pdfjs-dist/standard_fonts', 'public/pdfjs/standard_fonts', { recursive: true });

// PDF 편집기 OCR(8단계, src/client/pdf/ocrUi.ts) — tesseract.js의 worker·wasm 코어(브라우저 성능별 3종 중 하나를 받는다)와
// 한국어·영어 학습 데이터(4.0.0_best_int). 외부 CDN 없이 자체 호스팅해 문서 그림이 밖으로 나가지 않게 한다.
mkdirSync('public/ocr/lang', { recursive: true });
cpSync('node_modules/tesseract.js/dist/worker.min.js', 'public/ocr/worker.min.js');
for (const core of ['tesseract-core-lstm.wasm.js', 'tesseract-core-simd-lstm.wasm.js', 'tesseract-core-relaxedsimd-lstm.wasm.js']) {
  cpSync(`node_modules/tesseract.js-core/${core}`, `public/ocr/${core}`);
}
for (const lang of ['kor', 'eng']) {
  cpSync(`node_modules/@tesseract.js-data/${lang}/4.0.0_best_int/${lang}.traineddata.gz`, `public/ocr/lang/${lang}.traineddata.gz`);
}
