# 아트크래프트 앱 화면 + 콘텐츠만 보기 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** "아트크래프트로 열기"를 아트크래프트 앱 자체 화면(같은 출처 iframe, `?host=parent`)으로 열고 저장까지 서버에 반영하며, 모든 화면에 도구를 전부 숨기는 `ui=content`(콘텐츠만 보기) 옵션을 더한다.

**Architecture:** `tools/artcraft-app/build.ts`가 세 앱(wordcraft-web·deckcraft-web·gridcraft-web)을 ARTCRAFT.lock 커밋 그대로 크기 최적화 빌드해 `public/artcraft-app/<앱>craft-<해시>/`(+ `.br`, `manifest.json`)에 쓰고, 서버는 `/artcraft-app/`를 허용 목록·immutable로 내준다. `/api/view/info`가 `artcraftApp` 블록(앱 JS·글꼴·원본 주소)을 주고, 새 화면 `public/artcraft-app.html`이 iframe(`public/artcraft-frame.html`)에 앱을 띄워 글꼴 → 문서를 postMessage로 넣고 `:save` 바이트를 `PUT /api/view/artcraft-save`로 올린다(실패하면 엔진 보기 → 리브레오피스). `ui=content`는 순수 판정 `isContentUi` + `body.ui-content` CSS 숨김 + 보기 규칙의 키 막기(편집기는 canEdit을 끄고 방에 `readonly=1`로 들어간다)로 만든다.

**Tech Stack:** TypeScript(Node 20, Fastify 5, @fastify/static), esbuild 번들(scripts/build-editor.mjs), vitest, Playwright(Chromium), cargo(+stable-x86_64-pc-windows-gnu, wasm32-unknown-unknown) + wasm-bindgen 0.2.129.

**Spec:** `docs/superpowers/specs/2026-11-22-artcraft-app-screen-design.md` (근거 `C:\Users\OpenMinds\projects\artcraft-eval\SPIKE-WEBAPP.md`, 시험 코드 `spike-webapp\build-z.sh`·`host.html`·`func.mjs`·`run.mjs`). 제목 줄 규격은 `docs/superpowers/specs/2026-11-22-artcraft-ribbon-skin-design.md`(24px, `scripts/ribbonSkinLib.mjs` `SKIN_PX.title`).

## Global Constraints

- 고객 파일은 커밋하지 않는다(시험은 `audit/office/corpus/`의 합성 문서만 쓴다).
- 빌드 산출물은 git 밖: `public/artcraft-app/`, `tools/artcraft-app/.build/`, `public/js/artcraftApp.bundle.js(.map)`, `public/js/artcraftFrame.bundle.js(.map)`(.gitignore에 넣는다).
- 포트 8000·8010·8020·8100·8110·8130·8148은 쓰지 않는다. 이 계획의 시험 서버는 8170~8179만(8171은 시험(spike) 서버가 잡고 있을 수 있으니 8172부터). 다른 계획의 기존 e2e는 원래 범위 그대로.
- 프로세스는 내가 띄운 PID만 끈다(`taskkill /PID <pid> /T /F`). node를 이름으로 일괄 종료하지 않는다.
- 문서를 외부로 올리지 않는다(앱·글꼴·문서 모두 같은 출처, 외부 요청 없음).
- 화면 글자는 한국어(우리 띠·안내·오류). 앱 메뉴는 영어 그대로(spec §3, 한글화는 2차).
- 커밋은 경로를 지정해서(`git commit -F - -- <경로…>`), 메시지는 한국어, 끝 줄 `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`. 구현자는 push하지 않는다.
- 같은 가지에서 리본 꾸밈 작업이 진행 중이다 — 리본 마크업(`public/editor*.html`, `public/view.html`의 리본·제목 줄 구조, `ribbonMenus.ts`)은 고치지 않는다. `public/css/style.css`는 **파일 끝에 덧붙이기만** 한다.
- 기존 e2e는 배포 기본 동작(앱 화면 켜짐) 그대로 돌린다 — 환경 변수로 앱 화면을 끄고 옛 검사를 통과시키지 않는다. "아트크래프트로 열기"가 사용자 지시로 바뀌었으므로 **verify:artcraft·verify:readonly에서 그 단추·artcraft 토큰 viewerUrl을 다루는 단언만** 새 동작(앱 화면)으로 고친다. 엔진 보기 단언은 지우지 않고 엔진 보기 주소(`/view.html?token=…` = 앱 실패 때 가는 배포 경로)로 옮겨 살린다(Task 5 Step 1에 단언별로 적음). 나머지(verify:docx·pptx·xlsx·xlsx:browser·view-only·view-shell·ribbon-skin)는 스크립트를 고치지 않고 통과해야 한다.
- 시험은 묶음 끝(Task 5)에 한 번에: Task 1~4는 단위 시험·`npm run typecheck`·`npm run build:editor`까지만.

## 이 계획이 정한 것(spec이 비워 둔 곳)

1. **앱 자산 경로**: spec의 `<app>-<hash>.{wasm,js}` 대신 폴더 `public/artcraft-app/<앱>craft-<해시>/`(안의 이름은 wasm-bindgen 그대로). 접착 JS가 `./snippets/eframe-…/inline0.js`를 상대 경로로 불러서 파일 이름만 바꾸면 깨진다. 해시는 폴더 안 모든 파일 내용으로.
2. **토큰 응답**: `engine:'artcraft'`이고 앱이 되는 형식이면 `viewerUrl` 자체가 `/artcraft-app.html?token=…`(spec "기본 보기 = 앱 화면"). 따로 `appUrl`은 두지 않는다. `/api/view/info`의 `viewerUrl`은 지금 그대로(앱 실패 때 갈 엔진 보기·보기 화면).
3. **xlsx**: 앱 화면만 있고 엔진 보기는 없다 → 실패 전환은 `/viewer.html`(서버 격자). `ui=content`의 엑셀은 격자 편집기(`/editor-grid.html?…&ui=content`).
4. **워드·PPT 앱 화면은 엔진 manifest(글꼴)가 있어야** 켠다(본문 한글이 ☒가 되지 않게). 엑셀은 글꼴이 필요 없다.
5. **읽기 전용 암호**: 브라우저가 `crypto.getRandomValues`로 만든다(서버 왕복 없음). 엑셀은 `docInspect().sheets` 수만큼 `review.protectSheet {sheet, password}`.
6. **문서가 열렸는지**: 앱이 "열림" 메시지를 주지 않으므로 시험 코드(run.mjs)처럼 `uiInspect()`에 문서 이름이 나올 때까지 100ms마다 본다(15초). 읽기 전용인데 확인을 못 하면 보호를 못 걸므로 실패 전환.
7. **저장 규칙**: 버전 저장소는 아직 없다(버전 이력 기능은 보류) → hwp-save와 같은 규칙(canEdit·zip 폭탄·`edit_saved` 감사·보기 자료 다시 만들기) + artcraft 토큰만 + 같은 OOXML 형식만 + 편집 방이 열려 있으면 409. `saveAs:true`(다른 이름·내보내기)는 무시하고 안내. `.doc`(옛 형식 그대로 저장된 문서)는 저장하지 않는다.
8. **`ui=content` + 편집 토큰**: 화면에서 canEdit을 끄고, 편집 방 WebSocket에 `readonly=1`을 붙여 서버도 이 연결을 읽기 전용으로 둔다(권한을 넓히지 않고 좁히기만 — spec §4 "권한은 토큰이 정한다"와 충돌하지 않음). 엑셀 시트 탭은 내용 이동이라 남기고 시트 추가 단추만 숨긴다.
9. **메인 화면 "콘텐츠만 보기"**: `<a class="btn">` 링크(새 화면으로 가는 이동). verify:readonly가 세는 `.doc-info > button` 목록을 바꾸지 않는다. 한글·도면처럼 콘텐츠 화면이 없는 형식엔 없다.
10. **앱 화면 주소에 `ui=content`가 붙어 오면** 워드·PPT는 엔진 보기(`/view.html?…&ui=content`), 엑셀은 격자 편집기로 넘긴다.
11. **기존 e2e(조정자 판정)**: 앱을 끄는 환경 변수 방식은 쓰지 않는다. verify:artcraft·verify:readonly의 "아트크래프트로 열기"·artcraft viewerUrl 단언만 앱 화면으로 바꾸고, 엔진 보기 단언은 `/view.html?token=…`로 옮긴다. `ui=content` 엔진 보기는 verify:content-only가, manifest 없음은 단위 시험(tests/artcraftAppRoutes·artcraftAppAssets)이 맡는다.

---

## File Structure

| 파일 | 할 일 |
|---|---|
| `tools/artcraft-app/build.ts` (새) | 세 앱 빌드·크기 상한·해시 폴더·manifest |
| `scripts/artcraftBuildLib.ts` | `APP_SIZE_LIMITS`·`appSizeProblems`·`treeHash` 추가 |
| `src/artcraftManifest.ts` | mtime 캐시를 `cachedJsonReader`로 꺼내 두 manifest가 같이 쓴다 |
| `src/artcraftApp.ts` (새) | 앱 manifest 읽기·`APP_FORMATS`·`appKindFor`·`appInfo` |
| `src/artcraftStatic.ts` | 접두·허용 목록을 인자로(`APP_ASSET_PATH` 추가) |
| `src/app.ts` | `/artcraft-app/` 등록, public 전체 서빙에서 `/artcraft-app/` 막기 |
| `src/artcraftEngine.ts` | `engineFor` 5번째 인자 `appOk` |
| `src/routes/documents.ts` | `contentUrlFor`, 토큰 응답 `viewerUrl`(앱)·`contentUrl` |
| `src/routes/viewer.ts` | info `artcraftApp`, engine-source 엑셀(앱), `PUT /api/view/artcraft-save` |
| `src/routes/edit.ts` | `?readonly=1`이면 읽기 전용으로 방에 |
| `src/client/artcraftAppLogic.ts` (새) | 앱 화면 순수 판정(주소·메시지·읽기 전용 명령·저장 판정·이름) |
| `src/client/artcraftApp.ts` (새) | 앱 화면 DOM(제목 줄·iframe·글꼴·열기·저장·전환) |
| `src/client/artcraftFrame.ts` (새) | iframe 문서 — 허용된 앱 JS만 불러 실행, 실패를 부모에 |
| `public/artcraft-app.html`, `public/artcraft-frame.html` (새) | 위 두 화면 |
| `src/client/viewStrip.ts` | `isContentUi`·`markContentUi`, `fallBackToViewer(url, text)` |
| `src/client/officeShell.ts` | `ui=content`면 보기 규칙 막기 |
| `src/client/editor.ts`·`grid-editor.ts`·`editor-slides.ts` | `markContentUi`, canEdit 끄기, WebSocket `readonly` |
| `src/client/view.ts`, `public/css/view.css` | view.html의 `ui=content` |
| `public/css/style.css` (끝에 덧붙임) | 편집기 `body.ui-content` 숨김 |
| `public/js/main.js` | 엑셀도 아트크래프트, 앱 화면 "편집 열기" 키, "콘텐츠만 보기" 링크 |
| `scripts/build-editor.mjs`, `package.json`, `.gitignore`, `vitest.config.ts` | 진입점·명령·무시 목록·시험 환경 |
| `scripts/verify-artcraft-app-e2e.mjs`, `scripts/verify-content-only-e2e.mjs` (새) | 브라우저 e2e |
| 시험: `tests/artcraftAppAssets.test.ts`(새), `tests/artcraftBuild.test.ts`, `tests/artcraftEngine.test.ts`, `tests/artcraftAppRoutes.test.ts`(새), `tests/artcraftAppLogic.test.ts`(새), `tests/viewStrip.test.ts` | |

---

### Task 1: 앱 빌드·정적 서빙·manifest(없으면 꺼짐)

**Files:**
- Create: `tools/artcraft-app/build.ts`, `src/artcraftApp.ts`, `tests/artcraftAppAssets.test.ts`
- Modify: `scripts/artcraftBuildLib.ts`, `src/artcraftManifest.ts:33-56`, `src/artcraftStatic.ts`, `src/app.ts:8-9,134-140`, `tests/artcraftBuild.test.ts`, `vitest.config.ts`, `.gitignore`, `package.json` (scripts)

**Interfaces:**
- Produces (`scripts/artcraftBuildLib.ts`): `APP_SIZE_LIMITS: { word: number; deck: number; grid: number }`, `appSizeProblems(br: { word: number; deck: number; grid: number }): string[]`, `treeHash(files: ReadonlyArray<{ rel: string; bytes: Uint8Array }>): string`(16진 10자리)
- Produces (`src/artcraftManifest.ts`): `cachedJsonReader<T>(parse: (raw: unknown) => T | null): (file: string) => T | null`
- Produces (`src/artcraftApp.ts`): `type AppKind = 'word' | 'deck' | 'grid'`, `interface ArtcraftAppManifest { word: string; deck: string; grid: string }`(앱 JS 주소), `interface ArtcraftAppInfo { kind: AppKind; js: string; fonts: readonly string[]; originalUrl: string }`, `APP_FORMATS: ReadonlyMap<string, AppKind>`, `artcraftAppDir(): string`, `parseAppManifest(raw: unknown): ArtcraftAppManifest | null`, `readArtcraftAppManifest(dir?: string): ArtcraftAppManifest | null`, `appKindFor(ext: string, app: ArtcraftAppManifest | null, engine: ArtcraftManifest | null): AppKind | null`, `appInfo(ext: string, token: string, app: ArtcraftAppManifest, engine: ArtcraftManifest | null): ArtcraftAppInfo | undefined`
- Produces (`src/artcraftStatic.ts`): `APP_ASSET_PATH: RegExp`, `registerArtcraftStatic(app, root, prefix = '/artcraft/', assetPath = ASSET_PATH)`
- 앱 manifest 모양(빌드가 쓰고 서버가 읽는다): `{ "word": "/artcraft-app/wordcraft-<10hex>/wordcraft-web.js", "deck": "/artcraft-app/deckcraft-<10hex>/deckcraft-web.js", "grid": "/artcraft-app/gridcraft-<10hex>/gridcraft-web.js", "commits": {…} }`

- [ ] **Step 1: 빌드 도우미 시험을 먼저 쓴다** — `tests/artcraftBuild.test.ts`의 import 줄을 바꾸고 describe 안 끝에 시험 두 개를 덧붙인다.

```ts
import { APP_SIZE_LIMITS, MiB, SIZE_LIMITS, appSizeProblems, brotli, checkLock, hashedName, sizeProblems, treeHash, wasmMaxMemory } from '../scripts/artcraftBuildLib.js';
```

```ts
  test('앱 화면 크기 상한(brotli wasm): word 8.5 · deck 5.5 · grid 5 MB — 시험 실측 7.9/4.9/4.5 MB + 여유', () => {
    expect(APP_SIZE_LIMITS).toEqual({ word: 8.5 * MiB, deck: 5.5 * MiB, grid: 5 * MiB });
    expect(appSizeProblems({ word: 8.5 * MiB, deck: 5.5 * MiB, grid: 5 * MiB })).toEqual([]);
    const over = appSizeProblems({ word: 8.5 * MiB + 1, deck: 1, grid: 5 * MiB + 1 });
    expect(over).toHaveLength(2);
    expect(over[0]).toContain('word');
    expect(over[1]).toContain('grid');
  });

  test('폴더 해시: 같은 파일 묶음은 같은 값(순서 무관), 이름이나 내용이 바뀌면 바뀐다', () => {
    const a = [
      { rel: 'wordcraft-web.js', bytes: Buffer.from('js') },
      { rel: 'snippets/eframe-1/inline0.js', bytes: Buffer.from('s') },
    ];
    const h = treeHash(a);
    expect(h).toMatch(/^[0-9a-f]{10}$/);
    expect(treeHash([...a].reverse())).toBe(h);
    expect(treeHash([{ ...a[0], rel: 'other.js' }, a[1]])).not.toBe(h);
    expect(treeHash([{ ...a[0], bytes: Buffer.from('js2') }, a[1]])).not.toBe(h);
  });
```

- [ ] **Step 2: 앱 manifest·서빙 시험을 쓴다** — `tests/artcraftAppAssets.test.ts` 새 파일.

```ts
import { mkdirSync, mkdtempSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { brotliCompressSync } from 'node:zlib';
import Fastify, { type FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import { appKindFor, parseAppManifest, readArtcraftAppManifest } from '../src/artcraftApp.js';
import { APP_ASSET_PATH, registerArtcraftStatic } from '../src/artcraftStatic.js';

const APP = {
  word: '/artcraft-app/wordcraft-0123456789/wordcraft-web.js',
  deck: '/artcraft-app/deckcraft-abcdef0123/deckcraft-web.js',
  grid: '/artcraft-app/gridcraft-0123456789/gridcraft-web.js',
  commits: { wordcraft: 'f8ec3cc0' },
};
const ENGINE = {
  word: '/artcraft/artcraft_word-0123456789.wasm',
  deck: '/artcraft/artcraft_deck-0123456789.wasm',
  fonts: ['/artcraft/fonts/SCDream4-0123456789.otf', '/artcraft/fonts/SCDream6-0123456789.otf'],
};
const tmp = (tag: string) => mkdtempSync(path.join(os.tmpdir(), `artcraft-app-${tag}-`));

describe('앱 manifest 읽기', () => {
  test('맞는 manifest는 word·deck·grid 주소만 돌려준다', () => {
    expect(parseAppManifest(APP)).toEqual({ word: APP.word, deck: APP.deck, grid: APP.grid });
  });

  test('모양이 틀리거나 앱과 폴더가 다르거나 밖 주소·..이면 null', () => {
    expect(parseAppManifest(null)).toBeNull();
    expect(parseAppManifest({ ...APP, grid: undefined })).toBeNull();
    expect(parseAppManifest({ ...APP, word: '/artcraft-app/deckcraft-0123456789/deckcraft-web.js' })).toBeNull();
    expect(parseAppManifest({ ...APP, deck: 'https://evil.example/deckcraft-web.js' })).toBeNull();
    expect(parseAppManifest({ ...APP, word: '/artcraft-app/wordcraft-0123456789/../wordcraft-web.js' })).toBeNull();
  });

  test('파일이 없으면 null(앱 화면 꺼짐), 있으면 읽는다, 깨진 JSON은 null', () => {
    expect(readArtcraftAppManifest(tmp('none'))).toBeNull();
    const ok = tmp('ok');
    writeFileSync(path.join(ok, 'manifest.json'), JSON.stringify(APP));
    expect(readArtcraftAppManifest(ok)?.grid).toBe(APP.grid);
    const broken = tmp('broken');
    writeFileSync(path.join(broken, 'manifest.json'), '{ 깨짐');
    expect(readArtcraftAppManifest(broken)).toBeNull();
  });
});

describe('appKindFor — 앱 화면으로 열 형식', () => {
  const app = parseAppManifest(APP);
  test('manifest가 없으면 어떤 형식도 앱 화면이 아니다(옵션 꺼짐)', () => {
    for (const ext of ['docx', 'doc', 'pptx', 'xlsx']) expect(appKindFor(ext, null, ENGINE)).toBeNull();
  });
  test('docx·doc → word, pptx → deck, xlsx → grid, ppt·xls·hwp·pdf·constructor → null', () => {
    expect(appKindFor('docx', app, ENGINE)).toBe('word');
    expect(appKindFor('doc', app, ENGINE)).toBe('word');
    expect(appKindFor('pptx', app, ENGINE)).toBe('deck');
    expect(appKindFor('xlsx', app, ENGINE)).toBe('grid');
    for (const ext of ['ppt', 'xls', 'hwp', 'pdf', 'constructor']) expect(appKindFor(ext, app, ENGINE)).toBeNull();
  });
  test('엔진 manifest(한글 글꼴)가 없으면 워드·PPT는 끄고 엑셀만', () => {
    expect(appKindFor('docx', app, null)).toBeNull();
    expect(appKindFor('pptx', app, null)).toBeNull();
    expect(appKindFor('xlsx', app, null)).toBe('grid');
  });
});

describe('/artcraft-app/* 정적 서빙', () => {
  let app: FastifyInstance;
  const dir = tmp('static');
  const js = Buffer.from('export default async function init() {} // 아트크래프트 '.repeat(50));
  const wasm = Buffer.from('\0asm 앱 '.repeat(300));
  beforeAll(async () => {
    const d = path.join(dir, 'wordcraft-0123456789');
    mkdirSync(path.join(d, 'snippets', 'eframe-d31a8450ea12dfc1'), { recursive: true });
    writeFileSync(path.join(d, 'wordcraft-web.js'), js);
    writeFileSync(path.join(d, 'wordcraft-web.js.br'), brotliCompressSync(js));
    writeFileSync(path.join(d, 'wordcraft-web_bg.wasm'), wasm);
    writeFileSync(path.join(d, 'wordcraft-web_bg.wasm.br'), brotliCompressSync(wasm));
    writeFileSync(path.join(d, 'snippets', 'eframe-d31a8450ea12dfc1', 'inline0.js'), 'export {};');
    writeFileSync(path.join(dir, 'manifest.json'), JSON.stringify(APP));
    app = Fastify();
    await registerArtcraftStatic(app, dir, '/artcraft-app/', APP_ASSET_PATH);
    await app.ready();
  });
  afterAll(() => app.close());

  test('JS·wasm: br을 받으면 .br(content-encoding: br), immutable, wasm은 application/wasm, JS는 javascript', async () => {
    const w = await app.inject({ method: 'GET', url: '/artcraft-app/wordcraft-0123456789/wordcraft-web_bg.wasm', headers: { 'accept-encoding': 'br' } });
    expect(w.statusCode).toBe(200);
    expect(w.headers['content-encoding']).toBe('br');
    expect(w.headers['content-type']).toBe('application/wasm');
    expect(w.headers['cache-control']).toBe('public, max-age=31536000, immutable');
    const j = await app.inject({ method: 'GET', url: '/artcraft-app/wordcraft-0123456789/wordcraft-web.js', headers: { 'accept-encoding': 'identity' } });
    expect(j.statusCode).toBe(200);
    expect(String(j.headers['content-type'])).toContain('javascript');
    expect(Buffer.compare(j.rawPayload, js)).toBe(0);
  });

  test('접착 JS가 상대 경로로 부르는 snippets도 내준다', async () => {
    expect((await app.inject({ method: 'GET', url: '/artcraft-app/wordcraft-0123456789/snippets/eframe-d31a8450ea12dfc1/inline0.js' })).statusCode).toBe(200);
  });

  test('manifest·.br 직접·다른 이름·.. 은 404', async () => {
    for (const url of [
      '/artcraft-app/manifest.json',
      '/artcraft-app//manifest.json',
      '/artcraft-app/Manifest.json',
      '/artcraft-app/wordcraft-0123456789/wordcraft-web.js.br',
      '/artcraft-app/wordcraft-0123456789/../manifest.json',
      '/artcraft-app/wordcraft-0123456789/other.js',
      '/artcraft-app/wordcraft-XYZ/wordcraft-web.js',
    ]) {
      expect((await app.inject({ method: 'GET', url, headers: { 'accept-encoding': 'br' } })).statusCode, url).toBe(404);
    }
  });
});

describe('buildApp — public 전체 서빙으로 /artcraft-app/ manifest가 새지 않는다', () => {
  let app: FastifyInstance;
  beforeAll(async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_ARTCRAFT_APP_DIR', tmp('buildapp'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  test('대소문자·// 를 바꿔도 404', async () => {
    for (const url of ['/artcraft-app/manifest.json', '/ARTCRAFT-APP/manifest.json', '/Artcraft-App/manifest.json', '/artcraft-app//manifest.json', '/artcraft-app']) {
      expect((await app.inject({ method: 'GET', url })).statusCode, url).toBe(404);
    }
  });
});
```

- [ ] **Step 3: 시험이 실패하는지 본다**

Run: `npx vitest run tests/artcraftBuild.test.ts tests/artcraftAppAssets.test.ts`
Expected: FAIL — `appSizeProblems`/`treeHash` 없음, `../src/artcraftApp.js` 없음, `APP_ASSET_PATH` 없음.

- [ ] **Step 4: 빌드 도우미를 더한다** — `scripts/artcraftBuildLib.ts` 끝에 덧붙인다.

```ts
/** 앱 화면(설계 2026-11-22-artcraft-app-screen §3) wasm brotli 상한 — 시험(SPIKE-WEBAPP) 실측 word 7.9 · deck 4.9 · grid 4.5 MB에 여유 */
export const APP_SIZE_LIMITS = { word: 8.5 * MiB, deck: 5.5 * MiB, grid: 5 * MiB } as const;

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
```

- [ ] **Step 5: manifest 캐시를 꺼내 함께 쓴다** — `src/artcraftManifest.ts`의 `let cached …`부터 파일 끝(`readArtcraftManifest`)까지를 아래로 바꾼다(동작 같음).

```ts
/** 파일 수정 시각이 바뀔 때만 다시 읽는 JSON 읽개 — 빌드를 서버가 뜬 뒤에 돌려도 된다. 없거나 깨지면 null */
export function cachedJsonReader<T>(parse: (raw: unknown) => T | null): (file: string) => T | null {
  let cached: { readonly file: string; readonly mtimeMs: number; readonly value: T | null } | null = null;
  return (file) => {
    let mtimeMs: number;
    try {
      mtimeMs = statSync(file).mtimeMs;
    } catch {
      return null;
    }
    if (cached && cached.file === file && cached.mtimeMs === mtimeMs) return cached.value;
    let value: T | null;
    try {
      value = parse(JSON.parse(readFileSync(file, 'utf8')));
    } catch {
      value = null;
    }
    cached = { file, mtimeMs, value };
    return value;
  };
}

const readManifestFile = cachedJsonReader(parseManifest);

export function readArtcraftManifest(dir: string = artcraftDir()): ArtcraftManifest | null {
  return readManifestFile(path.join(dir, 'manifest.json'));
}
```

- [ ] **Step 6: 앱 manifest 모듈을 만든다** — `src/artcraftApp.ts` 새 파일.

```ts
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
```

- [ ] **Step 7: 정적 서빙을 접두·허용 목록 인자로 바꾼다** — `src/artcraftStatic.ts`.

`const ASSET_PATH = …` 줄 아래에 덧붙인다:

```ts
/** /artcraft-app/ 허용 목록 — 해시 폴더 안 접착 JS·wasm, 접착 JS가 상대 경로로 부르는 snippets/<이름>/inline<n>.js만 */
export const APP_ASSET_PATH = /^\/(word|deck|grid)craft-[0-9a-f]{10}\/((word|deck|grid)craft-web(_bg)?\.(js|wasm)|snippets\/[A-Za-z0-9_-]+\/inline[0-9]+\.js)$/;
```

함수 머리와 두 줄을 바꾼다:

```ts
export async function registerArtcraftStatic(app: FastifyInstance, root: string, prefix = '/artcraft/', assetPath: RegExp = ASSET_PATH): Promise<void> {
```

`app.log.warn({ err, root }, 'artcraft 자산 폴더를 만들 수 없어 /artcraft/ 서빙을 건너뜁니다');` → `app.log.warn({ err, root }, \`artcraft 자산 폴더를 만들 수 없어 ${prefix} 서빙을 건너뜁니다\`);`
`prefix: '/artcraft/',` → `prefix,`
`allowedPath: (pathName) => ASSET_PATH.test(pathName),` → `allowedPath: (pathName) => assetPath.test(pathName),`

- [ ] **Step 8: 서버에 등록한다** — `src/app.ts`.

import 두 줄을 바꾼다/더한다:

```ts
import { artcraftAppDir } from './artcraftApp.js';
import { artcraftDir } from './artcraftManifest.js';
import { APP_ASSET_PATH, registerArtcraftStatic } from './artcraftStatic.js';
```

`await registerArtcraftStatic(app, artcraftDir());` 바로 아래에:

```ts
  // 아트크래프트 앱 화면 자산(해시 폴더·.br·immutable, 설계 2026-11-22-artcraft-app-screen §3)
  await registerArtcraftStatic(app, artcraftAppDir(), '/artcraft-app/', APP_ASSET_PATH);
```

public 전체 서빙의 `allowedPath`를 바꾼다(`/artcraft-app/manifest.json`이 public 전체 서빙으로 새지 않게):

```ts
    allowedPath: (p) => !/^\/artcraft(-app)?(\/|$)/i.test(path.posix.normalize(p.replaceAll('\\', '/'))),
```

- [ ] **Step 9: 시험 환경에서 앱 화면을 끈다** — `vitest.config.ts`. 개발 PC에 `public/artcraft-app/`가 빌드돼 있어도 기존 시험(엑셀은 아트크래프트 안 됨 등)이 예전처럼 돌게.

맨 위 import에 `import os from 'node:os';`·`import path from 'node:path';`를 더하고 `env` 줄을 바꾼다:

```ts
    // 아트크래프트 앱 화면은 끈다(빈 임시 폴더) — 켠 상태는 tests/artcraftAppRoutes.test.ts가 값을 직접 넣어 시험한다
    env: { DOCVIEW_LO_WORKER: '0', DOCVIEW_ROOM_CLOSE_GRACE_MS: '0', DOCVIEW_ARTCRAFT_APP_DIR: path.join(os.tmpdir(), 'docview-test-no-artcraft-app') },
```

- [ ] **Step 10: 시험이 통과하는지 본다**

Run: `npx vitest run tests/artcraftBuild.test.ts tests/artcraftAppAssets.test.ts tests/artcraftAssets.test.ts tests/artcraftApp.test.ts tests/artcraftRoutes.test.ts`
Expected: PASS(기존 엔진 서빙·라우트 시험도 그대로).

- [ ] **Step 11: 빌드 스크립트를 쓴다** — `tools/artcraft-app/build.ts` 새 파일(시험의 `build-z.sh`와 같은 옵션).

```ts
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
import fs from 'node:fs';
import path from 'node:path';
import { appSizeProblems, brotli, checkLock, treeHash } from '../../scripts/artcraftBuildLib.js';

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
    env: { ...buildEnv, CARGO_TARGET_DIR: target },
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
```

- [ ] **Step 12: 명령과 무시 목록** — `package.json` scripts의 `"build:artcraft": …` 줄 아래에 `"build:artcraft-app": "tsx tools/artcraft-app/build.ts",`를 더한다. `.gitignore`의 `!/tools/artcraft-shell/` 아래에

```
!/tools/artcraft-app/
/tools/artcraft-app/.build/
```

를, `public/artcraft/` 줄 아래에

```
# 아트크래프트 앱 화면 산출물(npm run build:artcraft-app) — 폴더 이름에 내용 해시
public/artcraft-app/
```

를 더한다.

- [ ] **Step 13: 실제로 빌드하고 산출물이 git 밖인지 본다**(약 18분 — 백그라운드로 돌리고 끝나면 확인)

Run: `npm run build:artcraft-app`
Expected: 마지막 줄 `build:artcraft-app 완료 → public/artcraft-app/manifest.json`, wasm br이 word ≤ 8.5·deck ≤ 5.5·grid ≤ 5 MB.
Run: `git status --porcelain -- public/artcraft-app tools/artcraft-app/.build` → 출력 없음. `git status --porcelain -- tools/artcraft-app` → `?? tools/artcraft-app/build.ts`만.
Run: `git -C C:/Users/OpenMinds/projects/artcraft-eval/wordcraft status --porcelain` (deckcraft·gridcraft도) → 출력 없음(저장소를 고치지 않음).

- [ ] **Step 14: 타입 검사·전체 단위 시험**

Run: `npm run typecheck` → 오류 없음. Run: `npx vitest run` → PASS.

- [ ] **Step 15: 커밋**

```bash
git add tools/artcraft-app/build.ts src/artcraftApp.ts tests/artcraftAppAssets.test.ts
git commit -F - -- tools/artcraft-app/build.ts src/artcraftApp.ts tests/artcraftAppAssets.test.ts scripts/artcraftBuildLib.ts tests/artcraftBuild.test.ts src/artcraftManifest.ts src/artcraftStatic.ts src/app.ts vitest.config.ts .gitignore package.json <<'EOF'
앱 화면 1: 아트크래프트 세 앱 빌드(lock 고정·크기 최적화·brotli·해시 폴더·manifest)와 /artcraft-app/ 정적 서빙, manifest 없으면 꺼짐

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>
EOF
```

---

### Task 2: 서버 — info·토큰(viewerUrl·contentUrl)·engine-source·저장 경로

**Files:**
- Create: `tests/artcraftAppRoutes.test.ts`
- Modify: `src/artcraftEngine.ts:41-55`, `src/routes/documents.ts` (imports, `contentUrlFor` 새 함수, 토큰 라우트 `:526-591`), `src/routes/viewer.ts` (imports, info `:777-811`, engine-source `:1776`, 새 PUT 라우트), `tests/artcraftEngine.test.ts`

**Interfaces:**
- Consumes (Task 1): `readArtcraftAppManifest()`, `appKindFor(ext, app, engine)`, `appInfo(ext, token, app, engine)`, `ArtcraftAppInfo`
- Produces: `engineFor(ext, perms, requested, installed, appOk = false)`; `contentUrlFor(ext: string, viewKind: ViewKind | undefined, engine: Engine, perms: { readonly canEdit: boolean; readonly canDownload: boolean }, token: string): string | undefined` (documents.ts export)
- Produces (HTTP): 토큰 응답 `viewerUrl`(앱이 되면 `/artcraft-app.html?token=…`)·`contentUrl?`; `GET /api/view/info` → `data.artcraftApp?: { kind, js, fonts, originalUrl }`; `PUT /api/view/artcraft-save?token=…`(본문 `application/octet-stream`) → `200 { success: true, data: { pageCount } }` | 400·401·403·404·409·413 `{ success: false, error }`

- [ ] **Step 1: engineFor 표에 앱 줄을 더한다** — `tests/artcraftEngine.test.ts`의 `describe('engineFor …')` 안 끝에:

```ts
  test('앱 화면(appOk)이 되면 엔진이 안 그리는 형식도 artcraft — 내려받기 권한은 그대로 필요', () => {
    expect(engineFor('xlsx', { canDownload: true }, 'artcraft', true, true)).toEqual({ engine: 'artcraft' });
    expect(engineFor('xlsx', { canDownload: true }, 'artcraft', false, true)).toEqual({ engine: 'artcraft' });
    expect(engineFor('xlsx', { canDownload: false }, 'artcraft', true, true)).toEqual({ engine: 'libreoffice', reason: ENGINE_REASONS.noDownload });
    expect(engineFor('xlsx', { canDownload: true }, 'artcraft', true, false)).toEqual({ engine: 'libreoffice', reason: ENGINE_REASONS.unsupported });
    expect(engineFor('docx', { canDownload: true }, 'libreoffice', true, true)).toEqual({ engine: 'libreoffice' });
  });
```

- [ ] **Step 2: 라우트 시험을 쓴다** — `tests/artcraftAppRoutes.test.ts` 새 파일.

```ts
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import type { FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';

// 편집 방 열림 흉내 — 저장 경로가 방이 열린 문서를 거절하는지(실제 방을 열지 않고)
const { openRooms } = vi.hoisted(() => ({ openRooms: new Set<string>() }));
vi.mock('../src/collab/room.js', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/collab/room.js')>()),
  isRoomOpen: (docId: string) => openRooms.has(docId),
}));

type Perms = { canPrint?: boolean; canDownload?: boolean; canCopy?: boolean; canEdit?: boolean };

describe('아트크래프트 앱 화면 — 토큰·info·engine-source·저장', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let newDocumentId: (typeof import('../src/storage.js'))['newDocumentId'];
  let issueViewToken: (typeof import('../src/security.js'))['issueViewToken'];
  let contentUrlFor: (typeof import('../src/routes/documents.js'))['contentUrlFor'];
  const engineDir = mkdtempSync(path.join(os.tmpdir(), 'artcraft-engine-'));
  const appDir = mkdtempSync(path.join(os.tmpdir(), 'artcraft-appdir-'));
  const missing = mkdtempSync(path.join(os.tmpdir(), 'artcraft-missing-'));
  const dataDir = mkdtempSync(path.join(os.tmpdir(), 'artcraft-appdata-'));
  const ENGINE = {
    word: '/artcraft/artcraft_word-0123456789.wasm',
    deck: '/artcraft/artcraft_deck-0123456789.wasm',
    fonts: ['/artcraft/fonts/SCDream4-0123456789.otf', '/artcraft/fonts/SCDream6-0123456789.otf'],
  };
  const APP = {
    word: '/artcraft-app/wordcraft-0123456789/wordcraft-web.js',
    deck: '/artcraft-app/deckcraft-0123456789/deckcraft-web.js',
    grid: '/artcraft-app/gridcraft-0123456789/gridcraft-web.js',
  };

  beforeAll(async () => {
    writeFileSync(path.join(engineDir, 'manifest.json'), JSON.stringify(ENGINE));
    writeFileSync(path.join(appDir, 'manifest.json'), JSON.stringify(APP));
    mkdirSync(path.join(dataDir, 'audit'));
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_DATA_DIR', dataDir);
    vi.stubEnv('DOCVIEW_PERSIST_LOGS', '1');
    vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', engineDir);
    vi.stubEnv('DOCVIEW_ARTCRAFT_APP_DIR', appDir);
    ({ store, newDocumentId } = await import('../src/storage.js'));
    ({ issueViewToken } = await import('../src/security.js'));
    ({ contentUrlFor } = await import('../src/routes/documents.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  async function addDoc(ext: string, bytes: Buffer = Buffer.from(`원본-${ext}`), viewKind?: 'pages' | 'grid') {
    const id = newDocumentId();
    await store.saveOriginal(id, ext, bytes);
    await store.writeMeta({ id, name: `시험.${ext}`, ext, status: 'ready', pageCount: 1, createdAt: new Date().toISOString(), ...(viewKind ? { viewKind } : {}) });
    return id;
  }
  async function xlsxBytes(text: string): Promise<Buffer> {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('가').getCell('A1').value = text;
    return Buffer.from(await wb.xlsx.writeBuffer());
  }
  const issue = (id: string, body: Record<string, unknown>) =>
    app.inject({ method: 'POST', url: `/api/documents/${id}/tokens`, payload: { viewer: '시험자', ...body } });
  const tok = (docId: string, perms: Perms, engine?: 'artcraft') =>
    issueViewToken({ docId, viewer: '시험자', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...perms }, ...(engine ? { engine } : {}) }, 600);
  const info = async (token: string) => (await app.inject({ method: 'GET', url: `/api/view/info?token=${encodeURIComponent(token)}` })).json().data;
  const save = (token: string, body: Buffer) =>
    app.inject({ method: 'PUT', url: `/api/view/artcraft-save?token=${encodeURIComponent(token)}`, headers: { 'content-type': 'application/octet-stream' }, payload: body });
  const audits = (docId: string) =>
    readdirSync(path.join(dataDir, 'audit'))
      .flatMap((f) => readFileSync(path.join(dataDir, 'audit', f), 'utf8').trim().split('\n'))
      .map((l) => JSON.parse(l) as { type: string; docId: string; detail?: string })
      .filter((e) => e.docId === docId);

  describe('contentUrlFor — 같은 토큰 + ui=content', () => {
    const ro = { canEdit: false, canDownload: true };
    const none = { canEdit: false, canDownload: false };
    test.each([
      ['docx', 'pages', 'editor', ro, '/editor.html?token=t&ui=content'],
      ['xlsx', 'grid', 'editor', ro, '/editor-grid.html?token=t&ui=content'],
      ['pptx', 'pages', 'editor', ro, '/editor-slides.html?token=t&ui=content'],
      ['docx', 'pages', 'libreoffice', none, '/view.html?token=t&ui=content'],
      ['docx', 'pages', 'artcraft', ro, '/view.html?token=t&ui=content'],
      ['pdf', 'pages', 'libreoffice', none, '/view.html?token=t&ui=content'],
      ['xlsx', 'grid', 'artcraft', ro, '/editor-grid.html?token=t&ui=content'],
      ['xlsx', 'grid', 'libreoffice', ro, '/editor-grid.html?token=t&ui=content'],
      ['xlsx', 'grid', 'libreoffice', none, undefined],
      ['hwp', 'pages', 'libreoffice', ro, undefined],
      ['dwg', 'svg', 'libreoffice', ro, undefined],
    ] as const)('%s · %s · %s → %s', (ext, viewKind, engine, perms, want) => {
      expect(contentUrlFor(ext, viewKind, engine, perms, 't')).toBe(want);
    });
    test('토큰은 주소 인코딩', () => {
      expect(contentUrlFor('pdf', 'pages', 'libreoffice', none, 'a+b')).toBe('/view.html?token=a%2Bb&ui=content');
    });
  });

  test('토큰: docx + 내려받기 + artcraft → viewerUrl은 앱 화면, contentUrl은 엔진 쪽 그림 보기', async () => {
    const data = (await issue(await addDoc('docx'), { canDownload: true, engine: 'artcraft' })).json().data;
    expect(data.engine).toBe('artcraft');
    expect(data.viewerUrl).toBe(`/artcraft-app.html?token=${encodeURIComponent(data.token)}`);
    expect(data.contentUrl).toBe(`/view.html?token=${encodeURIComponent(data.token)}&ui=content`);
  });

  test('토큰: xlsx + 내려받기 + artcraft → 앱 화면(엔진은 엑셀을 못 그려도), contentUrl은 격자 편집기', async () => {
    const data = (await issue(await addDoc('xlsx', Buffer.from('x'), 'grid'), { canDownload: true, engine: 'artcraft' })).json().data;
    expect(data.engine).toBe('artcraft');
    expect(data.viewerUrl).toMatch(/^\/artcraft-app\.html\?token=/);
    expect(data.contentUrl).toMatch(/^\/editor-grid\.html\?token=[^&]+&ui=content$/);
  });

  test('토큰: 앱 manifest가 없으면 지금 그대로(docx → 엔진 보기, xlsx → 지원하지 않음)', async () => {
    vi.stubEnv('DOCVIEW_ARTCRAFT_APP_DIR', missing);
    try {
      const docx = (await issue(await addDoc('docx'), { canDownload: true, engine: 'artcraft' })).json().data;
      expect(docx.viewerUrl).toMatch(/^\/view\.html\?token=/);
      const xlsx = (await issue(await addDoc('xlsx', Buffer.from('x'), 'grid'), { canDownload: true, engine: 'artcraft' })).json().data;
      expect(xlsx).toMatchObject({ engine: 'libreoffice', engineReason: '이 형식은 브라우저 엔진이 지원하지 않습니다' });
    } finally {
      vi.stubEnv('DOCVIEW_ARTCRAFT_APP_DIR', appDir);
    }
  });

  test('토큰: 읽기 전용 편집기 화면(engine editor)의 contentUrl은 편집기 + ui=content', async () => {
    const data = (await issue(await addDoc('pptx'), { canDownload: true, engine: 'editor' })).json().data;
    expect(data.contentUrl).toBe(`/editor-slides.html?token=${encodeURIComponent(data.token)}&ui=content`);
  });

  test('info: artcraft 토큰이면 artcraftApp 블록(워드는 글꼴, 엑셀은 글꼴 없음), viewerUrl은 그대로 보기 화면', async () => {
    const d = await addDoc('docx');
    const t = tok(d, { canDownload: true }, 'artcraft');
    const i = await info(t);
    expect(i.artcraftApp).toEqual({ kind: 'word', js: APP.word, fonts: ENGINE.fonts, originalUrl: `/api/view/engine-source?token=${encodeURIComponent(t)}` });
    expect(i.viewerUrl).toMatch(/^\/view\.html\?token=/);
    const x = await info(tok(await addDoc('xlsx', Buffer.from('x'), 'grid'), { canDownload: true }, 'artcraft'));
    expect(x.artcraftApp).toMatchObject({ kind: 'grid', js: APP.grid, fonts: [] });
  });

  test('info: 내려받기 권한 없음·artcraft 요청 아님·엔진 글꼴 없음(워드)이면 artcraftApp 없음', async () => {
    expect((await info(tok(await addDoc('docx'), {}, 'artcraft'))).artcraftApp).toBeUndefined();
    expect((await info(tok(await addDoc('docx'), { canDownload: true }))).artcraftApp).toBeUndefined();
    vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', missing);
    try {
      expect((await info(tok(await addDoc('docx'), { canDownload: true }, 'artcraft'))).artcraftApp).toBeUndefined();
    } finally {
      vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', engineDir);
    }
  });

  test('engine-source: 앱이 있으면 엑셀 원본도 내준다(no-store, engine_open)', async () => {
    const id = await addDoc('xlsx', Buffer.from('엑셀-원본'), 'grid');
    const res = await app.inject({ method: 'GET', url: `/api/view/engine-source?token=${encodeURIComponent(tok(id, { canDownload: true }, 'artcraft'))}` });
    expect(res.statusCode).toBe(200);
    expect(res.rawPayload.toString('utf8')).toBe('엑셀-원본');
    expect(res.headers['cache-control']).toBe('private, no-store');
  });

  test('저장: 토큰 없음·잘못된 토큰 401', async () => {
    expect((await save('bad.token.x', Buffer.from('x'))).statusCode).toBe(401);
  });

  test('저장: artcraft 토큰이 아니거나 편집 권한이 없으면 403(edit_save_denied), 원본 그대로', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const before = await store.readOriginal(id, 'xlsx');
    const bytes = await xlsxBytes('바뀜');
    expect((await save(tok(id, { canDownload: true, canEdit: true }), bytes)).statusCode).toBe(403);
    expect((await save(tok(id, { canDownload: true }, 'artcraft'), bytes)).statusCode).toBe(403);
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), before)).toBe(0);
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_save_denied', detail: 'permission' }));
  });

  test('저장: 형식이 다른 바이트(docx 문서에 xlsx)는 400, 저장 안 되는 형식(pdf)도 400', async () => {
    const docx = await addDoc('docx');
    expect((await save(tok(docx, { canDownload: true, canEdit: true }, 'artcraft'), await xlsxBytes('x'))).statusCode).toBe(400);
    const pdf = await addDoc('pdf');
    expect((await save(tok(pdf, { canDownload: true, canEdit: true }, 'artcraft'), Buffer.from('%PDF'))).statusCode).toBe(400);
  });

  test('저장: 편집 방이 열려 있으면 409 + 안내, 원본 그대로', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const before = await store.readOriginal(id, 'xlsx');
    openRooms.add(id);
    try {
      const res = await save(tok(id, { canDownload: true, canEdit: true }, 'artcraft'), await xlsxBytes('바뀜'));
      expect(res.statusCode).toBe(409);
      expect(res.json().error).toContain('편집기');
    } finally {
      openRooms.delete(id);
    }
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), before)).toBe(0);
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_save_denied', detail: 'room-open' }));
  });

  test('저장: 엑셀 성공 → 원본 바뀜, 보기 격자 다시 만듦(시트 수), edit_saved(artcraft-app)', async () => {
    const id = await addDoc('xlsx', await xlsxBytes('원래'), 'grid');
    const bytes = await xlsxBytes('새 값');
    const res = await save(tok(id, { canDownload: true, canEdit: true }, 'artcraft'), bytes);
    expect(res.statusCode).toBe(200);
    expect(res.json().data.pageCount).toBe(1);
    expect(Buffer.compare(await store.readOriginal(id, 'xlsx'), bytes)).toBe(0);
    expect((await store.readMeta(id))?.viewKind).toBe('grid');
    expect(await store.hasDerived(id, 'grid')).toBe(true);
    expect(audits(id)).toContainEqual(expect.objectContaining({ type: 'edit_saved', detail: 'artcraft-app' }));
  });
});
```

- [ ] **Step 3: 시험이 실패하는지 본다**

Run: `npx vitest run tests/artcraftEngine.test.ts tests/artcraftAppRoutes.test.ts`
Expected: FAIL — `engineFor` 5번째 인자 무시(xlsx → unsupported), `contentUrlFor` 없음, `viewerUrl`이 `/view.html`, `artcraftApp` 없음, engine-source xlsx 400, 저장 경로 404.

- [ ] **Step 4: engineFor에 앱 인자** — `src/artcraftEngine.ts`의 `engineFor`를 바꾼다(머리 주석에 한 줄 더함).

```ts
/** ext는 meta.ext(저장된 바이트의 실제 종류) — 옛 형식 편집 업로드도 engine-source가 내주는 바이트와 형식이 맞는다.
 * appOk = 아트크래프트 앱 화면이 이 형식을 열 수 있음(src/artcraftApp.ts appKindFor) — 엔진이 안 그리는 엑셀도 앱 화면으로 연다 */
export function engineFor(
  ext: string,
  perms: { readonly canDownload: boolean },
  requested: Engine | undefined,
  installed: boolean,
  appOk = false,
): { engine: Engine; reason?: string } {
  if (requested === 'editor') {
    if (!EDITOR_VIEW_FORMATS.has(ext)) return { engine: 'libreoffice', reason: ENGINE_REASONS.noEditor };
    // 편집기 자료 경로는 열람만(편집·내려받기 둘 다 없음) 토큰을 받지 않는다(edit.ts isViewOnly → 4403)
    return perms.canDownload ? { engine: 'editor' } : { engine: 'libreoffice', reason: ENGINE_REASONS.noDownload };
  }
  if (requested !== 'artcraft') return { engine: 'libreoffice' };
  if (!installed && !appOk) return { engine: 'libreoffice', reason: ENGINE_REASONS.notInstalled };
  if (!(installed && ARTCRAFT_FORMATS.has(ext)) && !appOk) return { engine: 'libreoffice', reason: ENGINE_REASONS.unsupported };
  if (!perms.canDownload) return { engine: 'libreoffice', reason: ENGINE_REASONS.noDownload };
  return { engine: 'artcraft' };
}
```

- [ ] **Step 5: contentUrlFor와 토큰 응답** — `src/routes/documents.ts`.

import 추가: `import { appKindFor, readArtcraftAppManifest } from '../artcraftApp.js';`

`editorUrlFor` 함수 바로 아래에:

```ts
/**
 * 콘텐츠만 보기(ui=content, 설계 2026-11-22-artcraft-app-screen §4) 주소 — 같은 토큰 + ui=content. 편집기 엔진이면 편집기, 쪽 문서는 보기 화면
 * (아트크래프트 토큰이면 엔진 쪽 그림 보기), 엑셀 격자는 편집기(열람만 토큰은 편집기 방에 못 들어가 없음). 한글·도면은 콘텐츠 화면이 없다.
 */
export function contentUrlFor(
  ext: string,
  viewKind: ViewKind | undefined,
  engine: Engine,
  perms: { readonly canEdit: boolean; readonly canDownload: boolean },
  token: string,
): string | undefined {
  const query = `?token=${encodeURIComponent(token)}&ui=content`;
  if (engine === 'editor') return `${editorPathFor(ext)}${query}`;
  if (viewPathFor(ext, viewKind) === '/view.html') return `/view.html${query}`;
  if (EDITOR_VIEW_FORMATS.has(ext) && (perms.canEdit || perms.canDownload)) return `${editorPathFor(ext)}${query}`;
  return undefined;
}
```

토큰 라우트에서 `const decision = engineFor(meta.ext, { canDownload }, engine, readArtcraftManifest() !== null);`를 바꾼다:

```ts
      const engineManifest = readArtcraftManifest();
      const appKind = appKindFor(meta.ext, readArtcraftAppManifest(), engineManifest);
      const decision = engineFor(meta.ext, { canDownload }, engine, engineManifest !== null, appKind !== null);
```

`const viewerUrl = …` 줄 아래에:

```ts
      // 아트크래프트 앱 화면이 되는 형식이면 그 화면이 기본 보기(설계 2026-11-22-artcraft-app-screen §3) — 앱이 못 열면 그 화면이 엔진 보기·서버 변환으로 넘긴다
      const appUrl = decision.engine === 'artcraft' && appKind ? `/artcraft-app.html?token=${encodeURIComponent(token)}` : undefined;
      const contentUrl = contentUrlFor(meta.ext, meta.viewKind, decision.engine, { canEdit: effectiveCanEdit, canDownload }, token);
```

응답 `data`의 `viewerUrl:` 줄과 그 다음을 바꾼다:

```ts
          viewerUrl: decision.engine === 'editor' ? `${editorPathFor(meta.ext)}?token=${encodeURIComponent(token)}&ui=view` : (appUrl ?? viewerUrl),
          ...(decision.engine === 'editor' ? { fallbackUrl: viewerUrl } : {}),
          ...(contentUrl ? { contentUrl } : {}),
```

- [ ] **Step 6: info·engine-source·저장 경로** — `src/routes/viewer.ts`.

import 추가/변경:

```ts
import { appInfo, appKindFor, readArtcraftAppManifest } from '../artcraftApp.js';
import { runSerialized } from '../collab/saveSerializer.js';
```

`'../collab/room.js'` import 목록에 `discardPrewarmedXlsx, isRoomOpen`을 더하고, `'../legacyOffice.js'` import에 `ooxmlKind`를 더한다.

info 핸들러의 엔진 판정 세 줄을 바꾼다:

```ts
      const manifest = readArtcraftManifest();
      const appManifest = readArtcraftAppManifest();
      const appKind = appKindFor(view.meta.ext, appManifest, manifest);
      const engine = engineFor(view.meta.ext, view.claims.perms, view.claims.engine, manifest !== null, appKind !== null);
      const artcraft = engine.engine === 'artcraft' && manifest ? artcraftInfo(view.meta.ext, request.query.token!, manifest) : undefined;
      // 앱 화면(public/artcraft-app.html)이 쓴다 — viewerUrl은 앱이 못 열 때 넘어갈 보기 화면 그대로
      const artcraftApp = engine.engine === 'artcraft' && appManifest ? appInfo(view.meta.ext, request.query.token!, appManifest, manifest) : undefined;
```

응답의 `...(artcraft ? { artcraft: … } : {}),` 줄 아래에 `...(artcraftApp ? { artcraftApp } : {}),`를 더한다.

engine-source의 형식 검사를 바꾼다:

```ts
      // 엔진이 그리는 형식, 또는 앱 화면이 여는 형식(엑셀)
      if (!ARTCRAFT_FORMATS.has(meta.ext) && appKindFor(meta.ext, readArtcraftAppManifest(), readArtcraftManifest()) === null) {
```

파일 끝의 hwp-save 라우트 다음(함수 닫는 `}` 앞)에 저장 경로를 더한다:

```ts
  /**
   * 아트크래프트 앱 화면의 저장(설계 2026-11-22-artcraft-app-screen §3) — 앱이 부모에 준 :save 바이트로 원본을 통째로 바꾼다.
   * hwp-save와 같은 규칙(canEdit·zip 폭탄·edit_saved 감사·보기 자료 다시 만들기) + artcraft 토큰만 + 같은 OOXML 형식만 +
   * 편집 방이 열려 있으면 거절(방의 자동 저장이 이 저장을 덮거나 이 저장이 방의 편집을 덮지 않게). 버전 저장소는 아직 없다.
   */
  app.put<{ Querystring: TokenQuery }>(
    '/api/view/artcraft-save',
    { schema: { querystring: tokenQuerySchema }, bodyLimit: config.maxUploadBytes, onRequest: tokenBeforeBody, config: { rateLimit: HEAVY_RATE_LIMIT } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      const { meta, claims } = view;
      const deny = async (code: number, detail: string, error: string) => {
        await recordAudit({ type: 'edit_save_denied', docId: meta.id, viewer: claims.viewer, detail });
        return reply.code(code).send({ success: false, error });
      };
      if (claims.engine !== 'artcraft' || !claims.perms.canEdit) return deny(403, 'permission', '편집 권한이 없습니다.');
      if (!ARTCRAFT_SAVE_FORMATS.has(meta.ext)) return deny(400, 'unsupported', '이 형식은 아트크래프트 앱 화면에서 저장할 수 없습니다.');
      const data = request.body as Buffer | undefined;
      if (!Buffer.isBuffer(data) || data.length === 0) return reply.code(400).send({ success: false, error: '저장할 문서 내용이 없습니다.' });
      try {
        assertZipNotTooLarge(data, 'artcraft-save');
      } catch (error) {
        if (error instanceof ZipTooLargeError) return reply.code(413).send({ success: false, error: '압축을 풀면 너무 큰 파일입니다.' });
        throw error;
      }
      if (ooxmlKind(data)?.ext !== meta.ext) return deny(400, 'format', '문서 형식이 맞지 않습니다.');
      // 방의 저장(flushToStore)과 같은 줄에 선다 — 검사와 쓰기 사이에 방이 저장하지 않게
      const saved = await runSerialized(meta.id, async () => {
        if (isRoomOpen(meta.id)) return false;
        await store.saveOriginal(meta.id, meta.ext, data);
        return true;
      });
      if (!saved) return deny(409, 'room-open', '편집기 화면에서 이 문서를 열고 있어 저장하지 못했습니다. 편집기를 모두 닫은 뒤 다시 저장하세요.');
      // 보기 화면 자료 다시 만들기 — 실패해도 원본 저장은 된 것(hwp-save와 같다)
      let pageCount = meta.pageCount;
      try {
        if (meta.ext === 'xlsx') {
          discardPrewarmedXlsx(meta.id); // 미리 읽어 둔 옛 내용으로 편집기가 열리지 않게
          const built = await buildViewerGridFile(data);
          await saveGridFile(store, meta.id, built.buffer);
          pageCount = built.sheetCount;
          await store.writeMeta({ ...meta, pageCount, status: 'ready', viewKind: 'grid', progress: 100 });
        } else {
          const converter = findConverter(meta.ext);
          if (converter) {
            await waitForInitialConversion(meta.id); // 첫 변환이 늦게 끝나 새 PDF를 덮지 않게
            pageCount = (await convertToPdf({ ...meta }, converter)).pageCount;
            await store.writeMeta({ ...meta, pageCount, status: 'ready', viewKind: 'pages', progress: 100 });
          }
        }
      } catch (error) {
        request.log.error(error, '아트크래프트 저장 뒤 보기 자료 다시 만들기 실패');
      }
      invalidateRenderCache(meta.id);
      invalidateViewCaches(meta.id);
      await recordAudit({ type: 'edit_saved', docId: meta.id, viewer: claims.viewer, detail: 'artcraft-app' });
      return { success: true, data: { pageCount } };
    },
  );
```

파일 위쪽 상수들 옆(`const ARTCRAFT_WORKER = …` 아래)에:

```ts
/** 아트크래프트 앱 화면이 저장할 수 있는 형식 — 앱이 같은 형식으로 저장한다(.doc 등 옛 형식 그대로인 문서는 저장하지 않는다) */
const ARTCRAFT_SAVE_FORMATS: ReadonlySet<string> = new Set(['docx', 'xlsx', 'pptx']);
```

- [ ] **Step 7: 시험이 통과하는지 본다**

Run: `npx vitest run tests/artcraftEngine.test.ts tests/artcraftAppRoutes.test.ts tests/artcraftRoutes.test.ts tests/readonlyEditorTokens.test.ts`
Expected: PASS.

- [ ] **Step 8: 타입 검사·전체 단위 시험**

Run: `npm run typecheck` → 오류 없음. Run: `npx vitest run` → PASS.

- [ ] **Step 9: 커밋**

```bash
git add tests/artcraftAppRoutes.test.ts
git commit -F - -- tests/artcraftAppRoutes.test.ts tests/artcraftEngine.test.ts src/artcraftEngine.ts src/routes/documents.ts src/routes/viewer.ts <<'EOF'
앱 화면 2: 토큰 viewerUrl(앱 화면)·contentUrl, info artcraftApp, 엑셀 원본, PUT /api/view/artcraft-save(권한·형식·편집 방 열림 409·감사·보기 자료)

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>
EOF
```

---

### Task 3: 앱 화면(제목 줄 24px·iframe·글꼴·열기·저장·읽기 전용·실패 전환)

**Files:**
- Create: `src/client/artcraftAppLogic.ts`, `src/client/artcraftApp.ts`, `src/client/artcraftFrame.ts`, `public/artcraft-app.html`, `public/artcraft-frame.html`, `tests/artcraftAppLogic.test.ts`
- Modify: `src/client/viewStrip.ts` (`isContentUi`·`markContentUi`·`fallBackToViewer`), `tests/viewStrip.test.ts`, `scripts/build-editor.mjs:75-85`, `.gitignore`, `public/js/main.js` (아트크래프트 단추)

**Interfaces:**
- Consumes (Task 2): `GET /api/view/info` → `data.artcraftApp { kind: 'word'|'deck'|'grid'; js; fonts; originalUrl }`, `data.viewerUrl`, `data.permissions.canEdit`, `data.docId`, `data.name`, `data.ext`; `PUT /api/view/artcraft-save`; 토큰 응답 `viewerUrl`
- Consumes (spike 계약): iframe `?host=parent` → `<app>craft:ready`; 부모 → `{ type: '<app>craft:open', name, bytes }`(transfer); 앱 → `:save { name, saveAs, bytes }`, `:dirty { dirty }`, `:failed { error }`; `iframe.contentWindow.<app>craft.execute(id, params)`·`docInspect()`·`uiInspect()`(Promise)
- Produces (`artcraftAppLogic.ts`): `type AppKind`, `READY_TIMEOUT_MS = 20_000`, `OPEN_TIMEOUT_MS = 15_000`, `NOTICES`, `SAVE_LABELS`, `type SaveState`, `isAppJsUrl(kind, js)`, `frameUrl(kind, js, readOnly)`, `readAppMessage(e, origin, frame, kind)`, `readOnlyCommands(kind, password, sheetCount)`, `saveDecision(data, canSave)`, `openName(name, ext)`, `fontName(url)`, `uiShowsDocument(ui, fileName)`, `contentRedirect(kind, token)`, `toBase64(u8)`
- Produces (`viewStrip.ts`): `fallBackToViewer(url: string, text: string = CAPACITY_NOTICE): void`, `isContentUi(search: string): boolean`, `markContentUi(): boolean`(콘텐츠만 보기면 body에 `ui-content`를 붙이고 true — Task 4가 쓴다)
- Produces (DOM, e2e가 읽는다): `body[data-app-state]` = `loading|open|failed`, `body[data-save-state]` = `clean|dirty|saving|saved|failed`, `#doc-title`, `#readonly-badge`, `#save-status`, `#edit-open`, `#app-banner`, `#app-frame`(iframe), 실패 안내 `.view-fallback-notice`

- [ ] **Step 1: 순수 판정 시험을 쓴다** — `tests/artcraftAppLogic.test.ts` 새 파일.

```ts
import { describe, expect, test } from 'vitest';
import {
  NOTICES, contentRedirect, fontName, frameUrl, isAppJsUrl, openName, readAppMessage, readOnlyCommands, saveDecision, toBase64, uiShowsDocument,
} from '../src/client/artcraftAppLogic.js';

const JS = '/artcraft-app/wordcraft-0123456789/wordcraft-web.js';

describe('isAppJsUrl — iframe이 불러도 되는 앱 JS(같은 출처 해시 폴더의 그 앱 것만)', () => {
  test.each([
    ['word', JS, true],
    ['grid', '/artcraft-app/gridcraft-abcdef0123/gridcraft-web.js', true],
    ['deck', JS, false],
    ['word', 'https://evil.example/artcraft-app/wordcraft-0123456789/wordcraft-web.js', false],
    ['word', '//evil.example/artcraft-app/wordcraft-0123456789/wordcraft-web.js', false],
    ['word', '/artcraft-app/wordcraft-0123456789/../x/wordcraft-web.js', false],
    ['word', `${JS}?x=1`, false],
    ['word', 'data:text/javascript,alert(1)', false],
    ['nope', JS, false],
  ] as const)('%s · %s → %s', (kind, js, want) => {
    expect(isAppJsUrl(kind as never, js)).toBe(want);
  });
});

describe('frameUrl', () => {
  test('host=parent, 엑셀 읽기 전용만 chrome=none', () => {
    const word = new URL(frameUrl('word', JS, true), 'http://x');
    expect(word.pathname).toBe('/artcraft-frame.html');
    expect(Object.fromEntries(word.searchParams)).toEqual({ app: 'word', js: JS, host: 'parent' });
    const grid = new URL(frameUrl('grid', '/artcraft-app/gridcraft-0123456789/gridcraft-web.js', true), 'http://x');
    expect(grid.searchParams.get('chrome')).toBe('none');
    expect(new URL(frameUrl('grid', '/artcraft-app/gridcraft-0123456789/gridcraft-web.js', false), 'http://x').searchParams.has('chrome')).toBe(false);
  });
});

describe('readAppMessage — 같은 출처·우리 iframe·<앱>craft: 이름만', () => {
  const frame = {};
  const ok = { origin: 'http://h', source: frame, data: { type: 'wordcraft:ready', version: '0.18.1' } };
  test('맞으면 종류와 자료', () => {
    expect(readAppMessage(ok, 'http://h', frame, 'word')).toEqual({ kind: 'ready', data: ok.data });
    expect(readAppMessage({ ...ok, data: { type: 'wordcraft:save', bytes: new ArrayBuffer(1) } }, 'http://h', frame, 'word')?.kind).toBe('save');
  });
  test('다른 출처·다른 창·다른 앱·모르는 종류·모양 틀림은 null', () => {
    expect(readAppMessage({ ...ok, origin: 'http://evil' }, 'http://h', frame, 'word')).toBeNull();
    expect(readAppMessage({ ...ok, source: {} }, 'http://h', frame, 'word')).toBeNull();
    expect(readAppMessage(ok, 'http://h', frame, 'deck')).toBeNull();
    expect(readAppMessage({ ...ok, data: { type: 'wordcraft:open' } }, 'http://h', frame, 'word')).toBeNull();
    expect(readAppMessage({ ...ok, data: 'wordcraft:ready' }, 'http://h', frame, 'word')).toBeNull();
    expect(readAppMessage({ ...ok, data: null }, 'http://h', frame, 'word')).toBeNull();
  });
});

describe('readOnlyCommands', () => {
  test('워드 review.restrict, 엑셀 시트마다 protectSheet, PPT 없음(띠만)', () => {
    expect(readOnlyCommands('word', 'pw', 0)).toEqual([['review.restrict', { mode: 'readOnly', password: 'pw' }]]);
    expect(readOnlyCommands('grid', 'pw', 3)).toEqual([
      ['review.protectSheet', { sheet: 0, password: 'pw' }],
      ['review.protectSheet', { sheet: 1, password: 'pw' }],
      ['review.protectSheet', { sheet: 2, password: 'pw' }],
    ]);
    expect(readOnlyCommands('grid', 'pw', 0)).toHaveLength(1);
    expect(readOnlyCommands('deck', 'pw', 0)).toEqual([]);
  });
});

describe('saveDecision', () => {
  const bytes = new ArrayBuffer(4);
  test.each([
    [{ bytes, saveAs: false }, true, 'upload'],
    [{ bytes: new Uint8Array(4), saveAs: false }, true, 'upload'],
    [{ bytes, saveAs: false }, false, 'readonly'],
    [{ bytes, saveAs: true }, true, 'saveas'],
    [{ bytes: new ArrayBuffer(0) }, true, 'empty'],
    [{}, true, 'empty'],
  ] as const)('%o · 저장 가능 %s → %s', (data, canSave, want) => {
    expect(saveDecision(data as Record<string, unknown>, canSave)).toBe(want);
  });
});

describe('이름·주소 도우미', () => {
  test('openName: 확장자는 저장된 형식으로', () => {
    expect(openName('보고서.doc', 'docx')).toBe('보고서.docx');
    expect(openName('a.b.xlsx', 'xlsx')).toBe('a.b.xlsx');
    expect(openName('이름없음', 'pptx')).toBe('이름없음.pptx');
  });
  test('fontName: 해시를 뺀 원래 글꼴 파일 이름', () => {
    expect(fontName('/artcraft/fonts/SCDream4-0123456789.otf')).toBe('SCDream4.otf');
  });
  test('uiShowsDocument: 앱 화면 상태에 문서 이름(확장자 뺀)이 보이면 열린 것', () => {
    expect(uiShowsDocument({ title: '한글 보고서 - WordCraft' }, '한글 보고서.docx')).toBe(true);
    expect(uiShowsDocument({ title: 'Untitled' }, '한글 보고서.docx')).toBe(false);
    expect(uiShowsDocument({ t: 'a "q" b' }, 'a "q" b.docx')).toBe(true);
    expect(uiShowsDocument(null, '.docx')).toBe(false);
  });
  test('contentRedirect: 워드·PPT는 엔진 보기, 엑셀은 격자 편집기', () => {
    expect(contentRedirect('word', 'a+b')).toBe('/view.html?token=a%2Bb&ui=content');
    expect(contentRedirect('deck', 't')).toBe('/view.html?token=t&ui=content');
    expect(contentRedirect('grid', 't')).toBe('/editor-grid.html?token=t&ui=content');
  });
  test('toBase64: Buffer와 같은 값(0x8000 넘는 길이도)', () => {
    const u8 = Uint8Array.from({ length: 70_000 }, (_, i) => i % 256);
    expect(toBase64(u8)).toBe(Buffer.from(u8).toString('base64'));
  });
  test('안내 문구는 한국어', () => {
    for (const text of Object.values(NOTICES)) expect(text).toMatch(/[가-힣]/);
  });
});
```

- [ ] **Step 2: 시험이 실패하는지 본다**

Run: `npx vitest run tests/artcraftAppLogic.test.ts`
Expected: FAIL — `../src/client/artcraftAppLogic.js` 없음.

- [ ] **Step 3: 순수 판정 모듈** — `src/client/artcraftAppLogic.ts` 새 파일.

```ts
/**
 * 아트크래프트 앱 화면(public/artcraft-app.html)의 순수 판정 — tests/artcraftAppLogic.test.ts. 앱과의 계약은 시험(SPIKE-WEBAPP §3)의
 * craft-web-host 그대로: ?host=parent → <app>craft:ready, 부모 → <app>craft:open {name, bytes}, 앱 → :save {name, saveAs, bytes}·:dirty·:failed.
 */
export type AppKind = 'word' | 'deck' | 'grid';

/** 앱이 준비(ready)를 알리기까지 — 넘으면 엔진 보기로(설계 §3 실패 전환) */
export const READY_TIMEOUT_MS = 20_000;
/** 문서를 보낸 뒤 앱 화면에 문서 이름이 보이기까지 — 읽기 전용은 확인 못 하면 보호를 못 걸므로 전환 */
export const OPEN_TIMEOUT_MS = 15_000;

export const NOTICES = {
  failed: '아트크래프트 앱 화면을 열지 못해 다른 화면으로 보여 줍니다',
  readOnlyDeck: '이 화면의 편집은 저장되지 않습니다',
  readOnlyIgnored: '읽기 전용이라 저장하지 않았습니다',
  saveAs: '다른 이름으로 저장·내보내기는 지원하지 않습니다 — Ctrl+S로 저장하세요',
  appError: '아트크래프트 앱에 오류가 났습니다. 저장한 뒤 다시 열어 주세요',
} as const;

export const SAVE_LABELS = { clean: '', dirty: '저장 안 됨', saving: '저장하는 중…', saved: '저장됨', failed: '저장하지 못했습니다' } as const;
export type SaveState = keyof typeof SAVE_LABELS;

const APP_JS = /^\/artcraft-app\/(word|deck|grid)craft-[0-9a-f]{10}\/(word|deck|grid)craft-web\.js$/;

/** iframe이 불러도 되는 앱 JS — 같은 출처, 해시 폴더, 그 앱 것만(주소를 꾸며 다른 스크립트를 이 출처에서 돌리지 못하게) */
export function isAppJsUrl(kind: AppKind, js: string): boolean {
  const m = APP_JS.exec(js);
  return m !== null && m[1] === kind && m[2] === kind;
}

export function frameUrl(kind: AppKind, js: string, readOnly: boolean): string {
  const q = new URLSearchParams({ app: kind, js, host: 'parent' });
  // 엑셀 읽기 전용은 제목 줄·File 탭을 숨긴다(spike §5)
  if (readOnly && kind === 'grid') q.set('chrome', 'none');
  return `/artcraft-frame.html?${q.toString()}`;
}

export interface AppMessage {
  readonly kind: 'ready' | 'save' | 'dirty' | 'failed';
  readonly data: Record<string, unknown>;
}

/** 앱 → 부모 메시지만 — 같은 출처(origin), 우리 iframe(source), <앱>craft: 이름. 그 밖은 null(무시) */
export function readAppMessage(e: { readonly origin: string; readonly source: unknown; readonly data: unknown }, origin: string, frame: unknown, kind: AppKind): AppMessage | null {
  if (e.origin !== origin || e.source !== frame || typeof e.data !== 'object' || e.data === null) return null;
  const type = (e.data as { type?: unknown }).type;
  const prefix = `${kind}craft:`;
  if (typeof type !== 'string' || !type.startsWith(prefix)) return null;
  const k = type.slice(prefix.length);
  return k === 'ready' || k === 'save' || k === 'dirty' || k === 'failed' ? { kind: k, data: e.data as Record<string, unknown> } : null;
}

/** 읽기 전용으로 열 때 앱에 보낼 명령(spike §5 확인) — PPT는 막는 명령이 없어 1차는 띠만(설계 §3) */
export function readOnlyCommands(kind: AppKind, password: string, sheetCount: number): Array<readonly [string, Record<string, unknown>]> {
  if (kind === 'word') return [['review.restrict', { mode: 'readOnly', password }]];
  if (kind === 'grid') return Array.from({ length: Math.max(1, sheetCount) }, (_, sheet) => ['review.protectSheet', { sheet, password }] as const);
  return [];
}

export type SaveDecision = 'upload' | 'readonly' | 'saveas' | 'empty';

export function saveDecision(data: Record<string, unknown>, canSave: boolean): SaveDecision {
  if (!canSave) return 'readonly';
  if (data.saveAs === true) return 'saveas';
  const size = (data.bytes as { byteLength?: unknown } | undefined)?.byteLength;
  return typeof size === 'number' && size > 0 ? 'upload' : 'empty';
}

/** 앱에 줄 파일 이름 — 앱은 확장자로 형식을 고르므로 저장된 바이트 형식(meta.ext)으로 */
export function openName(name: string, ext: string): string {
  return `${name.replace(/\.[^./\\]+$/, '')}.${ext}`;
}

/** /artcraft/fonts/SCDream4-0123456789.otf → SCDream4.otf */
export function fontName(url: string): string {
  return (url.split('/').pop() ?? '').replace(/-[0-9a-f]{10}(\.otf)$/, '$1');
}

/** 앱 화면 상태(uiInspect 결과)에 문서 이름이 보이면 열린 것(시험 run.mjs와 같은 판정) */
export function uiShowsDocument(ui: unknown, fileName: string): boolean {
  const base = fileName.replace(/\.[^./\\]+$/, '');
  return base !== '' && JSON.stringify(ui ?? null).includes(JSON.stringify(base).slice(1, -1));
}

/** ui=content로 앱 화면에 오면 — 워드·PPT는 엔진 쪽 그림 보기, 엑셀은 격자 편집기의 콘텐츠만(설계 §4) */
export function contentRedirect(kind: AppKind, token: string): string {
  const q = `?token=${encodeURIComponent(token)}&ui=content`;
  return kind === 'grid' ? `/editor-grid.html${q}` : `/view.html${q}`;
}

/** font.add 명령은 base64를 받는다(spike host.html과 같다) */
export function toBase64(u8: Uint8Array): string {
  let s = '';
  for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
  return btoa(s);
}
```

- [ ] **Step 4: 시험이 통과하는지 본다**

Run: `npx vitest run tests/artcraftAppLogic.test.ts`
Expected: PASS.

- [ ] **Step 5: `ui=content` 판정 시험을 쓴다**(앱 화면이 `ui=content`로 오면 넘기고, Task 4 화면들도 쓴다) — `tests/viewStrip.test.ts`의 import에 `isContentUi`를 더하고 파일 끝에:

```ts
describe('isContentUi — 콘텐츠만 보기(토큰과 무관하게 주소의 ui=content)', () => {
  test.each([
    ['?token=t&ui=content', true],
    ['?ui=content&token=t', true],
    ['?token=t&ui=view', false],
    ['?token=t', false],
    ['?token=t&ui=CONTENT', false],
    ['?token=t&ui=content2', false],
    ['', false],
  ])('%s → %s', (search, expected) => {
    expect(isContentUi(search)).toBe(expected);
  });
});
```

Run: `npx vitest run tests/viewStrip.test.ts` → FAIL(`isContentUi` 없음).

- [ ] **Step 6: 판정·표시 함수와 실패 안내 문구 인자** — `src/client/viewStrip.ts`의 `isViewUi` 아래에:

```ts
/** 콘텐츠만 보기(ui=content, 설계 2026-11-22-artcraft-app-screen §4) — 토큰 종류와 무관하게 도구를 전부 숨기고 편집하지 않는다(권한은 토큰이 정한다) */
export function isContentUi(search: string): boolean {
  return new URLSearchParams(search).get('ui') === 'content';
}

/** 화면 입구 맨 위에서 — 콘텐츠만 보기면 body에 ui-content(도구를 CSS로 숨김)를 붙인다. 정보를 받기 전에 붙여 도구가 잠깐 보이지 않게 */
export function markContentUi(): boolean {
  const on = isContentUi(location.search);
  if (on) document.body.classList.add('ui-content');
  return on;
}
```

같은 파일의 `fallBackToViewer`를 바꾼다(기본값으로 기존 호출 그대로).

```ts
/** 안내를 띄우고 잠시 뒤 다른 화면으로(뒤로 가기로 다시 막힌 화면에 오지 않게 replace) — 기본 문구는 정원 초과 */
export function fallBackToViewer(url: string, text: string = CAPACITY_NOTICE): void {
  const note = document.createElement('div');
  note.className = 'view-fallback-notice';
  note.setAttribute('role', 'status');
  note.textContent = text;
  document.body.append(note);
  setTimeout(() => location.replace(url), 1500);
}
```

- [ ] **Step 7: iframe 문서** — `src/client/artcraftFrame.ts`와 `public/artcraft-frame.html` 새 파일.

```ts
/**
 * 아트크래프트 앱 iframe 문서(public/artcraft-frame.html) — 우리가 빌드한 같은 출처 앱 JS만 불러 캔버스에 띄운다(spike app-*.html과 같은 꼴).
 * 앱은 이 문서 주소의 ?host=parent(·chrome=none)를 읽는다. 불러오기·시작이 실패하면 부모에 <앱>craft:failed — 부모가 엔진 보기로 넘긴다.
 * sandbox 없음(WebGPU 필요, 설계 §3 보안) — 그래서 주소의 js를 허용 목록으로만 받는다.
 */
import { isAppJsUrl, type AppKind } from './artcraftAppLogic.js';

const q = new URLSearchParams(location.search);
const kind = (q.get('app') ?? '') as AppKind;
const js = q.get('js') ?? '';
const failToParent = (error: string): void => {
  if (window.parent !== window) window.parent.postMessage({ type: `${kind}craft:failed`, error }, location.origin);
};

if (!isAppJsUrl(kind, js)) {
  document.body.textContent = '잘못된 주소입니다.';
  failToParent('허용되지 않은 앱 주소');
} else {
  const canvas = document.createElement('canvas');
  canvas.id = `${kind}craft_canvas`;
  const loading = document.createElement('div');
  loading.id = `${kind}craft_loading`;
  loading.className = 'app-loading';
  loading.textContent = '불러오는 중…';
  document.body.append(canvas, loading);
  import(/* 주소는 위 허용 목록으로 확인 */ js)
    .then((m: { default: () => Promise<unknown> }) => m.default())
    .catch((error: unknown) => failToParent(String(error)));
}
```

```html
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>아트크래프트</title>
  <style>
    html, body { margin: 0; padding: 0; width: 100%; height: 100%; overflow: hidden; background: #e6e6e6; }
    canvas { position: absolute; inset: 0; width: 100%; height: 100%; display: block; }
    .app-loading { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; font: 13px system-ui, sans-serif; pointer-events: none; }
  </style>
</head>
<body>
  <script type="module" src="/js/artcraftFrame.bundle.js"></script>
</body>
</html>
```

- [ ] **Step 8: 앱 화면 HTML** — `public/artcraft-app.html` 새 파일. 제목 줄은 view.html과 같은 마크업·클래스(`#toolbar.ribbon-skin > .toolbar-row-title`) — 높이 24px은 style.css의 리본 꾸밈 규칙이 정한다.

```html
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>DocView — 아트크래프트</title>
  <link rel="stylesheet" href="/css/style.css" />
  <link rel="stylesheet" href="/css/view.css" />
  <style>
    /* 앱 화면(설계 2026-11-22-artcraft-app-screen §3) — 우리 띠는 제목 줄 하나(24px, 리본 꾸밈 규격), 나머지는 전부 앱(iframe) */
    html, body.artcraft-app-page { height: 100%; margin: 0; overflow: hidden; }
    .artcraft-app-page #toolbar { position: fixed; top: 0; left: 0; right: 0; z-index: 10; }
    .artcraft-app-page .save-status { margin-left: auto; font-size: 11px; color: var(--skin-text-dim); white-space: nowrap; }
    .artcraft-app-page .app-banner { position: fixed; top: calc(var(--skin-title-h) + 1px); left: 0; right: 0; z-index: 9; box-sizing: border-box; height: 18px; padding: 0 8px; font: 11px/18px var(--skin-font); background: #fff4ce; color: #4d3a00; }
    .artcraft-app-page .app-banner[hidden] { display: none; }
    #app-host { position: fixed; top: calc(var(--skin-title-h) + 1px); left: 0; right: 0; bottom: 0; }
    #app-host.with-banner { top: calc(var(--skin-title-h) + 19px); }
    #app-frame { display: block; width: 100%; height: 100%; border: 0; }
    .artcraft-app-page .view-notice { top: calc(var(--skin-title-h) + 0.3rem); }
  </style>
</head>
<body class="viewer-body artcraft-app-page" data-app-state="loading" data-save-state="clean">
  <header class="viewer-toolbar ribbon-skin" id="toolbar">
    <div class="toolbar-row toolbar-row-title">
      <div class="toolbar-group toolbar-left">
        <a href="/" class="icon-btn doclist-link" title="문서 목록으로 돌아가기" aria-label="문서 목록으로 돌아가기">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 12H5"/><path d="m12 19-7-7 7-7"/></svg>
        </a>
        <span id="doc-title" class="doc-title">문서 불러오는 중…</span>
        <span id="readonly-badge" class="engine-badge" hidden title="이 화면에서 고친 내용은 서버에 저장되지 않습니다">읽기 전용</span>
      </div>
      <span id="save-status" class="save-status" role="status" aria-live="polite"></span>
      <button id="edit-open" class="btn-pill view-open-edit" type="button" hidden title="편집할 수 있는 앱 화면으로 다시 엽니다">편집 열기</button>
    </div>
  </header>
  <div id="app-banner" class="app-banner" role="status" hidden></div>
  <div id="app-notice" class="view-notice" role="status" hidden></div>
  <div id="app-host"></div>
  <script type="module" src="/js/artcraftApp.bundle.js"></script>
</body>
</html>
```

- [ ] **Step 9: 앱 화면 DOM** — `src/client/artcraftApp.ts` 새 파일.

```ts
/**
 * 아트크래프트 앱 화면 입구(public/artcraft-app.html, 설계 2026-11-22-artcraft-app-screen §3). 앱 자체 화면(리본·도구 그대로)을 같은 출처 iframe에
 * 띄우고, 준비되면 한글 글꼴(font.add ×2) → 문서(:open)를 넣는다. 읽기 전용이면 워드 review.restrict·엑셀 protectSheet, PPT는 띠.
 * 저장(:save)은 canEdit 토큰이면 PUT /api/view/artcraft-save, 아니면 무시하고 안내. 준비 20초 초과·:failed·wasm 받기 실패면 info.viewerUrl
 * (워드·PPT = 엔진 보기 → 거기서도 실패하면 리브레오피스, 엑셀 = 서버 격자)로 넘긴다.
 */
import {
  NOTICES, OPEN_TIMEOUT_MS, READY_TIMEOUT_MS, SAVE_LABELS, contentRedirect, fontName, frameUrl, openName, readAppMessage, readOnlyCommands,
  saveDecision, toBase64, uiShowsDocument, type AppKind, type SaveState,
} from './artcraftAppLogic.js';
import { EDIT_DOC_KEY, fallBackToViewer, isContentUi } from './viewStrip.js';

interface AppInfo {
  readonly docId: string;
  readonly name: string;
  readonly ext: string;
  readonly viewerUrl: string;
  readonly permissions: { readonly canEdit: boolean };
  readonly artcraftApp?: { readonly kind: AppKind; readonly js: string; readonly fonts: readonly string[]; readonly originalUrl: string };
}
interface AppApi {
  execute(id: string, params?: Record<string, unknown>): Promise<unknown>;
  docInspect(): Promise<unknown>;
  uiInspect(): Promise<unknown>;
}

const token = new URLSearchParams(location.search).get('token') ?? '';
const el = (id: string) => document.getElementById(id) as HTMLElement;
const sleep = (ms: number) => new Promise((r) => setTimeout(r, ms));

function setSave(state: SaveState, detail = ''): void {
  document.body.dataset.saveState = state;
  el('save-status').textContent = detail ? `${SAVE_LABELS[state]} — ${detail}` : SAVE_LABELS[state];
}

let noticeTimer: ReturnType<typeof setTimeout> | null = null;
function notice(text: string): void {
  const box = el('app-notice');
  box.textContent = text;
  box.hidden = false;
  if (noticeTimer !== null) clearTimeout(noticeTimer);
  noticeTimer = setTimeout(() => (box.hidden = true), 8000);
}

function showBanner(text: string): void {
  const banner = el('app-banner');
  banner.textContent = text;
  banner.hidden = false;
  el('app-host').classList.add('with-banner');
}

function randomPassword(): string {
  return [...crypto.getRandomValues(new Uint8Array(16))].map((b) => b.toString(16).padStart(2, '0')).join('');
}

function readEditDoc(): string | null {
  try {
    return sessionStorage.getItem(EDIT_DOC_KEY);
  } catch {
    return null;
  }
}

/** 메인 화면에서 연 업로더 본인만 — 같은 문서의 편집 토큰(artcraft)을 받아 편집 가능한 앱 화면으로(officeShell.openEditFromView와 같은 규칙) */
async function openEdit(docId: string): Promise<void> {
  const res = await fetch(`/api/documents/${encodeURIComponent(docId)}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '업로더', canPrint: true, canDownload: true, canCopy: true, canEdit: true, engine: 'artcraft' }),
  }).catch(() => null);
  const json = (await res?.json().catch(() => null)) as { success?: boolean; data?: { viewerUrl?: string }; error?: string } | null;
  if (!res?.ok || !json?.success || !json.data?.viewerUrl) {
    notice(json?.error ?? '편집 화면을 열지 못했습니다. 잠시 뒤 다시 눌러 주세요.');
    return;
  }
  location.href = json.data.viewerUrl;
}

async function boot(): Promise<void> {
  if (!token) {
    el('doc-title').textContent = '열람 토큰이 없습니다. 발급받은 링크로 접속해 주세요.';
    return;
  }
  const res = await fetch(`/api/view/info?token=${encodeURIComponent(token)}`).catch(() => null);
  const json = (await res?.json().catch(() => null)) as { success?: boolean; data?: AppInfo; error?: string } | null;
  if (!res?.ok || !json?.success || !json.data) {
    el('doc-title').textContent = json?.error ?? '문서를 열 수 없습니다.';
    return;
  }
  const info = json.data;
  const app = info.artcraftApp;
  // 앱 화면을 못 쓰는 토큰(설치 안 됨·내려받기 권한 없음 등) — 보기 화면이 이유를 보인다
  if (!app) {
    location.replace(info.viewerUrl);
    return;
  }
  if (isContentUi(location.search)) {
    location.replace(contentRedirect(app.kind, token));
    return;
  }
  el('doc-title').textContent = info.name;
  document.title = `DocView — ${info.name}`;
  const canSave = info.permissions.canEdit && ['docx', 'xlsx', 'pptx'].includes(info.ext);
  const readOnly = !canSave;
  el('readonly-badge').hidden = !readOnly;
  if (readOnly && app.kind === 'deck') showBanner(NOTICES.readOnlyDeck);
  if (readOnly && readEditDoc() === info.docId) {
    const edit = el('edit-open') as HTMLButtonElement;
    edit.hidden = false;
    edit.addEventListener('click', () => {
      edit.disabled = true;
      void openEdit(info.docId).finally(() => (edit.disabled = false));
    });
  }

  let opened = false;
  let failed = false;
  let unsaved = false;
  const fail = (why: string): void => {
    if (failed || opened) return;
    failed = true;
    clearTimeout(readyTimer);
    console.warn('아트크래프트 앱 화면을 열지 못했습니다:', why);
    document.body.dataset.appState = 'failed';
    fallBackToViewer(info.viewerUrl, NOTICES.failed);
  };
  const readyTimer = setTimeout(() => fail('준비 20초 초과'), READY_TIMEOUT_MS);

  // 앱을 받는 동안 원본·글꼴도 함께 받는다
  const docBytes = fetch(app.originalUrl).then((r) => {
    if (!r.ok) throw new Error(`원본 ${r.status}`);
    return r.arrayBuffer();
  });
  const fontBytes = Promise.all(app.fonts.map(async (url) => [fontName(url), new Uint8Array(await (await fetch(url)).arrayBuffer())] as const));
  docBytes.catch(() => undefined);
  fontBytes.catch(() => undefined);

  const frame = document.createElement('iframe');
  frame.id = 'app-frame';
  frame.title = '아트크래프트 문서 화면';
  frame.src = frameUrl(app.kind, app.js, readOnly);
  el('app-host').append(frame);
  const fileName = openName(info.name, info.ext);

  async function openDoc(): Promise<void> {
    try {
      const api = (frame.contentWindow as unknown as Record<string, AppApi>)[`${app!.kind}craft`];
      for (const [name, u8] of await fontBytes) await api.execute('font.add', { data: toBase64(u8), name });
      const bytes = await docBytes;
      frame.contentWindow!.postMessage({ type: `${app!.kind}craft:open`, name: fileName, bytes }, location.origin, [bytes]);
      let shown = false;
      for (const t0 = Date.now(); !shown && Date.now() - t0 < OPEN_TIMEOUT_MS; ) {
        await sleep(100);
        shown = uiShowsDocument(await api.uiInspect().catch(() => null), fileName);
      }
      if (!shown && readOnly) throw new Error('문서 열기 확인 시간 초과(읽기 전용 보호를 걸 수 없음)');
      if (readOnly) {
        const inspect = app!.kind === 'grid' ? ((await api.docInspect()) as { sheets?: unknown[] } | null) : null;
        const sheets = Array.isArray(inspect?.sheets) ? inspect.sheets.length : 1;
        const password = randomPassword();
        for (const [id, params] of readOnlyCommands(app!.kind, password, sheets)) await api.execute(id, params);
      }
      opened = true;
      document.body.dataset.appState = 'open';
    } catch (error) {
      fail(error instanceof Error ? error.message : String(error));
    }
  }

  async function upload(bytes: ArrayBuffer | Uint8Array): Promise<void> {
    setSave('saving');
    const r = await fetch(`/api/view/artcraft-save?token=${encodeURIComponent(token)}`, {
      method: 'PUT',
      headers: { 'content-type': 'application/octet-stream' },
      body: bytes,
    }).catch(() => null);
    const j = (await r?.json().catch(() => null)) as { success?: boolean; error?: string } | null;
    if (r?.ok && j?.success) {
      unsaved = false;
      setSave('saved');
    } else {
      unsaved = true;
      setSave('failed', j?.error ?? '서버에 연결할 수 없습니다');
    }
  }

  window.addEventListener('message', (e) => {
    const msg = readAppMessage(e, location.origin, frame.contentWindow, app.kind);
    if (!msg) return;
    if (msg.kind === 'ready') {
      if (failed || opened) return;
      clearTimeout(readyTimer);
      void openDoc();
    } else if (msg.kind === 'failed') {
      // 열기 전이면 다른 화면으로, 연 뒤면(고친 내용이 있을 수 있다) 안내만
      if (opened) notice(NOTICES.appError);
      else fail(String(msg.data.error ?? ''));
    } else if (msg.kind === 'dirty') {
      if (canSave && msg.data.dirty === true) {
        unsaved = true;
        setSave('dirty');
      }
    } else {
      const decision = saveDecision(msg.data, canSave);
      if (decision === 'readonly') notice(NOTICES.readOnlyIgnored);
      else if (decision === 'saveas') notice(NOTICES.saveAs);
      else if (decision === 'empty') setSave('failed', '빈 문서');
      else void upload(msg.data.bytes as ArrayBuffer);
    }
  });
  window.addEventListener('beforeunload', (event) => {
    if (canSave && unsaved) event.preventDefault();
  });
}

boot().catch((error) => {
  console.error('앱 화면을 열지 못했습니다:', error);
  el('doc-title').textContent = '문서를 열 수 없습니다.';
});
```

- [ ] **Step 10: 번들 진입점·무시 목록** — `scripts/build-editor.mjs`의 `entryPoints`에서 `'src/client/view/artcraftWorker.ts',` 아래에:

```js
    // 아트크래프트 앱 화면(public/artcraft-app.html)과 그 iframe 문서(public/artcraft-frame.html)
    'src/client/artcraftApp.ts',
    'src/client/artcraftFrame.ts',
```

`.gitignore`의 `public/js/artcraftWorker.bundle.js.map` 아래에:

```
public/js/artcraftApp.bundle.js
public/js/artcraftApp.bundle.js.map
public/js/artcraftFrame.bundle.js
public/js/artcraftFrame.bundle.js.map
```

- [ ] **Step 11: 메인 화면 아트크래프트 단추** — `public/js/main.js`.

`function supportsArtcraft(doc) { … }` 아래에:

```js
/** "아트크래프트로 열기" 단추 — 엔진 형식 + 앱 화면이 여는 엑셀(앱이 빌드돼 있지 않으면 서버가 libreoffice + 이유로 돌린다).
 * 나란히 비교(compare.html)는 엔진 형식만이라 supportsArtcraft를 그대로 쓴다 */
function supportsArtcraftOpen(doc) {
  return supportsArtcraft(doc) || ['xlsx', 'xls'].includes(doc.legacy?.ext ?? doc.ext);
}
```

`const artOk = supportsArtcraft(doc);` → `const artOk = supportsArtcraftOpen(doc);`

`artView.addEventListener('click', () => {` 안 첫 줄 `artView.disabled = true;` 아래에:

```js
      // 앱 화면(읽기 전용)의 "편집 열기"가 같은 문서인지 본다 — 기본 "열기"와 같은 sessionStorage 키
      try {
        sessionStorage.setItem(EDIT_DOC_KEY, doc.id);
      } catch {
        /* 저장소를 못 쓰면 "편집 열기" 단추만 안 보인다 */
      }
```

- [ ] **Step 12: 번들·타입 검사·단위 시험**

Run: `npm run build:editor` → `public/js/artcraftApp.bundle.js`, `public/js/artcraftFrame.bundle.js` 생김.
Run: `npm run typecheck` → 오류 없음. Run: `npx vitest run` → PASS.
Run: `git status --porcelain -- public/js` → 새 번들이 보이지 않음(무시 목록).

- [ ] **Step 13: 커밋**

```bash
git add src/client/artcraftAppLogic.ts src/client/artcraftApp.ts src/client/artcraftFrame.ts public/artcraft-app.html public/artcraft-frame.html tests/artcraftAppLogic.test.ts
git commit -F - -- src/client/artcraftAppLogic.ts src/client/artcraftApp.ts src/client/artcraftFrame.ts public/artcraft-app.html public/artcraft-frame.html tests/artcraftAppLogic.test.ts src/client/viewStrip.ts tests/viewStrip.test.ts scripts/build-editor.mjs .gitignore public/js/main.js <<'EOF'
앱 화면 3: 아트크래프트 앱 화면(같은 출처 iframe·origin 검사·글꼴·열기·저장·읽기 전용·20초/실패 전환, 24px 한글 제목 줄), 메인 단추 엑셀까지

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>
EOF
```

---

### Task 4: 콘텐츠만 보기(`ui=content`) — 세 편집기·보기 화면·메인 링크

**Files:**
- Modify: `src/client/officeShell.ts:12,~322`, `src/client/editor.ts:46,602,637-639`, `src/client/grid-editor.ts:28,289,305-307`, `src/client/editor-slides.ts:11,1314,1360-1362`, `src/routes/edit.ts:9-16,101`, `src/client/view.ts`, `public/css/view.css`(끝), `public/css/style.css`(끝), `public/js/main.js`

**Interfaces:**
- Consumes (Task 2): 토큰 응답 `contentUrl`; (Task 3) `isContentUi(search)`, `markContentUi()` (`src/client/viewStrip.ts`)
- Produces (HTTP): `GET /ws/edit/:docId?token=…&readonly=1` → 편집 토큰이어도 읽기 전용 연결(`edit_joined_readonly`)
- Produces (DOM): `body.ui-content`; 메인 화면 `.doc-info > a.btn` "콘텐츠만 보기"

- [ ] **Step 1: 편집 방에 스스로 읽기 전용으로** — `src/routes/edit.ts`.

```ts
interface EditQuery {
  readonly token?: string;
  /** '1'이면 편집 토큰이어도 읽기 전용으로 들어온다(콘텐츠만 보기 ui=content) — 권한을 좁히기만 한다 */
  readonly readonly?: string;
}

const editQuerySchema = {
  type: 'object',
  properties: { token: { type: 'string' }, readonly: { type: 'string' } },
} as const;
```

`const readOnly = !claims.perms.canEdit;` →

```ts
      const readOnly = !claims.perms.canEdit || request.query.readonly === '1';
```

- [ ] **Step 2: 세 편집기** — 각 파일에서 같은 세 군데를 바꾼다.

`src/client/editor.ts`: import 줄 `import { handleSessionRejected } from './viewStrip.js';` → `import { handleSessionRejected, markContentUi } from './viewStrip.js';`, 그 import 묶음 바로 아래(모듈 맨 위 코드)에

```ts
/** 콘텐츠만 보기(ui=content) — 도구를 숨기고 편집 토큰이어도 읽기 전용으로 연다(설계 2026-11-22-artcraft-app-screen §4) */
const contentUi = markContentUi();
```

`canEdit = info.permissions.canEdit;` → `canEdit = info.permissions.canEdit && !contentUi;`
`params: { token },` → `params: { token: token ?? '', ...(contentUi ? { readonly: '1' } : {}) },`

`src/client/grid-editor.ts`(import `:28`, canEdit `:289`, params `:306`)와 `src/client/editor-slides.ts`(import `:11`, canEdit `:1314`, params `:1361`)도 똑같이 바꾼다(같은 세 줄·같은 주석).

- [ ] **Step 3: 편집기 틀에 보기 규칙 막기** — `src/client/officeShell.ts`.

import 줄을 `import { EDIT_DOC_KEY, installViewGuards, isContentUi, isViewUi, markReadOnly, viewControlAllowed } from './viewStrip.js';`로 바꾸고, `// ---------- 보기 꾸밈(ui=view …` 블록이 끝나는 `}` 바로 아래에:

```ts
  // ---------- 콘텐츠만 보기(ui=content, 설계 2026-11-22-artcraft-app-screen §4) ----------
  // 도구는 CSS(body.ui-content)가 숨기고, 편집기는 이 화면에서 canEdit을 끄고 방에 readonly=1로 들어온다(편집기 입구). 여기서는 보기 규칙의 키·붙여넣기·우클릭 막기
  if (isContentUi(location.search)) {
    document.body.classList.toggle('no-copy', !options.canCopy);
    installViewGuards(options.canCopy, options.canPrint);
  }
```

- [ ] **Step 4: 편집기 숨김 CSS** — `public/css/style.css` **파일 끝에** 덧붙인다(리본 꾸밈 묶음 뒤, 앞 규칙은 건드리지 않는다).

```css

/* ==================== 콘텐츠만 보기(ui=content, 설계 2026-11-22-artcraft-app-screen §4) ====================
   세 편집기에서 도구(제목 줄·탭 줄·리본·상태 줄·섬네일 레일·수식 줄·시트 추가·안내 띠·파일 화면)를 숨기고 본문만 전체 폭.
   엑셀 시트 탭은 내용 이동이라 남긴다. 숨김은 화면 꾸밈일 뿐 — 권한은 토큰이 정한다 */
body.ui-content :is(#toolbar, #statusbar, .formula-bar, .grid-status-row, #slides-rail, #sheet-add, .unsupported-warning, .backstage) { display: none !important; }
body.ui-content.no-copy :is(#editor, #grid-editor, #slide-canvas-wrap),
body.ui-content.no-copy :is(#editor, #grid-editor, #slide-canvas-wrap) * { user-select: none; -webkit-user-select: none; }
```

- [ ] **Step 5: 보기 화면(view.html)** — `src/client/view.ts` import에 `import { markContentUi } from './viewStrip.js';`를 더하고 `const token = …` 줄 바로 위에:

```ts
// 콘텐츠만 보기(ui=content) — 도구 막대를 숨기고 쪽만(설계 2026-11-22-artcraft-app-screen §4). 정보를 받기 전에 붙인다
markContentUi();
```

`public/css/view.css` 끝에:

```css

/* 콘텐츠만 보기(ui=content, 설계 2026-11-22-artcraft-app-screen §4) — 도구 막대·섬네일·찾기·측정 줄을 숨기고 쪽만(아트크래프트 엔진 보기도 같은 화면) */
body.ui-content { --view-chrome-h: 0px; }
body.ui-content :is(#toolbar, #thumbs-panel, #measure-bar, #search-bar) { display: none !important; }
```

- [ ] **Step 6: 메인 화면 "콘텐츠만 보기"** — `public/js/main.js`.

`openEditor` 함수 아래에:

```js
/** 콘텐츠만 보기(설계 2026-11-22-artcraft-app-screen §4) — 도구 없이 문서만. 편집기 형식은 읽기 전용 편집기 토큰, 그 밖은 보기 토큰의 contentUrl */
async function openContentOnly(doc, viewer) {
  const issued = await issueToken(doc.id, viewer, true, true, true, false, false, supportsEdit(doc) ? 'editor' : 'libreoffice');
  if (!issued.contentUrl) throw new Error('이 형식은 콘텐츠만 보기를 지원하지 않습니다');
  location.href = issued.contentUrl;
}

/** 콘텐츠 화면이 있는 형식 — 편집기 형식, 한글을 뺀 쪽 문서(서버 contentUrlFor와 같은 규칙) */
function supportsContentOnly(doc) {
  return supportsEdit(doc) || ((doc.viewKind ?? 'pages') === 'pages' && doc.ext !== 'hwp' && doc.ext !== 'hwpx');
}
```

`renderDocItem`의 `info.append(makeDeleteButton());`(준비된 문서 줄, `const form = …` 바로 위) 앞에:

```js
  // 콘텐츠만 보기 — 새 화면으로 가는 이동이라 링크(단추 줄 시험이 세는 button 목록은 그대로)
  if (supportsContentOnly(doc)) {
    const contentLink = document.createElement('a');
    contentLink.href = '#';
    contentLink.className = 'btn';
    contentLink.textContent = '콘텐츠만 보기';
    contentLink.title = '도구 없이 문서 내용만 봅니다';
    contentLink.addEventListener('click', (event) => {
      event.preventDefault();
      openContentOnly(doc, '업로더').catch((error) => setStatus(error.message, 'error'));
    });
    info.append(contentLink);
  }
```

- [ ] **Step 7: 번들·타입 검사·단위 시험**

Run: `npm run build:editor` → 성공. Run: `npm run typecheck` → 오류 없음. Run: `npx vitest run` → PASS.

- [ ] **Step 8: 커밋**

```bash
git commit -F - -- src/client/officeShell.ts src/client/editor.ts src/client/grid-editor.ts src/client/editor-slides.ts src/routes/edit.ts src/client/view.ts public/css/view.css public/css/style.css public/js/main.js <<'EOF'
콘텐츠만 보기: ui=content(isContentUi) — 세 편집기·보기 화면에서 도구 전부 숨김, 편집 토큰도 화면·방(readonly=1)은 읽기 전용, 메인 "콘텐츠만 보기" 링크

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>
EOF
```

---

### Task 5: 브라우저 e2e 두 개 + 기존 회귀 + 직접 조작 확인

**Files:**
- Create: `scripts/verify-artcraft-app-e2e.mjs`, `scripts/verify-content-only-e2e.mjs`
- Modify: `package.json` (scripts), `scripts/verify-artcraft-e2e.mjs`(아트크래프트 단추·viewerUrl 단언만), `scripts/verify-readonly-editor-e2e.mjs`(아트크래프트 단추 단언 한 줄)

**Interfaces:**
- Consumes: Task 1 산출물(`public/artcraft-app/manifest.json`), 엔진 산출물(`public/artcraft/manifest.json`, `npm run build:artcraft`), 번들(`npm run build:editor`), Task 3 DOM 계약(`body[data-app-state|data-save-state]`, `#app-frame`, `#readonly-badge`, `#save-status`, `#app-banner`, `.view-fallback-notice`), Task 4 `body.ui-content`·`contentUrl`, `scripts/ribbonSkinLib.mjs`의 `SKIN_PX`·`heightProblems`·`latinOnlyLabels`
- Produces: `npm run verify:artcraft-app`(포트 8172, 8170~8179), `npm run verify:content-only`(포트 8174, 8170~8179)

- [ ] **Step 1: 앱 화면 e2e** — `scripts/verify-artcraft-app-e2e.mjs` 새 파일.

```js
// 아트크래프트 앱 화면 — 설계 2026-11-22-artcraft-app-screen §5의 브라우저 검사 일곱 가지(docx·xlsx·pptx 각각):
//  (1) 앱 화면이 뜨고 문서가 보임(워드 document.text·엑셀 시트 수·PPT 슬라이드 수)  (2) 읽기 전용: 워드·엑셀 타자 → 변화 없음(엑셀 시트 모두 보호), PPT는 띠
//  (3) 편집 토큰 Ctrl+S/file.save → 서버 저장 → 받은 파일에 반영, 편집 방이 열려 있으면 거절 안내  (4) wasm 막기 → 안내 + 엔진 보기(워드)·서버 격자(엑셀)
//  (5) 서버 변환 요청 0건  (6) 두 번째 열기에서 앱 자산을 서버에서 다시 받지 않음  (7) 우리 제목 줄 24px·한글, 600px 가로 넘침 없음
// 이 스크립트가 시험 서버를 직접 띄운다(npm run build:editor·build:artcraft·build:artcraft-app 뒤):
//   node scripts/verify-artcraft-app-e2e.mjs [포트=8172, 8170~8179]
// 합성 문서(audit/office/corpus)만 쓴다. 다 통과하면 .audit-run/artcraft-app-e2e-<포트>/를 지운다. 실패하면 살펴보도록 남긴다.
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { chromium } from 'playwright';
import { SKIN_PX, heightProblems, latinOnlyLabels } from './ribbonSkinLib.mjs';

const PORT = Number(process.argv[2] ?? 8172);
// 운영 포트(8000·8010·8020·8100·8110·8130·8148)·다른 시험 범위와 겹치지 않게
if (!(PORT >= 8170 && PORT <= 8179)) throw new Error(`시험 포트는 8170~8179만 쓴다(받은 값 ${PORT})`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `artcraft-app-e2e-${PORT}`);
const FILES = {
  docx: 'audit/office/corpus/docx/a1-korean-report.docx',
  xlsx: 'audit/office/corpus/xlsx/x07-multisheet-print.xlsx',
  pptx: 'audit/office/corpus/pptx/a2-korean-deck.pptx',
};
const API = { docx: 'wordcraft', xlsx: 'gridcraft', pptx: 'deckcraft' };
const SERVER_CONVERT = /^\/api\/view\/(pages|derived-pdf|status|search|pdf-)/;
const FAILED_NOTICE = '아트크래프트 앱 화면을 열지 못해 다른 화면으로 보여 줍니다';
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

for (const [file, how] of [
  ['public/artcraft-app/manifest.json', 'npm run build:artcraft-app'],
  ['public/artcraft/manifest.json', 'npm run build:artcraft'],
  ['public/js/artcraftApp.bundle.js', 'npm run build:editor'],
]) {
  if (!fs.existsSync(file)) {
    console.error(`${file} 없음 — ${how} 먼저`);
    process.exit(1);
  }
}

let server = null;
async function startServer() {
  fs.rmSync(RUN, { recursive: true, force: true });
  fs.mkdirSync(path.join(RUN, 'tmp'), { recursive: true });
  const log = fs.openSync(path.join(RUN, 'server.log'), 'a');
  server = spawn(process.execPath, ['node_modules/tsx/dist/cli.mjs', 'src/server.ts'], {
    // 방 닫기 유예 0 — (3) 편집기를 닫으면 방이 바로 닫혀 저장이 다시 된다
    env: { ...process.env, DOCVIEW_PORT: String(PORT), DOCVIEW_HOST: '127.0.0.1', DOCVIEW_SECRET: randomBytes(24).toString('hex'), DOCVIEW_DATA_DIR: path.join(RUN, 'data'), TMP: path.join(RUN, 'tmp'), TEMP: path.join(RUN, 'tmp'), DOCVIEW_LO_PROFILE_ROOT: path.join(RUN, 'lo'), DOCVIEW_ROOM_CLOSE_GRACE_MS: '0' },
    stdio: ['ignore', log, log],
  });
  for (let i = 0; i < 120; i += 1) {
    if (await fetch(`${BASE}/api/formats`).then((r) => r.ok, () => false)) return;
    await sleep(500);
  }
  throw new Error('시험 서버가 뜨지 않음');
}
function stopServer() {
  // 내가 띄운 서버만(PID, 자식 포함) — 이름으로 node를 끄지 않는다
  const pid = server?.pid;
  server = null;
  if (pid) {
    try {
      execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' });
    } catch {
      // 이미 끝났다
    }
  }
}
process.on('SIGINT', () => {
  stopServer();
  process.exit(130);
});
process.on('exit', stopServer);
async function appHits() {
  await sleep(500);
  return fs.readFileSync(path.join(RUN, 'server.log'), 'utf8').split('\n').filter((line) => /"url":"\/artcraft-app\//.test(line)).length;
}

async function upload(file) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패 ${file}: ${JSON.stringify(r)}`);
  for (let i = 0; i < 1200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready') return st;
    if (st.status === 'failed') throw new Error(`변환 실패 ${file}: ${st.error}`);
    await sleep(250);
  }
  throw new Error(`대기 시간 초과 ${file}`);
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: 'app-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, engine: 'artcraft', ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);

/** 앱 화면을 열어 앱이 문서를 연 상태(data-app-state=open)까지 — 요청 경로·막은 요청을 모은다 */
async function openApp(context, url, { waitOpen = true } = {}) {
  const page = await context.newPage();
  const paths = [];
  page.on('request', (r) => paths.push(new URL(r.url()).pathname));
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.goto(`${BASE}${url}`);
  if (waitOpen) await page.waitForFunction(() => document.body.dataset.appState === 'open', null, { timeout: 120000 });
  return { page, paths, errors };
}
const frameOf = (page) => page.frames().find((f) => f.url().includes('/artcraft-frame.html'));
const exec = (page, api, id, params = {}) => frameOf(page).evaluate(([a, c, p]) => window[a].execute(c, p), [api, id, params]);
const inspect = (page, api) => frameOf(page).evaluate((a) => window[a].docInspect(), api);
async function clickCanvas(page, fx = 0.45, fy = 0.55) {
  const box = await page.locator('#app-frame').boundingBox();
  await page.mouse.click(box.x + box.width * fx, box.y + box.height * fy);
}
/** 받은 파일(지금 서버 원본) 안 XML 글자 */
async function savedXmlText(tokenStr) {
  const bytes = Buffer.from(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(tokenStr)}`)).arrayBuffer());
  const zip = await JSZip.loadAsync(bytes);
  const parts = await Promise.all(Object.values(zip.files).filter((f) => f.name.endsWith('.xml')).map((f) => f.async('string')));
  return parts.join('\n');
}

let browser = null;
let finished = false;
try {
  await startServer();
  // headless는 WebGPU 어댑터가 없어 앱이 WebGL2(SwiftShader)로 떨어진다 — 그것도 동작(spike §2)
  browser = await chromium.launch({ args: ['--enable-unsafe-webgpu'] });
  const ids = {};
  for (const [ext, file] of Object.entries(FILES)) ids[ext] = (await upload(file)).id;

  // (1)(2)(5)(6)(7) 읽기 전용 토큰
  for (const ext of Object.keys(FILES)) {
    const api = API[ext];
    const tok = await token(ids[ext], {});
    rec(`(1) ${ext}: 토큰 viewerUrl이 앱 화면`, /^\/artcraft-app\.html\?token=/.test(tok.viewerUrl ?? ''), tok.viewerUrl);
    const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const before = await appHits();
    const o = await openApp(context, tok.viewerUrl);
    if (ext === 'docx') {
      const text = (await exec(o.page, api, 'document.text')).text;
      rec('(1) docx: 문서 글자가 앱에 열림', text.includes('한글 보고서 시험 문서'), text.slice(0, 40));
      await clickCanvas(o.page);
      await o.page.keyboard.type('QQQ');
      await sleep(800);
      const after = (await exec(o.page, api, 'document.text')).text;
      rec('(2) docx 읽기 전용: 타자해도 글자 그대로', after === text, after.length - text.length);
    } else if (ext === 'xlsx') {
      const d = await inspect(o.page, api);
      rec('(1) xlsx: 시트 4개', d.sheets.length === 4, d.sheets.length);
      rec('(2) xlsx 읽기 전용: 모든 시트 보호', d.sheets.every((s) => s.protected), JSON.stringify(d.sheets.map((s) => s.protected)));
      const undoBefore = d.undo.length;
      await clickCanvas(o.page, 0.3, 0.4);
      await o.page.keyboard.type('QQQ', { delay: 100 });
      await o.page.keyboard.press('Enter');
      await sleep(800);
      const after = await inspect(o.page, api);
      rec('(2) xlsx 읽기 전용: 타자해도 되돌리기 목록 그대로', after.undo.length === undoBefore, `${undoBefore} → ${after.undo.length}`);
    } else {
      const d = await inspect(o.page, api);
      rec('(1) pptx: 슬라이드 3장', d.slides.length === 3, d.slides.length);
      rec('(2) pptx 읽기 전용: "이 화면의 편집은 저장되지 않습니다" 띠', (await o.page.textContent('#app-banner')) === '이 화면의 편집은 저장되지 않습니다' && (await o.page.isVisible('#app-banner')));
    }
    rec(`(2) ${ext}: 읽기 전용 표지`, await o.page.isVisible('#readonly-badge'));
    const converted = o.paths.filter((p) => SERVER_CONVERT.test(p));
    rec(`(5) ${ext}: 서버 변환 요청 0건`, converted.length === 0, converted.join(', '));
    rec(`(1) ${ext}: 화면 오류 없음`, o.errors.length === 0, o.errors.join(' | '));
    // (7) 제목 줄
    const title = await o.page.$eval('#toolbar .toolbar-row-title', (e) => e.getBoundingClientRect().height);
    const labels = await o.page.evaluate(() => [
      document.querySelector('.doclist-link')?.getAttribute('aria-label') ?? '',
      document.querySelector('#readonly-badge')?.textContent ?? '',
      document.querySelector('#edit-open')?.textContent ?? '',
    ]);
    rec(`(7) ${ext}: 제목 줄 ${SKIN_PX.title}px`, heightProblems({ title }).length === 0, `${title}px`);
    rec(`(7) ${ext}: 띠 글자 한글`, latinOnlyLabels(labels).length === 0 && labels.every((l) => /[가-힣]/.test(l)), labels.join(' / '));
    if (ext === 'docx') {
      await o.page.setViewportSize({ width: 600, height: 900 });
      await sleep(300);
      const m = await o.page.evaluate(() => ({ scroll: document.documentElement.scrollWidth, title: document.querySelector('#toolbar .toolbar-row-title').getBoundingClientRect().height }));
      rec('(7) 600px: 가로 넘침 없음, 제목 줄 한 줄', m.scroll <= 600 && heightProblems({ title: m.title }).length === 0, JSON.stringify(m));
    }
    const first = (await appHits()) - before;
    await o.page.close();
    // (6) 같은 문맥에서 다시 열기 — 앱 자산(immutable)을 서버에서 다시 받지 않는다
    const mid = await appHits();
    const again = await openApp(context, tok.viewerUrl);
    const second = (await appHits()) - mid;
    rec(`(6) ${ext}: 두 번째 열기 앱 자산 요청 0건(첫 열기 ${first}건)`, first > 0 && second === 0, `${first} → ${second}`);
    await again.page.close();
    await context.close();
  }

  // (3) 편집 토큰 저장
  {
    const context = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const tok = await token(ids.docx, { canEdit: true });
    const o = await openApp(context, tok.viewerUrl);
    rec('(3) docx 편집 토큰: 읽기 전용 표지 없음', !(await o.page.isVisible('#readonly-badge')));
    await clickCanvas(o.page);
    await o.page.keyboard.type('ZZQA');
    await o.page.keyboard.press('Control+s');
    await o.page.waitForFunction(() => ['saved', 'failed'].includes(document.body.dataset.saveState), null, { timeout: 60000 });
    rec('(3) docx: Ctrl+S → 저장됨', (await o.page.getAttribute('body', 'data-save-state')) === 'saved', await o.page.textContent('#save-status'));
    rec('(3) docx: 서버 원본에 친 글자', (await savedXmlText(tok.token)).includes('ZZQA'));
    // 편집 방이 열려 있으면 거절 + 안내, 닫으면 다시 저장됨
    const editorTok = await fetch(`${BASE}/api/documents/${ids.docx}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: 'app-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: true }) }).then((r) => r.json()).then((j) => j.data);
    const editor = await context.newPage();
    await editor.goto(`${BASE}${editorTok.editorUrl}`);
    await editor.waitForSelector('.ProseMirror', { timeout: 60000 });
    await exec(o.page, 'wordcraft', 'file.save');
    await o.page.waitForFunction(() => document.body.dataset.saveState === 'failed', null, { timeout: 30000 }).catch(() => {});
    rec('(3) 편집 방 열림 → 저장 거절 + "편집기" 안내', (await o.page.getAttribute('body', 'data-save-state')) === 'failed' && /편집기/.test((await o.page.textContent('#save-status')) ?? ''), await o.page.textContent('#save-status'));
    await editor.close();
    await sleep(1500);
    await exec(o.page, 'wordcraft', 'file.save');
    await o.page.waitForFunction(() => document.body.dataset.saveState === 'saved', null, { timeout: 30000 }).catch(() => {});
    rec('(3) 편집기를 닫으면 다시 저장됨', (await o.page.getAttribute('body', 'data-save-state')) === 'saved');
    await o.page.close();

    const xt = await token(ids.xlsx, { canEdit: true });
    const x = await openApp(context, xt.viewerUrl);
    await clickCanvas(x.page, 0.3, 0.4);
    await x.page.keyboard.press('F2');
    await x.page.keyboard.type('ZZQA', { delay: 100 });
    await x.page.keyboard.press('Enter');
    await exec(x.page, 'gridcraft', 'file.save');
    await x.page.waitForFunction(() => ['saved', 'failed'].includes(document.body.dataset.saveState), null, { timeout: 60000 });
    rec('(3) xlsx: 저장됨 + 서버 원본에 친 글자', (await x.page.getAttribute('body', 'data-save-state')) === 'saved' && (await savedXmlText(xt.token)).includes('ZZQA'), await x.page.textContent('#save-status'));
    await x.page.close();

    const pt = await token(ids.pptx, { canEdit: true });
    const p = await openApp(context, pt.viewerUrl);
    rec('(3) pptx 편집 토큰: 띠 없음', !(await p.page.isVisible('#app-banner')));
    await exec(p.page, 'deckcraft', 'file.save');
    await p.page.waitForFunction(() => ['saved', 'failed'].includes(document.body.dataset.saveState), null, { timeout: 60000 });
    rec('(3) pptx: file.save → 저장됨', (await p.page.getAttribute('body', 'data-save-state')) === 'saved', await p.page.textContent('#save-status'));
    await p.page.close();
    const reopened = await openApp(context, pt.viewerUrl);
    rec('(3) pptx: 저장 뒤 다시 열어도 슬라이드 3장', (await inspect(reopened.page, 'deckcraft')).slides.length === 3);
    await context.close();
  }

  // (4) wasm 막기 → 안내 + 엔진 보기(docx)·서버 격자(xlsx)
  for (const [ext, re] of [['docx', /\/view\.html\?token=/], ['xlsx', /\/viewer\.html\?token=/]]) {
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    await context.route('**/artcraft-app/**/*.wasm', (route) => route.abort());
    const tok = await token(ids[ext], {});
    const o = await openApp(context, tok.viewerUrl, { waitOpen: false });
    const note = await o.page.waitForSelector('.view-fallback-notice', { timeout: 30000 }).then((h) => h.textContent(), () => '');
    await o.page.waitForURL(re, { timeout: 30000 }).catch(() => {});
    rec(`(4) ${ext}: wasm 막기 → 안내`, note === FAILED_NOTICE, note);
    rec(`(4) ${ext}: 다른 화면으로 넘어감`, re.test(o.page.url()), o.page.url());
    if (ext === 'docx') {
      await o.page.waitForSelector('.page-frame[data-page="1"].page-ready', { timeout: 90000 });
      rec('(4) docx: 엔진 보기 표지', await o.page.isVisible('#engine-badge'));
    }
    await context.close();
  }
  finished = true;
} finally {
  await browser?.close().catch(() => undefined);
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
if (finished && !failed) {
  await sleep(500);
  fs.rmSync(RUN, { recursive: true, force: true });
} else {
  console.log(`시험 폴더를 살펴보도록 남김: ${RUN}`);
}
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed ? 1 : 0);
```

- [ ] **Step 2: 콘텐츠만 보기 e2e** — `scripts/verify-content-only-e2e.mjs` 새 파일.

```js
// 콘텐츠만 보기(ui=content) — 설계 2026-11-22-artcraft-app-screen §5: 세 편집기·보기 화면(view.html)·아트크래프트 엔진 보기 각각
//  도구 요소(제목 줄·탭 줄·리본·상태 줄·섬네일·수식 줄·도구 막대) 0개 보임, 본문 보임, 타자·붙여넣기 무효(편집 토큰이어도 — 서버 원본·감사로 확인), 600px 가로 넘침 없음,
//  메인 화면 "콘텐츠만 보기" 링크 → ui=content 주소.
// 이 스크립트가 시험 서버를 직접 띄운다(npm run build:editor·build:artcraft 뒤, LibreOffice 필요):
//   node scripts/verify-content-only-e2e.mjs [포트=8174, 8170~8179]
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { chromium } from 'playwright';

const PORT = Number(process.argv[2] ?? 8174);
if (!(PORT >= 8170 && PORT <= 8179)) throw new Error(`시험 포트는 8170~8179만 쓴다(받은 값 ${PORT})`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `content-only-e2e-${PORT}`);
const FILES = {
  docx: 'audit/office/corpus/docx/a1-korean-report.docx',
  xlsx: 'audit/office/corpus/xlsx/x07-multisheet-print.xlsx',
  pptx: 'audit/office/corpus/pptx/a2-korean-deck.pptx',
};
const BODY = { docx: '.ProseMirror', xlsx: '#grid-edit-table td', pptx: '#slide-canvas' };
const TOOLS = ['#toolbar', '.office-titlebar', '.ribbon-tabs', '.ribbon-panel', '#statusbar', '.formula-bar', '.grid-status-row', '#slides-rail', '.toolbar-row-tools', '#thumbs-panel'];
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

for (const [file, how] of [['public/artcraft/manifest.json', 'npm run build:artcraft'], ['public/js/editor.bundle.js', 'npm run build:editor']]) {
  if (!fs.existsSync(file)) {
    console.error(`${file} 없음 — ${how} 먼저`);
    process.exit(1);
  }
}

let server = null;
async function startServer() {
  fs.rmSync(RUN, { recursive: true, force: true });
  fs.mkdirSync(path.join(RUN, 'tmp'), { recursive: true });
  const log = fs.openSync(path.join(RUN, 'server.log'), 'a');
  server = spawn(process.execPath, ['node_modules/tsx/dist/cli.mjs', 'src/server.ts'], {
    env: { ...process.env, DOCVIEW_PORT: String(PORT), DOCVIEW_HOST: '127.0.0.1', DOCVIEW_SECRET: randomBytes(24).toString('hex'), DOCVIEW_DATA_DIR: path.join(RUN, 'data'), TMP: path.join(RUN, 'tmp'), TEMP: path.join(RUN, 'tmp'), DOCVIEW_LO_PROFILE_ROOT: path.join(RUN, 'lo') },
    stdio: ['ignore', log, log],
  });
  for (let i = 0; i < 120; i += 1) {
    if (await fetch(`${BASE}/api/formats`).then((r) => r.ok, () => false)) return;
    await sleep(500);
  }
  throw new Error('시험 서버가 뜨지 않음');
}
function stopServer() {
  const pid = server?.pid;
  server = null;
  if (pid) {
    try {
      execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' });
    } catch {
      // 이미 끝났다
    }
  }
}
process.on('SIGINT', () => {
  stopServer();
  process.exit(130);
});
process.on('exit', stopServer);

async function upload(file) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패 ${file}: ${JSON.stringify(r)}`);
  for (let i = 0; i < 1200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready') return st;
    if (st.status === 'failed') throw new Error(`변환 실패 ${file}: ${st.error}`);
    await sleep(250);
  }
  throw new Error(`대기 시간 초과 ${file}`);
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: 'content-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);
async function visibleTools(page) {
  const out = [];
  for (const sel of TOOLS) {
    for (const h of await page.locator(sel).all()) if (await h.isVisible()) out.push(sel);
  }
  return out;
}
async function savedXmlText(tokenStr) {
  const bytes = Buffer.from(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(tokenStr)}`)).arrayBuffer());
  const zip = await JSZip.loadAsync(bytes);
  return (await Promise.all(Object.values(zip.files).filter((f) => f.name.endsWith('.xml')).map((f) => f.async('string')))).join('\n');
}
function auditTypes(docId) {
  const dir = path.join(RUN, 'data', 'audit');
  return fs.readdirSync(dir)
    .flatMap((f) => fs.readFileSync(path.join(dir, f), 'utf8').trim().split('\n'))
    .map((l) => JSON.parse(l))
    .filter((e) => e.docId === docId)
    .map((e) => e.type);
}

let browser = null;
let finished = false;
try {
  await startServer();
  browser = await chromium.launch();
  const ids = {};
  for (const [ext, file] of Object.entries(FILES)) ids[ext] = (await upload(file)).id;

  // 세 편집기 — 편집 토큰 + ui=content
  for (const ext of Object.keys(FILES)) {
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] });
    const page = await context.newPage();
    const tok = await token(ids[ext], { canEdit: true });
    await page.goto(`${BASE}${tok.editorUrl}&ui=content`);
    await page.waitForSelector(BODY[ext], { timeout: 90000 });
    await sleep(1500);
    rec(`${ext} 편집기: body.ui-content`, await page.evaluate(() => document.body.classList.contains('ui-content')));
    const tools = await visibleTools(page);
    rec(`${ext} 편집기: 도구 0개 보임`, tools.length === 0, tools.join(', '));
    rec(`${ext} 편집기: 본문 보임`, await page.locator(BODY[ext]).first().isVisible());
    const box = await page.locator(BODY[ext]).first().boundingBox();
    await page.mouse.click(box.x + Math.min(box.width / 2, 200), box.y + Math.min(box.height / 2, 60));
    await page.keyboard.type('QQQ');
    await page.keyboard.press('Delete');
    await page.evaluate(() => navigator.clipboard.writeText('붙여넣기시험'));
    await page.keyboard.press('Control+v');
    await sleep(1500);
    const xml = await savedXmlText(tok.token);
    rec(`${ext} 편집기: 타자·붙여넣기 무효(서버 내용 그대로)`, !xml.includes('QQQ') && !xml.includes('붙여넣기시험'));
    const types = auditTypes(ids[ext]);
    rec(`${ext} 편집기: 방에 읽기 전용으로 들어감(편집 토큰이어도)`, types.includes('edit_joined_readonly') && !types.includes('edit_joined'), types.filter((t) => t.startsWith('edit_')).join(','));
    await page.setViewportSize({ width: 600, height: 900 });
    await sleep(400);
    rec(`${ext} 편집기: 600px 가로 넘침 없음`, (await page.evaluate(() => document.documentElement.scrollWidth)) <= 600);
    await context.close();
  }

  // 보기 화면(리브레오피스) — 토큰 contentUrl
  {
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const page = await context.newPage();
    const tok = await token(ids.docx, {});
    rec('보기: 토큰 contentUrl = view.html + ui=content', /^\/view\.html\?token=[^&]+&ui=content$/.test(tok.contentUrl ?? ''), tok.contentUrl);
    await page.goto(`${BASE}${tok.contentUrl}`);
    await page.waitForSelector('.page-frame[data-page="1"].page-ready', { timeout: 90000 });
    const tools = await visibleTools(page);
    rec('보기: 도구 막대 0개 보임', tools.length === 0, tools.join(', '));
    const top = await page.$eval('.page-frame[data-page="1"]', (e) => e.getBoundingClientRect().top);
    rec('보기: 쪽이 위쪽부터(도구 막대 자리 없음)', top < 60, `${top}px`);
    await page.setViewportSize({ width: 600, height: 900 });
    await sleep(400);
    rec('보기: 600px 가로 넘침 없음', (await page.evaluate(() => document.documentElement.scrollWidth)) <= 600);
    await context.close();
  }

  // 아트크래프트 엔진 보기 — artcraft 토큰 contentUrl(앱 화면을 띄우지 않음)
  {
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const page = await context.newPage();
    const paths = [];
    page.on('request', (r) => paths.push(new URL(r.url()).pathname));
    const tok = await token(ids.docx, { engine: 'artcraft' });
    rec('아트크래프트: contentUrl = view.html + ui=content', /^\/view\.html\?token=[^&]+&ui=content$/.test(tok.contentUrl ?? ''), tok.contentUrl);
    await page.goto(`${BASE}${tok.contentUrl}`);
    await page.waitForSelector('.page-frame[data-page="1"].page-ready', { timeout: 90000 });
    const tools = await visibleTools(page);
    rec('아트크래프트: 도구 0개 보임', tools.length === 0, tools.join(', '));
    rec('아트크래프트: 엔진(wasm)으로 그림, 앱 화면 자산 요청 없음', paths.some((p) => p.startsWith('/artcraft/')) && !paths.some((p) => p.startsWith('/artcraft-app/')), paths.filter((p) => p.startsWith('/artcraft')).join(', '));
    await context.close();
  }

  // 메인 화면 링크
  {
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const page = await context.newPage();
    await page.goto(`${BASE}/`);
    const row = page.locator('li.doc-item', { hasText: path.basename(FILES.docx) });
    await row.waitFor({ timeout: 30000 });
    await row.getByRole('link', { name: '콘텐츠만 보기' }).click();
    await page.waitForURL(/ui=content/, { timeout: 30000 }).catch(() => {});
    rec('메인: "콘텐츠만 보기" → ui=content 주소', /\/editor\.html\?token=[^&]+&ui=content$/.test(page.url()), page.url());
    await context.close();
  }
  finished = true;
} finally {
  await browser?.close().catch(() => undefined);
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
if (finished && !failed) {
  await sleep(500);
  fs.rmSync(RUN, { recursive: true, force: true });
} else {
  console.log(`시험 폴더를 살펴보도록 남김: ${RUN}`);
}
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed ? 1 : 0);
```

- [ ] **Step 3: 명령** — `package.json` scripts의 `"verify:artcraft": …` 줄 아래에:

```json
    "verify:artcraft-app": "node scripts/verify-artcraft-app-e2e.mjs",
    "verify:content-only": "node scripts/verify-content-only-e2e.mjs",
```

- [ ] **Step 4: 새 e2e 실행**

Run: `npm run build:editor && npm run verify:artcraft-app`
Expected: 마지막 줄 `N/N 통과`, exit 0.
Run: `npm run verify:content-only`
Expected: `N/N 통과`, exit 0.

- [ ] **Step 5: 바뀐 단추 동작에 맞춰 기존 두 e2e의 해당 단언만 고친다**(조정자 판정 — 다른 단언·다른 스크립트는 고치지 않는다)

`scripts/verify-artcraft-e2e.mjs`:
  1. 맨 위 준비 검사 목록(`for (const [file, how] of [...]`)에 `['public/artcraft-app/manifest.json', 'npm run build:artcraft-app']`를 더한다(배포 기본 = 앱 화면 켜짐).
  2. `token()` 아래에 엔진 보기 주소 도우미를 더한다 — artcraft 토큰의 기본 보기가 앱 화면으로 바뀌었으므로, 엔진 보기 단언은 앱이 실패할 때 가는 배포 경로(`/view.html?token=…`)에서 그대로 검사한다:
     ```js
     /** 엔진 쪽 그림 보기 — artcraft 토큰의 viewerUrl은 앱 화면(설계 2026-11-22-artcraft-app-screen §3), 엔진 보기는 앱 실패 때 가는 /view.html */
     const engineView = (tok) => `/view.html?token=${encodeURIComponent(tok.token)}`;
     ```
  3. 문서별 반복의 "토큰 engine=artcraft" 단언 바로 아래에 새 동작 단언을 더하고, 그 반복의 `openPage(context, tok.viewerUrl, { throttle: true })`와 (7) 두 번째 열기 `openPage(context, tok.viewerUrl)`를 `openPage(context, engineView(tok), …)`로 바꾼다(쪽 수·잉크·한글·CPU×4·캐시 단언은 그대로):
     ```js
     rec(`${c.name}: 토큰 viewerUrl = 앱 화면`, /^\/artcraft-app\.html\?token=/.test(tok.viewerUrl), tok.viewerUrl);
     ```
  4. (4) wasm 막기 블록과 "열린 뒤 Worker 죽음" 블록의 `openPage(context, tok.viewerUrl, …)`를 `engineView(tok)`로 바꾼다(단언 그대로). (5) 내려받기 권한 없음 블록은 viewerUrl이 지금도 `/view.html`이라 그대로 둔다.
  5. 메인 화면 블록 — 공유 폼: `const shared = await openPage(context, await link.getAttribute('href'));` + `rec('메인: 공유 폼 아트크래프트 링크 → 표지 보임', …)` 두 줄을 아래로 바꾼다:
     ```js
     const href = await link.getAttribute('href');
     rec('메인: 공유 폼 아트크래프트 링크 → 앱 화면 주소', /^\/artcraft-app\.html\?token=/.test(href ?? ''), href ?? '');
     const shared = await context.newPage();
     await shared.goto(`${BASE}${href}`);
     rec('메인: 공유 폼 아트크래프트 링크 → 앱이 문서를 엶', await shared.waitForFunction(() => document.body.dataset.appState === 'open', null, { timeout: 120000 }).then(() => true, () => false));
     await shared.close();
     ```
     (원래의 `await shared.page.close();` 줄은 지운다.) 메인 단추: `await page.waitForURL(/\/view\.html\?token=/, …)`·`page-ready` 대기·`rec('메인: "아트크래프트로 열기" 단추 → 표지 보임', …)` 세 줄을 아래로 바꾼다:
     ```js
     await page.waitForURL(/\/artcraft-app\.html\?token=/, { timeout: 15000 });
     const appOpen = await page.waitForFunction(() => document.body.dataset.appState === 'open', null, { timeout: 120000 }).then(() => true, () => false);
     rec('메인: "아트크래프트로 열기" 단추 → 앱 화면(문서 열림·읽기 전용 표지)', appOpen && (await page.isVisible('#readonly-badge')));
     ```
     `.ppt` 줄 흐림·"리브레오피스로 열기" 단언은 그대로.
  6. 맨 위 주석 (4)·메인 설명에 "엔진 보기 단언은 /view.html?token=…(앱 실패 경로), 단추·viewerUrl은 앱 화면"을 한 줄 더한다.

`scripts/verify-readonly-editor-e2e.mjs`: (6) 한 줄 단추 반복의 `['아트크래프트로 열기', /\/view\.html\?token=/],` 한 줄만 `['아트크래프트로 열기', /\/artcraft-app\.html\?token=/],`로 바꾼다(단추 목록 단언은 "콘텐츠만 보기"가 링크라 그대로). 앱 manifest가 없으면 이 단언이 실패하므로 준비 검사(`public/js/editor.bundle.js` 검사 옆)에 `public/artcraft-app/manifest.json`(없으면 `npm run build:artcraft-app 먼저`)을 더한다.

- [ ] **Step 6: 기존 회귀(배포 기본 동작 그대로)** — 단위·타입부터:

Run: `npx vitest run` → PASS. Run: `npm run typecheck` → 오류 없음.

스스로 서버를 띄우는 세 개(앱 화면 켠 배포 기본 그대로, 환경 변수로 끄지 않는다):

```bash
npm run verify:artcraft
npm run verify:readonly
npm run verify:ribbon-skin
```

Expected: 셋 다 `N/N 통과`(앞 둘은 Step 5에서 고친 단언 포함, ribbon-skin은 스크립트 그대로).

서버를 따로 띄워 BASE를 받는 검사들은 앱 화면을 켠 채 8176에 시험 서버를 띄워 차례로(동시에 돌리지 않는다):

```bash
mkdir -p .audit-run/regress-8176
DOCVIEW_PORT=8176 DOCVIEW_HOST=127.0.0.1 DOCVIEW_SECRET="regress-$RANDOM$RANDOM" DOCVIEW_DATA_DIR="$PWD/.audit-run/regress-8176/data" \
  node node_modules/tsx/dist/cli.mjs src/server.ts > .audit-run/regress-8176/server.log 2>&1 &
echo $! > .audit-run/regress-8176/server.pid
# /api/formats가 200이 될 때까지 기다린 뒤
npm run verify:docx -- http://127.0.0.1:8176
npm run verify:pptx -- http://127.0.0.1:8176
npm run verify:xlsx -- http://127.0.0.1:8176
npm run verify:xlsx:browser -- http://127.0.0.1:8176
npm run verify:view-only -- http://127.0.0.1:8176
npm run verify:view-shell -- http://127.0.0.1:8176
taskkill //PID "$(cat .audit-run/regress-8176/server.pid)" //T //F
```

Expected: 모두 실패 0. 끝나면 `.audit-run/regress-8176`을 지운다. 하나라도 실패하면 이번 변경을 원인으로 의심하고 제품 코드를 고친다(이 여섯 스크립트는 고치지 않는다).

- [ ] **Step 7: 직접 조작 확인(실제 브라우저, headed)** — 8178에 시험 서버(Step 6의 8176 명령에서 포트·폴더만 8178)를 띄우고 Playwright headed(`chromium.launch({ headless: false, args: ['--enable-unsafe-webgpu'] })`)나 PC 크롬으로 연다. 화면 그림만 보고 넘기지 말고 아래를 **모두 손으로 눌러** 확인한다:
  1. 메인 화면에서 docx·xlsx·pptx 각각 "아트크래프트로 열기" → 앱 화면, 읽기 전용 표지, 한글 본문 정상(☒ 없음), 앱 리본 메뉴를 몇 개 눌러 봄, 타자해도 안 바뀜(워드·엑셀), PPT 띠.
  2. 같은 화면의 "편집 열기" → 편집 가능한 앱 화면(표지 없음), 글자 고치고 Ctrl+S → "저장됨", 메인에서 "리브레오피스로 열기"로 열어 고친 내용 보임.
  3. 고친 뒤 저장 전에 창 닫기 → 떠나기 경고. ← "문서 목록으로 돌아가기" 링크 동작.
  4. 편집기("동시편집 열기")를 다른 탭에 열어 둔 채 앱 화면에서 Ctrl+S → "저장하지 못했습니다 — 편집기 화면에서…" 안내.
  5. 메인 "콘텐츠만 보기"(docx·xlsx·pptx·pdf) → 도구가 하나도 없고 문서만, 스크롤·화살표·Ctrl +/−로 이동·확대, 타자·Ctrl+V·우클릭 무효. 엑셀 시트 탭으로 시트 이동.
  6. 창 폭 600px로 줄여 제목 줄 한 줄·가로 스크롤 없음.
  끝나면 8178 서버를 PID로 끈다.

- [ ] **Step 8: 커밋**

```bash
git add scripts/verify-artcraft-app-e2e.mjs scripts/verify-content-only-e2e.mjs
git commit -F - -- scripts/verify-artcraft-app-e2e.mjs scripts/verify-content-only-e2e.mjs package.json scripts/verify-artcraft-e2e.mjs scripts/verify-readonly-editor-e2e.mjs <<'EOF'
앱 화면 5: 브라우저 검사 verify:artcraft-app(7항목×3형식)·verify:content-only, verify:artcraft·readonly의 아트크래프트 단추 단언을 앱 화면으로(엔진 보기 단언은 /view.html로 옮겨 유지)

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>
EOF
```

---

## Self-Review

- **Spec 범위**: §3 대상 형식(APP_FORMATS·.doc :failed 전환 — Task 1·3), 열기 옵션(토큰 viewerUrl·메인 단추 — Task 2·3), 읽기 전용(워드 restrict·엑셀 protectSheet·PPT 띠 — Task 3), 저장(Task 2·3), 우리 띠(Task 3 HTML·e2e (7)), 자산(Task 1), 글꼴(Task 2 `appInfo.fonts`·Task 3 font.add), 실패 전환(Task 3 `fail`·e2e (4)), 보안(Task 1 허용 목록·Task 3 `isAppJsUrl`·`readAppMessage`). §4 ui=content(Task 4, 엔진 보기 Task 2 `contentUrlFor`·Task 4 view.ts). §5 단위(manifest 없음·origin·저장 권한·판정 함수)·e2e 두 개·기존 회귀(Task 5, verify:artcraft·readonly는 단추 단언만 새 동작으로). §6 하지 않는 것(PPT 읽기 전용 패치·메뉴 한글화·동시편집 연동·폰 측정·wasm-opt)은 계획에 없다 — craft-i18n 확인은 2차로 남김.
- **이름 맞춤**: `appKindFor`·`appInfo`·`readArtcraftAppManifest`(Task 1) ↔ Task 2 사용, `contentUrlFor`(Task 2) ↔ Task 2 시험, `isContentUi`·`markContentUi`·`fallBackToViewer(url, text)`(Task 3·4) ↔ Task 3 `artcraftApp.ts`·Task 4 편집기, DOM 계약(`data-app-state`·`data-save-state`·`#app-banner` 등) ↔ Task 5.
