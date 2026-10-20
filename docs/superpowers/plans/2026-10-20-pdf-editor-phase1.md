# PDF 편집기 1단계(기반 + 보기·찾기·레이어) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 향기랩 문서도구에서 PDF를 올리면 pdfcraft와 비슷한 화면의 PDF 편집기(webviewer `editor-pdf.html`)가 열리고, 보기·확대·회전·썸네일·책갈피·레이어·찾기가 되며, 이후 단계(주석·정리·양식…)가 올라설 협업 방(kind `pdf`)이 준비된다.

**Architecture:** 서버는 기존 협업 방(room.ts)에 `pdf` 종류를 더하고(1단계엔 편집 모델이 비어 있어 저장 시 원본 그대로), 편집기용 원본 바이트 경로(`GET /api/view/pdf-source`)를 연다. 토큰 요청의 새 선택 항목 `pdfEditor`가 true이고 내려받기 권한이 있을 때만 pdf에 editorUrl을 준다(8000 유지보수 시스템 영향 차단). 브라우저는 pdf.js로 쪽을 canvas + 글자 층으로 그린다(보이는 쪽만).

**Tech Stack:** TypeScript, Fastify, Yjs/y-websocket, pdfjs-dist 4.2.67(브라우저 ESM 번들 + worker), pdf-lib(서버, 쪽 수), esbuild, vitest, Playwright.

**Spec:** `docs/superpowers/specs/2026-10-20-pdf-editor-design.md`

## Global Constraints

- 화면 문구는 모두 한국어. 배치는 pdfcraft를 따르되 pdfcraft의 아이콘·이미지·문구를 가져오지 않는다(아이콘은 직접 그린 SVG).
- 8000(유지보수 시스템)의 동작이 바뀌면 안 된다: `pdfEditor`가 없거나 false면 pdf의 토큰 응답은 지금과 같다(editorUrl 없음).
- PDF 원본 바이트는 `canDownload` 권한이 있는 토큰에만 준다. PDF 편집기 주소도 `canDownload`가 있을 때만 준다.
- pdf.js 보안: `isEvalSupported: false`, `enableScripting` 쓰지 않음, worker는 자체 호스팅(`/js/pdf.worker.mjs`).
- 기존 편집기·보기 화면이 그대로여야 한다: `npm test`, `npm run verify:xlsx`, `npm run verify:pptx`, 워드 리본 e2e가 고치지 않고 통과.
- 반영 절차: 브랜치 `feature/room-formula-fix-and-ppt-ribbon`(worktree `C:\Users\OpenMinds\projects\webviewer-wt`) 커밋 → `C:\Users\OpenMinds\projects\webviewer`(master)에서 merge → `node scripts/build-editor.mjs` → 서버 코드가 바뀌었으면 8010을 PID 체인으로 재시작 → `git push origin master:main`. calc-site는 `npx wrangler pages deploy public --project-name hyanggilab --branch main` 후 `git push origin master:main`. 커밋 전 diff에 비밀 문자열(`jZd-spcFtb`, `078213b5bc`)이 없는지 확인.
- 시험 서버: `DOCVIEW_PORT=8124 DOCVIEW_HOST=127.0.0.1 DOCVIEW_DATA_DIR="$S/e2e-data3" DOCVIEW_SECRET=e2e-test-secret npx tsx src/server.ts`(S = 세션 scratchpad). 끌 때는 PID로만(node 일괄 종료 금지).

---

## File Structure

| 파일 | 책임 |
|---|---|
| `src/routes/documents.ts` (수정) | 토큰 선택 항목 `pdfEditor`, editorUrl 결정 함수 `editorUrlFor`, pdf 업로드 `editReady` |
| `src/pdf/pdfModel.ts` (새) | PDF 방의 Y.Doc 이름 상수, `exportPdfFromRoom(original, ydoc)`(1단계: 바뀐 것이 없으면 원본 그대로), `pdfChangeLogOf` |
| `src/collab/room.ts` (수정) | DocKind `pdf`, 시딩·저장·현재 바이트·변경 기록 분기 |
| `src/collab/changeLogPersistence.ts` (수정) | docKind에 `pdf` |
| `src/routes/edit.ts` (수정) | 동시편집 허용 확장자에 pdf |
| `src/routes/viewer.ts` (수정) | `GET /api/view/pdf-source` |
| `scripts/build-editor.mjs` (수정) | 진입점 `src/client/editor-pdf.ts`, pdf.js worker·cMap·표준 글꼴 복사 |
| `.gitignore` (수정) | `public/js/editor-pdf.bundle.js*`, `public/js/pdf.worker.mjs`, `public/pdfjs/` |
| `public/editor-pdf.html` (새) | 편집기 틀(제목 줄·위 줄·도구 목록·본문·정보 창·오른쪽 줄·찾기 줄) |
| `public/css/pdf-editor.css` (새) | 편집기 전용 스타일(pdfcraft 배치) |
| `src/client/pdf/pdfDocument.ts` (새) | pdf.js 열기·쪽 크기·책갈피·레이어 읽기(화면과 분리된 순수 래퍼) |
| `src/client/pdf/pageView.ts` (새) | 쪽 목록 그리기: 보이는 쪽만 canvas+글자 층, 배율·회전, 현재 쪽 추적 |
| `src/client/pdf/findController.ts` (새) | 찾기: 쪽 글자 모으기, 일치 목록, 강조·이동 |
| `src/client/pdf/panels.ts` (새) | 정보 창: 썸네일·책갈피·레이어 |
| `src/client/editor-pdf.ts` (새) | 시작: 토큰·정보·Yjs 연결·참여자·← 돌아가기·docview:ready, 위 모듈 연결, 단축키 |
| `src/client/pdf/findText.ts` (새) | 찾기의 순수 함수(정규화·일치 위치) — 단위 테스트 대상 |
| `tests/pdfEditorRouting.test.ts` (새) | editorUrlFor·EDITABLE 결정 |
| `tests/pdfModel.test.ts` (새) | exportPdfFromRoom 1단계 동작 |
| `tests/pdfFindText.test.ts` (새) | 찾기 순수 함수 |
| `scripts/verify-pdf-editor-e2e.mjs` (새) + `package.json` script `verify:pdf` | 브라우저 확인 |
| `C:\Users\OpenMinds\calc-site\public\tools\index.html` (수정) | `.pdf` 받기, 토큰에 `pdfEditor: true` |

---

### Task 1: 토큰 옵션 `pdfEditor`와 editorUrl 결정

**Files:**
- Modify: `src/routes/documents.ts` (tokenBodySchema :22, TokenBody, `/tokens` 핸들러 :399-452, 업로드 meta :242-256, EDITABLE_EXTS/editorPathFor :54-60)
- Test: `tests/pdfEditorRouting.test.ts`

**Interfaces:**
- Produces: `export function editorUrlFor(ext: string, opts: { effectiveCanEdit: boolean; canDownload: boolean; pdfEditor: boolean }, token: string): string | undefined`, `editorPathFor('pdf') === '/editor-pdf.html'`

- [ ] **Step 1: 실패하는 테스트**

```ts
// tests/pdfEditorRouting.test.ts
import { describe, expect, test } from 'vitest';
import { editorPathFor, editorUrlFor } from '../src/routes/documents.js';

describe('PDF 편집기 연결', () => {
  const t = 'tok';
  test('pdf는 pdfEditor+canDownload일 때만 편집기 주소(8000은 pdfEditor를 안 보낸다)', () => {
    expect(editorUrlFor('pdf', { effectiveCanEdit: true, canDownload: true, pdfEditor: false }, t)).toBeUndefined();
    expect(editorUrlFor('pdf', { effectiveCanEdit: true, canDownload: false, pdfEditor: true }, t)).toBeUndefined();
    expect(editorUrlFor('pdf', { effectiveCanEdit: false, canDownload: true, pdfEditor: true }, t)).toBe('/editor-pdf.html?token=tok');
    expect(editorUrlFor('pdf', { effectiveCanEdit: true, canDownload: true, pdfEditor: true }, t)).toBe('/editor-pdf.html?token=tok');
  });
  test('기존 형식은 예전과 같다', () => {
    expect(editorUrlFor('docx', { effectiveCanEdit: false, canDownload: false, pdfEditor: false }, t)).toBe('/editor.html?token=tok');
    expect(editorUrlFor('hwp', { effectiveCanEdit: false, canDownload: true, pdfEditor: true }, t)).toBeUndefined();
    expect(editorUrlFor('hwp', { effectiveCanEdit: true, canDownload: false, pdfEditor: false }, t)).toBe('/editor-hwp.html?token=tok');
    expect(editorUrlFor('png', { effectiveCanEdit: true, canDownload: true, pdfEditor: true }, t)).toBeUndefined();
  });
  test('경로', () => {
    expect(editorPathFor('pdf')).toBe('/editor-pdf.html');
  });
});
```

- [ ] **Step 2: 실행해 실패 확인** — `npx vitest run tests/pdfEditorRouting.test.ts` → `editorUrlFor is not a function`

- [ ] **Step 3: 구현**

`src/routes/documents.ts`:

```ts
// tokenBodySchema.properties에 추가
    pdfEditor: { type: 'boolean', default: false },
// TokenBody에 추가
  readonly pdfEditor: boolean;

// editorPathFor에 한 줄 추가(맨 앞)
  if (ext === 'pdf') return '/editor-pdf.html';

/**
 * 토큰 응답의 editorUrl. 기존 형식(docx/xlsx/pptx/hwp/hwpx)은 예전 규칙 그대로다. pdf는 호출자가 `pdfEditor: true`를
 * 보낼 때만(향기랩) 그리고 내려받기 권한이 있을 때만 준다 — 8000 유지보수 시스템은 editorUrl이 있으면 그걸 우선 여는데
 * (WebviewerClient.java), 그쪽 PDF는 보기 화면(마크업)을 계속 써야 하고, 편집기는 원본 바이트를 브라우저로 보내므로
 * 내려받기 금지 문서에 열어 주면 안 된다.
 */
export function editorUrlFor(
  ext: string,
  opts: { effectiveCanEdit: boolean; canDownload: boolean; pdfEditor: boolean },
  token: string,
): string | undefined {
  const url = `${editorPathFor(ext)}?token=${encodeURIComponent(token)}`;
  if (ext === 'pdf') return opts.pdfEditor && opts.canDownload ? url : undefined;
  if (!EDITABLE_EXTS.has(ext)) return undefined;
  const supportsReadOnlyEditor = ext === 'docx' || ext === 'xlsx' || ext === 'pptx';
  return opts.effectiveCanEdit || supportsReadOnlyEditor ? url : undefined;
}
```

`/tokens` 핸들러: `const { viewer, canPrint, canDownload, canCopy, canEdit, pdfEditor } = request.body;`, `supportsReadOnlyEditor`와 기존 `editorUrl:` 식을 `editorUrl: editorUrlFor(meta.ext, { effectiveCanEdit, canDownload, pdfEditor }, token),`로 바꾼다. `effectiveCanEdit` 계산의 `isEditableExt`는 그대로(pdf는 EDITABLE_EXTS에 넣지 않는다 — 넣으면 `findConverter` 조건이 바뀐다).

업로드 meta의 editReady 줄: `...(ext === 'docx' || ext === 'pdf' || (ext === 'xlsx' && config.storageMode === 'memory') ? { editReady: true } : {}),` — pdf는 변환이 통과 복사라 편집기가 원본만으로 연다. 주석에 한 줄 덧붙인다.

- [ ] **Step 4: 통과 확인** — `npx vitest run tests/pdfEditorRouting.test.ts tests/routes-documents.test.ts` PASS

- [ ] **Step 5: 커밋** — `git add src/routes/documents.ts tests/pdfEditorRouting.test.ts && git commit -m "PDF 편집기 1: 토큰 pdfEditor 옵션·editorUrl 결정(8000 영향 없음)"`

---

### Task 2: 서버 — PDF 방(kind `pdf`)과 원본 경로

**Files:**
- Create: `src/pdf/pdfModel.ts`
- Modify: `src/collab/room.ts` (DocKind :173, kindOfExt :175, seedRoom :1403-1560 else 분기 앞, flushToStore :1568, getCurrentBytes :1801, changeLogOf :2525, closeConn의 persistChangeLog :1986)
- Modify: `src/collab/changeLogPersistence.ts` (docKind 타입)
- Modify: `src/routes/edit.ts` (:71 허용 확장자)
- Modify: `src/routes/viewer.ts` (새 경로, download 근처)
- Test: `tests/pdfModel.test.ts`

**Interfaces:**
- Produces: `PDF_PAGES_ARRAY = 'pdfPages'`, `PDF_ANNOTS_MAP = 'pdfAnnots'`, `PDF_FORM_MAP = 'pdfForm'`, `export async function exportPdfFromRoom(original: Uint8Array, ydoc: Y.Doc): Promise<Uint8Array>`, `export function pdfModelIsEmpty(ydoc: Y.Doc): boolean`, `export function pdfChangeLogOf(ydoc: Y.Doc): readonly object[]`; HTTP `GET /api/view/pdf-source?token=` → `application/pdf`(canDownload 필요, 열린 방이 있으면 현재 바이트).

- [ ] **Step 1: 실패하는 테스트**

```ts
// tests/pdfModel.test.ts
import { PDFDocument } from 'pdf-lib';
import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { exportPdfFromRoom, pdfChangeLogOf, pdfModelIsEmpty, PDF_ANNOTS_MAP } from '../src/pdf/pdfModel.js';

async function samplePdf(pages = 2): Promise<Uint8Array> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i += 1) doc.addPage([300, 400]);
  return doc.save();
}

describe('PDF 방 모델(1단계)', () => {
  test('편집이 없으면 저장 결과는 원본 바이트 그대로', async () => {
    const original = await samplePdf();
    const ydoc = new Y.Doc();
    expect(pdfModelIsEmpty(ydoc)).toBe(true);
    const out = await exportPdfFromRoom(original, ydoc);
    expect(out).toBe(original);
    expect(pdfChangeLogOf(ydoc)).toEqual([]);
  });
  test('모델에 무언가 있으면 비어 있지 않다(다음 단계가 채운다)', () => {
    const ydoc = new Y.Doc();
    ydoc.getMap(PDF_ANNOTS_MAP).set('a', { type: 'note' });
    expect(pdfModelIsEmpty(ydoc)).toBe(false);
  });
});
```

- [ ] **Step 2: 실패 확인** — `npx vitest run tests/pdfModel.test.ts` → 모듈 없음

- [ ] **Step 3: 구현**

```ts
// src/pdf/pdfModel.ts
/**
 * PDF 편집 방의 Y.Doc 모델 — 원본 PDF는 그대로 두고 바뀐 것만 담는다(설계 §2.2).
 * 1단계에는 편집 도구가 없어 맵이 비어 있고, 저장은 원본 바이트를 그대로 돌려준다. 2단계 이후가 이 파일의
 * exportPdfFromRoom을 채운다(주석·쪽 정리·양식).
 */
import * as Y from 'yjs';

export const PDF_PAGES_ARRAY = 'pdfPages';
export const PDF_ANNOTS_MAP = 'pdfAnnots';
export const PDF_FORM_MAP = 'pdfForm';

export function pdfModelIsEmpty(ydoc: Y.Doc): boolean {
  return ydoc.getArray(PDF_PAGES_ARRAY).length === 0 && ydoc.getMap(PDF_ANNOTS_MAP).size === 0 && ydoc.getMap(PDF_FORM_MAP).size === 0;
}

export async function exportPdfFromRoom(original: Uint8Array, ydoc: Y.Doc): Promise<Uint8Array> {
  if (pdfModelIsEmpty(ydoc)) return original;
  // 2단계부터 채운다 — 그 전까지는 모델이 생길 길이 없다(편집기에 쓰는 도구가 없음)
  return original;
}

export function pdfChangeLogOf(_ydoc: Y.Doc): readonly object[] {
  return [];
}
```

`src/collab/room.ts`:
- `type DocKind = 'docx' | 'xlsx' | 'pptx' | 'pdf';`, `kindOfExt`에 `if (ext === 'pdf') return 'pdf';`
- seedRoom: `} else if (kind === 'pptx') {` 분기 앞에 `} else if (kind === 'pdf') {\n      // 1단계: 시딩할 것이 없다 — 편집기가 원본을 /api/view/pdf-source로 받아 pdf.js로 그린다\n    }` 를 넣는다(실제 위치는 `if (kind === 'xlsx') {…}` 블록 바로 뒤).
- flushToStore: `if (kind === 'xlsx') {` 앞에
  ```ts
  if (kind === 'pdf') {
    if (pdfModelIsEmpty(ydoc)) return; // 바뀐 것이 없으면 원본을 다시 쓰지 않는다
    const original = await store.readOriginal(docId, meta.ext);
    const next = await exportPdfFromRoom(original, ydoc);
    await store.saveOriginal(docId, meta.ext, Buffer.from(next));
    await convertToPdf(docId, meta.ext); // 보기용 pdf·쪽 크기 갱신(통과 복사)
    invalidateRenderCache(docId);
    invalidateViewCaches(docId);
    return;
  }
  ```
  (convertToPdf/invalidate 함수 이름은 같은 함수의 docx 분기가 쓰는 것과 똑같이 맞춘다 — 그 분기를 열어 이름을 그대로 쓴다.)
- getCurrentBytes: `if (room.kind === 'xlsx') {` 앞에 `if (room.kind === 'pdf') return exportPdfFromRoom(original, room.ydoc);`
- changeLogOf: `if (room.kind === 'pdf') return pdfChangeLogOf(room.ydoc);`
- import: `import { exportPdfFromRoom, pdfChangeLogOf, pdfModelIsEmpty } from '../pdf/pdfModel.js';`

`src/collab/changeLogPersistence.ts`: `docKind: 'xlsx' | 'docx' | 'pptx' | 'pdf'`.

`src/routes/edit.ts:71`: `if (meta.ext !== 'docx' && meta.ext !== 'xlsx' && meta.ext !== 'pptx' && meta.ext !== 'pdf') {` — 위 주석에 "pdf는 PDF 편집기 방(kind pdf, 1단계엔 원본을 고치지 않음)" 한 줄.

`src/routes/viewer.ts` — `/api/view/download` 등록 바로 앞에:

```ts
  // PDF 편집기(editor-pdf.html)가 pdf.js로 그릴 원본. 바이트를 브라우저로 보내는 것이라 내려받기 권한이 있어야 한다
  // (editorUrlFor도 같은 조건으로만 편집기 주소를 준다). 방이 열려 있으면 아직 저장 전인 현재 상태를 준다.
  app.get<{ Querystring: { token?: string } }>('/api/view/pdf-source', async (request, reply) => {
    const view = await resolveView(request.query.token, reply);
    if (!view) return reply;
    if (view.meta.ext !== 'pdf') return reply.code(400).send({ success: false, error: 'PDF 문서가 아닙니다.' });
    if (!view.claims.perms.canDownload) return reply.code(403).send({ success: false, error: '이 문서는 편집기로 열 수 없습니다.' });
    const data = (await getCurrentBytes(view.meta.id)) ?? (await store.readOriginal(view.meta.id, view.meta.ext));
    return reply.type('application/pdf').header('cache-control', 'no-store').send(Buffer.from(data));
  });
```

- [ ] **Step 4: 통과 확인** — `npx vitest run tests/pdfModel.test.ts` PASS, `npx tsc --noEmit -p .` 오류 없음, `npm test` 전부 통과

- [ ] **Step 5: 커밋** — `git commit -m "PDF 편집기 2: 협업 방 kind pdf·원본 경로(1단계는 원본 그대로 저장)"`

---

### Task 3: pdf.js를 편집기 번들에 넣기

**Files:**
- Modify: `scripts/build-editor.mjs`, `.gitignore`
- Create: `src/client/editor-pdf.ts`(이 태스크에서는 pdf.js가 번들되는지만 확인하는 최소 시작 코드)

- [ ] **Step 1: build-editor.mjs**

entryPoints에 `'src/client/editor-pdf.ts'` 추가. 파일 끝에:

```js
// PDF 편집기가 쓰는 pdf.js 부속 — worker(별도 스레드에서 PDF 해석), cMap(글꼴이 안 들어 있는 한글·중국어·일본어 PDF),
// 표준 글꼴. 외부 CDN 없이 자체 호스팅한다.
import { cpSync, mkdirSync } from 'node:fs';
mkdirSync('public/pdfjs', { recursive: true });
cpSync('node_modules/pdfjs-dist/build/pdf.worker.min.mjs', 'public/js/pdf.worker.mjs');
cpSync('node_modules/pdfjs-dist/cmaps', 'public/pdfjs/cmaps', { recursive: true });
cpSync('node_modules/pdfjs-dist/standard_fonts', 'public/pdfjs/standard_fonts', { recursive: true });
```

(import 문은 파일 맨 위로 올린다.)

`.gitignore` 빌드 산출물 목록에 `public/js/editor-pdf.bundle.js`, `public/js/editor-pdf.bundle.js.map`, `public/js/pdf.worker.mjs`, `public/pdfjs/` 추가.

- [ ] **Step 2: 최소 시작 코드**

```ts
// src/client/editor-pdf.ts
import * as pdfjs from 'pdfjs-dist';
pdfjs.GlobalWorkerOptions.workerSrc = '/js/pdf.worker.mjs';
console.info('pdf.js', pdfjs.version);
```

- [ ] **Step 3: 빌드** — `node scripts/build-editor.mjs` 성공, `public/js/editor-pdf.bundle.js`·`public/js/pdf.worker.mjs`·`public/pdfjs/cmaps/Adobe-Korea1-UCS2.bcmap` 존재. 실패(top-level await 등)하면 esbuild `target`을 편집기 전체가 아니라 그대로 두고 `supported: { 'top-level-await': true }`를 추가한다.

- [ ] **Step 4: 커밋** — `git commit -m "PDF 편집기 3: pdf.js 번들·worker·cMap 자체 호스팅"`

---

### Task 4: 찾기 순수 함수

**Files:**
- Create: `src/client/pdf/findText.ts`
- Test: `tests/pdfFindText.test.ts`

**Interfaces:**
- Produces: `export interface TextItemLite { str: string }`, `export interface FindMatch { page: number; itemStart: number; offsetStart: number; itemEnd: number; offsetEnd: number }`, `export function normalizeForFind(s: string): string`, `export function findInPage(page: number, items: readonly TextItemLite[], query: string): FindMatch[]`

- [ ] **Step 1: 실패하는 테스트**

```ts
// tests/pdfFindText.test.ts
import { describe, expect, test } from 'vitest';
import { findInPage, normalizeForFind } from '../src/client/pdf/findText.js';

describe('PDF 찾기', () => {
  test('대소문자·전각·연속 공백을 무시하고, 글자 조각 경계를 넘는 일치도 찾는다', () => {
    const items = [{ str: '분기 보' }, { str: '고서 ' }, { str: 'Report  2026' }];
    const m = findInPage(3, items, '보고서');
    expect(m).toEqual([{ page: 3, itemStart: 0, offsetStart: 3, itemEnd: 1, offsetEnd: 2 }]);
    expect(findInPage(3, items, 'report 2026')).toHaveLength(1);
    expect(normalizeForFind('ＡＢ  c')).toBe('ab c');
  });
  test('빈 검색어는 결과 없음, 여러 번 나오면 모두', () => {
    expect(findInPage(1, [{ str: 'aa aa' }], '')).toEqual([]);
    expect(findInPage(1, [{ str: 'aa aa' }], 'aa')).toHaveLength(2);
  });
});
```

- [ ] **Step 2: 실패 확인** — `npx vitest run tests/pdfFindText.test.ts`

- [ ] **Step 3: 구현**

```ts
// src/client/pdf/findText.ts
/** pdf.js 글자 조각(items) 위에서 찾기 — 조각 경계를 넘는 일치도 찾고, 결과는 조각 번호·조각 안 위치로 돌려준다(강조용). */
export interface TextItemLite {
  readonly str: string;
}
export interface FindMatch {
  readonly page: number;
  readonly itemStart: number;
  readonly offsetStart: number;
  readonly itemEnd: number;
  readonly offsetEnd: number;
}

export function normalizeForFind(s: string): string {
  return s.normalize('NFKC').toLowerCase().replace(/\s+/g, ' ');
}

export function findInPage(page: number, items: readonly TextItemLite[], query: string): FindMatch[] {
  const q = normalizeForFind(query).trim();
  if (!q) return [];
  // 페이지 글자를 한 줄로 이으면서, 이은 글자의 각 위치가 어느 조각의 몇 번째 글자였는지 기록한다(공백 묶음은 한 칸)
  let text = '';
  const posItem: number[] = [];
  const posOffset: number[] = [];
  items.forEach((item, i) => {
    const s = item.str.normalize('NFKC').toLowerCase();
    for (let k = 0; k < s.length; k += 1) {
      const ch = /\s/.test(s[k]) ? ' ' : s[k];
      if (ch === ' ' && text.endsWith(' ')) continue;
      text += ch;
      posItem.push(i);
      posOffset.push(k);
    }
  });
  const out: FindMatch[] = [];
  for (let at = text.indexOf(q); at !== -1; at = text.indexOf(q, at + q.length)) {
    const end = at + q.length - 1;
    out.push({ page, itemStart: posItem[at], offsetStart: posOffset[at], itemEnd: posItem[end], offsetEnd: posOffset[end] + 1 });
  }
  return out;
}
```

- [ ] **Step 4: 통과 확인**
- [ ] **Step 5: 커밋** — `git commit -m "PDF 편집기 4: 찾기 순수 함수"`

---

### Task 5: 편집기 틀(HTML·CSS)

**Files:**
- Create: `public/editor-pdf.html`, `public/css/pdf-editor.css`

화면 구성(설계 §1): 제목 줄(← `.doclist-link`, 문서 탭(파일 이름·저장 점), 오른쪽 연결 상태·참여자), 위 줄(`메뉴` 단추, 탭 `모든 도구`·`읽기`·`편집`·`서명` — 1단계에 편집·서명 탭은 "다음 업데이트" 안내 상자를 보여 준다, 찾기 상자, ⓘ 문서 정보·💾 내려받기·🖨 인쇄), 왼쪽 도구 목록(1단계 항목: 주석 달기·페이지 정리·파일 합치기·페이지 추출·양식·서명·내려받기 — 아직 없는 기능은 "준비 중" 표시), 본문(`#pdf-scroll` > `#pdf-pages`), 떠 있는 도구 줄(선택·손), 오른쪽 정보 창(`#panel-thumbs`·`#panel-outline`·`#panel-layers`), 맨 오른쪽 줄(창 단추 3개, 쪽 번호 입력 `#page-input` / 전체 `#page-count`, ▲▼, 배율 `#zoom-level`, 맞춤 `#fit-btn`, 보기 회전 `#rotate-btn`, 확대 `#zoom-in`, 축소 `#zoom-out`), 찾기 줄(`#find-bar`: 입력·`n/m`·이전·다음·닫기).

`editor-slides.html` 머리(`<head>`, chrome/home/back 처리 스크립트)를 그대로 가져오고 title은 `DocView — PDF 편집`, body class는 `editor-body pdf-editor-page`, `<link rel="stylesheet" href="/css/style.css" /><link rel="stylesheet" href="/css/pdf-editor.css" />`, 끝에 `<script type="module" src="/js/editor-pdf.bundle.js"></script>`.

CSS 원칙: 밝은 회색 바탕(#f1f1f3), 흰 패널, 1px 회색 테두리(#e1dfdd), 강조색 `--pdf-accent: #1a73e8`, 위 줄 높이 44px·제목 줄 38px, 왼쪽 목록 폭 260px, 오른쪽 정보 창 300px, 맨 오른쪽 줄 44px, 쪽은 흰 바탕 + 그림자, 쪽 사이 16px, 떠 있는 도구 줄은 둥근 흰 상자(radius 10px, 그림자). 1100px 이하에서 왼쪽 목록, 900px 이하에서 정보 창을 기본으로 접는다. `.pdf-text-layer`(pdf.js 글자 층: absolute, 글자는 투명, `::selection`만 파란색), `.pdf-find-hit`(노랑 반투명), `.pdf-find-hit.current`(주황).

- [ ] **Step 1: 두 파일 작성**(위 요소 id 그대로 — Task 6~8 코드가 이 id를 쓴다)
- [ ] **Step 2: 시험 서버에서 `/editor-pdf.html` 열어 배치 스크린샷 확인**(데이터 없이도 틀이 보이는지)
- [ ] **Step 3: 커밋** — `git commit -m "PDF 편집기 5: 화면 틀(pdfcraft 배치)"`

---

### Task 6: 문서 열기·쪽 그리기·배율·회전·쪽 이동

**Files:**
- Create: `src/client/pdf/pdfDocument.ts`, `src/client/pdf/pageView.ts`
- Modify: `src/client/editor-pdf.ts`

**Interfaces:**
- Produces:
  - `pdfDocument.ts`: `export async function openPdf(url: string): Promise<PDFDocumentProxy>`(cMapUrl `/pdfjs/cmaps/`, cMapPacked true, standardFontDataUrl `/pdfjs/standard_fonts/`, isEvalSupported false), `export interface OutlineNode { title: string; page: number | null; children: OutlineNode[] }`, `export async function readOutline(doc): Promise<OutlineNode[]>`, `export interface LayerInfo { id: string; name: string; visible: boolean }`, `export async function readLayers(doc): Promise<LayerInfo[]>`
  - `pageView.ts`: `export class PageView { constructor(host: HTMLElement, scroller: HTMLElement, doc: PDFDocumentProxy); init(): Promise<void>; setScale(scale: number | 'fit-width' | 'fit-page'): void; get scale(): number; rotate(delta: 90 | -90): void; goTo(page: number): void; get currentPage(): number; onPageChange(cb: (page: number) => void): void; setLayerVisibility(id: string, visible: boolean): void; textItems(page: number): Promise<{ str: string }[]>; highlight(matches: FindMatch[], current: number): void; scrollToMatch(m: FindMatch): void }`

구현 요점:
- 쪽마다 `div.pdf-page[data-page]`(크기는 viewport 크기로 먼저 잡아 스크롤 길이를 정확히) 안에 canvas + `div.pdf-text-layer`.
- IntersectionObserver(rootMargin 1개 화면)로 보이는 쪽만 그린다. 그린 canvas 크기 = viewport × devicePixelRatio. 배율·회전이 바뀌면 그려 둔 것을 지우고 다시.
- 진행 중 렌더는 `RenderTask.cancel()`로 취소.
- 글자 층: `pdfjs.renderTextLayer({ textContentSource: page.streamTextContent(), container, viewport })`.
- 현재 쪽: 스크롤 중심선에 걸친 쪽(scroll 이벤트 rAF로).
- 레이어: `doc.getOptionalContentConfig()` 한 번 받아 두고 render에 `optionalContentConfigPromise`로 넘긴다. `setLayerVisibility`는 `config.setVisibility(id, visible)` 후 다시 그림.
- 배율: 25%~400%, 확대·축소 단계 [0.25,0.5,0.67,0.75,0.9,1,1.1,1.25,1.5,1.75,2,2.5,3,4], fit-width = (scroller 폭 - 48)/쪽 폭, fit-page = min(폭, 높이 맞춤). Ctrl+휠 = 확대·축소.
- 쪽 번호 입력 Enter → goTo, ▲▼ → 이전/다음, PageUp/PageDown·Home/End.

- [ ] **Step 1~3**: 두 모듈 작성 → editor-pdf.ts에서 `const doc = await openPdf('/api/view/pdf-source?token=' + encodeURIComponent(token)); const view = new PageView(...); await view.init();` 연결, 오른쪽 줄 단추·쪽 번호 연결.
- [ ] **Step 4**: 시험 서버 + 한글 PDF 두 개(글꼴 포함·미포함 — 미포함은 `scripts/fixtures`가 아니라 시험 스크립트가 pdf-lib로 만든다: Type0 글꼴이 Adobe-Korea1을 가리키고 FontFile이 없는 쪽 하나)로 화면 확인.
- [ ] **Step 5: 커밋** — `git commit -m "PDF 편집기 6: 쪽 그리기·배율·회전·쪽 이동"`

---

### Task 7: 정보 창(썸네일·책갈피·레이어)과 찾기

**Files:**
- Create: `src/client/pdf/panels.ts`, `src/client/pdf/findController.ts`
- Modify: `src/client/editor-pdf.ts`

**Interfaces:**
- Consumes: `PageView`, `readOutline`, `readLayers`, `findInPage`
- Produces: `export function setupPanels(opts: { doc; view: PageView; root: HTMLElement }): { show(panel: 'thumbs' | 'outline' | 'layers' | null): void }`, `export class FindController { constructor(view: PageView, doc, ui: { bar; input; count; prev; next; close }); open(): void; close(): void }`

구현 요점:
- 썸네일: 폭 140px, 보이는 것만 그림(IntersectionObserver), 현재 쪽 강조, 눌러 이동.
- 책갈피: 트리(▸ 펼치기), 눌러 이동, 없으면 "이 문서에는 책갈피가 없습니다."
- 레이어: 체크 상자 목록, 없으면 "이 문서에는 레이어가 없습니다."
- 찾기: Ctrl+F·위 찾기 상자 → 찾기 줄. 입력 200ms 뒤 쪽 1부터 차례로 `textItems` → `findInPage`(쪽마다 결과가 나오는 대로 개수 갱신, 새 입력이 오면 이전 작업 취소). Enter/▼ 다음, Shift+Enter/▲ 이전, Esc 닫기. 현재 일치로 스크롤하고 글자 층 span 위에 `.pdf-find-hit` 상자.

- [ ] **Step 1~3**: 작성·연결
- [ ] **Step 4**: 시험 서버에서 확인
- [ ] **Step 5: 커밋** — `git commit -m "PDF 편집기 7: 썸네일·책갈피·레이어·찾기"`

---

### Task 8: 시작 처리(협업 연결·참여자·돌아가기·준비 알림·내려받기·인쇄)

**Files:**
- Modify: `src/client/editor-pdf.ts`

- 토큰: `new URLSearchParams(location.search).get('token')`, 없으면 오류 화면.
- `/api/view/info?token=`로 이름·권한. 문서 탭에 파일 이름.
- Yjs: 다른 편집기와 같은 방식으로 `new WebsocketProvider(wsBase, docId, ydoc, { params: { token } })`(grid-editor.ts의 연결 코드를 그대로 따른다) → 연결 상태 배지, 참여자(`renderPresenceAvatars`).
- 첫 쪽이 그려지면 `window.parent.postMessage({ type: 'docview:ready' }, '*')`(다른 편집기와 같음).
- 💾 내려받기: `/api/view/download?token=`로 이동(새 창 아님), 🖨 인쇄: `/api/view/print?token=`(권한 없으면 단추 숨김).
- ⓘ 문서 정보: 쪽 수·쪽 크기·PDF 버전·제목·작성자(`doc.getMetadata()`)를 작은 창으로.
- 단축키: Ctrl+F 찾기, Ctrl+P 인쇄, Ctrl+=/− 확대·축소, Ctrl+0 맞춤.

- [ ] **Step 1~2**: 작성·빌드
- [ ] **Step 3: 커밋** — `git commit -m "PDF 편집기 8: 협업 연결·참여자·돌아가기·내려받기·인쇄"`

---

### Task 9: 브라우저 확인 스크립트

**Files:**
- Create: `scripts/verify-pdf-editor-e2e.mjs`
- Modify: `package.json` (`"verify:pdf": "node scripts/verify-pdf-editor-e2e.mjs"`)

확인 항목(시험 서버 8124, 업로드 → 토큰 `{ canDownload: true, pdfEditor: true }` → editorUrl):
1. 토큰: pdfEditor 없이는 editorUrl 없음, 있으면 `/editor-pdf.html`.
2. 한글 PDF(LibreOffice로 만든 글꼴 포함 + pdf-lib로 만든 Korea1 미포함) — 첫 쪽 canvas에 그려진 픽셀이 있고, 글자 층에 "분기 보고서"/"한글" 텍스트.
3. 쪽 수 표시, ▼로 2쪽 이동, 쪽 번호 입력 이동, 확대·축소·맞춤 배율 표시, 보기 회전 후 쪽 상자 가로세로 바뀜.
4. 썸네일 눌러 이동, 책갈피(pdf-lib로 개요를 넣은 PDF) 눌러 이동, 레이어(선택적 콘텐츠가 있는 PDF — 시험 스크립트가 OCProperties를 직접 써서 만든다) 끄면 그 쪽 픽셀이 바뀜.
5. 찾기: Ctrl+F → "보고서" → "1/1", 강조 상자 존재, Esc로 닫힘.
6. canDownload 없는 토큰으로 `/api/view/pdf-source` → 403.
7. 내려받기 단추 → 받은 바이트가 원본과 같음.
8. docview:ready 메시지(iframe 호스트 페이지로 확인).
9. 페이지 오류 없음.

- [ ] **Step 1~2**: 작성·실행 → 전부 OK
- [ ] **Step 3: 커밋** — `git commit -m "PDF 편집기 9: 브라우저 확인(verify:pdf)"`

---

### Task 10: 향기랩 연결

**Files:**
- Modify: `C:\Users\OpenMinds\calc-site\public\tools\index.html` (:32 accept, 토큰 요청 :369, 확장자 판정·안내 문구)

- accept에 `.pdf`, 확장자 분기에서 pdf를 오피스와 같은 webviewer 경로로, 토큰 요청 본문에 `pdfEditor: true`.
- 안내 문구(지원 형식 목록, i18n 키가 있으면 `public/assets/i18n-en.js`도)에 PDF 추가.
- 로컬 하네스(`node $S/hyang-harness.mjs C:/Users/OpenMinds/calc-site/public 8790 http://127.0.0.1:8124`)로 PDF 끌어 놓기 → 편집기 → ← 돌아가기 확인.

- [ ] **Step 1~2**: 수정·확인
- [ ] **Step 3: 커밋(calc-site 저장소)** — `git commit -m "문서도구: PDF도 편집기로(pdfEditor)"`

---

### Task 11: 회귀·반영·운영 확인

- [ ] `npm test`, `npm run verify:xlsx`, `npm run verify:pptx`, `npm run verify:pdf`, 워드 리본 e2e(`$S/_tmp_word_ribbon_e2e.mjs`를 프로젝트 폴더로 복사해 실행 후 삭제), 엑셀 홈/병합/조건부 서식 e2e
- [ ] 8000 영향 확인: 시험 서버에서 pdfEditor 없이 토큰 → editorUrl 없음, viewerUrl 그대로(Task 9-1)
- [ ] master merge → build-editor → 8010 재시작(PID 체인, 배터리면 bat 직접 실행) → push
- [ ] calc-site 배포·push
- [ ] 운영(hyanggilab.com)에서 한글 PDF 올려 Task 9의 2·3·5·7 항목과 ← 돌아가기 확인
- [ ] 메모리(project_hyanggilab_docview.md)에 1단계 반영 기록
