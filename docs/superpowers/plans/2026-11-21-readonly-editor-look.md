# 기본 열기 = 읽기 전용 편집기 화면(편집 도구 숨김) 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 메인 화면의 기본 "열기"가 워드·엑셀·PPT 문서를 편집기 화면 그대로, 다만 편집 도구 없이(리본·탭·빠른 실행·파일 화면·수식 입력줄 숨김, 얇은 도구 줄만) 그리고 편집되지 않게 연다.

**Architecture:** 서버는 토큰 발급 `engine: 'editor'`를 받으면 `canEdit`을 끄고(첫 겹) 응답 `viewerUrl`을 `<편집기 쪽>?token=…&ui=view`로, `fallbackUrl`을 같은 토큰의 보기 화면으로 준다. 화면은 `ui=view` + 읽기 전용 토큰일 때만(두 번째 겹) `officeShell`이 `body.ui-view`를 붙이고 제목 줄에 `.view-strip-tools`(확대·이동·인쇄·내려받기·편집 열기)를 단다 — 숨기기는 CSS 한 묶음, 판정·규칙은 새 순수 모듈 `src/client/viewStrip.ts`, 키·우클릭·복사 막기는 같은 모듈의 창 잡기 단계 수신기 하나. 세 편집기는 이전/다음(쪽·시트·슬라이드)과 정원 초과(4429) 자동 이동만 잇는다.

**Tech Stack:** Node 20 · Fastify 5 · TypeScript · esbuild(`npm run build:editor`) · ProseMirror/Yjs 편집기 · vitest · Playwright 1.62

**Spec:** `docs/superpowers/specs/2026-11-21-readonly-editor-look-design.md`

## Global Constraints

- 두 겹(설계 §6): 읽기 전용 = "토큰 `canEdit=false`(서버가 편집을 버림)" + "화면 `ui=view`(도구 숨김)". 보기 꾸밈은 **둘 다**일 때만 — 둘 중 하나라도 아니면 지금 편집기 그대로(향기랩·동시편집 영향 없음).
- `POST /api/documents/:id/tokens` 본문 `engine` 값: `'libreoffice' | 'artcraft' | 'editor'`. `engine: 'editor'`면 서버가 `canEdit`을 false로 강제, 응답 `viewerUrl` = `<editor page>?token=…&ui=view`(docx→`/editor.html`, xlsx→`/editor-grid.html`, pptx→`/editor-slides.html`, `editorPathFor` 그대로), `fallbackUrl` = 같은 토큰의 보기 화면 주소(`viewPathFor` 그대로).
- `engine` 생략·다른 값이면 지금과 같다 — **API 기본값은 바꾸지 않는다**(향기랩·saturn·유지관리 시스템 영향 없음). `editorUrl` 규칙(`editorUrlFor`)도 그대로 — 향기랩이 받는 `canEdit=true` 토큰의 `editorUrl`은 한 글자도 바뀌지 않고, 향기랩은 `ui=view`를 보내지 않는다.
- `canDownload`가 false인 `engine:'editor'` 요청은 거절하지 않고 `engine:'libreoffice'` + `engineReason: '내려받기 권한이 없어 서버 변환으로 엽니다'`.
- 지원 형식: docx·xlsx·pptx(`meta.ext` 기준). pdf·hwp·도면 등 그 밖에 `engine:'editor'`를 보내면 `engine:'libreoffice'`, 이유 없음(설계에 문구가 없어 이 계획이 정함), 주소는 지금 그대로.
- 정원 초과 안내 문구(그대로): `'지금은 보는 사람이 많습니다. 서버 변환으로 보여 줍니다'` — 뒤이어 같은 토큰의 보기 화면으로 자동 이동(토큰 재발급 없음).
- 메인 화면: 기존 단추("리브레오피스로 열기"/"바로 보기", "아트크래프트로 열기", "동시편집 열기", "PDF 편집")의 글자·동작은 그대로 두고 "▾ 다른 방식" 메뉴(`<details class="open-more">`) 안으로 옮기기만 한다. 기존 단추에는 id가 없다 — e2e가 쓰는 접근 이름(글자)을 바꾸지 않는다.
- `canCopy=false` 막기·우클릭 메뉴 없애기·편집 단축키 무시는 보기 꾸밈일 때만. 단 엑셀 시트 추가·삭제·이름 바꾸기·순서·틀 고정 쓰기는 `canEdit`으로 막는다(서버가 어차피 버리는 쓰기라 화면만 어긋나던 원인 — 읽기 전용 편집기 전체의 버그 수정, 이 계획의 결정).
- 운영 규칙: 포트 8000·8010·8020·8100·8110·8130·8148은 쓰지도 건드리지도 않는다(시험 서버는 8150~8159). 프로세스를 이름으로 끄지 않는다(`taskkill /IM node.exe` 금지) — 직접 띄운 것만 PID로 끈다. `public/js/*.bundle.js` 빌드 산출물은 커밋하지 않는다(`public/js/main.js`는 손으로 쓴 소스라 커밋한다). push 하지 않는다. 커밋은 경로를 지정해서(`git add <경로>`), 메시지는 한국어, 꼬리말 `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`.
- 시험 순서: 태스크마다 단위 시험·타입 검사·빌드만, 브라우저 e2e와 회귀 묶음은 Task 7에서 한 번에.
- 하지 않는 것(설계 §5): Yjs 없는 스냅샷 모드, 향기랩 화면 변경, API 기본값 변경. PPT 편집기에는 찾기·확대 기능이 원래 없어 얇은 줄에도 넣지 않는다.

---

## 파일 구조

| 파일 | 할 일 |
|---|---|
| `src/artcraftEngine.ts` | `Engine`에 `'editor'`, `EDITOR_VIEW_FORMATS`, `engineFor`의 editor 판정 |
| `src/routes/documents.ts` | 본문 enum, `canEdit` 강제, 응답 `viewerUrl`·`fallbackUrl` |
| `src/client/viewStrip.ts` (새) | 보기 꾸밈 판정·규칙(순수) + 키·우클릭·복사 막기, 정원 초과 이동(DOM) |
| `src/client/officeShell.ts` | `ui-view` 모드: 클래스·얇은 줄 단추·상태 줄 "읽기 전용"·`setViewNav`·편집 열기 |
| `public/css/style.css` | 보기 꾸밈 숨김 묶음·얇은 줄·안내·600px, 메인 화면 "▾ 다른 방식" 메뉴 |
| `src/client/editor.ts` | 쪽 이동, 4429 자동 이동 |
| `src/client/grid-editor.ts` | 시트 이동, 구조 쓰기 `canEdit` 막기, 시트 추가 단추 숨김, 4429 |
| `src/client/editor-slides.ts` | 슬라이드 이동, 4429 |
| `public/js/main.js` | 기본 "열기", "▾ 다른 방식", 공유 폼 라디오, 편집 열기용 sessionStorage |
| `scripts/verify-artcraft-e2e.mjs` | 메뉴 안으로 옮긴 단추를 누르기 전에 메뉴 열기(두 군데) |
| `scripts/verify-readonly-editor-e2e.mjs` (새) | 설계 §4 (1)~(8) 브라우저 검사 |
| `package.json` | `verify:readonly` |
| `tests/artcraftEngine.test.ts`, `tests/readonlyEditorTokens.test.ts` (새), `tests/viewStrip.test.ts` (새) | 단위 시험 |

---

### Task 1: 서버 — `engine: 'editor'`

**Files:**
- Modify: `src/artcraftEngine.ts:8` (`Engine`), `:34-45` (`engineFor`)
- Modify: `src/routes/documents.ts:41` (enum), `:542` (`effectiveCanEdit`), `:571-582` (응답)
- Test: `tests/artcraftEngine.test.ts` (덧붙임), `tests/readonlyEditorTokens.test.ts` (새)

**Interfaces:**
- Consumes: 없음
- Produces: `type Engine = 'libreoffice' | 'artcraft' | 'editor'`; `EDITOR_VIEW_FORMATS: ReadonlySet<string>`; `engineFor(ext, perms, requested, installed)`가 `requested === 'editor'`면 `{ engine: 'editor' }` 또는 `{ engine: 'libreoffice', reason? }`. 토큰 응답에 `fallbackUrl?: string`(engine이 editor일 때만), `viewerUrl`은 그때 `<editorPathFor(ext)>?token=<enc>&ui=view`.

- [ ] **Step 1: 판정 시험을 쓴다** — `tests/artcraftEngine.test.ts` 끝(마지막 `describe` 뒤)에 덧붙인다.

```ts
describe('engineFor — 읽기 전용 편집기(editor, 설계 readonly-editor-look §3.2)', () => {
  const table: Array<[string, boolean, Engine, string | undefined]> = [
    ['docx', true, 'editor', undefined],
    ['xlsx', true, 'editor', undefined],
    ['pptx', true, 'editor', undefined],
    ['docx', false, 'libreoffice', ENGINE_REASONS.noDownload],
    ['xlsx', false, 'libreoffice', ENGINE_REASONS.noDownload],
    ['pptx', false, 'libreoffice', ENGINE_REASONS.noDownload],
    // 편집기가 없는 형식은 지금 경로 그대로(이유 문구는 설계에 없다)
    ['pdf', true, 'libreoffice', undefined],
    ['hwp', true, 'libreoffice', undefined],
    ['doc', true, 'libreoffice', undefined],
    ['constructor', true, 'libreoffice', undefined],
  ];
  test.each(table)('%s · 내려받기 %s → %s', (ext, canDownload, engine, reason) => {
    // 아트크래프트 설치 여부(manifest)와 무관하다
    for (const installed of [true, false]) {
      expect(engineFor(ext, { canDownload }, 'editor', installed)).toEqual(reason ? { engine, reason } : { engine });
    }
  });
});
```

- [ ] **Step 2: 토큰 경로 시험을 쓴다** — `tests/readonlyEditorTokens.test.ts` 새 파일.

```ts
import { mkdirSync, mkdtempSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';

describe('읽기 전용 편집기 보기 꾸밈 — 토큰 engine:editor', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let newDocumentId: (typeof import('../src/storage.js'))['newDocumentId'];
  let verifyViewToken: (typeof import('../src/security.js'))['verifyViewToken'];
  const dataDir = mkdtempSync(path.join(os.tmpdir(), 'readonly-editor-data-'));

  beforeAll(async () => {
    mkdirSync(path.join(dataDir, 'audit'));
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_DATA_DIR', dataDir);
    vi.stubEnv('DOCVIEW_PERSIST_LOGS', '1');
    ({ store, newDocumentId } = await import('../src/storage.js'));
    ({ verifyViewToken } = await import('../src/security.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  async function addDoc(ext: string) {
    const id = newDocumentId();
    await store.saveOriginal(id, ext, Buffer.from(`원본-${ext}`));
    await store.writeMeta({ id, name: `시험.${ext}`, ext, status: 'ready', pageCount: 3, createdAt: new Date().toISOString() });
    return id;
  }
  const issue = (id: string, body: Record<string, unknown>) =>
    app.inject({ method: 'POST', url: `/api/documents/${id}/tokens`, payload: { viewer: '시험자', ...body } });
  const data = async (id: string, body: Record<string, unknown>) => (await issue(id, body)).json().data;

  test.each([
    ['docx', '/editor.html', '/view.html'],
    ['xlsx', '/editor-grid.html', '/viewer.html'],
    ['pptx', '/editor-slides.html', '/view.html'],
  ])('%s + 내려받기 + editor → %s?…&ui=view, fallbackUrl %s, canEdit 강제 false', async (ext, page, fallback) => {
    const id = await addDoc(ext);
    const d = await data(id, { canDownload: true, canEdit: true, engine: 'editor' });
    const t = encodeURIComponent(d.token);
    expect(d.engine).toBe('editor');
    expect(d.engineReason).toBeUndefined();
    expect(d.viewerUrl).toBe(`${page}?token=${t}&ui=view`);
    expect(d.fallbackUrl).toBe(`${fallback}?token=${t}`);
    const claims = verifyViewToken(d.token);
    expect(claims.perms.canEdit).toBe(false);
    expect('engine' in claims).toBe(false);
  });

  test('내려받기 권한이 없으면 libreoffice + 이유, 보기 화면 주소, fallbackUrl 없음, canEdit false', async () => {
    const id = await addDoc('docx');
    const d = await data(id, { canEdit: true, engine: 'editor' });
    expect(d.engine).toBe('libreoffice');
    expect(d.engineReason).toBe('내려받기 권한이 없어 서버 변환으로 엽니다');
    expect(d.viewerUrl).toBe(`/view.html?token=${encodeURIComponent(d.token)}`);
    expect(d.fallbackUrl).toBeUndefined();
    expect(verifyViewToken(d.token).perms.canEdit).toBe(false);
  });

  test('pdf에 editor → 지금 경로 그대로(이유 없음, 편집기 주소 없음)', async () => {
    const id = await addDoc('pdf');
    const d = await data(id, { canDownload: true, engine: 'editor' });
    expect(d.engine).toBe('libreoffice');
    expect(d.engineReason).toBeUndefined();
    expect(d.viewerUrl).toBe(`/view.html?token=${encodeURIComponent(d.token)}`);
    expect(d.fallbackUrl).toBeUndefined();
    expect(d.editorUrl).toBeUndefined();
  });

  test('engine 생략 → 지금과 같다(보기 화면 주소, editorUrl에 ui 없음, 편집 토큰은 편집 그대로)', async () => {
    const id = await addDoc('docx');
    const d = await data(id, { canDownload: true, canEdit: true });
    const t = encodeURIComponent(d.token);
    expect(d.engine).toBe('libreoffice');
    expect(d.viewerUrl).toBe(`/view.html?token=${t}`);
    expect(d.editorUrl).toBe(`/editor.html?token=${t}`);
    expect(d.fallbackUrl).toBeUndefined();
    expect(verifyViewToken(d.token).perms.canEdit).toBe(true);
  });

  test('모르는 engine 값은 400', async () => {
    const id = await addDoc('docx');
    expect((await issue(id, { engine: 'foo' })).statusCode).toBe(400);
  });
});
```

- [ ] **Step 3: 시험이 실패하는지 본다**

Run: `npx vitest run tests/artcraftEngine.test.ts tests/readonlyEditorTokens.test.ts`
Expected: FAIL — 판정 표의 editor 줄은 `{ engine: 'libreoffice' }`가 나와 다르고, 토큰 시험은 `engine: 'editor'`가 enum에 없어 400(`d`가 undefined라 TypeError). 생략·400 시험 두 개는 이미 통과.

- [ ] **Step 4: 판정을 넣는다** — `src/artcraftEngine.ts`

`:8`을 바꾼다.

```ts
/** libreoffice = 서버 변환 보기 화면, artcraft = 브라우저 wasm 엔진, editor = 읽기 전용 편집기 화면(보기 꾸밈, 설계 readonly-editor-look §3.2) */
export type Engine = 'libreoffice' | 'artcraft' | 'editor';
```

`ARTCRAFT_FORMATS` 정의 바로 아래에 더한다.

```ts
/** 편집기 화면이 있는 형식 — 옛 형식(.doc·.xls·.ppt)은 업로드 때 바뀐 저장본의 meta.ext로 들어온다(editorUrlFor와 같은 규칙) */
export const EDITOR_VIEW_FORMATS: ReadonlySet<string> = new Set(['docx', 'xlsx', 'pptx']);
```

`engineFor` 본문 첫 줄(`if (requested !== 'artcraft') …`) 앞에 넣는다.

```ts
  if (requested === 'editor') {
    if (!EDITOR_VIEW_FORMATS.has(ext)) return { engine: 'libreoffice' };
    // 편집기 자료 경로는 열람만(편집·내려받기 둘 다 없음) 토큰을 받지 않는다(edit.ts isViewOnly → 4403)
    return perms.canDownload ? { engine: 'editor' } : { engine: 'libreoffice', reason: ENGINE_REASONS.noDownload };
  }
```

- [ ] **Step 5: 토큰 경로를 바꾼다** — `src/routes/documents.ts`

`:40-41`을 바꾼다.

```ts
    // 여는 엔진(설계 2026-11-21 §3.1) — artcraft는 브라우저 wasm 엔진, editor는 읽기 전용 편집기 화면(보기 꾸밈, 설계 readonly-editor-look §3.2).
    // 안 되는 경우는 거절하지 않고 libreoffice + engineReason
    engine: { type: 'string', enum: ['libreoffice', 'artcraft', 'editor'], default: 'libreoffice' },
```

`:542`를 바꾼다.

```ts
      // engine:'editor'는 읽기 전용 참여만 — 요청의 canEdit과 무관하게 끈다(두 겹 중 서버 쪽 겹, 설계 §6)
      const effectiveCanEdit = engine !== 'editor' && canEdit && (!isEditableExt || findConverter(meta.ext) !== null);
```

`return { success: true, data: { … } }` 블록(`:571-582`)을 바꾼다.

```ts
      const viewerUrl = `${viewPathFor(meta.ext, meta.viewKind)}?token=${encodeURIComponent(token)}`;
      return {
        success: true,
        data: {
          token,
          // 쪽 문서는 통합 보기 껍데기, 엑셀 격자·도면·한글은 지금 보기 화면(viewPathFor) — 주소·토큰 계약은 그대로(설계 V2).
          // engine:'editor'면 편집기 화면 + ui=view(도구 숨김), fallbackUrl은 같은 토큰의 보기 화면(큰 엑셀 정원 초과 때 옮겨 간다)
          viewerUrl: decision.engine === 'editor' ? `${editorPathFor(meta.ext)}?token=${encodeURIComponent(token)}&ui=view` : viewerUrl,
          ...(decision.engine === 'editor' ? { fallbackUrl: viewerUrl } : {}),
          editorUrl: editorUrlFor(meta.ext, { effectiveCanEdit, canDownload, pdfEditor }, token),
          expiresInSeconds: config.tokenTtlSeconds,
          engine: decision.engine,
          ...(decision.reason ? { engineReason: decision.reason } : {}),
        },
      };
```

(`editorPathFor`는 같은 파일 `:69`에 있다. 토큰 클레임의 `engine`은 지금처럼 `'artcraft'`일 때만 넣는다 — `:555`는 고치지 않는다.)

- [ ] **Step 6: 시험이 통과하는지 본다**

Run: `npx vitest run tests/artcraftEngine.test.ts tests/readonlyEditorTokens.test.ts tests/artcraftRoutes.test.ts && npm run typecheck`
Expected: PASS, 타입 오류 0.

- [ ] **Step 7: 커밋**

```bash
git add src/artcraftEngine.ts src/routes/documents.ts tests/artcraftEngine.test.ts tests/readonlyEditorTokens.test.ts
git commit -m "보기 꾸밈 1: 토큰 engine:'editor' — canEdit 강제 끔, viewerUrl=편집기+ui=view, fallbackUrl, 내려받기 없으면 서버 변환+이유" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 2: 공용 틀 — `viewStrip.ts`와 `officeShell` 보기 꾸밈

**Files:**
- Create: `src/client/viewStrip.ts`
- Modify: `src/client/officeShell.ts` (옵션·반환 타입, 인쇄 함수 뽑기, 보기 꾸밈 블록)
- Modify: `src/client/editor.ts:708-715`, `src/client/grid-editor.ts:459-466`, `src/client/editor-slides.ts:1314-1321` (옵션 두 개 넘기기)
- Modify: `public/css/style.css` (끝에 덧붙임)
- Test: `tests/viewStrip.test.ts` (새)

**Interfaces:**
- Consumes: Task 1의 `?ui=view` 주소 모양
- Produces (`src/client/viewStrip.ts`):
  - `EDIT_DOC_KEY = 'docview-edit-doc'` — 메인 화면이 sessionStorage에 문서 id를 넣는 키(Task 6)
  - `CAPACITY_NOTICE: string`
  - `type StripAction = 'zoom-out' | 'zoom-in' | 'print' | 'download' | 'edit'`
  - `isViewUi(search: string, canEdit: boolean): boolean`
  - `stripActions(p: { canPrint: boolean; canDownload: boolean; hasZoom: boolean; canOpenEdit: boolean }): StripAction[]`
  - `blockedViewKey(e: { key: string; ctrl: boolean; inField: boolean; canCopy: boolean }): boolean`
  - `stepIndex(length: number, current: number, dir: 1 | -1): number`
  - `pickBoundary(tops: readonly number[], dir: 1 | -1): number`
  - `capacityFallback(code: number, viewUi: boolean, viewerUrl: string | undefined): string | null`
  - `installViewGuards(canCopy: boolean): void`, `fallBackToViewer(url: string): void` (DOM)
- Produces (`officeShell.ts`): `OfficeShellOptions`에 `docId: string`, `canCopy: boolean`; `interface ViewNav { prevLabel: string; nextLabel: string; prev(): void; next(): void }`; `OfficeShell.setViewNav(nav: ViewNav): void`(보기 꾸밈이 아니면 아무것도 안 함). 얇은 줄 단추는 모두 `.view-strip-tools [data-act=<이름>]`(`zoom-out`·`zoom-in`·`prev`·`next`·`print`·`download`·`edit`).

- [ ] **Step 1: 규칙 시험을 쓴다** — `tests/viewStrip.test.ts`

```ts
import { describe, expect, test } from 'vitest';
import { CAPACITY_NOTICE, blockedViewKey, capacityFallback, isViewUi, pickBoundary, stepIndex, stripActions } from '../src/client/viewStrip.js';

describe('isViewUi — 두 겹(토큰 읽기 전용 + ui=view)', () => {
  test.each([
    ['?token=t&ui=view', false, true],
    ['?ui=view&token=t', false, true],
    ['?token=t&ui=view', true, false],
    ['?token=t', false, false],
    ['?token=t&ui=edit', false, false],
    ['?token=t&ui=VIEW', false, false],
  ])('%s · canEdit %s → %s', (search, canEdit, expected) => {
    expect(isViewUi(search, canEdit)).toBe(expected);
  });
});

describe('stripActions — 얇은 줄 단추', () => {
  test('권한 다 있고 확대 있고 편집 열기 가능 → 정해진 순서로 전부', () => {
    expect(stripActions({ canPrint: true, canDownload: true, hasZoom: true, canOpenEdit: true })).toEqual(['zoom-out', 'zoom-in', 'print', 'download', 'edit']);
  });
  test('PPT(확대 없음)·인쇄 권한 없음·공유 링크(편집 열기 없음) → 내려받기만', () => {
    expect(stripActions({ canPrint: false, canDownload: true, hasZoom: false, canOpenEdit: false })).toEqual(['download']);
  });
  test('인쇄만', () => {
    expect(stripActions({ canPrint: true, canDownload: false, hasZoom: false, canOpenEdit: false })).toEqual(['print']);
  });
});

describe('blockedViewKey — 편집 단축키는 무시, 찾기·이동은 그대로', () => {
  test.each([
    ['s', true, false, true, true],
    ['S', true, false, true, true],
    ['h', true, false, true, true],
    ['v', true, false, true, true],
    ['x', true, false, true, true],
    ['z', true, false, true, true],
    ['y', true, false, true, true],
    ['c', true, false, true, false],
    ['c', true, false, false, true],
    ['f', true, false, false, false],
    ['a', true, false, false, false],
    ['Delete', false, false, true, true],
    ['Backspace', false, false, true, true],
    ['F2', false, false, true, true],
    ['ArrowDown', false, false, true, false],
    ['PageDown', false, false, true, false],
    ['a', false, false, true, false],
    ['Backspace', false, true, true, false],
    ['v', true, true, true, false],
  ])('%s · ctrl %s · 입력칸 %s · 복사 %s → %s', (key, ctrl, inField, canCopy, expected) => {
    expect(blockedViewKey({ key, ctrl, inField, canCopy })).toBe(expected);
  });
});

describe('stepIndex — 시트·슬라이드 한 칸 앞/뒤(끝에서 멈춤)', () => {
  test.each([
    [3, 0, 1, 1],
    [3, 2, 1, 2],
    [3, 0, -1, 0],
    [3, 2, -1, 1],
    [3, -1, 1, 0],
    [0, -1, 1, -1],
  ] as const)('길이 %s · 지금 %s · 방향 %s → %s', (length, current, dir, expected) => {
    expect(stepIndex(length, current, dir)).toBe(expected);
  });
});

describe('pickBoundary — 워드 쪽 경계(도구 줄 아래 기준 px)', () => {
  test('다음: 기준보다 아래인 첫 경계, 이전: 기준보다 위인 마지막 경계', () => {
    const tops = [-500, -2, 0.5, 300];
    expect(pickBoundary(tops, 1)).toBe(3);
    expect(pickBoundary(tops, -1)).toBe(1);
  });
  test('없으면 -1', () => {
    expect(pickBoundary([], 1)).toBe(-1);
    expect(pickBoundary([100], -1)).toBe(-1);
    expect(pickBoundary([-100], 1)).toBe(-1);
  });
});

describe('capacityFallback — 큰 엑셀 정원 초과(4429)', () => {
  const url = '/viewer.html?token=t';
  test('보기 꾸밈 + 4429 + 보기 주소 → 그 주소', () => expect(capacityFallback(4429, true, url)).toBe(url));
  test('편집 화면·다른 거부·주소 없음 → null(지금처럼 오류 화면)', () => {
    expect(capacityFallback(4429, false, url)).toBeNull();
    expect(capacityFallback(4403, true, url)).toBeNull();
    expect(capacityFallback(4429, true, undefined)).toBeNull();
  });
  test('안내 문구는 설계 그대로', () => expect(CAPACITY_NOTICE).toBe('지금은 보는 사람이 많습니다. 서버 변환으로 보여 줍니다'));
});
```

- [ ] **Step 2: 실패를 본다**

Run: `npx vitest run tests/viewStrip.test.ts`
Expected: FAIL — `Failed to resolve import "../src/client/viewStrip.js"`.

- [ ] **Step 3: `src/client/viewStrip.ts`를 쓴다**

```ts
/**
 * 읽기 전용 편집기의 "보기 꾸밈"(ui=view, 설계 2026-11-21-readonly-editor-look §3.1). 판정·규칙은 순수 함수(tests/viewStrip.test.ts),
 * 막기(installViewGuards)와 정원 초과 이동(fallBackToViewer)은 DOM. 얇은 줄 단추를 다는 일은 officeShell.ts가 한다.
 * 실제 쓰기 차단은 서버(토큰 canEdit=false → room.ts가 Update를 버림) — 여기 막기는 화면이 어긋나지 않게 하는 겹이다.
 */

/** 메인 화면이 "열기" 때 문서 id를 넣어 두는 sessionStorage 키 — 같은 문서면 얇은 줄에 "편집 열기"(public/js/main.js) */
export const EDIT_DOC_KEY = 'docview-edit-doc';
export const CAPACITY_NOTICE = '지금은 보는 사람이 많습니다. 서버 변환으로 보여 줍니다';

export type StripAction = 'zoom-out' | 'zoom-in' | 'print' | 'download' | 'edit';

/** 두 겹 — 토큰이 읽기 전용이고 주소에 ui=view가 있을 때만(하나라도 아니면 지금 편집기 그대로) */
export function isViewUi(search: string, canEdit: boolean): boolean {
  return !canEdit && new URLSearchParams(search).get('ui') === 'view';
}

/** 이전/다음 단추는 편집기가 setViewNav로 따로 단다(쪽·시트·슬라이드) */
export function stripActions(p: { readonly canPrint: boolean; readonly canDownload: boolean; readonly hasZoom: boolean; readonly canOpenEdit: boolean }): StripAction[] {
  const actions: StripAction[] = [];
  if (p.hasZoom) actions.push('zoom-out', 'zoom-in');
  if (p.canPrint) actions.push('print');
  if (p.canDownload) actions.push('download');
  if (p.canOpenEdit) actions.push('edit');
  return actions;
}

const BLOCKED_CTRL_KEYS = new Set(['s', 'h', 'v', 'x', 'z', 'y']);

/** 보기 꾸밈에서 먹이지 않을 키 — 찾기(Ctrl+F)·확대·화살표 이동은 그대로, 입력칸(찾기 상자) 안은 막지 않는다 */
export function blockedViewKey(e: { readonly key: string; readonly ctrl: boolean; readonly inField: boolean; readonly canCopy: boolean }): boolean {
  if (e.inField) return false;
  const key = e.key.toLowerCase();
  if (e.ctrl) return BLOCKED_CTRL_KEYS.has(key) || (!e.canCopy && key === 'c');
  return key === 'delete' || key === 'backspace' || key === 'f2';
}

/** 목록 안에서 한 칸 앞/뒤(끝에서 멈춤). 지금 위치를 모르면(-1) 첫 칸, 빈 목록이면 -1 */
export function stepIndex(length: number, current: number, dir: 1 | -1): number {
  if (length === 0) return -1;
  if (current < 0) return 0;
  return Math.min(length - 1, Math.max(0, current + dir));
}

/** 워드 쪽 이동 — 쪽 경계들의 위치(도구 줄 아래 기준 px)에서 다음(아래 첫)/이전(위 마지막) 경계 번호, 없으면 -1. ±1px은 "지금 여기" */
export function pickBoundary(tops: readonly number[], dir: 1 | -1): number {
  if (dir === 1) return tops.findIndex((top) => top > 1);
  for (let i = tops.length - 1; i >= 0; i -= 1) if (tops[i] < -1) return i;
  return -1;
}

/** 큰 엑셀 정원 초과(edit.ts 4429)면 보기 꾸밈만 같은 토큰의 보기 화면으로 — 편집 화면은 지금처럼 오류 화면 */
export function capacityFallback(code: number, viewUi: boolean, viewerUrl: string | undefined): string | null {
  return code === 4429 && viewUi && viewerUrl ? viewerUrl : null;
}

function isField(target: EventTarget | null): boolean {
  return target instanceof HTMLElement && (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA');
}

/**
 * 창의 잡기 단계에서 끊는다 — 편집기들의 키·우클릭 처리기는 모두 요소나 창의 거품 단계에 걸려 있어 여기서 멈추면 하나도 돌지 않는다.
 * 우클릭: 편집기 메뉴(행·열·시트 탭·칸)는 늘 없애고, 브라우저 기본 메뉴는 복사 권한이 있을 때만 둔다(보기 화면 view.ts setupDeterrents와 같은 규칙).
 */
export function installViewGuards(canCopy: boolean): void {
  window.addEventListener(
    'keydown',
    (event) => {
      if (!blockedViewKey({ key: event.key, ctrl: event.ctrlKey || event.metaKey, inField: isField(event.target), canCopy })) return;
      event.preventDefault();
      event.stopImmediatePropagation();
    },
    true,
  );
  window.addEventListener(
    'contextmenu',
    (event) => {
      event.stopImmediatePropagation();
      if (!canCopy) event.preventDefault();
    },
    true,
  );
  const types: Array<'paste' | 'cut' | 'drop' | 'copy'> = canCopy ? ['paste', 'cut', 'drop'] : ['paste', 'cut', 'drop', 'copy'];
  for (const type of types) {
    window.addEventListener(
      type,
      (event) => {
        if (isField(event.target)) return;
        event.preventDefault();
        event.stopImmediatePropagation();
      },
      true,
    );
  }
}

/** 안내를 띄우고 잠시 뒤 보기 화면으로(뒤로 가기로 다시 막힌 화면에 오지 않게 replace) */
export function fallBackToViewer(url: string): void {
  const note = document.createElement('div');
  note.className = 'view-fallback-notice';
  note.setAttribute('role', 'status');
  note.textContent = CAPACITY_NOTICE;
  document.body.append(note);
  setTimeout(() => location.replace(url), 1500);
}
```

- [ ] **Step 4: 규칙 시험 통과를 본다**

Run: `npx vitest run tests/viewStrip.test.ts`
Expected: PASS.

- [ ] **Step 5: `officeShell.ts` — 가져오기·옵션·반환 타입**

파일 맨 위 `import { mountBackToView } from './backToView.js';` 아래에 더한다.

```ts
import { officeAlert } from './officeDialog.js';
import { EDIT_DOC_KEY, installViewGuards, isViewUi, stripActions, type StripAction } from './viewStrip.js';
```

`OfficeShellOptions`의 `readonly token: string;` 아래와 `readonly canPrint: boolean;` 아래에 각각 더한다.

```ts
  /** /api/view/info 의 docId — 보기 꾸밈의 "편집 열기"가 같은 문서의 편집 토큰을 다시 받는다 */
  readonly docId: string;
```

```ts
  /** 보기 꾸밈(ui=view)에서만 쓴다 — false면 복사·선택·브라우저 우클릭 메뉴를 막는다(보기 화면 view.ts와 같은 규칙) */
  readonly canCopy: boolean;
```

`export interface OfficeShell { … }`를 바꾼다.

```ts
/** 보기 꾸밈 얇은 줄의 이전/다음 — 워드는 쪽, 엑셀은 시트, PPT는 슬라이드 */
export interface ViewNav {
  readonly prevLabel: string;
  readonly nextLabel: string;
  prev(): void;
  next(): void;
}

export interface OfficeShell {
  selectTab(tab: string): void;
  /** 상황별 탭(예: 표 안에 커서가 있을 때만 보이는 "표 레이아웃")을 보이거나 숨긴다. */
  setContextualTab(tab: string, visible: boolean): void;
  /** 보기 꾸밈일 때만 얇은 줄에 이전/다음 단추를 단다. 아니면 아무것도 하지 않는다 */
  setViewNav(nav: ViewNav): void;
}
```

`const EXT_LABELS …` 정의 아래에 더한다(아이콘은 `/view.html` 도구 막대와 같은 모양).

```ts
const svgIcon = (inner: string): string =>
  `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${inner}</svg>`;

const STRIP_BUTTONS: Readonly<Record<StripAction | 'prev' | 'next', { readonly label: string; readonly icon: string }>> = {
  'zoom-out': { label: '축소', icon: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/><path d="M8 11h6"/>' },
  'zoom-in': { label: '확대', icon: '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/><path d="M11 8v6"/><path d="M8 11h6"/>' },
  prev: { label: '이전', icon: '<path d="m15 18-6-6 6-6"/>' },
  next: { label: '다음', icon: '<path d="m9 18 6-6-6-6"/>' },
  print: {
    label: '인쇄',
    icon: '<path d="M6 9V4h12v5"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8" rx="1"/>',
  },
  download: { label: '내려받기', icon: '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>' },
  edit: { label: '편집 열기', icon: '' },
};

function readEditDoc(): string | null {
  try {
    return sessionStorage.getItem(EDIT_DOC_KEY);
  } catch {
    return null;
  }
}

/** 메인 화면에서 연 업로더 본인만 — 같은 문서의 편집 토큰을 지금 받아 편집기로(API 키 규칙은 메인 화면과 같다) */
async function openEditFromView(docId: string): Promise<void> {
  const res = await fetch(`/api/documents/${encodeURIComponent(docId)}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '업로더', canPrint: true, canDownload: true, canCopy: true, canEdit: true }),
  }).catch(() => null);
  const json = (await res?.json().catch(() => null)) as { success?: boolean; data?: { editorUrl?: string }; error?: string } | null;
  if (!res?.ok || !json?.success || !json.data?.editorUrl) {
    await officeAlert(json?.error ?? '편집 화면을 열지 못했습니다. 잠시 뒤 다시 눌러 주세요.');
    return;
  }
  location.href = json.data.editorUrl;
}
```

- [ ] **Step 6: `officeShell.ts` — 인쇄 함수 뽑기**

`backstage.querySelector('[data-action="print"]')?.addEventListener('click', async (event) => { … });` 블록 전체를 바꾼다(동작은 같다 — `window.open`은 여전히 클릭과 같은 틱에 불린다).

```ts
  /** 인쇄 — 백스테이지와 보기 꾸밈 얇은 줄이 같이 쓴다 */
  async function printDoc(): Promise<void> {
    // 새 창은 클릭과 같은 틱에 열어야 팝업 차단에 걸리지 않는다 — 저장이 끝나면 주소를 채운다.
    const printWindow = window.open('about:blank', '_blank');
    try {
      if (options.canEdit) await fetch(`/api/view/save?token=${encodeURIComponent(options.token)}`, { method: 'POST' });
    } catch {
      // 저장이 실패해도 마지막으로 저장된 상태를 인쇄한다
    }
    const url = `/api/view/print?token=${encodeURIComponent(options.token)}`;
    if (printWindow) printWindow.location.href = url;
    else window.open(url, '_blank');
  }
  backstage.querySelector('[data-action="print"]')?.addEventListener('click', async (event) => {
    const button = event.currentTarget as HTMLButtonElement;
    button.disabled = true;
    button.textContent = '저장하는 중…';
    await printDoc();
    button.disabled = false;
    button.textContent = '인쇄';
  });
```

- [ ] **Step 7: `officeShell.ts` — 보기 꾸밈 블록과 `setViewNav`**

`if (!tabs.some((tab) => tab.dataset.tab === activeTab && !tab.hidden)) activeTab = firstVisibleTab();` 줄 **바로 위**에 넣는다.

```ts
  // ---------- 보기 꾸밈(ui=view, 설계 2026-11-21-readonly-editor-look §3.1) ----------
  // 숨기기는 CSS(body.ui-view) 한 묶음, 여기서는 제목 줄에 얇은 도구 줄 단추를 단다(문서 이름·찾기·참석자는 제목 줄에 이미 있다).
  let stripTools: HTMLElement | null = null;
  if (isViewUi(location.search, options.canEdit)) {
    document.body.classList.add('ui-view');
    document.body.classList.toggle('no-copy', !options.canCopy);
    installViewGuards(options.canCopy);
    toolbar.querySelector('.office-titlebar')?.classList.add('view-strip');
    stripTools = document.createElement('div');
    stripTools.className = 'view-strip-tools';
    stripTools.setAttribute('role', 'toolbar');
    stripTools.setAttribute('aria-label', '보기 도구');
    const actions = stripActions({
      canPrint: options.canPrint,
      canDownload: options.canDownload,
      hasZoom: document.getElementById('zoom-in') !== null,
      canOpenEdit: readEditDoc() === options.docId,
    });
    for (const act of actions) stripTools.append(stripButton(act));
    titleRight?.insertBefore(stripTools, titleRight.querySelector('#presence-list'));
    const readOnlyItem = document.createElement('span');
    readOnlyItem.className = 'statusbar-item view-readonly';
    readOnlyItem.textContent = '읽기 전용';
    document.getElementById('statusbar')?.prepend(readOnlyItem);
  }

  function stripButton(act: StripAction | 'prev' | 'next', label: string = STRIP_BUTTONS[act].label): HTMLElement {
    if (act === 'edit') {
      const edit = document.createElement('button');
      edit.type = 'button';
      edit.className = 'btn-pill back-to-view';
      edit.dataset.act = act;
      edit.textContent = label;
      edit.title = '편집할 수 있는 화면으로 다시 엽니다';
      edit.addEventListener('click', () => {
        edit.disabled = true;
        void openEditFromView(options.docId).finally(() => {
          edit.disabled = false;
        });
      });
      return edit;
    }
    if (act === 'download') {
      const link = document.createElement('a');
      link.className = 'icon-btn';
      link.dataset.act = act;
      link.title = label;
      link.setAttribute('aria-label', label);
      link.innerHTML = svgIcon(STRIP_BUTTONS[act].icon);
      link.href = `/api/view/download?token=${encodeURIComponent(options.token)}`;
      return link;
    }
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'icon-btn';
    button.dataset.act = act;
    button.title = label;
    button.setAttribute('aria-label', label);
    button.innerHTML = svgIcon(STRIP_BUTTONS[act].icon);
    // 확대·축소는 상태 줄의 같은 id 단추(#zoom-in·#zoom-out)를 대신 누른다 — 확대 로직은 편집기 것 그대로
    if (act === 'zoom-out' || act === 'zoom-in') button.addEventListener('click', () => document.getElementById(act)?.click());
    if (act === 'print') button.addEventListener('click', () => void printDoc());
    return button;
  }
```

반환 객체(`return { selectTab, setContextualTab(tab, visible) { … }, };`)의 `setContextualTab` 뒤에 더한다.

```ts
    setViewNav(nav) {
      const tools = stripTools;
      if (!tools) return;
      const before = tools.querySelector('[data-act="print"], [data-act="download"], [data-act="edit"]');
      const pairs: Array<['prev' | 'next', string, () => void]> = [
        ['prev', nav.prevLabel, () => nav.prev()],
        ['next', nav.nextLabel, () => nav.next()],
      ];
      for (const [act, label, run] of pairs) {
        const button = stripButton(act, label);
        button.addEventListener('click', run);
        tools.insertBefore(button, before);
      }
    },
```

- [ ] **Step 8: 세 편집기가 옵션 두 개를 넘긴다**

`src/client/editor.ts:708-715`, `src/client/grid-editor.ts:459-466`, `src/client/editor-slides.ts:1314-1321`의 `initOfficeShell({` 객체에서 `docName: info.name,` 줄 바로 아래에 똑같이 더한다.

```ts
    docId: info.docId,
    canCopy: info.permissions.canCopy,
```

(세 편집기의 `ViewInfo`에는 `docId`(editor.ts:145·grid-editor.ts:71·editor-slides.ts:50)와 `permissions.canCopy`(:133·:64·:43)가 이미 있다.)

- [ ] **Step 9: CSS** — `public/css/style.css` 끝에 덧붙인다.

```css
/* 읽기 전용 편집기 보기 꾸밈(ui=view, src/client/viewStrip.ts·officeShell.ts) — 리본·탭·빠른 실행·파일 화면·수식 입력줄·바꾸기를 숨기고
   제목 줄을 얇은 도구 줄로 쓴다. 상태 줄은 "읽기 전용"과 확대 슬라이더만 */
body.ui-view .ribbon-tabs,
body.ui-view .ribbon-panel,
body.ui-view .quick-access,
body.ui-view .backstage,
body.ui-view #formula-input,
body.ui-view .formula-bar .fx-icon,
body.ui-view #status-words,
body.ui-view #status-chars,
body.ui-view #search-replace-row,
body.ui-view #search-replace-one,
body.ui-view #search-replace-all { display: none !important; }
body.ui-view .editor-toolbar.office-shell { padding-bottom: 0; }
.view-strip-tools { display: flex; align-items: center; gap: 2px; flex-wrap: wrap; }
.view-strip-tools .icon-btn { width: 30px; height: 30px; }
.view-strip-tools .icon-btn svg { width: 18px; height: 18px; }
.view-strip-tools .back-to-view { margin-left: 0.35rem; }
body.ui-view.no-copy :is(#editor, #grid-editor, #slides-rail, #slide-canvas-wrap),
body.ui-view.no-copy :is(#editor, #grid-editor, #slides-rail, #slide-canvas-wrap) * { user-select: none; -webkit-user-select: none; }
.view-fallback-notice {
  position: fixed;
  top: 56px;
  left: 50%;
  transform: translateX(-50%);
  z-index: 50;
  padding: 0.6rem 1rem;
  border-radius: 8px;
  background: #323130;
  color: #ffffff;
  font-size: 0.85rem;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
}
@media (max-width: 600px) {
  body.ui-view .office-titlebar { height: auto; min-height: 40px; flex-wrap: wrap; padding: 0.25rem 0.5rem; }
  body.ui-view .titlebar-left,
  body.ui-view .titlebar-right { flex: 1 1 auto; }
  body.ui-view .titlebar-right { flex-wrap: wrap; }
}
```

- [ ] **Step 10: 시험·타입·빌드**

Run: `npx vitest run tests/viewStrip.test.ts && npm run typecheck && npm run build:editor`
Expected: PASS, 타입 오류 0, 번들 빌드 성공.

- [ ] **Step 11: 커밋**

```bash
git add src/client/viewStrip.ts src/client/officeShell.ts src/client/editor.ts src/client/grid-editor.ts src/client/editor-slides.ts public/css/style.css tests/viewStrip.test.ts
git commit -m "보기 꾸밈 2: viewStrip(두 겹 판정·단추 규칙·막을 키·이동·정원 초과)과 officeShell ui-view(리본·탭·파일 화면 숨김, 얇은 줄 확대·인쇄·내려받기·편집 열기, setViewNav)" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 3: 워드 편집기 — 쪽 이동·정원 초과

워드는 이미 `canEdit=false`면 `editable: () => canEdit`(타자·붙여넣기·IME 무시), Ctrl+S 무시(`setupSaveStatus`가 `!canEdit`면 바로 끝), 바꾸기 줄 숨김(`openSearch`의 `withReplace && canEdit`)이다. 보기 꾸밈에서 남은 건 Task 2의 막기(Ctrl+H·Delete·F2·복사)와 이 태스크의 쪽 이동·4429뿐.

**Files:**
- Modify: `src/client/editor.ts` (가져오기, `:654-659` connection-close, `:708-724` 셸 뒤)

**Interfaces:**
- Consumes: `capacityFallback`, `fallBackToViewer`, `isViewUi`, `pickBoundary` (`./viewStrip.js`), `shell.setViewNav(nav: ViewNav)` (Task 2)
- Produces: 얇은 줄 `[data-act=prev]`·`[data-act=next]` = 이전/다음 쪽(첫 `.docx-page-gap`이 도구 줄 바로 아래로)

- [ ] **Step 1: 가져오기** — `import { initOfficeShell } from './officeShell.js';`(:45) 아래에 더한다.

```ts
import { capacityFallback, fallBackToViewer, isViewUi, pickBoundary } from './viewStrip.js';
```

- [ ] **Step 2: 4429 자동 이동** — `:654-659`의 `provider.on('connection-close', …)`를 바꾼다.

```ts
  provider.on('connection-close', (event: CloseEvent | null) => {
    if (event && event.code >= 4400) {
      provider.shouldConnect = false;
      // 보기 꾸밈이 정원 초과(4429)에 걸리면 같은 토큰의 보기 화면으로(설계 §3.4) — 그 밖은 지금처럼 이유를 보여 준다
      const fallback = capacityFallback(event.code, isViewUi(location.search, canEdit), info.viewerUrl);
      if (fallback) fallBackToViewer(fallback);
      else showFatalError(event.reason || '서버가 연결을 거부했습니다.');
    }
  });
```

- [ ] **Step 3: 쪽 이동** — `const shell = initOfficeShell({ … });`(:708-724) 바로 뒤에 넣는다.

```ts
  // 보기 꾸밈의 이전/다음 쪽 — 쪽 사이 간격(docxLayout.ts .docx-page-gap)으로 몸체를 스크롤한다. 본문은 body가 스크롤되고
  // 도구 줄(#toolbar)은 sticky라 그 아래를 기준으로 잰다. 쪽 나눔이 없는 문서(경계 0개)면 이전만 맨 위로 간다.
  shell.setViewNav({
    prevLabel: '이전 쪽',
    nextLabel: '다음 쪽',
    prev: () => jumpPage(-1),
    next: () => jumpPage(1),
  });
  function jumpPage(dir: 1 | -1): void {
    const offset = toolbar.getBoundingClientRect().bottom;
    const tops = [...editorHost.querySelectorAll('.docx-page-gap')].map((gap) => gap.getBoundingClientRect().top - offset);
    const index = pickBoundary(tops, dir);
    if (index >= 0) window.scrollBy(0, tops[index]);
    else if (dir === -1) window.scrollTo(0, 0);
  }
```

(`toolbar`는 :190, `editorHost`는 이 파일의 `#editor` 상수 — 둘 다 이미 있다.)

- [ ] **Step 4: 타입·단위·빌드**

Run: `npm run typecheck && npx vitest run tests/viewStrip.test.ts && npm run build:editor`
Expected: 타입 오류 0, PASS, 빌드 성공.

- [ ] **Step 5: 커밋**

```bash
git add src/client/editor.ts
git commit -m "보기 꾸밈 3: 워드 — 얇은 줄 이전/다음 쪽(쪽 간격 기준), 보기 꾸밈 4429면 안내 뒤 보기 화면으로" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 4: 엑셀 편집기 — 시트 이동·구조 쓰기 막기·정원 초과

엑셀은 이미 `canEdit=false`면 칸 편집(`g-editable` 없음)·Delete·행/열 우클릭 메뉴(:1447 `if (canEdit)`)·바꾸기(:5072·:5135)가 막혀 있다. 시트 탭 우클릭 메뉴(삽입·삭제·이름 바꾸기·이동), 두 번 클릭 이름 바꾸기, 시트 추가 단추, 틀 고정은 `canEdit`을 보지 않고 Yjs `structure`를 직접 써서 — 서버가 버리므로 — 화면만 어긋났다. 수식 입력줄은 Task 2 CSS가 숨기고 이름 상자(`#cell-ref-box`)만 남는다.

**Files:**
- Modify: `src/client/grid-editor.ts` (가져오기, `:339-344`, `:459-469` 셸 뒤, `:799`, `:871`, `:882`, `:905`, `:930`, `:1015`, `:3042`)

**Interfaces:**
- Consumes: `capacityFallback`, `fallBackToViewer`, `isViewUi`, `stepIndex` (`./viewStrip.js`), `shell.setViewNav` (Task 2)
- Produces: 얇은 줄 `[data-act=prev]`·`[data-act=next]` = 이전/다음 시트(`switchToSheet`), `.grid-sheet-tab.active`가 바뀐다

- [ ] **Step 1: 가져오기** — `import { initOfficeShell } from './officeShell.js';`(:27) 아래에 더한다.

```ts
import { capacityFallback, fallBackToViewer, isViewUi, stepIndex } from './viewStrip.js';
```

- [ ] **Step 2: 4429 자동 이동** — `:339-344`를 바꾼다.

```ts
  provider.on('connection-close', (event: CloseEvent | null) => {
    if (event && event.code >= 4400) {
      provider.shouldConnect = false;
      // 보기 꾸밈이 큰 엑셀 정원 초과(4429)에 걸리면 같은 토큰의 보기 화면으로(설계 §3.4) — 편집 화면은 지금처럼 이유를 보여 준다
      const fallback = capacityFallback(event.code, isViewUi(location.search, canEdit), info.viewerUrl);
      if (fallback) fallBackToViewer(fallback);
      else showFatalError(event.reason || '서버가 연결을 거부했습니다.');
    }
  });
```

- [ ] **Step 3: 시트 이동** — `const shell = initOfficeShell({ … });`(:459-469) 바로 뒤에 넣는다.

```ts
  // 보기 꾸밈의 이전/다음 시트 — 시트 탭을 누른 것과 같다(switchToSheet). 끝에서 멈춘다.
  shell.setViewNav({
    prevLabel: '이전 시트',
    nextLabel: '다음 시트',
    prev: () => stepSheet(-1),
    next: () => stepSheet(1),
  });
  function stepSheet(dir: 1 | -1): void {
    const indices = sortedSheetIndices();
    const index = stepIndex(indices.length, indices.indexOf(activeSheet), dir);
    if (index >= 0 && indices[index] !== activeSheet) switchToSheet(indices[index]);
  }
```

(`sortedSheetIndices`(:746)·`switchToSheet`(:758)은 `main` 안의 함수 선언이라 끌어올려지고, `activeSheet`(:496)는 단추를 누를 때 읽으므로 선언 순서에 걸리지 않는다.)

- [ ] **Step 4: 시트 추가 단추** — `:799`를 바꾼다.

```ts
  const sheetAddBtn = document.getElementById('sheet-add');
  // 읽기 전용이면 시트 추가 단추를 보이지 않는다 — 서버가 버리는 구조 변경이라 눌러도 화면만 어긋났다
  if (sheetAddBtn && !canEdit) sheetAddBtn.style.display = 'none';
  sheetAddBtn?.addEventListener('click', () => addSheet());
```

- [ ] **Step 5: 구조 쓰기 막기** — 다섯 함수의 본문 첫 줄에 한 줄씩 넣는다.

`function applySheetOrder(orderedKeys: readonly number[]): void {`(:871) 다음 줄:

```ts
    if (!canEdit) return; // 읽기 전용 — 서버가 버리는 구조 변경을 화면에만 적용하지 않는다(시트 순서·이름이 어긋났다)
```

`function addSheet(afterIdx?: number): void {`(:882), `async function deleteSheet(idx: number): Promise<void> {`(:905), `function startSheetRename(idx: number): void {`(:930), `function openSheetContextMenu(idx: number, x: number, y: number): void {`(:1015) 다음 줄에 각각:

```ts
    if (!canEdit) return;
```

틀 고정(:3042) 처리기의 첫 줄 `if (!activeCellKey) return;`을 바꾼다.

```ts
    if (!canEdit || !activeCellKey) return; // 틀 고정도 structure를 쓴다 — 읽기 전용이면 화면만 어긋났다
```

- [ ] **Step 6: 타입·단위·빌드**

Run: `npm run typecheck && npx vitest run tests/viewStrip.test.ts tests/xlsx-collab.test.ts && npm run build:editor`
Expected: 타입 오류 0, PASS, 빌드 성공.

- [ ] **Step 7: 커밋**

```bash
git add src/client/grid-editor.ts
git commit -m "보기 꾸밈 4: 엑셀 — 얇은 줄 이전/다음 시트, 읽기 전용이면 시트 추가·삭제·이름 바꾸기·순서·틀 고정 쓰기 막음, 보기 꾸밈 4429면 보기 화면으로" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 5: PPT 편집기 — 슬라이드 이동·정원 초과

PPT는 이미 `canEdit=false`면 글상자 `editable` 끔, 섬네일 `draggable=false`·복제/삭제 단추 제거(:818-821), Ctrl+M/D 무시(:1284)다. PPT 편집기에는 찾기·확대 기능이 없어 얇은 줄은 이전/다음·인쇄·내려받기(·편집 열기)만.

**Files:**
- Modify: `src/client/editor-slides.ts` (가져오기, `:938` 근처 새 함수, `:1314` 셸, `:1383-1388`)

**Interfaces:**
- Consumes: `capacityFallback`, `fallBackToViewer`, `isViewUi`, `stepIndex` (`./viewStrip.js`), `shell.setViewNav` (Task 2)
- Produces: 얇은 줄 `[data-act=prev]`·`[data-act=next]` = 이전/다음 슬라이드(`selectSlide`), `#slides-rail .slide-card.active`가 바뀐다

- [ ] **Step 1: 가져오기** — `import { initOfficeShell } from './officeShell.js';`(:10) 아래에 더한다.

```ts
import { capacityFallback, fallBackToViewer, isViewUi, stepIndex } from './viewStrip.js';
```

- [ ] **Step 2: 슬라이드 한 칸 이동 함수** — `function activeSlideIndex(): number { … }`(:938-940) 바로 뒤에 넣는다.

```ts
/** 보기 꾸밈의 이전/다음 단추 — 섬네일을 누른 것과 같다(selectSlide). 끝에서 멈추고, 그 섬네일이 레일에 보이게 한다. */
function stepSlide(dir: 1 | -1): void {
  const list = slidesArray.toArray();
  const index = stepIndex(list.length, activeSlideIndex(), dir);
  if (index < 0) return;
  selectSlide(list[index].id);
  railEl.querySelector<HTMLElement>(`.slide-card[data-id="${CSS.escape(list[index].id)}"]`)?.scrollIntoView({ block: 'nearest' });
}
```

- [ ] **Step 3: 셸 받기와 이동 잇기** — `:1314`의 `initOfficeShell({`를 `const shell = initOfficeShell({`로 바꾸고, 그 호출이 끝나는 `});`(:1328) 바로 뒤에 넣는다.

```ts
  shell.setViewNav({
    prevLabel: '이전 슬라이드',
    nextLabel: '다음 슬라이드',
    prev: () => stepSlide(-1),
    next: () => stepSlide(1),
  });
```

- [ ] **Step 4: 4429 자동 이동** — `:1383-1388`의 `provider.on('connection-close', …)`를 바꾼다.

```ts
  provider.on('connection-close', (event: CloseEvent | null) => {
    if (event && event.code >= 4400) {
      provider.shouldConnect = false;
      // 보기 꾸밈이 정원 초과(4429)에 걸리면 같은 토큰의 보기 화면으로(설계 §3.4) — 그 밖은 지금처럼 이유를 보여 준다
      const fallback = capacityFallback(event.code, isViewUi(location.search, canEdit), info.viewerUrl);
      if (fallback) fallBackToViewer(fallback);
      else showFatalError(event.reason || '서버가 연결을 거부했습니다.');
    }
  });
```

- [ ] **Step 5: 타입·단위·빌드**

Run: `npm run typecheck && npx vitest run tests/viewStrip.test.ts tests/editorSlidesAddTextboxGuard.test.ts && npm run build:editor`
Expected: 타입 오류 0, PASS, 빌드 성공.

- [ ] **Step 6: 커밋**

```bash
git add src/client/editor-slides.ts
git commit -m "보기 꾸밈 5: PPT — 얇은 줄 이전/다음 슬라이드(섬네일과 같은 selectSlide), 보기 꾸밈 4429면 보기 화면으로" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 6: 메인 화면 — 기본 "열기", "▾ 다른 방식", 공유 폼 기본값

**Files:**
- Modify: `public/js/main.js` (`:37-53` issueToken, `:79` 앞 새 함수, `:298-363` 단추, `:365-430` 공유 폼)
- Modify: `public/css/style.css` (끝에 덧붙임)
- Modify: `scripts/verify-artcraft-e2e.mjs:305-320` (메뉴 열기 두 군데)

**Interfaces:**
- Consumes: Task 1 `engine:'editor'` 응답(`viewerUrl`, `fallbackUrl`), Task 2 `EDIT_DOC_KEY = 'docview-edit-doc'`(main.js는 모듈이 아니라 같은 문자열을 상수로 둔다)
- Produces: 문서 줄의 `button` "열기"(`.btn.btn-primary`), `details.open-more > summary`("▾ 다른 방식") 안의 기존 단추들(글자 그대로), 공유 폼 `input[name=engine][value=editor]`(docx/xlsx/pptx면 기본 선택), 결과 링크 글자 "읽기 전용 편집기 화면 열기"

- [ ] **Step 1: issueToken이 editor도 보낸다** — `:49`의 `...(engine === 'artcraft' ? { engine } : {}),`를 바꾸고 위 주석 한 줄을 더한다.

```js
    // engine: 'editor' = 읽기 전용 편집기 화면(보기 꾸밈) — 서버가 canEdit 을 끄고 viewerUrl 을 편집기+ui=view 로 준다
    body: JSON.stringify({
      viewer,
      canPrint,
      canDownload,
      canCopy,
      canEdit: canEdit === true,
      ...(pdfEditor ? { pdfEditor: true } : {}),
      ...(engine !== 'libreoffice' ? { engine } : {}),
    }),
```

- [ ] **Step 2: 읽기 전용 편집기로 여는 함수** — `/** 업로더 본인 편집용 …` 주석(:78) 바로 위에 넣는다.

```js
/** 보기 꾸밈 "편집 열기"가 같은 문서인지 보는 sessionStorage 키 — src/client/viewStrip.ts EDIT_DOC_KEY와 같은 값 */
const EDIT_DOC_KEY = 'docview-edit-doc';

/** 기본 "열기"(docx/xlsx/pptx) — 편집기 화면 그대로, 편집 도구 없이(설계 readonly-editor-look §3.3). 같은 탭으로 가므로
 * sessionStorage 의 문서 id 가 편집기 화면에 남아 "편집 열기" 단추가 보인다(공유 링크로 연 사람에게는 없다). */
async function openReadonlyEditor(docId, viewer) {
  const issued = await issueToken(docId, viewer, true, true, true, false, false, 'editor');
  try {
    sessionStorage.setItem(EDIT_DOC_KEY, docId);
  } catch {
    /* 저장소를 못 쓰면 "편집 열기" 단추만 안 보인다 */
  }
  location.href = issued.viewerUrl;
}
```

- [ ] **Step 3: 기본 "열기"와 "▾ 다른 방식" 메뉴** — `renderDocItem`에서 `const quickView = document.createElement('button');`(:298) 줄 **바로 위**에 넣는다.

```js
  // 기본 "열기" — docx/xlsx/pptx 는 읽기 전용 편집기 화면, 그 밖은 지금 바로 보기와 같다(pdf 는 PDF 편집기).
  // 예전 단추들은 글자·동작 그대로 "▾ 다른 방식" 메뉴 안으로 옮긴다(e2e 가 글자로 찾는다).
  const editorLook = supportsEdit(doc);
  const openBtn = document.createElement('button');
  openBtn.type = 'button';
  openBtn.className = 'btn btn-primary';
  openBtn.textContent = '열기';
  openBtn.title = editorLook ? '편집기와 같은 화면으로 엽니다 — 편집 도구 없이 읽기만' : '';
  openBtn.addEventListener('click', () => {
    openBtn.disabled = true;
    (editorLook ? openReadonlyEditor(doc.id, '업로더') : openViewer(doc.id, '업로더', doc.ext)).catch((error) => {
      openBtn.disabled = false;
      setStatus(error.message, 'error');
    });
  });
  const more = document.createElement('details');
  more.className = 'open-more';
  const moreSummary = document.createElement('summary');
  moreSummary.className = 'btn';
  moreSummary.textContent = '▾ 다른 방식';
  const moreMenu = document.createElement('div');
  moreMenu.className = 'open-more-menu';
  more.append(moreSummary, moreMenu);
  info.append(openBtn, more);
```

이어서 기존 코드의 네 군데를 바꾼다.

`:309-310`

```js
  const artOk = supportsArtcraft(doc);
  if (artOk) quickView.textContent = '리브레오피스로 열기';
  info.append(quickView);
```

→

```js
  const artOk = supportsArtcraft(doc);
  if (artOk || editorLook) quickView.textContent = '리브레오피스로 열기';
  if (artOk || editorLook) quickView.title = '서버 변환';
  moreMenu.append(quickView);
```

`:334` `info.append(artWrap);` → `moreMenu.append(artWrap);`
`:348` `info.append(quickEdit);` → `moreMenu.append(quickEdit);`
`:362` `info.append(pdfEdit);` → `moreMenu.append(pdfEdit);`

(`info.append(makeDeleteButton());`(:365)는 그대로 — 삭제는 메뉴 밖.)

- [ ] **Step 4: 공유 폼 라디오** — `form.innerHTML`(:368-381)의 `<span class="engine-choice" …>…</span>`를 바꾼다.

```js
    <span class="engine-choice" role="radiogroup" aria-label="여는 엔진">
      ${supportsEdit(doc) ? '<label title="편집기와 같은 화면, 편집 도구 없이 — 저장 권한이 있어야 합니다"><input type="radio" name="engine" value="editor" checked /> 읽기 전용 편집기 화면(기본)</label>' : ''}
      <label><input type="radio" name="engine" value="libreoffice" ${supportsEdit(doc) ? '' : 'checked'} /> 리브레오피스</label>
      <label title="${artOk ? '브라우저 안에서 그립니다 — 저장 권한이 있어야 합니다' : ARTCRAFT_UNSUPPORTED}"><input type="radio" name="engine" value="artcraft" ${artOk ? '' : 'disabled'} /> 아트크래프트</label>
    </span>
```

`const result = document.createElement('p');`(:383) 바로 위에 넣는다.

```js
  // "동시편집"을 켜면 엔진 라디오는 끈다 — 편집 토큰은 지금처럼 편집기(editorUrl)로 열린다(꺼진 라디오는 FormData 에 없어 libreoffice)
  const editBox = form.querySelector('input[name=canEdit]');
  if (editBox && supportsEdit(doc)) {
    editBox.addEventListener('change', () => {
      for (const radio of form.querySelectorAll('input[name=engine]')) radio.disabled = editBox.checked || (radio.value === 'artcraft' && !artOk);
    });
  }
```

제출 처리기에서 엔진 인자(:396)를 바꾼다.

```js
        ['editor', 'artcraft'].includes(data.get('engine')) ? data.get('engine') : 'libreoffice',
```

결과 링크 글자(:403)를 바꾼다.

```js
      open.textContent =
        issued.engine === 'editor' ? '읽기 전용 편집기 화면 열기' : issued.engine === 'artcraft' ? '아트크래프트로 열기' : artOk ? '리브레오피스로 열기' : '뷰어 열기';
```

`if (issued.editorUrl) {`(:412)를 바꾼다(보기 꾸밈 링크와 나란히 읽기 전용 편집기 링크를 또 주면 헷갈린다).

```js
      if (issued.editorUrl && issued.engine !== 'editor') {
```

링크 복사(:427)의 주소를 바꾼다.

```js
        const link = issued.engine === 'editor' ? issued.viewerUrl : issued.editorUrl || issued.viewerUrl;
```

- [ ] **Step 5: 메뉴 CSS** — `public/css/style.css` 끝에 덧붙인다.

```css
/* 메인 화면 "▾ 다른 방식" — 기본 "열기" 옆, 예전 여는 단추들(리브레오피스·아트크래프트·동시편집·PDF 편집)을 담는 펼침 메뉴 */
.open-more { position: relative; display: inline-block; }
.open-more > summary { list-style: none; cursor: pointer; user-select: none; }
.open-more > summary::-webkit-details-marker { display: none; }
.open-more-menu {
  position: absolute;
  top: calc(100% + 4px);
  left: 0;
  z-index: 20;
  display: flex;
  flex-direction: column;
  gap: 4px;
  min-width: 190px;
  padding: 6px;
  border: 1px solid var(--color-line);
  border-radius: 8px;
  background: var(--color-surface);
  box-shadow: 0 6px 18px rgba(0, 0, 0, 0.12);
}
.open-more-menu .btn { width: 100%; text-align: left; }
.open-more-menu > span { display: block; }
```

- [ ] **Step 6: 아트크래프트 e2e가 메뉴를 열고 단추를 누르게** — `scripts/verify-artcraft-e2e.mjs`

`const pptButton = pptRow.getByRole('button', { name: '아트크래프트로 열기' });`(:306) 바로 위에 넣는다.

```js
    // 여는 단추들은 "▾ 다른 방식" 메뉴 안으로 옮겨졌다(readonly-editor-look §3.3) — 닫힌 메뉴 안은 접근성 트리에 없어 먼저 연다
    const openMore = (row) => row.locator('details.open-more').evaluate((d) => { d.open = true; });
    await openMore(docRow);
    await openMore(pptRow);
```

`await docRow.getByRole('button', { name: '아트크래프트로 열기' }).click();`(:320) 바로 위에 넣는다.

```js
    await openMore(docRow);
```

- [ ] **Step 7: 문법·단위 확인**

Run: `node --check public/js/main.js && node --check scripts/verify-artcraft-e2e.mjs && npm run typecheck`
Expected: 오류 없음.

- [ ] **Step 8: 커밋** — `scripts/verify-artcraft-e2e.mjs`가 아직 추적되지 않는 파일이면(`git ls-files --error-unmatch scripts/verify-artcraft-e2e.mjs`가 실패) 아트크래프트 작업 몫이므로 여기서 커밋하지 말고 보고에 적는다.

```bash
git add public/js/main.js public/css/style.css
git ls-files --error-unmatch scripts/verify-artcraft-e2e.mjs && git add scripts/verify-artcraft-e2e.mjs
git commit -m "보기 꾸밈 6: 메인 화면 기본 '열기'=읽기 전용 편집기 화면, 예전 단추는 '▾ 다른 방식' 메뉴로(글자·동작 그대로), 공유 폼 '읽기 전용 편집기 화면(기본)' 라디오, 편집 열기용 sessionStorage" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 7: 브라우저 e2e와 회귀 묶음

**Files:**
- Create: `scripts/verify-readonly-editor-e2e.mjs`
- Modify: `package.json` (scripts에 `verify:readonly`)

**Interfaces:**
- Consumes: Task 1~6 전부 — 토큰 `engine:'editor'`/`fallbackUrl`, `body.ui-view`, `.view-strip-tools [data-act=…]`, `.view-fallback-notice`, 메인 화면 "열기"·`details.open-more`·`input[name=engine][value=editor]`·"읽기 전용 편집기 화면 열기"
- Produces: `npm run verify:readonly [포트]` — 설계 §4 (1)~(8) 통과 여부, 실패 1개라도 있으면 종료 코드 1

- [ ] **Step 1: e2e 스크립트를 쓴다** — `scripts/verify-readonly-editor-e2e.mjs`

```js
// 기본 열기 = 읽기 전용 편집기 화면(ui=view) — 설계 2026-11-21-readonly-editor-look §4의 브라우저 검사 (1)~(8). (9) 기존 e2e는 따로 돌린다.
//  (1) 리본·탭·빠른 실행·파일 화면·수식 입력줄 숨김, 얇은 줄 단추  (2) 편집 시도가 화면·서버(새로 연 화면)를 바꾸지 않음
//  (3) 찾기·확대·이동·섬네일·인쇄·내려받기  (4) canCopy=false면 복사·우클릭·선택 막힘  (5) 다른 탭의 편집이 실시간 반영
//  (6) 메인 화면 기본 "열기"·"▾ 다른 방식" 세 항목·공유 폼 기본값·편집 열기  (7) 600px 폭  (8) 정원 초과(4429) → 안내 + 자동 이동
//  (8)은 큰 엑셀 네 개로 정원을 실제로 채우는 대신 Playwright routeWebSocket으로 서버가 4429로 닫은 것과 같은 상황을 만든다.
// 이 스크립트가 시험 서버를 직접 띄운다(npm run build:editor 뒤, LibreOffice 필요):
//   node scripts/verify-readonly-editor-e2e.mjs [포트=8152, 8150~8159]
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';

const PORT = Number(process.argv[2] ?? 8152);
// 운영 포트(8000·8010·8020·8100·8110·8130·8148)와 겹치지 않는 시험 전용 범위만
if (!(PORT >= 8150 && PORT <= 8159)) throw new Error(`시험 포트는 8150~8159만 쓴다(받은 값 ${PORT})`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `readonly-editor-e2e-${PORT}`);
// 쪽 이동 검사에는 쪽 나눔이 있는 문서가 필요하다 — a1에 .docx-page-gap이 없으면 d05-header-footer-sections.docx로 바꾼다
const FILES = {
  docx: 'audit/office/corpus/docx/a1-korean-report.docx',
  xlsx: 'audit/office/corpus/xlsx/x07-multisheet-print.xlsx',
  pptx: 'audit/office/corpus/pptx/a2-korean-deck.pptx',
};
const CAPACITY_REASON = '사용자가 많습니다. 잠시 후 다시 시도해주세요.';
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

if (!fs.existsSync('public/js/editor.bundle.js')) {
  console.error('public/js/editor.bundle.js 없음 — npm run build:editor 먼저');
  process.exit(1);
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
  // 내가 띄운 서버만(PID, 자식 포함) — 이름으로 node를 끄지 않는다
  if (server?.pid) execFileSync('taskkill', ['/PID', String(server.pid), '/T', '/F'], { stdio: 'ignore' });
  server = null;
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
    body: JSON.stringify({ viewer: 'readonly-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);

/** 편집기마다: 준비 표시, 누를 곳, 문서 지문(화면·서버 상태 비교용 — 페이지 안에서 돈다) */
const ED = {
  docx: {
    ready: '.ProseMirror',
    surface: '.ProseMirror',
    snap: () => document.querySelector('.ProseMirror')?.textContent ?? '',
  },
  xlsx: {
    ready: '#grid-edit-table td[data-row]',
    surface: '#grid-edit-table td[data-row]',
    snap: () =>
      [...document.querySelectorAll('.grid-sheet-tab')].map((t) => t.textContent).join('|') +
      '#' +
      [...document.querySelectorAll('#grid-edit-table td[data-row]')].map((td) => td.textContent).join('|'),
  },
  pptx: {
    ready: '#slides-rail .slide-card',
    surface: '#slide-canvas',
    snap: () =>
      [...document.querySelectorAll('#slides-rail .slide-card')].map((c) => c.dataset.id).join('|') + '#' + (document.querySelector('#slide-canvas')?.textContent ?? ''),
  },
};
const WANT_ACTS = {
  docx: ['zoom-out', 'zoom-in', 'prev', 'next', 'print', 'download'],
  xlsx: ['zoom-out', 'zoom-in', 'prev', 'next', 'print', 'download'],
  pptx: ['prev', 'next', 'print', 'download'],
};

async function openView(context, url, ext) {
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.goto(`${BASE}${url}`);
  await page.waitForSelector(ED[ext].ready, { timeout: 90000 });
  await page.waitForFunction(() => document.querySelector('#conn-status')?.getAttribute('data-status') === 'synced', null, { timeout: 30000 }).catch(() => {});
  await sleep(1500);
  return { page, errors };
}
const contextMenuPrevented = (page, sel) =>
  page.evaluate((s) => {
    const ev = new MouseEvent('contextmenu', { bubbles: true, cancelable: true });
    document.querySelector(s).dispatchEvent(ev);
    return ev.defaultPrevented;
  }, sel);

await startServer();
let browser;
try {
  const ids = {};
  for (const [ext, file] of Object.entries(FILES)) ids[ext] = (await upload(file)).id;
  browser = await chromium.launch();

  for (const ext of ['docx', 'xlsx', 'pptx']) {
    const id = ids[ext];
    const tok = await token(id, { engine: 'editor' });
    rec(`${ext} 토큰: engine=editor · ui=view 주소 · fallbackUrl`, tok.engine === 'editor' && /&ui=view$/.test(tok.viewerUrl) && typeof tok.fallbackUrl === 'string', tok.viewerUrl);
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] });
    const { page, errors } = await openView(context, tok.viewerUrl, ext);

    // (1) 숨김·얇은 줄
    const shown = await page.evaluate(() => {
      const none = (sel) => [...document.querySelectorAll(sel)].every((el) => getComputedStyle(el).display === 'none');
      return {
        body: document.body.classList.contains('ui-view'),
        tabs: none('.ribbon-tabs'),
        panels: none('.ribbon-panel'),
        qa: none('.quick-access'),
        backstage: none('.backstage'),
        formula: none('#formula-input'),
        acts: [...document.querySelectorAll('.view-strip-tools [data-act]')].map((b) => b.dataset.act),
      };
    });
    rec(`(1) ${ext} 리본·탭·빠른 실행·파일 화면·수식 입력줄 숨김`, shown.body && shown.tabs && shown.panels && shown.qa && shown.backstage && shown.formula, JSON.stringify(shown));
    rec(`(1) ${ext} 얇은 줄 단추(API로 연 화면엔 편집 열기 없음)`, WANT_ACTS[ext].every((a) => shown.acts.includes(a)) && !shown.acts.includes('edit'), shown.acts.join(','));

    // (2) 편집 시도
    const before = await page.evaluate(ED[ext].snap);
    await page.click(ED[ext].surface);
    await page.keyboard.type('ZZQ');
    for (const key of ['Delete', 'Backspace', 'F2', 'Control+v', 'Control+s', 'Control+h']) await page.keyboard.press(key);
    rec(`(2) ${ext} Ctrl+H — 바꾸기 줄 안 보임`, !(await page.locator('#search-replace-row').isVisible()));
    await page.keyboard.press('Escape');
    if (ext === 'xlsx') {
      const tab = page.locator('.grid-sheet-tab').first();
      await tab.click({ button: 'right' });
      rec('(2) xlsx 시트 탭 우클릭 메뉴 없음', (await page.locator('.sheet-context-menu').count()) === 0);
      await tab.dblclick();
      rec('(2) xlsx 시트 탭 두 번 클릭 — 이름 바꾸기 입력창 없음', (await page.locator('.grid-sheet-tab-rename').count()) === 0);
      rec('(2) xlsx 시트 추가 단추 안 보임', !(await page.locator('#sheet-add').isVisible()));
      await page.locator(ED.xlsx.surface).first().click({ button: 'right' });
      rec('(2) xlsx 칸 우클릭 메뉴 없음', (await page.locator('.grid-context-menu').count()) === 0);
    }
    if (ext === 'pptx') {
      const cards = page.locator('#slides-rail .slide-card');
      if ((await cards.count()) >= 2) await cards.nth(0).dragTo(cards.nth(1));
      rec('(2) pptx 섬네일 끌기 꺼짐·복제/삭제 단추 없음', (await page.locator('.slide-card-actions').count()) === 0 && (await cards.first().getAttribute('draggable')) === 'false');
    }
    await sleep(1500);
    const after = await page.evaluate(ED[ext].snap);
    const fresh = await openView(context, (await token(id, { engine: 'editor' })).viewerUrl, ext);
    const server = await fresh.page.evaluate(ED[ext].snap);
    await fresh.page.close();
    rec(`(2) ${ext} 편집 시도 뒤 화면·서버(새로 연 화면) 그대로`, after === before && server === before && !after.includes('ZZQ'), `${before.length}자`);

    // (3) 찾기·확대·이동·인쇄·내려받기
    if (ext !== 'pptx') {
      const body = ext === 'xlsx' ? before.split('#').slice(1).join('#') : before;
      const term = body.split(/[\s|#]+/).find((w) => w.length >= 2) ?? '';
      await page.click('#search-btn');
      await page.fill('#search-input', term);
      await page.keyboard.press('Enter');
      await sleep(800);
      const count = await page.textContent('#search-count');
      rec(`(3) ${ext} 찾기`, (await page.isVisible('#search-bar')) && /[1-9]/.test(count ?? ''), `${term} → ${count}`);
      await page.keyboard.press('Escape');
      const z0 = await page.textContent('#zoom-level');
      await page.click('.view-strip-tools [data-act=zoom-in]');
      await sleep(300);
      const z1 = await page.textContent('#zoom-level');
      rec(`(3) ${ext} 얇은 줄 확대`, z0 !== z1, `${z0} → ${z1}`);
      await page.click('.view-strip-tools [data-act=zoom-out]');
    }
    if (ext === 'docx') {
      await page.evaluate(() => window.scrollTo(0, 0));
      await page.click('.view-strip-tools [data-act=next]');
      await sleep(500);
      const gap = await page.evaluate(() => {
        const g = document.querySelector('.docx-page-gap');
        const bar = document.getElementById('toolbar').getBoundingClientRect().bottom;
        return g ? Math.round(g.getBoundingClientRect().top - bar) : null;
      });
      rec('(3) docx 다음 쪽 — 첫 쪽 간격이 도구 줄 바로 아래로', gap !== null && Math.abs(gap) <= 4, String(gap));
    }
    if (ext === 'xlsx') {
      const t0 = await page.textContent('.grid-sheet-tab.active');
      await page.click('.view-strip-tools [data-act=next]');
      await sleep(800);
      const t1 = await page.textContent('.grid-sheet-tab.active');
      rec('(3) xlsx 다음 시트', t0 !== t1, `${t0} → ${t1}`);
      await page.locator('.grid-sheet-tab').first().click();
      await sleep(800);
    }
    if (ext === 'pptx') {
      const num = () => page.textContent('#slides-rail .slide-card.active .slide-card-number');
      const n0 = await num();
      await page.click('.view-strip-tools [data-act=next]');
      await sleep(500);
      const n1 = await num();
      rec('(3) pptx 다음 슬라이드', n0 !== n1, `${n0} → ${n1}`);
      await page.locator('#slides-rail .slide-card').first().click();
      await sleep(500);
      rec('(3) pptx 섬네일로 이동', (await num()) === '1');
    }
    const [popup] = await Promise.all([context.waitForEvent('page', { timeout: 15000 }), page.click('.view-strip-tools [data-act=print]')]);
    await popup.waitForURL(/\/api\/view\/print\?token=/, { timeout: 15000 }).catch(() => {});
    rec(`(3) ${ext} 인쇄 — 새 창이 인쇄 주소로`, /\/api\/view\/print\?token=/.test(popup.url()), popup.url());
    await popup.close();
    const [download] = await Promise.all([page.waitForEvent('download', { timeout: 60000 }), page.click('.view-strip-tools [data-act=download]')]);
    rec(`(3) ${ext} 내려받기`, download.suggestedFilename().length > 0, download.suggestedFilename());

    // (4) 복사 권한
    rec(`(4) ${ext} canCopy=true — 브라우저 우클릭 메뉴는 그대로`, !(await contextMenuPrevented(page, ED[ext].surface)));
    {
      const ctx2 = await browser.newContext({ viewport: { width: 1280, height: 900 }, permissions: ['clipboard-read', 'clipboard-write'] });
      const noCopy = await openView(ctx2, (await token(id, { engine: 'editor', canCopy: false })).viewerUrl, ext);
      const p = noCopy.page;
      await p.evaluate(() => navigator.clipboard.writeText('SENTINEL'));
      await p.click(ED[ext].surface);
      await p.keyboard.press('Control+a');
      await p.keyboard.press('Control+c');
      await sleep(300);
      const clip = await p.evaluate(() => navigator.clipboard.readText());
      const prevented = await contextMenuPrevented(p, ED[ext].surface);
      const userSelect = await p.evaluate((s) => getComputedStyle(document.querySelector(s)).userSelect, ED[ext].surface);
      rec(`(4) ${ext} canCopy=false — 복사·우클릭·선택 막힘`, clip === 'SENTINEL' && prevented && userSelect === 'none', `${clip.slice(0, 20)} · ${prevented} · ${userSelect}`);
      await ctx2.close();
    }

    // (5) 실시간 반영 — 편집 토큰으로 연 다른 탭에서 고친다
    {
      const edit = await openView(context, (await token(id, { canEdit: true })).editorUrl, ext);
      const mark = `LIVE${Date.now() % 100000}`;
      if (ext === 'docx') {
        await edit.page.click('.ProseMirror');
        await edit.page.keyboard.press('Control+End');
        await edit.page.keyboard.type(mark);
      } else if (ext === 'xlsx') {
        await edit.page.locator(ED.xlsx.surface).first().click();
        await edit.page.keyboard.type(mark);
        await edit.page.keyboard.press('Enter');
      } else {
        await edit.page.locator('#slide-canvas .slide-textbox').first().click();
        await edit.page.keyboard.press('End');
        await edit.page.keyboard.type(mark);
      }
      const seen = await page.waitForFunction((m) => document.body.innerText.includes(m), mark, { timeout: 20000 }).then(() => true, () => false);
      rec(`(5) ${ext} 다른 탭의 편집이 보기 꾸밈 화면에 실시간`, seen, mark);
      await edit.page.close();
    }

    rec(`${ext} 화면 오류 없음`, errors.length === 0, errors.slice(0, 2).join(' | '));
    await context.close();
  }

  // (6) 메인 화면
  {
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const page = await context.newPage();
    const row = () => page.locator('li.doc-item', { hasText: path.basename(FILES.docx) });
    const home = async () => {
      await page.goto(`${BASE}/`);
      await row().waitFor({ timeout: 30000 });
    };
    await home();
    await row().getByRole('button', { name: '열기', exact: true }).click();
    await page.waitForURL(/\/editor\.html\?token=[^&]+&ui=view$/, { timeout: 30000 });
    const editBtn = await page.waitForSelector('.view-strip-tools [data-act=edit]', { timeout: 60000 }).then(() => true, () => false);
    rec('(6) 메인 기본 "열기" → 읽기 전용 편집기 화면 + 편집 열기 단추', editBtn && (await page.evaluate(() => document.body.classList.contains('ui-view'))), page.url());
    await page.click('.view-strip-tools [data-act=edit]');
    await page.waitForURL((u) => u.pathname === '/editor.html' && !u.search.includes('ui=view'), { timeout: 30000 });
    await page.waitForSelector('.ProseMirror', { timeout: 60000 });
    rec('(6) 편집 열기 → 편집 권한 편집기(리본 보임)', !(await page.evaluate(() => document.body.classList.contains('ui-view'))) && (await page.isVisible('.ribbon-tabs')));
    for (const [label, re] of [
      ['리브레오피스로 열기', /\/view\.html\?token=/],
      ['아트크래프트로 열기', /\/view\.html\?token=/],
      ['동시편집 열기', /\/editor\.html\?token=[^&]+$/],
    ]) {
      await home();
      await row().locator('details.open-more > summary').click();
      await row().getByRole('button', { name: label }).click();
      await page.waitForURL(re, { timeout: 30000 }).catch(() => {});
      rec(`(6) "▾ 다른 방식" → ${label}`, re.test(page.url()), page.url());
    }
    await home();
    const form = row().locator('form.token-form');
    rec('(6) 공유 폼 — "읽기 전용 편집기 화면(기본)"이 기본 선택', await form.locator('input[name=engine][value=editor]').isChecked());
    await form.locator('input[name=canEdit]').check();
    rec('(6) 공유 폼 — 동시편집을 켜면 엔진 라디오 꺼짐', await form.locator('input[name=engine][value=editor]').isDisabled());
    await form.locator('input[name=canEdit]').uncheck();
    await form.locator('input[name=viewer]').fill('공유받는 사람');
    await form.locator('input[name=canDownload]').check();
    await form.getByRole('button', { name: '열람 링크 발급' }).click();
    const link = row().locator('.token-result a', { hasText: '읽기 전용 편집기 화면 열기' });
    await link.waitFor({ timeout: 10000 });
    const href = (await link.getAttribute('href')) ?? '';
    rec('(6) 공유 링크 = 보기 꾸밈 주소', /&ui=view$/.test(href), href);
    const other = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const shared = await openView(other, href, 'docx');
    rec('(6) 공유 링크로 연 사람에게는 편집 열기 단추 없음', (await shared.page.locator('[data-act=edit]').count()) === 0);
    await other.close();
    await context.close();
  }

  // (7) 600px 폭
  for (const ext of ['docx', 'xlsx', 'pptx']) {
    const context = await browser.newContext({ viewport: { width: 600, height: 900 } });
    const { page } = await openView(context, (await token(ids[ext], { engine: 'editor' })).viewerUrl, ext);
    const m = await page.evaluate(() => ({
      scroll: document.documentElement.scrollWidth,
      inside: [...document.querySelectorAll('.view-strip-tools [data-act]')].map((b) => {
        const r = b.getBoundingClientRect();
        return r.width > 0 && r.right <= window.innerWidth + 0.5;
      }),
    }));
    rec(`(7) ${ext} 600px — 가로 넘침 없음, 얇은 줄 단추 모두 화면 안`, m.scroll <= 600 && m.inside.length > 0 && m.inside.every(Boolean), JSON.stringify(m));
    await context.close();
  }

  // (8) 정원 초과(4429)
  {
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const page = await context.newPage();
    await page.routeWebSocket(/\/ws\/edit/, (ws) => ws.close({ code: 4429, reason: CAPACITY_REASON }));
    const tok = await token(ids.xlsx, { engine: 'editor' });
    await page.goto(`${BASE}${tok.viewerUrl}`);
    const notice = await page.waitForSelector('.view-fallback-notice', { timeout: 30000 }).then((el) => el.textContent(), () => null);
    await page.waitForURL(/\/viewer\.html\?token=/, { timeout: 15000 }).catch(() => {});
    rec('(8) 정원 초과 → 안내 + 같은 토큰의 보기 화면으로 자동 이동', notice === '지금은 보는 사람이 많습니다. 서버 변환으로 보여 줍니다' && page.url().endsWith(tok.fallbackUrl), `${notice} · ${page.url()}`);
    const page2 = await context.newPage();
    await page2.routeWebSocket(/\/ws\/edit/, (ws) => ws.close({ code: 4429, reason: CAPACITY_REASON }));
    await page2.goto(`${BASE}${(await token(ids.xlsx, { canEdit: true })).editorUrl}`);
    const shownError = await page2.waitForSelector('#error-screen:not([hidden])', { timeout: 30000 }).then(() => true, () => false);
    rec('(8) 편집 화면은 4429면 지금처럼 오류 화면(자동 이동 없음)', shownError && page2.url().includes('/editor-grid.html'));
    await context.close();
  }
} finally {
  await browser?.close();
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed ? 1 : 0);
```

- [ ] **Step 2: `package.json`에 스크립트** — `"verify:docx": …` 줄 아래에 더한다.

```json
    "verify:readonly": "node scripts/verify-readonly-editor-e2e.mjs",
```

- [ ] **Step 3: 빌드하고 e2e를 돌린다**

Run: `npm run build:editor && npm run verify:readonly -- 8152`
Expected: 마지막 줄 `N/N 통과`, FAIL 0. 실패하면 그 줄의 detail을 보고 해당 태스크 코드를 고친다(docx 쪽 이동이 `null`이면 FILES.docx를 `d05-header-footer-sections.docx`로 바꾼다 — 문서에 쪽 나눔이 없다는 뜻).

- [ ] **Step 4: 실제 브라우저로 직접 눌러 본다**(사용자 규칙: UI는 실제 브라우저로 확인) — 시험 서버를 8153에 띄우고(아래 Step 5의 명령) Chrome으로 `http://127.0.0.1:8153/`에서 문서를 올려 "열기" → 얇은 줄의 확대·축소·이전·다음·인쇄·내려받기·편집 열기, 찾기, 타자·Delete·붙여넣기·우클릭(엑셀 시트 탭·칸, PPT 섬네일 끌기)을 세 형식 모두 직접 해 보고, 창 폭을 600px로 줄여 본다. 이상하면 고치고 Step 3부터 다시.

- [ ] **Step 5: 회귀 묶음(설계 §4 (9))**

```bash
npm test
npm run typecheck
```

Expected: 전부 PASS, 타입 오류 0.

기존 e2e들은 떠 있는 서버 주소를 받는다 — 시험 서버를 직접 띄운다(Bash, `run_in_background: true`).

```bash
mkdir -p .audit-run/regress-8153/tmp
DOCVIEW_PORT=8153 DOCVIEW_HOST=127.0.0.1 DOCVIEW_SECRET=regress-$(date +%s)-secret DOCVIEW_DATA_DIR=.audit-run/regress-8153/data TMP=.audit-run/regress-8153/tmp TEMP=.audit-run/regress-8153/tmp DOCVIEW_LO_PROFILE_ROOT=.audit-run/regress-8153/lo node node_modules/tsx/dist/cli.mjs src/server.ts
```

뜬 뒤(`curl -s http://127.0.0.1:8153/api/formats`가 응답) 차례로:

```bash
npm run verify:docx -- http://127.0.0.1:8153
npm run verify:pptx -- http://127.0.0.1:8153
npm run verify:xlsx:browser -- http://127.0.0.1:8153
npm run verify:view-only -- http://127.0.0.1:8153
npm run verify:artcraft -- 8154
```

Expected: 다섯 개 모두 FAIL 0(`verify:artcraft`는 자기 서버를 8154에 띄우며 `npm run build:artcraft` 산출물이 필요하다 — 없으면 그 사실을 보고한다). 끝나면 8153 서버는 띄운 PID로만 끈다: `netstat -ano | findstr :8153`으로 PID를 보고 `taskkill /PID <그 PID> /T /F`(이름으로 끄지 않는다).

- [ ] **Step 6: 커밋** — `package.json`에 이 태스크 줄 말고 아직 커밋 안 된 다른 줄(예: 아트크래프트 `verify:artcraft`)이 있으면 그 줄은 아트크래프트 작업 몫이므로 `git add -p package.json`으로 `verify:readonly` 줄만 담는다.

```bash
git add scripts/verify-readonly-editor-e2e.mjs
git add -p package.json
git commit -m "보기 꾸밈 7: 브라우저 e2e verify:readonly — 숨김·편집 시도·찾기/확대/이동/인쇄/내려받기·복사 권한·실시간·메인 화면·600px·정원 초과" -m "Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

## 설계 대비 확인(자체 검토)

| 설계 | 태스크 |
|---|---|
| §3.1 리본·탭·빠른 실행·백스테이지 숨김 | Task 2 CSS·`body.ui-view` |
| §3.1 수식 입력줄 숨김, 이름 상자만 | Task 2 CSS(`#formula-input`·`.fx-icon`) |
| §3.1 얇은 줄(문서 이름·찾기·확대·이동·인쇄·내려받기·참석자·편집 열기) | Task 2(제목 줄의 이름·찾기·참석자 + 단추), Task 3~5(이동). PPT는 찾기·확대 기능이 없어 빠짐 |
| §3.1 상태 줄 "읽기 전용" + 확대 슬라이더 | Task 2(`#statusbar` 앞 "읽기 전용", 워드 단어·문자 수 숨김), 엑셀은 원래 `#grid-status-text`가 "읽기 전용" |
| §3.1 메모 창·섬네일 그대로 | 손대지 않음 |
| §3.1 우클릭 메뉴 없음·편집 단축키 무시 | Task 2 `installViewGuards`·`blockedViewKey` |
| §3.1 시트 추가/삭제/이름/틀 고정 쓰기 막음 | Task 4 |
| §3.1 canCopy=false 복사·선택 막음 | Task 2(`no-copy` CSS + copy·Ctrl+C·우클릭) |
| §3.2 engine:'editor'·canEdit 강제·viewerUrl·fallbackUrl·내려받기 없으면 서버 변환 | Task 1 |
| §3.3 기본 "열기"·"▾ 다른 방식"·공유 폼 라디오·편집 열기(sessionStorage, 누를 때 재발급) | Task 6(메인), Task 2(단추·재발급) |
| §3.4 4429 → 안내 + 자동 이동 | Task 1(`fallbackUrl`), Task 2(`capacityFallback`·`fallBackToViewer`), Task 3~5(잇기). 편집기는 같은 토큰의 `info.viewerUrl`을 쓴다(응답 `fallbackUrl`과 같은 주소) |
| §4 단위 | Task 1·2 |
| §4 e2e (1)~(8), (9) | Task 7 |
