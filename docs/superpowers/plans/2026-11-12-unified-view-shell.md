# 통합 보기 껍데기 1단계 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 워드·PPT·PDF(와 그 밖의 쪽 문서, 한글 빼고)를 새 보기 화면 `public/view.html` 하나로 같은 모양으로 보여 준다 — 내려받기 권한이 있으면 브라우저 pdf.js, 없으면 지금처럼 서버 PNG만. 썸네일·찾기·인쇄·내려받기·표시(마크업)·측정·편집 단추와 편집기의 "보기로"까지 옮긴 뒤 `viewerUrl`을 이 화면으로 바꾼다.

**Architecture:** 새 묶음 `view.bundle.js`(입구 `src/client/view.ts`, 기존 `scripts/build-editor.mjs`에 입구 하나 추가). 안쪽은 `PagesBody`(설계 §4 `ViewBody` 구현) 하나가 쪽 틀(`.page-frame` — 지금 보기 화면과 같은 틀·CSS)·확대·회전·가상화·찾기 강조를 맡고, 쪽 그림만 원천(`PageSource`)이 그린다: `PngSource`(서버 `/api/view/pages`·`/api/view/search`) 또는 `PdfjsSource`(새 경로 `/api/view/derived-pdf` + PDF 편집기의 `pdfDocument.ts`·`findText.ts`, 동적 import라 PNG 모드에서는 pdf.js를 받지 않는다). 막대·썸네일·찾기·파일·측정·표시·편집 단추는 `src/client/view/*.ts` 모듈이 `ShellCtx`를 받아 붙는다. 서버는 `/api/view/info`에 `editorUrl`·`viewerUrl`·`processing`을 더하고 `/api/view/derived-pdf`·`/api/view/status`를 새로 둔다. 마지막에 토큰 응답의 `viewerUrl`을 `/view.html`로 바꾸고 `viewer.js`는 쪽 문서를 `/view.html`로 넘긴다(엑셀 격자·도면·한글은 그대로 `viewer.html`).

**Tech Stack:** TypeScript(esbuild 0.24 splitting·esm, vitest 3, `tsc --noEmit`), Fastify 5, pdfjs-dist 4.2.67(이미 있음 — 새 의존성 없음), Playwright(e2e).

**Spec:** `docs/superpowers/specs/2026-11-12-unified-view-shell-design.md` — §0(V1~V8) §1(1단계 줄) §2~§9 전부. 2~4단계(엑셀 격자·한글·DWG/NWD 껍데기 안, 연동 가이드)와 "이후"(자체 그리기 엔진)는 범위 밖. 기준 코드 webviewer master **37eb942**.

## Global Constraints

- **워크트리에서만 일한다:** `C:\Users\OpenMinds\projects\webviewer-view-shell`, 브랜치 `view-shell`(master에서). 주 작업 폴더 `C:\Users\OpenMinds\projects\webviewer`의 `src/`·`public/`은 **한 글자도 고치지 않는다** — 운영 두 서버(8000, 향기랩 8010)가 그 폴더에서 `tsx watch`로 돌아 src를 고치는 순간 재시작되고 모든 세션·토큰이 끊긴다.
- **시험 포트는 임시 포트만**(이 계획은 8124 — 쓰기 전에 `netstat -ano | grep ":8124 "`로 비어 있는지 확인, 차 있으면 8125~8129 중 빈 것). **8000·8010·8020·8100·8110 금지.** 시험 서버는 띄울 때 PID를 적어 두고 **PID로만** 끈다(`taskkill //PID <pid> //T //F`) — 운영 node 프로세스가 떠 있으므로 이름으로 죽이지 않는다.
- 시험 서버 띄우기(워크트리에서, `run_in_background`):
  ```bash
  S="<세션 scratchpad 경로>"
  cd /c/Users/OpenMinds/projects/webviewer-view-shell
  npm run build:editor
  DOCVIEW_PORT=8124 DOCVIEW_HOST=127.0.0.1 DOCVIEW_DATA_DIR="$S/view-shell-data" DOCVIEW_SECRET=e2e-test-secret npx tsx src/server.ts
  ```
  `npx tsx`(watch 아님)로 띄운다. API 키를 설정하지 않으므로 업로드·토큰 API가 키 없이 열린다(기존 verify 스크립트들과 같다).
- `git add`는 파일 이름으로 하나씩, 커밋은 `git commit -m "…" -- <경로들>`. **push·stash·`git add -A` 금지.** 줄 끝 LF(고친 텍스트 파일마다 커밋 전 `git diff --stat`으로 통째 다시 쓰기가 없는지 확인, 있으면 `sed -i 's/\r$//' <파일>`).
- 커밋 메시지는 한국어, 끝 줄은 정확히 `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- 산출물은 커밋하지 않는다: `public/js/view.bundle.js`·`view.bundle.js.map`(Task 3에서 `.gitignore`에 더한다), `public/js/chunks/`, `public/pdfjs/`, `public/js/pdf.worker.mjs`(이미 무시됨).
- UI·오류 문구는 한국어. **사용자에게 보이는 오류에 파일 경로를 넣지 않는다**(서버 오류는 고정 문구, 파일 이름은 됨).
- **권한 규칙(설계 V5):** `canDownload`가 없는 토큰은 어떤 경우에도 `/api/view/derived-pdf`·`/api/view/pdf-source` 요청을 만들지 않는다(브라우저 쪽 `chooseBodyKind` + 서버 403 둘 다). Task 13 e2e가 요청 수 0건을 잰다.
- **연동 제품 회귀 금지:** 전환(Task 12) 전에도 뒤에도 `npm run verify:view-only`·`npm run verify:viewer-quicktext`가 고치지 않은 채로 통과해야 한다. 그래서 새 화면은 지금 보기 화면의 DOM 약속을 그대로 쓴다 — 쪽 틀 `.page-frame`(그린 뒤 `.page-ready`), `#download-btn`·`#text-tool-btn`·`.text-tool-hint`·`.text-tool-ghost`·`.text-insert-input`·`.text-overlay-item`, 1024px 이상에서 도구 단추가 막대에 보임.
- 향기랩 페이지(calc-site `tools/index.html`)는 바꾸지 않는다 — 향기랩은 지금처럼 `editorUrl`을 바로 연다(설계 §0, §8).
- **모바일 성능(사용자 규칙):** 캔버스는 보이는 쪽 + 지금 쪽 ±2만 남기고(`pagesToKeep`), 동시 그리기 3장, 캔버스 한 장 1,677만 화소 이하, 썸네일은 칸에 보이는 것만, pdf.js는 동적 import(PNG 모드에서는 pdf.js 묶음 요청 0건). CPU×4 첫 쪽 시간은 Task 12(지금 화면 기준값)·Task 13(새 화면)에서 숫자로 남긴다.
- **시험 정책(사용자 규칙):** 작업마다 단위 시험·타입 검사·빌드만 — `npx vitest run`(전체), `npx tsc --noEmit`, `npm run build:editor`. 서버 e2e(`verify:*`)·**모든 새 컨트롤의 실제 브라우저 조작**·휴대폰 CPU×4 측정은 **Task 13에서 한 번**.
- **master 병합은 운영 반영이다 — 이 계획에 없다.** Task 13 보고 뒤 사용자 승인을 받고 따로 한다(반영 절차·8010 재시작은 그때).

---

## 설계와 맞춘 이름

| 설계가 쓴 것 | 실제(이 계획이 쓰는 것) |
|---|---|
| `ViewBody`(§4) | `src/client/view/types.ts`의 `ViewBody` 그대로 + 아래 두 차이. 구현은 `PagesBody`(`src/client/view/pagesBody.ts`) 하나 — `kind`는 지금 원천에 따라 `'pages-png'`·`'pages-pdfjs'` |
| `search?(q): AsyncIterable<{ page; rects: DOMRect[] }>` | `search?(q, signal): Promise<SearchHit[]>`, `SearchHit = { page; rects: NormRect[] }`(쪽 대비 0~1 비율 — 서버 `/api/view/search`의 `rects`와 같은 모양). 서버 찾기는 한 번에 다 오고, 화면은 지금처럼 "n / 전체"를 한 번에 보인다 |
| `zoom(z: number \| 'fit-width' \| 'fit-page')` | 그대로. 숫자는 **확대 %**(`ZOOM_STEPS` 50~500, 100 = 쪽 칸 폭 — 지금 보기 화면과 같은 뜻) |
| `pages-png` 안쪽 | `PagesBody` + `PngSource`(`src/client/view/pngSource.ts`) |
| `pages-pdfjs` 안쪽 | `PagesBody` + `PdfjsSource`(`src/client/view/pdfjsSource.ts`, `openPdfjsSource`) |
| PagesBody가 더 내놓는 것 | `container`·`zoomIndex`·`zoomPercent`·`currentPage`·`setZoomIndex`·`zoomAtPoint`·`pageFrame(page)`·`pageSize(page)`·`showHits`·`clearHits`·`onZoom`·`onDraw`·`onDrawError`·`onSourceChange`·`replaceSource` |
| 껍데기 모듈이 함께 쓰는 것 | `ShellCtx`(`src/client/view/ctx.ts`): `token`·`tokenQuery`·`info`·`body`·`onEscape(priority, fn)`·`notice(text)` |
| `GET /api/view/info`(또는 같은 역할의 세션 응답)에 `editorUrl`·`viewerUrl` | `/api/view/info`의 `data.editorUrl`(없으면 생략)·`data.viewerUrl`·`data.processing` |
| `GET /api/view/derived-pdf?token=` | 그대로(`src/routes/viewer.ts`) |
| V8 진행 표시 | 새 `GET /api/view/status?token=` → `{ status, progress, pageCount, error? }` + `src/client/view/prepare.ts` `waitUntilPrepared` |
| `viewerUrl`을 `/view.html`로 | `src/routes/documents.ts` `viewPathFor(ext, viewKind)` → `'/view.html' \| '/viewer.html'` |
| "보기로" 단추 | `src/client/backToView.ts` `mountBackToView(host, before, opts)`, HWP 세 갈래 확인은 `officeDialog.ts`에 더하는 `officeChoose` |
| 워터마크·함께 보는 사람 공용 모듈 | `watermark.ts` `applyWatermark` 그대로, `presence.ts`에 `renderPresenceNames(container, names)` 추가 |
| 순수 계산 | `src/client/view/logic.ts`: `chooseBodyKind`·`snapZoomIndex`·`clampZoomIndex`·`pngScale`·`needsRedraw`·`pagesToKeep`·`fitPagePercent`·`canvasPixelRatio`·`carryViewParams`·`editHref`·`offersBackToView`·`pathOf`·`wrapIndex`·`hitCountLabel`·`matchRects`·`paperMm`·`measureLabel`·`menuIdsFor`·`thumbsOpenAtStart`·`disabledReason`·`createEscapeChain` |
| `scripts/verify-view-shell-e2e.mjs`, `npm run verify:view-shell` | 그대로(Task 13) |

## 해석한 스펙 모호점(계획 작성자 결정)

1. **V4 "pageView.ts 재사용":** `pdfDocument.ts`(열기·보안 설정·자체 호스팅 worker/cMap)와 `findText.ts`(찾기 정규화)는 그대로 쓰고, **`PageView` 클래스는 쓰지 않는다.** PageView는 쪽을 돌린 viewport로 다시 그리고(`.pdf-page`, 자기 확대 단계 0.25~4) 자기 스크롤 칸을 가진다 — 그러면 표시·측정·찾기 겹침이 두 좌표계를 따로 다뤄야 하고, `verify:viewer-quicktext`가 기대하는 `.page-frame`도 없다. 대신 두 원천이 같은 `.page-frame` 틀에 그린다(pdf.js `page.render` — 그리기 품질은 PageView와 같다). `searchEngine.ts`는 편집기 찾기 창의 범위·제목/본문·절 기능이라 보기 화면에는 필요 없다. pdfjs-dist 4.2.67 그대로, 새 의존성 없음.
2. **"viewer.html은 얇은 쪽"(V2):** 엑셀 격자·도면·한글이 2·3단계까지 viewer.html을 쓰므로 viewer.html·viewer.js는 그 기능 그대로 남는다. 전환(Task 12) 뒤 viewer.js는 `/api/view/info`를 받은 직후 `viewerUrl`이 `/view.html`이면 같은 조회 문자열로 `location.replace`한다(예전에 받아 둔 주소도 새 화면으로). 반대로 view.html은 `viewerUrl`이 `/viewer.html`이면 그쪽으로 넘긴다. 기준이 서버 하나(`viewPathFor`)라 서로 넘기지 않는다. viewer.js 안의 쪽 보기 코드 삭제는 2단계에서(지금 지우면 되돌림 없이 큰 diff만 생긴다).
3. **껍데기가 맡는 형식:** `viewKind`가 `pages`인 모든 형식 중 한글(`hwp`·`hwpx`)만 뺀다 — 워드·PPT·PDF에 더해 doc·odt·odp·rtf·txt·그림(png·jpg·tif…)도 같은 쪽 문서라 함께 옮긴다(세 번째 길을 만들지 않는다). 한글은 설계 §3대로 2단계까지 viewer.html.
4. **편집 → 보기 왕복 뒤 새 쪽 그림(§6):** "보기로"는 쓰기 권한이 있으면 먼저 `POST /api/view/save`(방 저장 → PDF 다시 변환까지 끝나고 응답)를 기다린 뒤 이동한다(단추에 "저장하는 중…"). 그래서 돌아온 보기 화면은 새로 받은 PDF·PNG를 그린다(쪽 그림 응답은 이미 `no-store`, 렌더 캐시는 저장이 비운다 — 캐시 무효 쿼리 불필요). 서버에 "다시 변환 중" 표시를 새로 두지 않는다. V8 표시는 첫 변환(`editReady`) 때만.
5. **글자 층(§4 "canCopy일 때만 켠다"):** pdf.js 모드에서 `canCopy`일 때만 글자 층(`.pdf-text-layer`)을 그린다. 찾기 강조는 글자 층 없이 pdf.js 글 조각 좌표로 계산한다(`matchRects`) — 복사 권한이 없는 pdf.js 모드에서도 강조가 된다.
6. **"도구▾"(§2):** 1024px 이상은 도구 단추가 막대에 그대로 있고(`verify:viewer-quicktext`가 `#text-tool-btn`을 바로 누른다), 640~1023px은 표시·글·측정·반전·전체 화면이 "도구▾" 메뉴로, 640px 미만은 썸네일·회전·인쇄·내려받기도 들어간다(막대에는 확대·쪽·찾기·도구▾, 제목 줄에 ←·편집). 예전 "더보기" 메뉴 CSS(`.toolbar-more-wrap`·`.toolbar-more-menu`)를 그대로 쓴다.
7. **←(§2 "back=parent일 때만"):** 지금 보기 화면과 같은 인라인 스크립트를 그대로 쓴다 — ← 는 webviewer 문서 목록 링크(`home=hidden`이면 숨김), `back=parent`이고 iframe 안이면 부모에 `docview:back`만 보낸다. 편집 단추·"보기로"는 `home`·`back`을 들고 간다(`carryViewParams`).
8. **`docview:ready`:** 껍데기도 첫 쪽을 그린 뒤(또는 준비 화면·오류 화면을 보인 뒤) 부모에 보낸다 — 편집 → 보기 왕복 때 부모가 다시 받는다(§6).
9. **derived-pdf 세부(§5):** 권한 없음 403 문구는 내려받기와 같은 `저장(다운로드) 권한이 없습니다.` 원본이 PDF면 `currentPlainBytes`(pdf-current와 같은 규칙)이고 `pdfEditor` 표시는 요구하지 않는다(보기라서). 암호가 걸려 잠긴 PDF는 423 → 껍데기가 PNG로 내려간다. 보기라서 감사 기록(`downloaded`)을 남기지 않는다. 격자·도면 문서는 400.
10. **휴대폰 메모리(§7):** pdf.js 열기·첫 그리기가 실패하면 같은 몸체에 PNG 원천을 끼우고 한 줄 안내. 파일 크기로 미리 PNG를 고르는 기준은 두지 않는다(shortcut — Task 13 CPU×4·실기 시험에서 문제가 보이면 그때 크기 기준을 더한다).
11. **흐리게 vs 숨김(§2):** 강조·글 넣기(권한 없음)·측정·반전(PDF 아님)은 지금 화면과 달리 숨기지 않고 흐리게 + 이유 툴팁. 편집 단추만 V6대로 편집기 주소가 없으면 아예 없다. `verify:view-only`는 "숨김 또는 꺼짐"을 보므로 그대로 통과한다.
12. **우클릭 막기:** 지금처럼 막되, 글자 선택이 되는 화면(pdf.js + 복사 권한)에서 글자 층 위의 우클릭은 막지 않는다. 끌어서 옮기기도 글자 위에서 시작하면 하지 않는다(글자 고르기).
13. **V8 기다리기:** `/api/view/info`를 되풀이하면 감사 기록(`opened`)이 쌓이므로 기록 없는 `/api/view/status`를 1.5초마다 본다. 변환에 실패한 문서는 `resolveView`가 404를 주므로 status는 토큰만 확인하고 메타를 직접 읽는다(실패 문구를 보여 주려고). 끝나면 info를 한 번 더 받는다.
14. **한글 "보기로":** 1단계에서 한글은 viewer.html(편집 단추 없음)로 열리므로 보통 경로에서는 `from=view`가 붙지 않아 단추가 안 보인다. 설계 V7대로 만들어 두고 시험은 `from=view`를 붙인 주소로 한다(Task 13 실제 브라우저).
15. **CPU×4 기준값:** 전환 뒤에는 viewer.html이 쪽 문서를 넘겨 버리므로 지금 화면의 첫 쪽 시간은 Task 12 첫 단계(전환 전)에 잰다. 새 화면은 Task 13에서 같은 스크립트로 잰다.
16. **워터마크·함께 보는 사람 중복(§2):** 새 화면은 공용 모듈만 쓴다. viewer.js 안의 복제본은 viewer.js가 격자·도면을 맡는 동안(2·3단계) 남는다.

## 파일 구조

| 파일 | 책임 | 작업 |
|---|---|---|
| `src/routes/documents.ts`, `src/routes/viewer.ts`, `tests/viewShellRoutes.test.ts`(새) | `viewPathFor`, info 필드, `/api/view/status`·`/api/view/derived-pdf` | 1 |
| `src/client/view/types.ts`·`logic.ts`(새), `tests/viewShellLogic.test.ts`(새) | 공유 타입, 순수 계산 | 2 |
| `src/client/view/pagesBody.ts`·`pngSource.ts`·`ctx.ts`(새), `src/client/view.ts`(새), `src/client/presence.ts`, `public/view.html`·`public/css/view.css`(새), `scripts/build-editor.mjs`, `.gitignore` | PNG로 쪽이 보이는 껍데기 | 3 |
| `src/client/view/pdfjsSource.ts`(새), `src/client/view.ts`, `public/css/view.css` | pdf.js 원천·글자 층·PNG로 내려가기 | 4 |
| `src/client/view/toolbar.ts`(새), `src/client/view.ts`, `public/css/view.css` | 확대·맞춤·쪽 이동·회전·전체 화면·끌기·도구▾ | 5 |
| `src/client/view/thumbs.ts`(새), `src/client/view.ts`, `public/css/view.css` | 썸네일 칸 | 6 |
| `src/client/view/search.ts`(새), `src/client/view.ts` | 찾기(두 방식) | 7 |
| `src/client/view/fileActions.ts`·`measure.ts`(새), `src/client/view.ts` | 인쇄·내려받기·측정·반전 | 8 |
| `src/client/view/markup.ts`(새), `src/client/view.ts` | 강조·글 넣기 이전 | 9 |
| `src/client/view/editButton.ts`·`src/client/backToView.ts`(새), `officeDialog.ts`·`officeShell.ts`·`editor.ts`·`grid-editor.ts`·`editor-slides.ts`·`editor-pdf.ts`·`editor-hwp.ts`, `src/client/view.ts` | 편집 ↔ 보기 | 10 |
| `src/client/view/prepare.ts`(새), `src/client/view.ts`, `public/css/view.css` | "쪽을 준비하는 중", 오류 화면의 편집 링크 | 11 |
| `src/routes/documents.ts`, `public/js/viewer.js`, `public/api-docs.html`, `README.md`, `scripts/measure-view-first-paint.mjs`(새), `tests/viewShellRoutes.test.ts` | `viewerUrl` 전환, 기준값 측정 | 12 |
| `scripts/verify-view-shell-e2e.mjs`(새), `package.json` | 전체 검증 | 13 |

---

### Task 1: 서버 — 보기 화면 경로·info 필드·`/api/view/status`·`/api/view/derived-pdf`

**Files:**
- Modify: `src/routes/documents.ts`(import 18줄 근처, `editorUrlFor` 주석 바로 위 ~69줄)
- Modify: `src/routes/viewer.ts`(import 18줄, `/api/view/info` 응답 data ~670줄, 그 바로 뒤에 두 라우트)
- Test: `tests/viewShellRoutes.test.ts`(새)

**Interfaces:**
- Produces: `viewPathFor(ext: string, viewKind: ViewKind | undefined): '/view.html' | '/viewer.html'`(documents.ts, export)
- Produces: `GET /api/view/info` data에 `editorUrl?: string`(토큰 포함, `editorUrlFor` 규칙), `viewerUrl: string`(`viewPathFor` + `?token=`), `processing: boolean`
- Produces: `GET /api/view/status?token=` → 200 `{ success: true, data: { status: 'processing'|'ready'|'failed', progress: number, pageCount: number, error?: string } }`, 잘못된 토큰 401, 문서 없음 404
- Produces: `GET /api/view/derived-pdf?token=` → `canDownload` 아니면 403 `{ success:false, error:'저장(다운로드) 권한이 없습니다.' }`; 첫 변환 중(원본이 PDF가 아닐 때) 202 `{ success:true, data:{ processing:true } }`; 격자·도면 400; 그 밖 200 `application/pdf`, `cache-control: no-store`, `content-disposition` 없음; 잠긴 암호 PDF는 `PdfRoomError.status`(423)

- [ ] **Step 1: 워크트리 만들기**

```bash
cd /c/Users/OpenMinds/projects/webviewer
git worktree add ../webviewer-view-shell -b view-shell master
cd ../webviewer-view-shell
npm ci
cp -r ../webviewer/public/rhwp public/   # 한글 편집기 미러(무시되는 파일) — Task 13 실제 브라우저 시험용, 주 폴더는 읽기만
npx vitest run && npx tsc --noEmit
```

Expected: `npm ci`가 끝나며 `prepare`(build:editor·build:server)도 돈다. vitest·tsc 전부 통과(기준선).

- [ ] **Step 2: 실패하는 시험 쓰기** — `tests/viewShellRoutes.test.ts`

```ts
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';
import { makeSamplePdf } from './helpers.js';

type Perms = { canPrint?: boolean; canDownload?: boolean; canCopy?: boolean; canEdit?: boolean; pdfEditor?: boolean };

describe('보기 껍데기 서버 경로(info·status·derived-pdf)', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let newDocumentId: (typeof import('../src/storage.js'))['newDocumentId'];
  let issueViewToken: (typeof import('../src/security.js'))['issueViewToken'];
  let viewPathFor: (typeof import('../src/routes/documents.js'))['viewPathFor'];
  let pdf: Buffer;

  // 저장소는 메모리 — 모듈을 읽기 전에 환경을 바꿔야 해서 모두 동적 import(정적 import하면 디스크 저장소가 data/에 쓴다)
  beforeAll(async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    ({ store, newDocumentId } = await import('../src/storage.js'));
    ({ issueViewToken } = await import('../src/security.js'));
    ({ viewPathFor } = await import('../src/routes/documents.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
    pdf = await makeSamplePdf(2);
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  async function addDoc(ext: string, opts: { status?: 'ready' | 'processing' | 'failed'; viewKind?: 'pages' | 'grid'; derived?: boolean; error?: string } = {}) {
    const id = newDocumentId();
    await store.saveOriginal(id, ext, ext === 'pdf' ? pdf : Buffer.from('원본'));
    if (opts.derived !== false) await store.saveDerived(id, 'pdf', pdf);
    await store.writeMeta({
      id,
      name: `시험.${ext}`,
      ext,
      status: opts.status ?? 'ready',
      pageCount: 2,
      createdAt: new Date().toISOString(),
      ...(opts.viewKind ? { viewKind: opts.viewKind } : {}),
      ...(opts.status === 'processing' ? { editReady: true } : {}),
      ...(opts.error ? { error: opts.error } : {}),
    });
    return id;
  }
  const tok = (docId: string, perms: Perms = {}) =>
    issueViewToken({ docId, viewer: '시험자', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...perms } }, 600);
  const get = (url: string) => app.inject({ method: 'GET', url });

  test('보기 화면 경로: 쪽 문서는 껍데기, 엑셀·도면·한글은 지금 보기 화면', () => {
    expect(viewPathFor('docx', 'pages')).toBe('/view.html');
    expect(viewPathFor('pptx', undefined)).toBe('/view.html');
    expect(viewPathFor('pdf', 'pages')).toBe('/view.html');
    expect(viewPathFor('xlsx', undefined)).toBe('/viewer.html');
    expect(viewPathFor('xlsx', 'grid')).toBe('/viewer.html');
    expect(viewPathFor('dwg', 'svg')).toBe('/viewer.html');
    expect(viewPathFor('hwp', 'pages')).toBe('/viewer.html');
    expect(viewPathFor('hwpx', 'pages')).toBe('/viewer.html');
  });

  test('derived-pdf: 내려받기 권한이 없으면 403이고 PDF를 주지 않는다', async () => {
    const docx = await addDoc('docx', { viewKind: 'pages' });
    const res = await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(docx, { canEdit: true, canCopy: true, canPrint: true }))}`);
    expect(res.statusCode).toBe(403);
    expect(res.json()).toEqual({ success: false, error: '저장(다운로드) 권한이 없습니다.' });
    const pdfDoc = await addDoc('pdf', { viewKind: 'pages' });
    expect((await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(pdfDoc))}`)).statusCode).toBe(403);
  });

  test('derived-pdf: 내려받기 권한이 있으면 변환한 PDF(보기용 — 저장 이름 없음, 캐시 안 함)', async () => {
    const docx = await addDoc('docx', { viewKind: 'pages' });
    const res = await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(docx, { canDownload: true }))}`);
    expect(res.statusCode).toBe(200);
    expect(res.headers['content-type']).toBe('application/pdf');
    expect(res.headers['cache-control']).toBe('no-store');
    expect(res.headers['content-disposition']).toBeUndefined();
    expect(res.rawPayload.subarray(0, 5).toString('latin1')).toBe('%PDF-');
  });

  test('derived-pdf: 원본이 PDF면 원본 바이트, 첫 변환 중이면 202, 격자 문서는 400', async () => {
    const pdfDoc = await addDoc('pdf', { viewKind: 'pages', derived: false });
    const res = await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(pdfDoc, { canDownload: true }))}`);
    expect(res.statusCode).toBe(200);
    expect(Buffer.compare(res.rawPayload, pdf)).toBe(0);
    const busy = await addDoc('pptx', { status: 'processing', derived: false });
    const r202 = await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(busy, { canDownload: true }))}`);
    expect(r202.statusCode).toBe(202);
    expect(r202.json()).toEqual({ success: true, data: { processing: true } });
    const grid = await addDoc('xlsx', { viewKind: 'grid', derived: false });
    expect((await get(`/api/view/derived-pdf?token=${encodeURIComponent(tok(grid, { canDownload: true }))}`)).statusCode).toBe(400);
  });

  test('info: 편집기·보기 주소와 첫 변환 여부', async () => {
    const docx = await addDoc('docx', { viewKind: 'pages' });
    const t = tok(docx, { canEdit: true });
    const data = (await get(`/api/view/info?token=${encodeURIComponent(t)}`)).json().data;
    expect(data.editorUrl).toBe(`/editor.html?token=${encodeURIComponent(t)}`);
    expect(data.viewerUrl).toBe(`/view.html?token=${encodeURIComponent(t)}`);
    expect(data.processing).toBe(false);
    // 열람만이면 편집기 주소 없음(토큰 응답 규칙 editorUrlFor 그대로)
    const viewOnly = (await get(`/api/view/info?token=${encodeURIComponent(tok(docx))}`)).json().data;
    expect(viewOnly.editorUrl).toBeUndefined();
    // pdfEditor 없는 PDF(8000 연동)는 편집기 없이 보기만
    const pdfDoc = await addDoc('pdf', { viewKind: 'pages' });
    expect((await get(`/api/view/info?token=${encodeURIComponent(tok(pdfDoc, { canEdit: true, canDownload: true }))}`)).json().data.editorUrl).toBeUndefined();
    const hwp = await addDoc('hwp', { viewKind: 'pages' });
    expect((await get(`/api/view/info?token=${encodeURIComponent(tok(hwp))}`)).json().data.viewerUrl.startsWith('/viewer.html?')).toBe(true);
    const busy = await addDoc('pptx', { status: 'processing', derived: false });
    expect((await get(`/api/view/info?token=${encodeURIComponent(tok(busy))}`)).json().data.processing).toBe(true);
  });

  test('status: 진행률, 실패 문구(열 수 없는 문서도), 잘못된 토큰 401', async () => {
    const busy = await addDoc('pptx', { status: 'processing', derived: false });
    const s1 = (await get(`/api/view/status?token=${encodeURIComponent(tok(busy))}`)).json();
    expect(s1).toEqual({ success: true, data: { status: 'processing', progress: 0, pageCount: 2 } });
    const failed = await addDoc('docx', { status: 'failed', derived: false, error: '문서를 변환하지 못했습니다. 파일이 손상되었습니다.' });
    const s2 = (await get(`/api/view/status?token=${encodeURIComponent(tok(failed))}`)).json();
    expect(s2.data.status).toBe('failed');
    expect(s2.data.error).toBe('문서를 변환하지 못했습니다. 파일이 손상되었습니다.');
    expect((await get('/api/view/status?token=bad')).statusCode).toBe(401);
  });
});
```

- [ ] **Step 3: 실패 확인**

Run: `npx vitest run tests/viewShellRoutes.test.ts`
Expected: FAIL — `viewPathFor is not a function`, `/api/view/derived-pdf`·`/api/view/status` 404, info에 `viewerUrl` 없음.

- [ ] **Step 4: `src/routes/documents.ts` — `viewPathFor`**

import 줄 `import type { DocumentMeta } from '../types.js';`를 `import type { DocumentMeta, ViewKind } from '../types.js';`로 바꾸고, `/**\n * 토큰 응답의 editorUrl.` 주석 바로 위에 넣는다(`GRID_VIEW_EXTS`는 같은 파일 위쪽에 이미 있다):

```ts
/** 지금 보기 화면(viewer.html)에 남는 쪽 문서 — 한글은 2단계에서 껍데기로 옮긴다(설계 §1·§3) */
const LEGACY_VIEW_EXTS = new Set(['hwp', 'hwpx']);

/**
 * 보기 화면 경로 — 쪽 문서(한글 빼고)는 통합 보기 껍데기(/view.html), 엑셀 격자·도면·한글은 지금 보기 화면(/viewer.html).
 * 변환 전(editReady)이라 viewKind가 아직 없으면 확장자로 정한다(엑셀·csv는 격자, 나머지는 쪽).
 */
export function viewPathFor(ext: string, viewKind: ViewKind | undefined): '/view.html' | '/viewer.html' {
  const kind = viewKind ?? (GRID_VIEW_EXTS.has(ext) ? 'grid' : 'pages');
  return kind === 'pages' && !LEGACY_VIEW_EXTS.has(ext) ? '/view.html' : '/viewer.html';
}
```

- [ ] **Step 5: `src/routes/viewer.ts` — info 필드와 두 라우트**

import 줄 `import { convertToPdf } from './documents.js';` → `import { convertToPdf, editorUrlFor, viewPathFor } from './documents.js';`(viewer.ts ↔ documents.ts는 이미 서로 import한다 — 함수는 호출 때 읽으므로 문제없다).

`/api/view/info` 응답의 `data`에서 `slideSize,` 다음에 넣는다:

```ts
          // 통합 보기 껍데기(view.html)의 편집 단추와 편집기의 "보기로"가 쓴다 — 토큰 하나로 오가므로 새 토큰을 만들지 않는다
          editorUrl: editorUrlFor(
            view.meta.ext,
            { effectiveCanEdit: view.claims.perms.canEdit, canDownload: view.claims.perms.canDownload, pdfEditor: view.claims.perms.pdfEditor === true },
            request.query.token!,
          ),
          viewerUrl: `${viewPathFor(view.meta.ext, view.meta.viewKind)}?token=${encodeURIComponent(request.query.token!)}`,
          processing: view.meta.status === 'processing',
```

(`perms.canEdit`는 토큰을 낼 때 이미 `effectiveCanEdit`로 저장돼 있고, `pdfEditor`는 pdf·내려받기 권한일 때만 붙어 있다 — 토큰 응답과 같은 결과.)

info 라우트의 닫는 `);` 바로 다음(presence 라우트 앞)에 두 라우트를 넣는다(`verifyViewToken`·`ViewClaims`·`currentPlainBytes`·`tokenKey`·`PdfRoomError`·`store`는 이미 import돼 있다):

```ts
  /**
   * 변환 상태 — 보기 화면이 첫 변환(editReady)을 기다리는 동안 부른다("쪽을 준비하는 중", 설계 V8). 감사 기록을 남기지 않고,
   * 변환에 실패한 문서(열 수 없는 상태)도 그 문구를 돌려준다.
   */
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/status',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      let claims: ViewClaims;
      try {
        claims = verifyViewToken(request.query.token ?? '');
      } catch {
        return reply.code(401).send({ success: false, error: '열람 토큰이 없거나 유효하지 않습니다.' });
      }
      const meta = await store.readMeta(claims.docId);
      if (!meta) return reply.code(404).send({ success: false, error: '문서를 찾을 수 없습니다.' });
      return {
        success: true,
        data: {
          status: meta.status,
          progress: meta.progress ?? 0,
          pageCount: meta.pageCount,
          ...(meta.status === 'failed' ? { error: meta.error ?? '문서를 변환하지 못했습니다.' } : {}),
        },
      };
    },
  );

  /**
   * 보기용 PDF(설계 V5) — 통합 보기 껍데기가 내려받기 권한이 있을 때 pdf.js로 그린다. 권한이 없으면 403(PDF가 브라우저로 가지
   * 않는다). 원본이 PDF면 지금 상태(pdf-current와 같은 규칙: 편집 방이 열려 있으면 저장 전 상태, 암호 파일은 비밀번호를 맞힌 토큰만),
   * 그 밖에는 변환한 PDF. 첫 변환 중이면 202. 내려받기와 달리 보기라서 파일 이름·감사 기록을 붙이지 않는다.
   */
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/derived-pdf',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canDownload) {
        return reply.code(403).send({ success: false, error: '저장(다운로드) 권한이 없습니다.' });
      }
      const meta = view.meta;
      try {
        let bytes: Uint8Array;
        if (meta.ext === 'pdf') {
          bytes = await currentPlainBytes(meta.id, () => store.readOriginal(meta.id, meta.ext), tokenKey(request.query.token!));
        } else if (meta.status === 'processing') {
          return reply.code(202).send({ success: true, data: { processing: true } });
        } else if ((meta.viewKind ?? 'pages') !== 'pages') {
          return reply.code(400).send({ success: false, error: '쪽으로 보는 문서가 아닙니다.' });
        } else {
          bytes = await store.readDerived(meta.id, 'pdf');
        }
        return await reply
          .type('application/pdf')
          .header('cache-control', 'no-store')
          .send(Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength));
      } catch (error) {
        if (error instanceof PdfRoomError) return reply.code(error.status).send({ success: false, error: error.message });
        request.log.error(error, 'derived pdf failed');
        return reply.code(500).send({ success: false, error: '보기용 PDF를 불러오지 못했습니다.' });
      }
    },
  );
```

- [ ] **Step 6: 통과 확인**

Run: `npx vitest run tests/viewShellRoutes.test.ts && npx vitest run && npx tsc --noEmit`
Expected: 새 시험 6개 PASS, 전체 vitest PASS(기존 info를 쓰는 시험 포함), tsc 오류 0. 시험이 `data/`를 만들지 않았는지 `git status --short`로 확인(무시 대상이지만 남기지 않는다).

- [ ] **Step 7: 커밋**

```bash
git add tests/viewShellRoutes.test.ts
git commit -m "$(cat <<'EOF'
통합 보기 서버: 보기 화면 경로(viewPathFor), info에 편집·보기 주소, derived-pdf(내려받기 권한만)·status 경로

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/routes/documents.ts src/routes/viewer.ts tests/viewShellRoutes.test.ts
```

---

### Task 2: 껍데기 순수 계산과 공유 타입

**Files:**
- Create: `src/client/view/types.ts`, `src/client/view/logic.ts`
- Test: `tests/viewShellLogic.test.ts`

**Interfaces:**
- Consumes: `FindMatch`·`findInPage`(`src/client/pdf/findText.ts`, 그대로)
- Produces(types.ts): `PageSize {w,h}`, `NormRect {x,y,w,h}`, `SearchHit {page, rects}`, `ViewPermissions`, `ViewInfo`, `BodyKind`, `ZoomMode`, `ViewBody`
- Produces(logic.ts): 아래 코드의 export 전부 — 이후 작업이 이 이름 그대로 쓴다

- [ ] **Step 1: 실패하는 시험 쓰기** — `tests/viewShellLogic.test.ts`

```ts
import { describe, expect, test } from 'vitest';
import { findInPage } from '../src/client/pdf/findText.js';
import {
  canvasPixelRatio,
  carryViewParams,
  chooseBodyKind,
  createEscapeChain,
  disabledReason,
  editHref,
  fitPagePercent,
  hitCountLabel,
  matchRects,
  measureLabel,
  menuIdsFor,
  needsRedraw,
  offersBackToView,
  pagesToKeep,
  paperMm,
  pathOf,
  pngScale,
  snapZoomIndex,
  thumbsOpenAtStart,
  wrapIndex,
  ZOOM_STEPS,
} from '../src/client/view/logic.js';

const perms = (p: Partial<{ canPrint: boolean; canDownload: boolean; canCopy: boolean; canEdit: boolean }> = {}) => ({
  canPrint: false,
  canDownload: false,
  canCopy: false,
  canEdit: false,
  ...p,
});

describe('권한 → 안쪽', () => {
  test('내려받기 권한이 있을 때만 pdf.js, 아니면 서버 PNG', () => {
    expect(chooseBodyKind(perms({ canDownload: true }))).toBe('pages-pdfjs');
    expect(chooseBodyKind(perms({ canEdit: true, canCopy: true, canPrint: true }))).toBe('pages-png');
  });
});

describe('확대·그리기', () => {
  test('가장 가까운 확대 단계', () => {
    expect(ZOOM_STEPS[snapZoomIndex(100)]).toBe(100);
    expect(ZOOM_STEPS[snapZoomIndex(72)]).toBe(67);
    expect(ZOOM_STEPS[snapZoomIndex(10)]).toBe(50);
    expect(ZOOM_STEPS[snapZoomIndex(9999)]).toBe(500);
  });
  test('서버 PNG 배율은 viewer.js와 같다(2~16, dpr은 2까지)', () => {
    expect(pngScale(100, 1)).toBe(2);
    expect(pngScale(100, 1.5)).toBe(3);
    expect(pngScale(100, 3)).toBe(4);
    expect(pngScale(50, 1)).toBe(2);
    expect(pngScale(500, 2)).toBe(16);
  });
  test('다시 그리기: 처음·더 선명하게·두 배 넘게 큰 그림', () => {
    expect(needsRedraw(0, 4)).toBe(true);
    expect(needsRedraw(4, 4)).toBe(false);
    expect(needsRedraw(4, 6)).toBe(true);
    expect(needsRedraw(4, 3)).toBe(false);
    expect(needsRedraw(1000, 400)).toBe(true);
  });
  test('남길 쪽 = 보이는 쪽 + 지금 쪽 ±2(문서 범위 안)', () => {
    expect([...pagesToKeep(1, 10, 2, [])].sort((a, b) => a - b)).toEqual([1, 2, 3]);
    expect([...pagesToKeep(5, 10, 2, [9])].sort((a, b) => a - b)).toEqual([3, 4, 5, 6, 7, 9]);
    expect([...pagesToKeep(10, 10, 2, [])].sort((a, b) => a - b)).toEqual([8, 9, 10]);
  });
  test('쪽 맞춤: 높이가 넉넉하면 100, 아니면 줄인다', () => {
    expect(fitPagePercent(1032, 2000, { w: 595, h: 842 })).toBe(100);
    // 폭 1000(-32=968) → 100%에서 높이 968*842/595 ≈ 1369.9, 남은 높이 800-32=768 → 약 56%
    expect(fitPagePercent(1000, 800, { w: 595, h: 842 })).toBeCloseTo(56.06, 1);
  });
  test('캔버스 화소 상한', () => {
    expect(canvasPixelRatio(1000, 1000, 2)).toBe(2);
    const r = canvasPixelRatio(4000, 4000, 2);
    expect(4000 * 4000 * r * r).toBeLessThanOrEqual(16_777_216 + 1);
  });
});

describe('편집 ↔ 보기 주소', () => {
  test('home·back만 들고 가고 extra를 붙인다', () => {
    expect(carryViewParams('/editor.html?token=a.b', '?token=a.b&home=hidden&back=parent&from=x', { from: 'view' })).toBe(
      '/editor.html?token=a.b&home=hidden&back=parent&from=view',
    );
    expect(carryViewParams('/view.html?token=t', '?token=t&from=view')).toBe('/view.html?token=t');
  });
  test('편집 단추 주소에는 from=view', () => {
    expect(editHref('/editor-slides.html?token=t', '?token=t')).toBe('/editor-slides.html?token=t&from=view');
  });
  test('"보기로"는 from=view이고 viewerUrl이 있을 때만', () => {
    expect(offersBackToView('?token=t&from=view', '/view.html?token=t')).toBe(true);
    expect(offersBackToView('?token=t', '/view.html?token=t')).toBe(false);
    expect(offersBackToView('?token=t&from=view', undefined)).toBe(false);
  });
  test('경로만', () => {
    expect(pathOf('/view.html?token=t')).toBe('/view.html');
  });
});

describe('찾기', () => {
  test('순환 번호와 표시', () => {
    expect(wrapIndex(-1, 3)).toBe(2);
    expect(wrapIndex(3, 3)).toBe(0);
    expect(wrapIndex(0, 0)).toBe(-1);
    expect(hitCountLabel(0, 3)).toBe('1 / 3');
    expect(hitCountLabel(4, 500)).toBe('5 / 500+');
    expect(hitCountLabel(-1, 0)).toBe('결과 없음');
  });
  test('pdf.js 글 조각 → 쪽 대비 사각형(배율 1·회전 0, A4)', () => {
    const items = [{ str: 'hello world', transform: [10, 0, 0, 10, 100, 700], width: 55, height: 10 }];
    const [m] = findInPage(1, items, 'world');
    const vp = { transform: [1, 0, 0, -1, 0, 842], width: 595, height: 842 };
    const [r] = matchRects(m, items, vp);
    expect(r.x).toBeCloseTo((100 + (55 * 6) / 11) / 595, 6);
    expect(r.w).toBeCloseTo((55 * 5) / 11 / 595, 6);
    expect(r.y).toBeCloseTo((842 - 710) / 842, 6);
    expect(r.h).toBeCloseTo(10 / 842, 6);
  });
  test('조각 두 개에 걸친 일치는 사각형 두 개', () => {
    const items = [
      { str: 'abc', transform: [10, 0, 0, 10, 0, 100], width: 30, height: 10 },
      { str: 'def', transform: [10, 0, 0, 10, 30, 100], width: 30, height: 10 },
    ];
    const [m] = findInPage(1, items, 'cd');
    expect(matchRects(m, items, { transform: [1, 0, 0, -1, 0, 200], width: 100, height: 200 })).toHaveLength(2);
  });
});

describe('측정', () => {
  test('A4 폭 전체 = 210mm 남짓, 이름표 자리수', () => {
    expect(paperMm({ fx: 0, fy: 0 }, { fx: 1, fy: 0 }, { w: 595, h: 842 })).toBeCloseTo(209.9, 1);
    expect(measureLabel(209.9, 1, 'mm(용지)')).toBe('210 mm(용지)');
    expect(measureLabel(12.34, 1, 'mm')).toBe('12.3 mm');
  });
});

describe('막대 배치·단추 상태', () => {
  test('1024px 이상은 메뉴 없음, 640~1023은 도구만, 그 아래는 더', () => {
    expect(menuIdsFor(1400)).toEqual([]);
    expect(menuIdsFor(800)).toContain('text-tool-btn');
    expect(menuIdsFor(800)).not.toContain('print-btn');
    expect(menuIdsFor(390)).toContain('print-btn');
    expect(menuIdsFor(390)).not.toContain('zoom-in');
  });
  test('썸네일은 휴대폰 폭에서 닫힌 채', () => {
    expect(thumbsOpenAtStart(390)).toBe(false);
    expect(thumbsOpenAtStart(1280)).toBe(true);
  });
  test('흐린 단추의 이유', () => {
    expect(disabledReason('print', perms(), 'docx')).toBe('인쇄 권한이 없습니다');
    expect(disabledReason('print', perms({ canPrint: true }), 'docx')).toBeNull();
    expect(disabledReason('download', perms(), 'pdf')).toBe('내려받기 권한이 없습니다');
    expect(disabledReason('markup', perms(), 'pdf')).toBe('표시를 넣을 권한이 없습니다');
    expect(disabledReason('measure', perms(), 'docx')).toBe('PDF 문서에서만 쓸 수 있습니다');
    expect(disabledReason('measure', perms(), 'pdf')).toBeNull();
  });
  test('Esc는 우선순위 순서로 처음 받은 곳에서 멈춘다', () => {
    const calls: string[] = [];
    const chain = createEscapeChain();
    chain.add(90, () => (calls.push('search'), true));
    chain.add(10, () => (calls.push('measure'), false));
    chain.add(40, () => (calls.push('markup'), true));
    expect(chain.run()).toBe(true);
    expect(calls).toEqual(['measure', 'markup']);
  });
});
```

- [ ] **Step 2: 실패 확인**

Run: `npx vitest run tests/viewShellLogic.test.ts`
Expected: FAIL — `Cannot find module '../src/client/view/logic.js'`.

- [ ] **Step 3: `src/client/view/types.ts`**

```ts
/** 통합 보기 껍데기의 공유 타입 — 설계 §4 ViewBody(계획의 "설계와 맞춘 이름" 참고) */
export interface PageSize {
  readonly w: number;
  readonly h: number;
}

/** 쪽 크기 대비 비율(0~1) 사각형 — 서버 찾기(/api/view/search)의 rects와 같은 모양 */
export interface NormRect {
  readonly x: number;
  readonly y: number;
  readonly w: number;
  readonly h: number;
}

export interface SearchHit {
  readonly page: number;
  readonly rects: readonly NormRect[];
}

export interface ViewPermissions {
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  readonly canCopy: boolean;
  readonly canEdit: boolean;
  readonly pdfEditor?: boolean;
}

/** GET /api/view/info 의 data(이 화면이 쓰는 것만) */
export interface ViewInfo {
  readonly docId: string;
  readonly name: string;
  readonly ext: string;
  readonly pageCount: number;
  readonly viewKind: 'pages' | 'grid' | 'svg';
  readonly pages?: readonly PageSize[];
  readonly permissions: ViewPermissions;
  readonly viewer: string;
  /** 편집기 주소(토큰 포함) — 없으면 편집 단추가 없다(설계 V6) */
  readonly editorUrl?: string;
  /** 이 문서의 보기 화면 주소(토큰 포함) — 쪽 문서는 /view.html, 엑셀·도면·한글은 /viewer.html */
  readonly viewerUrl: string;
  /** 첫 변환 중(editReady) — 쪽 자리에 "쪽을 준비하는 중"(설계 V8) */
  readonly processing: boolean;
}

export type BodyKind = 'pages-png' | 'pages-pdfjs' | 'grid' | 'cad' | 'frame';

/** number = 확대 %(logic.ts ZOOM_STEPS, 100 = 쪽 칸 폭 전체) */
export type ZoomMode = number | 'fit-width' | 'fit-page';

export interface ViewBody {
  readonly kind: BodyKind;
  readonly textSelectable: boolean;
  mount(host: HTMLElement): Promise<void>;
  /** 쪽 개념이 없으면 0(막대의 쪽 칸 숨김) */
  pageCount(): number;
  goto(page: number): void;
  onPage(cb: (page: number) => void): void;
  zoom(z: ZoomMode): void;
  rotate(deg: 90 | -90): void;
  search?(q: string, signal: AbortSignal): Promise<SearchHit[]>;
  thumbnail?(page: number, width: number): Promise<HTMLCanvasElement | HTMLImageElement>;
  destroy(): void;
}
```

- [ ] **Step 4: `src/client/view/logic.ts`**

```ts
/**
 * 통합 보기 껍데기의 순수 계산 — DOM 없이 시험한다(tests/viewShellLogic.test.ts).
 * 값은 지금 보기 화면(public/js/viewer.js)과 같게 둔다(확대 단계·PNG 배율·쪽 맞춤·측정).
 */
import type { FindMatch } from '../pdf/findText.js';
import type { NormRect, PageSize, ViewPermissions } from './types.js';

/** 확대 단계(%) — viewer.js와 같다. 100% = 쪽 칸 폭 전체 */
export const ZOOM_STEPS = [50, 67, 80, 100, 125, 150, 200, 250, 300, 400, 500] as const;
export const DEFAULT_ZOOM_INDEX = 3;
/** 캔버스 한 장 최대 화소(pageView.ts와 같은 값) — 넘으면 해상도를 낮춰 그리고 CSS로 늘린다 */
export const MAX_CANVAS_PIXELS = 16_777_216;
/** 찾기 결과 상한(서버 /api/view/search limit 최대와 같다) */
export const SEARCH_LIMIT = 500;

/** 권한 → 안쪽: 내려받기 권한이 있을 때만 PDF를 브라우저로 받는다(설계 V5·§4) */
export function chooseBodyKind(perms: Pick<ViewPermissions, 'canDownload'>): 'pages-pdfjs' | 'pages-png' {
  return perms.canDownload ? 'pages-pdfjs' : 'pages-png';
}

export function clampZoomIndex(index: number): number {
  return Math.max(0, Math.min(ZOOM_STEPS.length - 1, index));
}

/** 가장 가까운 확대 단계(같으면 작은 쪽) */
export function snapZoomIndex(percent: number): number {
  let best = 0;
  ZOOM_STEPS.forEach((step, i) => {
    if (Math.abs(step - percent) < Math.abs(ZOOM_STEPS[best] - percent)) best = i;
  });
  return best;
}

/** 서버 PNG 배율(viewer.js desiredScale 그대로) */
export function pngScale(percent: number, dpr: number): number {
  const ratio = Math.min(dpr || 1, 2);
  return Math.min(16, Math.max(2, Math.ceil((percent / 100) * 2 * ratio)));
}

/** 다시 그릴지 — 처음이거나, 더 선명해야 하거나, 지금 그림이 필요의 두 배를 넘어 메모리를 낭비할 때 */
export function needsRedraw(drawnKey: number, wantKey: number): boolean {
  return drawnKey === 0 || wantKey > drawnKey || wantKey * 2 < drawnKey;
}

/** 캔버스를 남겨 둘 쪽 — 보이는 쪽 + 지금 쪽 앞뒤 radius(설계 §7 "보이는 쪽 ±2") */
export function pagesToKeep(current: number, count: number, radius: number, visible: Iterable<number>): Set<number> {
  const keep = new Set<number>(visible);
  for (let page = Math.max(1, current - radius); page <= Math.min(count, current + radius); page += 1) keep.add(page);
  return keep;
}

/** "쪽 맞춤" 확대(%) — viewer.js pageFitPercent 그대로(좌우·위아래 여백 32px). 100%에서 쪽이 남은 높이를 넘으면 줄인다 */
export function fitPagePercent(containerWidth: number, availableHeight: number, size: PageSize): number {
  const heightAt100 = (containerWidth - 32) * (size.h / size.w);
  const room = availableHeight - 32;
  return heightAt100 <= room ? 100 : (room / heightAt100) * 100;
}

/** 캔버스 화소 배율 — 화면 배율로 그리되 maxPixels를 넘지 않게 */
export function canvasPixelRatio(cssW: number, cssH: number, dpr: number, maxPixels = MAX_CANVAS_PIXELS): number {
  const ratio = dpr || 1;
  return cssW * cssH * ratio * ratio > maxPixels ? Math.sqrt(maxPixels / (cssW * cssH)) : ratio;
}

/** 편집기 ↔ 보기 화면을 오갈 때 들고 가는 임베드 옵션(editor*.html·viewer.html의 home·back) */
const CARRIED_PARAMS = ['home', 'back'] as const;

/** target(토큰 포함 상대 주소)에 지금 주소의 home·back과 extra를 붙인다. from 등 다른 값은 옮기지 않는다 */
export function carryViewParams(target: string, currentSearch: string, extra: Readonly<Record<string, string>> = {}): string {
  const url = new URL(target, 'http://local.invalid');
  const current = new URLSearchParams(currentSearch);
  for (const key of CARRIED_PARAMS) {
    const value = current.get(key);
    if (value !== null) url.searchParams.set(key, value);
  }
  for (const [key, value] of Object.entries(extra)) url.searchParams.set(key, value);
  return `${url.pathname}${url.search}`;
}

/** 편집 단추가 갈 주소 — 편집기 "보기로"가 보이게 from=view를 붙인다(설계 §6) */
export function editHref(editorUrl: string, currentSearch: string): string {
  return carryViewParams(editorUrl, currentSearch, { from: 'view' });
}

/** 편집기의 "보기로" 단추를 보일지 — 껍데기에서 편집으로 들어온 경우만(설계 V7) */
export function offersBackToView(currentSearch: string, viewerUrl: string | undefined): boolean {
  return Boolean(viewerUrl) && new URLSearchParams(currentSearch).get('from') === 'view';
}

export function pathOf(url: string): string {
  return new URL(url, 'http://local.invalid').pathname;
}

/** 순환 번호(찾기 이전·다음) — n이 0이면 -1 */
export function wrapIndex(index: number, n: number): number {
  return n === 0 ? -1 : ((index % n) + n) % n;
}

/** 찾기 칸 표시 — viewer.js gotoHit와 같은 "3 / 500+" */
export function hitCountLabel(current: number, total: number): string {
  if (total === 0) return '결과 없음';
  return `${current + 1} / ${total >= SEARCH_LIMIT ? `${SEARCH_LIMIT}+` : String(total)}`;
}

/** pdf.js 글 조각의 모양(TextItem에서 쓰는 것만) */
export interface TextItemGeom {
  readonly str: string;
  readonly transform: readonly number[];
  readonly width: number;
  readonly height: number;
}

/** 배율 1·회전 0 viewport의 변환·크기 */
export interface ViewportGeom {
  readonly transform: readonly number[];
  readonly width: number;
  readonly height: number;
}

const applyTransform = (t: readonly number[], x: number, y: number): [number, number] => [t[0] * x + t[2] * y + t[4], t[1] * x + t[3] * y + t[5]];

/**
 * 찾은 곳(findText.ts FindMatch — 조각 번호·조각 안 글자 위치) → 쪽 대비 사각형. 글자 층(DOM) 없이 계산하므로 글자 선택이
 * 꺼진(canCopy 없음) pdf.js 모드에서도 찾기 강조가 된다. 조각 안 글자 폭은 고르게 나눈다.
 * shortcut: 가로 글자만 정확하다(돌아간·세로 글 조각은 기준점에서 가로로 근사) — 도면 PDF에서 어긋나 보이면 transform의 a·b로 방향을 쓴다.
 */
export function matchRects(m: FindMatch, items: readonly TextItemGeom[], vp: ViewportGeom): NormRect[] {
  const out: NormRect[] = [];
  for (let i = m.itemStart; i <= m.itemEnd; i += 1) {
    const item = items[i];
    if (!item || item.str.length === 0) continue;
    const len = item.str.length;
    const start = i === m.itemStart ? Math.min(m.offsetStart, len) : 0;
    const end = i === m.itemEnd ? Math.min(m.offsetEnd, len) : len;
    if (end <= start) continue;
    const [, , c, d, tx, ty] = item.transform;
    const height = item.height || Math.hypot(c, d);
    const x0 = tx + (item.width * start) / len;
    const x1 = tx + (item.width * end) / len;
    const [ax, ay] = applyTransform(vp.transform, x0, ty);
    const [bx, by] = applyTransform(vp.transform, x1, ty + height);
    const left = Math.min(ax, bx);
    const top = Math.min(ay, by);
    out.push({ x: left / vp.width, y: top / vp.height, w: (Math.max(ax, bx) - left) / vp.width, h: (Math.max(ay, by) - top) / vp.height });
  }
  return out;
}

/** 측정: 두 점(쪽 비율) 사이 용지 거리(mm) — viewer.js addMeasurement 그대로(1pt = 1/72in) */
export function paperMm(a: { readonly fx: number; readonly fy: number }, b: { readonly fx: number; readonly fy: number }, size: PageSize): number {
  return (Math.hypot((b.fx - a.fx) * size.w, (b.fy - a.fy) * size.h) * 25.4) / 72;
}

/** 측정 이름표 — viewer.js measurementLabelText 그대로 */
export function measureLabel(mm: number, factor: number, unit: string): string {
  const value = mm * factor;
  return `${value >= 100 ? value.toFixed(0) : value.toFixed(1)} ${unit}`;
}

/** 640~1023px에서 "도구▾" 메뉴로 들어가는 것(설계 §2 도구▾). 1024px 이상은 막대에 그대로 둔다(verify:viewer-quicktext가 #text-tool-btn을 누른다) */
export const TOOL_MENU_IDS = ['markup-add-btn', 'text-tool-btn', 'markup-include-wrap', 'measure-btn', 'invert-btn', 'fullscreen-btn'] as const;
/** 640px 미만에서 더 들어가는 것 — 막대에는 확대·쪽·찾기·도구▾만 남는다(설계 §2) */
export const NARROW_MENU_IDS = ['thumbs-btn', 'rotate-btn', 'print-btn', 'download-btn', 'download-as-btn'] as const;

export function menuIdsFor(width: number): string[] {
  if (width >= 1024) return [];
  return width >= 640 ? [...TOOL_MENU_IDS] : [...TOOL_MENU_IDS, ...NARROW_MENU_IDS];
}

/** 썸네일 칸 — 640px 미만은 닫힌 채 시작(설계 §2) */
export function thumbsOpenAtStart(width: number): boolean {
  return width >= 640;
}

export type ControlName = 'print' | 'download' | 'markup' | 'measure';

/** 흐린 단추의 이유(툴팁) — null이면 쓸 수 있다(설계 §2 "흐리게 하고 이유를 툴팁으로") */
export function disabledReason(control: ControlName, perms: ViewPermissions, ext: string): string | null {
  switch (control) {
    case 'print':
      return perms.canPrint ? null : '인쇄 권한이 없습니다';
    case 'download':
      return perms.canDownload ? null : '내려받기 권한이 없습니다';
    case 'markup':
      return perms.canEdit ? null : '표시를 넣을 권한이 없습니다';
    case 'measure':
      return ext === 'pdf' ? null : 'PDF 문서에서만 쓸 수 있습니다';
  }
}

/** Esc 처리 순서 — priority가 작은 것부터 불러 true를 돌려준 곳에서 멈춘다 */
export function createEscapeChain(): { add(priority: number, handler: () => boolean): void; run(): boolean } {
  const list: Array<{ readonly priority: number; readonly handler: () => boolean }> = [];
  return {
    add(priority, handler) {
      list.push({ priority, handler });
      list.sort((a, b) => a.priority - b.priority);
    },
    run() {
      return list.some((entry) => entry.handler());
    },
  };
}
```

- [ ] **Step 5: 통과 확인**

Run: `npx vitest run tests/viewShellLogic.test.ts && npx tsc --noEmit`
Expected: 19개 PASS(계획 작성 때 같은 코드로 확인함), tsc 오류 0.

- [ ] **Step 6: 커밋**

```bash
git add src/client/view/types.ts src/client/view/logic.ts tests/viewShellLogic.test.ts
git commit -m "$(cat <<'EOF'
통합 보기: ViewBody 타입과 순수 계산(권한→안쪽, 확대·PNG 배율·가상화, 편집↔보기 주소, 찾기 사각형, 측정, 막대 배치)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/client/view/types.ts src/client/view/logic.ts tests/viewShellLogic.test.ts
```

---

### Task 3: PNG로 쪽이 보이는 껍데기 — PagesBody·PngSource·view.html·view.ts·빌드

**Files:**
- Create: `src/client/view/pagesBody.ts`, `src/client/view/pngSource.ts`, `src/client/view/ctx.ts`, `src/client/view.ts`, `public/view.html`, `public/css/view.css`
- Modify: `src/client/presence.ts`(끝부분 `renderPresenceAvatars`), `scripts/build-editor.mjs`(entryPoints), `.gitignore`

**Interfaces:**
- Consumes: Task 2의 `logic.ts`·`types.ts`, Task 1의 info 필드
- Produces: `PageSource` 인터페이스·`DrawOpts`·`PagesBody`(아래 코드의 public 멤버 전부), `PngSource(tokenQuery, pageCount, sizes)`, `ShellCtx`, `renderPresenceNames(container, names)`
- Produces(DOM 약속, 이후 작업이 쓴다): `#pages.pages-container` 안에 쪽마다 `div.page-frame[data-page]`(그린 뒤 `.page-ready`, 실패 `.page-error`) > `canvas` + `span.page-number` (+ 글자 층이면 `div.pdf-text-layer`). view.html의 id 목록은 Step 5 끝 참고.

- [ ] **Step 1: `src/client/view/pagesBody.ts`**

```ts
/**
 * 쪽 문서 안쪽(ViewBody) — 쪽마다 .page-frame(지금 보기 화면 viewer.js와 같은 틀·CSS)을 만들고, 그림은 원천(PageSource)이
 * 그린다: 서버 PNG(pngSource.ts) 또는 브라우저 pdf.js(pdfjsSource.ts). 확대(쪽 칸 폭 %)·보기 회전(CSS)·현재 쪽·가상화(보이는 쪽 +
 * 지금 쪽 ±2만 캔버스 유지)·찾기 강조는 두 원천이 같다 — 마크업·측정·찾기 겹침이 .page-frame 비율 좌표 하나로 붙는다.
 * pdf.js가 실패하면 replaceSource로 같은 몸체에 PNG 원천을 끼운다(등록한 콜백·확대는 그대로).
 */
import { DEFAULT_ZOOM_INDEX, ZOOM_STEPS, clampZoomIndex, fitPagePercent, needsRedraw, pagesToKeep, snapZoomIndex } from './logic.js';
import type { BodyKind, PageSize, SearchHit, ViewBody, ZoomMode } from './types.js';

export interface DrawOpts {
  /** 확대 %(ZOOM_STEPS) */
  readonly percent: number;
  /** 쪽 틀의 CSS 폭(px) */
  readonly cssWidth: number;
  /** min(devicePixelRatio, 2) */
  readonly dpr: number;
}

export interface PageSource {
  readonly kind: 'pages-png' | 'pages-pdfjs';
  readonly pageCount: number;
  /** 글자 층을 그리는가(pdf.js + canCopy) */
  readonly textSelectable: boolean;
  /** 쪽 크기(pt) — 아직 모르면 null */
  sizeOf(page: number): PageSize | null;
  /** 같은 값이면 다시 그리지 않는다(PNG: 서버 배율, pdf.js: 화소 폭) */
  qualityKey(page: number, opts: DrawOpts): number;
  /** canvas에 그린다(크기도 정한다). 실패하면 throw, 취소면 AbortError·RenderingCancelledException */
  draw(page: number, canvas: HTMLCanvasElement, opts: DrawOpts, signal: AbortSignal): Promise<void>;
  drawText?(page: number, layer: HTMLElement, cssWidth: number): Promise<void>;
  search(q: string, signal: AbortSignal): Promise<SearchHit[]>;
  thumbnail(page: number, width: number): Promise<HTMLCanvasElement | HTMLImageElement>;
  destroy(): void;
}

const MAX_CONCURRENT_DRAWS = 3;
const KEEP_RADIUS = 2;
const DEFAULT_SIZE: PageSize = { w: 612, h: 792 };

function isCancel(error: unknown): boolean {
  const name = (error as { name?: string } | null)?.name;
  return name === 'AbortError' || name === 'RenderingCancelledException';
}

export class PagesBody implements ViewBody {
  /** #pages(.pages-container) — mount 뒤에 쓴다 */
  container!: HTMLElement;
  zoomIndex = DEFAULT_ZOOM_INDEX;
  rotateDeg = 0;
  private frames: HTMLElement[] = [];
  private current = 1;
  private readonly visible = new Set<number>();
  private readonly drawnKey = new Map<number, number>();
  private readonly pendingKey = new Map<number, number>();
  private abort = new AbortController();
  private active = 0;
  private readonly waiters: Array<() => void> = [];
  private reloadTimer: ReturnType<typeof setTimeout> | null = null;
  private drawObserver: IntersectionObserver | null = null;
  private pageObserver: IntersectionObserver | null = null;
  private readonly pageCbs: Array<(page: number) => void> = [];
  private readonly zoomCbs: Array<(percent: number) => void> = [];
  private readonly drawCbs: Array<(page: number) => void> = [];
  private readonly errorCbs: Array<(error: unknown) => void> = [];
  private readonly sourceCbs: Array<() => void> = [];

  constructor(
    private source: PageSource,
    private readonly initialSizes: readonly PageSize[] | undefined,
  ) {}

  get kind(): BodyKind {
    return this.source.kind;
  }
  get textSelectable(): boolean {
    return this.source.textSelectable;
  }
  get currentPage(): number {
    return this.current;
  }
  get zoomPercent(): number {
    return ZOOM_STEPS[this.zoomIndex];
  }

  async mount(host: HTMLElement): Promise<void> {
    this.container = host;
    host.classList.add('pages-container');
    this.buildFrames();
    this.applyZoom();
  }

  pageCount(): number {
    return this.source.pageCount;
  }
  pageFrame(page: number): HTMLElement | null {
    return this.frames[page - 1] ?? null;
  }
  /** 쪽 크기(pt) — 원천이 아직 모르면 서버가 준 크기(info.pages), 그것도 없으면 Letter */
  pageSize(page: number): PageSize {
    return this.source.sizeOf(page) ?? this.initialSizes?.[page - 1] ?? this.initialSizes?.[0] ?? DEFAULT_SIZE;
  }

  onPage(cb: (page: number) => void): void {
    this.pageCbs.push(cb);
  }
  onZoom(cb: (percent: number) => void): void {
    this.zoomCbs.push(cb);
  }
  /** 쪽 그림이 (다시) 그려질 때마다 */
  onDraw(cb: (page: number) => void): void {
    this.drawCbs.push(cb);
  }
  /** 그리기 실패(취소 빼고) — view.ts가 pdf.js → PNG로 내려간다 */
  onDrawError(cb: (error: unknown) => void): void {
    this.errorCbs.push(cb);
  }
  /** replaceSource 뒤(쪽 틀을 다시 만들었다 — 겹침·썸네일을 다시 붙인다) */
  onSourceChange(cb: () => void): void {
    this.sourceCbs.push(cb);
  }

  replaceSource(next: PageSource): void {
    this.abort.abort();
    this.abort = new AbortController();
    this.source.destroy();
    this.source = next;
    this.buildFrames();
    this.applyZoom();
    for (const cb of this.sourceCbs) cb();
  }

  goto(page: number): void {
    const target = Math.min(Math.max(Math.round(page), 1), this.frames.length || 1);
    this.frames[target - 1]?.scrollIntoView({ block: 'start' });
    this.setCurrent(target);
  }

  zoom(z: ZoomMode): void {
    if (z === 'fit-width' || z === 'fit-page') {
      // 1120px 읽기 폭 제한을 풀어 창 끝까지(viewer.js setupZoomFit와 같다)
      this.container.classList.add('pages-edge-to-edge');
      const toolbarHeight = document.getElementById('toolbar')?.offsetHeight ?? 0;
      const percent = z === 'fit-width' ? 100 : fitPagePercent(this.container.clientWidth, window.innerHeight - toolbarHeight, this.pageSize(this.current));
      this.zoomIndex = snapZoomIndex(percent);
    } else {
      this.zoomIndex = snapZoomIndex(z);
    }
    this.applyZoom();
  }

  setZoomIndex(index: number): void {
    const next = clampZoomIndex(index);
    if (next === this.zoomIndex) return;
    this.zoomIndex = next;
    this.applyZoom();
  }

  /** 커서가 가리키는 곳을 그대로 둔 채 확대(viewer.js zoomAtPoint 그대로 — 세로는 창, 가로는 쪽 칸이 스크롤한다) */
  zoomAtPoint(nextIndex: number, clientX: number, clientY: number): void {
    const bounded = clampZoomIndex(nextIndex);
    if (bounded === this.zoomIndex) return;
    const ratio = ZOOM_STEPS[bounded] / ZOOM_STEPS[this.zoomIndex];
    const rect = this.container.getBoundingClientRect();
    const containerDocTop = rect.top + window.scrollY;
    const contentX = this.container.scrollLeft + (clientX - rect.left);
    const contentY = clientY + window.scrollY - containerDocTop;
    this.zoomIndex = bounded;
    this.applyZoom();
    this.container.scrollLeft = contentX * ratio - (clientX - rect.left);
    window.scrollTo(window.scrollX, containerDocTop + contentY * ratio - clientY);
  }

  /** 보기 전용 회전(저장 안 함) — 틀을 CSS로 돌린다(style.css .pages-rotated) */
  rotate(deg: 90 | -90): void {
    this.rotateDeg = (this.rotateDeg + deg + 360) % 360;
    this.container.style.setProperty('--page-rotate', `${this.rotateDeg}deg`);
    this.container.classList.toggle('pages-rotated', this.rotateDeg !== 0);
  }

  search(q: string, signal: AbortSignal): Promise<SearchHit[]> {
    return this.source.search(q, signal);
  }
  thumbnail(page: number, width: number): Promise<HTMLCanvasElement | HTMLImageElement> {
    return this.source.thumbnail(page, width);
  }

  /** 찾은 곳 강조(viewer.js renderPageHitOverlays 그대로) — 지금 결과의 첫 사각형을 돌려준다 */
  showHits(hits: readonly SearchHit[], current: number): HTMLElement | null {
    this.clearHits();
    let currentEl: HTMLElement | null = null;
    hits.forEach((hit, index) => {
      const frame = this.frames[hit.page - 1];
      if (!frame) return;
      for (const rect of hit.rects) {
        const el = document.createElement('div');
        el.className = index === current ? 'search-hit-overlay current' : 'search-hit-overlay';
        el.style.left = `${rect.x * 100}%`;
        el.style.top = `${rect.y * 100}%`;
        el.style.width = `${Math.max(rect.w * 100, 0.5)}%`;
        el.style.height = `${Math.max(rect.h * 100, 0.5)}%`;
        frame.append(el);
        if (index === current && !currentEl) currentEl = el;
      }
    });
    return currentEl;
  }

  clearHits(): void {
    for (const el of this.container.querySelectorAll('.search-hit-overlay')) el.remove();
  }

  destroy(): void {
    this.abort.abort();
    this.drawObserver?.disconnect();
    this.pageObserver?.disconnect();
    if (this.reloadTimer !== null) clearTimeout(this.reloadTimer);
    this.source.destroy();
  }

  private buildFrames(): void {
    this.drawObserver?.disconnect();
    this.pageObserver?.disconnect();
    this.visible.clear();
    this.drawnKey.clear();
    this.pendingKey.clear();
    // 화면 근처(한 화면 위아래)에 들어오면 그리고, 멀어지면 trim()이 비운다
    this.drawObserver = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          const page = Number((entry.target as HTMLElement).dataset.page);
          if (entry.isIntersecting) {
            this.visible.add(page);
            void this.ensureDrawn(page);
          } else {
            this.visible.delete(page);
          }
        }
        this.trim();
      },
      { rootMargin: '100% 0px' },
    );
    this.pageObserver = new IntersectionObserver(
      (entries) => {
        const shown = entries.find((entry) => entry.isIntersecting);
        if (shown) this.setCurrent(Number((shown.target as HTMLElement).dataset.page));
      },
      { threshold: 0.4 },
    );
    this.frames = [];
    for (let page = 1; page <= this.source.pageCount; page += 1) {
      const frame = document.createElement('div');
      frame.className = 'page-frame page-loading';
      frame.dataset.page = String(page);
      // 크기를 미리 알면 그리기 전에도 틀 높이가 맞아 쪽 이동·스크롤이 어긋나지 않는다
      const size = this.pageSize(page);
      frame.style.aspectRatio = `${size.w} / ${size.h}`;
      frame.style.minHeight = '0';
      const canvas = document.createElement('canvas');
      canvas.dataset.page = String(page);
      const label = document.createElement('span');
      label.className = 'page-number';
      label.textContent = String(page);
      frame.append(canvas, label);
      if (this.source.textSelectable) {
        const layer = document.createElement('div');
        layer.className = 'pdf-text-layer';
        frame.append(layer);
      }
      this.frames.push(frame);
      this.drawObserver.observe(frame);
      this.pageObserver.observe(frame);
    }
    this.container.replaceChildren(...this.frames);
    // 옵저버 콜백(다음 틱)을 기다리지 않고 1쪽부터(viewer.js buildPages와 같다)
    void this.ensureDrawn(1);
  }

  private applyZoom(): void {
    this.container.style.setProperty('--page-width', `${this.zoomPercent}%`);
    for (const cb of this.zoomCbs) cb(this.zoomPercent);
    this.scheduleQualityReload();
  }

  /** 확대가 멈춘 뒤 보이는 쪽을 새 배율로(viewer.js scheduleQualityReload와 같은 350ms) */
  private scheduleQualityReload(): void {
    if (this.reloadTimer !== null) clearTimeout(this.reloadTimer);
    this.reloadTimer = setTimeout(() => {
      this.reloadTimer = null;
      for (const page of this.visible) void this.ensureDrawn(page);
    }, 350);
  }

  private setCurrent(page: number): void {
    if (page === this.current) return;
    this.current = page;
    this.trim();
    for (const cb of this.pageCbs) cb(page);
  }

  private drawOpts(frame: HTMLElement): DrawOpts {
    return { percent: this.zoomPercent, cssWidth: frame.clientWidth || this.container.clientWidth, dpr: Math.min(window.devicePixelRatio || 1, 2) };
  }

  /** 대용량 쪽 그림을 한꺼번에 받으면 스크롤·단추 반응이 늦어진다 — 동시에 3장까지(viewer.js withLoadSlot) */
  private async withSlot(task: () => Promise<void>): Promise<void> {
    if (this.active >= MAX_CONCURRENT_DRAWS) await new Promise<void>((resolve) => this.waiters.push(resolve));
    this.active += 1;
    try {
      await task();
    } finally {
      this.active -= 1;
      this.waiters.shift()?.();
    }
  }

  private async ensureDrawn(page: number): Promise<void> {
    const frame = this.frames[page - 1];
    if (!frame) return;
    const opts = this.drawOpts(frame);
    const want = this.source.qualityKey(page, opts);
    if (this.pendingKey.get(page) === want || !needsRedraw(this.drawnKey.get(page) ?? 0, want)) return;
    this.pendingKey.set(page, want);
    const source = this.source;
    const signal = this.abort.signal;
    await this.withSlot(async () => {
      if (signal.aborted || source !== this.source || this.pendingKey.get(page) !== want) return;
      // 다 그린 canvas로 한 번에 바꿔 끼운다(그리는 동안 빈 쪽이 보이지 않게)
      const next = document.createElement('canvas');
      try {
        await source.draw(page, next, opts, signal);
      } catch (error) {
        this.pendingKey.delete(page);
        if (isCancel(error) || signal.aborted) return;
        frame.classList.add('page-error');
        for (const cb of this.errorCbs) cb(error);
        return;
      }
      if (signal.aborted || source !== this.source) return;
      next.dataset.page = String(page);
      frame.querySelector('canvas')?.replaceWith(next);
      frame.style.aspectRatio = `${next.width} / ${next.height}`;
      frame.classList.remove('page-loading', 'page-error');
      frame.classList.add('page-ready');
      this.drawnKey.set(page, want);
      this.pendingKey.delete(page);
      const layer = frame.querySelector<HTMLElement>('.pdf-text-layer');
      if (layer && source.drawText) await source.drawText(page, layer, opts.cssWidth).catch(() => undefined);
      for (const cb of this.drawCbs) cb(page);
    });
  }

  /** 멀어진 쪽의 캔버스를 비워 메모리를 돌려준다 — 틀의 가로세로 비율은 남아 스크롤 위치가 그대로다 */
  private release(page: number): void {
    const frame = this.frames[page - 1];
    if (!frame || !this.drawnKey.has(page)) return;
    const canvas = frame.querySelector('canvas');
    if (canvas) {
      canvas.width = 0;
      canvas.height = 0;
    }
    frame.querySelector('.pdf-text-layer')?.replaceChildren();
    this.drawnKey.delete(page);
    frame.classList.remove('page-ready');
    frame.classList.add('page-loading');
  }

  private trim(): void {
    const keep = pagesToKeep(this.current, this.frames.length, KEEP_RADIUS, this.visible);
    for (const page of [...this.drawnKey.keys()]) if (!keep.has(page)) this.release(page);
  }
}
```

- [ ] **Step 2: `src/client/view/pngSource.ts`**

```ts
/**
 * 서버 PNG 원천 — 내려받기 권한이 없는 토큰(설계 V5): 쪽 그림(/api/view/pages)·썸네일(?thumb=true)·찾기(/api/view/search)를
 * 지금 보기 화면과 같은 경로로 받는다. PDF·원본은 브라우저로 오지 않는다.
 */
import { SEARCH_LIMIT, pngScale } from './logic.js';
import type { DrawOpts, PageSource } from './pagesBody.js';
import type { PageSize, SearchHit } from './types.js';

export class PngSource implements PageSource {
  readonly kind = 'pages-png' as const;
  readonly textSelectable = false;

  constructor(
    private readonly tokenQuery: string,
    readonly pageCount: number,
    private readonly sizes: readonly PageSize[] | undefined,
  ) {}

  sizeOf(page: number): PageSize | null {
    return this.sizes?.[page - 1] ?? null;
  }

  qualityKey(_page: number, opts: DrawOpts): number {
    return pngScale(opts.percent, opts.dpr);
  }

  async draw(page: number, canvas: HTMLCanvasElement, opts: DrawOpts, signal: AbortSignal): Promise<void> {
    const res = await fetch(`/api/view/pages/${page}?${this.tokenQuery}&scale=${pngScale(opts.percent, opts.dpr)}`, { signal });
    if (!res.ok) throw new Error('쪽 그림을 불러오지 못했습니다.');
    const bitmap = await createImageBitmap(await res.blob());
    canvas.width = bitmap.width;
    canvas.height = bitmap.height;
    canvas.getContext('2d')?.drawImage(bitmap, 0, 0);
    bitmap.close();
  }

  async search(q: string, signal: AbortSignal): Promise<SearchHit[]> {
    const res = await fetch(`/api/view/search?${this.tokenQuery}&q=${encodeURIComponent(q)}&limit=${SEARCH_LIMIT}`, { signal });
    const json = (await res.json().catch(() => null)) as { success?: boolean; data?: { hits: SearchHit[] }; error?: string } | null;
    if (!res.ok || !json?.success || !json.data) throw new Error(json?.error || '검색에 실패했습니다.');
    return json.data.hits.map((hit) => ({ page: hit.page, rects: hit.rects }));
  }

  async thumbnail(page: number): Promise<HTMLImageElement> {
    const img = new Image();
    img.alt = `${page}쪽`;
    img.decoding = 'async';
    img.src = `/api/view/pages/${page}?${this.tokenQuery}&thumb=true`;
    await img.decode();
    return img;
  }

  destroy(): void {
    // 들고 있는 것이 없다
  }
}
```

- [ ] **Step 3: `src/client/view/ctx.ts`**

```ts
import type { PagesBody } from './pagesBody.js';
import type { ViewInfo } from './types.js';

/** 껍데기 모듈(막대·썸네일·찾기·표시·측정·편집 단추)이 함께 쓰는 것 — view.ts가 만든다 */
export interface ShellCtx {
  readonly token: string;
  /** `token=…`(인코딩됨) */
  readonly tokenQuery: string;
  readonly info: ViewInfo;
  /** 안쪽 — pdf.js가 실패해도 같은 객체가 원천만 바꾼다(replaceSource). 모듈은 body를 꺼내 두지 말고 ctx.body로 쓴다 */
  readonly body: PagesBody;
  /** Esc — priority가 작은 것부터(측정 10 · 쓰던 글 20 · 글 도구 30 · 강조 40 · 찾기 90), true를 돌려준 곳에서 멈춘다 */
  onEscape(priority: number, handler: () => boolean): void;
  /** 쪽 위 한 줄 안내(빈 글이면 숨김, 8초 뒤 저절로 숨김) */
  notice(text: string): void;
}
```

- [ ] **Step 4: `src/client/presence.ts` — 이름 목록으로 그리기**

파일 끝의 `renderPresenceAvatars` 함수(주석 포함)를 아래로 바꾼다(아바타 만들기를 한 곳으로 모으고, 보기 화면용 함수를 더한다 — 편집기 동작은 같다):

```ts
function presenceBadge(name: string, color: string): HTMLSpanElement {
  const avatar = document.createElement('span');
  avatar.className = 'presence-badge';
  avatar.style.setProperty('--presence-color', color);
  avatar.textContent = initialsOf(name);
  avatar.title = name;
  return avatar;
}

/** 오피스 365 스타일의 겹쳐진 원형 아바타 스택으로 현재 접속자를 렌더링한다. */
export function renderPresenceAvatars(container: HTMLElement, awareness: Awareness): void {
  const states = [...awareness.getStates().entries()] as Array<[number, AwarenessUserState]>;
  container.replaceChildren(
    ...dedupeByName(states).map(([, state]) => presenceBadge(state.user!.name!, state.user!.color ?? pickPresenceColorForName(state.user!.name!))),
  );
}

/** 이름 목록으로 같은 아바타를 그린다 — 보기 화면(view.ts)의 /api/view/presence 하트비트 응답용 */
export function renderPresenceNames(container: HTMLElement, names: readonly string[]): void {
  container.replaceChildren(...names.map((name) => presenceBadge(name, pickPresenceColorForName(name))));
}
```

- [ ] **Step 5: `public/view.html` — viewer.html에서 만든다**

워크트리 루트에서 아래 스크립트를 한 번 돌린다(계획 작성 때 같은 스크립트로 만들어 확인함, 229줄 → 190줄). 스크립트 파일은 커밋하지 않는다(scratchpad에 둔다).

```bash
python - <<'PY'
import re
src = open('public/viewer.html', encoding='utf8').read()
s = src

def sub(pattern, repl, flags=re.S):
    global s
    new, n = re.subn(pattern, repl, s, count=1, flags=flags)
    assert n == 1, pattern
    s = new

sub(r'<title>DocView — 뷰어</title>', '<title>DocView — 보기</title>')
sub(r'(<link rel="stylesheet" href="/css/style.css" />)', r'\1\n  <link rel="stylesheet" href="/css/view.css" />')
# 엑셀 격자·도면은 viewer.html에 남는다 — 이 화면에는 쪽 보기만
sub(r'\n  <!-- 도면\(dxf/dwg\) 벡터 뷰.*?</div>\n  </div>\n', '\n')
sub(r'\n  <main id="grid" class="grid-container" hidden>.*?</main>\n', '\n')
# 예전 "더보기" 메뉴·"페이지 보기(준비 중)" 자리는 도구▾·썸네일이 맡는다
sub(r'\n        <div id="toolbar-more-wrap" class="toolbar-more-wrap">.*?</div>\n        </div>\n', '\n')
sub(r'\n        <div class="toolbar-sep" aria-hidden="true"></div>\n        <button id="page-view-btn".*?</button>\n', '\n')
# 강조·글 도구는 숨기지 않고 권한이 없으면 흐리게(이유는 툴팁)
sub(r'(<button id="markup-add-btn"[^>]*?) hidden>', r'\1>')
sub(r'(<button id="text-tool-btn"[^>]*?) hidden>', r'\1>')
# 썸네일 단추 — 왼쪽 도구 묶음 맨 앞
sub(r'(<div id="toolbar-tools-left" class="[^"]*">)', r'''\1
        <button id="thumbs-btn" class="icon-btn" title="쪽 미리 보기" aria-label="쪽 미리 보기" aria-pressed="false">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="9" rx="1"/><rect x="3" y="15" width="7" height="6" rx="1"/><rect x="14" y="15" width="7" height="6" rx="1"/></svg>
        </button>
        <div class="toolbar-sep" aria-hidden="true"></div>''')
# 이전·다음 쪽
sub(r'(\n        <div class="page-jump-group")', r'''
        <button id="page-prev" class="icon-btn" title="이전 쪽" aria-label="이전 쪽">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m15 18-6-6 6-6"/></svg>
        </button>\1''')
sub(r'(<span id="page-jump-total" class="page-jump-total">의 1</span>\n        </div>)', r'''\1
        <button id="page-next" class="icon-btn" title="다음 쪽" aria-label="다음 쪽">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m9 18 6-6-6-6"/></svg>
        </button>''')
# 도구▾ — 오른쪽 묶음 끝(style.css의 예전 "더보기" 메뉴 모양을 그대로 쓴다)
sub(r'(</button>\n      </div>\n    </div>\n  </header>)', r'''</button>
        <div id="tools-wrap" class="toolbar-more-wrap">
          <button id="tools-btn" class="icon-btn tools-btn" title="도구" aria-label="도구" aria-haspopup="true" aria-expanded="false" hidden>도구 ▾</button>
          <div id="tools-menu" class="toolbar-more-menu" role="menu" hidden></div>
        </div>
      </div>
    </div>
  </header>''')
# 편집 단추 — 제목 줄, 함께 보는 사람 앞(설계 V6)
sub(r'(\n      <div id="presence-list")', r'''
      <button id="edit-btn" class="btn-pill btn-pill-primary edit-btn" type="button" hidden>편집</button>\1''')
# 썸네일 칸·한 줄 안내·준비 화면 — 쪽 칸 앞
sub(r'(\n  <main id="pages" class="pages-container"></main>)', r'''
  <aside id="thumbs-panel" class="thumbs-panel" aria-label="쪽 미리 보기" hidden></aside>
  <div id="view-notice" class="view-notice" role="status" hidden></div>
  <div id="prep-screen" class="prep-screen" role="status" hidden>
    <p>쪽을 준비하는 중</p>
    <progress id="prep-progress" max="100" value="0"></progress>
  </div>\1''')
# 오류 화면 — 편집기가 있으면 "편집 화면으로 열기"(설계 §7)
sub(r'(<p id="error-message"></p>)', r'''\1
      <a id="error-edit" class="btn-pill" hidden>편집 화면으로 열기</a>''')
sub(r'<script src="/js/viewer.js"></script>', '<script type="module" src="/js/view.bundle.js"></script>')
for bad in ('id="grid', 'svg-view', '/js/viewer.js'):
    assert bad not in s, bad
open('public/view.html', 'w', encoding='utf8', newline='\n').write(s)
print(len(src.splitlines()), '->', len(s.splitlines()))
PY
grep -o 'id="[^"]*"' public/view.html | tr '\n' ' '
```

Expected 마지막 줄의 id(이후 작업이 쓰는 DOM 약속): `toolbar back-btn doc-title edit-btn presence-list toolbar-tools-left thumbs-btn measure-btn invert-btn markup-add-btn text-tool-btn markup-include-wrap markup-include-checkbox zoom-out zoom-in zoom-fit-btn zoom-level page-prev page-jump-input page-jump-total page-next rotate-btn search-btn print-btn download-btn download-as-btn fullscreen-btn tools-wrap tools-btn tools-menu measure-bar measure-hint measure-calibrate measure-clear measure-close search-bar search-input search-count search-prev search-next search-close thumbs-panel view-notice prep-screen prep-progress pages error-screen error-message error-edit`. 인라인 스크립트(home=hidden·back=parent → `docview:back`)는 viewer.html 것이 그대로 남는다.

- [ ] **Step 6: `public/css/view.css`(이 작업분 — 이후 작업이 뒤에 더한다)**

```css
/* 통합 보기 껍데기(view.html) — style.css의 보기 화면(viewer.html) 모양 위에 새 부분만 둔다 */

/* 편집 단추(설계 V6) — 제목 줄 오른쪽, 함께 보는 사람 앞 */
.edit-btn { margin-left: auto; margin-right: 0.6rem; }

/* 한 줄 안내(pdf.js → 서버 그림 등) */
.view-notice {
  position: fixed;
  top: 6.4rem;
  left: 50%;
  transform: translateX(-50%);
  z-index: 12;
  max-width: calc(100vw - 2rem);
  padding: 0.35rem 0.9rem;
  border-radius: 6px;
  background: var(--color-surface);
  border: 1px solid var(--v-glass-line);
  box-shadow: 0 4px 14px rgba(0, 0, 0, 0.12);
  font-size: 0.85rem;
}

#error-edit { margin-top: 0.6rem; }

/* 아래 칸들은 display를 정하므로 hidden 속성이 이기도록 따로 적는다(계획 작성 때 hidden인데 보이던 것을 확인함) */
.thumbs-panel[hidden], .prep-screen[hidden], .view-notice[hidden] { display: none; }
```

- [ ] **Step 7: `src/client/view.ts`(이 작업분)**

```ts
/**
 * 통합 보기 껍데기(public/view.html) 입구 — 설계 docs/superpowers/specs/2026-11-12-unified-view-shell-design.md.
 * 쪽 문서(워드·PPT·PDF 등, 한글 빼고)를 같은 모양으로 보여 주고, 편집은 형식별 편집기로 넘긴다. 엑셀 격자·도면·한글은
 * 지금 보기 화면(viewer.html)으로 넘긴다(서버 viewerUrl 기준).
 */
import { renderPresenceNames } from './presence.js';
import { applyWatermark } from './watermark.js';
import type { ShellCtx } from './view/ctx.js';
import { createEscapeChain, pathOf } from './view/logic.js';
import { PagesBody, type PageSource } from './view/pagesBody.js';
import { PngSource } from './view/pngSource.js';
import type { ViewInfo } from './view/types.js';

const token = new URLSearchParams(location.search).get('token') ?? '';
const tokenQuery = `token=${encodeURIComponent(token)}`;
const PRESENCE_HEARTBEAT_MS = 8000;
const NOTICE_MS = 8000;

let readyPosted = false;
/** 부모 창(향기랩·연동 제품)에 "다 그렸다" — 편집기들의 docview:ready와 같다 */
function postReady(): void {
  if (readyPosted || window.parent === window) return;
  readyPosted = true;
  requestAnimationFrame(() => requestAnimationFrame(() => window.parent.postMessage({ type: 'docview:ready' }, '*')));
}

function showError(message: string): void {
  (document.getElementById('toolbar') as HTMLElement).hidden = true;
  (document.getElementById('pages') as HTMLElement).hidden = true;
  (document.getElementById('error-screen') as HTMLElement).hidden = false;
  (document.getElementById('error-message') as HTMLElement).textContent = message;
  postReady();
}

let noticeTimer: ReturnType<typeof setTimeout> | null = null;
function notice(text: string): void {
  const el = document.getElementById('view-notice') as HTMLElement;
  el.textContent = text;
  el.hidden = text === '';
  if (noticeTimer !== null) clearTimeout(noticeTimer);
  noticeTimer = text === '' ? null : setTimeout(() => (el.hidden = true), NOTICE_MS);
}

async function fetchInfo(): Promise<ViewInfo> {
  let res: Response;
  try {
    res = await fetch(`/api/view/info?${tokenQuery}`);
  } catch {
    throw new Error('서버에 연결할 수 없습니다.');
  }
  const json = (await res.json().catch(() => null)) as { success?: boolean; data?: ViewInfo; error?: string } | null;
  if (!res.ok || !json?.success || !json.data) throw new Error(json?.error || '문서를 열 수 없습니다.');
  return json.data;
}

/** "지금 이 문서를 보고 있는 사람" — 지금 보기 화면과 같은 8초 하트비트(/api/view/presence), 그림은 presence.ts */
function startPresence(): void {
  const container = document.getElementById('presence-list');
  if (!container) return;
  const sessionId = globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random()}`;
  const beat = async () => {
    try {
      const res = await fetch(`/api/view/presence?${tokenQuery}`, {
        method: 'POST',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ sessionId }),
      });
      const json = (await res.json()) as { success?: boolean; data?: { viewers: string[] } };
      if (json.success && json.data) renderPresenceNames(container, json.data.viewers);
    } catch {
      // 접속자 표시 실패는 열람을 막지 않는다
    }
  };
  void beat();
  setInterval(() => void beat(), PRESENCE_HEARTBEAT_MS);
}

/** 우클릭·저장/인쇄 단축키 막기는 억지력일 뿐 — 실제 권한은 서버가 지킨다(viewer.js setupDeterrents와 같다) */
function setupDeterrents(ctx: ShellCtx, runEscape: () => boolean): void {
  document.addEventListener('contextmenu', (event) => {
    // 글자 선택이 되는 화면(pdf.js + 복사 권한)에서는 글자 층의 오른쪽 단추(복사)를 막지 않는다
    if (ctx.body.textSelectable && (event.target as Element | null)?.closest?.('.pdf-text-layer')) return;
    event.preventDefault();
  });
  document.addEventListener('keydown', (event) => {
    const key = event.key.toLowerCase();
    if ((event.ctrlKey || event.metaKey) && (key === 'p' || key === 's')) event.preventDefault();
    if (key === 'escape') runEscape();
  });
}

/** 권한에 맞는 원천 — 이 작업에서는 PNG만(Task 4에서 내려받기 권한이면 pdf.js) */
async function openSource(_info: ViewInfo, png: () => PageSource): Promise<PageSource> {
  return png();
}

async function boot(): Promise<void> {
  if (!token) {
    showError('열람 토큰이 없습니다. 발급받은 링크로 접속해 주세요.');
    return;
  }
  let info: ViewInfo;
  try {
    info = await fetchInfo();
  } catch (error) {
    showError(error instanceof Error ? error.message : '문서를 열 수 없습니다.');
    return;
  }
  // 이 껍데기가 맡지 않는 형식은 지금 보기 화면으로(같은 조회 문자열) — 기준은 서버 viewerUrl(viewPathFor)
  if (pathOf(info.viewerUrl) !== '/view.html') {
    location.replace(`/viewer.html${location.search}`);
    return;
  }
  (document.getElementById('doc-title') as HTMLElement).textContent = info.name;
  document.title = `DocView — ${info.name}`;
  applyWatermark(info.viewer);
  startPresence();

  const png = () => new PngSource(tokenQuery, info.pageCount, info.pages);
  const body = new PagesBody(await openSource(info, png), info.pages);
  const escape = createEscapeChain();
  const ctx: ShellCtx = { token, tokenQuery, info, body, onEscape: escape.add, notice };
  setupDeterrents(ctx, escape.run);
  body.onDraw(() => postReady());
  await body.mount(document.getElementById('pages') as HTMLElement);
}

boot().catch((error) => {
  console.error('보기 화면을 열지 못했습니다:', error);
  showError('문서를 열 수 없습니다.');
});
```

(`let info`는 Task 11에서 다시 받는다. `ctx`는 Task 5부터 `setupX(ctx)`에 쓴다 — 이 작업에서는 쓰지 않는 변수 경고가 없도록 `setupDeterrents`가 받는다.)

- [ ] **Step 8: 빌드 입구·무시 목록**

`scripts/build-editor.mjs`의 `entryPoints` 배열에서 `'src/client/editor-pdf.ts',` 다음 줄에:

```js
    // 통합 보기 껍데기(public/view.html) — pdf.js 원천은 동적 import라 PNG 모드에서는 받지 않는다
    'src/client/view.ts',
```

`.gitignore`의 `public/js/editor-pdf.bundle.js.map` 다음 줄에:

```
public/js/view.bundle.js
public/js/view.bundle.js.map
```

- [ ] **Step 9: 확인**

Run: `npx tsc --noEmit && npx vitest run && npm run build:editor && ls -la public/js/view.bundle.js && git status --short`
Expected: tsc 0, vitest 전부 PASS(presence.test.ts 포함), 빌드 성공, `view.bundle.js` 수십 KB 이하(계획 작성 때 정적 부분 26 KB + 공유 조각 몇 KB), `git status`에 bundle·chunks가 나오지 않는다.

- [ ] **Step 10: 커밋**

```bash
git add src/client/view/pagesBody.ts src/client/view/pngSource.ts src/client/view/ctx.ts src/client/view.ts public/view.html public/css/view.css
git commit -m "$(cat <<'EOF'
통합 보기 껍데기 골격: view.html·view.bundle.js, 쪽 틀 몸체(PagesBody)와 서버 PNG 원천, 워터마크·함께 보는 사람은 공용 모듈

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/client/view/pagesBody.ts src/client/view/pngSource.ts src/client/view/ctx.ts src/client/view.ts src/client/presence.ts public/view.html public/css/view.css scripts/build-editor.mjs .gitignore
```

---

### Task 4: pdf.js 원천 — 내려받기 권한일 때, 글자 층, 실패하면 PNG

**Files:**
- Create: `src/client/view/pdfjsSource.ts`
- Modify: `src/client/view.ts`(`openSource`, `boot`의 `body.onDraw` 다음), `public/css/view.css`(끝에 더하기)

**Interfaces:**
- Consumes: `openPdfData`·`pdfjs`(`src/client/pdf/pdfDocument.ts`), `findInPage`(`findText.ts`), `matchRects`·`canvasPixelRatio`·`SEARCH_LIMIT`·`chooseBodyKind`(logic.ts), `PageSource`·`DrawOpts`(pagesBody.ts)
- Produces: `openPdfjsSource(tokenQuery, canCopy, signal): Promise<PdfjsSource>`, `class PdfjsSource implements PageSource`, `class DerivedPdfUnavailable extends Error`

- [ ] **Step 1: `src/client/view/pdfjsSource.ts`**

```ts
/**
 * pdf.js 원천 — 내려받기 권한이 있을 때만(설계 V5) /api/view/derived-pdf로 받은 PDF를 브라우저에서 그린다. 여는 방법·보안 설정은
 * PDF 편집기와 같은 pdfDocument.ts(PDF 안 자바스크립트 실행 안 함, worker·cMap·표준 글꼴 자체 호스팅), 찾기 정규화는 findText.ts.
 * view.ts가 동적 import로만 부른다 — PNG 모드에서는 pdf.js 묶음을 받지 않는다.
 */
import { findInPage } from '../pdf/findText.js';
import { openPdfData, pdfjs, type PDFDocumentProxy } from '../pdf/pdfDocument.js';
import { SEARCH_LIMIT, canvasPixelRatio, matchRects, type TextItemGeom, type ViewportGeom } from './logic.js';
import type { DrawOpts, PageSource } from './pagesBody.js';
import type { PageSize, SearchHit } from './types.js';

type PageProxy = Awaited<ReturnType<PDFDocumentProxy['getPage']>>;
type Viewport = ReturnType<PageProxy['getViewport']>;

/** 받을 수 없음(403·202·423 등) — view.ts가 PNG로 내려간다 */
export class DerivedPdfUnavailable extends Error {}

export async function openPdfjsSource(tokenQuery: string, canCopy: boolean, signal: AbortSignal): Promise<PdfjsSource> {
  const res = await fetch(`/api/view/derived-pdf?${tokenQuery}`, { cache: 'no-store', signal });
  if (res.status !== 200) throw new DerivedPdfUnavailable(String(res.status));
  const doc = await openPdfData(new Uint8Array(await res.arrayBuffer()));
  return new PdfjsSource(doc, canCopy);
}

export class PdfjsSource implements PageSource {
  readonly kind = 'pages-pdfjs' as const;
  private readonly sizes: Array<PageSize | null>;
  private readonly texts = new Map<number, Promise<{ items: TextItemGeom[]; vp: ViewportGeom }>>();

  constructor(
    private readonly doc: PDFDocumentProxy,
    /** 글자 층은 복사 권한(canCopy)이 있을 때만(설계 §4) */
    readonly textSelectable: boolean,
  ) {
    this.sizes = new Array<PageSize | null>(doc.numPages).fill(null);
  }

  get pageCount(): number {
    return this.doc.numPages;
  }

  sizeOf(page: number): PageSize | null {
    return this.sizes[page - 1] ?? null;
  }

  qualityKey(_page: number, opts: DrawOpts): number {
    return Math.round(opts.cssWidth * opts.dpr);
  }

  private async open(page: number): Promise<{ proxy: PageProxy; base: Viewport }> {
    const proxy = await this.doc.getPage(page);
    const base = proxy.getViewport({ scale: 1 });
    this.sizes[page - 1] = { w: base.width, h: base.height };
    return { proxy, base };
  }

  async draw(page: number, canvas: HTMLCanvasElement, opts: DrawOpts, signal: AbortSignal): Promise<void> {
    const { proxy, base } = await this.open(page);
    const viewport = proxy.getViewport({ scale: opts.cssWidth / base.width });
    const ratio = canvasPixelRatio(viewport.width, viewport.height, opts.dpr);
    canvas.width = Math.floor(viewport.width * ratio);
    canvas.height = Math.floor(viewport.height * ratio);
    const ctx = canvas.getContext('2d', { alpha: false });
    if (!ctx) throw new Error('쪽을 그릴 캔버스를 만들지 못했습니다.');
    const task = proxy.render({ canvasContext: ctx, viewport, transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined });
    const cancel = () => task.cancel();
    signal.addEventListener('abort', cancel, { once: true });
    try {
      await task.promise;
    } finally {
      signal.removeEventListener('abort', cancel);
    }
  }

  async drawText(page: number, layer: HTMLElement, cssWidth: number): Promise<void> {
    const { proxy, base } = await this.open(page);
    const viewport = proxy.getViewport({ scale: cssWidth / base.width });
    layer.replaceChildren();
    layer.style.setProperty('--scale-factor', String(viewport.scale));
    await pdfjs.renderTextLayer({ textContentSource: proxy.streamTextContent(), container: layer, viewport, textDivs: [] }).promise;
  }

  private textOf(page: number): Promise<{ items: TextItemGeom[]; vp: ViewportGeom }> {
    let hit = this.texts.get(page);
    if (!hit) {
      hit = this.open(page).then(async ({ proxy, base }) => {
        const content = await proxy.getTextContent();
        const items: TextItemGeom[] = [];
        for (const item of content.items) {
          if ('str' in item) items.push({ str: item.str, transform: item.transform as number[], width: item.width, height: item.height });
        }
        return { items, vp: { transform: base.transform, width: base.width, height: base.height } };
      });
      this.texts.set(page, hit);
    }
    return hit;
  }

  async search(q: string, signal: AbortSignal): Promise<SearchHit[]> {
    const hits: SearchHit[] = [];
    for (let page = 1; page <= this.doc.numPages && hits.length < SEARCH_LIMIT; page += 1) {
      if (signal.aborted) throw new DOMException('찾기를 멈췄습니다.', 'AbortError');
      const { items, vp } = await this.textOf(page);
      for (const match of findInPage(page, items, q)) {
        hits.push({ page, rects: matchRects(match, items, vp) });
        if (hits.length >= SEARCH_LIMIT) break;
      }
    }
    return hits;
  }

  async thumbnail(page: number, width: number): Promise<HTMLCanvasElement> {
    const { proxy, base } = await this.open(page);
    const viewport = proxy.getViewport({ scale: width / base.width });
    const ratio = Math.min(2, window.devicePixelRatio || 1);
    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(viewport.width * ratio);
    canvas.height = Math.floor(viewport.height * ratio);
    const ctx = canvas.getContext('2d', { alpha: false });
    if (!ctx) throw new Error('썸네일을 그릴 캔버스를 만들지 못했습니다.');
    await proxy.render({ canvasContext: ctx, viewport, transform: ratio !== 1 ? [ratio, 0, 0, ratio, 0, 0] : undefined }).promise;
    return canvas;
  }

  destroy(): void {
    void this.doc.destroy();
  }
}
```

- [ ] **Step 2: `src/client/view.ts` — 권한에 맞는 원천과 PNG로 내려가기**

import에서 `import { createEscapeChain, pathOf } from './view/logic.js';` → `import { chooseBodyKind, createEscapeChain, pathOf } from './view/logic.js';`.

`openSource` 함수를 통째로 바꾼다:

```ts
const FALLBACK_NOTICE = 'PDF를 바로 그리지 못해 서버 그림으로 보여 줍니다 — 글자 선택은 되지 않습니다.';

/** 권한에 맞는 원천 — 내려받기 권한이 있으면 pdf.js(동적 import), 받지 못하면 PNG와 한 줄 안내 */
async function openSource(info: ViewInfo, png: () => PageSource): Promise<PageSource> {
  if (chooseBodyKind(info.permissions) === 'pages-png') return png();
  try {
    const { openPdfjsSource } = await import('./view/pdfjsSource.js');
    return await openPdfjsSource(tokenQuery, info.permissions.canCopy, new AbortController().signal);
  } catch (error) {
    console.warn('pdf.js로 열지 못해 서버 그림으로 보여 줍니다:', error);
    notice(FALLBACK_NOTICE);
    return png();
  }
}
```

`boot()`에서 `body.onDraw(() => postReady());` 다음 줄에:

```ts
  // pdf.js로 그리다 실패하면(깨진 파일·휴대폰 메모리) 같은 자리에 서버 PNG로(설계 §4·§7)
  body.onDrawError(() => {
    if (body.kind !== 'pages-pdfjs') return;
    notice(FALLBACK_NOTICE);
    body.replaceSource(png());
  });
```

- [ ] **Step 3: `public/css/view.css` 끝에 — 글자 층**

```css
/* pdf.js 글자 층(복사 권한이 있을 때만) — pdf-editor.css의 .pdf-text-layer 규칙과 같은 값. 글자는 투명, 고른 곳만 보인다 */
.pdf-text-layer { position: absolute; inset: 0; overflow: hidden; text-align: initial; opacity: 1; line-height: 1; text-size-adjust: none; -webkit-text-size-adjust: none; forced-color-adjust: none; transform-origin: 0 0; caret-color: CanvasText; z-index: 2; }
.pdf-text-layer :is(span, br) { color: transparent; position: absolute; white-space: pre; cursor: text; transform-origin: 0% 0%; }
.pdf-text-layer span.markedContent { top: 0; height: 0; }
.pdf-text-layer ::selection { background: rgba(26, 115, 232, 0.28); }
.pdf-text-layer br::selection { background: transparent; }
.pdf-text-layer .endOfContent { display: block; position: absolute; inset: 100% 0 0; z-index: -1; cursor: default; user-select: none; }
.pdf-text-layer .endOfContent.active { top: 0; }
/* 측정·강조·글 넣기 중에는 글자 층이 누름을 가로채지 않게 */
.pages-container.measuring .pdf-text-layer,
.pages-container.markup-drawing .pdf-text-layer,
.pages-container.text-tool-active .pdf-text-layer { pointer-events: none; }
```

(쪽 틀은 CSS로 돌리므로 pdf-editor.css의 `data-main-rotation` 규칙은 필요 없다 — 글자 층은 늘 회전 0으로 그린다.)

- [ ] **Step 4: 확인**

Run: `npx tsc --noEmit && npx vitest run && npm run build:editor && ls public/js/chunks | grep -i pdfjsSource && grep -c "pdfjs-dist\|GlobalWorkerOptions" public/js/view.bundle.js`
Expected: tsc 0, vitest PASS, `chunks/pdfjsSource-<해시>.js`가 따로 생기고 `view.bundle.js` 안에는 pdf.js 본체가 없다(마지막 grep 0 — 계획 작성 때 정적 묶음 26 KB로 확인).

- [ ] **Step 5: 커밋**

```bash
git add src/client/view/pdfjsSource.ts
git commit -m "$(cat <<'EOF'
통합 보기: 내려받기 권한이면 pdf.js로 그리기(동적 묶음), 복사 권한이면 글자 층, 실패하면 서버 PNG로

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/client/view/pdfjsSource.ts src/client/view.ts public/css/view.css
```

---

### Task 5: 막대 — 확대·맞춤·쪽 이동·회전·전체 화면·끌기·"도구▾"

**Files:**
- Create: `src/client/view/toolbar.ts`
- Modify: `src/client/view.ts`(import, `body.mount` 다음), `public/css/view.css`(끝에 더하기)

**Interfaces:**
- Consumes: `ShellCtx`, `PagesBody`의 `onZoom`·`setZoomIndex`·`zoom`·`zoomAtPoint`·`goto`·`onPage`·`onSourceChange`·`rotate`·`container`·`textSelectable`, `ZOOM_STEPS`·`menuIdsFor`
- Produces: `setupToolbar(ctx: ShellCtx): void`. 다른 모듈과의 약속: 측정 중에는 `#pages`에 `.measuring` 클래스(Task 8)가 붙어 있고 그동안 두 번 누르기 확대·끌기를 하지 않는다.

- [ ] **Step 1: `src/client/view/toolbar.ts`**

```ts
/**
 * 막대 — 확대·맞춤·쪽 이동·회전·전체 화면, Ctrl+휠·두 번 누르기 확대, 끌어서 옮기기, 좁은 화면의 "도구▾" 메뉴.
 * 동작은 지금 보기 화면(viewer.js setupToolbar·setupZoomFit·setupPageJump·setupRotate·setupFullscreen·setupDrawingTools)과 같다.
 */
import type { ShellCtx } from './ctx.js';
import { ZOOM_STEPS, menuIdsFor } from './logic.js';

const $ = <T extends HTMLElement = HTMLElement>(id: string): T => document.getElementById(id) as T;

export function setupToolbar(ctx: ShellCtx): void {
  setupZoom(ctx);
  setupPageNav(ctx);
  $('rotate-btn').addEventListener('click', () => ctx.body.rotate(90));
  setupFullscreen();
  setupPointer(ctx);
  setupToolsMenu();
}

function setupZoom(ctx: ShellCtx): void {
  const zoomIn = $<HTMLButtonElement>('zoom-in');
  const zoomOut = $<HTMLButtonElement>('zoom-out');
  const level = $('zoom-level');
  const sync = (percent: number) => {
    level.textContent = `${percent}%`;
    zoomOut.disabled = ctx.body.zoomIndex === 0;
    zoomIn.disabled = ctx.body.zoomIndex === ZOOM_STEPS.length - 1;
  };
  ctx.body.onZoom(sync);
  sync(ctx.body.zoomPercent);
  zoomIn.addEventListener('click', () => ctx.body.setZoomIndex(ctx.body.zoomIndex + 1));
  zoomOut.addEventListener('click', () => ctx.body.setZoomIndex(ctx.body.zoomIndex - 1));

  // 맞춤 토글 — 단추는 "다음에 누르면 적용될" 모드를 보인다(viewer.js와 같다)
  const fit = $<HTMLButtonElement>('zoom-fit-btn');
  let next: 'fit-width' | 'fit-page' = 'fit-width';
  fit.addEventListener('click', () => {
    ctx.body.zoom(next);
    next = next === 'fit-width' ? 'fit-page' : 'fit-width';
    const pageNext = next === 'fit-page';
    // svg 요소는 hidden 프로퍼티가 속성에 반영되지 않는 엔진이 있어 속성을 직접 바꾼다
    fit.querySelector('.zoom-fit-icon-width')?.toggleAttribute('hidden', pageNext);
    fit.querySelector('.zoom-fit-icon-page')?.toggleAttribute('hidden', !pageNext);
    const label = pageNext ? '페이지에 맞추기' : '너비에 맞추기';
    fit.title = label;
    fit.setAttribute('aria-label', label);
    fit.setAttribute('aria-pressed', 'true');
    setTimeout(() => fit.setAttribute('aria-pressed', 'false'), 400);
  });
}

function setupPageNav(ctx: ShellCtx): void {
  const input = $<HTMLInputElement>('page-jump-input');
  const total = $('page-jump-total');
  const prev = $<HTMLButtonElement>('page-prev');
  const next = $<HTMLButtonElement>('page-next');
  const show = (page: number) => {
    if (document.activeElement !== input) input.value = String(page);
    total.textContent = `의 ${ctx.body.pageCount()}`;
    prev.disabled = page <= 1;
    next.disabled = page >= ctx.body.pageCount();
  };
  ctx.body.onPage(show);
  ctx.body.onSourceChange(() => show(ctx.body.currentPage));
  show(ctx.body.currentPage);
  const jump = () => {
    const typed = Math.round(Number(input.value));
    if (Number.isFinite(typed)) ctx.body.goto(typed);
    input.value = String(ctx.body.currentPage);
  };
  input.addEventListener('keydown', (event) => {
    if (event.key === 'Enter') {
      event.preventDefault();
      jump();
      input.blur();
    } else if (event.key === 'Escape') {
      input.blur();
    }
  });
  input.addEventListener('blur', jump);
  input.addEventListener('focus', () => input.select());
  prev.addEventListener('click', () => ctx.body.goto(ctx.body.currentPage - 1));
  next.addEventListener('click', () => ctx.body.goto(ctx.body.currentPage + 1));
}

/** 전체 화면 — Esc·브라우저 메뉴로 나가는 경우도 fullscreenchange로 맞춘다(viewer.js setupFullscreen) */
function setupFullscreen(): void {
  const btn = $<HTMLButtonElement>('fullscreen-btn');
  if (!document.fullscreenEnabled) {
    btn.disabled = true;
    btn.title = '이 브라우저에서는 전체 화면을 쓸 수 없습니다';
    return;
  }
  const sync = () => {
    const on = Boolean(document.fullscreenElement);
    btn.querySelector('.fullscreen-icon-enter')?.toggleAttribute('hidden', on);
    btn.querySelector('.fullscreen-icon-exit')?.toggleAttribute('hidden', !on);
    const label = on ? '전체화면 종료' : '전체화면 보기';
    btn.title = label;
    btn.setAttribute('aria-label', label);
    btn.setAttribute('aria-pressed', String(on));
  };
  btn.addEventListener('click', () => {
    if (document.fullscreenElement) void document.exitFullscreen();
    else document.documentElement.requestFullscreen().catch(() => undefined);
  });
  document.addEventListener('fullscreenchange', sync);
  sync();
}

/** Ctrl+휠·두 번 누르기 확대, 끌어서 옮기기(viewer.js setupDrawingTools 앞부분) */
function setupPointer(ctx: ShellCtx): void {
  const pages = ctx.body.container;
  pages.addEventListener(
    'wheel',
    (event) => {
      if (!event.ctrlKey) return;
      event.preventDefault();
      ctx.body.zoomAtPoint(ctx.body.zoomIndex + (event.deltaY < 0 ? 1 : -1), event.clientX, event.clientY);
    },
    { passive: false },
  );
  pages.addEventListener('dblclick', (event) => {
    if (pages.classList.contains('measuring')) return;
    ctx.body.zoomAtPoint(ctx.body.zoomIndex + 2, event.clientX, event.clientY);
  });
  let pan: { x: number; y: number; left: number; top: number } | null = null;
  pages.addEventListener('mousedown', (event) => {
    if (pages.classList.contains('measuring')) return;
    if (event.button !== 0 && event.button !== 1) return;
    const target = event.target as Element;
    // 글자 선택이 되는 화면에서 글자 위를 누르면 끌기 대신 글자 고르기
    if (ctx.body.textSelectable && target.closest('.pdf-text-layer span')) return;
    // 칸 스크롤바를 누른 것은 브라우저에 맡긴다
    if (target === pages && (event.offsetX > pages.clientWidth || event.offsetY > pages.clientHeight)) return;
    pan = { x: event.clientX, y: event.clientY, left: pages.scrollLeft, top: window.scrollY };
    pages.classList.add('panning');
    event.preventDefault();
  });
  document.addEventListener('mousemove', (event) => {
    if (!pan) return;
    pages.scrollLeft = pan.left - (event.clientX - pan.x);
    window.scrollTo(window.scrollX, pan.top - (event.clientY - pan.y));
  });
  document.addEventListener('mouseup', () => {
    pan = null;
    pages.classList.remove('panning');
  });
}

/**
 * "도구▾" — 화면 폭에 따라 정해진 단추(logic.ts menuIdsFor)를 메뉴로 옮기고, 넓어지면 제자리로 돌린다. 요소를 그대로
 * 옮기므로 걸어 둔 클릭 처리가 그대로 산다(viewer.js toolbar-more와 같은 방법).
 */
function setupToolsMenu(): void {
  const btn = $<HTMLButtonElement>('tools-btn');
  const menu = $('tools-menu');
  const home = new Map<string, { parent: Node; next: Node | null }>();
  const close = () => {
    menu.hidden = true;
    btn.setAttribute('aria-expanded', 'false');
  };
  const layout = () => {
    const wanted = new Set(menuIdsFor(window.innerWidth));
    // 메뉴에 있을 필요가 없어진 것부터 제자리로 — 뒤에서부터 넣어 기억해 둔 다음 형제가 먼저 돌아와 있게
    for (const [id, place] of [...home.entries()].reverse()) {
      if (wanted.has(id)) continue;
      const el = document.getElementById(id);
      // 기억해 둔 다음 형제가 아직 메뉴에 있으면 끝에 붙인다(insertBefore가 던지지 않게)
      if (el) place.parent.insertBefore(el, place.next?.parentNode === place.parent ? place.next : null);
      home.delete(id);
    }
    for (const id of wanted) {
      const el = document.getElementById(id);
      if (!el || home.has(id)) continue;
      home.set(id, { parent: el.parentNode as Node, next: el.nextSibling });
      menu.append(el);
    }
    btn.hidden = wanted.size === 0;
    if (wanted.size === 0) close();
  };
  btn.addEventListener('click', (event) => {
    event.stopPropagation();
    menu.hidden = !menu.hidden;
    btn.setAttribute('aria-expanded', String(!menu.hidden));
  });
  // 메뉴 안 단추를 누르면 그 동작을 한 뒤 닫는다
  menu.addEventListener('click', (event) => {
    if ((event.target as Element).closest('button')) close();
  });
  document.addEventListener('click', (event) => {
    if (!menu.hidden && !(event.target as Element).closest('#tools-wrap')) close();
  });
  window.addEventListener('resize', layout);
  layout();
}
```

- [ ] **Step 2: `src/client/view.ts`**

import 추가: `import { setupToolbar } from './view/toolbar.js';`
`boot()`의 `await body.mount(document.getElementById('pages') as HTMLElement);` 다음 줄에 `setupToolbar(ctx);`

- [ ] **Step 3: `public/css/view.css` 끝에 — 좁은 화면 막대**

```css
.tools-btn { width: auto; padding: 0 0.6rem; font-size: 0.85rem; }

@media (max-width: 639px) {
  /* 왼쪽 도구는 모두 도구▾로 들어가 비어 있다 — 막대에는 확대·쪽·찾기·도구▾만(설계 §2). 계획 작성 때 390px에서 도구▾가
     화면 밖으로 밀리던 것을 이 규칙으로 맞춤 */
  #toolbar-tools-left { display: none; }
  .toolbar-tools-center { flex: 1 1 auto; justify-content: flex-start; gap: 0.15rem; }
  .toolbar-tools-center .toolbar-sep { display: none; }
  .toolbar-tools-right { flex: 0 0 auto; gap: 0.15rem; }
  .edit-btn { margin-right: 0.3rem; }
}
```

- [ ] **Step 4: 확인**

Run: `npx tsc --noEmit && npx vitest run && npm run build:editor`
Expected: 모두 통과.

- [ ] **Step 5: 커밋**

```bash
git add src/client/view/toolbar.ts
git commit -m "$(cat <<'EOF'
통합 보기 막대: 확대·맞춤·쪽 이동(이전·다음)·회전·전체 화면·Ctrl+휠·끌기, 좁은 화면의 도구▾ 메뉴

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/client/view/toolbar.ts src/client/view.ts public/css/view.css
```

---

### Task 6: 썸네일 칸

**Files:**
- Create: `src/client/view/thumbs.ts`
- Modify: `src/client/view.ts`, `public/css/view.css`(끝에 더하기)

**Interfaces:**
- Consumes: `PagesBody.thumbnail`·`pageCount`·`pageSize`·`goto`·`currentPage`·`onPage`·`onSourceChange`, `thumbsOpenAtStart`
- Produces: `setupThumbs(ctx: ShellCtx): void`; DOM — `#thumbs-panel > button.thumb-item[data-page]`(지금 쪽 `.current`) > `span.thumb-box`(그림 `canvas` 또는 `img`) + `span.thumb-label`; 열리면 `body.thumbs-open`

- [ ] **Step 1: `src/client/view/thumbs.ts`**

```ts
/**
 * 썸네일 칸(설계 §3 "새로") — 쪽마다 작은 그림, 누르면 그 쪽으로. 칸에 보이는 것만 그린다(PNG: 서버 ?thumb=true,
 * pdf.js: 그 자리에서 폭 120px). 640px 미만은 닫힌 채 시작하고, 닫혀 있는 동안은 아무것도 받지 않는다.
 */
import type { ShellCtx } from './ctx.js';
import { thumbsOpenAtStart } from './logic.js';

const THUMB_WIDTH = 120;

export function setupThumbs(ctx: ShellCtx): void {
  const btn = document.getElementById('thumbs-btn') as HTMLButtonElement;
  const panel = document.getElementById('thumbs-panel') as HTMLElement;
  let items: HTMLButtonElement[] = [];
  let marked: HTMLButtonElement | null = null;
  let observer: IntersectionObserver | null = null;

  const mark = (page: number) => {
    marked?.classList.remove('current');
    marked = items[page - 1] ?? null;
    marked?.classList.add('current');
    if (marked && !panel.hidden) marked.scrollIntoView({ block: 'nearest' });
  };

  const setOpen = (open: boolean) => {
    panel.hidden = !open;
    document.body.classList.toggle('thumbs-open', open);
    btn.setAttribute('aria-pressed', String(open));
    if (open) mark(ctx.body.currentPage);
  };

  const build = () => {
    observer?.disconnect();
    // 칸 안에서 보이는(또는 곧 보일) 것만 그린다 — 숨겨진 칸에서는 교차가 없어 요청도 없다
    observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (!entry.isIntersecting) continue;
          const item = entry.target as HTMLButtonElement;
          observer?.unobserve(item);
          const box = item.querySelector('.thumb-box') as HTMLElement;
          ctx.body
            .thumbnail(Number(item.dataset.page), THUMB_WIDTH)
            .then((picture) => box.replaceChildren(picture))
            .catch(() => box.classList.add('thumb-error'));
        }
      },
      { root: panel, rootMargin: '200px 0px' },
    );
    items = [];
    marked = null;
    for (let page = 1; page <= ctx.body.pageCount(); page += 1) {
      const item = document.createElement('button');
      item.type = 'button';
      item.className = 'thumb-item';
      item.dataset.page = String(page);
      item.setAttribute('aria-label', `${page}쪽으로 이동`);
      const size = ctx.body.pageSize(page);
      const box = document.createElement('span');
      box.className = 'thumb-box';
      box.style.aspectRatio = `${size.w} / ${size.h}`;
      const label = document.createElement('span');
      label.className = 'thumb-label';
      label.textContent = String(page);
      item.append(box, label);
      item.addEventListener('click', () => {
        ctx.body.goto(page);
        // 휴대폰에서는 고른 뒤 칸을 닫아 쪽을 넓게 보인다
        if (window.innerWidth < 640) setOpen(false);
      });
      items.push(item);
      observer.observe(item);
    }
    panel.replaceChildren(...items);
    mark(ctx.body.currentPage);
  };

  btn.addEventListener('click', () => setOpen(panel.hidden));
  ctx.body.onPage(mark);
  ctx.body.onSourceChange(build);
  build();
  setOpen(thumbsOpenAtStart(window.innerWidth));
}
```

- [ ] **Step 2: `src/client/view.ts`** — import `import { setupThumbs } from './view/thumbs.js';`, `setupToolbar(ctx);` 다음 줄에 `setupThumbs(ctx);`

- [ ] **Step 3: `public/css/view.css` — 썸네일 칸**

`@media (max-width: 639px)` 블록 **앞**에 넣고, 그 미디어 블록 안 끝에 두 줄을 더한다:

```css
/* 썸네일 칸(설계 §3 "새로") — 막대 아래 왼쪽, 열리면 쪽 칸을 그만큼 민다 */
.thumbs-panel {
  position: fixed;
  top: 6.2rem;
  left: 0;
  bottom: 0;
  z-index: 9;
  width: 152px;
  overflow-y: auto;
  padding: 0.6rem 0.8rem 2rem;
  display: flex;
  flex-direction: column;
  gap: 0.7rem;
  background: var(--v-glass);
  border-right: 1px solid var(--v-glass-line);
}
body.thumbs-open .pages-container { padding-left: calc(1rem + 152px); }
.thumb-item {
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 0.25rem;
  padding: 4px;
  border: 2px solid transparent;
  border-radius: 6px;
  background: none;
  cursor: pointer;
  font: inherit;
  color: inherit;
}
.thumb-item.current { border-color: var(--v-accent); }
.thumb-box { display: block; width: 120px; background: #fff; box-shadow: 0 1px 2px rgba(0, 0, 0, 0.3); overflow: hidden; }
.thumb-box > canvas, .thumb-box > img { display: block; width: 100%; height: 100%; }
.thumb-box.thumb-error { background: repeating-linear-gradient(45deg, #f3f3f3 0 6px, #e7e7e7 6px 12px); }
.thumb-label { font-size: 0.75rem; opacity: 0.75; }
```

미디어 블록 안에 더할 두 줄(휴대폰에서는 칸이 쪽 위에 겹친다):

```css
  body.thumbs-open .pages-container { padding-left: 1rem; }
  .thumbs-panel { box-shadow: 4px 0 18px rgba(0, 0, 0, 0.18); }
```

- [ ] **Step 4: 확인** — `npx tsc --noEmit && npx vitest run && npm run build:editor` 통과.

- [ ] **Step 5: 커밋**

```bash
git add src/client/view/thumbs.ts
git commit -m "$(cat <<'EOF'
통합 보기 썸네일 칸: 보이는 것만 그림(PNG는 서버 작은 그림, pdf.js는 그 자리), 지금 쪽 표시, 휴대폰 폭은 닫힌 채

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/client/view/thumbs.ts src/client/view.ts public/css/view.css
```

---

### Task 7: 찾기 — PNG는 서버, pdf.js는 브라우저

**Files:**
- Create: `src/client/view/search.ts`
- Modify: `src/client/view.ts`

**Interfaces:**
- Consumes: `PagesBody.search`·`showHits`·`clearHits`·`pageFrame`·`setZoomIndex`·`zoomIndex`·`zoomPercent`·`container`·`onSourceChange`, `ctx.onEscape`, `ZOOM_STEPS`·`hitCountLabel`·`wrapIndex`
- Produces: `setupSearch(ctx: ShellCtx): void`(Ctrl+F·F3·Shift+F3·Enter·Shift+Enter, Esc 우선순위 90)

- [ ] **Step 1: `src/client/view/search.ts`**

```ts
/**
 * 문서 안 찾기 — PNG 모드는 서버(/api/view/search), pdf.js 모드는 브라우저(findText.ts)가 찾고, 강조·이동은 같다
 * (viewer.js runSearch·gotoHit·gotoPageHit·centerOnElement 그대로). Ctrl+F·F3·Shift+F3·Enter·Shift+Enter·Esc.
 */
import type { ShellCtx } from './ctx.js';
import { ZOOM_STEPS, hitCountLabel, wrapIndex } from './logic.js';
import type { SearchHit } from './types.js';

/** 강조가 이보다 작으면(도면처럼 아주 작은 글자) 읽을 수 있는 배율까지 확대한다 */
const MIN_READABLE_HIT_PX = 11;
const TARGET_HIT_PX = 16;

const $ = <T extends HTMLElement = HTMLElement>(id: string): T => document.getElementById(id) as T;

/** 요소를 화면 가운데로 — 쪽 칸은 가로, 창은 세로로 스크롤한다(viewer.js centerOnElement) */
function centerOn(container: HTMLElement, el: HTMLElement): void {
  const rect = el.getBoundingClientRect();
  const box = container.getBoundingClientRect();
  container.scrollLeft += rect.left + rect.width / 2 - (box.left + container.clientWidth / 2);
  const dy = rect.top + rect.height / 2 - window.innerHeight / 2;
  const before = container.scrollTop;
  container.scrollTop += dy;
  const consumed = container.scrollTop - before;
  if (Math.abs(dy - consumed) > 1) window.scrollBy(0, dy - consumed);
}

export function setupSearch(ctx: ShellCtx): void {
  const bar = $('search-bar');
  const input = $<HTMLInputElement>('search-input');
  const count = $('search-count');
  const prev = $<HTMLButtonElement>('search-prev');
  const next = $<HTMLButtonElement>('search-next');
  let hits: SearchHit[] = [];
  let current = -1;
  let lastQuery = '';
  let running: AbortController | null = null;

  const setCount = (text: string) => {
    count.textContent = text;
    prev.disabled = hits.length === 0;
    next.disabled = hits.length === 0;
  };
  const open = () => {
    bar.hidden = false;
    input.focus();
    input.select();
  };
  const close = () => {
    bar.hidden = true;
    running?.abort();
    ctx.body.clearHits();
    hits = [];
    current = -1;
    lastQuery = '';
    setCount('');
  };
  const show = (index: number) => {
    if (hits.length === 0) return;
    current = wrapIndex(index, hits.length);
    setCount(hitCountLabel(current, hits.length));
    const el = ctx.body.showHits(hits, current);
    if (el) {
      const height = el.getBoundingClientRect().height;
      if (height > 0 && height < MIN_READABLE_HIT_PX) {
        const target = Math.min(ZOOM_STEPS[ZOOM_STEPS.length - 1], ctx.body.zoomPercent * (TARGET_HIT_PX / height));
        const found = ZOOM_STEPS.findIndex((step) => step >= target);
        const resolved = found === -1 ? ZOOM_STEPS.length - 1 : found;
        if (resolved > ctx.body.zoomIndex) ctx.body.setZoomIndex(resolved);
      }
    }
    const target = el ?? ctx.body.pageFrame(hits[current].page);
    if (target) centerOn(ctx.body.container, target);
  };
  const run = async () => {
    const query = input.value.trim();
    if (query === '') return;
    if (query === lastQuery) {
      if (hits.length > 0) show(current + 1);
      return;
    }
    lastQuery = query;
    ctx.body.clearHits();
    hits = [];
    setCount('검색 중…');
    running?.abort();
    const ac = new AbortController();
    running = ac;
    try {
      const found = await ctx.body.search(query, ac.signal);
      if (ac.signal.aborted) return;
      hits = found;
      current = -1;
      if (hits.length === 0) setCount('결과 없음');
      else show(0);
    } catch (error) {
      if ((error as { name?: string } | null)?.name === 'AbortError') return;
      lastQuery = '';
      setCount(error instanceof Error && error.message ? error.message : '검색 실패');
    }
  };

  $('search-btn').addEventListener('click', open);
  $('search-close').addEventListener('click', close);
  prev.addEventListener('click', () => show(current - 1));
  next.addEventListener('click', () => show(current + 1));
  input.addEventListener('keydown', (event) => {
    if (event.key !== 'Enter') return;
    event.preventDefault();
    if (event.shiftKey) show(current - 1);
    else void run();
  });
  document.addEventListener('keydown', (event) => {
    const key = event.key.toLowerCase();
    // 브라우저 기본 Ctrl+F는 그리지 않은 쪽을 찾지 못한다 — 자체 찾기로
    if ((event.ctrlKey || event.metaKey) && key === 'f') {
      event.preventDefault();
      open();
    } else if (key === 'f3') {
      event.preventDefault();
      if (hits.length > 0) show(current + (event.shiftKey ? -1 : 1));
      else open();
    }
  });
  ctx.onEscape(90, () => {
    if (bar.hidden) return false;
    close();
    return true;
  });
  // 원천이 바뀌면(pdf.js → PNG) 쪽 틀이 새로 만들어져 강조가 사라진다 — 찾기도 처음부터
  ctx.body.onSourceChange(close);
}
```

- [ ] **Step 2: `src/client/view.ts`** — import `import { setupSearch } from './view/search.js';`, `setupThumbs(ctx);` 다음에 `setupSearch(ctx);`

- [ ] **Step 3: 확인** — `npx tsc --noEmit && npx vitest run && npm run build:editor` 통과. (계획 작성 때 같은 코드로 PNG·pdf.js 두 모드에서 3쪽 PDF "Target" → `1 / 3`, 강조 3개를 헤드리스로 확인했다 — 정식 확인은 Task 13.)

- [ ] **Step 4: 커밋**

```bash
git add src/client/view/search.ts
git commit -m "$(cat <<'EOF'
통합 보기 찾기: 서버 PNG 모드는 서버 찾기, pdf.js 모드는 브라우저 찾기 — 강조·이동·작은 글자 자동 확대는 같게

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/client/view/search.ts src/client/view.ts
```

---

### Task 8: 인쇄·내려받기·다른 이름으로, 측정·반전(PDF)

**Files:**
- Create: `src/client/view/fileActions.ts`, `src/client/view/measure.ts`
- Modify: `src/client/view.ts`

**Interfaces:**
- Consumes: `disabledReason`·`paperMm`·`measureLabel`, `PagesBody.container`·`pageSize`·`onSourceChange`, `ctx.onEscape`, 서버 `/api/view/print`·`/api/view/download`(그대로)
- Produces: `setupFileActions(ctx)`, `setupMeasure(ctx)`. 다른 모듈과의 약속: `#markup-include-checkbox`가 체크돼 있으면 인쇄·내려받기에 `&includeMarkup=1`(체크 상자 표시는 Task 9). 측정 중 `#pages.measuring`, 반전 `#pages.pages-invert`(localStorage `docview-invert`, 지금 화면과 같은 열쇠).

- [ ] **Step 1: `src/client/view/fileActions.ts`**

```ts
/**
 * 인쇄·내려받기·다른 이름으로 저장(표시 포함/제외) — 지금 보기 화면(viewer.js setupToolbar)과 같은 서버 경로·진행률 표시.
 * 권한이 없는 단추는 흐리게 하고 이유를 툴팁으로(설계 §2).
 */
import type { ShellCtx } from './ctx.js';
import { disabledReason } from './logic.js';

const $ = <T extends HTMLElement = HTMLElement>(id: string): T => document.getElementById(id) as T;

/** content-length 기준 진행률과 함께 본문을 받는다 */
async function fetchWithProgress(url: string, onProgress: (received: number, total: number) => void): Promise<Uint8Array[]> {
  const res = await fetch(url);
  if (!res.ok || !res.body) {
    const json = (await res.json().catch(() => null)) as { error?: string } | null;
    throw new Error(json?.error || `요청 실패 (HTTP ${res.status})`);
  }
  const total = Number(res.headers.get('content-length')) || 0;
  const reader = res.body.getReader();
  const chunks: Uint8Array[] = [];
  let received = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    received += value.length;
    onProgress(received, total);
  }
  return chunks;
}

function formatProgress(received: number, total: number): string {
  return total > 0 ? `${Math.round((received / total) * 100)}%` : `${(received / 1048576).toFixed(1)}MB`;
}

/** 단추를 진행률 표시로 바꿨다가 끝나면 되돌린다(아이콘 단추는 툴팁에) */
async function runButtonTask(button: HTMLButtonElement, prefix: string, task: (onProgress: (r: number, t: number) => void) => Promise<void>): Promise<void> {
  const original = button.title;
  button.disabled = true;
  try {
    await task((received, total) => {
      button.title = `${prefix} ${formatProgress(received, total)}`;
    });
    button.title = original;
  } catch (error) {
    button.title = '실패';
    setTimeout(() => (button.title = original), 2500);
    throw error;
  } finally {
    button.disabled = false;
  }
}

function saveBlob(chunks: Uint8Array[], name: string): void {
  const url = URL.createObjectURL(new Blob(chunks as BlobPart[]));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

export function setupFileActions(ctx: ShellCtx): void {
  const { permissions, ext, name } = ctx.info;
  const includeMarkup = () => (($('markup-include-checkbox') as HTMLInputElement | null)?.checked ? '&includeMarkup=1' : '');
  const print = $<HTMLButtonElement>('print-btn');
  const download = $<HTMLButtonElement>('download-btn');
  const downloadAs = $<HTMLButtonElement>('download-as-btn');

  const printReason = disabledReason('print', permissions, ext);
  print.disabled = printReason !== null;
  print.title = printReason ?? '워터마크가 포함된 출력용 PDF';
  if (!printReason) {
    print.addEventListener('click', () => {
      // 팝업 차단을 피하려면 누른 그 순간에 창을 먼저 연다
      const printWindow = window.open('', '_blank');
      runButtonTask(print, '생성', async (onProgress) => {
        const chunks = await fetchWithProgress(`/api/view/print?${ctx.tokenQuery}${includeMarkup()}`, onProgress);
        const url = URL.createObjectURL(new Blob(chunks as BlobPart[], { type: 'application/pdf' }));
        if (printWindow && !printWindow.closed) printWindow.location.href = url;
        else window.open(url, '_blank', 'noopener');
        setTimeout(() => URL.revokeObjectURL(url), 60_000);
      }).catch(() => {
        if (printWindow && !printWindow.closed) printWindow.close();
      });
    });
    document.addEventListener('keydown', (event) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'p') print.click();
    });
  }

  const downloadReason = disabledReason('download', permissions, ext);
  for (const btn of [download, downloadAs]) {
    btn.disabled = downloadReason !== null;
    btn.title = downloadReason ?? (btn === download ? '원본 파일 다운로드' : '다른 이름으로 저장');
  }
  if (downloadReason) return;
  const run = (btn: HTMLButtonElement, fileName: string) =>
    runButtonTask(btn, '저장', async (onProgress) => {
      saveBlob(await fetchWithProgress(`/api/view/download?${ctx.tokenQuery}${includeMarkup()}`, onProgress), fileName);
    }).catch(() => undefined);
  download.addEventListener('click', () => void run(download, name));
  downloadAs.addEventListener('click', () => {
    const custom = window.prompt('저장할 파일 이름을 입력하세요', name);
    if (custom) void run(downloadAs, custom);
  });
}
```

- [ ] **Step 2: `src/client/view/measure.ts`**

```ts
/**
 * 거리 측정·도면 모드(반전) — PDF 문서에서만(지금 보기 화면과 같다, viewer.js measure·applyInvert). 다른 형식은 흐리게 + 이유.
 */
import type { ShellCtx } from './ctx.js';
import { disabledReason, measureLabel, paperMm } from './logic.js';

const $ = <T extends HTMLElement = HTMLElement>(id: string): T => document.getElementById(id) as T;
const SVG_NS = 'http://www.w3.org/2000/svg';
const INVERT_STORAGE_KEY = 'docview-invert';

interface Point {
  readonly frame: HTMLElement;
  readonly fx: number;
  readonly fy: number;
}

export function setupMeasure(ctx: ShellCtx): void {
  const measureBtn = $<HTMLButtonElement>('measure-btn');
  const invertBtn = $<HTMLButtonElement>('invert-btn');
  const reason = disabledReason('measure', ctx.info.permissions, ctx.info.ext);
  for (const btn of [measureBtn, invertBtn]) {
    btn.disabled = reason !== null;
    if (reason) btn.title = reason;
  }
  if (reason) return;
  const pages = ctx.body.container;
  const hint = $('measure-hint');
  const state = {
    mode: false,
    calibrating: false,
    pending: null as (Point & { dot: HTMLElement }) | null,
    items: [] as Array<{ paperMm: number; label: HTMLElement; els: Element[] }>,
    factor: 1,
    unit: 'mm(용지)',
  };

  const layerOf = (frame: HTMLElement): SVGSVGElement => {
    let layer = frame.querySelector<SVGSVGElement>('.measure-layer');
    if (!layer) {
      layer = document.createElementNS(SVG_NS, 'svg');
      layer.setAttribute('class', 'measure-layer');
      layer.setAttribute('viewBox', '0 0 100 100');
      layer.setAttribute('preserveAspectRatio', 'none');
      frame.append(layer);
    }
    return layer;
  };
  const dot = (frame: HTMLElement, fx: number, fy: number): HTMLElement => {
    const el = document.createElement('div');
    el.className = 'measure-dot';
    el.style.left = `${fx * 100}%`;
    el.style.top = `${fy * 100}%`;
    frame.append(el);
    return el;
  };
  const clear = () => {
    for (const item of state.items) for (const el of item.els) el.remove();
    state.items = [];
    state.pending?.dot.remove();
    state.pending = null;
  };
  const add = (a: Point & { dot: HTMLElement }, b: Point) => {
    const mm = paperMm(a, b, ctx.body.pageSize(Number(a.frame.dataset.page)));
    const line = document.createElementNS(SVG_NS, 'line');
    line.setAttribute('x1', String(a.fx * 100));
    line.setAttribute('y1', String(a.fy * 100));
    line.setAttribute('x2', String(b.fx * 100));
    line.setAttribute('y2', String(b.fy * 100));
    layerOf(a.frame).append(line);
    const dotB = dot(a.frame, b.fx, b.fy);
    const label = document.createElement('div');
    label.className = 'measure-label';
    label.style.left = `${((a.fx + b.fx) / 2) * 100}%`;
    label.style.top = `${((a.fy + b.fy) / 2) * 100}%`;
    label.textContent = measureLabel(mm, state.factor, state.unit);
    a.frame.append(label);
    state.items.push({ paperMm: mm, label, els: [line, a.dot, dotB, label] });
    if (!state.calibrating) return;
    state.calibrating = false;
    const value = Number((window.prompt('방금 잰 구간의 실제 길이를 입력하세요 (예: 1500 → mm)', '') || '').replace(/[^\d.]/g, ''));
    if (value > 0) {
      state.factor = value / mm;
      state.unit = 'mm';
      for (const item of state.items) item.label.textContent = measureLabel(item.paperMm, state.factor, state.unit);
      hint.textContent = '보정 완료 — 실제 치수(mm)로 표시됩니다';
    } else {
      hint.textContent = '보정이 취소되었습니다';
    }
  };
  const setMode = (on: boolean) => {
    state.mode = on;
    state.calibrating = false;
    measureBtn.setAttribute('aria-pressed', String(on));
    $('measure-bar').hidden = !on;
    pages.classList.toggle('measuring', on);
    if (!on) clear();
    else hint.textContent = '두 점을 클릭해 거리를 측정하세요';
  };

  measureBtn.addEventListener('click', () => setMode(!state.mode));
  $('measure-close').addEventListener('click', () => setMode(false));
  $('measure-clear').addEventListener('click', clear);
  $('measure-calibrate').addEventListener('click', () => {
    state.calibrating = true;
    hint.textContent = '실제 길이를 아는 구간(치수선 등)의 양 끝을 클릭하세요';
  });
  pages.addEventListener('click', (event) => {
    if (!state.mode) return;
    const frame = (event.target as Element).closest<HTMLElement>('.page-frame');
    if (!frame) return;
    const rect = frame.getBoundingClientRect();
    const point = { frame, fx: (event.clientX - rect.left) / rect.width, fy: (event.clientY - rect.top) / rect.height };
    if (!state.pending || state.pending.frame !== frame) {
      state.pending?.dot.remove();
      state.pending = { ...point, dot: dot(frame, point.fx, point.fy) };
    } else {
      add(state.pending, point);
      state.pending = null;
    }
  });
  ctx.onEscape(10, () => {
    if (!state.mode) return false;
    setMode(false);
    return true;
  });
  // 쪽 틀이 새로 만들어지면(pdf.js → PNG) 잰 것은 사라진다 — 상태도 비운다
  ctx.body.onSourceChange(() => setMode(false));

  const applyInvert = (on: boolean) => {
    pages.classList.toggle('pages-invert', on);
    invertBtn.setAttribute('aria-pressed', String(on));
    try {
      localStorage.setItem(INVERT_STORAGE_KEY, on ? '1' : '0');
    } catch {
      // 시크릿 모드 등에서 저장 실패해도 무시
    }
  };
  invertBtn.addEventListener('click', () => applyInvert(!pages.classList.contains('pages-invert')));
  try {
    if (localStorage.getItem(INVERT_STORAGE_KEY) === '1') applyInvert(true);
  } catch {
    // 저장소를 못 쓰면 기본값
  }
}
```

- [ ] **Step 3: `src/client/view.ts`** — import 두 줄(`./view/fileActions.js`의 `setupFileActions`, `./view/measure.js`의 `setupMeasure`), `setupSearch(ctx);` 다음에 `setupFileActions(ctx);`·`setupMeasure(ctx);`

- [ ] **Step 4: 확인** — `npx tsc --noEmit && npx vitest run && npm run build:editor` 통과.

- [ ] **Step 5: 커밋**

```bash
git add src/client/view/fileActions.ts src/client/view/measure.ts
git commit -m "$(cat <<'EOF'
통합 보기: 인쇄·내려받기·다른 이름으로(표시 포함), PDF 측정·반전 이전 — 권한·형식 밖 단추는 흐리게 + 이유

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/client/view/fileActions.ts src/client/view/measure.ts src/client/view.ts
```

---

### Task 9: 강조 표시·글 넣기(markup) 이전

**Files:**
- Create: `src/client/view/markup.ts`
- Modify: `src/client/view.ts`

**Interfaces:**
- Consumes: `ShellCtx`(`tokenQuery`·`info.permissions.canEdit`·`body.container`·`body.pageSize`·`body.onSourceChange`·`onEscape`), `disabledReason('markup', …)`, 서버 `/api/view/markup`(GET·POST·PATCH·DELETE), `/api/view/pages/:n/text-style`(그대로)
- Produces: `setupMarkup(ctx: ShellCtx): void`. DOM은 지금 보기 화면과 **같은 클래스**(`.markup-layer`·`.markup-rect`·`.markup-popover`·`.text-overlay-item`·`.text-insert-wrapper`·`.text-insert-input`·`.text-tool-ghost`·`.text-tool-hint`, `#pages.markup-drawing`·`#pages.text-tool-active`) — `verify:viewer-quicktext`가 이 이름으로 확인한다.

이 작업은 **옮기기**다. `public/js/viewer.js` 2013~2592줄(`/* ==================== 마크업(리뷰 코멘트) ==================== */`부터 `setupTextTool` 끝까지, 함수 `markupLayer`·`clearMarkupOverlays`·`deleteMarkupItem`·`pageHeightPtFor`·글꼴 크기 상수·`clampFontSizePt`·`applyFontSizeToEl`·`persistFontSize`·`scheduleFontSizePersist`·`startTextDrag`·`cssFontFamilyFor`·`renderTextOverlayItem`·`renderMarkupItem`·`renderAllMarkup`·`updateMarkupIncludeVisibility`·`loadMarkup`·`closeMarkupPopover`·`cancelMarkupDraft`·`setMarkupDrawMode`·`openMarkupPopover`·`setupMarkupTool`·`textTool`·`setTextToolActive`·`showTextGhost`·`cancelTextInsert`·`adjustPendingFontSize`·`commitTextInsert`·`openTextInsertBox`·`setupTextTool`)를 동작을 바꾸지 않고 TypeScript로 옮긴다. 상태 객체(`markup`·`textTool`·`textDragState`)는 `setupMarkup` 안의 지역 변수로 두고, 함수들은 그 안의 지역 함수(또는 같은 파일의 비공개 함수 + 상태 인자)로 둔다.

**바꿔 쓰는 것(이 표 밖의 것은 글자 그대로):**

| viewer.js | markup.ts |
|---|---|
| `pagesContainer` | `ctx.body.container` |
| `[...pagesContainer.children].indexOf(frame) + 1` | `Number(frame.dataset.page)` |
| `pageHeightPtFor(pageIdx)` | `ctx.body.pageSize(pageIdx + 1).h`(함수는 지우고 부르는 곳을 바꾼다) |
| `fetchJsonView(url, options)` | 같은 파일의 `fetchJson`(아래) |
| `tokenQuery` | `ctx.tokenQuery` |
| `permissions.canEdit` | `ctx.info.permissions.canEdit` |
| `setupMarkupTool()`의 `addBtn.hidden = false` + `if (permissions.canEdit)` | 아래 "단추 상태" 코드 |
| `setupTextTool()`의 `if (!permissions.canEdit) return; btn.hidden = false;` | 같은 "단추 상태" 코드(권한이 없으면 흐림 + 이유, 듣기를 걸지 않는다) |
| `setupDeterrents()` 안의 Esc 갈래(`textTool.pending` → `commitTextInsert`, `textTool.active` → `setTextToolActive(false)`, `markup.drawMode` → `setMarkupDrawMode(false)`) | `ctx.onEscape(20, …)`·`ctx.onEscape(30, …)`·`ctx.onEscape(40, …)` — 각각 해당 상태일 때만 처리하고 `true` |
| `alert(…)`·`confirm(…)` | 그대로 `window.alert`·`window.confirm`(지금 화면과 같은 동작) |
| `init()`의 `loadMarkup()` 부르기 | `setupMarkup` 끝에서 `void loadMarkup();` |

**새로 더하는 것:**

```ts
/** viewer.js fetchJsonView와 같은 응답 모양({success, data, error}) — 실패면 그 문구로 throw */
async function fetchJson<T = unknown>(url: string, options: RequestInit = {}): Promise<T> {
  const res = await fetch(url, options);
  const json = (await res.json().catch(() => null)) as { success?: boolean; data?: T; error?: string } | null;
  if (!res.ok || !json?.success) throw new Error(json?.error || '요청에 실패했습니다.');
  return json.data as T;
}
```

`setupMarkup(ctx)` 시작부의 단추 상태:

```ts
  const reason = disabledReason('markup', ctx.info.permissions, ctx.info.ext);
  const addBtn = document.getElementById('markup-add-btn') as HTMLButtonElement;
  const textBtn = document.getElementById('text-tool-btn') as HTMLButtonElement;
  for (const btn of [addBtn, textBtn]) {
    btn.disabled = reason !== null;
    if (reason) btn.title = reason;
  }
```

이어서 `if (!reason) { …setupMarkupTool 본문(addBtn 듣기 포함)… ; …setupTextTool 본문… }`, 표시 목록 읽기(`loadMarkup`)는 권한과 무관하게 한다(지금 화면도 열람만 토큰에 표시를 보여 준다). 쪽 틀이 다시 만들어지면(pdf.js → PNG) 겹침을 다시 그린다:

```ts
  ctx.body.onSourceChange(() => {
    closeMarkupPopover();
    if (textTool.pending) cancelTextInsert();
    renderAllMarkup();
  });
```

- [ ] **Step 1: 위 규칙대로 `src/client/view/markup.ts`를 쓴다.** 표시 항목 타입은 서버 것을 그대로 쓴다 — `import type { MarkupItem, CommentMarkupItem, TextMarkupItem } from '../../markup.js';`(`kind: 'comment'`는 사각형+코멘트 `w`·`h`, `kind: 'text'`는 `fontSizePt`·`bold`·`fontFamily`; 공통 `id`·`page`·`x`·`y`·`text`·`author`·`createdAt`). 타입만 가져오므로 서버 코드가 묶음에 들어가지 않는다. 글 스타일 응답(`/api/view/pages/:n/text-style`)은 viewer.js가 쓰는 필드 그대로 `{ fontSizePt: number; bold: boolean; fontFamily: string; baselineY: number; matched: boolean }`로 적는다. `any`는 쓰지 않는다.

- [ ] **Step 2: 옮긴 뒤 빠진 것이 없는지 확인**

Run: `grep -oE "\b(markupLayer|clearMarkupOverlays|deleteMarkupItem|clampFontSizePt|applyFontSizeToEl|persistFontSize|scheduleFontSizePersist|startTextDrag|cssFontFamilyFor|renderTextOverlayItem|renderMarkupItem|renderAllMarkup|updateMarkupIncludeVisibility|loadMarkup|closeMarkupPopover|cancelMarkupDraft|setMarkupDrawMode|openMarkupPopover|setTextToolActive|showTextGhost|cancelTextInsert|adjustPendingFontSize|commitTextInsert|openTextInsertBox)\b" src/client/view/markup.ts | sort -u | wc -l`
Expected: `24`(지운 `pageHeightPtFor`·`setupMarkupTool`·`setupTextTool` 셋 빼고 전부 있다).

- [ ] **Step 3: `src/client/view.ts`** — import `import { setupMarkup } from './view/markup.js';`, `setupMeasure(ctx);` 다음에 `setupMarkup(ctx);`

- [ ] **Step 4: 확인** — `npx tsc --noEmit && npx vitest run && npm run build:editor` 통과. (`verify:viewer-quicktext`는 Task 13에서 — 이 작업의 정식 확인이다.)

- [ ] **Step 5: 커밋**

```bash
git add src/client/view/markup.ts
git commit -m "$(cat <<'EOF'
통합 보기: 강조 표시·글 넣기(markup) 이전 — 서버 저장·화면 클래스는 지금 보기 화면과 같게, 권한 없으면 흐리게

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/client/view/markup.ts src/client/view.ts
```

---

### Task 10: 편집 단추와 편집기의 "보기로"

**Files:**
- Create: `src/client/view/editButton.ts`, `src/client/backToView.ts`
- Modify: `src/client/officeDialog.ts`, `src/client/officeShell.ts`, `src/client/editor.ts`(ViewInfo ~128줄, `initOfficeShell` ~660줄), `src/client/grid-editor.ts`(ViewInfo ~54줄, ~437줄), `src/client/editor-slides.ts`(ViewInfo ~44줄, ~1215줄), `src/client/editor-pdf.ts`(ViewInfo 26줄, `main` ~278줄), `src/client/editor-hwp.ts`, `src/client/view.ts`

**Interfaces:**
- Consumes: `editHref`·`carryViewParams`·`offersBackToView`(logic.ts), info `editorUrl`·`viewerUrl`(Task 1), `POST /api/view/save`(그대로 — 방 저장 + PDF 다시 변환까지 끝나고 응답)
- Produces: `setupEditButton(ctx: Pick<ShellCtx, 'info'>): void`; `mountBackToView(host: Element, before: Element | null, opts: BackToViewOptions): HTMLButtonElement | null`(`BackToViewOptions = { token; viewerUrl: string | undefined; canEdit; beforeLeave?: () => Promise<boolean> }`, 단추 `#back-to-view`); `officeChoose(message, { okLabel, altLabel, title? }): Promise<'ok' | 'alt' | 'cancel'>`; `OfficeShellOptions.viewerUrl?: string`

- [ ] **Step 1: `src/client/view/editButton.ts`**

```ts
/** 편집 단추(설계 V6) — 세션 응답에 편집기 주소가 있을 때만. 같은 창(iframe)에서 편집기로, 돌아올 수 있게 from=view */
import type { ShellCtx } from './ctx.js';
import { editHref } from './logic.js';

export function setupEditButton(ctx: Pick<ShellCtx, 'info'>): void {
  const editorUrl = ctx.info.editorUrl;
  const btn = document.getElementById('edit-btn') as HTMLButtonElement;
  if (!editorUrl) return;
  btn.hidden = false;
  btn.addEventListener('click', () => location.assign(editHref(editorUrl, location.search)));
}
```

`src/client/view.ts`: import `import { setupEditButton } from './view/editButton.js';`, `boot()`의 `startPresence();` 다음 줄에 `setupEditButton({ info });`(쪽을 그리기 전부터 쓸 수 있게 — 설계 V8).

- [ ] **Step 2: `src/client/backToView.ts`**

```ts
/**
 * 편집기의 "보기로" 단추(설계 V7) — 통합 보기 껍데기의 편집 단추로 들어온 경우(from=view)에만 넣는다. 향기랩처럼 편집기를
 * 바로 연 곳에는 생기지 않는다. 자동 저장 편집기(워드·엑셀·PPT·PDF)는 쓰기 권한이 있으면 서버에 지금 저장(/api/view/save —
 * 저장 뒤 보기용 PDF까지 다시 만든다)한 뒤 이동하고, 한글은 beforeLeave(저장 안 한 내용 확인)를 쓴다.
 */
import { officeAlert } from './officeDialog.js';
import { carryViewParams, offersBackToView } from './view/logic.js';

export interface BackToViewOptions {
  readonly token: string;
  /** /api/view/info 의 viewerUrl */
  readonly viewerUrl: string | undefined;
  readonly canEdit: boolean;
  /** 떠나도 되는지(한글) — 없으면 자동 저장 편집기 규칙 */
  readonly beforeLeave?: () => Promise<boolean>;
}

/** host 안 before 앞에 단추를 넣는다. from=view가 아니면 아무것도 하지 않고 null */
export function mountBackToView(host: Element, before: Element | null, opts: BackToViewOptions): HTMLButtonElement | null {
  const viewerUrl = opts.viewerUrl;
  if (!viewerUrl || !offersBackToView(location.search, viewerUrl)) return null;
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.id = 'back-to-view';
  btn.className = 'btn-pill back-to-view';
  btn.textContent = '보기로';
  btn.title = '보기 화면으로 돌아가기';

  const flush = async (): Promise<boolean> => {
    if (!opts.canEdit) return true;
    btn.textContent = '저장하는 중…';
    const res = await fetch(`/api/view/save?token=${encodeURIComponent(opts.token)}`, { method: 'POST' }).catch(() => null);
    btn.textContent = '보기로';
    if (res?.ok) return true;
    await officeAlert('편집한 내용을 저장하지 못했습니다. 잠시 뒤 다시 눌러 주세요.');
    return false;
  };

  btn.addEventListener('click', async () => {
    btn.disabled = true;
    try {
      const go = opts.beforeLeave ? await opts.beforeLeave() : await flush();
      if (go) location.assign(carryViewParams(viewerUrl, location.search));
    } finally {
      btn.disabled = false;
    }
  });
  host.insertBefore(btn, before);
  return btn;
}
```

- [ ] **Step 3: `src/client/officeDialog.ts` — 세 갈래 대화상자**

`DialogOptions`에 `readonly danger?: boolean;` 다음 줄로 더한다:

```ts
  /** 세 번째 단추(확인·취소 사이) — 누르면 결과가 'alt'(officeChoose) */
  readonly altLabel?: string;
```

`open`의 반환 타입을 `Promise<{ ok: boolean; alt: boolean; value: string }>`로, 그리고 `actions.append(okBtn);` 다음에:

```ts
    let altBtn: HTMLButtonElement | null = null;
    if (opts.altLabel) {
      altBtn = document.createElement('button');
      altBtn.type = 'button';
      altBtn.className = 'btn-mini';
      altBtn.textContent = opts.altLabel;
      actions.append(altBtn);
    }
```

`finish`를 `const finish = (ok: boolean, alt = false) => { …; resolve({ ok, alt, value: input?.value ?? '' }); };`로, Tab 순환 배열을 `[input, okBtn, altBtn, cancelBtn]`로, `cancelBtn?.addEventListener(...)` 앞에 `altBtn?.addEventListener('click', () => finish(false, true));`. 파일 끝에:

```ts
/** 세 갈래(확인 / 다른 동작 / 취소) — Esc·바깥 누르기는 'cancel' */
export async function officeChoose(message: string, opts: DialogOptions & { readonly okLabel: string; readonly altLabel: string }): Promise<'ok' | 'alt' | 'cancel'> {
  const r = await open(message, 'confirm', opts);
  return r.ok ? 'ok' : r.alt ? 'alt' : 'cancel';
}
```

(기존 `officeAlert`·`officeConfirm`·`officePrompt`는 `alt`를 읽지 않으므로 동작이 같다.)

- [ ] **Step 4: `src/client/officeShell.ts` — 워드·엑셀·PPT 제목 표시줄**

파일 머리 주석 다음에 `import { mountBackToView } from './backToView.js';`. `OfficeShellOptions`의 `readonly canPrint: boolean;` 다음에:

```ts
  /** /api/view/info 의 viewerUrl — 보기 화면에서 들어왔으면(from=view) 제목 표시줄에 "보기로"(backToView.ts) */
  readonly viewerUrl?: string;
```

`initOfficeShell`의 `const toolbar = document.getElementById('toolbar') as HTMLElement;` 다음에:

```ts
  const titleRight = toolbar.querySelector('.titlebar-right');
  if (titleRight) {
    mountBackToView(titleRight, titleRight.querySelector('#presence-list'), { token: options.token, viewerUrl: options.viewerUrl, canEdit: options.canEdit });
  }
```

`editor.ts`·`grid-editor.ts`·`editor-slides.ts` 각각: 파일 안 `interface ViewInfo`에 `readonly viewerUrl?: string;`을 더하고, `initOfficeShell({` 호출의 `canPrint: info.permissions.canPrint,` 다음 줄에 `viewerUrl: info.viewerUrl,`.

- [ ] **Step 5: `src/client/editor-pdf.ts`**

import 추가 `import { mountBackToView } from './backToView.js';`, `interface ViewInfo`에 `readonly viewerUrl?: string;`, `main()`의 `document.title = \`${info.name} — PDF 편집\`;` 다음 줄에:

```ts
  // 보기 화면에서 편집으로 들어왔으면 머리에 "보기로"(설계 V7) — 저장은 방이 하고, 떠나기 전에 지금 저장한다
  mountBackToView(document.querySelector('.pdf-titlebar') as HTMLElement, $('presence'), {
    token,
    viewerUrl: info.viewerUrl,
    canEdit: info.permissions.canEdit,
  });
```

- [ ] **Step 6: `src/client/editor-hwp.ts`**

import 줄 `import { officeAlert } from './officeDialog.js';` → 

```ts
import { mountBackToView } from './backToView.js';
import { officeAlert, officeChoose } from './officeDialog.js';
```

`interface ViewInfo`에 `readonly viewerUrl?: string;`. `saveToServer`가 성공 여부를 돌려주게 바꾼다 — 시그니처를 `async function saveToServer(editor: RhwpEditor, ext: string): Promise<boolean>`로(위에 `/** 저장 — 성공하면 true(실패는 안내까지 하고 false) */`), `try` 끝(`saveStatusEl.textContent = \`저장됨 …\`;` 다음)에 `return true;`, `catch` 끝(`void officeAlert(…);` 다음)에 `return false;`. 기존 저장 단추는 `void saveToServer(…)`라 그대로다. `main()`의 `if (canEdit) { saveBtn.addEventListener(…) }` 블록 다음에:

```ts
  // 보기 화면에서 편집으로 들어왔으면 "보기로"(설계 V7) — 한글은 자동 저장이 아니라 저장 안 한 내용을 묻는다
  mountBackToView(saveStatusEl.parentElement as HTMLElement, null, {
    token,
    viewerUrl: info.viewerUrl,
    canEdit,
    beforeLeave: async () => {
      // 상태를 못 읽으면 바뀐 것으로 친다(묻는 쪽이 안전)
      const dirty = canEdit && (await editor.getDocumentState().then((state) => state.dirty, () => true));
      if (!dirty) return true;
      const choice = await officeChoose('저장하지 않은 내용이 있습니다 — 저장하고 돌아갈까요?', {
        title: '보기로',
        okLabel: '저장 후 이동',
        altLabel: '버리고 이동',
      });
      if (choice === 'cancel') return false;
      return choice === 'alt' || (await saveToServer(editor, info.ext));
    },
  });
```

- [ ] **Step 6b: 단추 모양** — 새 CSS 없음. `public/css/view.css`는 보기 화면 전용이라 편집기에 닿지 않고, 편집기 머리 단추는 `style.css`의 기존 `.btn-pill`을 쓴다. 제목 표시줄 색 위에서 읽히는지는 Task 13 실제 브라우저 표("보기로" 줄)에서 본다.

- [ ] **Step 7: 확인**

Run: `npx tsc --noEmit && npx vitest run && npm run build:editor`
Expected: 모두 통과(`tests/officeDialog*`가 있으면 그것도 — 기존 함수 동작은 같다).

- [ ] **Step 8: 커밋**

```bash
git add src/client/view/editButton.ts src/client/backToView.ts
git commit -m "$(cat <<'EOF'
통합 보기 ↔ 편집: 편집 단추(from=view), 편집기 넷과 한글 편집기에 "보기로" — 자동 저장은 지금 저장 뒤 이동, 한글은 저장 안 한 내용 확인

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/client/view/editButton.ts src/client/backToView.ts src/client/officeDialog.ts src/client/officeShell.ts src/client/editor.ts src/client/grid-editor.ts src/client/editor-slides.ts src/client/editor-pdf.ts src/client/editor-hwp.ts src/client/view.ts
```

---

### Task 11: "쪽을 준비하는 중"(editReady)과 오류 화면의 편집 링크

**Files:**
- Create: `src/client/view/prepare.ts`
- Modify: `src/client/view.ts`(`showError`, `boot`), `public/css/view.css`

**Interfaces:**
- Consumes: `GET /api/view/status`(Task 1), `#prep-screen`·`#prep-progress`·`#error-edit`(Task 3), `editHref`
- Produces: `waitUntilPrepared(tokenQuery: string): Promise<void>`(끝나면 resolve, 실패면 그 문구로 reject), `showError(message, editorUrl?)`

- [ ] **Step 1: `src/client/view/prepare.ts`**

```ts
/**
 * "쪽을 준비하는 중"(설계 V8) — 첫 변환(editReady)이 끝나지 않은 문서. 쪽 자리에 안내와 진행 막대를 두고 /api/view/status를
 * 1.5초마다 본다(감사 기록 없음). 그동안 편집 단추는 이미 쓸 수 있다. 끝나면 resolve, 실패하면 그 문구로 reject.
 */
const POLL_MS = 1500;

interface StatusData {
  readonly status: 'processing' | 'ready' | 'failed';
  readonly progress: number;
  readonly error?: string;
}

export async function waitUntilPrepared(tokenQuery: string): Promise<void> {
  const screen = document.getElementById('prep-screen') as HTMLElement;
  const bar = document.getElementById('prep-progress') as HTMLProgressElement;
  screen.hidden = false;
  try {
    for (;;) {
      const res = await fetch(`/api/view/status?${tokenQuery}`, { cache: 'no-store' }).catch(() => null);
      const json = res ? ((await res.json().catch(() => null)) as { success?: boolean; data?: StatusData; error?: string } | null) : null;
      if (res && !res.ok) throw new Error(json?.error || '문서를 열 수 없습니다.');
      const data = json?.data;
      if (data?.status === 'ready') return;
      if (data?.status === 'failed') throw new Error(data.error || '문서를 변환하지 못했습니다.');
      if (data) bar.value = data.progress;
      await new Promise((resolve) => setTimeout(resolve, POLL_MS));
    }
  } finally {
    screen.hidden = true;
  }
}
```

(네트워크가 잠깐 끊기면(`res`가 null) 다음 차례에 다시 본다.)

- [ ] **Step 2: `src/client/view.ts`**

import 추가: `import { waitUntilPrepared } from './view/prepare.js';`, logic import에 `editHref` 더하기.

`showError`를 바꾼다:

```ts
function showError(message: string, editorUrl?: string): void {
  (document.getElementById('toolbar') as HTMLElement).hidden = true;
  (document.getElementById('pages') as HTMLElement).hidden = true;
  (document.getElementById('error-screen') as HTMLElement).hidden = false;
  (document.getElementById('error-message') as HTMLElement).textContent = message;
  // 쪽을 못 그려도 편집기는 원본으로 열린다(설계 §7)
  const edit = document.getElementById('error-edit') as HTMLAnchorElement;
  if (editorUrl) {
    edit.href = editHref(editorUrl, location.search);
    edit.hidden = false;
  }
  postReady();
}
```

`boot()`의 `setupEditButton({ info });` 다음에:

```ts
  if (info.processing) {
    postReady();
    try {
      await waitUntilPrepared(tokenQuery);
      info = await fetchInfo();
    } catch (error) {
      showError(error instanceof Error ? error.message : '문서를 열 수 없습니다.', info.editorUrl);
      return;
    }
  }
```

- [ ] **Step 3: `public/css/view.css` — 준비 화면**(`@media (max-width: 639px)` 블록 앞)

```css
/* "쪽을 준비하는 중"(설계 V8) */
.prep-screen {
  position: fixed;
  inset: 6.2rem 0 0;
  display: grid;
  place-content: center;
  justify-items: center;
  gap: 0.8rem;
  z-index: 8;
  font-size: 1rem;
}
.prep-screen progress { width: min(320px, 70vw); }
```

- [ ] **Step 4: 확인** — `npx tsc --noEmit && npx vitest run && npm run build:editor` 통과.

- [ ] **Step 5: 커밋**

```bash
git add src/client/view/prepare.ts
git commit -m "$(cat <<'EOF'
통합 보기: 첫 변환 중이면 "쪽을 준비하는 중"과 진행 막대(편집 단추는 먼저), 변환 실패 화면에 편집 화면으로 열기

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/client/view/prepare.ts src/client/view.ts public/css/view.css
```

---

### Task 12: `viewerUrl` 전환 — 쪽 문서는 view.html, viewer.html은 넘기기

설계 §3 표의 모든 줄이 Task 3~11에 있다(확대·쪽 이동·회전·전체 화면 5, 찾기 7, 인쇄·내려받기 8, 표시 9, 측정·반전 8, 워터마크·함께 보는 사람 3, 썸네일 6, 글자 선택 4, 편집·보기로 10). 그래서 이 작업에서 전환한다.

**Files:**
- Create: `scripts/measure-view-first-paint.mjs`
- Modify: `src/routes/documents.ts`(토큰 응답 ~520줄), `public/js/viewer.js`(`init()`의 info 확인 직후 ~2756줄), `public/api-docs.html`(~284·292줄), `README.md`(~419줄), `tests/viewShellRoutes.test.ts`

**Interfaces:**
- Consumes: `viewPathFor`(Task 1)
- Produces: `POST /api/documents/:id/tokens` 응답 `viewerUrl` = `viewPathFor(meta.ext, meta.viewKind) + '?token=…'`; `scripts/measure-view-first-paint.mjs <BASE> <경로> [횟수]` → 토큰 두 종류(열람만·내려받기)의 CPU×4 첫 쪽 시간 중앙값(ms)

- [ ] **Step 1: 전환 전 기준값 — 측정 스크립트를 쓰고 지금 화면을 잰다**

`scripts/measure-view-first-paint.mjs`:

```js
// 휴대폰(CPU×4) 첫 쪽 시간 — 쪽 틀 하나가 그려질 때까지(.page-frame.page-ready). 시험 서버에서만.
//   node scripts/measure-view-first-paint.mjs [http://127.0.0.1:8124] [/view.html | /viewer.html] [횟수=5]
// 시험 문서는 매번 만든다(20쪽 글자 PDF). 열람만(서버 PNG)·내려받기(새 화면은 pdf.js) 토큰 각각의 중앙값을 찍는다.
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const PATH = process.argv[3] ?? '/view.html';
const RUNS = Number(process.argv[4] ?? 5);

const doc = await PDFDocument.create();
const font = await doc.embedFont(StandardFonts.Helvetica);
for (let i = 1; i <= 20; i += 1) {
  const page = doc.addPage([595, 842]);
  for (let y = 780; y > 60; y -= 18) page.drawText(`First paint page ${i} line ${y}`, { x: 50, y, size: 11, font });
}
const form = new FormData();
form.append('file', new Blob([await doc.save()]), '첫화면.pdf');
const up = (await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json()).data;
for (let i = 0; i < 200; i += 1) {
  if ((await (await fetch(`${BASE}/api/documents/${up.id}`)).json()).data.status === 'ready') break;
  await new Promise((r) => setTimeout(r, 250));
}
const token = async (body) =>
  (await (await fetch(`${BASE}/api/documents/${up.id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: '측정', canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...body }) })).json()).data.token;

const browser = await chromium.launch();
const median = (xs) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];
for (const [label, body] of [['열람만', {}], ['내려받기', { canDownload: true, canCopy: true }]]) {
  const t = await token(body);
  const times = [];
  for (let run = 0; run < RUNS; run += 1) {
    const ctx = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 3, isMobile: true, hasTouch: true });
    const page = await ctx.newPage();
    const cdp = await ctx.newCDPSession(page);
    await cdp.send('Emulation.setCPUThrottlingRate', { rate: 4 });
    const t0 = Date.now();
    await page.goto(`${BASE}${PATH}?token=${encodeURIComponent(t)}`);
    await page.waitForSelector('.page-frame.page-ready', { timeout: 120000 });
    times.push(Date.now() - t0);
    await ctx.close();
  }
  console.log(`${PATH} ${label}: 중앙값 ${median(times)}ms (${times.join(', ')})`);
}
await browser.close();
await fetch(`${BASE}/api/documents/${up.id}`, { method: 'DELETE' });
```

시험 서버(Global Constraints대로 8124)를 띄우고: `node scripts/measure-view-first-paint.mjs http://127.0.0.1:8124 /viewer.html 5` → 두 줄을 scratchpad의 `first-paint-baseline.txt`에 적어 둔다(Task 13 보고에 쓴다). 서버는 PID로 끈다.

- [ ] **Step 2: 실패하는 시험 더하기** — `tests/viewShellRoutes.test.ts`의 마지막 `test` 다음에:

```ts
  test('토큰 응답 viewerUrl: 쪽 문서는 껍데기, 격자·한글은 지금 보기 화면', async () => {
    const issue = async (id: string) =>
      (await app.inject({ method: 'POST', url: `/api/documents/${id}/tokens`, payload: { viewer: '연동' } })).json().data;
    const docx = await addDoc('docx', { viewKind: 'pages' });
    expect((await issue(docx)).viewerUrl.startsWith('/view.html?token=')).toBe(true);
    const grid = await addDoc('xlsx', { viewKind: 'grid', derived: false });
    expect((await issue(grid)).viewerUrl.startsWith('/viewer.html?token=')).toBe(true);
    const hwp = await addDoc('hwp', { viewKind: 'pages' });
    expect((await issue(hwp)).viewerUrl.startsWith('/viewer.html?token=')).toBe(true);
  });
```

Run: `npx vitest run tests/viewShellRoutes.test.ts` → 새 시험 FAIL(`/viewer.html?token=`).

- [ ] **Step 3: `src/routes/documents.ts` 토큰 응답**

`viewerUrl: \`/viewer.html?token=${encodeURIComponent(token)}\`,` → 

```ts
          // 쪽 문서는 통합 보기 껍데기, 엑셀 격자·도면·한글은 지금 보기 화면(viewPathFor) — 주소·토큰 계약은 그대로(설계 V2)
          viewerUrl: `${viewPathFor(meta.ext, meta.viewKind)}?token=${encodeURIComponent(token)}`,
```

- [ ] **Step 4: `public/js/viewer.js` — 쪽 문서는 넘기기**

`init()`에서 `if (!res.ok || !json.success) { showError(...); return; }` 블록 바로 다음에:

```js
  // 쪽 문서는 통합 보기 껍데기(view.html)로 — 예전에 받아 둔 viewer.html 주소도 같은 화면으로 열린다(같은 조회 문자열).
  // 기준은 서버 viewerUrl(viewPathFor) 하나라 view.html과 서로 넘기지 않는다.
  if (json.data.viewerUrl && new URL(json.data.viewerUrl, location.href).pathname === '/view.html') {
    location.replace(`/view.html${location.search}`);
    return;
  }
```

- [ ] **Step 5: 문서 — `public/api-docs.html`·`README.md`**

api-docs.html 토큰 응답 예시의 `"viewerUrl": "/viewer.html?token=eyJhbGciOi…",` → `"viewerUrl": "/view.html?token=eyJhbGciOi…",`, 필드 표의 `viewerUrl` 설명 칸을 `항상 내려준다 — 워드·PPT·PDF 등 쪽 문서는 <code>/view.html</code>(통합 보기 화면), 엑셀·도면·한글은 <code>/viewer.html</code>. 이 서비스의 상대경로이므로, 호출하는 쪽에서 이 서비스의 접속 주소를 앞에 붙여야 한다. 예전에 받은 <code>/viewer.html</code> 주소도 그대로 열린다(쪽 문서면 새 화면으로 넘어간다).`로.

README.md 3번 항목의 `` `editorUrl`/`viewerUrl`은 이 `` 줄 앞에 한 문장: `` `viewerUrl`은 쪽 문서(워드·PPT·PDF 등)면 통합 보기 화면 `/view.html`, 엑셀·도면·한글이면 `/viewer.html`이다(예전 `/viewer.html` 주소도 쪽 문서면 새 화면으로 넘어간다). ``

- [ ] **Step 6: 확인**

Run: `npx vitest run && npx tsc --noEmit && npm run build:editor`
Expected: 새 시험 포함 전부 PASS.

- [ ] **Step 7: 커밋**

```bash
git add scripts/measure-view-first-paint.mjs
git commit -m "$(cat <<'EOF'
통합 보기 전환: 쪽 문서 viewerUrl을 view.html로, viewer.html은 쪽 문서를 넘김(엑셀·도면·한글은 그대로), 첫 쪽 시간 측정 스크립트

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- src/routes/documents.ts public/js/viewer.js public/api-docs.html README.md tests/viewShellRoutes.test.ts scripts/measure-view-first-paint.mjs
```

---

### Task 13: 전체 검증 — e2e·회귀·실제 브라우저 조작·휴대폰 CPU×4 (병합 전 멈춤)

**Files:**
- Create: `scripts/verify-view-shell-e2e.mjs`
- Modify: `package.json`(scripts에 `"verify:view-shell": "node scripts/verify-view-shell-e2e.mjs"` — `verify:view-only` 다음 줄)

**Interfaces:**
- Consumes: 전부. 시험 서버는 Global Constraints대로 8124(PID 기록).

- [ ] **Step 1: `scripts/verify-view-shell-e2e.mjs` 쓰기**

머리·도구 함수:

```js
// 통합 보기 껍데기(view.html) 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-view-shell-e2e.mjs [BASE] [스크린샷 폴더]
// 워드·PPT·PDF × 열람만(서버 PNG)·내려받기(pdf.js)·편집 토큰. 열람만은 PDF가 브라우저로 가지 않는다(derived-pdf·pdf-source 0건).
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import JSZip from 'jszip';
import { strToU8, zipSync } from 'fflate';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'view-shell-'));
fs.mkdirSync(shotDir, { recursive: true });
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
```

시험 문서(매번 만든다 — 고객 문서를 쓰지 않는다):
- `makePdf()` — pdf-lib로 3쪽, 쪽마다 `ViewShell Target page ${i}`(Helvetica 18pt, x 60 y 760).
- `makeDocx()` — `scripts/verify-view-only-e2e.mjs`의 `makeDocx`(42~55줄)를 그대로 옮기고 문단 글만 `보기 껍데기 시험 문서`·`두 번째 문단`으로.
- `buildFixture()`(pptx) — `scripts/verify-pptx-fidelity-e2e.mjs` 24~96줄(`NS`·`REL`·`rels`·`tree`·`xfrm`·`text`·`THEME`·`buildFixture`)을 그대로 옮기고 `SHAPES`만 글상자 하나로: `` const SHAPES = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="글"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(914400, 914400, 6400800, 1828800)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${text('보기 껍데기 슬라이드')}</p:txBody></p:sp>`; ``
- `upload(bytes, name, { wait = true })` — `verify-view-only-e2e.mjs`의 `upload`(57~69줄)에 `wait`이 false면 상태를 기다리지 않고 id만 돌려주는 갈래를 더한 것.
- `token(id, body)` — 같은 파일의 `token`(70~77줄), 기본 이름 `'보기시험'`.

브라우저 도구:

```js
const browser = await chromium.launch();
async function openView(url, { viewport = { width: 1400, height: 900 }, setup } = {}) {
  const ctx = await browser.newContext({ viewport });
  const page = await ctx.newPage();
  const errors = [];
  const requests = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));
  page.on('request', (r) => requests.push(new URL(r.url()).pathname));
  if (setup) await setup(page);
  await page.goto(BASE + url);
  await page.waitForSelector('.page-frame.page-ready', { timeout: 90000 });
  await page.waitForTimeout(800);
  return { ctx, page, errors, requests };
}
const pdfCalls = (reqs) => reqs.filter((p) => p === '/api/view/derived-pdf' || p === '/api/view/pdf-source').length;
const pngCalls = (reqs) => reqs.filter((p) => p.startsWith('/api/view/pages/')).length;
const shot = (page, name) => page.screenshot({ path: path.join(shotDir, `${name}.png`) });
```

확인 항목(각각 `check(조건, '한국어 설명', 세부)` 한 줄 이상, 끝에 오류 0건):

1. **권한별 안쪽**(PDF·워드·PPT 셋 다, 문서마다 열람만 `{}`·내려받기 `{ canDownload:true, canCopy:true, canPrint:true }` 토큰):
   - 토큰 응답 `viewerUrl`이 `/view.html?`로 시작.
   - 열람만: `pdfCalls === 0`, `pngCalls > 0`, `.pdf-text-layer span` 0개, `#edit-btn` 숨김, `#download-btn` 꺼짐이고 `title === '내려받기 권한이 없습니다'`, `#print-btn` 꺼짐.
   - 내려받기: `pdfCalls === 1`, `pngCalls === 0`(썸네일도 브라우저가 그림), `.pdf-text-layer span` 1개 이상, `#download-btn` 켜짐.
   - 둘 다 `#page-jump-total`이 `의 <쪽수>`, `#presence-list .presence-badge` 1개 이상, `getComputedStyle(document.documentElement).getPropertyValue('--watermark-bg')`가 비어 있지 않음, 오류 0. 스크린샷 `${형식}-${모드}`.
2. **막대(PDF, 내려받기 토큰)**: `#zoom-in` → `#zoom-level` `125%`; `#zoom-fit-btn` 두 번(너비 → 쪽) 뒤 첫 쪽 틀 높이가 창 높이 이하; `#page-next` → 입력값 `2`; 입력에 `3` + Enter → `3`; `#page-prev` → `2`; `#rotate-btn` → `#pages`에 `.pages-rotated`; Ctrl+휠(위) → 확대 %가 커짐.
3. **썸네일**: 1400px에서 `#thumbs-panel` 보임, `.thumb-item` 3개, 그림(`.thumb-box canvas, .thumb-box img`) 1개 이상, 세 번째를 누르면 입력값 `3`이고 `.thumb-item.current`가 세 번째. 열람만 토큰에서도 그림이 `img`로 옴.
4. **찾기(두 모드)**: Ctrl+F → `#search-bar` 보임; `Target` + Enter → `#search-count` `1 / 3`, `.search-hit-overlay.current` 1개 이상; Enter → `2 / 3`; Shift+Enter → `1 / 3`; Esc → 막대 숨김·강조 0개. 열람만(PNG)·내려받기(pdf.js) 둘 다.
5. **측정·반전(PDF)**: `#measure-btn` → `#measure-bar` 보임, 쪽 위 두 점 누르기 → `.measure-label` 글자가 `mm`를 포함; Esc → 측정 막대 숨김. `#invert-btn` → `.pages-invert`. 워드 문서에서는 두 단추가 꺼짐이고 `title === 'PDF 문서에서만 쓸 수 있습니다'`.
6. **표시(markup)**: 편집 토큰(`canEdit:true` 등)으로 PDF — `#markup-add-btn` 켜짐, 누르고 쪽 위를 끌면 팝오버가 뜨고 저장하면 `/api/view/markup`에 1건; 열람만 토큰에서는 꺼짐이고 `title === '표시를 넣을 권한이 없습니다'`. 글 넣기는 `verify:viewer-quicktext`가 맡는다.
7. **pdf.js 실패 → PNG**: `setup`에서 `page.route('**/api/view/derived-pdf*', (r) => r.fulfill({ status: 200, contentType: 'application/pdf', body: 'not a pdf' }))` — 쪽이 그려지고 `#view-notice`가 `서버 그림`을 포함, `pngCalls > 0`.
8. **편집 → 보기로 왕복(워드)**: 편집 토큰 `{ canEdit:true, canDownload:true, canCopy:true, canPrint:true }`의 `viewerUrl + '&home=hidden'`을 열고 `#edit-btn` 보임 → 누름 → 주소가 `/editor.html`이고 `from=view`·`home=hidden`이 있음 → `.ProseMirror[contenteditable="true"]` 기다림 → 첫 문단 끝에 `keyboard.insertText(' 껍데기왕복')` → `#back-to-view` 누름 → `/view.html`로 돌아옴(120초 안) → `.pdf-text-layer`의 글에 `껍데기왕복`이 있음(새로 변환한 PDF).
9. **편집기 바로 열기(향기랩 방식)**: 같은 토큰의 `editorUrl`을 그대로 열면 5초 뒤에도 `#back-to-view` 0개.
10. **PPT·PDF 편집기 "보기로"**: PPT 내려받기 토큰(읽기 전용 편집기)의 `editorUrl + '&from=view'` → `#back-to-view` 보임 → 누르면 `/view.html`. PDF는 `{ canDownload:true, canEdit:true, pdfEditor:true }` 토큰으로 보기 화면 → `#edit-btn` → `/editor-pdf.html?…from=view` → `.pdf-titlebar #back-to-view` → 누르면 `/view.html`.
11. **넘기기**: 워드 열람만 토큰으로 `/viewer.html?token=…`을 열면 `/view.html`에 도착하고 쪽이 그려짐. 엑셀(`ExcelJS`로 한 칸짜리 통합 문서 — `verify-view-only-e2e.mjs` 21줄 `makeXlsx`의 첫 시트·A1~B2만) 토큰의 `viewerUrl`은 `/viewer.html?`로 시작하고, `/view.html?token=…`으로 열면 `/viewer.html`로 넘어가 `#grid-table`이 그려짐.
12. **첫 변환 중(V8)**: PPT를 `wait:false`로 올리고 바로 내려받기 토큰으로 `viewerUrl`을 연다(`openView` 대신 직접 `goto` 후) — 2초 안에 `#prep-screen`이 보였거나 이미 쪽이 그려졌는지 기록(변환이 빠르면 못 볼 수 있다 — 세부에 어느 쪽인지 적고 통과 조건은 "결국 `.page-frame.page-ready`"와 "준비 화면이 보이는 동안 `#edit-btn` 보임"), `/api/view/info` 요청은 2번 이하(되풀이하지 않음).
13. **좁은 화면(390×844)**: 내려받기 토큰 PDF — `#tools-btn` 보임·화면 안(`boundingBox().x + width <= 390`), `#thumbs-panel` 숨김, `#text-tool-btn`·`#print-btn`이 `#tools-menu` 안, `#tools-btn` 누르면 메뉴 열림, 창을 1400으로 넓히면 메뉴가 비고 단추가 제자리. 스크린샷 `narrow`.
14. **iframe**: 같은 출처 페이지(`${BASE}/api-docs.html`)를 연 뒤 `window.addEventListener('message', …)`를 먼저 걸고 `iframe`(`viewerUrl + '&back=parent'`)을 넣는다 → `docview:ready` 받음; iframe 안 `#back-btn` 누름 → `docview:back` 받음.

끝:

```js
await browser.close();
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
```

올린 문서는 끝에 `DELETE /api/documents/:id`로 지운다.

- [ ] **Step 2: `package.json`** — `"verify:view-only": …,` 다음 줄에 `"verify:view-shell": "node scripts/verify-view-shell-e2e.mjs"`(앞 줄 끝 쉼표 확인).

- [ ] **Step 3: 단위·타입·빌드**

Run: `npx vitest run && npx tsc --noEmit && npm run build:editor`
Expected: 전부 통과.

- [ ] **Step 4: 시험 서버 띄우고 e2e 전부**(8124, PID 기록)

```bash
npm run verify:view-shell
npm run verify:view-only
npm run verify:viewer-quicktext
npm run verify:xlsx
npm run verify:pptx
npm run verify:pdf
npm run verify:pdf-search
```

Expected: 각각 마지막 줄 `ALL CHECKS PASSED`(또는 그 스크립트의 통과 표시). `verify:view-only`·`verify:viewer-quicktext`는 **고치지 않은 채로** 통과해야 한다(연동 제품 회귀 금지). 실패하면 원인을 고치고(스크립트를 고쳐 통과시키지 않는다) 다시.

- [ ] **Step 5: 실제 브라우저로 모든 새 컨트롤 조작(사용자 규칙 — 시험·스크린샷만으로 끝내지 않는다)**

헤드 있는 크롬(Playwright `headless:false` 또는 직접 크롬)으로 시험 서버의 워드·PPT·PDF를 열람만/내려받기/편집 토큰으로 열고 아래를 **손으로 눌러** 확인한다. 하나라도 이상하면 고친 뒤 Step 3부터 다시.

| 컨트롤 | 확인 |
|---|---|
| ←(home 링크 / `back=parent`) | 목록으로 감 / iframe이면 부모에 알림만 |
| 편집 | 편집기로, 돌아오면 쪽에 고친 글 |
| 썸네일 단추·썸네일 누르기 | 칸 열고 닫힘, 그 쪽으로 감, 지금 쪽 테두리 |
| 축소·확대·맞춤(너비↔쪽) | 단계·아이콘 바뀜, 맞춤 뒤 화면에 맞음 |
| 이전·다음·쪽 번호 입력 | 끝 쪽에서 단추 꺼짐, 입력 Enter·Esc·포커스 잃기 |
| 회전 | 90도씩, 표시·측정 겹침도 같이 돈다 |
| 찾기(Ctrl+F·F3·Shift+F3·Enter·Shift+Enter·Esc·×) | 두 모드, 작은 글자 자동 확대 |
| 인쇄(Ctrl+P 포함)·내려받기·다른 이름으로·"마크업 포함" | 새 창 PDF(워터마크), 파일 이름, 포함/제외 |
| 강조(끌기·팝오버 저장·지우기)·글 넣기(누름·Esc 두 번·글자 크기 −/+·끌어 옮기기) | 새로 고쳐도 남음 |
| 측정(두 점·기준 보정·지우기·×·Esc)·도면 모드 | mm 표시, 보정 뒤 실제 치수, 반전 기억 |
| 전체 화면 | 들어가고 Esc로 나옴, 아이콘 맞음 |
| 도구▾(800px·390px 폭) | 메뉴 열림·밖 누르면 닫힘·안 단추 동작 후 닫힘, 넓히면 제자리 |
| 글자 고르기·복사(pdf.js + 복사 권한) | 끌어 고르기·Ctrl+C·우클릭 복사, 복사 권한 없으면 글자 층 없음 |
| 편집기 "보기로" — 워드·엑셀·PPT(제목 표시줄)·PDF(머리) | `from=view`일 때만 보임, 저장 중 표시, 돌아옴 |
| 한글 "보기로"(`/editor-hwp.html?token=…&from=view`, 편집 토큰) | 고치지 않았으면 바로 감, 고쳤으면 [저장 후 이동 / 버리고 이동 / 취소] 세 갈래 각각 |
| 첫 변환 중(큰 PPT) | "쪽을 준비하는 중"·진행 막대, 편집 단추 먼저 됨 |

- [ ] **Step 6: 휴대폰 CPU×4 첫 쪽 시간**

`node scripts/measure-view-first-paint.mjs http://127.0.0.1:8124 /view.html 5` → Task 12 기준값(`first-paint-baseline.txt`)과 나란히 적는다. 열람만(PNG) 줄은 기준값 + 10% 이내여야 한다(같은 서버 PNG를 받으므로 — 넘으면 원인(묶음 크기·초기 요청 순서)을 찾아 고친다). 내려받기(pdf.js) 줄은 숫자를 그대로 보고한다(설계 §9 "지금 viewer.html 대비").

- [ ] **Step 7: 시험 서버 끄기(PID로만)·정리**

`taskkill //PID <기록한 PID> //T //F`. `git status --short`가 깨끗한지(산출물이 무시되는지) 확인.

- [ ] **Step 8: 커밋**

```bash
git add scripts/verify-view-shell-e2e.mjs
git commit -m "$(cat <<'EOF'
통합 보기 e2e(verify:view-shell): 권한별 안쪽·PDF 요청 0건·막대·썸네일·찾기·측정·표시·편집 왕복·넘기기·좁은 화면·iframe

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
EOF
)" -- scripts/verify-view-shell-e2e.mjs package.json
```

- [ ] **Step 9: 보고하고 멈춘다(병합 안 함)**

사용자에게: 브랜치 `view-shell`의 커밋 목록, e2e 결과(각 스크립트 통과 수), 실제 브라우저 확인표, CPU×4 숫자(기준값 대비), 남은 한계(아래), 그리고 "master 병합 = 운영 반영(8000·8010 재시작, 세션 끊김) — 승인해 주시면 반영 절차대로 진행"을 묻는다.

알려 둘 한계: 한글 문서는 1단계에서 지금 보기 화면(편집 단추 없음 — 2단계), 돌아간·세로 글자의 pdf.js 찾기 강조는 근사(`matchRects` shortcut), 큰 PDF를 휴대폰에서 pdf.js로 열 때 크기 기준 없이 실패 시에만 PNG로 내려감, viewer.js 안 쪽 보기 코드는 2단계에서 정리.

---

## 자체 점검(계획 작성자)

- **설계 범위:** V1(Task 3) V2(Task 12, 모호점 2) V3(Task 2·3) V4(Task 4, 모호점 1) V5(Task 1·4, Global) V6(Task 1·10) V7(Task 10) V8(Task 1·11). §2 화면(Task 3·5·6, 640px·도구▾·흐림+툴팁·워터마크·함께 보는 사람). §3 표 전 줄(Task 12 머리 참고). §4 두 구현·권한 선택·실패 시 PNG(Task 3·4). §5 서버 전부(Task 1, build 입구 Task 3). §6(Task 10, 모호점 4). §7 실패 문구·편집 링크(Task 11)·±2·휴대폰(Task 3·4)·경로 없는 오류(Global). §8 향기랩 변경 없음(Global). §9 단위·e2e·실제 브라우저·CPU×4(Task 1·2·13).
- **이름 일치:** `PagesBody` 멤버(`onZoom`·`onSourceChange`·`setZoomIndex`·`zoomAtPoint`·`showHits`·`clearHits`·`pageFrame`·`pageSize`·`currentPage`·`zoomPercent`·`container`)는 Task 3 코드와 Task 5~9 사용처가 같다. `ShellCtx.onEscape` 우선순위(10·20·30·40·90)는 ctx.ts 주석·measure·markup·search가 같다. `setupEditButton({ info })`는 `Pick<ShellCtx,'info'>`.
- **계획 작성 때 실제로 돌려 본 것**(버리는 워크트리, 운영 폴더·포트 무관): Task 1 시험 6개·Task 2 시험 19개 통과, Task 3~11 코드 전체 `tsc --noEmit` 0, `build-editor.mjs`로 `view.bundle.js`(정적 26 KB, pdf.js는 별도 조각) 빌드, 헤드리스로 PNG·pdf.js 두 모드의 쪽·썸네일·찾기(`1 / 3`)·쪽 이동·pdf.js 실패 → PNG 안내·390px 도구▾ 확인(그때 고친 것: `[hidden]` CSS, 좁은 화면 막대 규칙). Task 9(markup 옮기기)·Task 12 측정·Task 13 e2e는 쓰지 않았다.
