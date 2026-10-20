# PDF 편집기 3단계 — 페이지 정리 실행 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 쪽 회전·삭제·빈 쪽·순서 바꾸기·다른 PDF 넣기(파일 합치기)·고른 쪽 추출·나누기를 편집기에서 하고, 저장 파일에 그대로 반영한다.

**Architecture:** 쪽 목록을 방의 `pdfPages`(Y.Array)에 둔다. 주석은 지금처럼 `pageKey`로 쪽에 붙고, 쪽 목록 항목도 같은 열쇠를 쓴다
(`o:<원래 순번>`, `p:<부속 id>:<순번>`, `b:<무작위>`). 저장은 **방을 열 때의 원본**에서 매번 다시 만든다(여러 번 저장해도 같은 결과).
넣은 PDF(부속)는 방 메모리에 두고 저장 때 `copyPages`로 복사한다. 편집기는 쪽 목록으로 화면을 다시 짠다.

**Tech Stack:** pdf-lib(저장), pdfjs-dist(그리기), Yjs, fflate(나누기 zip), Playwright(e2e).

**Spec:** docs/superpowers/specs/2026-10-20-pdf-editor-design.md §1(페이지 정리 화면), §2.2, §2.3(2·6), §4-3

## Global Constraints

- 8000(유지보수 시스템) 동작은 바뀌지 않는다 — 새 경로는 모두 `perms.pdfEditor`가 있어야 한다.
- 기존 편집기(워드·엑셀·PPT)와 1·2단계 기능은 그대로: `npm test`, `verify:xlsx`, `verify:pptx`, `verify:pdf`, `verify:pdf-annots`, 워드 리본 e2e.
- 화면 문구는 한국어. 비밀값은 저장소에 넣지 않는다.
- 저장은 바꾼 것이 있을 때만(`editedSinceSave`).

---

### Task 1: 쪽 목록 모델과 저장(서버)

**Files:**
- Modify: `src/pdf/annotTypes.ts` — `PdfPageEntry`, `parsePageKey`, `PDF_META_MAP`
- Modify: `src/pdf/pdfModel.ts` — 시딩(쪽 목록·원본 지문), `exportPdfFromRoom(original, ydoc, edited, parts?)`, `extractPdf`
- Modify: `src/collab/room.ts` — `originalPdfBytes`(방을 열 때의 원본), flush·getCurrentBytes가 그것을 기준으로
- Test: `tests/pdfPages.test.ts`

**Interfaces:**
- `interface PdfPageEntry { key: string; rotate: 0|90|180|270; blank?: { w: number; h: number } }`
- `parsePageKey(key): { kind: 'orig'; index } | { kind: 'part'; partId; index } | { kind: 'blank' } | null`
- `exportPdfFromRoom(original, ydoc, edited = true, parts?: (id) => Uint8Array | undefined): Promise<Uint8Array>`
- `extractPdf(fullBytes: Uint8Array, positions: number[][]): Promise<Uint8Array[]>`
- `pdfBaseFingerprint(bytes): string` (sha1 hex 앞 16자)

- [ ] 시험 먼저: 순서 바꾸기·지우기·회전(원래 /Rotate에 더함)·빈 쪽·부속 쪽, 지운 쪽의 주석은 빠짐, 같은 열쇠가 두 번이면 첫 번째만, 두 번 저장해도 같은 쪽 수,
      추출(고른 쪽만·주석 포함), 시딩(쪽 목록·지문, 시딩은 편집으로 안 셈)
- [ ] 구현: pdf-lib로 원본을 열어 다루는 주석을 빼고, 원래 쪽 객체를 모두 페이지 트리에서 뺀 뒤 목록 순서대로 다시 넣는다
      (원래 쪽은 같은 객체를 다시 끼움 — 양식·책갈피가 끊기지 않게, 부속 쪽은 copyPages, 빈 쪽은 addPage([w,h])).
      주석은 `writeAnnots(doc, annots, key => pageOfKey.get(key))`, 목록에 없는 쪽의 주석은 버린다.
- [ ] room.ts: pdf 방을 열 때 `room.originalPdfBytes = original`, flush는 `exportPdfFromRoom(room.originalPdfBytes, …, parts)`,
      getCurrentBytes도 같은 기준(바꾼 것이 없으면 저장본 그대로).
- [ ] 커밋

### Task 2: 부속 PDF·추출 경로, 변경 기록

**Files:**
- Modify: `src/collab/room.ts` — `room.pdfParts`, `addPdfPart`, `readPdfPart`, `pdfRoomSnapshot`
- Modify: `src/routes/viewer.ts` — `POST /api/view/pdf-part`, `GET /api/view/pdf-part/:partId`, `POST /api/view/pdf-extract`, pdf-source는 방이 있으면 방의 원본 기준 + `x-pdf-base` 머리글
- Modify: `src/pdf/pdfModel.ts` — `preparePart(bytes, partId)`(다루는 주석을 읽어 열쇠를 바꾸고 뺀 PDF), PdfChangeTracker의 쪽 정리 요약
- Test: `tests/pdfPartRoutes.test.ts`

- [ ] 부속 올리기: 토큰(pdfEditor·canEdit·canDownload), 방이 있어야 함(없으면 409), 크기 ≤ maxUploadBytes, 방마다 합계 ≤ maxUploadBytes×4,
      pdf-lib로 열리지 않거나 암호면 400. 응답 `{ partId, pageCount, sizes:[{w,h,rotate}], annots }`(주석 id·열쇠는 부속 기준으로 새로).
- [ ] 부속 받기: 같은 방의 것만, 없으면 404.
- [ ] 추출: `{ groups: string[][] }`(쪽 열쇠) → 한 묶음이면 pdf, 여럿이면 zip(`이름-1.pdf`…). canDownload 필요. 방이 있으면 방 상태 기준, 없으면 저장본(`o:i`).
- [ ] 변경 기록: 쪽 목록이 처음과 다르면 한 줄 `pages_changed`(지움 n · 넣음 n · 돌림 n · 순서 바꿈), 마지막으로 바꾼 사람·시각.
- [ ] 커밋

### Task 3: 편집기 — 쪽 목록으로 보기

**Files:**
- Modify: `src/client/pdf/pageView.ts` — 쪽 목록(`setPages(entries)`), 원본·부속·빈 쪽 그리기, `drawThumbnail`, 찾기 글자 캐시는 열쇠 기준
- Modify: `src/client/pdf/panels.ts` — 썸네일을 view로 그리고 쪽 목록이 바뀌면 다시 짬, 책갈피는 `pageOfKey('o:i')`
- Modify: `src/client/pdf/annotUi.ts` — 메모 목록의 쪽 번호는 지금 순서 기준, 목록에 없는 쪽의 주석은 숨김, 실행 취소는 쪽 목록도 함께
- Modify: `src/client/pdf/findController.ts` — 쪽 목록이 바뀌면 다시 찾기
- Modify: `src/client/editor-pdf.ts` — `pdfPages` 관찰 → `view.setPages`, 동기화 뒤 원본 지문이 다르면 원본을 다시 받음

- [ ] 처음엔 원본 순서로 바로 그리고(기다림 없음), 방과 맞춰지면 목록대로 다시 짠다.
- [ ] 같은 열쇠의 쪽 자리는 다시 쓰고(그림 유지), 회전이 바뀐 쪽만 다시 그린다.
- [ ] 커밋

### Task 4: 페이지 정리 화면

**Files:**
- Create: `src/client/pdf/pageOps.ts` — 쪽 목록 바꾸기(LOCAL_ORIGIN 트랜잭션): rotate, remove(전부는 안 됨), insertBlank, insertEntries, move
- Create: `src/client/pdf/organizeView.ts` — 카드 격자, 고르기(클릭·Ctrl·Shift·Ctrl+A), 끌어 순서 바꾸기, 위 도구 줄
- Modify: `public/editor-pdf.html`, `public/css/pdf-editor.css`, `src/client/editor-pdf.ts`(모든 도구: 페이지 정리·파일 합치기·페이지 추출 연결)

- [ ] 도구 줄: 왼쪽으로 돌리기 · 오른쪽으로 돌리기 · 지우기 · 빈 쪽 넣기 · 파일에서 넣기 · 앞으로 · 뒤로 · 고른 쪽 추출 · 나누기(N쪽씩/고른 쪽 앞에서) · 닫기
- [ ] 단축키: Delete(지우기), Ctrl+A, Ctrl+Z/Y, Esc(닫기)
- [ ] 읽기 전용이면 정리·합치기 숨김, 추출은 canDownload면 보임
- [ ] 커밋

### Task 5: 시험·반영

- [ ] `scripts/verify-pdf-organize-e2e.mjs`(`npm run verify:pdf-organize`): 정리 화면 조작 → 저장 파일(서버 pdf.js로 쪽 수·글자·회전·주석) → 다시 열기, 다른 참여자 화면, 읽기 전용, 추출·나누기 내려받기, 실행 취소
- [ ] 회귀 전부 → master merge → build → 비밀값 검사(따로) → push → 8010 재시작 → 향기랩에서 운영 확인 → 메모리 갱신
