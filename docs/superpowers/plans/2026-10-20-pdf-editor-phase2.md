# PDF 편집기 2단계(주석 — 새로 만들기 + 기존 주석 고치기) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** PDF 편집기의 "편집" 모드에서 메모·형광펜·밑줄·취소선·펜·글상자·사각형·타원·선·화살표를 만들고 고치고 지울 수 있고, 파일에 원래 있던 같은 종류의 주석도 고칠 수 있으며, 저장하면 다른 PDF 뷰어에서도 보이는 표준 PDF 주석(모양 포함)으로 들어간다.

**Architecture:** 편집기가 "다루는 종류"(Text·Highlight·Underline·StrikeOut·Ink·FreeText·Square·Circle·Line, 답글(IRT)이 아닌 것)를 정한다. 서버는 (1) 방을 열 때 그 주석들을 읽어 `pdfAnnots` 맵에 넣고(시딩 origin), (2) 편집기에 줄 PDF(`/api/view/pdf-source`)에서는 그 주석들을 빼며(편집기가 직접 그리므로 두 번 그려지지 않게), (3) 저장할 때 원본에서 그 주석들을 빼고 맵의 주석을 모양(AP) 스트림과 함께 다시 쓴다. 다루지 않는 주석(링크·양식·답글·도장 등)은 손대지 않는다. 사용자가 실제로 바꾼 것이 있을 때만 저장한다(`room.editedSinceSave`). 주석은 "원본 몇 쪽"(`pageKey: 'o:<0부터>'`)에 붙어 3단계 쪽 정리에도 따라간다.

**Tech Stack:** pdf-lib(서버 읽기·쓰기), @pdf-lib/fontkit + 동봉 Pretendard(한글 글상자 모양), Yjs(Y.Map + UndoManager), pdf.js 글자 층(형광펜 위치), SVG/DOM 겹침 층.

**Spec:** `docs/superpowers/specs/2026-10-20-pdf-editor-design.md` (§2.2·§2.3·§3 "기존 주석 고치기")

## Global Constraints

- 1단계 Global Constraints 그대로(한국어 문구, 8000 영향 없음, 회귀 검사, 반영 절차, 시험 서버 8124, PID로만 종료).
- 비밀 값 검사는 push와 한 명령으로 잇지 않는다. 검사 문자열을 문서·코드에 쓰지 않는다.
- 글꼴 넣기는 `subset: false`(pdf-lib+fontkit의 subset은 글자가 빠진다). 한글이 없는 글상자는 Helvetica(표준 글꼴, 파일 크기 그대로).
- 좌표는 원래 쪽의 PDF 좌표(회전 전, 아래가 y=0). 화면 변환은 pdf.js viewport(`convertToViewportPoint`/`convertToPdfPoint`)로만.

---

## File Structure

| 파일 | 책임 |
|---|---|
| `src/pdf/annotTypes.ts` (새, 브라우저·서버 공용) | `PdfAnnot` 형식, 종류 목록, `pageKeyOf`/`origIndexOf`, 기본 색 |
| `src/pdf/annotRead.ts` (새, 서버) | `readManagedAnnots(doc)`: 파일 주석 → `PdfAnnot[]`, `stripManagedAnnots(doc)`: 쪽에서 다루는 주석(+Popup) 빼기 |
| `src/pdf/annotWrite.ts` (새, 서버) | `writeAnnots(doc, annots)`: 표준 주석 사전 + 모양(AP) 스트림 |
| `src/pdf/pdfModel.ts` (수정) | `seedPdfModel`, `exportPdfFromRoom`(편집 있으면 빼고 다시 쓰기), `basePdfForEditor`, `PDF_SEED_ORIGIN`, 변경 기록 |
| `src/collab/room.ts` (수정) | pdf 시딩·저장 조건(editedSinceSave)·시딩 origin 자동저장 제외 |
| `src/routes/viewer.ts` (수정) | `pdf-source`가 다루는 주석을 뺀 PDF를 준다(작은 캐시) |
| `src/client/pdf/pageView.ts` (수정) | 쪽마다 주석 층(`div.pdf-annot-layer`)과 viewport 알림(`onPageLayout`) |
| `src/client/pdf/annotLayer.ts` (새) | 주석 그리기(SVG·글상자 DOM), 선택·이동·삭제, 도구별 만들기(끌기·클릭·글자 선택) |
| `src/client/pdf/annotUi.ts` (새) | 떠 있는 도구 줄(편집 모드 도구), 속성 상자(색·굵기·투명도), 메모 창(정보 창 "메모"), 실행 취소/다시 실행 |
| `src/client/editor-pdf.ts`·`public/editor-pdf.html`·`public/css/pdf-editor.css` (수정) | 편집 모드 연결, 도구 줄 단추, 메모 창 |
| `tests/pdfAnnots.test.ts` (새) | 읽기·빼기·쓰기 왕복, 한글 글상자, 다루지 않는 주석 보존, 편집 없으면 원본 그대로 |
| `scripts/verify-pdf-editor-e2e.mjs` (수정) | 2단계 항목 추가 |

## Tasks

### Task 1: 주석 형식과 읽기·빼기 (서버)
- `PdfAnnot`:
  ```ts
  export type PdfAnnotType = 'note' | 'highlight' | 'underline' | 'strike' | 'ink' | 'freetext' | 'rect' | 'ellipse' | 'line' | 'arrow';
  export interface PdfAnnot {
    id: string; pageKey: string; type: PdfAnnotType;
    rect: [number, number, number, number];          // x1,y1,x2,y2 (PDF 좌표)
    quads?: number[];                                 // 형광펜류: 8개씩(왼위·오른위·왼아래·오른아래 — Acrobat 순서)
    paths?: number[][];                               // 펜: 획마다 [x,y,x,y…]
    line?: [number, number, number, number];          // 선·화살표
    color: string; fill?: string | null; opacity: number; width: number;
    text?: string; fontSize?: number;
    author: string; createdAt: string; modifiedAt: string;
  }
  ```
- 읽기: Subtype Text→note, Highlight/Underline/StrikeOut→highlight/underline/strike(QuadPoints), Ink(InkList), FreeText(Contents·DA 글자 크기·C 또는 DA 색), Square→rect, Circle→ellipse, Line(L, LE에 OpenArrow/ClosedArrow 있으면 arrow), C/IC 색(RGB·회색·CMYK→RGB), CA 투명도, BS/W 또는 Border 굵기, T 작성자, M/CreationDate, NM은 id로(없으면 `f<쪽>-<순번>`). IRT가 있는 답글·Popup·그 밖의 종류는 읽지 않는다.
- 빼기: 각 쪽 `/Annots`에서 읽은 것과 같은 조건의 주석과 그 `/Popup`을 뺀다(배열이 비면 키 삭제).
- 테스트: pdf-lib로 각 종류 + Link + 답글 + Popup 있는 PDF를 만들어 → 읽은 목록·뺀 뒤 남은 주석(Link·답글)이 맞는지.

### Task 2: 쓰기(모양 스트림 포함)
- 종류별 사전(Type Annot, Subtype, Rect, NM, T, M, CreationDate, C, CA, Border/BS, F=4(인쇄)) + AP/N 스트림:
  - 형광펜: 사각형 채우기(곱하기 혼합 `/ca`, ExtGState BM Multiply), 밑줄: 사각 아래쪽 선, 취소선: 가운데 선 — 모두 quads 기준.
  - 펜: 획마다 `m l … S`(둥근 끝·이음), 선·화살표: 선 + 끝 화살촉(LE), 사각형·타원(베지어 4개): 테두리(+채우기 IC).
  - 메모: 노란 말풍선 아이콘(24×24pt) + Contents(UTF-16 hex), Popup 없이(뷰어가 자체 창).
  - 글상자: 상자 테두리 없음, 글자 줄바꿈(상자 폭 기준, 글자 폭은 글꼴 widthOfTextAtSize), DA `/Helv 12 Tf 0 0 0 rg`, 한글 등 WinAnsi 밖 글자가 있으면 Pretendard(subset false)로 AP 글꼴.
- 테스트: 쓴 PDF를 pdf-lib로 다시 열어 Task 1 읽기로 왕복(같은 값), pdf.js(서버 쪽 pdfjs-dist legacy)의 getAnnotations로 subtype·hasAppearance 확인, 한글 글상자 AP에 글꼴 리소스가 있는지.

### Task 3: 방 모델 연결
- `PDF_SEED_ORIGIN`(자동저장 제외), `seedPdfModel(original, ydoc)`(pdf-lib로 못 열면 비워 둠), `exportPdfFromRoom(original, ydoc)`: 원본 load → strip → write(맵 전체) → save(`useObjectStreams: false`), `basePdfForEditor(bytes)`: strip만.
- room.ts: pdf 시딩, update 핸들러 제외 origin에 PDF_SEED_ORIGIN, flush/종료 저장은 `editedSinceSave`일 때만, 변경 기록 `PdfChangeTracker`(추가·삭제·변경, 열람자·시각).
- viewer.ts `pdf-source`: `basePdfForEditor(현재 바이트)` — 문서 id + 바이트 길이 + 해시 앞부분으로 작은 캐시(최근 8개).
- 테스트: 편집 없으면 원본 그대로, 맵에 주석 추가 → 저장 결과에 주석, 원래 주석 하나 지움 → 저장 결과에서 사라지고 다루지 않는 Link는 남음.

### Task 4: 편집기 — 주석 층과 그리기
- pageView: 쪽마다 `div.pdf-annot-layer`(글자 층 위), 그린 뒤 `onPageLayout(page, layer, viewport)` 알림, `viewportOf(page)`.
- annotLayer: 맵 → 쪽별로 SVG(형광펜 multiply, 밑줄·취소선 선, 펜 폴리라인, 사각·타원·선·화살표) + 메모 아이콘 단추 + 글상자 div. 선택하면 파란 테두리·조절점.

### Task 5: 편집기 — 도구
- "편집" 탭 → 떠 있는 도구 줄에 메모·형광펜(▾ 밑줄·취소선)·펜·글상자·도형(▾ 사각형·타원·선·화살표)·지우개·실행 취소·다시 실행이 나타난다(읽기 전용 토큰은 없음).
- 형광펜류: 글자를 끌어 고르면 손을 뗄 때 그 줄들로 주석(글자 층 선택 → 줄별 사각형 합치기 → PDF quads).
- 펜: 끄는 동안 획, 손을 떼면 저장(점 줄이기). 도형·글상자: 끌어 상자. 메모: 누른 곳에 아이콘 + 바로 내용 입력.
- 선택 도구: 주석 클릭 선택, 끌어 이동, Delete로 삭제, 글상자 두 번 클릭 → 글 고치기, 메모 클릭 → 내용 창.
- 속성 상자: 색 6가지 + 직접 고르기, 굵기, 투명도 — 고른 도구 기본값과 선택한 주석에 적용.
- 정보 창 "메모" 단추(오른쪽 줄)와 목록(쪽·종류·작성자·시각·내용, 눌러 이동, 지우기).
- Ctrl+Z / Ctrl+Y(UndoManager, 내 편집만).

### Task 6: 확인·반영
- 단위: `tests/pdfAnnots.test.ts`. 브라우저: verify-pdf-editor-e2e.mjs에 2단계 항목(종류별 만들기 → 저장 파일을 서버 쪽 pdf.js로 열어 subtype 확인, 기존 주석 고치기·지우기, 한글 메모·글상자, 실행 취소, 다른 참여자 화면에 보임, 읽기 전용).
- 회귀 전부 → master merge → build → 8010 재시작 → push(검사와 분리) → 운영 확인.
