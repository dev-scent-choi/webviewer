# 통합 보기 껍데기 — 1단계 설계(껍데기 핵심과 쪽 문서)

작성 2026-11-12. 기준: webviewer master `b294bd8`. 이 문서는 전체 방향(§1)과 **1단계**만 설계한다. 2~4단계는 단계마다 따로 설계한다.

## 0. 결정 요약

**사용자가 정한 것(바꾸지 않는다)**

- 모든 형식을 **같은 모양의 보기 껍데기**로 보여 주고, **편집은 형식별 기존 편집기**로 넘어간다("껍데기만 통일, 편집은 다른 UI").
- 쓰는 곳: **향기랩 문서도구와 연동 제품(saturn 등) 둘 다**. 연동 제품은 지금처럼 `viewerUrl`을 받는다.
- **향기랩은 지금처럼 편집 모드로 바로 연다**(사용자 결정 2026-11-12). 향기랩에서 껍데기는 편집기가 없는 경우(편집 주소가 없을 때)에만 보인다.
- 쪽 그리기: **권한 따라 둘 다** — 내려받기 권한이 있으면 바꾼 PDF를 브라우저 pdf.js로, 없으면 지금처럼 서버 PNG만(PDF·원본이 브라우저로 가지 않음). 화면 모양과 단추는 같다.
- 형식 범위: **전부, 안쪽은 형식별** — 워드·PPT·PDF·HWP는 쪽 보기, 엑셀은 지금 격자 보기, DWG·NWD는 cadview가 껍데기 안에서 그린다.
- 구성: webviewer에 새 껍데기를 만들고 안쪽 그리기를 갈아 끼우는 방식(1안). 나중에 자체 그리기 엔진(엑셀 인쇄 쪽 → PPT → 워드)을 붙일 자리를 남긴다.

**이 설계가 정한 것**

| # | 결정 | 이유 |
|---|---|---|
| V1 | 새 껍데기는 TypeScript 묶음 `view.bundle.js`(기존 `scripts/build-editor.mjs`에 입구 추가), 화면 `public/view.html` | 편집기들과 같은 빌드·시험 틀. 2,800줄 묶이지 않은 `viewer.js`에 덧붙이면 형식마다 확대·검색 코드가 또 생긴다 |
| V2 | 이전 동안 `viewer.html`·`viewer.js`는 그대로 둔다. 쪽 문서 기능이 다 옮겨지면(§3 표 전부) `viewerUrl`을 `/view.html`로 바꾸고, `viewer.html`은 같은 조회 문자열로 `/view.html`에 넘기는 얇은 쪽이 된다 | 연동 제품의 주소·토큰 계약이 바뀌지 않는다 |
| V3 | 껍데기와 안쪽 사이 약속 `ViewBody`(§4) | 안쪽(PNG·pdf.js·격자·cadview·자체 엔진)을 갈아 끼울 수 있게 |
| V4 | pdf.js 경로는 PDF 편집기의 `src/client/pdf/pageView.ts`·`pdfDocument.ts`·`searchEngine.ts`를 재사용한다(pdfjs-dist 4.2.67 그대로) | 같은 그리기·검색 품질, 새 의존성 없음 |
| V5 | 바꾼 PDF를 주는 새 경로 `GET /api/view/derived-pdf` — `canDownload` 토큰만, 그 밖은 403. 원본이 PDF면 원본 바이트 | 권한 없는 열람에서 PDF가 브라우저로 가지 않음(사용자 결정) |
| V6 | 편집 단추는 `GET /api/view/info`(또는 같은 역할의 기존 세션 응답)에 더한 `editorUrl`이 있을 때만 보인다. 누르면 같은 창(iframe) 안에서 그 주소로 이동 | 토큰 계약(`editorUrlFor`)을 그대로 따른다 |
| V7 | 편집기 넷(워드·엑셀·PPT·PDF)과 HWP 편집기에 "보기로" 단추 — **껍데기에서 편집으로 들어온 경우(`from=view`)에만** 보인다. 자동 저장 편집기는 바로, HWP는 저장 안 한 내용이 있으면 확인 | 편집 → 보기 왕복. 향기랩처럼 편집기로 바로 연 곳에는 단추가 생기지 않는다 |
| V8 | 변환이 끝나지 않은 문서(`editReady`)는 쪽 자리에 "쪽을 준비하는 중"과 진행 표시, 편집 단추는 먼저 쓸 수 있다 | 지금 `/api/view/pages`가 첫 변환을 기다리는 동작과 맞춤 |

## 1. 전체 단계

| 단계 | 내용 |
|---|---|
| **1(이 문서)** | 껍데기 핵심과 쪽 문서: 워드·PPT·PDF, 권한별 PNG·pdf.js, 썸네일·검색·인쇄·내려받기, 편집 단추와 "보기로", 지금 보기 화면 기능 이전, `viewerUrl` 전환 |
| 2 | 엑셀 격자와 HWP를 껍데기 안으로(HWP는 LibreOffice 변환 품질이 낮아 rhwp 그리기 검토) |
| 3 | DWG·NWD: cadview 화면 숨김 모드와 iframe 명령 통로(데스크톱 다리 명령 재사용), 보안 조건 |
| 4 | 연동 가이드 갱신(향기랩은 편집 모드 바로 열기 유지 — 바꾸지 않는다) |
| 이후 | 자체 그리기 엔진: 엑셀 인쇄 쪽 → PPT → 워드는 마지막까지 LibreOffice |

## 2. 화면

```
┌───────────────────────────────────────────────────────────────────────────┐
│ ← 이름.docx │ ▢썸네일 │ − 100% + 맞춤▾ │ ◀ 3 / 12 ▶ │ ⟳ │ 🔍 │ 도구▾ │ 인쇄 내려받기 │ [편집] │
├──────┬────────────────────────────────────────────────────────────────────┤
│ 썸네일 │                       쪽 보기(안쪽)                                │
└──────┴────────────────────────────────────────────────────────────────────┘
```

- ← 는 `back=parent`일 때만(지금처럼 `docview:back`), `home=hidden` 그대로 따른다.
- 형식·권한에 없는 단추는 흐리게 하고 이유를 툴팁으로 보인다(인쇄 권한 없음 등).
- 도구▾ = 형광펜·글상자 표시(markup), 측정·반전(PDF), 전체 화면. 지금 보기 화면에 있는 것만.
- 640 px 미만: 썸네일 칸은 닫힌 채 시작, 막대는 두 줄로 접고 확대·쪽·편집만 첫 줄.
- 워터마크·함께 보는 사람 표시는 `src/client/watermark.ts`·`presence.ts`를 그대로 쓴다(지금 `viewer.js`의 중복 구현을 없앤다).

## 3. 지금 보기 화면에서 옮길 기능(전환 조건)

| 기능 | 지금 위치 | 1단계 |
|---|---|---|
| 확대·축소·맞춤(폭/쪽) | viewer.js | 옮김 |
| 쪽 번호 이동 | viewer.js | 옮김 |
| 회전 | viewer.js | 옮김 |
| 전체 화면 | viewer.js | 옮김 |
| 검색(서버) | `/api/view/search` | PNG 모드에서 그대로, pdf.js 모드는 브라우저 검색 |
| 인쇄 | `/api/view/print` | 그대로 |
| 내려받기·다른 이름으로(표시 포함/제외) | `/api/view/download` | 그대로 |
| 형광펜·글상자 표시 | `/api/view/markup` | 옮김(서버 저장 그대로) |
| 측정·반전(PDF) | viewer.js | 옮김 |
| 워터마크·함께 보는 사람 | viewer.js 중복 구현 | 공용 모듈로 |
| 썸네일 | 없음("준비 중") | **새로** |
| 글자 선택·복사(`canCopy`) | 없음 | **새로**(pdf.js 모드만) |
| 편집 단추·보기로 | 없음 | **새로** |

이 표가 모두 끝나야 V2의 `viewerUrl` 전환을 한다. 엑셀·DWG·HWP는 2·3단계까지 지금 화면(`viewer.html` 격자·svg)으로 열린다 — 전환 때 껍데기는 그 형식이면 `viewer.html`을 그대로 쓴다(되돌림 없음).

## 4. 안쪽 약속 `ViewBody`

```ts
interface ViewBody {
  readonly kind: 'pages-png' | 'pages-pdfjs' | 'grid' | 'cad' | 'frame';
  mount(host: HTMLElement): Promise<void>;
  pageCount(): number;                 // 쪽 개념이 없으면 0(막대의 쪽 칸 숨김)
  goto(page: number): void;
  onPage(cb: (page: number) => void): void;
  zoom(z: number | 'fit-width' | 'fit-page'): void;
  rotate(deg: 90 | -90): void;
  search?(q: string): AsyncIterable<{ page: number; rects: DOMRect[] }>;
  thumbnail?(page: number, width: number): Promise<HTMLCanvasElement | HTMLImageElement>;
  textSelectable: boolean;
  destroy(): void;
}
```

- `pages-png`: `/api/view/pages/:n?scale=`(지금 서버 그리기·캐시), 썸네일 `?thumb=true`, 검색 `/api/view/search`.
- `pages-pdfjs`: V5의 PDF를 받아 `pdfDocument.ts`로 열고 `pageView.ts`로 보이는 쪽만 그린다(가상 스크롤, 화면 밖 캔버스 해제). 글자 층은 `canCopy`일 때만 켠다.
- 껍데기는 토큰 권한으로 고른다: `canDownload` → `pages-pdfjs`, 아니면 `pages-png`. pdf.js 열기가 실패하면(암호·깨짐·메모리) `pages-png`로 내려가고 한 줄 안내.

## 5. 서버 변경

- `GET /api/view/derived-pdf?token=` — `canDownload` 아니면 403(같은 문구). 변환 전이면 202 `{processing:true}`. `Cache-Control: no-store`, `Content-Disposition` 없음(보기용). 원본이 PDF면 원본(또는 PDF 편집기가 저장한 현재본 — `pdf-current`와 같은 규칙).
- 세션 응답에 `editorUrl`(있으면)과 `viewerUrl`. 토큰 하나로 보기·편집 화면을 오가므로 새 토큰을 만들지 않는다.
- `scripts/build-editor.mjs` 입구에 `view` 추가. `public/view.html` 정적.
- 속도 제한·인증 미들웨어는 기존 `/api/view/*`와 같게.

## 6. 편집 ↔ 보기

- 편집 단추: `location.assign(editorUrl + 지금 조회 문자열의 home·back + from=view)`. 향기랩·연동 제품 모두 iframe 안에서 그대로 바뀐다(부모는 `docview:ready`를 다시 받는다).
- "보기로": `officeShell.ts` 제목 막대(워드·엑셀·PPT), PDF 편집기 머리, HWP 편집기 머리에 단추 하나. 주소는 세션 응답의 `viewerUrl`. HWP는 저장 안 한 내용이 있으면 "저장하지 않은 내용이 있습니다 — 저장하고 돌아갈까요?" [저장 후 이동 / 버리고 이동 / 취소].
- 돌아오면 껍데기는 쪽 그림을 새로 받는다(PNG는 캐시 무효 쿼리, pdf.js는 새로 받기). 편집 저장 → 다시 변환이 진행 중이면 V8 표시.

## 7. 오류·한계

- 변환 실패: 지금 보기 화면 문구 그대로. 편집기가 있으면 "편집 화면으로 열기"를 함께 보인다.
- 큰 PDF(pdf.js): 보이는 쪽 ±2만 캔버스 유지. 휴대폰에서 메모리 부족이면 PNG로 내려간다.
- 사용자에게 보이는 오류에 파일 경로를 넣지 않는다.

## 8. 미리 알려 둔 손해

향기랩은 지금 변환을 기다리지 않고 편집기를 바로 연다(`editReady`). 껍데기를 먼저 보여 주면 워드·PPT는 LibreOffice 변환이 끝날 때까지 쪽이 늦게 뜬다. 그래서 **향기랩은 편집 모드 바로 열기를 유지한다**(사용자 결정). 연동 제품에서 보기 화면을 쓸 때만 이 대기가 생기며, 지금 `viewer.html`과 같다.

## 9. 시험

- 단위(vitest): 권한 → 안쪽 선택, `ViewBody` 두 구현의 쪽 이동·확대·썸네일 요청, 검색 결과 모양, 편집·보기 주소 만들기, derived-pdf 권한(403·202·200).
- e2e(Playwright, `scripts/verify-view-shell-e2e.mjs`, `npm run verify:view-shell`): 워드·PPT·PDF × PNG·pdf.js × 권한(열람만·내려받기·편집). 열람만 토큰은 `/api/view/derived-pdf`·`pdf-source` 요청 0건, 편집 → 보기로 왕복 뒤 바뀐 글자가 쪽에 보임, `viewerUrl` 전환 뒤 기존 `verify:view-only`·`verify:viewer-quicktext` 통과.
- 실제 브라우저에서 모든 단추 조작, 휴대폰 CPU×4 첫 화면 시간(지금 `viewer.html` 대비).
- 시험 포트는 임시 포트만 — 8000·8010·8020·8100·8110 금지.
