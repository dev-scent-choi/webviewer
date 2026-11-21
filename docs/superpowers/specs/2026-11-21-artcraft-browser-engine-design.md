# 두 가지로 열기 — 리브레오피스(서버 변환) / 아트크래프트(브라우저 처리) 설계

작성 2026-11-21. 가지 `office-artcraft`(office-quality-all 8103231 위). 근거: `C:\Users\OpenMinds\projects\artcraft-eval\REPORT.md`(PC 시험), 같은 폴더 `SPIKE-WASM.md`(브라우저 wasm 타당성 검사, 시험 코드 `spike-wasm\`).

## 1. 목표

문서를 여는 방식(토큰 발급 → 보기 주소)은 그대로 두고, **여는 엔진을 두 가지 중에서 고른다.**

| 옵션 | 처리 위치 | 쪽 그림의 출처 |
|---|---|---|
| **리브레오피스** (`engine: 'libreoffice'`, 기본) | 서버 — 지금 경로 그대로 | 서버가 LibreOffice로 PDF 변환 → `/view.html`이 pdf.js·PNG로 그림 |
| **아트크래프트** (`engine: 'artcraft'`) | 브라우저 — 원본 파일을 받아 wasm 엔진이 그림 | `/view.html`이 Worker 안 ArtCraft 엔진으로 쪽을 그림, 서버 변환 없음 |

같은 `/view.html`, 같은 도구 막대(확대·회전·쪽 이동·섬네일·찾기·인쇄·내려받기·마크업·참석자·편집 단추)를 쓴다. 다른 것은 **쪽 그림을 만드는 원천(PageSource) 하나**뿐이다.

## 2. 범위

| 형식 | 아트크래프트 옵션 | 이유 |
|---|---|---|
| .docx | **된다** (WordCraft) | 쪽 수 LibreOffice와 같음(고객 문서 1건 1쪽 차이), 한 쪽 20~56 ms |
| .pptx | **된다** (DeckCraft) | 장 수 같음, 합성 38~61 ms, 고객 문서 0.4초/장 |
| .doc | **된다, "실험" 표시** | 표 테두리·칸 너비가 깨지는 파일이 있음 → 실패·의심 시 리브레오피스로 자동 전환 |
| .ppt | 안 됨 → 리브레오피스 | DeckCraft가 .ppt를 못 읽음 |
| .xlsx .xls | 안 됨 → 지금 격자 화면 | GridCraft는 화면 그리기가 편집기(egui)에 묶여 있어 그림 출력이 없음(HTML 내보내기는 줄바꿈·테두리 빠짐). 엑셀은 원래 리브레오피스를 쓰지 않는다 |
| .hwp .pdf 도면 그 밖 | 안 됨 → 지금 경로 | 엔진 없음 |

안 되는 형식에 `engine: 'artcraft'`를 보내면 **거절하지 않고** 지금 경로로 열고, 응답에 `engine: 'libreoffice'`와 `engineReason`을 돌려준다. 메인 화면의 "아트크래프트로 열기" 단추는 안 되는 형식에서는 비활성 + 이유 말풍선.

## 3. 호출 방식

### 3.1 토큰 발급 `POST /api/documents/:id/tokens`
본문에 `engine?: 'libreoffice' | 'artcraft'`(기본 `libreoffice`) 하나를 더한다. 다른 필드는 그대로.

서버 판정(`engineFor(ext, perms, requested)`):
1. 요청이 `artcraft`이고 형식이 docx/pptx/doc이고 **`canDownload`가 참**이면 `artcraft`.
   - 원본 바이트가 브라우저로 내려가므로 내려받기 권한과 같은 뜻이다. 권한이 없으면 `libreoffice` + `engineReason: '내려받기 권한이 없어 서버 변환으로 엽니다'`.
2. 그 밖은 `libreoffice`(+ 형식이면 `engineReason: '이 형식은 브라우저 엔진이 지원하지 않습니다'`).

토큰 클레임에 `engine`을 넣고(위조 방지 — 주소의 쿼리는 보지 않는다), 응답은 `{ token, viewerUrl, editorUrl?, engine, engineReason? }`. `viewerUrl`은 지금과 같이 `/view.html?token=…` — 화면은 `/api/view/info`에서 `engine`을 받는다.

### 3.2 `/api/view/info`
응답에 `engine: 'libreoffice' | 'artcraft'`, `artcraft?: { kind: 'word' | 'deck', experimental: boolean, originalUrl, assets: { wasm, fonts[] } }`를 더한다. `originalUrl`은 지금 있는 원본 내려받기 경로(토큰 포함, `canDownload` 검사 그대로). `assets`는 내용 해시가 붙은 정적 주소(`/artcraft/wordcraft-<hash>.wasm`, `/artcraft/fonts/SCDream4-<hash>.otf`, `…SCDream6…`).

### 3.3 메인 화면(index.html)
docx/pptx/doc 줄에 "리브레오피스로 열기"(지금 "뷰어 열기") 옆에 **"아트크래프트로 열기"** 단추. 공유 링크 만들기 폼에도 `engine` 라디오 두 개(기본 리브레오피스). 안 되는 형식은 단추 비활성 + 이유.

## 4. 화면 구조(`/view.html`)

```
view.ts
  openSource(info)
    engine === 'artcraft' → openArtcraftSource(info)   ← 새 파일 view/artcraftSource.ts
    아니면               → 지금 그대로(pdf.js / PNG)
  body = new PagesBody(source, info.pages)             ← 변경 없음
  source 실패 → body.replaceSource(리브레오피스 원천) + 한 줄 안내   ← 지금 pdf.js→PNG 전환과 같은 길
```

### 4.1 `ArtcraftSource implements PageSource`
- `kind: 'pages-artcraft'`(BodyKind 합집합에 추가), `pageCount` = 엔진이 연 뒤 알려 준 값, `textSelectable` = 엔진이 글자 위치를 주면 `canCopy`, 아니면 false.
- `sizeOf(page)`: 엔진이 준 쪽 크기(pt). 열기 전에는 `info.pages`(서버가 알면) 아니면 null.
- `qualityKey = ceil(cssWidth × dpr)` (pdf.js와 같음).
- `draw(page, canvas, opts, signal)`: Worker에 `{render, page, scale}`(`scale = 원하는 화소 폭 ÷ 쪽 폭 pt`, 엔진 `sp_render(h, i, round(scale×1000))`) → RGBA 버퍼(transferable) → `putImageData`. 취소는 요청 id로 무시.
  - **두 단계 그리기**: 그 쪽을 요청 배율로 아직 그린 적이 없으면 먼저 `scale = 1`(쪽 폭 pt = 화소)로 그려 바로 보이고, 이어서 요청 배율로 다시 그려 바꿔 끼운다(고객 pptx 1600px 0.7~1.5초, 1× 는 그 1/3~1/4). 같은 쪽·같은 배율의 버퍼는 JS 쪽에서 쪽당 1개 보관(엔진에는 그림 캐시가 없다).
- `drawText`: 엔진 `sp_text(h, i)` → `{w, h, runs:[{t, x0, x1, y, s}]}`(pt). pdf.js와 같은 글자 층(`.textLayer` 안 span, 쪽 비율 좌표, 글자 크기 `s`). 세로·회전 글자는 엔진이 주지 않으므로 없음. `textSelectable = canCopy`.
- `search(q)`: 쪽마다 `sp_text`의 run 글자에서 찾고(대소문자 무시, 쪽당 0.3~23 ms, 결과는 보관), `SearchHit`(쪽, 비율 사각형 — run 상자 안에서 글자 수 비율로 자른 구간)을 만든다.
- `thumbnail(page, width)`: 같은 Worker에 작은 폭으로 그린다(한 번 만든 것은 보관).
- `release(page)`: 큰 배율 버퍼는 Worker 쪽에서 쪽당 1개만 유지. `destroy()`: Worker 종료(문서당 wasm 인스턴스 1개, 탭 닫힘과 함께 해제).

### 4.2 Worker(`view/artcraftWorker.ts`)
- 시작: wasm(brotli, `immutable` 캐시) + 글꼴 2개(S-Core Dream 4·6, 합계 brotli 약 0.45 MB) 받기 → `add_font` 두 번 → 원본 바이트 받기(`originalUrl`) → `open`.
- 외부 함수(엔진이 요구하는 import — 시간·시간대, 예외, wasm-bindgen 내부용)는 Worker 쪽에서 채운다(spike §8). 시간대(`Date`·타임존)만 진짜 값으로, 나머지는 호출되면 콘솔 경고 + 안전한 기본값.
- 도구 막대 상태 줄에 "브라우저 엔진(아트크래프트)" 표지를 보인다 — 쪽 수가 Word와 다를 수 있다는 것을 사용자가 알도록(고객 docx 1건: Word 7쪽·LibreOffice 6쪽·아트크래프트 5쪽 — 엔진의 맑은 고딕 줄 높이 계산 차이, spike §7).
- 요청은 큐로 하나씩, **보이는 쪽 우선**(PagesBody가 보이는 쪽부터 `draw`를 부르므로 큐 순서 = 호출 순서; `signal` 취소된 요청은 큐에서 뺀다).
- 실패 규칙(원천 → 전환): wasm 받기 실패, `open` 실패/오류, 첫 쪽 그리기 20초 초과, 쪽 수 0, 글꼴 받기 실패 → `ArtcraftUnavailable` 예외 → `view.ts`가 리브레오피스 원천으로 교체하고 안내 "브라우저 엔진이 이 문서를 열지 못해 서버 변환으로 보여 줍니다". .doc(실험)은 쪽 수가 서버 쪽 수(`info.pageCount`, 변환이 끝나 있을 때)와 **2쪽 넘게** 다르면 전환.
- 메모리: Worker 하나·문서당 약 35 MB(엔진 측정). 탭마다 별개.

### 4.3 그대로 두는 것
인쇄·내려받기·PDF 변환·마크업·측정·참석자·편집 단추·섬네일 막대·확대 단계·회전·가상화(보이는 쪽 ±2) — 모두 PagesBody·셸의 일이므로 변경 없음. 인쇄와 내려받기는 지금처럼 서버 PDF를 쓴다(브라우저 엔진 그림을 인쇄하지 않는다).

## 5. 자산과 빌드

- `tools/artcraft-shell/`(새, 저장소에 둠): spike의 C ABI 껍데기 크레이트 2개(word·deck)를 정리해서 옮긴다. ArtCraft 저장소 3개는 저장소 밖(`ARTCRAFT_DIR`, 기본 `C:\Users\OpenMinds\projects\artcraft-eval`)에서 **고정 커밋**으로 쓰고, 커밋 해시는 `tools/artcraft-shell/ARTCRAFT.lock`에 적는다.
- `npm run build:artcraft`: `cargo build --release --target wasm32-unknown-unknown`(GNU 도구 묶음, wasm-bindgen 없이) → `wasm-bindgen-cli`가 `tools/`에 있으면 크기 10% 줄이기(없으면 건너뜀) → brotli(`.br`) 함께 → `public/artcraft/<이름>-<내용해시>.wasm(.br)` + 글꼴 `public/artcraft/fonts/SCDream{4,6}-<해시>.otf(.br)` + `public/artcraft/manifest.json`(해시 주소 목록, `/api/view/info`가 읽음). 산출물은 git에 넣지 않는다(지금 `public/js` 번들과 같음). `npm run start`·`dev`는 manifest가 없으면 **아트크래프트 옵션을 끈 채**(`engine` 요청을 전부 libreoffice로, `engineReason: '브라우저 엔진이 설치되지 않았습니다'`) 뜬다 — 빌드 없이도 서버는 그대로 돈다.
- 정적 서빙: `/artcraft/*`는 `.br`가 있으면 그것을(`content-encoding: br`, 요청 `accept-encoding` 확인), `cache-control: public, max-age=31536000, immutable`(이름에 해시).
- 크기 상한(검사로 고정): word wasm brotli ≤ 5 MB, deck ≤ 2.5 MB, 글꼴 2개 brotli 합 ≤ 0.6 MB.

## 6. 보안

- 원본 바이트는 `canDownload` 토큰에만 내려간다(3.1). 토큰의 `engine`은 클레임으로만 믿는다.
- wasm은 같은 출처 정적 파일. Worker는 원본 주소·정적 자산 외에 요청하지 않는다. CSP의 `worker-src 'self'`, `script-src`에 `'wasm-unsafe-eval'`을 더한다(지금 CSP 값 확인 후 최소 추가).
- 조작된 문서로 Worker가 멈추거나 메모리를 다 쓰면: 그 탭만 영향. 20초 시간 제한 → 리브레오피스 전환. Worker 메모리 상한은 wasm 최대 메모리(엔진 설정)로 둔다.
- 서버 부담은 **줄어든다**(변환 없음). 단 메인 화면 "아트크래프트로 열기"도 서버가 `editReady`(리브레오피스 변환)를 **미리 돌리지 않게** 한다 — 전환이 일어날 때만 지금 경로의 변환이 시작된다(`info.processing` 흐름 그대로).

## 7. 검사

- 단위(vitest): `engineFor` 판정표(형식 × 권한 × 요청), `/api/view/info`의 `artcraft` 블록, manifest 없을 때 끄기, 정적 `.br` 선택.
- 브라우저(Playwright, 임시 포트): `scripts/verify-artcraft-e2e.mjs` — docx·pptx·doc 각 1개 이상 + 고객형 한글 문서(`.audit-run/test-server/originals`의 docx·pptx 사본, 저장소에 넣지 않음)로 (1) 아트크래프트로 열어 쪽 수·첫 쪽 그림(한글 글자 보임 — 쪽 그림의 비어 있지 않은 화소 비율과 LibreOffice 쪽 그림과의 유사도) (2) 확대·회전·쪽 이동·섬네일·찾기(글자 있을 때) (3) 서버 변환 요청이 **0건**인지 (4) wasm 주소를 막았을 때 리브레오피스로 자동 전환 + 안내 (5) 내려받기 권한 없는 토큰은 서버 변환으로 열림 (6) CPU×4에서 첫 쪽까지 시간 — 환산값(CDP 스로틀은 Worker를 늦추지 않으므로 메인 스레드 실측 + 4 × Worker 안 시간). 목표: 받기 제외 docx 2.5초 안, pptx·doc 2초 안, 고객 pptx 3초 안 (7) 두 번째 열기에서 wasm·글꼴이 캐시에서 옴.
- 회귀: 기존 verify:docx·verify:pptx·view-only e2e 그대로 통과(리브레오피스 경로는 변경 없음).
- 쪽 수 비교 기록: `docs/quality-audit/out/artcraft-pages.md` — 시험 문서별 LibreOffice 쪽 수 vs 아트크래프트 쪽 수.

## 8. 하지 않는 것

- 엑셀·한글·옛 .ppt의 브라우저 엔진(엔진 없음). 엑셀은 GridCraft에 화면 그리기 출력이 생기면 다시 본다.
- 브라우저 엔진으로 편집(편집기는 우리 것 그대로).
- 아트크래프트 그림으로 인쇄·PDF 만들기.
- 서비스 워커·오프라인.
- 글꼴 추가 등록(맑은 고딕 이름 등록은 메모리 3배, 효과 없음 — spike).
- 첫 쪽 속도 최적화(이번 범위 밖, 실기기 측정 뒤 별도 작업). 후보: wasm compileStreaming 캐시, 웜 스타트.

## 9. 결정 기록

- 엑셀을 아트크래프트 옵션에서 뺀 이유: HTML 내보내기는 지금 격자 화면보다 못하다(줄바꿈·테두리·인쇄 쪽 나눔 없음). 옵션 단추는 있되 비활성.
- 글꼴은 S-Core Dream 4·6만(한글 11,172자 전부, brotli 0.45 MB). Noto Sans KR(4.1 MB)은 쓰지 않는다.
- 쪽 그림 원천만 바꾸고 PagesBody·셸은 손대지 않는다 — 두 옵션의 화면이 같아야 한다는 요구 때문.
- (2026-11-21 최종 검토) CPU×4는 환산값이다 — Worker 시간 ×4라 실기기보다 비관적. docx 첫 쪽이 2.06~2.29초로 나와 docx 기준을 2.5초로 둔다. 최적화 후보: compileStreaming 캐시·웜 스타트.
