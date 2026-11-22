# 아트크래프트 앱 화면 열기 + 콘텐츠만 보기 옵션 설계

작성 2026-11-22(사용자 재지시로 되살림). 가지 `office-artcraft`(리본 스킨 작업 뒤에 이어짐). 근거: `C:\Users\OpenMinds\projects\artcraft-eval\SPIKE-WEBAPP.md`(시험 코드 `spike-webapp\`, 화면 `shots\`).

## 1. 사용자 요구(2026-11-22)
세 가지 열기 방식이 있어야 한다.
1. **아트크래프트로 열기** = 아트크래프트 앱 **자체 화면**(앱의 리본·도구 그대로). 우리 편집기에 입힌 것이 아니다.
2. **리브레오피스(서버 변환)·우리 편집기로 열기** = 우리 편집기 화면 + 아트크래프트식 도구 UI(별도 설계 `2026-11-22-artcraft-ribbon-skin-design.md`, 이 문서 범위 밖). 열기/편집 모드는 같은 화면이고 도구 활성·비활성만 다르다.
3. **콘텐츠만 보기** = 도구를 **전부 없애고** 문서 내용만 보이는 옵션.

## 2. 시험으로 확인된 사실(SPIKE-WEBAPP)
- 세 앱(wordcraft-web·deckcraft-web·gridcraft-web)은 cargo + wasm-bindgen 0.2.129 로 빌드되고, 같은 출처 iframe 안에서 WebGPU·WebGL2 둘 다 동작. COOP/COEP 불필요.
- 크기(brotli, 크기 최적화 빌드): 워드 7.9 MB, PPT 4.9 MB, 엑셀 4.5 MB (+ JS 0.02 MB). 첫 화면 0.22~0.39초, 문서 표시까지 워드 1.1~1.3초·PPT 0.7초·엑셀 0.55초. wasm 메모리 64/31/9 MB, 브라우저 탭 전체 +210~340 MB(GPU 포함).
- 문서 넣기: 패치 없이 `?host=parent` → 앱이 `<app>:ready` postMessage → 부모가 `{type:"<app>:open", name, bytes}` 전송. 앱 → 부모: `:save {name, saveAs, bytes}`, `:dirty`, `:failed`. 명령 API `window.<app>.execute()`.
- 한글: 본문은 `execute("font.add", base64)` 로 S-Core Dream 4·6 을 넣으면 정상(0.2초), 메뉴·칸 글자는 시스템 글꼴로 나옴. 메뉴 자체는 영어.
- 읽기 전용: 워드 `review.restrict readOnly`(+암호), 엑셀 `?chrome=none` + `review.protectSheet` 로 입력 차단 확인. PPT는 막는 명령이 없음(Session::execute 약 10줄 + web.rs `?readonly` 약 3줄 패치 추정, 미시험).
- 저장: Ctrl+S/`file.save` 때 부모가 바이트를 받음(다운로드 없음).
- 고객 xlsx 하나는 열자마자 "묶인 도형 건너뜀" 경고 대화상자가 뜨고 부모가 막을 수 없음.

## 3. 아트크래프트 앱 화면 — 범위와 결정
| 항목 | 결정 |
|---|---|
| 대상 형식 | docx·xlsx·pptx. .doc/.xls/.ppt 는 앱이 못 열면 `:failed` → 리브레오피스 경로로 자동 전환 |
| 열기 옵션 | `engine:'artcraft'` 토큰의 기본 보기 = **앱 화면**. 메인 화면 "아트크래프트로 열기" = 앱 화면. 지금의 엔진 쪽 그림 보기(ArtcraftSource)는 `ui=content`(§4)와 실패 전환용으로 남긴다 |
| 읽기 전용 | 토큰 canEdit=false 면 워드 `review.restrict`, 엑셀 `protectSheet`, PPT 는 1차에서는 편집 가능 상태로 열되 저장을 무시하고 "이 화면의 편집은 저장되지 않습니다" 띠(PPT 읽기 전용 패치는 2차) |
| 저장 | 부모가 `:save` 바이트를 받아 canEdit 토큰이면 `PUT /api/view/artcraft-save`(기존 저장 경로의 권한·감사·버전 규칙 그대로, 편집 방이 열려 있으면 거절 + 안내). canEdit 아니면 무시 + 띠 |
| 우리 띠 | 앱 위 **제목 줄 하나**(24px, 리본 스킨과 같은 규격): 문서 이름·돌아가기·읽기 전용 표지·저장 상태·"편집 열기". 전부 한글. 앱 메뉴는 영어 그대로(craft-i18n 유무를 1차 중 확인, 되면 켜고 안 되면 2차) |
| 자산 | `tools/artcraft-shell` 옆 `tools/artcraft-app/`: 세 앱 빌드 스크립트(`npm run build:artcraft-app`, ARTCRAFT.lock 고정, 크기 최적화, brotli, 해시 이름) → `public/artcraft-app/<app>-<hash>.{wasm,js}` + manifest. 산출물은 git 밖. manifest 없으면 앱 화면은 꺼지고 엔진 보기로 |
| 글꼴 | 기존 `/artcraft/fonts/SCDream{4,6}` 재사용, 앱 준비 뒤 `font.add` 두 번 → 문서 전송 |
| 실패 전환 | wasm 받기 실패·`:failed`·준비 20초 초과 → 엔진 보기 → 그것도 실패면 리브레오피스 + 안내 한 줄 |
| 보안 | 원본은 `/api/view/engine-source`(canDownload+artcraft 토큰). iframe 같은 출처, `sandbox` 없음(WebGPU 필요), 우리가 빌드한 정적 파일만. 앱 → 부모 메시지는 `origin` 검사 |

## 4. 콘텐츠만 보기(`ui=content`)
- 주소 꾸밈 `ui=content`(기존 `ui=view` 옆). 토큰 종류와 무관하게 **도구를 전부 숨기고** 문서만 보인다. 편집 가능 토큰이라도 이 화면에선 편집하지 않는다(키·클릭 모두 `ui=view` 보기 규칙 적용 + 화면 요소 숨김).
- 적용 화면:
  - 우리 편집기(docx·xlsx·pptx 편집기 화면): 제목 줄·탭 줄·리본·상태 줄·섬네일 레일·수식 줄을 숨기고 본문 영역만 전체 폭. 확대·쪽 이동은 키보드(기존 허용 키)와 스크롤로.
  - 리브레오피스 보기 화면(`/view.html`): 도구 막대 전체 숨김, 쪽만.
  - 아트크래프트: 앱 화면 대신 **엔진 쪽 그림 보기**(ArtcraftSource, 도구 없음)를 쓴다 — 이미 콘텐츠만 그리는 화면이라 앱을 띄울 필요가 없고 메모리도 작다.
- 메인 화면·토큰 발급: `/api/documents/:id/tokens` 응답의 `viewerUrl` 옆에 `contentUrl`(= 같은 토큰 + `ui=content`)을 더한다. 메인 화면 단추 줄에 "콘텐츠만 보기" 한 개 추가.
- 보안 규칙 변화 없음: 숨김은 화면 꾸밈일 뿐, 권한은 토큰이 정한다(canEdit 토큰으로 `ui=content` 를 열어도 서버 쪽 편집 권한은 그대로 — 화면에서만 막음).

## 5. 검사
- 단위: manifest 없음 → 앱 화면 옵션 꺼짐, origin 검사, 저장 경로 권한(열람만·canEdit·편집 방 열림), `ui=content` 판정 함수(viewStrip 식 순수 함수).
- 브라우저 e2e `scripts/verify-artcraft-app-e2e.mjs`: docx·xlsx·pptx 각각 (1) 앱 화면이 뜨고 문서가 보임(앱 `execute` 로 쪽/시트 수 확인) (2) 읽기 전용에서 타자 → 변화 없음(워드·엑셀), PPT 는 띠 (3) canEdit 토큰 Ctrl+S → 서버 저장 → 다시 열면 반영 (4) wasm 막기 → 전환 + 안내 (5) 서버 변환 요청 0건 (6) 두 번째 열기 캐시 (7) 우리 제목 줄 24px·한글.
- `scripts/verify-content-only-e2e.mjs`: 세 편집기·view.html·아트크래프트 각각 `ui=content` 에서 도구 요소(제목 줄·탭·리본·상태 줄·도구 막대) 0개 보임, 본문 보임, 타자·붙여넣기 무효, 600px 가로 스크롤 없음.
- 기존 verify:artcraft·readonly·docx·pptx·xlsx:browser·view-only·view-shell·ribbon-skin 그대로 통과.

## 6. 하지 않는 것(1차)
- PPT 읽기 전용 패치(2차), 앱 메뉴 한글화(craft-i18n 확인 뒤), 동시편집 연동(앱 화면은 혼자 편집), 폰·저사양 기기 측정, wasm-opt 추가 축소.
