# 오피스 품질 세 갈래 통합 (office-quality-all)

기준 커밋 94ab62b에서 세 갈래를 차례로 `--no-ff`로 합쳤다. 합칠 때마다 검증했고, 통합 때 고친 것은 그 병합 바로 뒤에 따로 커밋했다.
master에는 합치지 않았다. push·배포도 하지 않았다. 운영 폴더(`webviewer`, 8000·8010)는 건드리지 않았다.
시험 서버는 이 작업 폴더에서 8146(검증용)과 8147(audit·verify:reconnect가 스스로 띄움)에 띄웠다. 데이터·임시·LibreOffice 프로필은 `.audit-run/` 아래에 따로 뒀다.
정답 PDF는 `webviewer-quality/docs/quality-audit/out`에서 복사했다. Office로 다시 만들지 않았다(`--no-office`).

| 순서 | 커밋 | 내용 |
|---|---|---|
| 1 | 516c9e1 | 병합: office-quality-xlsx @9389052 (충돌 없음) |
| 2 | 2c0219a | 병합: office-quality-pptx @2c2678a (충돌 2개, 아래) |
| 2a | e3e906c | 통합 시험: 방 닫기 유예·세대 × 엑셀 값 개수 표·표 규칙·전역 상한 (`tests/integrationRoomGraceXlsx.test.ts`) |
| 3 | dd4cc0e | 병합: office-quality-docx @ddd4304 (충돌 3개, 아래) |
| 3a | 6142bae | 통합: 방이 닫히면 워드 보존 조각의 문서별 기록을 놓음 + 유예·세대 × 조각 저장소 시험 (`tests/integrationRoomGraceDocx.test.ts`) |

## 충돌과 해결

| 병합 | 파일 | 충돌 | 해결 |
|---|---|---|---|
| pptx | `src/collab/room.ts` closeConn "마지막 참여자 퇴장" 분기 | xlsx 갈래가 바로 닫는 코드 안에 `cfTallyRelease(room.cfRanges)`를 넣었다. pptx 갈래는 같은 코드를 `teardownRoom()`으로 옮기고 그 자리를 `lingerRoom ? 저장만 : teardownRoom`으로 바꿨다 | pptx 쪽 분기를 쓰고, `cfTallyRelease`는 `teardownRoom` 안(`largeRoomDocIds.delete` 다음)으로 옮겼다. 그래서 유예가 끝날 때, 유예 상한을 넘어 밀려날 때, `forceCloseRoom`으로 닫을 때 모두 표를 놓는다. 유예 중에는 놓지 않는다(같은 세대로 다시 붙으면 쓴다) |
| pptx | `audit/office/run.mjs` | 두 갈래가 따로 `--no-office`를 넣었다. xlsx는 정답이 없으면 알리고 계속, pptx는 멈춘다. xlsx는 `--only` 형식만 정답을 확인한다 | xlsx의 `NO_OFFICE` 상수와 `--only` 범위 확인을 쓰고, 정답이 없을 때는 pptx·docx처럼 **멈춘다**(조용히 틀린 점수를 내지 않게). pptx의 서버 로그 파일 직접 쓰기(파이프 교착 해결)는 자동으로 합쳐졌다. xlsx의 "캡처 실패면 멈춤"(`capture ${ext} 실패`)도 그대로다 |
| docx | `audit/office/run.mjs` | 같은 `--no-office`의 세 번째 사본(설명·오류 문구만 다름) | 동작이 같아서 HEAD 문구를 두고, docx 쪽 이유 설명을 한 줄 덧붙였다 |
| docx | `src/client/editor.ts` import | pptx `guardRoomEpoch` import와 docx `docxLayout`·`docxPaste`·타입 import | 둘 다 둔다. `guardRoomEpoch(provider, ydoc)` 호출 한 줄은 자동으로 합쳐졌다(639행) |
| docx | `src/routes/viewer.ts` room.js import | xlsx `getCfRemoteAnswer` + `parseCfRemoteRequest` + `CF_REMOTE_RATE_LIMIT`와 docx `docxFragmentsOf` | 이름 목록을 하나로 합치고 xlsx 쪽 두 줄은 그대로 둔다 |

자동으로 합쳐진 공유 파일(`public/css/style.css`, `public/js/viewer.js`, `package.json`, `capture.mjs`, `compare.py`, 점검표, `src/converters/libreOffice.ts`, `vitest.config.ts`, `src/storage.ts`, `public/editor.html`, `src/client/grid-editor.ts`)은 tsc·시험·audit로 확인했다.
`libreOffice.ts`의 전처리 선택은 `ext === '.docx' ? literalizeDocxThemeFonts : ext === '.pptx' ? boldPptxTableHeaders : undefined`이다. docx 쪽은 동기 함수로 가장자리 글상자 처리를 하고, pptx 쪽은 비동기 함수다. 타입이 `Uint8Array | Promise<Uint8Array>`라서 둘이 함께 돈다.

## 상호작용 점검

### 1. 방 세대·닫기 유예·teardownRoom × xlsx 방 추가물

`vitest.config.ts`는 유예를 끄고(`DOCVIEW_ROOM_CLOSE_GRACE_MS=0`) 돌린다. 그래서 두 갈래 시험 어디에서도 이 조합을 보지 않았다. 새 시험은 유예를 400ms로 켜서 본다(`tests/integrationRoomGraceXlsx.test.ts`, 4개).

- **값 개수 표(cfRanges)**
  - 유예 중에도 `getCfRemoteAnswer`가 답한다(재접속 직전에 온 xlsx-cf-ranges).
  - 같은 세대로 다시 붙으면 **같은 표 객체, 같은 rev**를 쓴다. 그래서 편집 화면의 rev 캐시(`sameRev`)가 맞다.
  - 유예가 끝나면 `tally.sheets.size === 0`이고, 방이 없어 경로는 409를 낸다.
  - 지난 세대 탭은 메시지 8로 거절되고, 방을 새로 만들지 않는다.
  - 새로 불러온 탭은 새 방과 새 표로 답을 받는다.
  - **변이 확인**: teardownRoom의 `cfTallyRelease` 줄을 주석 처리하면 이 시험과 다음 시험이 실패한다.
- **강제 닫기**: 유예 중인 방을 `forceCloseRoom`으로 닫으면 표를 바로 놓는다.
- **전역 LRU 상한**: 유예 중인 방은 `lastUsed`가 가장 오래돼 먼저 버려진다. 그 방에 다시 물으면 다시 세어 맞게 답한다(600칸, K1 2개).
- **표 규칙**
  - 유예 뒤 같은 세대로 다시 붙어도 머리 줄만 지우는 행 삭제는 거절된다.
  - 데이터 행을 모두 지우면 한 행이 남는다(r0 0 → r1 2, totalsRows 1).
  - `roomTablesCache`는 방을 키로 하는 WeakMap이라 방과 함께 사라진다.
  - 큰 엑셀 방(`largeRoomDocIds`)은 유예하지 않는다. 메모리를 그대로 지킨다.
- **차트 맵 검사, 리비전**: 둘 다 방(Y.Doc) 수명 안의 관찰자다. 유예 중에도 같은 방이 살아 있어서 바뀌는 것이 없다. 방을 다시 만들면 seedRoom이 다시 건다.
- **실제 브라우저**(`.audit-run/probe/reconnect-cf.mjs`, 8146, Chromium, 1,500행 COUNTIF + 하이퍼링크): **7/7**. 머지 2 뒤와 머지 3 뒤 두 번 돌렸다.
  - 편집 화면이 연 웹소켓을 페이지 안에서 닫았다. 그러자 `epoch=`를 실은 새 연결이 열렸고, 옛 세대 안내는 뜨지 않았다.
  - 그 뒤 다른 참여자가 편집하자 중복이 생기고 없어지는 것이 칠해졌다. 이때 xlsx-cf-ranges 응답은 모두 200이었다. xlsx-extras도 답했다.
- 서버를 다시 켠 뒤 옛 세대 거절(엑셀 새 번들·옛 번들)은 `verify:reconnect`의 엑셀 항목이 본다. 거절 뒤에는 새로 불러오므로, 처음 여는 흐름과 같다.

### 2. 감사 도구(run.mjs 하나)

- `--only`, `--no-office`(정답이 없으면 멈춤), 서버 로그 파일 직접 쓰기, 캡처 실패면 멈춤이 한 파일에 다 있다.
- **증거**: 머지 1 뒤(xlsx 갈래만, run.mjs는 아직 파이프 방식) 전체 audit를 돌리자 pptx 캡처에서 서버가 멈췄다. 로그는 1.6KB에서 멈췄고 서버는 응답하지 않았다. pptx 갈래가 고친 그 파이프 교착이다. 내가 띄운 audit 프로세스만 PID로 끄고 `--only xlsx`로 다시 쟀다(27/27·27/27).
- 머지 2 뒤부터는 전체 audit가 끝까지 돈다.

### 3. package.json 스크립트

`verify:xlsx`, `verify:xlsx:browser`(xlsx), `verify:pptx`, `verify:reconnect`(pptx), `verify:docx`, `perf:docx`(docx), `audit:office`가 모두 있다. 기존 `verify:pdf*`·`verify:view-*`·`verify:legacy-edit`도 그대로다.

### 4. 방 세대·유예 × docx 보존 조각 저장소(docxfrag.json)

`tests/integrationRoomGraceDocx.test.ts`(d04-images.docx, 유예 400ms):

- 그림 문단을 지우고 나가도, 유예 중에는 `docxFragmentsOf`가 **같은 저장소**를 돌려준다. 지운 그림 조각도 그대로 남는다.
- 같은 세대로 다시 붙으면 거절되지 않고 같은 저장소를 쓴다.
- 유예가 끝나 방이 닫히면 `docxfrag.json`에 지운 그림 조각이 남는다. 지난 세대 탭은 거절된다.
- 새로 연 방은 파일의 조각을 모두 이어받는다.
- 워드 편집기의 재접속 안내와 복구 화면은 `verify:reconnect`의 워드 항목이 본다(유예·반쯤 열림·재시작·옛 번들, 모두 통과).

**통합 수정(6142bae)**: docx 갈래의 문서별 모듈 맵 `inheritedFragments`·`writtenFragments`는 방이 닫혀도 지워지지 않았다. writtenFragments에는 문서마다 최대 4MB JSON이 든다. 서버가 켜져 있는 동안 저장한 워드 문서 수만큼 쌓였다.
지금은 teardownRoom의 마지막 단계(저장과 `ydoc.destroy` 뒤)에서 그 문서의 방이 없으면 두 맵을 지운다. 새 방은 `loadSavedDocxFragments`가 파일에서 다시 채우므로 "한 세대만" 동작은 그대로다.
저장보다 먼저 지우면 마지막 저장이 지난 세대 조각을 다시 써 버린다. 그래서 `.finally` 안에서 지운다. 시험은 테스트 전용 훅 `hasDocxFragmentRecordsForTest`로 이를 확인한다.

### 5. editor.ts의 roomEpoch 한 줄 × docx 변경

import 충돌만 났고 둘 다 남겼다. `guardRoomEpoch(provider, ydoc)`는 WebsocketProvider를 만든 바로 다음 줄에 그대로 있다. `verify:docx` 32/32, `verify:reconnect` 워드 항목 모두 통과.

### 6. docx 경로의 접근 확인

`/api/view/slide-media`(word/media)와 `/api/view/docx-drawing`은 각자 `resolveEditorDataView(token)`를 부르고 문서 종류를 확인한다. xlsx 쪽 `/api/view/xlsx-cf-ranges`(편집 토큰, 분당 240회)와 `/api/view/xlsx-extras`는 따로 등록된 경로라 서로 바꾸는 것이 없다. 시험 `cfRemoteRoute.test.ts`(xlsx)와 docx 경로 시험은 전체 vitest에서 통과했다.

## 검증 결과

| 시점 | tsc | vitest | verify:xlsx | verify:xlsx:browser | verify:pptx | verify:docx | verify:reconnect | 브라우저 재접속×cf |
|---|---|---|---|---|---|---|---|---|
| 머지 1 뒤(xlsx) | 통과 | 1036/1037(122파일). 실패 1개는 부하 탓 — `cfRangeValues.test.ts` C1 메모리·멈춤 시험이고, 그 파일만 다시 돌리면 10/10 | 34/34 | 39/39 | 9/9(기준) | — | — | — |
| 머지 2 뒤(+pptx) | 통과 | **1093/1093**(126파일) | 34/34 | 39/39 | 24/24 | — | 39/39 | 7/7 |
| 머지 3 뒤(+docx, 최종) | 통과 | **1158/1158**(130파일) | 34/34 | 39/39 | 24/24 | **32/32** | **39/39** | 7/7 |

docx 왕복(`npx tsx audit/office/roundtrip-docx.ts`, 말뭉치 10 + 검토 재현 10)은 최종에서 **20/20**이다.

### audit:office 점수(`--no-office`, 8147)

| 시점 | docx 보기 | docx 편집 점검표 | xlsx 보기/편집 | pptx 편집 점검표 | pptx 편집 모양 |
|---|---|---|---|---|---|
| 머지 1 뒤 | (전체 실행은 파이프 교착으로 멈춤, 위 2절) | — | 27/27 · 27/27(`--only xlsx`) | — | — |
| 머지 2 뒤 | 8/8, 모양 0.807 | 0/11(예상대로) | 27/27 · 27/27 | **52/52** | **0.975** |
| 최종 | **8/8**, 모양 0.816 | **11/11** | **27/27 · 27/27** | **52/52** | **0.975** |

최종 점검표 실패 항목은 0개이고, 30개 문서 모두 캡처됐다. docx 보기 모양 0.816과 편집 점검표 11/11은 FIX-DOCX.md 4차 수치와 같다.

## 이상한 점·남는 위험

- **`--no-office`에서 정답이 없을 때 이제 멈춘다.** xlsx 갈래 혼자일 때는 알리고 계속했다. 정답이 다 있는 지금은 차이가 없다.
- **유예 중인 방은 값 개수 표 메모리를 최대 45초 더 쥔다.** 전역 상한이 그 방부터 버리고(시험으로 확인), 큰 엑셀 방은 유예하지 않으며, 유예하는 방은 한꺼번에 20개까지라 위험은 작다고 봤다.
- **편집 화면의 cf rev 캐시와 방 재생성**: 방이 새로 만들어지면 rev가 0부터 다시 오른다. 그래서 이미 "새로 불러오기" 안내를 받은 옛 탭이 계속 물으면, 우연히 rev가 같아 범위 개수를 섞을 수 있다. 그 탭은 편집이 막힌 상태라 표시만 틀린다. 손대지 않았다.
- **docx 조각 저장 경합(기존부터 있던 것)**: 닫히는 방의 마지막 저장과, 같은 문서로 막 새로 만든 방의 `loadSavedDocxFragments`가 겹치면 문서별 맵을 서로 덮을 수 있다. 세대 거절 덕분에 새 방은 새로 불러온 탭에서만 생긴다. 이번 통합으로 생긴 문제가 아니라서 고치지 않았다.
- **세대 알림(메시지 7)과 Node 쪽 참여자**: `guardRoomEpoch`가 없는 y-websocket 참여자(시험 스크립트의 Node 참여자 등)는 메시지 7을 받으면 콘솔에 "Unable to compute message"를 남긴다. 해는 없다. 옛 번들도 같고, pptx 갈래 보고서가 다룬 범위다.
- 머지 1 뒤 vitest의 실패 1개는 부하 탓이었다(단독 10/10). 머지 2·3 뒤에는 실패가 없다.

## 최종 검토 반영 (FINAL-REVIEW.md, @9ffd696 → 926660b)

각 항목은 재현 시험을 먼저 써서 실패(RED)를 확인한 뒤 고쳤다. 고친 커밋마다 시험이 하나씩 붙어 있다.

| 항목 | 커밋 | 고친 것 | 고치기 전(RED) → 뒤 |
|---|---|---|---|
| C1 COUNTIF 와일드카드 정규식 되돌아가기 | d573405 | 새 `src/collab/wildcard.ts`가 `*`로 나눈 조각을 차례로 찾는다(칸 길이 × 조각 길이). 엑셀 `~*` `~?` `~~`를 처리하고, 대소문자 무시도 그대로다. 조건부 서식 `countIf`가 이 매처를 쓴다. 같은 패턴인 fast-formula-parser의 `WildCard.toRegex`(COUNTIF·SUMIF·AVERAGEIF·MATCH·VLOOKUP·HLOOKUP·SEARCH)는 `formulaOverrides.ts`에서 이 매처로 덮어쓴다. 서버 수식 엔진, 조건부 서식 대체 경로, 브라우저 chunk가 모두 이 모듈을 거친다(빌드한 chunk에 `patchFormulaWildcards` 들어 있는 것 확인). 그 밖의 `new RegExp`는 상수나 이스케이프한 값만 넣는다(검토 원문과 같음) | `*` 8개 × 200자 칸에서 시험 프로세스가 끝나지 않음 → 2~11ms (`tests/finalReviewC1Wildcard.test.ts` 16개) |
| C2 docxViewLayoutFixes 제곱 정규식 | ec19099 | sectPr는 `indexOf`·`lastIndexOf`로 찾는다. 예전 정규식과 **같은 범위**(첫 sectPr부터 마지막 sectPr 뒤의 첫 닫는 태그까지)다. anchor와 개체 안의 positionH도 indexOf로 차례로 찾는다 | 닫힌 sectPr 400개 + 1MB 꼬리 41초 → 100ms 안. 안 닫힌 anchor 1MB 5초 → 100ms 안. 말뭉치 10 + 검토 문서 10의 결과는 고치기 전 sha256과 바이트 그대로 (`tests/finalReviewC2DocxLayout.test.ts` 22개) |
| I1 /api/view/info 전체 풀기 | b56a4cd | `docxRender`는 워드 편집기가 `editor=1`을 붙였을 때만 만든다. word/ 아래 `.xml`·`.rels`만 푼다. 원본 sha1 판마다 한 번만 만든다(`SizedLruCache` 32MB·64개). 읽기 전용 info는 예전 비용(쪽 설정·기본 글꼴·간격)으로 돌아간다 | 읽기 전용도 만들고 매번 다시 풀던 것 → 만들지 않음·캐시 적중·원본이 바뀌면 다시 만듦 (`tests/finalReviewI1DocxInfo.test.ts` 3개) |
| I2 유예가 개수로만 묶임 | c53a7ad | 방 크기 어림 = PDF 원본 + 암호화 원본 + 엑셀 원본 + Y.Doc 인코딩 길이. 방 하나가 **16MB**를 넘으면 유예하지 않는다(예전처럼 바로 닫음). 유예 중인 방들의 합이 **64MB**를 넘으면 가장 오래된 방부터 닫는다. 개수 상한 20은 그대로다 | 큰 방도 유예, 합 제한 없음 → 바로 닫힘, 오래된 방부터 닫힘 (`tests/finalReviewI2LingerBytes.test.ts` 2개) |
| M1 teardown 정리가 진행 중 저장보다 먼저 | 437110f | 편집이 없어도 teardown이 `runSerialized(docId, 빈 작업)`에 선다. 그래서 진행 중인 저장 뒤에 `ydoc.destroy`와 조각 기록 삭제가 돈다 | 순서 ['destroyed','saved'] → ['saved','destroyed'], 조각 기록이 다시 생기지 않음 |
| M2 닫힌 방으로 cf 답 | 437110f | 기다린 뒤와 센 뒤에 `rooms.get(docId) === room`을 다시 본다. 아니면 null을 돌려준다(경로는 409). 세는 중 저장소가 닫혀 던진 오류도 409로 바꾼다 | 닫힌 방의 답을 냄 → null (`tests/finalReviewM1M2Teardown.test.ts` 2개) |
| M3 각주 복제 | 926660b | 각주 파트를 indexOf로 한 번 훑어 색인(Map)을 만든다: id → 각주, 복제 표지 → id, 가장 큰 id. 복제할 때 색인을 이어서 고친다. 쓰지 않는 복제본 지우기도 같은 순회로 한다. 번호 정의의 `maxOf`도 루프로 바꿨다 | 각주 20만 개에서 RangeError → 저장됨, 복제 id = 가장 큰 id + 1 (`tests/finalReviewM3NoteClone.test.ts`). 왕복 20/20 그대로 |

**I2 상한의 근거**

- 검토가 든 메모리 저장소 배포의 예산은 4GB다. Y.Doc은 메모리에서 인코딩 길이의 몇 배를 차지한다(어림 3~5배). 그래서 합 64MB 인코딩은 실제로 약 200~300MB(예산의 5~8%)이다.
- 방 하나 16MB면 보통 문서·덱(인코딩 수 MB 이하)은 유예된다. 수십~수백 MB PDF와 큰 덱은 예전처럼 바로 닫힌다.
- 큰 엑셀(`largeRoomDocIds`)은 원래 유예하지 않는다.
- 중간 크기 엑셀은 원본 바이트를 어림에 넣는다. lazyGrid 디스크 캐시와 cf 값 개수 표는 어림에 넣지 않았다. 값 개수 표는 전역 LRU 상한이 따로 묶는다.
- 시험할 때는 `DOCVIEW_ROOM_LINGER_MAX_BYTES`·`DOCVIEW_ROOM_LINGER_TOTAL_BYTES`로 바꾼다.

**최종 검토 반영 뒤 검증**

| 항목 | 결과 |
|---|---|
| tsc | 통과 |
| vitest | **1204/1204**(136파일) |
| verify:xlsx | 34/34 |
| verify:xlsx:browser | 39/39 |
| verify:pptx | 24/24 |
| verify:docx | 32/32 |
| verify:reconnect | 39/39 |
| docx 왕복 | 20/20 |

시험 서버는 8146(내가 띄움)과 8147(verify:reconnect가 띄움)이었고, PID로 껐다. audit:office는 이번에 다시 돌리지 않았다. C2의 보기 변환 결과는 말뭉치에서 바이트가 같고, 편집기 화면은 verify:docx로 확인했다.

**남는 것**

- C1에서 fast-formula-parser 라이브러리의 COUNTIF처럼 원래 대소문자를 구분하던 경로(`Criteria.parse`가 'i' 없이 부름)는 그대로 구분한다. 동작을 바꾸지 않으려고 손대지 않았다. 엑셀은 구분하지 않는다.
- I2의 어림에는 Y.Doc 인코딩이 든다. 방을 닫을 때 한 번 문서 크기만큼 일한다. 저장도 같은 크기만큼 일하므로 감수했다.

## 재검토 반영 (FINAL-REVIEW.md "재검토" 절, fe23856 → af2856d)

| 항목 | 커밋 | 고친 것 | 고치기 전(RED) → 뒤 |
|---|---|---|---|
| C1 남은 것 — 총 작업량 | badb0f1 | `wildcard.ts`가 맞추기마다 실제로 한 일을 예산에서 뺀다. 가운데 조각은 칸 길이 × 조각 길이, SEARCH 찾기는 칸 길이 × 조건 길이다. 예산이 모자라면 `WildcardBudgetError`를 던진다. 예산은 **5·10⁷**(`WILDCARD_WORK_BUDGET`)이고 다음 범위마다 하나씩이다: **xlsx-cf-ranges 요청 하나**(넘은 범위는 `skipped`, 답에 `overBudget: true` → 편집 화면 "일부 생략"), **업로드·재생성 굽기 하나**(그 규칙은 칠하지 않고 `cfSkipped`), **업로드 수식 계산 전체 하나**(넘은 칸은 계산 실패 → 수식 원문). 범위 밖에서는 매처 하나(엑셀 함수 호출 하나)마다 예산 하나다. 편집 방의 비동기 재계산과 브라우저가 여기에 해당한다 | 조건 2,000개 × 32k자 칸 200개 요청 13초 → 1초 안 + `overBudget`. 수식 엔진 COUNTIF 한 번 3.7초 → 1초 안. 업로드 수식 100개 5초 안. 굽기 `cfSkipped`. 정상 요청(`a*1`·`*99`)은 결과 그대로 (`tests/finalReviewC1Budget.test.ts`) |
| I2 틈 | da3eb82 | `room.seeded`(시딩이 잘 끝남)가 아니면 유예하지 않는다. 채우는 중인 방은 예전처럼 바로 닫는다. `pdfParts`도 크기 어림에 넣었다 | 원본 읽기가 느린 PDF를 채우는 중 나가면 유예됨 → 바로 닫힘 |
| M3 남은 것 | 9eb376e | `ensureBuiltinLists`의 기본 번호 정의 찾기(`abstractNumIdNamed`)를 indexOf 순회로 바꿨다. 닫는 태그와 이름 자리는 앞으로만 찾는다 | 안 닫힌 `<w:abstractNum>` 1만 개 4.7초 → 300ms 안. 두 번 저장해도 numId·정의 수 그대로 (`tests/finalReviewM3Numbering.test.ts`) |
| 4. teardown 순서(기준부터 있던 것) | af2856d | `persistChangeLog`·`closeLazyGridStore`를 M1과 같은 줄(runSerialized) 뒤, `ydoc.destroy` 직전으로 옮겼다 | 저장 도중 닫으면 저장이 읽는 중에 행 저장소 파일이 이미 지워져 있었다 → 저장 중에는 남아 있고, 저장본에 편집이 들어가며, 끝나면 지워진다 |

**예산 5·10⁷의 근거**: 이 PC에서 와일드카드 맞추기는 약 1.4ms에 10⁶ 단위를 처리한다. 그래서 예산 하나를 다 써도 약 70ms다. 정상 문서는 크게 밑돈다. 예를 들어 10만 칸 × 평균 20자 × `abc*`는 사실상 앞 조각 비교뿐이다.

**바뀐 계약**: xlsx-cf-ranges 답에 `overBudget?: true`가 생겼다. 편집 화면은 이 필드를 몰라도 그 범위가 `skipped`라 "일부 생략"을 보인다.

**검증(af2856d)**

| 항목 | 결과 |
|---|---|
| tsc | 통과 |
| vitest | 1210/1211(138파일). 실패 1개는 부하 때 멈춤 시간 시험(`cfRangeValues.test.ts` C1, 479ms > 100ms)이고, 그 파일만 돌리면 10/10 |
| verify:xlsx | 34/34 |
| verify:xlsx:browser | 39/39 |
| verify:reconnect | 39/39 |
| docx 왕복 | 20/20 |

시험 서버는 8146(내가 띄움)과 8147(verify:reconnect)이었고, PID로 껐다. 다른 작업의 8148(PID 33544)은 건드리지 않았다. `npm run build:editor`가 이 폴더의 `public/js` 번들을 다시 만들었으므로, 8148 서버도 다음 요청부터 새 번들을 내보낸다.

**남는 것**

- 편집 방의 비동기 수식 재계산은 함수 호출 하나마다 예산 하나다. 그래서 거의 맞는 COUNTIF가 든 수식 칸이 많으면, 칸마다 최대 약 70ms씩 막는다. 칸 사이에는 디스크 읽기에서 양보한다. 페이지 단위로 예산을 묶으려면 비동기 경계를 넘어 예산을 넘겨야 해서 손대지 않았다.
- 예산을 넘은 COUNTIF 수식은 오류값이나 계산 실패로 남는다. 틀린 개수를 내지는 않는다.

## 재검토 2 반영 (f612d0b → 6557ce6)

| 항목 | 고친 것 | 고치기 전(RED) → 뒤 |
|---|---|---|
| 1) 값싼 조건 | 매처가 칸을 바꾸기 **전에** 최소 2 × (칸 길이 + 1)을 뺀다. 2는 countIf의 slice와 소문자 바꾸기, 두 번 훑기다. 가운데 조각이 있으면 (칸 길이 + 1) × (가운데 길이 + 2) | `q123*z`·`~q` 2,000개 × 32k자 칸 200개 요청 9.2초 → 1초 안 |
| 2) 편집 방 재계산 | `recomputeFormulasForPage`는 쪽 하나에 예산 하나(`WILDCARD_WORK_BUDGET`)를 쓰고, 칸 하나는 그중 단계 하나(`runWildcardStep`, `WILDCARD_STEP_BUDGET` 5e7 ≈ 114ms)까지 쓴다. 예산 객체는 `computeGridCellDisplayAsync`의 마지막 인자로 넘기고, 시도마다 동기 계산 구간에서만 `withWildcardBudget`으로 쓴다. 비동기 경계에서는 전역 변수를 쓰지 않는다. 125행 묶음을 `Promise.all`로 돌리던 것을 칸마다 차례로 돌리고 20ms마다 양보하게 바꿨다. cf-ranges 요청은 조건마다 단계 하나씩 쓰고 20ms마다 양보한다. 굽기는 칸마다 단계 하나를 쓰고 시간 예산도 확인한다 | 캐시 없는 `COUNTIF($A$1:$A$200,"*"&B1&"*")` 무거운 열 500행: 최대 지연 15.8초·전체 49초 → **116ms·약 2초**. 예산을 넘은 칸은 `#ERROR!` 대신 수식 원문이다 |
| 3) 정상 파일 회귀 | 예산 기본값을 1e9로 올렸다(근거는 `wildcard.ts` 주석: 5e7 ≈ 114ms → 1e9 ≈ 2.3초, 1,000행 정상 파일 전체 약 1.5e8). `formulas.ts`의 두 catch가 예산 오류(FormulaError.details로 감싼 것 포함)를 다시 던진다. 업로드에서는 null(수식 원문)로, 편집 방에서는 계산값 자리에 수식 원문으로 남는다. `spreadsheet.ts`의 주석도 바로잡았다 | 1,000행 `COUNTIF($A$1:$A$1000,"*"&B1&"*")`: 808칸만 숫자 → **1,000칸 모두 숫자** |
| Minor(그대로 둠) | 요청 중간에 예산이 다 떨어지면 뒤 범위가 모두 skipped가 된다. 편집 화면 문구는 `cf-skip-notice` "조건부 서식 일부를 생략했습니다 — …"로 "일부 생략" 뜻이 맞다 | — |

**운영(master)에도 있던 문제인가**: 있었다. master는 같은 `recomputeFormulasForPage` 묶음 `Promise.all`(room.ts 1165행)과 라이브러리 정규식 와일드카드(`WildCard.toRegex`, 지수 되돌아가기)를 쓴다. 그래서 같은 입력에서 더 나빴다(쪽 하나가 끝나지 않는다).

**남는 것**: cf-ranges 요청 하나는 예산을 다 쓰면 약 2초를 쓴다. 20ms마다 양보하지만, 분당 240회 제한 안에서 반복하면 서버 CPU를 계속 쓸 수 있다. 요청 제한을 낮추거나 문서별 동시 1개(이미 방마다 세기 하나)로 묶는 것은 따로 정할 일이다.

**검증(6557ce6 + 시험 고정)**

| 항목 | 결과 |
|---|---|
| tsc | 통과 |
| vitest | 1212/1217(139파일). 실패 5개는 모두 부하 때 시간 시험이다(cfRangeValues C1 멈춤, finalReviewC1Budget·C1Budget2 시간, finalReviewI1 info 비교). 해당 4개 파일만 순서대로 돌리면 19/19. I1 시험은 docxRender만 비교하도록 좁혔다 |
| verify:xlsx | 34/34 |
| verify:xlsx:browser | 39/39 |
| verify:reconnect | 39/39 |

시험 서버는 8146·8147이었고 PID로 껐다. 8148(다른 작업)은 건드리지 않았다.

**cf-ranges 남용 막기(재검토 2 판정)**: 문서마다 답 만들기를 동시에 하나만 한다. 하나는 뒤에서 기다리고, 그 뒤로 오는 것은 409다. 토큰마다 1분 와일드카드 작업량은 5e8(단계 예산의 10배)이고, 넘으면 바로 `overBudget`으로 답한다. 요청 수는 토큰마다 분당 40회로 낮췄다(전에는 IP당 240회). 근거는 verify:xlsx:browser에서 1분에 부른 최대 18회의 2배 남짓이다. 다시 돌린 39항목에서 429는 0회였다. 시험은 `tests/cfRemoteAbuse.test.ts`(36회 통과·41번째 429·다른 토큰 따로)이고, `cfRangeValues`와 브라우저 동시 4요청 항목은 "2개 답 + 2개 409"로 바꿨다.

## 재검토 3 반영 (d458f85 → b4c95be)

| 항목 | 고친 것 | 고치기 전(RED) → 뒤 |
|---|---|---|
| 1) 409·429·예산 초과가 영구 생략 | **편집 화면**(`grid-editor.ts`)에서 세 경우를 모두 "잠시 뒤 다시"로 바꿨다. 범위에 `waitUntil`을 두고, 서버가 알려 준 시간(`retryAfterMs`: 409는 1초, 예산 초과는 토큰 1분 창의 남은 시간, 429는 `Retry-After` 헤더)과 늘려 가는 간격(1초 → 최대 1분, 답을 받으면 0) 중 긴 쪽 뒤에 다시 묻는다. 그사이에는 지난 답으로 칠한다. 보낸 조건은 되돌려 두어 잃지 않는다. 진짜 생략(규칙 밖·상한)만 계속 생략이다. 조건 개수를 통째로 바꿀지는 범위마다 받은 판(`st.rev`)으로 정한다. 예산 초과로 빠진 답의 판을 따라가 새 답을 낡은 답에 합치던 일이 있었다(시험에서 발견). **서버**: 범위마다 센 조건의 답을 기억한다(`RectTally.critCache`, 범위당 4,096개, 표가 바뀌거나 그 범위 칸이 편집되면 버림, 방 teardown 때 표와 함께 사라짐). 다시 물어도 예산을 쓰지 않는다. 범위가 예산 때문에 생략돼도 그 전에 센 조건은 남아, 다음 물음이 이어서 센다. 답의 판은 세기 시작할 때 것을 쓴다. 예산 때문에 생략한 범위에는 `retry: true`를 붙인다 | 브라우저(`verify:xlsx:browser` 6번, 3항목): 옛 편집 화면 번들로 **3/3 실패**(새로 보이는 행이 칠해지지 않음 / 탭 3개 중 둘이 새 결과로 안 바뀜 / 한 토큰 3명 중 일부가 안 칠해짐) → 새 번들 결과는 아래 |
| 2) 시간 단언이 병렬 부하에서 흔들림 | 벽시계 기준을 넓히고, 이유를 주석에 적었다: 결과(overBudget·cfSkipped·0이 아님)로 먼저 보고, 시간은 고치기 전 값과 크게 떨어지는 선으로. `finalReviewC1Budget`·`C1Budget2`·`cfRemoteAbuse`·`C1Wildcard`·`C2DocxLayout`·`M3Numbering`과 `cfRangeValues`의 멈춤 시험 두 개 | — |
| 3) 기다린 요청이 같은 토큰 예산을 따로 받음 | 경로가 토큰 예산을 빌려 준다(`CfBudgetLease`). 답을 만들기 시작할 때 `begin`, 끝날 때(슬롯을 놓기 전) `end`로 쓴 만큼 적는다. 그래서 기다리던 요청은 앞 요청이 쓴 뒤의 남은 양을 받는다 | 시험: 앞 요청이 다 쓰면 기다린 요청은 0을 받아 overBudget |

**브라우저 6번(새 번들)**

| 항목 | 결과 |
|---|---|
| 예산 초과 | 새로 보이는 행 201~212가 약 123초 뒤 칠해졌다. 예산 초과 답 3회였고, 토큰 1분 창이 끝나기를 기다렸다 |
| 탭 3개 | 409를 받고도 약 57초 뒤 세 화면 모두 새 결과로 칠했다 |
| 한 토큰 3명 | 429를 3회 받고도 약 38초 뒤 칠했다 |

**검증(b4c95be)**

| 항목 | 결과 |
|---|---|
| tsc | 통과 |
| vitest 전체 | 두 번 모두 **1224/1224**(140파일). 시간 시험과 유예 시험을 고친 뒤 실패가 없다 |
| verify:xlsx | 34/34 |
| verify:xlsx:browser | **42/42**(6번 3항목 포함) |
| verify:reconnect | 39/39 |

시험 서버는 8146·8147이었고 PID로 껐다.

**남는 것**

- 아주 무거운 시트(칸 1,000개 × 수천 자)에서는 토큰의 1분 예산(5e8)이 화면에 보이는 조건 수십 개를 다 세지 못한다. 그래서 다 칠할 때까지 1~2분 걸릴 수 있다. 그동안 지난 답으로 칠하고, 결국은 다 칠한다.
- 범위 칸을 편집하면 그 범위에서 센 조건을 모두 버린다. 편집된 칸만큼 −1/+1로 이어서 고치는 것은 하지 않았다.
