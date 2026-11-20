# 엑셀(xlsx) 표시 품질 고침 — 작업 보고(FIX-XLSX)

가지 `office-quality-xlsx`(감사 가지 94ab62b에서 갈라짐) · 작업 폴더 `projects\webviewer-q-xlsx`
운영 폴더·운영 포트(8000/8010)는 건드리지 않았다. 시험 서버는 이 폴더에서 8144(전/후 비교용 기준 폴더 `webviewer-q-xlsx-base`는 8154 — 측정 뒤 지움), 전용 데이터·임시·LibreOffice 프로필.
MS Office는 띄우지 않았다 — 정답 PDF는 감사 때 만든 것(`webviewer-quality\docs\quality-audit\out\ref`)을 복사해 썼고, `audit:office`에 `--no-office`를 더했다.

## 1. 점수(점검표, `npm run audit:office -- --only xlsx --no-office`)

| | 고치기 전(94ab62b) | 고친 뒤 |
|---|---|---|
| 감사 때 점검표 17항목 — 보기 / 편집 | 0/17 / 4/17 | **17/17 / 17/17** |
| 확장 점검표 27항목(아래 화면 요소 항목 추가) — 보기 / 편집 | 0/27 / 7/27 | **27/27 / 27/27** |

확장 항목은 글자로는 못 재는 것을 화면 요소 수로 잰다(`checklist.json`의 `viewDom`/`editDom`, `capture.mjs domProbes`):
차트 수, 데이터 막대, 조건부 서식 아이콘, 빨간 테두리·점선/이중선 테두리, 무늬 채우기, 회전 글자, 틀 고정 칸, 8pt 행 높이, 좁은 칸 `###`.
"고치기 전" 값은 같은 측정 도구를 기준 커밋 94ab62b 폴더(`webviewer-q-xlsx-base`, 8154)에 그대로 얹어 잰 것이다.
엑셀은 모양 점수(보기 모양·편집 모양)를 compare.py가 내지 않으므로(정답이 인쇄용 PDF) 판정은 점검표와 나란히 그림(`out/sbs/*-grid.png`)을 눈으로 본 것이다.

## 2. 결함별

| 번호 | 결함 | 결과 | 어떻게 · 증거 |
|---|---|---|---|
| X5 | 오류값 `[object Object]` | **고침** | `cellTexts`가 ExcelJS 오류 객체 `{error}`를 그대로 글자로(엑셀처럼 가운데 정렬). 시험 `numberFormat.test.ts`, x05 I6 `#DIV/0!` 보기·편집 |
| X4 | 공유 수식 칸 빈칸(x08 F10·F11) | **고침** | ExcelJS가 공유 수식 복제 칸을 `{sharedFormula, result}`로 주는데 `'formula' in`만 봤다 → 둘 다 수식으로. 캐시 결과가 빈 글자(`=IF(D13="","",…)`)인 칸도 엑셀처럼 빈칸 |
| X1·X3·X12 | 숫자 서식(시간·날짜·괄호 음수·빨강·지수·분수·앞 0·천 단위·글자 붙은 서식·조건 구역·회계·USD·요일·로캘) | **고침**(회계 채움 간격만 근사) | 손으로 짠 서식기를 없애고 **SheetJS SSF**(Apache-2.0, fast-formula-parser에 이미 들어 있는 사본 — 새 의존성 없음)로. 보기(서버)·편집(브라우저)이 **같은 함수** `src/collab/numberFormat.ts`만 쓴다. SSF가 모르는 것 보탬: 한국어 요일 `aaa/aaaa`, `[$-412]`의 요일·오전/오후, 서식 색(`[Red]` 등, 조건 구역 포함), 좁은 칸 `####`, ExcelJS가 영어 코드로 바꾸는 기본 서식 번호 14·22·37~40 → 한국어 엑셀 코드(x01 B6 `-4,321` 빨강이 엑셀과 같아짐). 그리드 칸에 값 글자 `n`을 따로 둬 서식 모양("(4,321)", "12,346천원")이 수식 계산·편집기 시딩·저장에 섞이지 않는다(예전엔 표시 글자에서 쉼표만 벗겨 값으로 썼고, 반올림된 표시값으로 계산되는 문제도 있었다). 시각은 초까지 보존. 날짜·서식 숫자 칸은 서식이 있으면 오른쪽 정렬(X12). x01 sbs: 보기·편집 모두 엑셀과 같은 글자 |
| X6 | 보기 화면 차트 없음 | **고침**(모양은 편집 화면과 같은 근사) | 편집 화면 그리기를 `src/xlsx/chartSvg.ts`로 옮겨 공유, 서버가 지금 칸 값으로 SVG를 만들어 `/api/view/xlsx-chart-images`로 내려 주면 보기 화면이 그림처럼 얹는다. x06: 3개 차트 보기·편집 모두. 남은 차이: 파일의 차트 색(2007 팔레트)·부드러운 선·원형 데이터 레이블·축 제목은 두 화면 모두 아직 그리지 않는다(우리 차트 모델이 종류·제목·계열만 읽음 — 차트 XML 스타일 해석은 별도 작업) |
| X7 | 조건부 서식: 보기 화면 없음, 편집 화면 아이콘 집합·수식 규칙 없음 | **고침** | `condFormatShared.ts`에 아이콘 집합(문턱 percent/percentile/num, gte, reverse, showValue=0, 20종 기호)과 수식 규칙(상대 참조를 칸마다 옮기는 작은 계산기 — 비교·산술·& ·AND/OR/NOT/IF/ISBLANK/MOD/ROW/LEN 등, 모르는 함수·범위 참조는 "안 맞음") 추가. 보기 화면은 같은 판정을 그리드 파일 만들 때 칸 서식에 굽는다(`xlsx/viewerCondFormat.ts`). x03 sbs: 색조·막대·아이콘·셀 강조·A열 초록 굵게 모두 엑셀과 같음. 한계: 수식 규칙의 범위 참조·다른 시트 참조·SUM 등은 계산하지 않음, 큰 파일(시트 XML 48MB 넘는 스트리밍 경로) 보기 화면에는 굽지 않음(범위 통계에 행을 두 번 읽어야 함 — 코드에 shortcut 주석) |
| X8 | 테두리 색·선 종류 | **고침** | `GridStyle.bd`(변마다 [엑셀 선 종류, 색])와 공용 `collab/cellBoxCss.ts`(점선·파선·이중선·굵은 선·hair·쇄선 → CSS). 저장 때도 선 종류·색 그대로 씀. 툴바로 테두리를 바꾸면 그 칸의 `bd`는 버린다. 한계: 쇄선(dashDot 계열)·slantDashDot은 CSS에 없어 파선으로 |
| X9 | 무늬 채우기·글자 회전·셀에 맞춤·(들여쓰기) | **고침**(무늬는 근사) | 무늬: gray 계열은 섞은 색, 줄·격자·trellis는 반복 그라데이션. 회전 1~180°·세로쓰기(255), 들여쓰기, 셀에 맞춤(글자 크기를 줄여 칸에 맞춤, 넘침 없음). 보기·편집 같은 함수. 들여쓰기 한 단계 폭(9px)은 어림 |
| X10 | 숫자 모양 글자·TEXT 결과가 오른쪽 정렬 | **고침** | 값이 글자인 칸이 숫자처럼 보이거나 숫자 서식이 걸려 있으면 왼쪽 정렬을 칸 서식에 박는다(엑셀 기본과 같은 모양). x01 C열·x05 I5 |
| X11 | 편집 화면이 8pt 행을 키움 | **고침** | 편집 화면도 칸 내용을 행 높이에서 자르고, 기본보다 낮은 행은 칸 위아래 여백·행 번호 글자 크기를 줄인다. x04 3행: 보기 15→11px, 편집 38→11px(엑셀 8pt=11px) |
| (틀 고정) | 보기 화면 틀 고정 없음(감사 땐 정적 화면이라 확인 못 함) | **고침** | 보기 화면도 위 N행·왼쪽 N열을 sticky 칸으로, 열 머리글도 함께. sticky가 transform 위치를 몰라 틀 고정 시트만 표를 `top`으로 옮긴다. 실제 스크롤해서 확인(x04, 세로 900px·가로 150px) |
| (좋음 유지) | 병합·숨긴 행/열·숨긴 시트·줄 바꿈·열 너비·행 높이·그림·여러 시트 탭 | 회귀 없음 | `verify:xlsx` 34/34, x02·x04·x07·x08 sbs |

못 고친 것·근사로 남긴 것(이유):
- 회계 서식의 채움(`* `) — 엑셀은 `₩`를 칸 왼쪽 끝, 숫자를 오른쪽 끝에 둔다. SSF는 채움 문자를 버리므로 `-₩25,000`처럼 붙어 보인다. 칸 안 양 끝 배치는 보기·편집 렌더러 양쪽에 칸 내부 구조를 새로 넣어야 해서 이번엔 하지 않음(글자는 같다).
- 차트 세부 모양(위 X6).
- 무늬 채우기·쇄선은 CSS로 정확히 같은 무늬를 못 낸다(근사).
- 감사 표에 없던 것을 이번에 봤으나 고치지 않음: a3 둘째 시트의 **표 서식(Excel 표) 머리 줄 색·줄무늬**가 보기·편집 모두 없다 — 표 파트(xl/tables)의 스타일 해석이 필요한 별도 기능이라 이번 범위 밖.
- 좁은 칸 `####` 판정은 글자 폭 어림(숫자 0.45em 등, 들어가는 값은 가리지 않게 보수적으로). 일반(General) 서식 숫자는 엑셀도 `####` 대신 줄여 보이므로 대상 아님.

## 3. 성능(큰 시트 회귀 확인)

`parseXlsxToGrid` 측정(같은 PC, 두 번째 실행 시간):

| 파일 | 전 | 후 | 그리드 파일 크기 전→후 |
|---|---|---|---|
| a3-fidelity-grid(1,218행, 시트 3개) | 27ms | 약 35ms | 37,780 → 37,780 B |
| 합성 5만 행×5열(전부 서식 숫자·날짜·빨강 음수·백분율) | 약 1.0초 | 약 1.8초 | 6.29 → 8.46 MB |

느는 몫은 서식기(SSF) 자체와, 값이 표시와 다른 칸에 값 글자 `n`을 싣는 것(정확한 값을 위해 필요). 서식 코드별 준비 결과·날짜 서식 판정·색/정렬 변형 서식을 기억해 처음 2.4초에서 1.8초로 줄였다. 지연 그리드(LazyGridStore, 큰 파일 편집)는 같은 칸 변환 함수(`gridCellFromExcel`)를 쓰고 구조는 그대로다. a3를 보기·편집 화면으로 여는 것은 audit 캡처에서 오류 없이 열렸다(변환 1.2초대, 전과 같은 수준).

## 4. 고친 파일

엑셀 전용: `src/spreadsheet.ts`, `src/xlsx/{collab,condFormatParse,formulas,lazyGridStore,viewerGrid,sheetGeometry}.ts`, 새 `src/xlsx/{chartSvg,viewerCondFormat}.ts`, `src/client/{grid-editor,gridCharts}.ts`, `src/client/fast-formula-parser.d.ts`, `src/collab/{cellValueShared,condFormatShared,gridShared}.ts`, 새 `src/collab/{numberFormat,cellBoxCss}.ts`.

**다른 형식과 함께 쓰는 파일(병합 때 주의)**:
- `public/js/viewer.js` — 그리드(엑셀) 부분만: 서식 글자는 서버 것 그대로, 차트 그림, 조건부 서식·상자 CSS, 틀 고정, 낮은 행.
- `public/css/style.css` — 그리드 규칙만 추가(`.grid-edit-table td[data-cf-icon]`, `.g-frozen`, `.g-head-frozen`, `.g-row-short`).
- `src/routes/viewer.ts` — `/api/view/xlsx-chart-images` 추가, `/api/view/grid/meta` 서식에 CSS(x/xi) 덧붙임.
- `src/collab/room.ts` — 엑셀 서식만 바꿀 때 값 글자 `n` 보존, 변경 기록 값.
- `audit/office/{run.mjs,capture.mjs,compare.py,checklist.json}` — `--no-office`, `--only`일 때 그 형식 정답만 확인, xlsx 화면 요소 점검 항목. (docx·pptx 레인도 이 파일을 고쳤다면 병합 시 합칠 것 — 기존 항목·출력 형식은 그대로다.)
- 그리드 캐시 형식 버전 `GRID_CACHE_FORMAT_VERSION` 4→5: 운영에 반영되면 기존 문서의 보기용 그리드가 다음 조회 때 한 번 다시 만들어진다(기존 자가 치유 경로).

## 5. 검증

- `npx tsc --noEmit` 통과, `npx vitest run` 전체 112파일 891개 통과(시작 때 861개), 새 시험: `tests/numberFormat.test.ts`, `tests/condFormatIconExpr.test.ts`, `tests/cellBoxCss.test.ts`(바뀐 기대: `xlsxCondFormat.test.ts`의 수식 규칙 → `expr`).
- `npm run verify:xlsx`(내 8144 서버) 34/34.
- `audit:office --only xlsx --no-office` 후 x01·x02·x03·x04·x05·x06·x08 나란히 그림을 직접 보고 판정, 틀 고정·8pt 행은 실제 브라우저에서 스크롤·크기 측정.
