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

## 6. 반박 검토(REVIEW-XLSX.md) 대응 — 2차

검토 대상은 1bdd19d. 아래 커밋은 6c260bc..(이 보고 커밋)이다. 항목마다 회귀 시험을 넣었다.

| 항목 | 결과 | 무엇을 · 증거(시험) |
|---|---|---|
| C1 SSF가 못 읽는 서식 → 일련번호(`yyyy.mm.dd` 등) | **고침** | 서식 코드를 SSF에 넘기기 전에 고친다(`numberFormat.ts normalizeForSsf`): 따옴표 없는 한글 등 비ASCII를 `"…"`로, 날짜 구역의 점(`.`·`\.`)을 `"."`로(초 소수 `ss.0`은 둔다), 로캘·숫자 표기 지정(`[$-412]`·`[$-ko-KR]`·`[DBNum1~4]`)은 지운다. 그래도 실패하면 날짜는 날짜 글자, 숫자는 쉼표(자리 표시가 없으면 일반). **ExcelJS가 모르는 한국어 엑셀 기본 서식 번호**(통화 5~8, 회계 41~44, 한글 날짜 27~36·50~58)도 ExcelJS 표를 한국어 코드로 채워 해결(예전엔 일반 서식·일련번호). `tests/numberFormatCodes.test.ts` — 흔한 서식 52가지 + 기본 서식 0~58 전부 |
| C2 조건부 서식 굽기 제곱 비용·멈춤 | **고침** | 규칙을 행마다 색인(`cfRowIndex`), 행 묶음마다 이벤트 루프 양보, 일의 양 상한 2천만(넘으면 나머지는 굽지 않고 경고). 편집 화면도 같은 색인. 측정(행마다 규칙 하나, `parseViewerGrid`): 1천 166→133ms, 4천 1,252→414ms, 1만2천 **15,760→1,128ms**, 3만 2,709ms. 가장 긴 멈춤 1만2천에서 0.2초, 3만에서 0.46초(시트 XML 동기 파싱 몫). `tests/condFormatReview.test.ts`(1만2천 규칙, 굽는 동안 타이머가 도는지) |
| I1 `[$-412]AM/PM`이 24시간제 | **고침** | AM/PM은 그대로 SSF가 12시간제로 그리고 결과의 AM/PM만 오전/오후로. `오후 2:24:00`, `오후 06:00` 시험 |
| I2 날짜 칸·TODAY() 비교, 보기 캐시 굳음 | **고침** | 조건부 서식의 값 비교에서 날짜 글자를 일련번호로(`cfNumber`), 피연산자가 수식이면 계산. TODAY()/NOW() 규칙이 있으면 그리드에 구운 날(`cfDay`)을 적고, 보기 화면이 날짜가 바뀐 그리드(디스크·메모리 캐시 둘 다)를 다시 만든다(`routes/viewer.ts gridIsStale`) |
| I3 셀 값 규칙의 상대 참조·식 | **고침** | 피연산자를 범위 왼쪽 위 기준으로 칸마다 옮겨 계산(수식 규칙과 같은 계산기), `=$B$1*2` 같은 식도. 참조 옮기기는 공용 `shiftFormulaRefs`(정렬의 `shiftFormulaRows`도 이것을 쓴다) |
| I4 stopIfTrue·cfvo 종류 | **고침** | 높은 우선순위부터 아직 정해지지 않은 속성만 채우고 "true이면 중지"에서 멈춤. 색조·막대·아이콘 문턱값 min/max/num/percent/percentile/formula(x14 autoMin/autoMax), 막대 길이 minLength/maxLength(기본 10~90) |
| I5 CSS 주입(큰 파일 경로) | **고침** | `sanitizeCellStyle`: 색은 `#RRGGBB`, 선 종류·무늬·정렬은 알려진 이름, 회전·들여쓰기·소수 자릿수는 범위 안 정수, 글꼴 이름은 CSS 특수 글자 제거, 모르는 칸은 버림. 서버 진입점(`fromCellStyleAttrs` — 편집자가 보낸 서식)과 보기 화면 서식 목록(`/api/view/grid/meta`) 두 곳. 그 밖에 CSS/HTML로 가는 값 점검: 조건부 서식 색(편집자가 만든 규칙 포함)은 구워진 뒤 같은 검사를 지나고, 아이콘 기호는 고정 표, 차트 글자는 `esc` + `<img src=data:>`, 편집 화면은 `style.setProperty`라 규칙 밖으로 못 나간다 |
| I6 회계 채움(`* `) | **고침** | 채움 자리를 표시 글자(FILL_MARK)로 남겨 보기·편집 모두 `₩`는 칸 왼쪽 끝, 숫자는 오른쪽 끝(`.g-fill`). x01 5행이 엑셀처럼 `-₩      25,000` |
| M1 `$` 새어 나옴 | **고침** | C1과 같이 `[$-…]`·`[DBNum]` 제거. `[DBNum]`의 한자·한글 숫자 표기 자체는 하지 않는다(엑셀 전용 표기 — 숫자는 보통 숫자로) |
| M2 `[ColorN]` | **고침** | 56색 표(`EXCEL_INDEXED_COLORS`, 칸 색 indexed와 공용) |
| M3 자정 직전 | **고침** | 날짜·시각 서식은 초 단위로 반올림해 그림 |
| M4 1900-03-01 이전 날짜 | **안 고침** | 파일에서 읽은 값은 ExcelJS와 같은 기준이라 맞고, 편집기에 그 이전 날짜를 직접 칠 때만 하루 밀린다. 고치려면 ExcelJS 변환과 우리 변환을 함께 바꿔야 해 위험이 크고 쓰임이 거의 없다 |
| M5 보기 화면 수식 입력줄·찾기 | **고침** | 입력줄은 값(`n ?? v`), 찾기는 보이는 글자와 값 둘 다 |
| M6 툴바 테두리가 bd 통째 삭제 | **고침** | 다시 그린 변만 버린다(`bordersWithout`) |
| M7 nf 칸의 0 표시 | **바뀐 동작(의도)** | 예전 편집 화면은 툴바 서식(nf) 칸의 0을 빈칸으로 보였다. 이제 엑셀처럼 `0`(서식이 회계면 `-`). |
| 10 범위·함수 수식 규칙 | **고침** | 작은 계산기가 못 하는 것(범위·다른 시트·모르는 함수)은 이 앱의 수식 엔진(fast-formula-parser + formulaOverrides)으로 — 서버는 바로, 편집 화면은 그런 규칙이 있을 때만 엔진 조각(326KB)을 따로 불러온다. `$A:$A`는 엔진이 잘못 읽어 시트 끝 행 범위로 바꿔 넘긴다. 브라우저에서 COUNTIF 중복 강조가 보기·편집 모두 칠해지는 것 확인. 한계: 다른 시트 참조는 빈 칸으로 본다 |
| 11 x14 확장 | **고침** | x14 아이콘 집합(3Stars·3Triangles·5Boxes 등)은 화면 판정에만 쓰고(저장은 파일 확장 부분 그대로), 데이터 막대 확장(길이·자동 최소/최대·음수 색)은 본 규칙에 붙인다. 사용자 지정 아이콘(x14:cfIcon) 조합은 첫 집합 모양으로 그린다 |
| 12 엑셀 표 스타일 | **고침(근사)** | `xlsx/tableStyles.ts`: tableStyleInfo(머리 줄·줄무늬 행/열·요약 행·첫/끝 열)를 기본 표 스타일 60가지의 계열 규칙과 테마 색으로 칸 서식에 얹는다. 보기(통째 읽기)·편집(스트리밍 읽기) 같은 함수, 칸에 직접 준 서식이 이긴다. a3 둘째 시트 표 머리 줄이 엑셀처럼 파랑. 한계: 사용자 정의 표 스타일(styles.xml tableStyles)은 Medium2로 본다, 편집 화면에서는 열 때의 모양이 칸 서식으로 들어가 정렬하면 줄무늬가 행을 따라간다(편집기 '표 서식' 삽입과 같은 방식) |
| 13 차트 세부 모양 | **고침(대부분)** | 계열 색(직접·테마+밝기), 기본 색=문서 테마 강조색, 누적·100% 누적, 부드러운 선·표식, 데이터 레이블(값·백분율·항목·서식), 축 제목, 범례 위치, 값 축 숫자 서식, 눈금선 유무, 보조 축·콤보, 영역·도넛·분산형 — 보기·편집 공용 `chartSvg.ts`. x06이 엑셀과 같은 2007 팔레트·오른쪽 범례·부드러운 선·원형 레이블. 못 한 것: 축 최소/최대 고정값, 3차원 효과, 방사형(꺾은선으로 그림)·거품형(분산형으로), 글꼴 크기·제목 위치 등 서식 세부 |
| 14 큰 파일(스트리밍) 조건부 서식 | **고침** | `serializeGridFromLazyStore`가 규칙 있는 시트를 행 묶음으로 한 번 더 읽어 통계를 내고 같은 판정·상한으로 굽는다(편집 중 바뀐 규칙 반영). 한계: 다른 행을 가리키는 수식 규칙은 지금 읽은 500행 묶음 안의 값만 본다. 이 경로의 서식 코드 색([Red])도 이제 싣는다(1차의 shortcut 제거) |

2차 측정:
- 점검표(감사 17항목 / 확장 27항목): 보기 17/17 · 27/27, 편집 17/17 · 27/27(1차와 같음 — 검토 결함은 점검표 고정 파일에 없던 서식이라 단위 시험으로 잡았다).
- 5만 행×5열 서식 시트 `parseXlsxToGrid`: 약 1.7초(1차 1.8초, 처음 1.0초). a3: 약 40ms.
- 조건부 서식 굽기: 위 C2 표.
- vitest 116파일 974개 통과, tsc 통과, verify:xlsx 34/34, 틀 고정·8pt 행·COUNTIF 강조는 실제 브라우저에서 확인, x01·x06·a3 나란히 그림 확인.

2차에서 더 고친 파일: `src/collab/{numberFormat,cellBoxCss,condFormatShared}.ts`, `src/xlsx/{cfFormulaEngine(새),tableStyles(새),viewerCondFormat,viewerGrid,condFormatParse,chartModel,chartSvg,drawingPassthrough,lazyGridStore,collab}.ts`, `src/spreadsheet.ts`, `src/client/{grid-editor,gridCharts,gridSort,fast-formula-parser.d.ts}.ts`, 함께 쓰는 파일 `src/routes/viewer.ts`(찾기·수식 입력줄 값, 서식 검사, 날짜별 그리드 재생성), `public/js/viewer.js`·`public/css/style.css`(채움 배치 `.g-fill`).
