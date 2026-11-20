# 오피스 문서(워드·엑셀·PPT) 표시 품질 점검 보고서

작성 2026-11-20 · 브랜치 `office-quality`(master 102fc69에서 갈라짐) · 작업 폴더 `C:\Users\OpenMinds\projects\webviewer-quality`
운영 폴더(`projects\webviewer`)·운영 포트(8000/8010)·운영 LibreOffice는 건드리지 않았다. 시험 서버는 이 작업 폴더에서 8141 포트, 전용 데이터·임시·LibreOffice 프로필(`.audit-run/`)로 따로 띄웠다.

## 1. 결론 먼저 (쉬운 말로)

| 형식 | 보기 화면(LibreOffice가 만든 PDF) | 편집 화면(우리가 직접 그림) — 향기랩은 늘 이 화면 |
|---|---|---|
| 워드(docx) | **좋음.** 10개 문서 모두 쪽 수가 Word와 같다. 표·그림·머리글·바닥글·쪽 번호·가로 쪽·2단·각주가 다 나온다. 눈에 띄는 차이는 한글 줄 바꿈 위치(낱자 단위)와 일부 표 열 너비 정도. | **약함.** 그림·글상자가 **하나도 안 보이고**, 번호 매기기(1. / 1.1. / 가))가 전부 "•"로 바뀌고, 표 스타일 테두리·머리글·바닥글·각주·탭이 사라진다. 본문 글자 자체는 거의 다 있다(글자 일치 0.92). |
| 엑셀(xlsx) | (엑셀은 LibreOffice를 쓰지 않는다 — 보기·편집 모두 우리 격자) **보통 이하.** 숫자 서식이 많이 틀리고(시간이 "1899-12-30 14:05", "(4,321)"이 "-4,321" 등), **차트와 조건부 서식이 보기 화면에 없고**, 공유 수식 칸 일부가 빈칸, 오류값이 "[object Object]". | 보기보다 낫다: 차트·색 눈금·데이터 막대가 나온다. 숫자 서식 문제는 같고, 아이콘 집합·수식 규칙 조건부 서식, 테두리 색·점선, 무늬 채우기, 글자 회전은 빠진다. |
| PPT(pptx) | **좋음.** 10개 묶음 모두 장 수·배치가 PowerPoint와 같다. 차트 글자에 띄어쓰기가 끼는 것("1 월", "추세 ( 꺾은선 )")과 표 머리 줄 굵기 정도. | **보통.** 표·차트는 LibreOffice 그림을 잘라 붙여 같다. 그러나 **자동 맞춤(글자 축소)을 무시해 글이 슬라이드 밖으로 넘치고**, 그림 자르기·회전이 무시되고, 줄 간격·단락 간격·제목 테마 글꼴·자리표시자 글자색이 빠진다. |

- 정답은 **이 PC의 MS Office(Word·Excel·PowerPoint)** 로 만든 PDF다(COM 자동화, 창 숨김, 매크로 끔, 원본은 읽기 전용으로 열고 저장 안 함, 문서마다 따로 띄우고 띄운 프로세스만 번호로 끔).
- 가장 큰 문제는 LibreOffice가 아니라 **우리 편집기(직접 그리는 부분)** 다. 향기랩은 늘 편집 화면이고, 8000 연동도 내려받기 권한이 있으면 편집기 주소(editorUrl)를 받으므로, 사람들이 실제로 보는 것은 대부분 편집 화면이다.
- 엑셀은 원래 LibreOffice를 쓰지 않으므로 엑셀 문제는 전부 우리 코드에서 고칠 수 있다.

## 2. 형식별 처리 흐름

### 워드(docx)
```
올림 → LibreOffice(상주 변환기 loWorker.py, 실패 시 soffice 명령) → PDF
       (src/converters/libreOffice.ts officeConverter + 테마 글꼴 전처리 literalizeDocxThemeFonts, 글꼴 대체 loFontSubstitution.ts)
  ├ 보기 화면 /view.html : 그 PDF를 pdf.js로(내려받기 권한 있을 때) 또는 서버가 그린 PNG로  (src/client/view/*, /api/view/derived-pdf, /api/view/pages/N)
  └ 편집 화면 /editor.html : 원본 docx를 우리가 직접 읽어 ProseMirror로 그림 — LibreOffice와 무관  (src/docx/import.ts → schema.ts → src/client/editor.ts)
```
### 엑셀(xlsx)
```
올림 → (LibreOffice 안 씀) 우리 격자 추출 src/xlsx/viewerGrid.ts / src/spreadsheet.ts(ExcelJS)
  ├ 보기 화면 /viewer.html : public/js/viewer.js 가 격자를 HTML 표로 그림
  ├ 편집 화면 /editor-grid.html : src/client/grid-editor.ts (+ gridCharts.ts 차트, collab/condFormatShared.ts 조건부 서식)
  └ 인쇄용 PDF만 요청 때 LibreOffice(gridPrintConverter)
```
### PPT(pptx)
```
올림 → LibreOffice → PDF
  ├ 보기 화면 /view.html : 그 PDF(워드와 같음)
  └ 편집 화면 /editor-slides.html : 섞어 그림 — 글상자·그림·단순 도형은 우리가 HTML로(src/pptx/slideContent.ts, src/client/editor-slides.ts),
     표·차트·스마트아트·복잡한 도형은 LibreOffice 쪽 그림(/api/view/pages/N)에서 그 영역만 잘라 붙임(renderSnapshot)
```

## 3. 시험 문서와 방법

- 고객 문서 없음. `audit/office/gen_corpus.py`가 만든 합성 한글 업무 문서 24개 + ArtCraft 시험 때 쓴 합성 문서 6개(`a1~a3-*`) = 워드 10·엑셀 10·PPT 10. 원본은 `audit/office/corpus/{docx,xlsx,pptx}`.
- 엑셀·PPT는 MS Office로 한 번 다시 저장한 파일을 썼다(실제 사용자 파일과 같은 모양). 워드는 python-docx가 만든 파일이 Word 호환 모드라 그대로 썼다(호환 모드 문서를 다시 저장하면 Word가 숨은 확인 창에서 멈췄다) — 그래서 표 열 너비 차이(D-v3)는 Word로 저장한 실제 파일에선 덜할 수 있다.
- 화면은 Playwright Chromium(1400×1000, 배율 1)으로 찍었다: 워드·PPT는 보기 화면 + 그 PDF + 편집 화면(슬라이드마다), 엑셀은 시트마다 보기·편집 격자.
- 나란히 그림: `docs/quality-audit/out/sbs/`(git 제외, 다시 만들려면 7장 명령) — `<문서>-view.png`(왼쪽 Office · 오른쪽 LibreOffice), `<문서>-edit.png`(왼쪽 Office · 오른쪽 편집기), `<문서>-grid.png`(엑셀: 왼쪽 Excel · 오른쪽 위 보기/아래 편집). 모든 그림을 직접 눈으로 보고 판정했다.

## 4. 결함 표

심각도: 상 = 내용이 빠지거나 틀려 보임(읽는 사람이 오해), 중 = 모양이 확연히 다름, 하 = 눈썰미 있는 사람만 앎. 난이도: 하 = 하루 안, 중 = 며칠, 상 = 1주 이상.

### 4.1 워드(docx)

| 번호 | 결함 | 문서·위치 | 심각도 | 원인 | 코드 위치 | 난이도 |
|---|---|---|---|---|---|---|
| D1 | 편집 화면에 **그림이 전혀 안 보임**(줄 안 그림·떠 있는 그림·글자 사이 아이콘 모두) | d04-images 전체(그림 3개 → 0), a1-korean-report 1쪽 파란 그림, d08 공문 오른쪽 아래 로고 | 상 | 우리 편집기: `w:drawing`을 건너뜀("v1 미지원") | src/docx/import.ts:222 `inlineContentOf`, 768 `docxHasUnsupportedContent`; schema.ts에 그림 노드 없음 | 중(줄 안 그림) / 상(떠 있는 그림·글 흐름) |
| D2 | 편집 화면 **번호 매기기가 사라짐**: 1. / 1.1. / 1.1.1. / 가) / (ㄱ)이 전부 같은 수준 "•"로, 들여쓰기 수준도 없음. 스타일로 매긴 목록(List Number/List Bullet)은 기호조차 없음 | d02-numbering 1~2쪽 | 상 | 우리 편집기: numbering.xml 해석 안 함, 문단에 직접 붙은 numPr만 목록으로 봄 | src/docx/import.ts:271 `isListItem`, 698 `groupListItems`(주석에 "알려진 단순화") | 중 |
| D3 | 편집 화면에 **머리글·바닥글·쪽 번호가 없음**, 쪽 구분 없이 한 장으로 이어짐, 가로 쪽 구역도 세로 폭 | d05 전체("대외비", "- 2 / 5 -" 없음), a2-korean-layout | 중 | 우리 편집기: header/footer 파트를 읽지 않음, 쪽 나눔 없음 | src/docx/import.ts(머리글 파트 읽기 없음), src/client/editor.ts:181 applyPageSetup | 중(표시만) / 상(쪽 나눔) |
| D4 | 편집 화면에 **각주가 없음**(본문 위첨자 번호도 없음), 2단이 1단으로 | d06 1쪽 | 중 | 우리 편집기: footnotes.xml·w:cols 미해석 | src/docx/import.ts | 중 |
| D5 | 편집 화면에 **글상자가 없음**(안의 글 "글상자: 주의 사항 / 점검 전 전원 차단…"까지 사라짐) | d07 1쪽 | 상 | 우리 편집기: wps 글상자도 w:drawing이라 통째로 건너뜀 | src/docx/import.ts:222 | 중(글만 꺼내기) / 상(제자리 배치) |
| D6 | 편집 화면 **탭 정지·점선 채움이 사라져 글자가 붙음**: "제1장 사업 개요 ……… 1" → "제1장 사업 개요1", "왼쪽가운데오른쪽" | d07 1쪽 목차 4줄·소수점 맞춤 2줄 | 중 | 우리 편집기: w:tab·w:tabs 미해석 | src/docx/import.ts `inlineContentOf` | 중 |
| D7 | 편집 화면 **표 스타일이 사라짐**: "Table Grid" 표의 칸 선이 없어지고, "Light Grid Accent 1"의 테두리·줄무늬·굵은 첫 열이 없는 맨 글자 표로 | d03 1~2쪽 세 표, a1·a2·d08 표 | 상 | 우리 편집기: tblStyle의 tblBorders·조건부 서식(firstRow/band) 미해석(기본 스타일의 칸 여백만 읽음) | src/docx/import.ts:391~410 `tableCellMarginOf` 주변 | 중 |
| D8 | 편집 화면 **표 열 너비가 원본 무시하고 균등**, 칸 위아래 여백이 커서 표가 1.5배 길어짐; 이중선이 한 줄로 | d03 첫 표(구분 열 2cm가 넓어짐) | 중 | 우리 편집기 | src/docx/import.ts `tableCellFrom`(342) | 하~중 |
| D9 | 편집 화면 **제목·부제·인용 스타일 색·글꼴이 빠짐**: 파란 제목1·2가 검정, 부제 기울임·회색 없음, a1 제목 아래 선 없음, 제목 크기 더 큼 | d01, a1 1쪽 | 중 | 우리 편집기: 제목 스타일 캐스케이드 미해석(주석에 명시) | src/docx/import.ts:64 `headingLevel`, 89 `namedStyleInfoOf` | 중 |
| D10 | 편집 화면 **한글이 낱자 단위로 줄 바뀜**("향/기랩", "주식/회사") — Word는 어절 단위 | d01·d04·a1 본문 | 중 | 브라우저 CSS 기본값(word-break: normal) | 편집기 본문 CSS(`.ProseMirror`) | 하 |
| D11 | 편집 화면 줄 간격이 Word보다 촘촘(맑은 고딕 문서에서 Word가 줄 높이를 더 크게 잡음) | a1, d01 | 하 | 글꼴 줄 높이 계산 차이 | src/client/editor.ts 기본 줄 간격 | 중 |
| D-v1 | 보기 화면(LibreOffice) **한글 줄 바뀜이 낱자 단위**, Word는 어절 단위 → 줄 끝 모양이 다르고 긴 문서는 쪽이 밀릴 수 있음 | d01·d04·d06 본문 | 중 | LibreOffice(26.2)의 한글 줄 나눔 | 우리 코드 밖 | LibreOffice 한계 |
| D-v2 | 보기 화면 글상자 옆 글이 **양쪽으로 흘러 오른쪽에 한 글자씩 세로로** 남음("회/풍/년/계/업"), Word는 왼쪽에만 | d07 1쪽 | 중 | LibreOffice 글 흐름(좁은 틈 최소 폭 없음) | 전처리로 wrapText="bothSides"→"largest"(libreOffice.ts 전처리 단계) | 하~중 |
| D-v3 | 보기 화면 표 열 너비가 Word와 다름(구분 열이 넓어지고 긴 칸이 3줄) → 아래 긴 표가 한 줄 일찍 다음 쪽으로 | d03 1쪽 | 하 | LibreOffice는 tblGrid, Word는 칸 너비(tcW)를 따름 — Word로 저장한 실제 파일에선 덜함 | — | — |
| D-v4 | 보기 화면 2단 문서의 각주가 단마다가 아니라 쪽 폭 전체에, 각주 번호가 위첨자 아님 | d06 1쪽 | 하 | LibreOffice | — | LibreOffice 한계 |
| D-v5 | 보기 화면에서 영문 글꼴(Times New Roman·Arial) 칸 안의 한글이 굵은 다른 글꼴로 대체 | d01 1쪽 5·6줄 | 하 | 글꼴 대체 | src/converters/loFontSubstitution.ts | 하 |
| D-v6 | 보기 화면 "2026 년"처럼 Word가 한글과 숫자 사이에 넣는 작은 간격(autoSpaceDN)이 없음 | 모든 문서 | 하 | LibreOffice | — | LibreOffice 한계 |

### 4.2 엑셀(xlsx) — 보기·편집 모두 우리 코드

| 번호 | 결함 | 문서·위치 | 심각도 | 원인 | 코드 위치 | 난이도 |
|---|---|---|---|---|---|---|
| X1 | **숫자 서식이 많이 틀림**(보기·편집 둘 다): 시간 `14:05:09`→`1899-12-30 14:05`, `PM 2:05`→`1899-12-30 14:05`, `(4,321)`→`-4,321`, 빨간 음수 색 없음, `1.23E+08`→`123456789`, `3/4`→`0.75`, `2026년 11월 20일`→`2026-11-20`, 요일 `(금)` 없음, `00042`→`42`, `12,346천원`→`₩12,345,678`, 회계 `-₩ 25,000`→`₩-25,000`, 조건 서식 `많음 1,500` 없음, 좁은 칸 `####` 대신 숫자가 잘림 | x01-number-formats 시트1 B2~B22, E2 | 상 | 우리 코드: 쉼표 여부만 보는 자체 서식기 | src/spreadsheet.ts:358 `formatNumber`, 366 `formatDate`, 648 `classifyNumFmt`; src/collab/cellValueShared.ts; src/client/grid-editor.ts 표시부 | 중(검증된 서식 라이브러리 SSF로 교체) |
| X3 | **글자가 붙은 사용자 서식이 통째로 바뀜**: `일금 9,108,000원정` → `₩9,108,000`(없던 ₩까지) | x08-korean-quotation C6(합계 금액 칸), 보기·편집 | 상 | X1과 같은 서식기 | X1과 같음 | X1에 포함 |
| X4 | 보기 화면 **공유 수식 칸 일부가 빈칸**: 공급가액 F10 `900,000`, F11 `3,840,000`이 안 보임(파일엔 계산값 있음, F9·F12는 보임). 편집 화면은 정상 | x08 F10·F11 | 상(견적서 금액 누락) | 우리 코드: 공유 수식(shared formula) 복제 칸의 값 처리 | src/spreadsheet.ts:395~410 `displayValue`, 544 `applyComputedFormulaValues` | 하~중 |
| X5 | 보기 화면 **오류값이 `[object Object]`**(`#DIV/0!`이어야) | x05-formulas I6 | 중 | 우리 코드: `String(result)`에 ExcelJS 오류 객체 `{error:'#DIV/0!'}`가 들어옴 | src/spreadsheet.ts:410 | 하(한 줄) |
| X6 | 보기 화면에 **차트가 없음**(막대·꺾은선·원형 모두). 편집 화면은 그림 | x06-charts 시트1 F1·F18·A17 | 상 | 보기 화면(viewer.js)에 차트 그리기 없음 | public/js/viewer.js ↔ src/client/gridCharts.ts 재사용 | 중 |
| X7 | 보기 화면에 **조건부 서식이 전혀 없음**(색 눈금·데이터 막대·아이콘·강조). 편집 화면도 **아이콘 집합(신호등·화살표)과 수식 규칙(A열 초록 굵게)은 빠짐** | x03 B·C·D·E열, A2/A10/A12, 19행 | 상 | 보기: 기능 없음 / 편집: iconSet·expression 규칙 미구현 | public/js/viewer.js; src/xlsx/condFormatParse.ts:112~(colorScale·dataBar만), src/collab/condFormatShared.ts | 중 |
| X8 | **테두리 색·선 종류가 사라짐**: 빨간 테두리가 검정, 점선·파선·이중선·쇄선이 실선/없음 | x02 10행 A~H | 중 | 우리 코드: 굵기 1/2만 남김 | src/spreadsheet.ts:627 `borderWeight`, 733~740 | 중 |
| X9 | **무늬 채우기**(darkGrid·lightTrellis·gray125) 없음, **글자 회전**(세로쓰기·45도) 없음, **셀에 맞춤 축소** 없음(글자 잘림) | x02 12행 B~D, 14행 A·B·D | 하~중 | 우리 코드 미구현 | src/spreadsheet.ts 스타일 변환 + viewer.js·grid-editor.ts | 중 |
| X10 | 문자로 저장된 숫자 모양 글자(`0.00`, `00000`)와 TEXT 함수 결과가 오른쪽 정렬(엑셀은 왼쪽) | x01 C3·C8·C9·C18, x05 I5 | 하 | 정렬을 값 모양으로 판단 | public/js/viewer.js, grid-editor.ts | 하 |
| X11 | 편집 화면이 사용자가 정한 낮은 행 높이(8pt)를 내용에 맞춰 키움(보기 화면·엑셀은 8pt) | x04 3행 | 하 | 편집기 행 높이 자동 맞춤 | src/client/grid-editor.ts | 하 |
| X12 | 편집 화면 날짜 칸 정렬이 칸마다 다름(왼쪽/오른쪽 섞임) | x01 B11·B13·B14 | 하 | X1과 같은 서식 판단 | grid-editor.ts | X1에 포함 |
| (좋음) | 병합·숨긴 행/열·숨긴 시트·줄 바꿈·열 너비·행 높이(보기)·그림(로고)·여러 시트 탭은 엑셀과 같다 | x02·x04·x07·x08 | — | — | — | — |

### 4.3 PPT(pptx)

| 번호 | 결함 | 문서·위치 | 심각도 | 원인 | 코드 위치 | 난이도 |
|---|---|---|---|---|---|---|
| P1 | 편집 화면이 **자동 맞춤(normAutofit fontScale=47.5%)을 무시** → 글이 원래 크기로 커져 슬라이드 아래로 넘쳐 잘림(13단락 중 3개만 보임) | p03-autofit 1장 | 상 | 우리 편집기: normAutofit·fontScale·lnSpcReduction 미해석(코드에 해당 단어 없음) | src/pptx/slideContent.ts(글상자 bodyPr), src/client/editor-slides.ts(글상자 그리기) | 하~중 |
| P2 | 편집 화면 **그림 자르기 무시**(잘린 그림 대신 원본 전체를 찌그려 넣음), **회전 무시**(15° 그림·30° 화살표가 똑바로) | p06 1장 가운데·오른쪽 그림, p07 1장 오른쪽 위 화살표 | 중 | 우리 편집기: srcRect·xfrm rot를 그림/도형에 안 씀(잘라 붙이기 영역 계산에만 씀) | src/client/editor-slides.ts:406 `renderSlideVisual`, src/pptx/slideContent.ts:695 | 하 |
| P3 | 편집 화면 **줄 간격·단락 뒤 간격이 빠져** 글이 촘촘(줄 간격 1.3·뒤 6pt 무시), 글머리 단락 사이 간격도 좁음 | p08 1장 왼쪽 글상자, p01 2장 | 중 | 우리 편집기: lnSpc·spcBef·spcAft를 화면에 안 씀 | src/client/editor-slides.ts, src/pptx/textSchema.ts | 하~중 |
| P4 | 편집 화면 **자리표시자가 물려받는 글자색·글꼴이 빠짐**: 부제목 회색이 검정, 제목 테마 글꼴(Georgia/HY견고딕 굵게)이 기본 고딕 | p01 1장, a2 1장 부제, p02 1장 제목 | 중 | 우리 편집기: 레이아웃·마스터 lstStyle·테마 major 글꼴 상속 일부만 | src/pptx/inheritance.ts, src/pptx/theme.ts | 중 |
| P5 | 편집 화면 도형 **테마 효과(그라데이션·그림자) 없음**, 캔·구름 도형이 사각형, 도형 안 글 정렬이 원본(왼쪽)과 달리 가운데 | p02·p07 1장 | 하 | 우리 편집기: style(fillRef/effectRef)·일부 prstGeom 미지원 | src/client/editor-slides.ts `PRESET_CLIP` 등 | 중 |
| P6 | 편집 화면 한글 낱자 단위 줄 바꿈(PowerPoint는 어절) | p03 2장, p08 | 하 | CSS | `.slide-textbox` CSS | 하 |
| P-v1 | 보기 화면(그리고 편집 화면의 표·차트 조각) 차트 글자에 **띄어쓰기가 끼고 제목이 굵게**: "1월"→"1 월", "추세(꺾은선)"→"추세 ( 꺾은선 )" | p05 1장, p08 1장 차트 | 하~중 | LibreOffice 차트 글자 처리 | — | LibreOffice 한계 |
| P-v2 | 보기 화면 표 머리 줄 글자가 굵은 흰 글씨가 아니라 가는 글씨 | p04 1장, a2 2장 | 하 | LibreOffice 표 스타일(firstRow b=1) 해석 | 전처리로 굵기 직접 지정 가능 | 하~중 |
| P-v3 | 보기 화면 자동 맞춤을 PowerPoint보다 더 작게 줄임, 글머리 기호가 글과 떨어짐 | p03 1장 | 하 | LibreOffice 자체 축소 계산 | — | LibreOffice 한계 |
| (좋음) | 보기 화면은 장 수·배치·그림 자르기·회전·그라데이션·그림자·그룹·연결선·표 병합·차트 종류 모두 PowerPoint와 같다 | p04~p08 | — | — | — | — |

## 5. 고칠 순서 상위 15 (세 형식 통합)

순서 = (읽는 사람이 받는 피해 × 자주 나오는 정도) ÷ 고치는 수고. "편집"은 향기랩·8000 연동에서 사람들이 실제로 보는 화면.

| 순위 | 결함 | 화면 | 원인 | 난이도 | 고칠 곳 |
|---|---|---|---|---|---|
| 1 | X5 오류값 `[object Object]` | 엑셀 보기 | 우리 코드 | 하(한 줄) | src/spreadsheet.ts:410 — 결과가 `{error}`면 그 문자열 |
| 2 | X4 공유 수식 칸 빈칸(견적 금액 누락) | 엑셀 보기 | 우리 코드 | 하~중 | src/spreadsheet.ts displayValue / applyComputedFormulaValues — 캐시값 우선 |
| 3 | X1+X3 숫자 서식(시간·날짜·괄호 음수·빨강·지수·분수·앞 0·천 단위·글자 붙은 서식·####) | 엑셀 보기·편집 | 우리 코드 | 중 | src/spreadsheet.ts formatNumber/formatDate/classifyNumFmt, cellValueShared.ts, grid-editor 표시 — SSF 하나로 통일 |
| 4 | D1 그림이 안 보임 | 워드 편집 | 우리 코드 | 중(줄 안) / 상(떠 있는 그림) | src/docx/import.ts:222 + schema.ts 그림 노드 + 미디어 내려주는 경로; 저장 때 원본 그림 보존 확인 |
| 5 | D2 번호 매기기 사라짐 | 워드 편집 | 우리 코드 | 중 | src/docx/import.ts numbering.xml 해석, groupListItems(698), 스타일의 numPr도 |
| 6 | D7+D8 표 스타일 테두리·줄무늬·열 너비 | 워드 편집 | 우리 코드 | 중 | src/docx/import.ts 표 읽기(tblStyle·tblGrid) |
| 7 | X6 차트가 없음 | 엑셀 보기 | 우리 코드 | 중 | public/js/viewer.js에 src/client/gridCharts.ts 재사용 |
| 8 | X7 조건부 서식 없음(보기) / 아이콘·수식 규칙 없음(편집) | 엑셀 보기·편집 | 우리 코드 | 중 | viewer.js + condFormatParse.ts(iconSet·expression) |
| 9 | P1 자동 맞춤 무시로 글 넘침 | PPT 편집 | 우리 코드 | 하~중 | slideContent.ts bodyPr normAutofit → editor-slides.ts 글자 배율 |
| 10 | D10+P6 한글 어절 단위 줄 바꿈 | 워드·PPT 편집(엑셀 줄 바꿈 칸도) | CSS | 하 | 편집 본문·글상자·격자 칸에 `word-break: keep-all; overflow-wrap: anywhere` |
| 11 | D5 글상자 안 보임 | 워드 편집 | 우리 코드 | 중 | src/docx/import.ts — 최소한 글상자 안 글을 문단으로 꺼내 표시 |
| 12 | P2 그림 자르기·회전 무시 | PPT 편집 | 우리 코드 | 하 | editor-slides.ts renderSlideVisual — srcRect로 크기·위치, rot로 transform |
| 13 | D3+D4 머리글·바닥글·쪽 번호·각주 | 워드 편집 | 우리 코드 | 중(표시만) | src/docx/import.ts header/footer/footnotes 파트를 읽어 위·아래에 읽기 전용으로 |
| 14 | P3+P4 줄·단락 간격, 상속 글자색·테마 글꼴 | PPT 편집 | 우리 코드 | 중 | editor-slides.ts 문단 스타일, src/pptx/inheritance.ts·theme.ts |
| 15 | X8 테두리 색·선 종류(+X9 무늬 채우기·회전·맞춤 축소) | 엑셀 보기·편집 | 우리 코드 | 중 | src/spreadsheet.ts:627 borderWeight → 색·종류까지 GridStyle에 |

다음 후보: D6 탭 정지(워드 편집), D9 제목 스타일 색, D-v2 글상자 양쪽 흐름 전처리, P-v2 표 머리 줄 굵기 전처리, X10 문자 정렬.

## 6. LibreOffice를 계속 쓰는 한 못 고치는(또는 매우 어려운) 것

- **한글 어절 단위 줄 바꿈(보기 화면, D-v1)** — LibreOffice 26.2도 한글을 낱자에서 끊는다. 줄 끝 모양이 Word와 다르고 긴 문서는 쪽이 밀릴 수 있다. 전처리로 바꿀 설정을 찾지 못했다.
- Word의 한글–숫자/영문 사이 자동 간격(D-v6), 2단 문서의 단별 각주(D-v4), 차트 글자 띄어쓰기·굵기(P-v1), 자동 맞춤 축소 비율(P-v3).
- PPT 편집 화면의 표·차트는 LibreOffice 그림을 잘라 쓰므로 P-v1·P-v2가 편집 화면에도 그대로 나온다.
- 반대로 **엑셀은 LibreOffice와 무관**하고, **워드·PPT 편집 화면 결함은 전부 우리 코드**라 LibreOffice를 바꾸지 않아도 고칠 수 있다.

## 7. 다시 재는 법(한 명령)

```
npm run audit:office                              # 정답 PDF가 없으면 MS Office로 만들고(ref_office.ps1), 시험 서버를 8141에 따로 띄워 찍고, 점수를 낸다
npm run audit:office -- --only docx --port 8142
```
- 결과: `docs/quality-audit/out/scores.md`(형식별 요약 + 문서별 + 점검표 실패 항목), `scores.json`, 나란히 그림 `out/sbs/`. 도구는 `audit/office/`(gen_corpus.py 문서 만들기, ref_office.ps1·ref_one.ps1 정답 PDF, capture.mjs 화면 찍기, compare.py 비교·점수, checklist.json 점검표, run.mjs 묶음).
- 점수 세 가지(2026-11-20 기준값):
  1. **점검표 통과율**(`audit/office/checklist.json`) — Office에 보이는 글자·그림이 우리 화면에도 있는가(결함 번호별). 고치면 바로 오른다. 지금: 워드 보기 8/8 · 워드 편집 **0/11** · PPT 보기 2/2 · PPT 편집 1/1 · 엑셀 보기 **0/17** · 엑셀 편집 **4/17**.
  2. **글자 일치**(Office PDF 글자 순서와 비교, 0~1) — 워드 보기 0.99 · 편집 0.92, PPT 보기 0.95 · 편집 0.77(표·차트는 그림 조각이라 글자로 안 잡혀 낮게 나옴).
  3. **모양 일치**(잉크 겹침 F1, 0~1) — 워드 보기 0.81(낱자 줄 바꿈 차이가 깎음) · PPT 보기 0.95 · PPT 편집 0.92. 결함을 콕 집지는 못하므로(그림 자르기 오류도 0.97로 나옴) 판정은 점검표와 나란히 그림으로.
- 새 결함을 고칠 땐 checklist.json에 그 결함을 잡는 항목을 먼저 넣어 실패를 확인하고, 고친 뒤 다시 돌린다. PPT 편집(P1~P5)은 지금 점검표 항목이 거의 없어 그림 판정에 기대므로, 고칠 때 DOM 검사 항목(글자 배율·transform 등)을 같이 넣기를 권한다.

## 8. 확인하지 못한 것·주의

- **스마트아트**는 python-pptx로 만들 수 없어 빠졌다. 고객 문서(data/originals)는 규칙대로 쓰지 않았다 — 합성 문서에 없는 기능(수식 개체, 변경 내용 추적, 메모, 워드아트, 큰 문서 성능)은 미확인.
- 워드 시험 파일은 Word로 다시 저장하지 못했다(호환 모드). 엑셀 a3-fidelity-grid(1200행)는 정답 PDF가 늦게 만들어져 눈 비교는 생략했다.
- 엑셀 정답 PDF는 "격자선·행열 머리 인쇄, 폭 1쪽 맞춤"으로 만든 것이라 엑셀 화면 그 자체는 아니다. 틀 고정은 정적 화면이라 확인 못 함. 보기 화면의 PNG 경로(내려받기 권한 없음)는 같은 PDF를 서버가 그린 것이라 따로 찍지 않았다.
- 작업 중 사용자의 엑셀 창(PID 26272, 13:07 시작)이 어느 순간 사라졌다. 스크립트는 자기가 새로 띄운 Office 프로세스 번호만 끄도록 되어 있어 끈 기록은 없지만, 확인 바람.
