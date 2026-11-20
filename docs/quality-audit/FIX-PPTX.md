# PPT(pptx) 표시 품질 고침 — 작업 보고

가지 `office-quality-pptx`(점검 가지 94ab62b에서 갈라짐) · 작업 폴더 `projects\webviewer-q-pptx` · 2026-11-20
운영 폴더·운영 포트(8000/8010)·MS Office는 건드리지 않았다. 시험 서버는 이 폴더에서 8145(·공식 점검 명령은 8155), 전용 데이터·임시·LibreOffice 프로필(`.audit-run/`), 끝나면 띄운 PID만 껐다. 정답 PDF는 점검 때 만든 `out/ref`를 복사해 썼다(`--no-office`).

## 1. 숫자(전 → 후)

측정은 `audit:office`와 같은 capture.mjs·compare.py를 이 폴더의 격리 서버(8145)에 돌린 값이다. 공식 한 줄 명령(`npm run audit:office -- --only pptx --port 8155 --no-office`)도 세 번 돌렸는데, 세 번 모두 중간(p03~p04)에 시험 서버나 캡처 프로세스가 바깥에서 끝나 버려 끝까지 가지 못했다(오류 기록 없이 종료 — 같은 시각 다른 갈래가 돌던 때. 같은 문서를 8145 서버로 찍으면 매번 끝까지 통과). `--no-office` 덕분에 Office는 한 번도 뜨지 않았다.

| 지표 | 전 | 후 |
|---|---|---|
| PPT 편집 화면 점검표(DOM 항목, `audit/office/pptx_checks.json`) | 9/36 (아래 P6 항목 교체 반영 시 11/34) | **34/34** |
| PPT 편집 모양(잉크 겹침, 10개 평균) | 0.923 | **0.953** |
| PPT 보기 모양 / 보기 글자 | 0.95 / 0.945 | 0.95 / 0.945 (회귀 없음) |
| PPT 편집 글자 | 0.769 | 0.769 (표·차트는 그림 조각이라 글자로 안 잡힘 — 결함 아님) |

문서별 편집 모양: p01 0.912→0.990 · p03 0.772→0.839 · p06 0.973→1.000 · p07 0.937→0.989 · p08 0.906→0.955 · a2 0.961→0.976 · p05 0.844→0.853 · p02 0.932→0.931 · p04·a3 그대로(0.999/0.995).

점검표에 결함마다 DOM 항목을 새로 넣었다(보고서 7장 권고): 자동 맞춤 뒤 글자 pt·마지막 단락 위치·줄 간격 배수, 그림 자르기 비율(왼·위·오른·아래)·회전각, 도형 회전각, 줄 간격·단락 뒤 pt·첫 줄 공백, 상속 글자색(rgb)·테마 글꼴·굵기·글머리 문자·수준 크기, 테마 그라데이션·그림자·캔/구름 모양·도형 글자 정렬, 한글 줄 바꿈 방식, 배경색. 기대값은 MS Office 정답 PDF에서 읽었다. 측정 코드 `audit/office/pptx_dom.mjs`(capture.mjs가 슬라이드마다 부름), compare.py는 결과를 "점검표 실패 항목"에 `editDom`으로 더한다. 말뭉치에 없는 기능(하이퍼링크·문단 안 줄 바꿈·자동 번호·eaLnBrk=0·세로 쓰기·회전 글상자·저장 왕복)은 `npm run verify:pptx`(16/16)에 넣었다 — 말뭉치에 새 파일을 넣으면 정답 PDF가 없어 다른 갈래의 점검이 Office를 띄우려 하므로 넣지 않았다.

## 2. 결함별

| 번호 | 결함 | 결과 | 근거·고친 곳 |
|---|---|---|---|
| P1 | 자동 맞춤(normAutofit fontScale·lnSpcReduction) 무시로 글 넘침 | **고침** | 글자 크기만 배율(`--slide-fpt`), 줄 간격·비율 단락 간격에서 줄임 비율을 뺀다. 레이아웃·마스터 상속, spAutoFit/noAutofit은 맞춤 없음. p03 1장 본문 32pt→15.2pt, 13단락 모두 슬라이드 안(점검 3/3). 정답보다 단락 사이가 약간 촘촘(PowerPoint 내부 맞춤 계산과 1~2px 차). inheritance.ts·slideTextLook.ts·paragraphLook.ts |
| P1+ | 세로 쓰기(bodyPr vert) | **고침** | vert/vert270/eaVert/wordArtVert/mongolianVert → writing-mode. verify:pptx |
| P2 | 그림 자르기(srcRect)·회전 무시 | **고침** | 그림 자르기 비율대로 확대·이동(측정 0.25/0.10/0.05/0), 그림·도형·글상자 회전·뒤집기, **그룹 회전·뒤집기를 자식에 합성**(예전엔 무시). p06 모양 1.000, p07 화살표 30°. slideContent.ts placeXfrm·cropOf, editor-slides.ts |
| P3 | 줄 간격·단락 앞/뒤 간격 빠짐 | **고침** | 문단 직접값(pPr)과 수준 상속값 모두. PowerPoint처럼 글상자 첫 문단 앞 간격은 두지 않는다(p01 첫 글머리 높이 일치). 줄 앞 공백 유지(p08 "    - 항목"), 문단 marL/indent 직접값도 반영. textSchema toDOM·CSS |
| P4 | 상속 글자색·테마 글꼴·제목 굵게 빠짐 | **고침** | tint/shade를 선형 밝기로(부제 #404040→#898989, p01·a2), 테마 글꼴(+mj/+mn, 한글 칸이 비면 Hang 글꼴: Georgia/HY견고딕, Segoe UI/맑은 고딕), 수준별 굵게·기울임·대문자(구역 머리글 굵게), 글머리 글꼴·색·크기. theme.ts·inheritance.ts |
| P4+ | 글머리 자동 번호(buAutoNum) | **고침** | 1. a) (1) Ⅰ. ① 一. 등 체계, 수준별 이어 세기·다시 시작(플러그인). 단위 시험 + verify:pptx |
| P4+ | 그림 글머리(buBlip) | **일부** | 그림 대신 •로 보인다 — 글머리 그림은 레이아웃·마스터 파트의 관계를 따라가 미디어를 내려야 해서 이번엔 하지 않았다 |
| P4+ | 문단 직접 글머리 색이 테마 색(schemeClr)일 때 | **일부** | 수준 상속 글머리 색은 테마까지 풀지만, 문단에 직접 적힌 buClr는 RGB만 반영(문단 속성은 편집기 안에서 테마 없이 읽음) |
| P5 | 테마 효과(그라데이션·그림자) 없음, 캔·구름이 사각형, 도형 글자 가운데 | **고침** | fillRef→fmtScheme 채우기(phClr), lnRef→선 두께, effectRef/outerShdw→그림자(잘린 모양은 drop-shadow), 캔(윗면 포함)·구름 곡선 모양, 프리셋 모양의 글자 칸(타원·별·삼각형 등). 정렬 없는 도형 글자는 마스터 기본(왼쪽) — 예전 "그린 도형은 가운데" 규칙은 정답 p07과 반대라 뺐다(별 안 흰 글자가 밖으로 나가던 원래 문제는 글자 칸으로 해결). p07 0.937→0.989 |
| P5+ | 슬라이드 배경 bgRef | **고침** | 테마 배경 채우기 목록(idx 1001~, 그라데이션 포함) |
| P6 | 한글 낱자 단위 줄 바꿈 | **고침(보고서 판정 정정)** | 정답 PDF를 다시 보니 PowerPoint도 이 파일들(p03 2장 "주\|식회사는", p08 "달\|라질")은 **글자 단위**로 끊는다 — 마스터의 `eaLnBrk="1"` 때문. 그래서 편집기를 일괄 어절 단위로 바꾸지 않고 eaLnBrk를 따르게 했다(1/기본 = 글자 단위, 0 = 어절 단위 keep-all, 수준 상속·문단 직접값). p03 2장 줄 끝이 정답과 글자 하나까지 같아졌다. verify:pptx에 eaLnBrk=0 항목 |
| P-v1 | 차트 글자 띄어쓰기·굵기("1 월") | **못 고침** | LibreOffice 차트 글자 처리. 편집 화면 차트도 그 PDF 조각. 우리 차트 그리기(gridCharts.ts의 근사 SVG)로 바꾸면 축 범위·색 변화·눈금 등에서 오히려 원본과 더 달라져 바꾸지 않았다 |
| P-v2 | 보기 화면 표 머리 줄이 가는 글씨 | **고침** | 변환 사본에서만 머리글 행 글자에 b=1(src/pptx/loPreprocess.ts, 내장 스타일·tableStyles.xml 정의 따름). 보기·편집 표 조각 모두 굵게 확인(p04·a2 2장) |
| P-v3 | 보기 화면 자동 맞춤 축소 비율 | **못 고침** | LibreOffice가 자체 계산(파일의 fontScale을 쓰지 않음) — 편집 화면은 이제 파일 값을 쓴다 |
| 표(테두리·채우기·병합·칸 여백) | 편집 화면 | **원래 같음** | 표는 LibreOffice 그림 조각 — 정답과 같다(p04 모양 0.999). 머리 줄 굵기는 P-v2로 해결 |
| 선·연결선·화살촉 | 편집 화면 | **원래 같음** | 연결선은 그림 조각(그룹 회전 합성 후 영역 계산) |
| 하이퍼링크 | 편집 화면 | **고침** | 테마 하이퍼링크 색 + 밑줄, 저장 보존 |
| 레이아웃·마스터 자리표시자 | 편집 화면 | **원래 맞음** | 빈 자리표시자는 PowerPoint도 안 보인다. 꼬리말·쪽 번호는 슬라이드에 있을 때만 — 그대로 |
| 스마트아트 | — | **미확인** | 말뭉치에 없음(보고서 8장). 그림 조각 경로라 보기와 같을 것 |

## 3. 저장(열기 → 저장 → 다시 열기)

새로 그리는 값은 모두 표시 전용 meta·문단 속성에서 읽고, 저장은 예전처럼 고친 글상자의 `p:txBody`만 갈아 끼운다. 이번에 저장에서 빠지던 것을 더 막았다:
- 글자 속성 원본(`<a:rPr>` 전체 — lang·자간·하이퍼링크·테마 글꼴 참조·강조색)을 `runProps` 마크에 보관, 저장 때 바뀐 속성만 고쳐 쓴다(굵게만 끄면 b만 빠짐).
- 문단 안 줄 바꿈 `<a:br>`, 필드 `<a:fld>`(쪽 번호·날짜), 문단 끝 `<a:endParaRPr>` — 예전엔 글자를 한 자만 고쳐도 사라졌다. 이제 원문 순서대로 읽고 그대로 쓴다.
- 고치지 않은 글상자는 예전처럼 한 글자도 안 바뀐다(바이트 동일 시험 통과).
- 확인: 단위 시험(`tests/pptxQualityFixes.test.ts` 18개), verify:pptx "고친 글상자 저장" 항목(실제 편집기에서 타이핑 → 저장 → 내려받은 slide1.xml에 hlinkClick·a:br·buAutoNum·eaLnBrk·lang 유지).

## 4. 건드린 공유 파일(다른 갈래와 합칠 때)

- `src/converters/libreOffice.ts` — officeConverter 전처리 선택 한 줄(pptx면 boldPptxTableHeaders) + import 한 줄
- `public/css/style.css` — `.slide-textbox`·`.slide-visual` 구역만
- `audit/office/capture.mjs`(pptx_dom 호출 3줄), `audit/office/compare.py`(editDom 결과 합치기 5줄), `audit/office/run.mjs`(`--no-office` 2줄) — 모두 덧붙임
- 나머지는 pptx 전용: src/pptx/*, src/client/editor-slides.ts·slideTextLook.ts, src/collab/pptxShared.ts, tests/pptx*, scripts/verify-pptx-fidelity-e2e.mjs, audit/office/pptx_checks.json·pptx_dom.mjs

## 5. 검증

vitest 전체 110파일 879개 통과, tsc, build:editor·build:server, verify:pptx 16/16, audit:office(pptx) 34/34. 나란히 그림(`out/sbs/*-edit.png`)을 슬라이드마다 직접 봤다 — p01~p08·a2·a3 편집 화면이 정답과 배치·색·글꼴·줄 끝까지 맞고, 보기 화면은 표 머리 줄 굵기만 바뀌었다.
알려진 남은 차이: p03 1장 자동 맞춤 상자의 단락 사이가 정답보다 1~2px 촘촘, 제목 상자 글자가 정답보다 10px쯤 아래(제목 줄 높이 근사), 구름 모양은 원 9개 합으로 근사. vitest의 pptxCollab "모든 슬라이드 시딩 후 바이트 동일" 시험이 한 번 실패했다가 다시 돌리면 통과(3회 연속 통과) — 시간에 따라 흔들리는 시험으로 보이며 이번 변경과의 관련은 확인하지 못했다.
