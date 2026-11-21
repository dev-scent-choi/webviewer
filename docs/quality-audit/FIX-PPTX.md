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

## 6. 적대적 검토(REVIEW-PPTX.md) 항목별 답

검토 대상 5f950ef → 이번 고침 c9f8f9a…42b163d. 재현 파일(review-repro/r1·r2·r4·r5·r6·r7, pipe-deadlock)은 모두 시험으로 옮겼다
(`tests/pptxQualityFixes.test.ts` 37개, `npm run verify:pptx` 23항목). 큰 출력물은 커밋하지 않았다.

| 항목 | 결과 | 무엇을·근거 |
|---|---|---|
| C1 ReDoS(표 스타일 id를 정규식에) | **고침** | GUID 모양만 받고 indexOf로 찾는다. r4 값(`(a|a)+$`, 40자 역추적, `{X(}`)이 100ms 안에 끝나고 굵게 안 함(시험). 레인 전체 `new RegExp` 점검: 나머지는 상수 태그 이름뿐 |
| I1 큰 파일 통째 재압축 | **고침** | JSZip 비동기로 표 있는 슬라이드만 다시 쓰고, 나머지 파트는 원래 압축 바이트 그대로(무압축 미디어는 STORE 유지). 60MB 미디어: 3.1초 동기 정지 → 0.11초, 이벤트 루프 최대 공백 약 50ms(시험: 30MB에서 1.5초·400ms 안, 미디어 바이트 동일) |
| I2 협업자 원문 문자열 주입 | **고침** | 방(Y.Doc)마다 "서버가 원본 파일에서 읽은 조각" 목록(TrustedFragments)을 두고, 문단·글자 속성·글자 테마 색·줄 바꿈 속성·필드·수식·문단 끝 속성 원문은 목록에 있을 때만 XML에 붙인다(아니면 모델링한 값만, 필드는 글자만). 목록은 방이 사는 동안 쌓기만 하고 저장 파일로 다시 만들지 않는다(DOCX 갈래의 되돌리기 손실 함정 회피). 시험: r6 + 위조 pPr, 실제 브라우저: 별도 Yjs 클라이언트가 끝 속성에 다른 도형을 끼워 넣은 뒤 저장 → 주입 없음, XML 올바름 |
| I3 숫자 문자 참조 | **고침** | 세 파서 htmlEntities, 조각 빌더가 직접 이스케이프(속성의 줄 바꿈·탭은 &#10;·&#9;). 화면 "가나다", 고쳐 저장해도 `&amp;#` 없음, `typeface="Evil&#10;Font"` 유지(시험 + 브라우저 왕복) |
| I4 문단 안 수식 소실 | **고침** | 문단 바로 아래 자식 순서를 정확히 읽고, 모르는 자식(mc:AlternateContent/a14:m 등)은 원문째 보존 조각으로 그 자리에 대체 글자(mc²)를 보인다. 한 글자 고쳐 저장해도 수식 원문 그대로(시험 + 브라우저 왕복). 잠금은 필요 없어졌다 |
| I5 최상위 mc:AlternateContent 도형 | **고침** | Fallback(없으면 Choice)의 도형·그림·연결선·표/차트·그룹 자리를 원본 렌더링 조각으로. 편집 대상이 아니라 저장은 원문 그대로(시험 + 브라우저: 조각 표시·저장 원문 유지) |
| I6 run.mjs 파이프 교착 | **고침** | 서버 stdout/stderr를 파일 기술자로. 공식 `npm run audit:office -- --only pptx --port 8155 --no-office`가 끝까지 돌았다(아래 숫자) |
| 운영 실행기 확인(읽기만) | **문제 없음** | `run-webviewer.bat`(8000)·`start.bat`, 예약 작업 HyanggilabWebviewer가 부르는 `hyanggilab-deploy/run-webviewer-hyanggilab.bat`(8010) 모두 `npm run dev`/`npm start`를 리디렉션 없이 자기 콘솔 창에 쓴다 — 콘솔이 출력을 비우므로 아무도 안 읽는 파이프에 막히는 일은 없다 |
| Minor 1 zip 시각 | **고침** | 저장 zip 엔트리 시각 고정 → 같은 내용이면 같은 바이트(시계를 61초 옮겨도 동일 — 시험) |
| Minor 2 Shift+Enter | **고침** | 문단 안 줄 바꿈 `<a:br/>`(브라우저 왕복) |
| Minor 3 자동 맞춤 고정 | **고침** | 내가 글을 고친 맞춤 상자만 PowerPoint 단계(줄 간격 10/20% → 글자 7.5%씩)로 다시 맞춰 meta에 적고, 저장 때 normAutofit 값을 새로 씀(열기만으로는 안 바꿈). 레이아웃 배율은 상속하지 않고 종류만. 브라우저: 글을 늘리면 24px → 7.3px, 저장 fontScale 32500 |
| Minor 4 빈 줄 높이 | **고침** | endParaRPr sz |
| Minor 5 색 변환 순서 | **고침** | 파일 순서대로, hue/sat/lum Mod·Off, comp·inv·gray·gamma·invGamma·채널 Mod/Off |
| Minor 6 hlinkClr | **고침** | `ahyp:hlinkClr val="tx"`면 글자 색 그대로 |
| Minor 7 자동 번호 | **고침** | ST_TextAutonumberScheme 41가지 전부, 한자 수 十一·二十·百五 |
| Minor 8 테마 색·그림 글머리 | **고침** | 문단 직접 buClr은 서버가 테마로 풀어 둠, buBlip은 그 수준이 적힌 파트(슬라이드·레이아웃·마스터) 관계로 그림을 찾아 글자 대신 그림 |
| Minor 9 뒤집힌 그룹 그라데이션 | **결정·고침** | ECMA-376 20.1.8.33(rotWithShape)은 회전만 정하고 뒤집기는 정하지 않는다 → 보기 화면(LibreOffice)과 같게 그라데이션은 뒤집지 않음. MS Office 정답은 아직 없음(생기면 확인) |
| Minor 10 점검표가 자기 확인 | **고침** | `officePos` 18개: 글자 위치를 MS Office 정답 PDF의 같은 글자 위치와 비교(±1.5%). 처음 돌렸을 때 4개 실패 → 아래 Minor 11로 고친 뒤 전부 통과 |
| Minor 11 남은 위치 차이 | **고침** | 원인은 글꼴 메트릭이 아니었다(정답 실측: PowerPoint 한 줄 = 1.2×크기가 맞음). ① 넘친 세로 가운데 상자에서 편집 영역이 줄어 글이 아래로 밀림(제목 10pt) → 줄이지 않음, ② 문단 기준선 글꼴과 글자 글꼴이 달라 줄마다 1~2pt 벌어짐(p08 8줄에 10pt) → 같은 글꼴 문단은 문단도 그 글꼴, ③ 자동 맞춤 글자 크기를 PowerPoint처럼 0.5pt 단위로 내림(32×47.5%=15.2 → 정답 글자 높이 15.0pt) — 15.2로 그리면 첫 단락이 한 줄 늘어 마지막 단락이 20pt 밀렸다. 앞 커밋의 "단락 간격도 줄임" 근사는 정답 실측과 달라 되돌림. p03 편집 모양 0.723 → 0.992 |
| 편집 화면 차트 | **원래 있음** | 차트는 원본 렌더링 조각으로 보인다(p05·p08·a2). 글자 띄어쓰기("1 월")는 LibreOffice 한계(P-v1) — 근사 SVG로 바꾸면 축·색이 더 달라져 유지 |
| Minor 12 스마트아트 | **일부(합성 표본만)** | 이 PC의 공개 가능한 시험 폴더(webviewer tests·audit, artcraft-eval)에 스마트아트가 든 pptx가 없다. 쪽 그림이 있으면 원본 렌더링 조각(예전 그대로), 없으면 그림 대체본(dsp:drawing) 도형·글자를 그리게 했고, ECMA/Microsoft 확장 구조로 만든 합성 표본으로만 시험 — 실제 PowerPoint 스마트아트로는 미확인 |

### 숫자(공식 `audit:office --only pptx --no-office`, 8155)

| 지표 | 첫 보고(5f950ef) | 지금 |
|---|---|---|
| PPT 편집 점검표 | 34/34(구현 값 확인 위주) | **52/52**(정답 위치 비교 18개 포함) |
| PPT 편집 모양 | 0.953 | **0.975**(p03 0.839→0.992, p02 0.931→0.971, p05 0.853→0.879, p08 0.955→0.963) |
| PPT 보기 모양·글자 | 0.95·0.945 | 0.95·0.945(회귀 없음) |
| verify:pptx | 16/16 | **23/23** |

### 성능(200장 덱, 검토와 같은 `perf-200.mjs`, 8145)
업로드→준비 145초(LibreOffice 변환, 검토 131초와 같은 수준 — 기기 부하 영향), 편집 화면 첫 캔버스 0.56초·카드 200개,
슬라이드 넘김 2.4~6.3초(변환 직후 서버가 쪽 그림을 만드는 동안 — 검토 5.5~9.1초와 같은 성질), 타자 지연 p50 16.6ms·p95 19.3ms·최대 22.8ms
(검토 16.8/17.9/33.6ms — 자동 맞춤 다시 맞춤을 넣고도 한 프레임 수준), JS 힙 10MB.

### 그 밖
- vitest 전체: 896개 중 2개가 전체 실행 때 30초 시간 초과(room.ts pptx 자동 저장 id 유지, xlsx 창 단위 시딩) — 다른 갈래 작업으로 기기가 바쁠 때였고, 그 파일만 다시 돌리면 모두 통과(67/67). 이번 변경과 무관한 시간 초과로 본다.
- 건드린 공유 파일 추가: `src/converters/libreOffice.ts`(전처리 결과 await — 타입이 Promise 허용), `audit/office/run.mjs`(로그 파일 기술자), `audit/office/compare.py`(officePos 비교 덧붙임).

## 7. 2차 검토(REVIEW-PPTX-r2.md) 항목별 답

커밋 28d58db…b064914. 새 재현(r10, e2e-room-recreate 등)은 시험으로 옮겼다(`tests/pptxQualityFixes.test.ts`, `npm run verify:pptx` 24항목, 새 `npm run verify:reconnect` 12항목).

| 항목 | 결과 | 무엇을·근거 |
|---|---|---|
| RX-C1 재연결 시 옛 상태가 새 방에 합쳐짐(운영에도 있던 결함) | **고침(공유 동시편집 층)** | 방마다 세대(epoch). 서버가 접속마다 세대를 알리고(메시지 7), 클라이언트는 다음 접속 주소에 `epoch`로 싣는다. 세대가 다르거나 방이 없으면 서버는 동기화 메시지를 하나도 받지 않고 거절(메시지 8, 닫기 4409) — 방을 새로 만들지도 않는다. 클라이언트는 다시 접속하지 않고 한국어 안내 후 새로 불러온다. 끊긴 동안 입력한 내용이 있었으면 "반영되지 않았다"고 알리고 사용자가 누를 때 다시 불러온다 — 옛 세대의 Yjs 변경을 새 세대 문서에 다시 적용하는 것이 바로 이 중복이라 안전하게 되살릴 방법이 없어 다시 보내지 않는다. PPT 저장 직전 안전망: 같은 원본 슬라이드·같은 항목 id가 둘이면 첫 것만 저장하고 알림(sldId 중복 손상 방지). 실제 브라우저(`verify:reconnect`): PPT·워드·엑셀 모두 끊기→방 재생성→재접속에서 거절·안내·새로 불러오기, 화면·저장 파일 한 벌. 같은 시험을 거절을 끈 서버에서 돌리면 **워드 문단 25 → 50, 저장 파일도 두 벌**로 실패(시험이 결함을 잡는다). 검토의 `e2e-room-recreate`도 카드 3·sldId 3·링크 유지. PDF 편집기에도 같은 한 줄을 넣었다(브라우저 시험은 세 형식만) |
| R2-I1 속성 값 안 `>`로 문단 자식 순서 뒤섞임(회귀) | **고침** | 정규식 대신 작은 XML 토큰 읽개(따옴표·주석·CDATA·처리 지시·DOCTYPE). r10과 속성 값 흔들기(`a>b`, `/>`, 따옴표 섞임, CDATA `C<>`) 시험 |
| R2-I2 방 재생성 뒤 편집했던 런의 하이퍼링크가 조용히 사라짐 | **고침** | 글자 속성(runProps)은 바이트 일치 신뢰 대신 구조 검증: 허용 속성(값 모양 확인)·허용 자식만, 관계 id는 이 슬라이드 rels에 있고 요소에 맞는 종류(하이퍼링크→hyperlink/slide, 그림→image, 소리→audio). 방이 새로 만들어져도 유지(시험: 빈 신뢰 목록에서도 링크·lang 남음). 그래도 원문을 못 지키면 저장 알림 띠로 보인다(조용히 버리지 않음) |
| R2-I3 자동 맞춤이 창 크기에 따라 다름 | **고침** | 측정을 슬라이드 좌표(1pt=4px) 화면 밖 사본(문서 직렬화)에서, 안쪽 크기는 meta(EMU)로. 지금 단계에서 시작해 "고친 뒤 높이 > 상자 높이와 고치기 전 높이 중 큰 쪽"일 때만 작게, 한 단계 키워도 상자에 들어갈 때만 크게. verify:pptx: p03 한 글자 편집이 창 1400·900·1920에서 모두 47.5% 유지(검토 재현: 40%로 떨어졌었다) |
| R2-m1 다른 슬라이드의 믿는 조각 재사용 | **고침** | 모든 원문 조각의 관계 id를 이 슬라이드 rels·종류로 확인(시험: rId2가 그림 관계면 링크 빼고 알림) |
| R2-m2 전처리의 압축 폭탄 상한 빠짐 | **고침** | `assertZipNotTooLarge`(서버 공통 상한)를 풀기 전에. 시험: 적힌 풀린 크기 4GB → 거절 |
| R2-m3 레이아웃에서만 맞춤을 물려받은 도형 | **고침** | 슬라이드 bodyPr에 맞춤 요소가 없으면 저장 때 `normAutofit` 삽입(prstTxWarp 뒤, 100%면 안 넣음, 다른 맞춤 종류면 그대로) |
| R2-m4 감싼 도형 위치 | **고침** | Choice·Fallback 같은 개체 id의 두 자리를 합친 영역을 잘라 온다(LibreOffice가 어느 쪽을 그려도 조각 안) |
| Minor 10 위치 허용 오차 | **조임** | ±1.5% → ±1%(지금 최대 차이 0.96%, 720pt 높이 기준 약 5pt) |
| Minor 9 뒤집힌 그라데이션·Minor 12 스마트아트 | **그대로(검토도 연기 수용)** | MS Office 정답·실제 스마트아트 표본이 생기면 확인 |

### 공유 파일 변경(다른 갈래와 합칠 때)
- `src/collab/room.ts`: ① `/* ===== [RX-C1] … ===== */` 블록 하나(joinEditSession 바로 위: `roomEpochs` WeakMap, `epochOf`, `rejectStaleReconnect`, `sendRoomEpoch` — Room 타입은 안 바꿈), ② `joinEditSession` 첫머리 2줄(거절 확인, 세대 알림), ③ import 2줄(`node:crypto` randomUUID, `./roomEpoch.js`), ④ pptx 저장 블록 2줄(저장 알림 `pptxSaveNotice`) + import 1줄.
- 새 공유 파일: `src/collab/roomEpoch.ts`(메시지 번호 7·8, 매개변수 `epoch`, 닫기 코드 4409), `src/client/roomEpoch.ts`(`guardRoomEpoch`, 안내 화면).
- 편집기 한 줄씩: `src/client/editor.ts`(워드), `src/client/grid-editor.ts`(엑셀), `src/client/editor-pdf.ts`(PDF), `src/client/editor-slides.ts` — `guardRoomEpoch(provider, ydoc)` + import.
- `package.json`: `verify:reconnect` 스크립트 한 줄. `scripts/verify-reconnect-e2e.mjs` 새 파일.
- 그 밖: `src/converters/libreOffice.ts`(앞 라운드), `src/zipGuard.ts`는 가져다 쓰기만.

### 숫자
- 공식 `audit:office -- --only pptx --no-office`: PPT 편집 점검표 **52/52**(위치 허용 ±1%), 편집 모양 **0.975**, 보기 0.95·0.945(회귀 없음).
- verify:pptx **24/24**, verify:reconnect **12/12**, tsc·build 통과, vitest 전체 **906/906**(110파일, 한 번에 통과 — 앞 실행에서 기기 부하로 room 시험이 시간 초과했을 때는 그 파일만 다시 돌려 통과).

## 8. 3차 검토(REVIEW-PPTX-r3.md) 항목별 답

커밋 24f40fb…017cb96.

| 항목 | 결과 | 무엇을·근거 |
|---|---|---|
| R3-I1 반쯤 열린 연결 + 혼자 편집하는 방이 바로 닫혀 입력 손실 | **고침(공유 동시편집 층)** | ① **방 닫기 유예**: 마지막 참여자가 나가면 저장은 지금처럼 바로 하되(편집이 있었을 때만), 방은 같은 세대로 45초 남긴다(`DOCVIEW_ROOM_CLOSE_GRACE_MS`, 0=끔, 최대 120초). 그 안에 다시 붙으면 같은 세대라 Yjs가 빠진 변경을 정상적으로 다시 보낸다 — 새로 고침·손실·중복 없음. 남겨 두는 방은 최대 20개(넘으면 가장 오래된 방부터 닫음), 큰 엑셀 방(largeRoom)은 남기지 않음, 문서 삭제(`forceCloseRoom`)는 유예 없이 바로 닫음, 타이머는 `unref`(서버 종료를 붙잡지 않음 — 남은 방은 이미 저장됨). 유예 중 추가 자동저장은 없다(편집이 없으니 예약 자체가 없음). ② **그래도 방이 새로 만들어진 경우**(유예 뒤, 서버 재시작): 안내에 "연결이 끊긴 뒤 입력한 내용은 반영되지 않았을 수 있다"를 적고, 이 화면에서 고친 곳(문단·셀)의 현재 글자를 모아 **'내 입력 내용 복사'·'내려받기'**를 건넨다. 반쯤 열린 연결은 보낸 것이 닿았는지 알 수 없으므로 고친 기록은 접속 상태와 무관하게 모두 남긴다(보수적). 그런 내용이 있으면 **저절로 새로 고치지 않는다**. 실제 브라우저(`verify:reconnect`): PPT·워드·엑셀 반쯤 열림 → 재접속에서 안내 없이 이어지고 그 사이 입력이 저장 파일에 한 번씩 |
| R3-m2 2.5초 자동 새로 고침 | **고침** | 잃을 입력이 없을 때만, 보이는 10초 초읽기(aria-live) 뒤 다시 불러온다. 입력이 있었으면 사용자가 누를 때만 |
| R3-I2 세대 값을 모르는 옛 번들은 여전히 합쳐짐 | **고침(서버 판정)** | 주소에 `epoch`가 없는 연결은 첫 SyncStep1의 상태 벡터를 이 방 세대의 Yjs 상태와 비교해, 비어 있지 않은데 **겹치는 client id가 하나도 없으면**(= 이전 세대 문서를 가진 탭) 동기화 메시지를 하나도 받지 않고 메시지 8 + 닫기 **4409**(이유 문구 한국어)로 거절한다. 옛 번들은 4400대를 영구 거부로 다뤄 다시 붙지 않고 이유를 보여 준다. 새 번들은 보통 닫기(1000)에 자체 안내. 빈 상태 벡터(새 탭)·같은 세대 상태(유예 안 재접속)는 통과. y-websocket은 열리자마자 SyncStep1부터 보내므로 다른 메시지가 먼저 오면 판정 없이 통과. 브라우저 시험(옛 번들 흉내: 주소의 epoch 제거 + 메시지 7·8 버림): **PPT·워드·엑셀·PDF 넷 다** 서버 재시작 뒤 4409·다시 붙지 않음, 새 탭 내용·저장 파일 한 벌. 워드 옛 번들의 유예 안 재접속은 받아 주고 이어서 입력도 저장 |
| 배포 절차(선택) | 문서화 | 서버 판정으로 비밀 값 교체는 **필수가 아니다**. 그래도 옛 번들 탭을 확실히 모두 새로 고치게 하려면: 운영 실행 파일을 읽기만 해 확인한 결과 8010(`start.bat`)은 `start.local.bat`에서 `DOCVIEW_SECRET`을 고정해 넣고, 8000(`run-webviewer.bat` → `npm run dev`)은 그 줄이 주석이다(값은 보지도 옮기지도 않음). 배포 때 `start.local.bat`의 값을 새로 바꾸고 예약 작업을 다시 시작하면, 열린 탭은 재접속 때 토큰 오류(4401)로 멈추고 새로 고침 안내를 받는다. 대가: 그때까지 발급된 열람 링크(연동 쪽 토큰 포함)가 모두 무효 |
| R3-m1 rPr의 이름공간 선언 삭제 | **고침** | `@_xmlns`/`@_xmlns:접두어` 속성은 남긴다(값은 공백·따옴표·`<>` 없는 URI 글자 512자까지). 모르는 속성·자식을 버릴 때도 저장 알림("알 수 없는 글자 속성을 빼고 저장했습니다.") — 전에는 조용히. 검토의 r12: `xmlns:x` kept, `<x:foo` kept, `x:tag`(스키마에 없는 속성)만 빠지고 알림. 시험 추가 |
| 저장 두 번 시험("expected 1, got 2") | **원인 확정: 시험 간섭 + 별개의 실제 이중 저장 하나 발견·수정** | 부하(방 시험 3벌 동시)에서 재현해 실행 기록으로 확인: 같은 파일의 앞 시험이 **room.js 첫 변환 때문에 20초 한도를 넘겨 시간 초과**되면 vitest는 다음 시험으로 넘어가지만 그 몸통은 뒤에서 계속 돌아, 다음 시험이 `resetModules`로 새로 만든 store에 저장한다 → 다음 시험의 첫 확인에서 2번. 실제 서버의 이중 저장이 아니다. 조치: 그 파일만 시간 한도 90초(`vi.setConfig`) — 같은 부하 3벌에서 그 파일 통과. 조사 중 **진짜 이중 저장 하나**를 찾았다: 2차에서 넣은 PPT 저장 알림(`PPTX_SAVE_NOTICE_MAP`) 쓰기가 자동저장 예약에서 빠져 있지 않아, 알림이 있는 저장마다 8초 뒤 같은 내용을 한 번 더 저장했다 → savedAt 쓰기처럼 `suppressAutosaveReschedule`로 감쌌다. 유예는 이 경로를 바꾸지 않는다(닫을 때 저장은 그대로 한 번, 편집 없으면 0번) |
| Minor 9 뒤집힌 그라데이션·Minor 12 스마트아트 | **연기(메모 유지)** | 9: MS Office 정답 PDF로 flipH/flipV + 그라데이션 각도 조합을 확인하기 전엔 손대지 않는다(LibreOffice 그림과 Office가 다를 수 있음). 12: 실제 스마트아트 표본(dsp:drawing 없는 파일 포함)이 생기면 그리기 경로 확인 |

### 공유 파일 변경(3차, 다른 갈래와 합칠 때)
- `src/collab/room.ts`:
  - `[RX-C1]` 블록 확장(joinEditSession 바로 위): `clientEpochOf`, `refuseStale`, `passesStateCheck`(+`stateChecked` WeakSet), 유예용 `MAX_LINGERING_ROOMS`·`lingeringRooms`·`closeWithoutGrace`·`roomCloseGraceMs`·`lingerRoom`·`wakeLingeringRoom`.
  - `joinEditSession`: 첫머리(세대 읽기 → 거절 확인 → 유예 방 깨우기 → 방 → 세대 알림), 메시지 처리와 대기 메시지 처리에 `passesStateCheck` 한 겹.
  - 닫기 정리를 `teardownRoom(room)` 함수로 분리(내용은 그대로) — `closeConn`의 마지막 참여자 분기는 `lingerRoom`이면 저장만, 아니면 `teardownRoom`.
  - `forceCloseRoom`: `closeWithoutGrace` 표시 + 유예 중이던 방 바로 정리.
  - pptx 저장 알림 쓰기를 `suppressAutosaveReschedule`로 감쌈.
  - import: `./roomEpoch.js`에서 `MESSAGE_ROOM_STALE`·`ROOM_STALE_CLOSE_CODE`·`ROOM_STALE_REASON` 추가.
- `src/collab/roomEpoch.ts`: `ROOM_STALE_REASON` 추가, 4409 설명 갱신.
- `src/client/roomEpoch.ts`: 고친 곳 기록(afterTransaction), `recoveryText`, 안내 화면(복사·내려받기·초읽기) 다시 씀.
- `src/client/editor-pdf.ts`: 거부로 멈춘 뒤 status 이벤트가 '보기 전용'을 덮지 않게 한 줄.
- `vitest.config.ts`: env `DOCVIEW_ROOM_CLOSE_GRACE_MS: '0'`(기존 방 시험은 닫히면 바로 정리되길 기대 — 유예 시험은 따로 켬).
- 시험: `tests/roomCloseGrace.test.ts`(새), `tests/roomEpochRecovery.test.ts`(새), `tests/roomAutosaveLoop.test.ts`(시간 한도 한 줄), `tests/pptxQualityFixes.test.ts`(R3-m1).
- `scripts/verify-reconnect-e2e.mjs`: 시험 서버를 직접 띄우고 다시 띄우도록 다시 씀(39항목).

### 숫자
- 공식 `audit:office -- --only pptx --port 8155 --no-office`: PPT 편집 점검표 **52/52**, 편집 모양 **0.975**, 보기 **0.95·0.945**(변화 없음). Office는 띄우지 않음.
- verify:pptx **24/24**, verify:reconnect **39/39**(유예·반쯤 열림 각 3형식, 서버 재시작 3형식 + 입력 없는 탭 초읽기, 옛 번들 4형식), tsc 통과, vitest 전체 **913/913**(112파일).
- 남은 것: 부하(방 시험 여러 벌 동시)에서 다른 room 시험 몇 개가 첫 import 때문에 20초 시간 초과하는 것은 그대로(기준 94ab62b에도 있음, 혼자·전체 실행에선 통과). 엑셀 '내 입력 내용'은 셀 키(`시트:행:열`) 그대로 보인다.
