# 향기랩 문서도구 — 모든 확장자를 처음부터 편집 모드로 연다

작성 2026-11-12. 기준: webviewer `view-shell` a96bf34(통합 보기 껍데기, 함께 병합), calc-site `cadview-ocr` d16fbee, cadview `e1-editor-core` 7d9a2e8(지시받은 02fbbc6 뒤로 문자 인식 커밋 3개가 더 있다 — 그 위에서 한다. 커밋 안 된 `base.ts`·`main.ts`·`test/desktop.test.ts` 세 파일은 다른 작업이라 건드리지 않는다).

## 0. 결정 요약

**사용자가 정한 것(영구 규칙, 바꾸지 않는다)**

1. 향기랩(calc-site `public/tools/index.html`)은 받는 **모든 확장자를 처음부터 편집 모드**로 연다. 열람 전용 화면은 어떤 경우에도 띄우지 않는다.
2. `.doc`·`.xls`·`.ppt`: webviewer가 LibreOffice로 `.docx`·`.xlsx`·`.pptx`로 바꿔 기존 편집기로 연다. **내려받기는 원래 형식**. 고치지 않았으면 올린 바이트 그대로, 고쳤으면 LibreOffice로 원래 형식으로 다시 바꾼다. 다른 호출자(8000 유지보수 시스템 등)의 동작은 바뀌지 않는다 — 향기랩만 보내는 표시로 켠다.
3. 이미지(png·jpg·jpeg·bmp·tif·tiff·webp): cadview(브라우저 안)에서 **래스터 편집이 켜진 채** 연다(문자 인식 이미지 보기 아님). 저장은 원래 형식. 문자 인식은 리본에서 그대로 쓸 수 있다. 데스크톱·8020 동작은 바꾸지 않는다.
4. NWD는 향기랩 받는 목록·경로에서 뺀다(편집기 없음). DWG는 cadview 그대로(이미 편집 모드 — §5에서 확인).
5. HWP는 rhwp 미러 그대로(이미 편집), PDF는 PDF 편집기 그대로(`pdfEditor:true`).

**이 설계가 정한 것**

| # | 결정 | 이유 |
|---|---|---|
| A1 | 켜는 표시는 **업로드 조회 문자열 `legacyEdit=1`**(`POST /api/documents/stream`), 향기랩의 `functions/api/docview/upload.js`만 붙인다. 토큰 옵션이 아니다 | 바꾸기는 편집기를 열기 **전**에 끝나야 한다 — 업로드가 토큰보다 먼저 오고, 바꾼 문서는 `editReady`라 향기랩이 변환을 기다리지 않는다. 토큰 때 바꾸면 업로드 직후 시작된 .doc→PDF 변환을 기다린 뒤 또 바꾸게 된다. `pdfEditor`처럼 "향기랩만 보내는 이름 붙은 표시"라는 뜻은 같다 |
| A2 | 바꾼 OOXML이 그 문서의 원본(`ext`=docx/xlsx/pptx)이 되고, 올린 바이트는 파생 산출물 `legacy`, 메타에 `legacy: { ext, sha256, hasMacros }`. 표시 이름(`meta.name`)은 `보고서.doc` 그대로 | 편집기·자동저장·쪽 그림·인쇄가 지금 코드 그대로 돈다. 바뀌는 곳은 업로드와 내려받기 두 곳뿐 |
| A3 | "고쳤는가" = 열린 방에 저장 안 된 편집이 있음(`editedSinceSave`, 시딩 끝난 뒤) **또는** 저장된 OOXML의 sha256이 바꾼 직후와 다름 | 방은 편집이 없으면 다시 저장하지 않는다(room.ts 시딩·종료 규칙) — 그래서 해시가 같으면 손대지 않은 것이다. 고쳤다가 되돌린 것은 "고침"으로 친다(안전한 쪽) |
| A4 | 다시 바꾸기가 실패하면 편집본을 OOXML 이름(`보고서.docx`)으로 내려준다 | 고친 내용을 잃지 않는 것이 형식보다 먼저. 정상 경로의 형식은 원래 형식 하나뿐 |
| A5 | 매크로(VBA): 업로드 때 찾아 두고, 편집기 맨 위에 경고 줄 하나 — "고치지 않고 내려받으면 원본 그대로, 고친 뒤 내려받으면 매크로가 빠집니다". 막지도 묻지도 않는다 | 손대지 않은 내려받기는 바이트가 같아 매크로가 남는다. 손실은 고칠 때만 생기므로 "고치기 전에" 알리면 충분하다. 열지 않거나 확인 창으로 막으면 "늘 편집 모드" 규칙과 부딪친다 |
| A6 | 바꾸기에 실패한 옛 파일(손상·암호)은 업로드를 422로 거절하고 이유를 보인다. 열람 화면으로 내려가지 않는다 | 규칙 1. 암호·손상 파일은 열람 변환도 대개 실패한다 |
| A7 | 향기랩 페이지는 `editorUrl`이 없으면 열람 화면 대신 오류를 보인다(`editorUrl \|\| viewerUrl` → `editorUrl`만) | 규칙 1을 페이지에서도 지킨다. 지금 경로(docx·doc·xlsx·xls·pptx·ppt·pdf)는 모두 `editorUrl`을 받는다 |
| A8 | 이미지: cadview가 **빈 도면 틀**(`blank.dwg`, 저장소에 넣는 작은 R2000 DWG)을 이 탭의 wasm 편집 세션으로 열고 그 위에 그림을 원점·축척 1(1 px = 1 단위)로 붙인다(`add_image`). 래스터 명령은 "이미지가 하나뿐이면 그것"(`onlyImage`) 규칙으로 바로 그 그림을 고친다. 켜는 표시는 부모 메시지 `cadview:open`의 `imageEdit: true`(향기랩만 보낸다) | 래스터 편집·저장(RSAVE)·실행 취소·문자 인식(OCR 명령)이 모두 "도면 안 IMAGE" 위에 이미 있다. 엔진에 "도면 없는 그림 세션"을 새로 만드는 것보다 작다. 엔진에는 빈 도면을 처음부터 만드는 기능이 없어(`Document::open`은 바이트에서만) 틀 파일이 필요하다 |
| A9 | 빈 도면 틀은 우리가 만든 빈 DXF(ezdxf, MIT)를 LibreDWG `dxf2dwg`(0.14)로 R2000 DWG로 바꿔 만든다. 시험 자료 도면(LibreDWG 저장소의 example_*.dwg)에서 지워 만들지 않는다 | 출처가 깨끗한 우리 파일. R2000(AC1015)은 편집 가능 판이고 이미지를 넣을 수 있다(registry: R13·R14만 불가) |
| A10 | 이미지 편집 모드의 저장(Ctrl+S·SAVE)은 그림만 내려받는다(도면 .dwg는 쓰지 않음). 문구에서 "도면 옆 같은 이름 파일로 바꾸라"는 꼬리를 뺀다 | 사용자는 그림을 열었다 — 빈 도면은 보이지 않는 그릇일 뿐 |
| A11 | WEBP: 엔진이 읽고 쓰지 못한다. 브라우저에서 PNG로 풀어 편집하고, 저장 때 브라우저가 WEBP로 다시 쓸 수 있으면(`convertToBlob` 결과가 `image/webp`) `.webp`로, 못 쓰면(사파리) `.png`로 내려받고 그렇게 알린다 | 원래 형식 규칙을 될 수 있는 데까지 지킨다. 브라우저 WEBP 인코딩은 손실 압축이다(알려진 한계) |
| A12 | 향기랩 `cadview:replace`(뷰어 안에 새 파일을 놓음)는 `openFile`(확장자 경로표)로 보낸다 | 그래야 NWD를 뷰어 안에 끌어 놓아도 받지 않고, 워드 파일은 webviewer로 간다 |

## 1. 형식별 경로(향기랩)

| 확장자 | 여는 곳 | 편집 | 내려받기 |
|---|---|---|---|
| docx·xlsx·pptx | webviewer 편집기 | 그대로 | 같은 형식 |
| **doc·xls·ppt** | webviewer — 업로드 때 OOXML로 바꿔 편집기 | 워드·엑셀·PPT 편집기 | **원래 형식**(고치지 않았으면 같은 바이트) |
| pdf | webviewer PDF 편집기(`pdfEditor:true`) | 그대로 | pdf |
| hwp·hwpx | `/rhwp/` 미러 | 그대로 | 같은 형식 |
| dwg | `/cadview/` 브라우저 안 | 그대로(§5) | dwg |
| **png·jpg·jpeg·bmp·tif·tiff·webp** | `/cadview/` 브라우저 안 — 빈 도면 + 그림, 래스터 편집 | 래스터 명령(리본 래스터 탭) | **원래 형식**(webp는 A11) |
| ~~nwd~~ | **받지 않는다**(지원하지 않는 형식 안내) | — | — |

## 2. webviewer: 옛 오피스 형식

### 2.1 업로드(`ingestUpload`)

- `POST /api/documents/stream?legacyEdit=1`(스키마 `enum: ['1']`). 조각 업로드는 조각마다 같은 값을 붙인다(마지막 조각이 처리한다). multipart `POST /api/documents`에는 더하지 않는다(쓰는 곳이 없다).
- `legacyEdit`이고 확장자가 doc·xls·ppt면: `convertOfficeFile(바이트, 'doc', 'docx')` 등 → 성공하면 원본 = 바꾼 OOXML, `ext` = docx/xlsx/pptx, `meta.name`은 그대로, `store.saveDerived(id, 'legacy', 올린 바이트)`, `meta.legacy = { ext, sha256(바꾼 OOXML), hasMacros }`. 실패하면 422(A6).
  - (2026-11-13 최종 점검 반영) 들어온 바이트가 ZIP이면 바꾸지 않는다: `[Content_Types].xml` 주 파트로 docx/xlsx/pptx를 정해(.doc 이름의 xlsx는 엑셀 편집기) 그 편집기로 열고, `meta.legacy`(올린 바이트·원래 확장자)는 남겨 안 고친 내려받기는 같은 바이트·같은 이름. 고친 내려받기는 종류가 맞으면(.doc 이름의 docx 등) 옛 형식과 똑같이 원래 확장자로 다시 바꾸고, 종류가 다르면(.doc 이름의 xlsx 등) 편집본을 제 OOXML 이름으로 주며 `/api/view/info`에 `legacy`를 싣지 않는다(옛 형식 문구 없음).
  - ZIP이 아니면(OLE·RTF·HTML·XML 2003 등) LibreOffice로 바꾸고, 결과가 `assertZipNotTooLarge`를 통과한 올바른 OOXML일 때만 받는다. 정말 바꾸지 못한 것만 422.
- 그 뒤는 OOXML이 올라온 것과 같다: docx·pptx는 `editReady`, xlsx는 메모리 저장 모드(향기랩)에서 `editReady`.
- `legacyEdit`이 없으면 지금과 **완전히 같다**: .doc·.ppt는 PDF 열람(편집기 주소 없음), .xls는 지금처럼 xlsx로 정규화하고 이름도 `.xlsx`로.

### 2.2 바꾸기 함수

`src/converters/libreOffice.ts`에 `convertOfficeFile(data, fromExt, toExt)` — 지금 `convertLegacySpreadsheetToXlsx`와 같은 꼴(프로필 슬롯 + `soffice --convert-to <필터>`). 필터: `docx:MS Word 2007 XML`, `doc:MS Word 97`, `xlsx:Calc MS Excel 2007 XML`, `xls:MS Excel 97`, `pptx:Impress MS PowerPoint 2007 XML`, `ppt:MS PowerPoint 97`. 상주 변환기(loWorker.py)는 PDF 내보내기만 하므로 쓰지 않는다 — 한 번에 soffice 새로 띄우기(수 초). 느리면 상주 변환기에 `storeToURL` 필터를 더하는 것이 다음 단계(코드에 `shortcut:` 주석). 기존 xls 정규화 함수는 다른 호출자 경로라 손대지 않는다.

### 2.3 매크로 찾기(`hasVbaMacros`)

- 워드·엑셀: 복합 파일(OLE, 머리 `D0 CF 11 E0`)의 디렉터리 항목 중 종류 1(저장소)이고 이름이 `Macros`(워드)·`_VBA_PROJECT_CUR`(엑셀)·`VBA`인 것.
- PPT: VBA는 `PowerPoint Document` 흐름 안의 압축 OLE 객체라 디렉터리에 없다 — `VBAInfoContainer`(recType 0x03FF, recVer 0xF, 길이 20) 바로 뒤 `VBAInfoAtom`(0x0400, 길이 12) 레코드 머리를 찾는다.
- 2026-11-12 시험(Apache POI test-data): `SimpleMacro.doc/.xls/.ppt` 셋 다 찾음, 매크로 없는 표본 12개 중 11개 없음, `44010-TwoCharts.xls`는 실제로 `_VBA_PROJECT_CUR` 저장소가 있어 찾음(빈 VBA 프로젝트로 보임 — 경고 한 줄이 더 뜰 뿐 해는 없다).

### 2.4 내려받기(`GET /api/view/download`)

마크업 PDF 분기 뒤, `meta.legacy`가 있으면 `legacyDownload(meta)`:

1. 고침 판단(A3): `await hasUnsavedEdits(id)`(room.ts 새 함수 — 방이 있으면 `room.ready`를 기다린 뒤 `editedSinceSave`) 또는 `sha256(저장된 OOXML) !== legacy.sha256`.
2. 안 고침 → 파생 `legacy` 바이트 그대로, 이름 `meta.name`(`보고서.doc`).
3. 고침 → `getCurrentBytes(id) ?? 저장된 OOXML` → `convertOfficeFile(…, ext, legacy.ext)`, 이름 `meta.name`.
4. 3이 실패 → 편집본 그대로, 이름 `보고서.docx`(A4), 서버 기록에 남김.
5. (2026-11-13) 3의 결과는 `옛 확장자:sha256(지금 OOXML)`로 메모리 캐시(64MB, LRU) — 같은 상태를 거듭 받으면 LibreOffice는 한 번만. 캐시에 없어 LibreOffice를 실제로 띄울 때만 `HEAVY_RATE_LIMIT`(IP당 분당 20)을 세고, 넘으면 429.

감사 기록 `downloaded`의 detail에 `legacy original|converted|fallback`. 인쇄(`/api/view/print`)·쪽 그림은 지금처럼 OOXML에서 만든다.

### 2.5 편집기 화면

- `/api/view/info`에 `legacy: { ext, hasMacros }`(옛 형식일 때만).
- `officeShell`(워드·엑셀·PPT 공용) 옵션 `legacy`: 파일 → 정보의 "형식"은 `Word 97-2003 문서` 등, 다운로드 칸 문구는 "원래 형식(.doc)으로 내려받습니다. 고치지 않았으면 올린 파일 그대로이고, 고쳤으면 LibreOffice로 다시 바꾸므로 모양이 조금 달라질 수 있습니다." 단추 "DOC 파일로 다운로드". `hasMacros`면 리본 아래 경고 줄(A5, 기존 `.unsupported-warning` 모양).
- 워드 편집기의 기존 경고("이미지 등 지원하지 않는 요소 — 저장하면 사라짐")는 바꾼 docx에도 그대로 적용된다.

## 3. 알려진 한계(사용자에게 알린 것)

같은 LibreOffice 그리기로 앞뒤를 비교한 시험(MS Office로 만든 Apache POI 표본 12개, doc·xls·ppt 각 4) 결과, 손대지 않고 한 바퀴(옛 → OOXML → 옛) 돌렸을 때:

- 보통 쪽 그림 화소 차이 0~2%.
- `.doc` 떠 있는 그림: 품은 그림 56 → 40개.
- `.ppt` 잘라 쓴 비트맵: 7.9% 차이.
- 복잡한 `.xls` 인쇄 배치: 16 → 14쪽, 11% 차이. `.xls` 차트 글자 239 → 208자.
- 워드아트 도형: 약 2% 어긋남.
- 시험하지 않은 것: 한글 글꼴, 매크로 실제 동작, SmartArt, 변경 내용 추적, 메모, 피벗.

여기에 편집기 자체 한계(워드 편집기가 지원하지 않는 요소는 저장 때 빠짐 등)가 겹친다. 그래도 **고치지 않으면 손실이 없다**(같은 바이트) — 이것이 사용자가 원래 형식을 고른 이유이므로 내려받기 기본값을 OOXML로 바꾸지 않는다.

## 4. cadview: 이미지를 래스터 편집으로

### 4.1 지금 있는 것(확인함)

- browser-local에서 이미지만 들어오면 `openDoc` → 문자 인식 이미지 보기(`ocr/imageview.ts`, 2D 캔버스, 편집 없음).
- 편집 백엔드 `LocalBackend`(`edit/backend-local.ts`): 이 탭 wasm의 `cv_edit_*`. 원본 = `localFile`(도면 바이트), 그림 바이트 = `localImages`(메모리). `rasterSave`는 바뀐 그림을 이름 그대로 `showSaveFilePicker`/내려받기로 쓴다(형식은 이름 확장자, JPG는 손실 경고).
- `IMAGEATTACH`의 결과 op는 `addImageOp(at, path, [w, h], scale)`(`edit/tools.ts`), 축척 1이면 1 px = 1 단위.
- 래스터 명령은 이미지가 하나뿐이면 고르지 않고 그것을 연다(`rasterCmd` → `ctx.onlyImage`).
- `RNEW`는 열린 도면에 새 빈 그림을 더한다 — 바깥 그림을 여는 데는 맞지 않다.
- 엔진에는 빈 도면을 처음부터 만드는 길이 없다 → 틀 파일(A8·A9).
- 엔진 래스터는 PNG·JPG·BMP·TIF만(WEBP 없음, A11).

### 4.2 흐름(`imageEdit`)

1. 향기랩이 `{type:'cadview:open', file, imageEdit:true}`를 보낸다(이미지일 때만). `acceptOpenMessage`가 `imageEdit`을 돌려준다(없으면 false — 다른 host·지금 동작 그대로).
2. 시작 화면에서 이미지 + `imageEdit`이면 이미지 보기 대신: (WEBP면 PNG로 풀어 `이름.png` File로) `localImages.add([그림])` → 빈 도면 틀(`blank.dwg`, 묶음에 바이트로 들어 있음)을 `open()`으로 변환·그리기 → 그린 뒤 `session.ensure()` → `session.apply([addImageOp([0,0], 그림 이름, [w,h], 1)], …)` → 카메라를 그림 사각형에 맞춤 → 리본 "래스터" 탭.
3. 문서 제목·`localName`은 그림 이름. 저장(Ctrl+S·SAVE·RSAVE)은 `LocalBackend`의 `imageOnly`로 그림만(A10). WEBP는 `saveAs` 훅이 다시 WEBP로(A11).
4. 문자 인식은 리본의 OCR 명령(도면 이미지용, 이미 있음)으로.
5. 래스터 상한을 넘는 큰 그림은 지금 엔진 문구("이미지가 너무 큽니다")가 뜨고 그림은 보이기만 한다 — 상한은 바꾸지 않는다.

### 4.3 바뀌지 않는 것

데스크톱(server 모드)·8020(server·browser-remote)은 `cadview:open` 메시지를 받지 않으므로 그대로. browser-local 묶음의 자체 host(index.html)도 `imageEdit`을 보내지 않아 이미지 보기 그대로.

## 5. DWG 확인

browser-local은 `editAllowed`가 늘 참(`main.ts` — `mode !== 'browser-remote'`)이고 편집 세션은 첫 편집 명령에서 열린다. "편집 켜기" 단계가 없으므로 이미 편집 모드다. 커밋 안 된 `editOff()`(`edit=off` 조회)가 있어도 향기랩은 그 값을 보내지 않는다. 마지막 검증에서 실제로 선 긋기 → 저장으로 확인한다.

## 6. 향기랩 페이지(calc-site)

- `functions/api/docview/upload.js`: 위로 보낼 때 늘 `legacyEdit=1`(이 함수는 향기랩만 쓴다).
- `tools/index.html`: `ROUTES`·`accept`에서 nwd 삭제, 이미지 `cadview:open`에 `imageEdit: true`, `cadview:replace` → `openFile`(A12), `editorUrl`만(A7). 지원 형식 문구·안내 표(워드 docx·doc, 엑셀 xlsx·xls, PPT pptx·ppt, 도면 dwg "보기·편집, 길이 재기, 인쇄", 이미지 줄 새로)와 영어 문구(`i18n-en.js`).
- 이 가지는 `cadview-ocr` 위(새 가지 `always-edit`). `cadview-ocr`은 master보다 뒤처져 있다(그림 지도 등) — 배포 때 master에 합치며 맞춘다(이 작업 밖).

## 7. 통합 보기 설계 §8 대체

`2026-11-12-unified-view-shell-design.md` §8의 "향기랩에서도 편집기가 없는 .doc·.ppt는 새 보기 화면으로 열린다"는 이 설계로 대체한다: 향기랩의 .doc·.xls·.ppt는 업로드 때 OOXML로 바꿔 편집기로 연다. 새 보기 화면으로 열리는 .doc·.ppt는 `legacyEdit`을 보내지 않는 연동 제품뿐이다.

## 8. 시험

- 단위(webviewer vitest): `hasVbaMacros`(POI 표본 6개), `convertOfficeFile` 왕복(LibreOffice 있을 때), 업로드 `legacyEdit` 있음/없음(메타·이름·파생·편집기 주소), 내려받기 안 고침 = 같은 바이트 / 고침 = OLE 머리 + `.doc` 이름 + 고친 글자 / 바꾸기 실패 = `.docx`.
- 단위(cadview): `acceptOpenMessage`의 `imageEdit`, `LocalBackend` `imageOnly` 저장·`saveAs`, 엔진 시험(edit-core)으로 빈 도면 틀이 편집 가능·그림 붙이기·래스터 열기·저장이 되는지.
- e2e(webviewer): `scripts/verify-legacy-edit-e2e.mjs`(`npm run verify:legacy-edit`) — doc·xls·ppt 업로드·편집·내려받기, `legacyEdit` 없는 업로드가 지금과 같은지.
- 실제 브라우저(향기랩 로컬, 임시 포트): doc·xls·ppt 편집 + 원래 형식 내려받기 + 안 고친 내려받기 바이트 같음, 매크로 경고 줄, 이미지 래스터 편집 + 저장(png·jpg·tif·webp), 문자 인식 리본, DWG 선 긋기·저장, nwd 거절(고르기·뷰어 안에 놓기 둘 다). 회귀: `verify:view-shell`·`verify:view-only`·`verify:xlsx`·`verify:pptx`.
- 시험 포트는 임시 포트만 — 8000·8010·8020·8100·8110 금지(8130도 운영 게임 서버라 피한다).

## 9. 하지 않는 것

- odt·ods·odp, rtf 등 다른 옛 형식(향기랩 경로에 없다).
- 상주 변환기로 옛 형식 바꾸기 빠르게 하기(느리면 다음 단계).
- 매크로 보존(.docm 등으로 바꾸기) — 사용자가 원래 형식만 원했다.
- cadview 판 올리기·향기랩 `/cadview/` 묶음 커밋·배포·master 병합(모두 사용자 승인 뒤 따로).
