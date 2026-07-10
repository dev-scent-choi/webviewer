# DocView — 권한 제어형 문서 웹뷰어

모든 입력 포맷을 **서버에서 PDF로 정규화 → 페이지 이미지로 렌더링**하여 브라우저에는
원본 파일이 절대 전달되지 않는 구조의 문서 뷰어입니다. 출력/저장 권한은 서버가 발급하는
단기 JWT에 담겨 API 레벨에서 강제됩니다.

## 실행

```bash
npm install
npm run dev        # http://127.0.0.1:8000
```

### 외부 변환 도구 (선택)

- **Office/HWP/DXF/스프레드시트**: [LibreOffice](https://www.libreoffice.org/) 설치 필요 (미설치 시 해당 포맷 자동 비활성화)
- **DWG**: [LibreDWG 릴리스](https://github.com/LibreDWG/libredwg/releases)에서 win64 zip을 받아
  실행 파일들을 `tools/` 폴더에 풀어 넣는다 (`tools/dwg2dxf.exe`가 있으면 자동 활성화,
  또는 `DOCVIEW_DWG2DXF` 환경변수로 경로 지정)

환경변수:

| 변수 | 기본값 | 설명 |
|------|--------|------|
| `DOCVIEW_SECRET` | (없음 — 부팅 시 랜덤 생성) | JWT 서명 시크릿. **운영에서는 반드시 설정** (미설정 시 재시작마다 토큰 전체 무효화) |
| `DOCVIEW_PORT` | `8000` | 서버 포트 |
| `DOCVIEW_HOST` | `127.0.0.1` | 바인딩 호스트 |
| `DOCVIEW_DATA_DIR` | `./data` | 원본/변환본/메타/감사로그 저장 경로 |
| `DOCVIEW_TOKEN_TTL` | `3600` | 열람 토큰 유효기간(초) |
| `DOCVIEW_STORAGE` | `disk` | `disk`: 파일 저장 / `memory`: 서버 메모리에만 보관 — 디스크 비잔류, 재시작 시 소멸 (변환 순간의 임시 파일은 예외적으로 생성 후 즉시 삭제) |
| `DOCVIEW_MEMORY_BUDGET_MB` | `1024` | memory 모드 총 보관 용량. 초과 시 오래된 문서부터 자동 제거 |
| `DOCVIEW_ODA_CONVERTER` | (자동 감지) | ODA File Converter 경로. 설치되어 있으면 DWG 변환에 우선 사용 (전 버전 지원·최고 충실도) |

검증:

```bash
npm run typecheck
npm test
```

## 아키텍처

```
업로드 → Converter 플러그인 (포맷 → PDF 정규화)
       → data/converted/{id}.pdf   (뷰어는 이것만 접근)
       → data/originals/{id}.{ext} (download 권한 있을 때만 제공)

열람   → POST /api/documents/:id/tokens 로 권한 담은 JWT 발급
       → 뷰어는 페이지 PNG API만 호출 (원본/PDF 직접 접근 불가)
       → 출력: 서버가 열람자 워터마크 입힌 PDF를 즉석 생성
       → 모든 열람/출력/다운로드/거부 이벤트를 data/audit/audit.jsonl에 기록
```

### API

| 메서드 | 경로 | 설명 |
|--------|------|------|
| POST | `/api/documents` | multipart 업로드. 즉시 201(`status: processing`) 응답 후 백그라운드 변환 |
| GET | `/api/documents/:id` | 단일 문서 상태 조회 (변환 진행 폴링용, `processing`/`ready`/`failed`) |
| GET | `/api/documents` | 문서 목록 |
| POST | `/api/documents/:id/tokens` | `{viewer, canPrint, canDownload}` → 열람 토큰 발급 |
| GET | `/api/view/info?token=…` | 문서 정보 + 권한 |
| GET | `/api/view/pages/:n?token=…&scale=1~4` | 페이지 PNG 렌더링. 뷰어가 확대 수준에 맞춰 배율을 자동 요청 (페이지당 최장변 6000px 상한) |
| GET | `/api/view/print?token=…` | 워터마크 PDF (canPrint 필요) |
| GET | `/api/view/download?token=…` | 원본 파일 (canDownload 필요) |

### 지원 포맷

| 그룹 | 확장자 | 변환 엔진 |
|------|--------|-----------|
| PDF | pdf | 통과 (변환 없음) |
| 이미지 | png, jpg, jpeg | pdf-lib |
| 이미지(확장) | tif, tiff, bmp, webp, gif | sharp — 다페이지 TIFF는 페이지별로 분리 |
| 문서 | doc, docx, ppt, pptx, odt, odp, rtf, txt | LibreOffice headless (미설치 시 자동 비활성화) |
| 스프레드시트 | xls, xlsx, ods, csv | LibreOffice Calc — `SinglePageSheets` 옵션으로 시트 1장 = PDF 1페이지 (인쇄 폭 분할로 화면이 조각나는 문제 방지). xlsx는 변환 전 고정 행 높이를 제거해 Excel↔LibreOffice 글꼴 폭 차이로 인한 텍스트 잘림을 방지 |
| 도면 | dxf | LibreOffice Draw |
| 도면 | dwg | ODA File Converter(설치 시 자동 감지, 전 버전 지원) → LibreDWG 순으로 시도 후 LibreOffice Draw. **LibreDWG만으로는 AutoCAD 2000/2004 형식 위주로 동작**하며 R14·2007~2018 형식은 실패할 수 있음 — 실패 시 파일의 AutoCAD 버전과 안내가 오류 메시지에 표시됨 |

`GET /api/formats`가 현재 서버에서 활성화된 확장자 목록을 반환하며, 업로드 화면의
파일 선택 필터도 이 목록을 따른다.

### 아직 미지원 포맷과 이유

- **HWP/HWPX**: LibreOffice에 [H2Orestart](https://github.com/ebandal/H2Orestart) 확장(.oxt)을
  설치하면 Office 변환기와 같은 경로로 지원 가능하다.

## 로드맵

- **HWP**: H2Orestart 확장 설치 후 `officeConverter` 확장자에 hwp/hwpx 추가.
- **DWG 충실도**: 복잡한 실무 도면에서 LibreDWG 변환 품질이 부족하면 ODA File Converter
  (무료, [opendesign.com](https://www.opendesign.com/guestfiles/oda_file_converter) 수동 설치)로 교체.
- **권한 고도화**: 관리자 인증, 문서별 ACL, 감사 로그 조회 화면, 워터마크에 한글 폰트
  임베딩(fontkit), 변환 작업 큐 분리(대용량 대응).

## 보안 설계 노트

- 원본과 변환 PDF는 뷰어 클라이언트가 직접 접근할 수 없고, 토큰 검증을 통과한
  페이지 이미지만 전달됩니다. 파일 경로는 서버가 UUID로만 관리하며 경로 조작을 차단합니다.
- 뷰어의 우클릭/Ctrl+P/Ctrl+S 차단은 **억지력일 뿐**이며, 실제 강제 지점은 서버 API입니다.
- 화면 캡처는 웹 기술로 막을 수 없습니다(상용 DRM도 동일). 출력물 워터마크(열람자+날짜)가
  현실적인 억지 수단입니다.
- 업로드는 확장자 허용 목록 + 100MB 용량 제한으로 검증됩니다. 관리 API(업로드/토큰 발급)에는
  아직 인증이 없으므로 **Phase 4 전까지는 신뢰된 네트워크에서만 사용**해야 합니다.
