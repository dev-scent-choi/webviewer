# DocView 기술 명세 및 연동 가이드

DocView는 **독립 실행형 문서 뷰어 서버**입니다. 업무 시스템(그룹웨어, EDMS, 포털 등)은
REST API로 문서를 등록하고 열람 토큰을 발급받은 뒤, **뷰어 URL을 새 창이나 iframe으로
열기만 하면** 열람·출력·다운로드·복사 권한이 모두 서버에서 강제된 상태로 동작합니다.

```
[업무 시스템]                         [DocView 서버]
    │ ① POST /api/documents (파일)        │
    │ ──────────────────────────────────► │ 저장 + 백그라운드 변환
    │ ② GET /api/documents/:id (폴링)      │
    │ ◄────────────────────────────────── │ status: ready
    │ ③ POST /api/documents/:id/tokens     │
    │    {viewer, canPrint, ...}           │
    │ ◄────────────────────────────────── │ { token, viewerUrl }
    │ ④ 사용자 브라우저에서 viewerUrl 오픈  │
    │    (새 창 또는 iframe)               │ 토큰 검증 후 페이지/그리드 서빙
```

- 원본 파일은 브라우저로 전송되지 않습니다. 페이지 이미지(PNG) 또는 표시용 셀 데이터만 전달됩니다.
- 출력은 열람자 이름·날짜 워터마크가 삽입된 PDF로만 제공되며, 다운로드는 `canDownload`
  권한이 있는 토큰만 가능합니다.
- 모든 열람/출력/다운로드/거부 이벤트는 감사 로그(`data/audit/audit.jsonl`)에 기록됩니다.

---

## 1. 설치 및 실행

**요구 사항**: Node.js 20 이상 (`node -v`로 확인).

```bash
# 1) 소스 받기
git clone https://github.com/dev-scent-choi/webviewer.git
cd webviewer                 # 저장소 루트가 프로젝트 폴더(package.json 위치)

# 2) 의존성 설치
npm install

# 3) 실행
npm run dev                  # 개발 모드 (파일 변경 자동 반영, 기본 http://127.0.0.1:8000)
npm start                    # 운영 모드
```

- 실행 후 브라우저에서 `http://<호스트>:<포트>` 접속 → 업로드·열람 화면.
- 외부에서 접속하게 하려면 `DOCVIEW_HOST=0.0.0.0`으로 바인딩한다 (아래 환경변수 참고).
- 운영 배포 시에는 `DOCVIEW_SECRET`을 반드시 고정값으로 설정한다 (미설정 시 재시작마다 발급된 토큰이 전부 무효화됨).
- Windows에서 상시 구동하려면 `npm start`를 pm2/NSSM 등 프로세스 관리자로 감싸 서비스로 등록한다.
- 렌더링·파싱 캐시는 `DOCVIEW_CACHE_BUDGET_MB`(기본 512) 안에서 관리된다. 고배율 도면/대형
  시트를 많이 동시 열람하는 환경이면 이 값과 Node 힙(`NODE_OPTIONS=--max-old-space-size=2048`)을
  함께 올린다. (네이티브 이미지 렌더링 메모리 때문에 실제 사용량은 예산보다 다소 크다.)

### 환경변수

| 변수 | 기본값 | 설명 |
|------|--------|------|
| `DOCVIEW_SECRET` | (없음 — 부팅 시 랜덤) | JWT 서명 시크릿. **운영 필수** (미설정 시 재시작마다 토큰 전체 무효화) |
| `DOCVIEW_PORT` | `8000` | 포트 |
| `DOCVIEW_HOST` | `127.0.0.1` | 바인딩 호스트. 외부 연동 시 `0.0.0.0` |
| `DOCVIEW_DATA_DIR` | `./data` | 저장 경로 (disk 모드) |
| `DOCVIEW_STORAGE` | `disk` | `disk` 파일 저장 / `memory` 메모리 보관(디스크 비잔류, 재시작 시 소멸) |
| `DOCVIEW_MEMORY_BUDGET_MB` | `1024` | memory 모드 총 용량. 초과 시 오래된 문서부터 제거 |
| `DOCVIEW_CACHE_BUDGET_MB` | `512` | 렌더링·파싱 캐시(페이지 이미지/그리드/PDF 텍스트) 총 상한. 초과 시 LRU 제거 |
| `DOCVIEW_TOKEN_TTL` | `3600` | 열람 토큰 유효기간(초) |
| `DOCVIEW_SOFFICE` | (자동 감지) | LibreOffice 경로 |
| `DOCVIEW_DWG2DXF` | `tools/dwg2dxf.exe` | LibreDWG 변환기 경로 |
| `DOCVIEW_ODA_CONVERTER` | (자동 감지) | ODA File Converter 경로. 설치 시 DWG 변환에 우선 사용 |

### 외부 변환 도구

| 도구 | 용도 | 설치 |
|------|------|------|
| LibreOffice | Office/스프레드시트(xls·ods)/DXF → PDF | https://libreoffice.org (필수에 가까움) |
| LibreDWG | DWG → DXF | [릴리스 zip](https://github.com/LibreDWG/libredwg/releases)을 `tools/`에 압축 해제 |
| ODA File Converter | DWG 전 버전 고충실 변환 | https://opendesign.com 수동 설치 (권장) |

---

## 2. 지원 포맷 (23종)

| 그룹 | 확장자 | 뷰 방식 | 비고 |
|------|--------|---------|------|
| PDF | pdf | 페이지 이미지 | 무변환 |
| 이미지 | png, jpg, jpeg | 페이지 이미지 | |
| 이미지(확장) | tif, tiff, bmp, webp, gif | 페이지 이미지 | 다페이지 TIFF는 페이지 분리 |
| 스프레드시트 | **xlsx, csv** | **그리드 뷰** | 엑셀형 시트 탭+가상 스크롤. 수십만 행 지원. PDF는 출력 시 지연 생성 |
| 스프레드시트 | xls, ods | 페이지 이미지 | LibreOffice 변환 |
| 문서 | doc, docx, ppt, pptx, odt, odp, rtf, txt | 페이지 이미지 | LibreOffice 변환 |
| 도면 | dxf | 페이지 이미지 | 여백 자동 크롭 |
| 도면 | dwg | 페이지 이미지 | ODA(있으면) → LibreDWG 폴백. LibreDWG 단독은 AutoCAD 2000/2004 위주 |

`GET /api/formats`가 현재 서버에서 활성화된 확장자 목록을 반환합니다.

---

## 3. API 레퍼런스

모든 응답은 `{ "success": boolean, "data"?: ..., "error"?: string }` 형식입니다.

### 3.1 문서 등록

```
POST /api/documents
Content-Type: multipart/form-data (필드명: file)
```

```bash
curl -F "file=@계약서.xlsx" http://viewer-server:8000/api/documents
```

응답 `201`:

```json
{
  "success": true,
  "data": {
    "id": "9b13ce91-961d-4b22-857d-6189d338128c",
    "name": "계약서.xlsx",
    "ext": "xlsx",
    "status": "processing",
    "pageCount": 0,
    "createdAt": "2026-07-11T02:00:00.000Z",
    "progress": 10
  }
}
```

업로드는 즉시 반환되고 변환은 백그라운드로 진행됩니다. 오류: `400`(파일 없음),
`413`(100MB 초과), `415`(미지원 형식).

### 3.2 변환 상태 폴링

```
GET /api/documents/:id
```

`status`가 `processing` → `ready` 또는 `failed`로 바뀔 때까지 1초 간격 폴링을 권장합니다.
`progress`(0~100)로 게이지를 표시할 수 있고, 실패 시 `error`에 사유가 담깁니다
(예: `"DWG 변환에 실패했습니다 (AutoCAD 2013 형식). ODA File Converter를 설치하면 지원됩니다."`).
`viewKind`는 `pages`(페이지 이미지) 또는 `grid`(스프레드시트 그리드)입니다.

### 3.3 문서 목록

```
GET /api/documents          → { success, data: DocumentMeta[] }
GET /api/formats            → { success, data: ["pdf","xlsx",...] }
```

### 3.4 열람 토큰 발급 (권한 지정)

```
POST /api/documents/:id/tokens
Content-Type: application/json
```

```json
{ "viewer": "홍길동", "canPrint": true, "canDownload": false, "canCopy": false }
```

| 필드 | 필수 | 설명 |
|------|------|------|
| `viewer` | ✓ | 열람자 표시 이름. 워터마크·감사 로그에 기록 |
| `canPrint` | | 워터마크 PDF 출력 허용 (기본 false) |
| `canDownload` | | 원본 다운로드 허용 (기본 false) |
| `canCopy` | | 그리드 뷰 셀 선택·복사 허용 (기본 false) |

응답:

```json
{
  "success": true,
  "data": {
    "token": "eyJhbGciOi...",
    "viewerUrl": "/viewer.html?token=eyJhbGciOi...",
    "expiresInSeconds": 3600
  }
}
```

문서가 `ready`가 아니면 `404`. 토큰은 JWT이며 만료 시 재발급이 필요합니다.

### 3.5 뷰어 열기 (연동 시스템에서 할 일의 전부)

```html
<!-- 새 창 -->
<a href="http://viewer-server:8000/viewer.html?token=..." target="_blank">문서 보기</a>

<!-- iframe 임베드 -->
<iframe src="http://viewer-server:8000/viewer.html?token=..."
        style="width:100%;height:90vh;border:0"></iframe>
```

뷰어가 토큰의 권한에 따라 출력/저장/복사 버튼을 자동으로 활성·비활성화하며,
버튼을 우회해 API를 직접 호출해도 서버가 401/403으로 차단합니다.

### 3.6 뷰어 데이터 API (뷰어 내부에서 사용 — 직접 연동도 가능)

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/api/view/info?token=…` | `{name, pageCount, viewKind, permissions, viewer}` |
| GET | `/api/view/pages/:n?token=…&scale=1~16` | 페이지 PNG. `scale`은 해상도 배율(뷰어가 확대 수준·디스플레이 밀도에 맞춰 자동 요청, 페이지당 최장변 6000px 상한) |
| GET | `/api/view/grid/meta?token=…` | 그리드 문서의 시트 목록·열 너비·병합·스타일 팔레트 |
| GET | `/api/view/grid/rows?token=…&sheet=0&from=1&count=200` | 행 데이터 범위 조회 (최대 500행/요청) |
| GET | `/api/view/search?token=…&q=…&limit=200` | 문서 내 문자 검색 (OCR 아님 — 그리드 셀 값 / PDF 텍스트 레이어). grid: `{kind:'grid', hits:[{sheet,r,c,snippet}]}` / pages: `{kind:'pages', hits:[{page, rects:[{x,y,w,h}(비율 0~1)], snippet}]}`. 뷰어에서 Ctrl+F 또는 돋보기 버튼 |
| GET | `/api/view/print?token=…` | 워터마크 PDF (`canPrint` 필요, 403) |
| GET | `/api/view/download?token=…` | 원본 파일 (`canDownload` 필요, 403) |

오류 코드: `401` 토큰 없음/만료/위조, `403` 권한 없음, `404` 문서/페이지/시트 없음.

---

## 4. 권한·보안 모델

- **원본 비전송**: 클라이언트는 페이지 PNG 또는 표시용 셀 JSON만 받습니다. 원본과
  변환 PDF는 서버에만 존재합니다.
- **토큰 단위 권한**: 출력/다운로드/복사가 열람자·문서별 JWT에 서명되어 담기며,
  모든 강제는 서버 API에서 수행됩니다 (뷰어 UI 차단은 보조 수단).
- **출력 워터마크**: `viewer` 이름 + 날짜가 모든 페이지에 삽입된 PDF만 제공됩니다.
- **감사 로그**: `uploaded / converted / token_issued / opened / printed / print_denied /
  downloaded / download_denied / conversion_failed` 이벤트가 JSONL로 기록됩니다.
- **알려진 한계**: 화면 캡처는 웹 기술로 차단할 수 없습니다(상용 DRM 동일).
  관리 API(업로드·토큰 발급)에는 인증이 없으므로 **신뢰된 내부망에서만 노출**하거나
  게이트웨이 뒤에 두어야 합니다.

## 5. 스토리지 모드

| | disk (기본) | memory |
|---|---|---|
| 원본/변환본 | `data/` 디렉터리 | 서버 메모리 (변환 순간의 임시 파일만 생성 후 즉시 삭제) |
| 서버 재시작 | 유지 | 전부 소멸 |
| 용량 관리 | 디스크 | `DOCVIEW_MEMORY_BUDGET_MB` 초과 시 오래된 문서 자동 제거 |
| 용도 | 일반 운영 | 원본 비잔류가 필요한 보안 환경 |

저장 계층은 `DocumentStore` 인터페이스(`src/storage.ts`)로 추상화되어 있어
프로그래밍 방식으로도 사용할 수 있습니다.

## 6. 성능 특성

| 작업 | 소요 (측정 기준) |
|------|------|
| PDF/이미지 열람 | 즉시 |
| Office 문서 변환 | 1~3초 (프로필 웜업 후) |
| xlsx 그리드 추출 | 27만 행 파일 기준 약 20초, 이후 열람 즉시 |
| DWG 변환 | 수 초 (도면 복잡도에 따라) |
| 페이지 렌더링 | 확대 수준에 맞춰 배율 1~16 동적 렌더링, 뷰어는 화면 근처 페이지만 로딩 |
