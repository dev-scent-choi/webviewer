# 두 가지로 열기 — 아트크래프트(브라우저 엔진) 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 토큰 발급 때 `engine: 'artcraft'`를 고르면 `/view.html`이 서버 변환 없이 브라우저 Worker 안의 ArtCraft wasm 엔진으로 docx·pptx·doc 쪽을 그리고, 안 되면 지금 리브레오피스 경로로 저절로 넘어간다.

**Architecture:** spike 껍데기 크레이트를 `tools/artcraft-shell/`로 옮겨 `npm run build:artcraft`가 해시 붙은 wasm·글꼴·`manifest.json`을 `public/artcraft/`에 만든다. 서버는 `engineFor()` 하나로 토큰 응답과 `/api/view/info`에서 엔진을 정하고(클레임에는 "요청한 엔진"만), 화면은 새 `PageSource`(`ArtcraftSource`) 하나만 더해 `PagesBody`·셸은 그대로 쓴다.

**Tech Stack:** Rust(wasm32-unknown-unknown, GNU 도구 묶음) · Node 20 · Fastify 5 + @fastify/static(preCompressed) · TypeScript · esbuild · Web Worker · vitest · Playwright · sharp

**Spec:** `docs/superpowers/specs/2026-11-21-artcraft-browser-engine-design.md` (근거: `C:\Users\OpenMinds\projects\artcraft-eval\SPIKE-WASM.md` §7·§8, 시험 코드 `...\artcraft-eval\spike-wasm\`)

## Global Constraints

- 형식: docx = WordCraft(된다), pptx = DeckCraft(된다), doc = WordCraft("실험" 표시), ppt·xlsx·xls·hwp·pdf·도면·그 밖 = 안 됨 → 지금 경로. 안 되는 형식에 `engine: 'artcraft'`를 보내면 **거절하지 않고** 지금 경로로 열고 `engine: 'libreoffice'` + `engineReason`.
- 이유 문구(그대로): `'내려받기 권한이 없어 서버 변환으로 엽니다'` · `'이 형식은 브라우저 엔진이 지원하지 않습니다'` · `'브라우저 엔진이 설치되지 않았습니다'`.
- 원본 바이트는 `canDownload` 토큰에만 내려간다. 토큰의 `engine`은 클레임으로만 믿는다(주소 쿼리는 보지 않는다).
- 토큰 응답: `{ token, viewerUrl, editorUrl?, expiresInSeconds, engine, engineReason? }`. `viewerUrl`은 지금처럼 `/view.html?token=…`.
- `/api/view/info` 추가: `engine`, `engineReason?`, `artcraft?: { kind: 'word' | 'deck', ext, experimental, originalUrl, assets: { wasm, fonts[] } }`. `originalUrl` = 지금 있는 `/api/view/download?token=…`.
- 크기 상한(검사로 고정, MB = 1024×1024 바이트): word wasm brotli ≤ 5 MB, deck ≤ 2.5 MB, 글꼴 2개(S-Core Dream 4·6) brotli 합 ≤ 0.6 MB.
- 정적 `/artcraft/*`: `.br`가 있고 `accept-encoding`에 br이 있으면 `.br`(`content-encoding: br`), `cache-control: public, max-age=31536000, immutable`.
- `public/artcraft/manifest.json`이 없으면 서버는 그대로 뜨고 아트크래프트 요청은 전부 libreoffice + `'브라우저 엔진이 설치되지 않았습니다'`.
- 시간 제한: 첫 쪽 그리기 20초(`RENDER_TIMEOUT_MS = 20_000`, 그 뒤 모든 Worker 요청에도 같은 값). 열기(wasm·글꼴·원본 받기 + `open`)는 60초(`OPEN_TIMEOUT_MS = 60_000` — 설계에 값이 없어 이 계획이 정함).
- 전환 규칙(→ 리브레오피스 원천 + 안내 `'브라우저 엔진이 이 문서를 열지 못해 서버 변환으로 보여 줍니다'`): wasm 받기 실패, `open` 실패/오류, 첫 쪽 20초 초과, 쪽 수 0, 글꼴 받기 실패, 열린 뒤 그리기 오류·Worker 멈춤. .doc(실험)은 엔진 쪽 수가 `info.pageCount`(서버 변환이 끝나 있을 때, >0)와 **2쪽 넘게** 다르면 전환.
- 표지 문구: `'브라우저 엔진(아트크래프트)'`(.doc는 `' · 실험'`을 붙임).
- 두 단계 그리기: 그 쪽의 선명한 그림이 아직 없고 요청 화소 폭이 쪽 폭(pt)의 1.5배를 넘으면 먼저 `scale = 1`, 이어서 요청 배율. `qualityKey = ceil(cssWidth × dpr)`. 엔진 호출 `sp_render(h, i, round(scale×1000))`, `scale = 화소 폭 ÷ 쪽 폭 pt`.
- `PagesBody`·셸 모듈(막대·썸네일·찾기·마크업·측정·인쇄·내려받기)은 바꾸지 않는다 — `PageSource.kind`·`BodyKind` 합집합에 `'pages-artcraft'`만 더한다. 인쇄·내려받기는 서버 PDF 그대로.
- 리브레오피스 경로는 바꾸지 않는다 — `verify:docx`·`verify:pptx`·`verify:view-only`·`verify:view-shell`이 그대로 통과해야 한다.
- ArtCraft 저장소 3개는 저장소 밖(`ARTCRAFT_DIR`, 기본 `C:\Users\OpenMinds\projects\artcraft-eval`)에서 `tools/artcraft-shell/ARTCRAFT.lock`의 고정 커밋으로만 쓴다(HEAD가 다르거나 고친 파일이 있으면 빌드 중단). 그 저장소들은 한 줄도 고치지 않는다.
- 글꼴은 S-Core Dream 4·6만, 별칭 등록 안 함(메모리 3배).
- 운영 규칙: 포트 8000·8010·8020·8100·8110·8130·8148은 쓰지도 건드리지도 않는다. 프로세스를 이름으로 끄지 않는다(`taskkill /IM node.exe` 금지) — 직접 띄운 것만 PID로 끈다. `public/artcraft/`·`public/js/` 빌드 산출물은 커밋하지 않는다. push 하지 않는다. 커밋은 경로를 지정해서(`git add <경로>`), 메시지는 한국어, 꼬리말은 실행 세션의 지시를 따른다.
- 고객형 시험 문서(`...webviewer-q-all\.audit-run\test-server\originals`의 docx `62c3881f-247e-48ad-aacb-977010a23ef1`, pptx `8c76f45a-e15a-40ca-b6be-d9ca5ed63137`)는 저장소에 넣지 않는다 — e2e가 실행 때 `ARTCRAFT_E2E_ORIGINALS` 폴더에서 `.audit-run/`으로 복사한다.
- 시험 순서: 태스크마다 단위 시험·빌드·타입 검사만, 브라우저 e2e와 회귀 묶음은 Task 5에서 한 번에.
- 하지 않는 것: 엑셀·한글·.ppt 브라우저 엔진, 브라우저 엔진 편집·인쇄·PDF, 서비스 워커, 글꼴 추가 등록, `release()`(메모리는 64 MB LRU로 묶는다).

---

## 파일 구조

| 파일 | 할 일 |
|---|---|
| `tools/artcraft-shell/ARTCRAFT.lock` (새) | ArtCraft 저장소 3개의 주소·고정 커밋(JSON) |
| `tools/artcraft-shell/abi.rs` (새) | 공통 C ABI(spike `abi.rs` 정리 + `sp_page_w/h`) |
| `tools/artcraft-shell/{word,deck}/src/lib.rs` (새) | 엔진별 껍데기(spike 정리 + `page_size`) |
| `tools/artcraft-shell/{word,deck}/Cargo.toml.in` (새) | 경로 자리표시(`@ARTCRAFT_DIR@`·`@SHELL@`)가 있는 Cargo.toml 틀 |
| `tools/artcraft-shell/{word,deck}/Cargo.lock` (새) | spike 잠금 파일(패키지 이름만 바꿈) |
| `scripts/artcraftBuildLib.ts` (새) | 순수 도우미: 해시 이름, 크기 검사, 잠금 검사, brotli |
| `scripts/build-artcraft.ts` (새) | `npm run build:artcraft` 본체 |
| `src/artcraftManifest.ts` (새) | `manifest.json` 읽기·검사(없거나 틀리면 null) |
| `src/artcraftStatic.ts` (새) | `/artcraft/*` 정적 서빙(.br·immutable) |
| `src/artcraftEngine.ts` (새) | `engineFor()` 판정표, `sourceExtOf()`, `artcraftInfo()` |
| `src/app.ts` · `src/types.ts` · `src/security.ts` · `src/routes/documents.ts` · `src/routes/viewer.ts` (고침) | 정적 등록, 클레임 `engine`, 토큰·info 응답 |
| `src/client/view/artcraftLogic.ts` (새) | 순수: 프로토콜 타입, 배율·두 단계 판단, 찾기 사각형, 글자 층 좌표, LRU, 작업 줄, 바깥 함수 채우기 |
| `src/client/view/artcraftWorker.ts` (새) | Worker: wasm·글꼴·원본 받기, open, render, text |
| `src/client/view/artcraftSource.ts` (새) | `ArtcraftSource implements PageSource`, `WorkerClient`, `openArtcraftSource()` |
| `src/client/view/types.ts` · `pagesBody.ts` · `src/client/view.ts` (고침) | 타입 합집합, 원천 고르기·전환·표지 |
| `scripts/build-editor.mjs` (고침) | Worker 진입점 추가 |
| `public/view.html` · `public/css/view.css` · `public/js/main.js` · `public/css/style.css` (고침) | 표지, 단추·라디오 |
| `scripts/verify-artcraft-e2e.mjs` (새) | 설계 §7 브라우저 검사 7가지 + 메인 화면 + 쪽 수 기록 |
| `tests/artcraftBuild.test.ts` · `tests/artcraftAssets.test.ts` · `tests/artcraftEngine.test.ts` · `tests/artcraftRoutes.test.ts` · `tests/artcraftLogic.test.ts` (새) | 단위 시험 |

---

### Task 1: 껍데기 크레이트 · `build:artcraft` · manifest · 정적 서빙

**Files:**
- Modify: `.gitignore`
- Create: `tools/artcraft-shell/ARTCRAFT.lock`, `tools/artcraft-shell/abi.rs`, `tools/artcraft-shell/word/src/lib.rs`, `tools/artcraft-shell/word/Cargo.toml.in`, `tools/artcraft-shell/word/Cargo.lock`, `tools/artcraft-shell/deck/src/lib.rs`, `tools/artcraft-shell/deck/Cargo.toml.in`, `tools/artcraft-shell/deck/Cargo.lock`
- Create: `scripts/artcraftBuildLib.ts`, `scripts/build-artcraft.ts`
- Create: `src/artcraftManifest.ts`, `src/artcraftStatic.ts`
- Modify: `src/app.ts:120-130` (fastifyStatic 등록 바로 앞), `package.json` (scripts)
- Test: `tests/artcraftBuild.test.ts`, `tests/artcraftAssets.test.ts`

**Interfaces:**
- Consumes: 없음(첫 태스크).
- Produces:
  - wasm C ABI(Worker가 씀): `sp_alloc(n)→ptr`, `sp_free(ptr,n)`, `sp_add_font(ptr,len,aliasPtr,aliasLen)→faces`, `sp_open(namePtr,nameLen,ptr,len)→handle(0=실패)`, `sp_count(h)→n`, `sp_page_w(h,i)→f32 pt`, `sp_page_h(h,i)→f32 pt`, `sp_render(h,i,scaleMilli)→0|1`, `sp_text(h,i)→0|1`, `sp_out_ptr()`, `sp_out_len()`, `sp_out_w()`, `sp_out_h()`, `sp_out_clear()`, `sp_msg_ptr()`, `sp_msg_len()`, `sp_close(h)`. `i`는 0부터. 글자 JSON = `{"w":pt,"h":pt,"runs":[{"t","x0","x1","y"(기준선),"s"}]}`.
  - `public/artcraft/manifest.json` = `{ "word": "/artcraft/artcraft_word-<10hex>.wasm", "deck": "/artcraft/artcraft_deck-<10hex>.wasm", "fonts": ["/artcraft/fonts/SCDream4-<10hex>.otf", "/artcraft/fonts/SCDream6-<10hex>.otf"], "commits": { "wordcraft": "<sha>", "deckcraft": "<sha>" } }`
  - `src/artcraftManifest.ts`: `interface ArtcraftManifest { readonly word: string; readonly deck: string; readonly fonts: readonly string[] }`, `artcraftDir(): string`, `parseManifest(raw: unknown): ArtcraftManifest | null`, `readArtcraftManifest(dir?: string): ArtcraftManifest | null`
  - `src/artcraftStatic.ts`: `registerArtcraftStatic(app: FastifyInstance, root: string): Promise<void>`
  - Worker 번들이 받는 파일 주소: `/js/artcraftWorker.bundle.js`(Task 3이 만든다)

- [ ] **Step 1: `.gitignore` — `tools/`는 그대로 무시하되 껍데기만 넣고, 빌드 산출물은 무시**

`.gitignore`에서 `tools/` 한 줄을 아래로 바꾸고, `public/js/.build.lock` 줄 아래에 산출물 줄을 더한다.

```gitignore
# 외부 변환 도구 바이너리 (README의 설치 안내 참조) — 단 아트크래프트 껍데기 크레이트(소스)는 저장소에 둔다
tools/*
!tools/artcraft-shell/
tools/artcraft-shell/.build/
```

```gitignore
# 아트크래프트 브라우저 엔진 산출물(npm run build:artcraft) — 이름에 내용 해시가 붙어 소스가 아니다
public/artcraft/
```

Run: `git check-ignore -v tools/artcraft-shell/abi.rs; echo "exit=$?"; git check-ignore -v tools/artcraft-shell/.build/x public/artcraft/manifest.json tools/wasm-bindgen/x`
Expected: 첫 줄은 출력 없이 `exit=1`(무시 안 됨), 둘째 명령은 세 경로 모두 무시 규칙을 출력.

- [ ] **Step 2: `ARTCRAFT.lock` 쓰기**

`tools/artcraft-shell/ARTCRAFT.lock`:

```json
{
  "note": "ArtCraft 저장소는 ARTCRAFT_DIR(기본 C:/Users/OpenMinds/projects/artcraft-eval) 아래에 두고 이 커밋에서만 빌드한다. 판을 올리면 커밋을 바꾸고 npm run build:artcraft + npm run verify:artcraft를 다시 돌린다.",
  "repos": {
    "wordcraft": { "url": "https://github.com/storytold/wordcraft.git", "commit": "f8ec3cc0fa884369c4f3ebbd9ca3085398f5ec31" },
    "gridcraft": { "url": "https://github.com/storytold/gridcraft.git", "commit": "61ebc0e3df5ead3289b3d5a986956b19e18b45d9" },
    "deckcraft": { "url": "https://github.com/storytold/deckcraft.git", "commit": "b547f10884dbae3cf147b8efd67e88d2a5c7546a" }
  }
}
```

Run: `for r in wordcraft gridcraft deckcraft; do git -C /c/Users/OpenMinds/projects/artcraft-eval/$r rev-parse HEAD; done`
Expected: 위 세 커밋과 같다(다르면 lock을 실제 HEAD로 고치지 말고 멈춰서 보고 — 설계 근거가 된 시험이 이 커밋이다).

- [ ] **Step 3: 공통 C ABI `tools/artcraft-shell/abi.rs`**

spike `abi.rs`를 옮기면서 (1) 핸들 0이 들어와도 넘침 없이 None이 되게 `wrapping_sub`, (2) 쪽 크기 함수 `sp_page_w/h`를 더한다.

```rust
// ArtCraft 브라우저 엔진 공통 C ABI — 엔진 껍데기(word·deck)가 include!로 넣는다.
// JS: p = sp_alloc(n); 바이트 복사; 호출; sp_out_ptr/len, sp_out_w/h 읽기; sp_free(p, n).
use std::cell::RefCell;

thread_local! {
    static OUT: RefCell<(Vec<u8>, u32, u32)> = RefCell::new((Vec::new(), 0, 0));
    static MSG: RefCell<String> = RefCell::new(String::new());
    static DOCS: RefCell<Vec<Option<Doc>>> = RefCell::new(Vec::new());
}

#[unsafe(no_mangle)]
pub extern "C" fn sp_alloc(n: usize) -> *mut u8 {
    let mut v = Vec::<u8>::with_capacity(n.max(1));
    let p = v.as_mut_ptr();
    std::mem::forget(v);
    p
}

#[unsafe(no_mangle)]
pub unsafe extern "C" fn sp_free(p: *mut u8, n: usize) {
    drop(unsafe { Vec::from_raw_parts(p, 0, n.max(1)) });
}

unsafe fn slice<'a>(p: *const u8, n: usize) -> &'a [u8] {
    unsafe { std::slice::from_raw_parts(p, n) }
}

fn set_out(v: Vec<u8>, w: u32, h: u32) {
    OUT.with(|o| *o.borrow_mut() = (v, w, h));
}
fn set_msg(s: String) {
    MSG.with(|m| *m.borrow_mut() = s);
}

/// 핸들 h(1부터)의 문서로 f를 부른다. 없는 핸들이면 None.
fn with_doc<R>(h: u32, f: impl FnOnce(&Doc) -> R) -> Option<R> {
    DOCS.with(|ds| ds.borrow().get((h as usize).wrapping_sub(1)).and_then(|d| d.as_ref()).map(f))
}

#[unsafe(no_mangle)]
pub extern "C" fn sp_out_ptr() -> *const u8 { OUT.with(|o| o.borrow().0.as_ptr()) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_out_len() -> usize { OUT.with(|o| o.borrow().0.len()) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_out_w() -> u32 { OUT.with(|o| o.borrow().1) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_out_h() -> u32 { OUT.with(|o| o.borrow().2) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_out_clear() { set_out(Vec::new(), 0, 0) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_msg_ptr() -> *const u8 { MSG.with(|m| m.borrow().as_ptr()) }
#[unsafe(no_mangle)]
pub extern "C" fn sp_msg_len() -> usize { MSG.with(|m| m.borrow().len()) }

/// 이름의 확장자로 형식을 고른다. 핸들(1 이상), 실패면 0(sp_msg에 이유).
#[unsafe(no_mangle)]
pub unsafe extern "C" fn sp_open(name: *const u8, name_len: usize, p: *const u8, n: usize) -> u32 {
    let name = String::from_utf8_lossy(unsafe { slice(name, name_len) }).into_owned();
    match open_doc(&name, unsafe { slice(p, n) }) {
        Ok(d) => DOCS.with(|ds| {
            let mut ds = ds.borrow_mut();
            ds.push(Some(d));
            ds.len() as u32
        }),
        Err(e) => {
            set_msg(e);
            0
        }
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn sp_count(h: u32) -> u32 {
    with_doc(h, |d| count(d) as u32).unwrap_or(0)
}

/// i쪽(0부터)의 폭(pt). 없으면 0.
#[unsafe(no_mangle)]
pub extern "C" fn sp_page_w(h: u32, i: u32) -> f32 {
    with_doc(h, |d| page_size(d, i as usize).0).unwrap_or(0.0)
}

/// i쪽(0부터)의 높이(pt). 없으면 0.
#[unsafe(no_mangle)]
pub extern "C" fn sp_page_h(h: u32, i: u32) -> f32 {
    with_doc(h, |d| page_size(d, i as usize).1).unwrap_or(0.0)
}

/// i쪽을 scale_milli/1000 px/pt로 OUT에(RGBA8 직선 알파). 0 성공, 1 실패.
#[unsafe(no_mangle)]
pub extern "C" fn sp_render(h: u32, i: u32, scale_milli: u32) -> u32 {
    match with_doc(h, |d| render(d, i as usize, scale_milli as f32 / 1000.0)).unwrap_or_else(|| Err("bad handle".into())) {
        Ok((v, w, hh)) => {
            set_out(v, w, hh);
            0
        }
        Err(e) => {
            set_msg(e);
            1
        }
    }
}

/// i쪽의 글자 조각 JSON을 OUT에: {"w":pt,"h":pt,"runs":[{"t","x0","x1","y"(기준선),"s"}]}, pt, y 아래로. 0 성공, 1 실패.
#[unsafe(no_mangle)]
pub extern "C" fn sp_text(h: u32, i: u32) -> u32 {
    match with_doc(h, |d| text(d, i as usize)).unwrap_or_else(|| Err("bad handle".into())) {
        Ok(s) => {
            set_out(s.into_bytes(), 0, 0);
            0
        }
        Err(e) => {
            set_msg(e);
            1
        }
    }
}

#[unsafe(no_mangle)]
pub extern "C" fn sp_close(h: u32) {
    DOCS.with(|ds| {
        if let Some(d) = ds.borrow_mut().get_mut((h as usize).wrapping_sub(1)) {
            *d = None;
        }
    });
}

/// 글꼴(TTF/OTF) 넣기. alias(빈 값 가능)는 엔진이 허락하면 그 이름으로도 등록(메모리 주의 — 화면은 빈 값만 쓴다). 넣은 face 수.
#[unsafe(no_mangle)]
pub unsafe extern "C" fn sp_add_font(p: *const u8, n: usize, alias: *const u8, alias_len: usize) -> u32 {
    let alias = String::from_utf8_lossy(unsafe { slice(alias, alias_len) }).into_owned();
    add_font(unsafe { slice(p, n) }.to_vec(), &alias) as u32
}

fn json_str(s: &str) -> String {
    let mut o = String::with_capacity(s.len() + 2);
    o.push('"');
    for c in s.chars() {
        match c {
            '"' => o.push_str("\\\""),
            '\\' => o.push_str("\\\\"),
            c if (c as u32) < 0x20 => o.push_str(&format!("\\u{:04x}", c as u32)),
            c => o.push(c),
        }
    }
    o.push('"');
    o
}

fn run_json(t: &str, x0: f64, x1: f64, y: f64, s: f64) -> String {
    format!("{{\"t\":{},\"x0\":{:.1},\"x1\":{:.1},\"y\":{:.1},\"s\":{:.1}}}", json_str(t), x0, x1, y, s)
}
```

- [ ] **Step 4: 워드 껍데기 `tools/artcraft-shell/word/src/lib.rs`**

```rust
include!("../../abi.rs");
use wordcraft_doc::Document;
use wordcraft_fonts::FontDb;

struct Doc {
    doc: Document,
    layout: wordcraft_layout::DocLayout,
}

fn open_doc(name: &str, bytes: &[u8]) -> Result<Doc, String> {
    let doc = wordcraft_engine::io::open_bytes(name, bytes)?;
    let doc = wordcraft_engine::cmd::references::output_fields(&doc).unwrap_or(doc);
    let lo = wordcraft_layout::LayoutOptions { output_page_refs: true, ..Default::default() };
    let layout = wordcraft_layout::layout(&doc, &mut wordcraft_layout::LayoutCache::new(), &lo);
    Ok(Doc { doc, layout })
}

fn count(d: &Doc) -> usize {
    d.layout.pages.len()
}

fn page_size(d: &Doc, i: usize) -> (f32, f32) {
    d.layout.pages.get(i).map_or((0.0, 0.0), |p| (p.w, p.h))
}

fn render(d: &Doc, i: usize, scale: f32) -> Result<(Vec<u8>, u32, u32), String> {
    let p = d.layout.pages.get(i).ok_or_else(|| format!("no page {}", i + 1))?;
    let img = wordcraft_render::render_page(&d.doc, p, scale, &Default::default());
    Ok((img.to_straight(), img.width, img.height))
}

fn add_font(bytes: Vec<u8>, alias: &str) -> usize {
    if alias.is_empty() { FontDb::global().add_font(bytes) } else { FontDb::global().add_document_fonts(vec![(bytes, alias.to_string())]) }
}

fn text(d: &Doc, i: usize) -> Result<String, String> {
    use wordcraft_layout::display::{Draw, page_display};
    let p = d.layout.pages.get(i).ok_or_else(|| format!("no page {}", i + 1))?;
    // shortcut: glyphs inside Turned/Rotated (vertical cell text, rotated shapes) are skipped; add their transform if search must cover them.
    fn walk(ds: &[Draw], out: &mut Vec<String>) {
        for dr in ds {
            match dr {
                Draw::Glyphs { face, size, glyphs, text, .. } if !glyphs.is_empty() && !text.trim().is_empty() => {
                    let (g0, gl) = (glyphs[0], glyphs[glyphs.len() - 1]);
                    let adv = face.advance(gl.0) * *size as f64 / face.units_per_em().max(1.0);
                    out.push(run_json(text, g0.1 as f64, gl.1 as f64 + adv, g0.2 as f64, *size as f64));
                }
                Draw::Figure { draws, .. } => walk(draws, out),
                Draw::Clip { items, .. } => walk(items, out),
                _ => {}
            }
        }
    }
    let mut runs = Vec::new();
    walk(&page_display(&d.doc, p, &Default::default()), &mut runs);
    Ok(format!("{{\"w\":{},\"h\":{},\"runs\":[{}]}}", p.w, p.h, runs.join(",")))
}
```

- [ ] **Step 5: PPT 껍데기 `tools/artcraft-shell/deck/src/lib.rs`**

```rust
include!("../../abi.rs");
use deckcraft_fonts::FontDb;
use deckcraft_model::Presentation;

struct Doc {
    doc: Presentation,
}

fn open_doc(name: &str, bytes: &[u8]) -> Result<Doc, String> {
    deckcraft_engine::cmd::file::open_presentation(name, bytes).map(|doc| Doc { doc }).map_err(|e| e.to_string())
}

fn count(d: &Doc) -> usize {
    d.doc.slides.len()
}

fn page_size(d: &Doc, i: usize) -> (f32, f32) {
    if i < d.doc.slides.len() { (d.doc.slide_size.width as f32, d.doc.slide_size.height as f32) } else { (0.0, 0.0) }
}

fn render(d: &Doc, i: usize, scale: f32) -> Result<(Vec<u8>, u32, u32), String> {
    if i >= d.doc.slides.len() {
        return Err(format!("no slide {i}"));
    }
    let scale = deckcraft_engine::cmd::file::image_scale(&d.doc, scale as f64);
    let opts = deckcraft_render::RenderOpts { scale, ..Default::default() };
    let img = deckcraft_render::render_slide(&d.doc, i, &opts);
    Ok((img.to_straight(), img.width, img.height))
}

// shortcut: deckcraft's FontDb has no alias API, so `alias` is ignored (Hangul relies on fallback).
fn add_font(bytes: Vec<u8>, _alias: &str) -> usize {
    FontDb::global().add_font(bytes)
}

fn text(d: &Doc, i: usize) -> Result<String, String> {
    if i >= d.doc.slides.len() {
        return Err(format!("no slide {i}"));
    }
    let mut runs = Vec::new();
    for pt in deckcraft_render::place_slide(&d.doc, i).texts {
        let mut chars: std::collections::HashMap<usize, Vec<char>> = Default::default();
        for run in &pt.layout.runs {
            if run.glyphs.is_empty() || run.size <= 0.0 {
                continue;
            }
            let cs = chars.entry(run.para).or_insert_with(|| {
                pt.body.paragraphs.get(run.para).map(|p| p.runs.iter().flat_map(|r| r.content().chars()).collect()).unwrap_or_default()
            });
            let (a, b) = (run.chars.0.min(cs.len()), run.chars.1.min(cs.len()));
            let t: String = cs.get(a..b.max(a)).unwrap_or_default().iter().collect();
            if t.trim().is_empty() {
                continue;
            }
            let (g0, gl) = (run.glyphs[0], run.glyphs[run.glyphs.len() - 1]);
            let end = gl.1 + run.face.advance(gl.0) * run.size / run.face.upem.max(1.0);
            // shortcut: rotated text gets its bounding box's left/right only.
            let p0 = pt.transform * kurbo::Point::new(g0.1, g0.2);
            let p1 = pt.transform * kurbo::Point::new(end, gl.2);
            runs.push(run_json(&t, p0.x.min(p1.x), p0.x.max(p1.x), p0.y, run.size));
        }
    }
    Ok(format!("{{\"w\":{},\"h\":{},\"runs\":[{}]}}", d.doc.slide_size.width, d.doc.slide_size.height, runs.join(",")))
}
```

- [ ] **Step 6: Cargo.toml 틀 두 개**

`tools/artcraft-shell/word/Cargo.toml.in` (빌드 스크립트가 `@ARTCRAFT_DIR@`·`@SHELL@`을 슬래시 경로로 바꿔 `.build/word/Cargo.toml`로 쓴다):

```toml
[package]
name = "artcraft-word"
version = "0.0.0"
edition = "2024"
publish = false

[lib]
name = "artcraft_word"
crate-type = ["cdylib"]
path = "@SHELL@/word/src/lib.rs"

[workspace]

[dependencies]
wordcraft-engine = { path = "@ARTCRAFT_DIR@/wordcraft/crates/engine" }
wordcraft-doc = { path = "@ARTCRAFT_DIR@/wordcraft/crates/doc" }
wordcraft-fonts = { path = "@ARTCRAFT_DIR@/wordcraft/crates/fonts" }
wordcraft-layout = { path = "@ARTCRAFT_DIR@/wordcraft/crates/layout" }
wordcraft-render = { path = "@ARTCRAFT_DIR@/wordcraft/crates/render" }

[profile.release]
opt-level = "z"
lto = true
codegen-units = 1
panic = "abort"
strip = true
debug = 0

[patch.crates-io]
krilla = { git = "https://github.com/storytold/craft-libs", rev = "93dac78c93a7f1fd66e2022edeb089233b1dc45a" }
ecolor = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
emath = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
epaint = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
epaint_default_fonts = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui-winit = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui-wgpu = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_glow = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_extras = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_kittest = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_system_font_provider = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
eframe = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_demo_lib = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_inspection = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
winit = { git = "https://github.com/storytold/winit.git", rev = "a517b3846ee8457d84ee4975110bfda6797c44dc" }
```

`tools/artcraft-shell/deck/Cargo.toml.in`:

```toml
[package]
name = "artcraft-deck"
version = "0.0.0"
edition = "2024"
publish = false

[lib]
name = "artcraft_deck"
crate-type = ["cdylib"]
path = "@SHELL@/deck/src/lib.rs"

[workspace]

[dependencies]
deckcraft-engine = { path = "@ARTCRAFT_DIR@/deckcraft/crates/engine" }
deckcraft-fonts = { path = "@ARTCRAFT_DIR@/deckcraft/crates/fonts" }
deckcraft-model = { path = "@ARTCRAFT_DIR@/deckcraft/crates/model" }
deckcraft-render = { path = "@ARTCRAFT_DIR@/deckcraft/crates/render" }
kurbo = "0.13"

[profile.release]
opt-level = "z"
lto = true
codegen-units = 1
panic = "abort"
strip = true
debug = 0

[patch.crates-io]
krilla = { git = "https://github.com/storytold/craft-libs", rev = "63196e66f3033a6623eb629af95a104d26986725", version = "=0.8.2" }
ecolor = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
emath = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
epaint = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
epaint_default_fonts = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui-winit = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui-wgpu = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_glow = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_extras = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_kittest = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_system_font_provider = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
eframe = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_demo_lib = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
egui_inspection = { git = "https://github.com/storytold/egui.git", rev = "6d518fc799681571963842f95b55a6b68431b15a" }
winit = { git = "https://github.com/storytold/winit.git", rev = "a517b3846ee8457d84ee4975110bfda6797c44dc" }
```

- [ ] **Step 7: 잠금 파일 옮기기(패키지 이름만 바꿈)**

```bash
cd /c/Users/OpenMinds/projects/webviewer-artcraft
SP=/c/Users/OpenMinds/projects/artcraft-eval/spike-wasm
cp "$SP/word/Cargo.lock" tools/artcraft-shell/word/Cargo.lock
cp "$SP/deck/Cargo.lock" tools/artcraft-shell/deck/Cargo.lock
sed -i 's/^name = "spike-word"$/name = "artcraft-word"/' tools/artcraft-shell/word/Cargo.lock
sed -i 's/^name = "spike-deck"$/name = "artcraft-deck"/' tools/artcraft-shell/deck/Cargo.lock
grep -c '^name = "artcraft-word"$' tools/artcraft-shell/word/Cargo.lock
grep -c '^name = "artcraft-deck"$' tools/artcraft-shell/deck/Cargo.lock
grep -c 'spike' tools/artcraft-shell/word/Cargo.lock tools/artcraft-shell/deck/Cargo.lock
```
Expected: `1`, `1`, 그리고 두 파일 모두 `:0`. (path 패키지는 잠금에 경로·체크섬이 없어서 ARTCRAFT_DIR 위치가 바뀌어도 `--locked`가 통한다.)

- [ ] **Step 8: 빌드 도우미 실패하는 시험 쓰기**

`tests/artcraftBuild.test.ts`:

```ts
import { brotliDecompressSync } from 'node:zlib';
import { describe, expect, test } from 'vitest';
import { MiB, SIZE_LIMITS, brotli, checkLock, hashedName, sizeProblems } from '../scripts/artcraftBuildLib.js';

describe('build:artcraft 도우미', () => {
  test('해시 이름: 같은 내용은 같은 이름, 내용이 바뀌면 이름도 바뀐다(10자리 16진수)', () => {
    const a = hashedName('artcraft_word', '.wasm', Buffer.from('하나'));
    expect(a).toMatch(/^artcraft_word-[0-9a-f]{10}\.wasm$/);
    expect(hashedName('artcraft_word', '.wasm', Buffer.from('하나'))).toBe(a);
    expect(hashedName('artcraft_word', '.wasm', Buffer.from('둘'))).not.toBe(a);
  });

  test('크기 상한: word 5 MB, deck 2.5 MB, 글꼴 합 0.6 MB(brotli)', () => {
    expect(SIZE_LIMITS).toEqual({ word: 5 * MiB, deck: 2.5 * MiB, fonts: 0.6 * MiB });
    expect(sizeProblems({ word: 5 * MiB, deck: 2.5 * MiB, fonts: [0.3 * MiB, 0.3 * MiB] })).toEqual([]);
    const over = sizeProblems({ word: 5 * MiB + 1, deck: 2.5 * MiB + 1, fonts: [0.3 * MiB, 0.3 * MiB + 1] });
    expect(over).toHaveLength(3);
    expect(over[0]).toContain('word');
    expect(over[1]).toContain('deck');
    expect(over[2]).toContain('글꼴');
  });

  test('잠금 검사: HEAD가 다르거나 고친 파일이 있으면 이유, 같고 깨끗하면 null', () => {
    const want = 'f8ec3cc0fa884369c4f3ebbd9ca3085398f5ec31';
    expect(checkLock('wordcraft', want, want, '')).toBeNull();
    expect(checkLock('wordcraft', want, 'b'.repeat(40), '')).toContain('ARTCRAFT.lock');
    expect(checkLock('wordcraft', want, want, ' M crates/doc/src/lib.rs\n')).toContain('고친 파일');
  });

  test('brotli: 풀면 원래 바이트', () => {
    const raw = Buffer.from('아트크래프트 '.repeat(500));
    const br = brotli(raw);
    expect(br.length).toBeLessThan(raw.length);
    expect(Buffer.compare(brotliDecompressSync(br), raw)).toBe(0);
  });
});
```

- [ ] **Step 9: 시험이 실패하는지 보기**

Run: `npx vitest run tests/artcraftBuild.test.ts`
Expected: FAIL — `Failed to load url ../scripts/artcraftBuildLib.js`(파일 없음).

- [ ] **Step 10: `scripts/artcraftBuildLib.ts` 쓰기**

```ts
/**
 * build:artcraft의 순수 도우미 — 시험(tests/artcraftBuild.test.ts)이 직접 부른다.
 * 크기 상한은 설계 §5(word ≤ 5 MB, deck ≤ 2.5 MB, 글꼴 2개 합 ≤ 0.6 MB, 모두 brotli). MB = 1024×1024.
 */
import { createHash } from 'node:crypto';
import { brotliCompressSync, constants } from 'node:zlib';

export const MiB = 1024 * 1024;
export const SIZE_LIMITS = { word: 5 * MiB, deck: 2.5 * MiB, fonts: 0.6 * MiB } as const;

const mb = (n: number) => (n / MiB).toFixed(2);

/** `<base>-<내용 sha256 앞 10자리><ext>` — 이름이 바뀌어야 immutable 캐시가 새 파일을 받는다 */
export function hashedName(base: string, ext: string, bytes: Uint8Array): string {
  return `${base}-${createHash('sha256').update(bytes).digest('hex').slice(0, 10)}${ext}`;
}

/** 상한을 넘은 것마다 한 줄 — 빈 배열이면 통과 */
export function sizeProblems(br: { readonly word: number; readonly deck: number; readonly fonts: readonly number[] }): string[] {
  const out: string[] = [];
  if (br.word > SIZE_LIMITS.word) out.push(`word wasm brotli ${mb(br.word)} MB > ${mb(SIZE_LIMITS.word)} MB`);
  if (br.deck > SIZE_LIMITS.deck) out.push(`deck wasm brotli ${mb(br.deck)} MB > ${mb(SIZE_LIMITS.deck)} MB`);
  const fonts = br.fonts.reduce((a, b) => a + b, 0);
  if (fonts > SIZE_LIMITS.fonts) out.push(`글꼴 brotli 합 ${mb(fonts)} MB > ${mb(SIZE_LIMITS.fonts)} MB`);
  return out;
}

/** ArtCraft 저장소가 고정 커밋 그대로인지 — 문제면 이유, 아니면 null */
export function checkLock(repo: string, wantCommit: string, head: string, porcelain: string): string | null {
  if (head.trim() !== wantCommit) return `${repo}: HEAD ${head.trim().slice(0, 12)} ≠ ARTCRAFT.lock ${wantCommit.slice(0, 12)}`;
  if (porcelain.trim() !== '') return `${repo}: 고친 파일이 있습니다(git status) — 저장소는 고치지 않는다`;
  return null;
}

export function brotli(bytes: Uint8Array): Buffer {
  return brotliCompressSync(bytes, { params: { [constants.BROTLI_PARAM_QUALITY]: 11, [constants.BROTLI_PARAM_SIZE_HINT]: bytes.length } });
}
```

- [ ] **Step 11: 시험이 통과하는지 보기**

Run: `npx vitest run tests/artcraftBuild.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 12: 빌드 스크립트 `scripts/build-artcraft.ts`**

```ts
/**
 * npm run build:artcraft — 아트크래프트 브라우저 엔진(설계 §5).
 *  1. ARTCRAFT_DIR(기본 C:/Users/OpenMinds/projects/artcraft-eval)의 wordcraft·deckcraft가 ARTCRAFT.lock 커밋 그대로인지 본다.
 *  2. tools/artcraft-shell/{word,deck}을 wasm32-unknown-unknown(GNU 도구 묶음, wasm-bindgen 없이)으로 빌드한다.
 *  3. wasm-bindgen CLI(WASM_BINDGEN 또는 tools/wasm-bindgen/bin/wasm-bindgen.exe)가 있으면 크기를 줄인다(없으면 건너뜀).
 *  4. 크기 상한을 넘으면 아무것도 쓰지 않고 멈춘다. 통과하면 public/artcraft/를 비우고 해시 이름 파일 + .br + manifest.json을 쓴다.
 */
import { execFileSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';
import { brotli, checkLock, hashedName, sizeProblems } from './artcraftBuildLib.js';

const ROOT = process.cwd();
const SHELL = path.join(ROOT, 'tools', 'artcraft-shell');
const BUILD = path.join(SHELL, '.build');
const OUT = path.join(ROOT, 'public', 'artcraft');
const ARTCRAFT_DIR = path.resolve(process.env.ARTCRAFT_DIR ?? 'C:/Users/OpenMinds/projects/artcraft-eval');
const WASM_BINDGEN = process.env.WASM_BINDGEN ?? path.join(ROOT, 'tools', 'wasm-bindgen', 'bin', 'wasm-bindgen.exe');
/** wasm 최대 메모리(설계 §6 — 조작된 문서가 탭 메모리를 다 쓰지 못하게). 고객 문서 실측 35 MB, 별칭 8개 115 MB */
const MAX_MEMORY = 512 * 1024 * 1024;
const FONT_DIR = path.join(ROOT, 'public', 'fonts', 's-core-dream');
const CRATES = [
  { id: 'word', repo: 'wordcraft', lib: 'artcraft_word' },
  { id: 'deck', repo: 'deckcraft', lib: 'artcraft_deck' },
] as const;

const fwd = (p: string) => p.replace(/\\/g, '/');
function fail(message: string): never {
  console.error(`build:artcraft 실패 — ${message}`);
  process.exit(1);
}
const git = (dir: string, ...args: string[]) => execFileSync('git', ['-C', dir, ...args], { encoding: 'utf8' });

const lock = JSON.parse(fs.readFileSync(path.join(SHELL, 'ARTCRAFT.lock'), 'utf8')) as { repos: Record<string, { url: string; commit: string }> };
for (const { repo } of CRATES) {
  const want = lock.repos[repo]?.commit;
  if (!want) fail(`ARTCRAFT.lock에 ${repo} 커밋이 없습니다`);
  const dir = path.join(ARTCRAFT_DIR, repo);
  if (!fs.existsSync(dir)) fail(`${dir} 없음 — ${lock.repos[repo].url}을(를) ${want}로 받아 두세요`);
  const problem = checkLock(repo, want, git(dir, 'rev-parse', 'HEAD'), git(dir, 'status', '--porcelain', '--untracked-files=no'));
  if (problem) fail(problem);
}

const wasmBytes: Record<string, Buffer> = {};
for (const c of CRATES) {
  const dir = path.join(BUILD, c.id);
  fs.mkdirSync(dir, { recursive: true });
  const toml = fs
    .readFileSync(path.join(SHELL, c.id, 'Cargo.toml.in'), 'utf8')
    .replace(/@ARTCRAFT_DIR@/g, fwd(ARTCRAFT_DIR))
    .replace(/@SHELL@/g, fwd(SHELL));
  fs.writeFileSync(path.join(dir, 'Cargo.toml'), toml);
  fs.copyFileSync(path.join(SHELL, c.id, 'Cargo.lock'), path.join(dir, 'Cargo.lock'));
  console.log(`[${c.id}] cargo build (몇 분 걸린다)…`);
  // wasm32 대상은 GNU 도구 묶음에만 설치돼 있다(MSVC로 하면 "can't find crate for core" — spike §4)
  execFileSync('cargo', ['+stable-x86_64-pc-windows-gnu', 'build', '--release', '--locked', '--target', 'wasm32-unknown-unknown'], {
    cwd: dir,
    stdio: 'inherit',
    env: {
      ...process.env,
      CARGO_TARGET_DIR: path.join(BUILD, 'target'),
      CARGO_TARGET_WASM32_UNKNOWN_UNKNOWN_RUSTFLAGS: `-C link-arg=--max-memory=${MAX_MEMORY}`,
    },
  });
  let wasm = path.join(BUILD, 'target', 'wasm32-unknown-unknown', 'release', `${c.lib}.wasm`);
  if (fs.existsSync(WASM_BINDGEN)) {
    const out = path.join(BUILD, 'wbg');
    execFileSync(WASM_BINDGEN, ['--target', 'web', '--no-typescript', '--out-dir', out, wasm], { stdio: 'inherit' });
    wasm = path.join(out, `${c.lib}_bg.wasm`);
  } else {
    console.log(`[${c.id}] wasm-bindgen 없음(${WASM_BINDGEN}) — 크기 줄이기 건너뜀(약 10% 큼)`);
  }
  wasmBytes[c.id] = fs.readFileSync(wasm);
}

const fontBytes = ['SCDream4', 'SCDream6'].map((name) => ({ name, bytes: fs.readFileSync(path.join(FONT_DIR, `${name}.otf`)) }));
const br = {
  word: brotli(wasmBytes.word),
  deck: brotli(wasmBytes.deck),
  fonts: fontBytes.map((f) => brotli(f.bytes)),
};
const problems = sizeProblems({ word: br.word.length, deck: br.deck.length, fonts: br.fonts.map((b) => b.length) });
if (problems.length > 0) fail(`크기 상한 초과:\n  ${problems.join('\n  ')}`);

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(path.join(OUT, 'fonts'), { recursive: true });
function emit(sub: string, base: string, ext: string, raw: Buffer, compressed: Buffer): string {
  const name = hashedName(base, ext, raw);
  const file = path.join(OUT, sub, name);
  fs.writeFileSync(file, raw);
  fs.writeFileSync(`${file}.br`, compressed);
  console.log(`  ${sub ? `${sub}/` : ''}${name}  원본 ${(raw.length / 1048576).toFixed(2)} MB · br ${(compressed.length / 1048576).toFixed(2)} MB`);
  return `/artcraft/${sub ? `${sub}/` : ''}${name}`;
}
const manifest = {
  word: emit('', 'artcraft_word', '.wasm', wasmBytes.word, br.word),
  deck: emit('', 'artcraft_deck', '.wasm', wasmBytes.deck, br.deck),
  fonts: fontBytes.map((f, i) => emit('fonts', f.name, '.otf', f.bytes, br.fonts[i])),
  commits: Object.fromEntries(CRATES.map((c) => [c.repo, lock.repos[c.repo].commit])),
};
fs.writeFileSync(path.join(OUT, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`);
console.log('build:artcraft 완료 → public/artcraft/manifest.json');
```

`package.json` `scripts`에 `"build:editor"` 줄 다음으로 더한다:

```json
    "build:artcraft": "tsx scripts/build-artcraft.ts",
```

- [ ] **Step 13: 빌드를 돌려 산출물·크기 보기**

wasm-bindgen으로 줄이려면 먼저(선택) spike에만 설치된 CLI를 무시되는 `tools/`로 복사한다(같은 0.2.129, 잠금 파일의 `wasm-bindgen` 판과 같아야 한다):

```bash
cd /c/Users/OpenMinds/projects/webviewer-artcraft
mkdir -p tools/wasm-bindgen/bin && cp /c/Users/OpenMinds/projects/artcraft-eval/spike-wasm/tools/bin/wasm-bindgen.exe tools/wasm-bindgen/bin/
npm run build:artcraft
ls -la public/artcraft public/artcraft/fonts && cat public/artcraft/manifest.json
```
Expected: 끝에 `build:artcraft 완료`. word br 약 4.2 MB(≤ 5), deck 약 1.8 MB(≤ 2.5), 글꼴 둘 약 0.22 MB씩. `public/artcraft/`에 `artcraft_word-*.wasm(.br)`·`artcraft_deck-*.wasm(.br)`·`fonts/SCDream{4,6}-*.otf(.br)`·`manifest.json`. `git status --short public/artcraft`는 빈 출력(무시됨).

- [ ] **Step 14: manifest·정적 서빙 실패하는 시험 쓰기**

`tests/artcraftAssets.test.ts`:

```ts
import { mkdtempSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { brotliCompressSync } from 'node:zlib';
import Fastify, { type FastifyInstance } from 'fastify';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import { parseManifest, readArtcraftManifest } from '../src/artcraftManifest.js';
import { registerArtcraftStatic } from '../src/artcraftStatic.js';

const GOOD = {
  word: '/artcraft/artcraft_word-0123456789.wasm',
  deck: '/artcraft/artcraft_deck-abcdef0123.wasm',
  fonts: ['/artcraft/fonts/SCDream4-0123456789.otf', '/artcraft/fonts/SCDream6-0123456789.otf'],
  commits: { wordcraft: 'f8ec3cc0', deckcraft: 'b547f108' },
};
const tmp = (tag: string) => mkdtempSync(path.join(os.tmpdir(), `artcraft-${tag}-`));

describe('manifest.json 읽기', () => {
  test('맞는 manifest는 word·deck·fonts만 돌려준다', () => {
    expect(parseManifest(GOOD)).toEqual({ word: GOOD.word, deck: GOOD.deck, fonts: GOOD.fonts });
  });

  test('모양이 틀리거나 /artcraft/ 밖 주소·..이 있으면 null', () => {
    expect(parseManifest(null)).toBeNull();
    expect(parseManifest({ ...GOOD, word: 42 })).toBeNull();
    expect(parseManifest({ ...GOOD, fonts: [] })).toBeNull();
    expect(parseManifest({ ...GOOD, deck: 'https://evil.example/x.wasm' })).toBeNull();
    expect(parseManifest({ ...GOOD, fonts: ['/artcraft/../secret.otf'] })).toBeNull();
  });

  test('파일이 없으면 null(아트크래프트 옵션이 꺼진 채 서버는 뜬다), 있으면 읽는다, 깨진 JSON은 null', () => {
    expect(readArtcraftManifest(tmp('none'))).toBeNull();
    const ok = tmp('ok');
    writeFileSync(path.join(ok, 'manifest.json'), JSON.stringify(GOOD));
    expect(readArtcraftManifest(ok)?.word).toBe(GOOD.word);
    const broken = tmp('broken');
    writeFileSync(path.join(broken, 'manifest.json'), '{ 깨짐');
    expect(readArtcraftManifest(broken)).toBeNull();
  });
});

describe('/artcraft/* 정적 서빙', () => {
  let app: FastifyInstance;
  const dir = tmp('static');
  const raw = Buffer.from('\0asm 아트크래프트 시험 '.repeat(200));
  const br = brotliCompressSync(raw);
  beforeAll(async () => {
    writeFileSync(path.join(dir, 'artcraft_word-0123456789.wasm'), raw);
    writeFileSync(path.join(dir, 'artcraft_word-0123456789.wasm.br'), br);
    writeFileSync(path.join(dir, 'manifest.json'), JSON.stringify(GOOD));
    app = Fastify();
    await registerArtcraftStatic(app, dir);
    await app.ready();
  });
  afterAll(() => app.close());

  test('br을 받는 브라우저에는 .br(content-encoding: br, application/wasm, immutable)', async () => {
    const res = await app.inject({ method: 'GET', url: '/artcraft/artcraft_word-0123456789.wasm', headers: { 'accept-encoding': 'gzip, deflate, br' } });
    expect(res.statusCode).toBe(200);
    expect(res.headers['content-encoding']).toBe('br');
    expect(res.headers['content-type']).toBe('application/wasm');
    expect(res.headers['cache-control']).toBe('public, max-age=31536000, immutable');
    expect(String(res.headers.vary)).toContain('accept-encoding');
    expect(Buffer.compare(res.rawPayload, br)).toBe(0);
  });

  test('br을 안 받으면 원본 그대로', async () => {
    const res = await app.inject({ method: 'GET', url: '/artcraft/artcraft_word-0123456789.wasm', headers: { 'accept-encoding': 'identity' } });
    expect(res.statusCode).toBe(200);
    expect(res.headers['content-encoding']).toBeUndefined();
    expect(Buffer.compare(res.rawPayload, raw)).toBe(0);
  });

  test('manifest.json은 내주지 않는다(서버만 읽는다)', async () => {
    expect((await app.inject({ method: 'GET', url: '/artcraft/manifest.json' })).statusCode).toBe(404);
  });
});
```

- [ ] **Step 15: 시험이 실패하는지 보기**

Run: `npx vitest run tests/artcraftAssets.test.ts`
Expected: FAIL — `../src/artcraftManifest.js`를 찾지 못함.

- [ ] **Step 16: `src/artcraftManifest.ts` 쓰기**

```ts
/**
 * 아트크래프트 브라우저 엔진 자산 목록(public/artcraft/manifest.json, npm run build:artcraft가 만든다) — /api/view/info가
 * 해시 붙은 wasm·글꼴 주소를 여기서 읽는다. 파일이 없거나 틀리면 null → 아트크래프트 옵션을 끈 채 서버는 그대로 돈다(설계 §5).
 * 빌드를 서버가 뜬 뒤에 돌려도 되게 파일 수정 시각이 바뀌면 다시 읽는다(assetVersion.ts와 같은 방식).
 */
import { readFileSync, statSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

export interface ArtcraftManifest {
  readonly word: string;
  readonly deck: string;
  readonly fonts: readonly string[];
}

const DEFAULT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'public', 'artcraft');
/** 같은 출처 정적 주소만 — manifest가 바뀌어도 Worker가 밖으로 요청하지 않게(설계 §6) */
const ASSET_URL = /^\/artcraft\/[A-Za-z0-9_./-]+$/;

/** 자산 폴더 — 시험은 DOCVIEW_ARTCRAFT_DIR로 바꾼다(부를 때마다 읽는다) */
export function artcraftDir(): string {
  return process.env.DOCVIEW_ARTCRAFT_DIR ?? DEFAULT_DIR;
}

export function parseManifest(raw: unknown): ArtcraftManifest | null {
  if (typeof raw !== 'object' || raw === null) return null;
  const { word, deck, fonts } = raw as Record<string, unknown>;
  const ok = (v: unknown): v is string => typeof v === 'string' && ASSET_URL.test(v) && !v.includes('..');
  if (!ok(word) || !ok(deck) || !Array.isArray(fonts) || fonts.length === 0 || !fonts.every(ok)) return null;
  return { word, deck, fonts: [...fonts] };
}

let cached: { readonly file: string; readonly mtimeMs: number; readonly manifest: ArtcraftManifest | null } | null = null;

export function readArtcraftManifest(dir: string = artcraftDir()): ArtcraftManifest | null {
  const file = path.join(dir, 'manifest.json');
  let mtimeMs: number;
  try {
    mtimeMs = statSync(file).mtimeMs;
  } catch {
    return null;
  }
  if (cached && cached.file === file && cached.mtimeMs === mtimeMs) return cached.manifest;
  let manifest: ArtcraftManifest | null;
  try {
    manifest = parseManifest(JSON.parse(readFileSync(file, 'utf8')));
  } catch {
    manifest = null;
  }
  cached = { file, mtimeMs, manifest };
  return manifest;
}
```

- [ ] **Step 17: `src/artcraftStatic.ts` 쓰기**

```ts
/**
 * /artcraft/* — 아트크래프트 wasm·글꼴(이름에 내용 해시). 브라우저가 br을 받으면 미리 압축한 .br을(content-encoding: br),
 * 아니면 원본을 내준다. 이름이 내용과 함께 바뀌므로 1년 immutable(설계 §5). manifest.json은 서버만 읽는다.
 */
import { mkdirSync } from 'node:fs';
import fastifyStatic from '@fastify/static';
import type { FastifyInstance } from 'fastify';

export async function registerArtcraftStatic(app: FastifyInstance, root: string): Promise<void> {
  // 빌드 전이라 폴더가 없어도 서버는 뜬다(빈 폴더 = 아트크래프트 꺼짐)
  mkdirSync(root, { recursive: true });
  await app.register(fastifyStatic, {
    root,
    prefix: '/artcraft/',
    decorateReply: false,
    preCompressed: true,
    index: false,
    allowedPath: (pathName) => pathName !== '/manifest.json',
    setHeaders(res) {
      res.setHeader('cache-control', 'public, max-age=31536000, immutable');
      res.setHeader('vary', 'accept-encoding');
    },
  });
}
```

- [ ] **Step 18: `src/app.ts`에 등록**

import 줄 묶음에 더한다:

```ts
import { artcraftDir } from './artcraftManifest.js';
import { registerArtcraftStatic } from './artcraftStatic.js';
```

`await app.register(fastifyStatic, {` (publicDir 등록) 바로 앞에 더한다:

```ts
  // 아트크래프트 브라우저 엔진 자산(.br·immutable) — 아래 public 전체 서빙보다 구체적인 /artcraft/* 경로가 먼저 걸린다
  await registerArtcraftStatic(app, artcraftDir());
```

- [ ] **Step 19: 시험·타입 검사**

Run: `npx vitest run tests/artcraftAssets.test.ts tests/artcraftBuild.test.ts tests/appHardening.test.ts && npm run typecheck`
Expected: PASS(세 파일 모두), typecheck 오류 0.

- [ ] **Step 20: 커밋**

```bash
git add .gitignore package.json src/app.ts src/artcraftManifest.ts src/artcraftStatic.ts scripts/artcraftBuildLib.ts scripts/build-artcraft.ts tests/artcraftBuild.test.ts tests/artcraftAssets.test.ts tools/artcraft-shell/ARTCRAFT.lock tools/artcraft-shell/abi.rs tools/artcraft-shell/word/Cargo.toml.in tools/artcraft-shell/word/Cargo.lock tools/artcraft-shell/word/src/lib.rs tools/artcraft-shell/deck/Cargo.toml.in tools/artcraft-shell/deck/Cargo.lock tools/artcraft-shell/deck/src/lib.rs
git status --short   # public/artcraft·tools/wasm-bindgen·.build가 목록에 없어야 한다
git commit -m "아트크래프트 1: 껍데기 크레이트(tools/artcraft-shell)·build:artcraft·manifest·/artcraft 정적 서빙(.br·immutable)"
```

---

### Task 2: 서버 — 엔진 판정·토큰·`/api/view/info`

**Files:**
- Create: `src/artcraftEngine.ts`
- Modify: `src/types.ts:47-51` (ViewClaims), `src/security.ts` (issueViewToken·verifyViewToken), `src/routes/documents.ts:26-47,520-574` (tokenBodySchema·TokenBody·토큰 라우트), `src/routes/viewer.ts:714-805` (`/api/view/info`)
- Test: `tests/artcraftEngine.test.ts`, `tests/artcraftRoutes.test.ts`

**Interfaces:**
- Consumes: `readArtcraftManifest(): ArtcraftManifest | null` (Task 1, `src/artcraftManifest.ts`).
- Produces:
  - `src/artcraftEngine.ts`: `type Engine = 'libreoffice' | 'artcraft'`; `ENGINE_REASONS = { unsupported, notInstalled, noDownload }`; `ARTCRAFT_FORMATS: ReadonlyMap<string, { kind: 'word' | 'deck'; experimental: boolean }>`; `sourceExtOf(meta: Pick<DocumentMeta, 'ext' | 'legacy'>): string`; `engineFor(ext: string, perms: { canDownload: boolean }, requested: Engine | undefined, installed: boolean): { engine: Engine; reason?: string }`; `interface ArtcraftInfo { kind; ext; experimental; originalUrl; assets: { wasm; fonts } }`; `artcraftInfo(ext: string, token: string, manifest: ArtcraftManifest): ArtcraftInfo | undefined`
  - `ViewClaims.engine?: 'artcraft'` — **요청한** 엔진(판정은 info가 매번 다시 한다).
  - 토큰 응답 `data.engine`, `data.engineReason?`; info `data.engine`, `data.engineReason?`, `data.artcraft?`(모양은 `ArtcraftInfo`).

- [ ] **Step 1: 판정표 실패하는 시험 쓰기**

`tests/artcraftEngine.test.ts`:

```ts
import { describe, expect, test } from 'vitest';
import { ENGINE_REASONS, artcraftInfo, engineFor, sourceExtOf, type Engine } from '../src/artcraftEngine.js';

const M = { word: '/artcraft/artcraft_word-0123456789.wasm', deck: '/artcraft/artcraft_deck-0123456789.wasm', fonts: ['/artcraft/fonts/SCDream4-0123456789.otf', '/artcraft/fonts/SCDream6-0123456789.otf'] };

describe('engineFor — 형식 × 권한 × 요청 × 설치', () => {
  const R = ENGINE_REASONS;
  const table: Array<[string, boolean, Engine | undefined, boolean, Engine, string | undefined]> = [
    ['docx', true, 'artcraft', true, 'artcraft', undefined],
    ['pptx', true, 'artcraft', true, 'artcraft', undefined],
    ['doc', true, 'artcraft', true, 'artcraft', undefined],
    ['docx', false, 'artcraft', true, 'libreoffice', R.noDownload],
    ['ppt', true, 'artcraft', true, 'libreoffice', R.unsupported],
    ['xlsx', true, 'artcraft', true, 'libreoffice', R.unsupported],
    ['xls', true, 'artcraft', true, 'libreoffice', R.unsupported],
    ['hwp', true, 'artcraft', true, 'libreoffice', R.unsupported],
    ['pdf', false, 'artcraft', true, 'libreoffice', R.unsupported],
    ['docx', true, 'artcraft', false, 'libreoffice', R.notInstalled],
    ['docx', true, 'libreoffice', true, 'libreoffice', undefined],
    ['docx', true, undefined, true, 'libreoffice', undefined],
    ['xlsx', false, undefined, false, 'libreoffice', undefined],
  ];
  test.each(table)('%s · 내려받기 %s · 요청 %s · 설치 %s → %s', (ext, canDownload, requested, installed, engine, reason) => {
    expect(engineFor(ext, { canDownload }, requested, installed)).toEqual(reason ? { engine, reason } : { engine });
  });

  test('이유 문구는 설계 그대로', () => {
    expect(ENGINE_REASONS).toEqual({
      unsupported: '이 형식은 브라우저 엔진이 지원하지 않습니다',
      notInstalled: '브라우저 엔진이 설치되지 않았습니다',
      noDownload: '내려받기 권한이 없어 서버 변환으로 엽니다',
    });
  });

  test('constructor 같은 이름도 형식으로 보지 않는다', () => {
    expect(engineFor('constructor', { canDownload: true }, 'artcraft', true).engine).toBe('libreoffice');
  });
});

describe('sourceExtOf — 내려받기로 실제 내려가는 바이트의 형식', () => {
  test('보통 문서는 ext, 향기랩 옛 형식(legacy)은 올린 형식', () => {
    expect(sourceExtOf({ ext: 'docx' })).toBe('docx');
    expect(sourceExtOf({ ext: 'docx', legacy: { ext: 'doc', sha256: 'x', hasMacros: false } })).toBe('doc');
    expect(sourceExtOf({ ext: 'pptx', legacy: { ext: 'ppt', sha256: 'x', hasMacros: false } })).toBe('ppt');
  });
});

describe('artcraftInfo — info의 artcraft 블록', () => {
  test('docx → word, doc → word 실험, pptx → deck, 그 밖 undefined', () => {
    expect(artcraftInfo('docx', 'a.b+c', M)).toEqual({
      kind: 'word',
      ext: 'docx',
      experimental: false,
      originalUrl: '/api/view/download?token=a.b%2Bc',
      assets: { wasm: M.word, fonts: M.fonts },
    });
    expect(artcraftInfo('doc', 't', M)).toMatchObject({ kind: 'word', experimental: true });
    expect(artcraftInfo('pptx', 't', M)).toMatchObject({ kind: 'deck', experimental: false, assets: { wasm: M.deck } });
    expect(artcraftInfo('xlsx', 't', M)).toBeUndefined();
  });
});
```

- [ ] **Step 2: 시험이 실패하는지 보기**

Run: `npx vitest run tests/artcraftEngine.test.ts`
Expected: FAIL — `../src/artcraftEngine.js`를 찾지 못함.

- [ ] **Step 3: `src/artcraftEngine.ts` 쓰기**

```ts
/**
 * 여는 엔진 고르기(설계 §2·§3.1) — 토큰 발급과 /api/view/info가 같은 판정을 쓴다. 토큰 클레임에는 "요청한 엔진"만 넣고
 * (위조 방지 — 주소 쿼리는 보지 않는다) 판정은 부를 때마다 다시 한다(그사이 manifest가 사라져도 안전하게).
 * 원본 바이트가 브라우저로 내려가므로 내려받기 권한(canDownload)과 같은 뜻이다.
 */
import type { ArtcraftManifest } from './artcraftManifest.js';
import type { DocumentMeta } from './types.js';

export type Engine = 'libreoffice' | 'artcraft';

export const ENGINE_REASONS = {
  unsupported: '이 형식은 브라우저 엔진이 지원하지 않습니다',
  notInstalled: '브라우저 엔진이 설치되지 않았습니다',
  noDownload: '내려받기 권한이 없어 서버 변환으로 엽니다',
} as const;

/** 브라우저 엔진이 그리는 형식 — .doc는 표가 깨지는 파일이 있어 "실험"(설계 §2). .ppt·엑셀·한글은 엔진 없음 */
export const ARTCRAFT_FORMATS: ReadonlyMap<string, { readonly kind: 'word' | 'deck'; readonly experimental: boolean }> = new Map([
  ['docx', { kind: 'word', experimental: false }],
  ['pptx', { kind: 'deck', experimental: false }],
  ['doc', { kind: 'word', experimental: true }],
] as const);

export interface ArtcraftInfo {
  readonly kind: 'word' | 'deck';
  /** Worker가 엔진에 넘길 파일 이름의 확장자(엔진은 확장자로 형식을 고른다) */
  readonly ext: string;
  readonly experimental: boolean;
  /** 지금 있는 원본 내려받기 경로(토큰 포함, canDownload 검사 그대로) */
  readonly originalUrl: string;
  readonly assets: { readonly wasm: string; readonly fonts: readonly string[] };
}

/** /api/view/download가 실제로 내주는 바이트의 형식 — 향기랩 옛 형식(legacy)은 OOXML로 바꿔 두었어도 올린 .doc·.ppt를 내준다 */
export function sourceExtOf(meta: Pick<DocumentMeta, 'ext' | 'legacy'>): string {
  return meta.legacy?.ext ?? meta.ext;
}

export function engineFor(
  ext: string,
  perms: { readonly canDownload: boolean },
  requested: Engine | undefined,
  installed: boolean,
): { engine: Engine; reason?: string } {
  if (requested !== 'artcraft') return { engine: 'libreoffice' };
  if (!ARTCRAFT_FORMATS.has(ext)) return { engine: 'libreoffice', reason: ENGINE_REASONS.unsupported };
  if (!installed) return { engine: 'libreoffice', reason: ENGINE_REASONS.notInstalled };
  if (!perms.canDownload) return { engine: 'libreoffice', reason: ENGINE_REASONS.noDownload };
  return { engine: 'artcraft' };
}

export function artcraftInfo(ext: string, token: string, manifest: ArtcraftManifest): ArtcraftInfo | undefined {
  const format = ARTCRAFT_FORMATS.get(ext);
  if (!format) return undefined;
  return {
    kind: format.kind,
    ext,
    experimental: format.experimental,
    originalUrl: `/api/view/download?token=${encodeURIComponent(token)}`,
    assets: { wasm: manifest[format.kind], fonts: manifest.fonts },
  };
}
```

- [ ] **Step 4: 시험이 통과하는지 보기**

Run: `npx vitest run tests/artcraftEngine.test.ts`
Expected: PASS.

- [ ] **Step 5: 토큰·info 실패하는 시험 쓰기**

`tests/artcraftRoutes.test.ts`:

```ts
import { mkdtempSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';

type Perms = { canPrint?: boolean; canDownload?: boolean; canCopy?: boolean; canEdit?: boolean };

describe('아트크래프트 엔진 — 토큰 발급·/api/view/info', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let newDocumentId: (typeof import('../src/storage.js'))['newDocumentId'];
  let issueViewToken: (typeof import('../src/security.js'))['issueViewToken'];
  let verifyViewToken: (typeof import('../src/security.js'))['verifyViewToken'];
  const installed = mkdtempSync(path.join(os.tmpdir(), 'artcraft-on-'));
  const missing = mkdtempSync(path.join(os.tmpdir(), 'artcraft-off-'));
  const M = {
    word: '/artcraft/artcraft_word-0123456789.wasm',
    deck: '/artcraft/artcraft_deck-0123456789.wasm',
    fonts: ['/artcraft/fonts/SCDream4-0123456789.otf', '/artcraft/fonts/SCDream6-0123456789.otf'],
  };

  beforeAll(async () => {
    writeFileSync(path.join(installed, 'manifest.json'), JSON.stringify(M));
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', installed);
    ({ store, newDocumentId } = await import('../src/storage.js'));
    ({ issueViewToken, verifyViewToken } = await import('../src/security.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  async function addDoc(ext: string, status: 'ready' | 'processing' = 'ready') {
    const id = newDocumentId();
    await store.saveOriginal(id, ext, Buffer.from(`원본-${ext}`));
    await store.writeMeta({
      id,
      name: `시험.${ext}`,
      ext,
      status,
      pageCount: status === 'ready' ? 3 : 0,
      createdAt: new Date().toISOString(),
      ...(status === 'processing' ? { editReady: true } : {}),
    });
    return id;
  }
  const issue = (id: string, body: Record<string, unknown>) =>
    app.inject({ method: 'POST', url: `/api/documents/${id}/tokens`, payload: { viewer: '시험자', ...body } });
  const tok = (docId: string, perms: Perms, engine?: 'artcraft') =>
    issueViewToken({ docId, viewer: '시험자', perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...perms }, ...(engine ? { engine } : {}) }, 600);
  const info = async (token: string) => (await app.inject({ method: 'GET', url: `/api/view/info?token=${encodeURIComponent(token)}` })).json().data;

  test('클레임: engine은 넣은 토큰에만 있고 그대로 돌아온다', () => {
    expect(verifyViewToken(tok('d', {}, 'artcraft')).engine).toBe('artcraft');
    expect('engine' in verifyViewToken(tok('d', {}))).toBe(false);
  });

  test('토큰: docx + 내려받기 + artcraft → engine=artcraft, 클레임에 engine, 주소는 지금 그대로', async () => {
    const id = await addDoc('docx');
    const res = await issue(id, { canDownload: true, engine: 'artcraft' });
    expect(res.statusCode).toBe(200);
    const data = res.json().data;
    expect(data.engine).toBe('artcraft');
    expect(data.engineReason).toBeUndefined();
    expect(data.viewerUrl).toMatch(/^\/view\.html\?token=/);
    expect(verifyViewToken(data.token).engine).toBe('artcraft');
  });

  test('토큰: 내려받기 권한이 없으면 libreoffice + 이유(거절하지 않음)', async () => {
    const id = await addDoc('pptx');
    const data = (await issue(id, { engine: 'artcraft' })).json().data;
    expect(data.engine).toBe('libreoffice');
    expect(data.engineReason).toBe('내려받기 권한이 없어 서버 변환으로 엽니다');
    expect(data.viewerUrl).toMatch(/^\/view\.html\?token=/);
  });

  test('토큰: 안 되는 형식(pdf)은 libreoffice + 이유', async () => {
    const id = await addDoc('pdf');
    const data = (await issue(id, { canDownload: true, engine: 'artcraft' })).json().data;
    expect(data).toMatchObject({ engine: 'libreoffice', engineReason: '이 형식은 브라우저 엔진이 지원하지 않습니다' });
  });

  test('토큰: engine을 안 보내면 예전과 같다(engine=libreoffice, 이유·클레임 없음)', async () => {
    const id = await addDoc('docx');
    const data = (await issue(id, { canDownload: true })).json().data;
    expect(data.engine).toBe('libreoffice');
    expect(data.engineReason).toBeUndefined();
    expect('engine' in verifyViewToken(data.token)).toBe(false);
  });

  test('토큰: 모르는 engine 값은 400', async () => {
    const id = await addDoc('docx');
    expect((await issue(id, { engine: 'chrome' })).statusCode).toBe(400);
  });

  test('토큰·info: manifest가 없으면 libreoffice + "설치되지 않았습니다"', async () => {
    const id = await addDoc('docx');
    vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', missing);
    try {
      const data = (await issue(id, { canDownload: true, engine: 'artcraft' })).json().data;
      expect(data).toMatchObject({ engine: 'libreoffice', engineReason: '브라우저 엔진이 설치되지 않았습니다' });
      const i = await info(data.token);
      expect(i.engine).toBe('libreoffice');
      expect(i.artcraft).toBeUndefined();
    } finally {
      vi.stubEnv('DOCVIEW_ARTCRAFT_DIR', installed);
    }
  });

  test('info: artcraft 토큰이면 engine·artcraft 블록(원본 주소·해시 자산)', async () => {
    const id = await addDoc('docx');
    const t = tok(id, { canDownload: true }, 'artcraft');
    const i = await info(t);
    expect(i.engine).toBe('artcraft');
    expect(i.engineReason).toBeUndefined();
    expect(i.artcraft).toEqual({
      kind: 'word',
      ext: 'docx',
      experimental: false,
      originalUrl: `/api/view/download?token=${encodeURIComponent(t)}`,
      assets: { wasm: M.word, fonts: M.fonts },
    });
  });

  test('info: pptx → deck, doc → word 실험', async () => {
    expect((await info(tok(await addDoc('pptx'), { canDownload: true }, 'artcraft'))).artcraft).toMatchObject({ kind: 'deck', experimental: false });
    expect((await info(tok(await addDoc('doc'), { canDownload: true }, 'artcraft'))).artcraft).toMatchObject({ kind: 'word', ext: 'doc', experimental: true });
  });

  test('info: 클레임에 engine이 있어도 내려받기 권한이 없으면 libreoffice + 이유, artcraft 블록 없음', async () => {
    const i = await info(tok(await addDoc('docx'), {}, 'artcraft'));
    expect(i.engine).toBe('libreoffice');
    expect(i.engineReason).toBe('내려받기 권한이 없어 서버 변환으로 엽니다');
    expect(i.artcraft).toBeUndefined();
  });

  test('info: 기본 토큰 → engine=libreoffice, 이유·artcraft 없음', async () => {
    const i = await info(tok(await addDoc('docx'), { canDownload: true }));
    expect(i.engine).toBe('libreoffice');
    expect(i.engineReason).toBeUndefined();
    expect(i.artcraft).toBeUndefined();
  });

  test('첫 변환 중(editReady) docx도 artcraft로 열리고 원본을 바로 받는다(서버 변환을 기다리지 않음)', async () => {
    const id = await addDoc('docx', 'processing');
    const i = await info(tok(id, { canDownload: true }, 'artcraft'));
    expect(i.engine).toBe('artcraft');
    expect(i.processing).toBe(true);
    const res = await app.inject({ method: 'GET', url: i.artcraft.originalUrl });
    expect(res.statusCode).toBe(200);
    expect(res.rawPayload.toString('utf8')).toBe('원본-docx');
  });
});
```

- [ ] **Step 6: 시험이 실패하는지 보기**

Run: `npx vitest run tests/artcraftRoutes.test.ts`
Expected: FAIL — 첫 시험에서 `engine`이 undefined(클레임에 안 들어감), 토큰 시험들은 `engine` 400(`additionalProperties: false`) 등.

- [ ] **Step 7: 클레임 — `src/types.ts`·`src/security.ts`**

`src/types.ts`의 `ViewClaims`를 바꾼다:

```ts
export interface ViewClaims {
  readonly docId: string;
  readonly viewer: string;
  readonly perms: Permissions;
  /** 토큰 요청에 engine:'artcraft'를 보낸 경우만 — "요청한" 엔진. 실제 판정은 /api/view/info가 매번 한다(artcraftEngine.ts) */
  readonly engine?: 'artcraft';
}
```

`src/security.ts`의 `issueViewToken` 본문:

```ts
export function issueViewToken(claims: ViewClaims, ttlSeconds: number): string {
  return jwt.sign(
    { docId: claims.docId, viewer: claims.viewer, perms: claims.perms, ...(claims.engine ? { engine: claims.engine } : {}) },
    jwtSecret,
    { expiresIn: ttlSeconds, algorithm: JWT_ALGORITHM },
  );
}
```

`verifyViewToken`에서 구조 분해와 반환을 바꾼다:

```ts
  const { docId, viewer, perms, engine } = payload as Record<string, unknown>;
```

```ts
  return { docId, viewer, perms: permissions, ...(engine === 'artcraft' ? { engine: 'artcraft' as const } : {}) };
```

- [ ] **Step 8: 토큰 라우트 — `src/routes/documents.ts`**

import 묶음에 더한다:

```ts
import { engineFor, sourceExtOf, type Engine } from '../artcraftEngine.js';
import { readArtcraftManifest } from '../artcraftManifest.js';
```

`tokenBodySchema.properties`의 `pdfEditor` 다음에 더한다:

```ts
    // 여는 엔진(설계 2026-11-21 §3.1) — artcraft는 브라우저 wasm 엔진. 안 되는 경우는 거절하지 않고 libreoffice + engineReason
    engine: { type: 'string', enum: ['libreoffice', 'artcraft'], default: 'libreoffice' },
```

`TokenBody`에 더한다:

```ts
  readonly engine: Engine;
```

토큰 핸들러에서 구조 분해를 바꾸고 판정을 더한다(`const { viewer, ... } = request.body;` 줄 교체, 그 아래 `effectiveCanEdit` 계산 다음에 `decision`):

```ts
      const { viewer, canPrint, canDownload, canCopy, canEdit, pdfEditor, engine } = request.body;
```

```ts
      const decision = engineFor(sourceExtOf(meta), { canDownload }, engine, readArtcraftManifest() !== null);
```

`issueViewToken({ docId: meta.id, viewer, perms: {...} }, …)` 첫 인자 객체의 `perms: { ... },` 다음 줄에 더한다:

```ts
          ...(engine === 'artcraft' ? { engine: 'artcraft' as const } : {}),
```

감사 기록 `detail` 문자열을 바꾼다:

```ts
        detail: `print=${canPrint} download=${canDownload} copy=${canCopy} edit=${effectiveCanEdit} engine=${decision.engine}${callerProject ? ` caller=${callerProject}` : ''}`,
```

응답 `data`의 `expiresInSeconds: config.tokenTtlSeconds,` 다음에 더한다:

```ts
          engine: decision.engine,
          ...(decision.reason ? { engineReason: decision.reason } : {}),
```

- [ ] **Step 9: `/api/view/info` — `src/routes/viewer.ts`**

import 묶음에 더한다:

```ts
import { artcraftInfo, engineFor, sourceExtOf } from '../artcraftEngine.js';
import { readArtcraftManifest } from '../artcraftManifest.js';
```

`/api/view/info` 핸들러에서 `const slideSize = …;` 다음, `return {` 앞에 더한다:

```ts
      // 여는 엔진 — 클레임은 "요청"만 담고 판정은 매번(설계 2026-11-21 §3.2). 아트크래프트는 서버 변환 없이 원본을 브라우저가 그린다
      const manifest = readArtcraftManifest();
      const ext = sourceExtOf(view.meta);
      const engine = engineFor(ext, view.claims.perms, view.claims.engine, manifest !== null);
      const artcraft = engine.engine === 'artcraft' && manifest ? artcraftInfo(ext, request.query.token!, manifest) : undefined;
```

응답 `data`의 `processing: view.meta.status === 'processing',` 다음에 더한다:

```ts
          engine: artcraft ? 'artcraft' : 'libreoffice',
          ...(engine.reason ? { engineReason: engine.reason } : {}),
          ...(artcraft ? { artcraft } : {}),
```

- [ ] **Step 10: 시험이 통과하는지 보기**

Run: `npx vitest run tests/artcraftRoutes.test.ts tests/artcraftEngine.test.ts tests/security.test.ts tests/viewShellRoutes.test.ts tests/routes-documents.test.ts`
Expected: PASS(모두). 기존 security·viewShellRoutes·routes-documents가 그대로 통과해야 한다(토큰 모양은 engine을 안 보내면 예전과 같다).

- [ ] **Step 11: 타입 검사·커밋**

```bash
npm run typecheck
git add src/artcraftEngine.ts src/types.ts src/security.ts src/routes/documents.ts src/routes/viewer.ts tests/artcraftEngine.test.ts tests/artcraftRoutes.test.ts
git commit -m "아트크래프트 2: engineFor 판정표, 토큰 engine(클레임은 요청만)·engineReason, /api/view/info engine·artcraft 블록"
```

---

### Task 3: 화면 — Worker와 `ArtcraftSource`

**Files:**
- Create: `src/client/view/artcraftLogic.ts`, `src/client/view/artcraftWorker.ts`, `src/client/view/artcraftSource.ts`
- Modify: `src/client/view/types.ts` (ViewInfo·BodyKind·ArtcraftInfo), `src/client/view/pagesBody.ts:20` (PageSource.kind), `scripts/build-editor.mjs` (entryPoints), `.gitignore` (Worker 번들)
- Test: `tests/artcraftLogic.test.ts`

**Interfaces:**
- Consumes: wasm C ABI(Task 1), `info.artcraft: ArtcraftInfo`·`info.engine`(Task 2 응답 모양), `PageSource`·`DrawOpts`(`src/client/view/pagesBody.ts`), `canvasPixelRatio`·`SEARCH_LIMIT`(`src/client/view/logic.ts`).
- Produces:
  - `src/client/view/types.ts`: `interface ArtcraftInfo { kind: 'word' | 'deck'; ext: string; experimental: boolean; originalUrl: string; assets: { wasm: string; fonts: readonly string[] } }`; `ViewInfo.engine?: 'libreoffice' | 'artcraft'`, `ViewInfo.engineReason?: string`, `ViewInfo.artcraft?: ArtcraftInfo`; `BodyKind`에 `'pages-artcraft'`.
  - `src/client/view/artcraftSource.ts`: `class ArtcraftUnavailable extends Error`; `OPEN_TIMEOUT_MS = 60_000`; `RENDER_TIMEOUT_MS = 20_000`; `openArtcraftSource(info: ViewInfo): Promise<ArtcraftSource>`(실패는 모두 `ArtcraftUnavailable`); `class ArtcraftSource implements PageSource` (`kind = 'pages-artcraft'`).
  - 빌드 산출물 `public/js/artcraftWorker.bundle.js`(module worker).

- [ ] **Step 1: 순수 로직 실패하는 시험 쓰기**

`tests/artcraftLogic.test.ts`:

```ts
import { describe, expect, test, vi } from 'vitest';
import {
  ByteLru,
  JobQueue,
  drawPlan,
  pageCountSuspicious,
  quickPixelWidth,
  runHits,
  runSpanGeom,
  stubImports,
  targetPixelWidth,
  type PageText,
} from '../src/client/view/artcraftLogic.js';

const LETTER = { w: 612, h: 792 };

describe('배율', () => {
  test('요청 화소 폭 = 쪽 CSS 폭 × dpr(캔버스 상한 안)', () => {
    expect(targetPixelWidth(800, 2, LETTER)).toBe(1600);
    expect(targetPixelWidth(800, 1, LETTER)).toBe(800);
  });
  test('아주 크게 확대해도 캔버스 한 장 1677만 화소를 넘지 않는다', () => {
    const px = targetPixelWidth(5000, 2, LETTER);
    expect(px).toBeGreaterThan(3500);
    expect(px * Math.ceil(px * (792 / 612))).toBeLessThanOrEqual(16_777_216 * 1.002);
  });
  test('1배 그림 화소 폭 = 쪽 폭 pt', () => {
    expect(quickPixelWidth(LETTER)).toBe(612);
    expect(quickPixelWidth({ w: 0.2, h: 1 })).toBe(1);
  });
});

describe('두 단계 그리기 판단', () => {
  test('같은 화소 폭이 이미 있으면 그대로', () => {
    expect(drawPlan([1600], 1600, 612)).toBe('cached');
  });
  test('선명한 그림이 아직 없고 1.5배를 넘으면 1배 먼저', () => {
    expect(drawPlan([], 1600, 612)).toBe('two-stage');
    expect(drawPlan([612], 1600, 612)).toBe('two-stage');
  });
  test('이미 선명한 그림이 있으면(확대 바꿈) 바로 요청 배율', () => {
    expect(drawPlan([1200], 1600, 612)).toBe('direct');
  });
  test('요청이 1.5배 이하면 바로', () => {
    expect(drawPlan([], 900, 612)).toBe('direct');
  });
});

describe('찾기 — run 상자 안에서 글자 수 비율로 자른 사각형', () => {
  const text: PageText = {
    w: 400,
    h: 800,
    runs: [
      { t: '가나다라', x0: 100, x1: 140, y: 200, s: 10 },
      { t: 'Hello World', x0: 0, x1: 110, y: 400, s: 20 },
    ],
  };
  test('한글 가운데 두 글자', () => {
    expect(runHits(3, text, '나다')).toEqual([{ page: 3, rects: [{ x: 0.275, y: 0.24, w: 0.05, h: 0.0125 }] }]);
  });
  test('대소문자 무시, 한 run에 여러 번', () => {
    const hits = runHits(1, text, 'O');
    expect(hits).toHaveLength(2);
    expect(hits[0].rects[0].x).toBeCloseTo(40 / 400);
    expect(runHits(1, text, 'world')[0].rects[0]).toMatchObject({ y: (400 - 16) / 800, h: 20 / 800 });
  });
  test('빈 검색어·없는 글자는 결과 없음', () => {
    expect(runHits(1, text, '  ')).toEqual([]);
    expect(runHits(1, text, '없음')).toEqual([]);
  });
});

describe('글자 층 좌표', () => {
  test('쪽 비율 % 위치, CSS 화소 글자 크기·폭', () => {
    expect(runSpanGeom({ t: '가나', x0: 100, x1: 140, y: 200, s: 10 }, { w: 400, h: 800 }, 800)).toEqual({ left: 25, top: 24, fontPx: 20, widthPx: 80 });
  });
});

describe('.doc(실험) 쪽 수 전환 규칙', () => {
  test('서버 쪽 수와 2쪽 넘게 다르면 전환, 2쪽까지는 그대로', () => {
    expect(pageCountSuspicious(true, 5, 8)).toBe(true);
    expect(pageCountSuspicious(true, 5, 7)).toBe(false);
    expect(pageCountSuspicious(true, 9, 6)).toBe(true);
  });
  test('실험이 아니거나 서버 쪽 수를 모르면(0) 보지 않는다', () => {
    expect(pageCountSuspicious(false, 1, 9)).toBe(false);
    expect(pageCountSuspicious(true, 5, 0)).toBe(false);
  });
});

describe('그림 LRU(바이트 상한)', () => {
  test('넘으면 오래 안 쓴 것부터 버리고, 방금 쓴 것은 남긴다', () => {
    const lru = new ByteLru<{ n: number }>(10, (v) => v.n);
    lru.set('1:100', { n: 4 });
    lru.set('2:100', { n: 4 });
    expect(lru.get('1:100')).toEqual({ n: 4 }); // 1을 최근으로
    lru.set('3:100', { n: 4 }); // 12 > 10 → 2를 버림
    expect(lru.keys()).toEqual(['1:100', '3:100']);
    expect(lru.bytes).toBe(8);
  });
  test('하나가 상한보다 커도 그것은 남긴다', () => {
    const lru = new ByteLru<{ n: number }>(10, (v) => v.n);
    lru.set('a', { n: 3 });
    lru.set('b', { n: 50 });
    expect(lru.keys()).toEqual(['b']);
  });
  test('같은 키를 다시 넣으면 크기를 바꿔 센다', () => {
    const lru = new ByteLru<{ n: number }>(100, (v) => v.n);
    lru.set('a', { n: 30 });
    lru.set('a', { n: 10 });
    expect(lru.bytes).toBe(10);
  });
});

describe('작업 줄 — 보이는 쪽(급함) 먼저, 호출 순서 유지, 취소는 빼기', () => {
  test('급한 것이 먼저, 같은 급끼리는 넣은 순서', () => {
    const q = new JobQueue<string>();
    q.push('선명하게 1', false);
    q.push('1쪽', true);
    q.push('2쪽', true);
    expect([q.shift(), q.shift(), q.shift(), q.shift()]).toEqual(['1쪽', '2쪽', '선명하게 1', undefined]);
  });
  test('remove는 줄에 있던 것만 true', () => {
    const q = new JobQueue<string>();
    q.push('a', true);
    q.push('b', false);
    expect(q.remove('b')).toBe(true);
    expect(q.remove('b')).toBe(false);
    expect(q.size).toBe(1);
  });
});

describe('바깥 함수 채우기(wasm-bindgen 없이)', () => {
  const descs: WebAssembly.ModuleImportDescriptor[] = [
    { module: 'wbg', name: '__wbg_now_2c95c9de', kind: 'function' },
    { module: 'wbg', name: '__wbg_getTimezoneOffset_6b5752021c499c47', kind: 'function' },
    { module: 'wbg', name: '__wbg_unknown_99', kind: 'function' },
    { module: 'wbg', name: '__wbindgen_throw', kind: 'function' },
    { module: 'env', name: 'memory', kind: 'memory' },
  ];
  test('시각·시간대만 진짜, 나머지는 0 + 이름마다 경고 한 번', () => {
    const warn = vi.fn();
    const imports = stubImports(descs, warn) as unknown as Record<string, Record<string, () => number>>;
    expect(Math.abs(imports.wbg.__wbg_now_2c95c9de() - Date.now())).toBeLessThan(1000);
    expect(imports.wbg.__wbg_getTimezoneOffset_6b5752021c499c47()).toBe(new Date().getTimezoneOffset());
    expect(imports.wbg.__wbg_unknown_99()).toBe(0); // "unknown" 안의 now에 걸리지 않는다
    imports.wbg.__wbg_unknown_99();
    expect(imports.wbg.__wbindgen_throw()).toBe(0);
    expect(warn.mock.calls).toEqual([['__wbg_unknown_99'], ['__wbindgen_throw']]);
    expect(imports.env).toBeUndefined();
  });
});
```

- [ ] **Step 2: 시험이 실패하는지 보기**

Run: `npx vitest run tests/artcraftLogic.test.ts`
Expected: FAIL — `../src/client/view/artcraftLogic.js`를 찾지 못함.

- [ ] **Step 3: `src/client/view/artcraftLogic.ts` 쓰기**

```ts
/**
 * 아트크래프트 원천의 순수 계산과 Worker 주고받기 모양 — DOM 없이 시험한다(tests/artcraftLogic.test.ts).
 * 엔진 글자 JSON·C ABI는 tools/artcraft-shell/abi.rs, 설계 docs/superpowers/specs/2026-11-21-artcraft-browser-engine-design.md §4.
 */
import { canvasPixelRatio } from './logic.js';
import type { NormRect, PageSize, SearchHit } from './types.js';

/** 엔진 sp_text의 글자 조각 — pt, y 아래로, y는 기준선, s는 글자 크기 */
export interface TextRun {
  readonly t: string;
  readonly x0: number;
  readonly x1: number;
  readonly y: number;
  readonly s: number;
}
export interface PageText {
  readonly w: number;
  readonly h: number;
  readonly runs: readonly TextRun[];
}

export type WorkerCall =
  | { readonly cmd: 'init'; readonly wasm: string; readonly fonts: readonly string[]; readonly original: string; readonly name: string }
  | { readonly cmd: 'render'; readonly page: number; readonly scale: number }
  | { readonly cmd: 'text'; readonly page: number };
export type WorkerRequest = WorkerCall & { readonly id: number };
export interface InitReply { readonly id: number; readonly count: number; readonly sizes: PageSize[] }
export interface RenderReply { readonly id: number; readonly w: number; readonly h: number; readonly rgba: ArrayBuffer }
export interface TextReply { readonly id: number; readonly text: PageText }
export interface ErrorReply { readonly id: number; readonly error: string }
export type WorkerReply = InitReply | RenderReply | TextReply | ErrorReply;

/** 글자 상자 위 = 기준선 − 0.8 × 글자 크기, 높이 = 글자 크기(엔진은 기준선과 크기만 준다) */
const ASCENT = 0.8;
/** 요청 화소 폭이 쪽 폭 pt의 이만큼을 넘어야 두 단계로 그린다(그 아래면 1배와 차이가 작다) */
const TWO_STAGE_FACTOR = 1.5;

/** 그릴 화소 폭 — CSS 폭 × dpr, 캔버스 한 장 상한(MAX_CANVAS_PIXELS)을 넘지 않게 */
export function targetPixelWidth(cssWidth: number, dpr: number, size: PageSize): number {
  const cssHeight = cssWidth * (size.h / size.w);
  return Math.max(1, Math.ceil(cssWidth * canvasPixelRatio(cssWidth, cssHeight, dpr)));
}

/** 1배(scale = 1, 쪽 폭 pt = 화소) 그림의 화소 폭 */
export function quickPixelWidth(size: PageSize): number {
  return Math.max(1, Math.round(size.w));
}

/**
 * 이 쪽을 어떻게 그릴지 — cachedPx는 이 쪽에 대해 보관 중인 그림의 화소 폭들.
 * 같은 폭이 있으면 그대로, 선명한(1.5배 넘는) 그림이 아직 없으면 1배 먼저 보이고 이어서 요청 배율(설계 §4.1), 그 밖은 바로.
 */
export function drawPlan(cachedPx: readonly number[], wantPx: number, pageWidthPt: number): 'cached' | 'direct' | 'two-stage' {
  if (cachedPx.includes(wantPx)) return 'cached';
  const sharp = TWO_STAGE_FACTOR * pageWidthPt;
  if (wantPx > sharp && !cachedPx.some((px) => px > sharp)) return 'two-stage';
  return 'direct';
}

/**
 * 한 쪽에서 찾기(대소문자 무시) — run 상자 안을 글자 수 비율로 잘라 쪽 비율 사각형을 만든다.
 * shortcut: matches that cross two runs are not found; join neighbouring runs if users report misses.
 */
export function runHits(page: number, text: PageText, q: string): SearchHit[] {
  const needle = q.trim().toLowerCase();
  if (!needle || text.w <= 0 || text.h <= 0) return [];
  const hits: SearchHit[] = [];
  for (const run of text.runs) {
    const hay = run.t.toLowerCase();
    // 소문자로 바꾸며 길이가 바뀌는 글자(예: İ)가 있으면 run 전체를 칠한다
    const exact = hay.length === run.t.length;
    for (let at = hay.indexOf(needle); at >= 0; at = hay.indexOf(needle, at + needle.length)) {
      const a = exact ? at / run.t.length : 0;
      const b = exact ? (at + needle.length) / run.t.length : 1;
      const x0 = run.x0 + (run.x1 - run.x0) * a;
      const x1 = run.x0 + (run.x1 - run.x0) * b;
      const rect: NormRect = { x: x0 / text.w, y: (run.y - run.s * ASCENT) / text.h, w: (x1 - x0) / text.w, h: run.s / text.h };
      hits.push({ page, rects: [rect] });
    }
  }
  return hits;
}

/** 글자 층 span 하나의 자리 — left·top은 쪽 비율 %, 글자 크기·폭은 지금 쪽 CSS 폭 기준 화소 */
export function runSpanGeom(run: TextRun, size: PageSize, cssWidth: number): { left: number; top: number; fontPx: number; widthPx: number } {
  const scale = cssWidth / size.w;
  return {
    left: (run.x0 / size.w) * 100,
    top: ((run.y - run.s * ASCENT) / size.h) * 100,
    fontPx: run.s * scale,
    widthPx: (run.x1 - run.x0) * scale,
  };
}

/** .doc(실험): 서버 변환 쪽 수(알 때만, >0)와 2쪽 넘게 다르면 의심 → 리브레오피스로(설계 §4.2) */
export function pageCountSuspicious(experimental: boolean, engineCount: number, serverCount: number): boolean {
  return experimental && serverCount > 0 && Math.abs(engineCount - serverCount) > 2;
}

/** 바이트 상한 LRU — 쪽 그림(ImageData)을 보관한다. 방금 넣은 것은 상한보다 커도 남긴다 */
export class ByteLru<V> {
  private readonly map = new Map<string, V>();
  private total = 0;
  constructor(
    private readonly maxBytes: number,
    private readonly sizeOf: (value: V) => number,
  ) {}
  get bytes(): number {
    return this.total;
  }
  get(key: string): V | undefined {
    const value = this.map.get(key);
    if (value !== undefined) {
      this.map.delete(key);
      this.map.set(key, value);
    }
    return value;
  }
  set(key: string, value: V): void {
    const old = this.map.get(key);
    if (old !== undefined) {
      this.total -= this.sizeOf(old);
      this.map.delete(key);
    }
    this.map.set(key, value);
    this.total += this.sizeOf(value);
    for (const [k, v] of this.map) {
      if (this.total <= this.maxBytes || k === key) break;
      this.map.delete(k);
      this.total -= this.sizeOf(v);
    }
  }
  keys(): string[] {
    return [...this.map.keys()];
  }
}

/** Worker 작업 줄 — 급한 것(보이는 쪽) 먼저, 같은 급끼리는 넣은 순서(PagesBody가 보이는 쪽부터 부른다) */
export class JobQueue<T> {
  private readonly urgent: T[] = [];
  private readonly later: T[] = [];
  get size(): number {
    return this.urgent.length + this.later.length;
  }
  push(job: T, urgent: boolean): void {
    (urgent ? this.urgent : this.later).push(job);
  }
  shift(): T | undefined {
    return this.urgent.shift() ?? this.later.shift();
  }
  remove(job: T): boolean {
    for (const list of [this.urgent, this.later]) {
      const at = list.indexOf(job);
      if (at >= 0) {
        list.splice(at, 1);
        return true;
      }
    }
    return false;
  }
}

/**
 * wasm-bindgen 없이 돌릴 때 엔진이 요구하는 바깥 함수(spike §7.6) — 시각·시간대만 진짜, 나머지는 0을 돌려주고 이름마다 한 번 경고.
 * 시험 중 실제로 불린 것은 없었다(불리면 엔진 바인딩 문제).
 */
export function stubImports(descriptors: readonly WebAssembly.ModuleImportDescriptor[], warn: (name: string) => void): WebAssembly.Imports {
  const imports: Record<string, Record<string, () => number>> = {};
  const warned = new Set<string>();
  for (const d of descriptors) {
    if (d.kind !== 'function') continue;
    const fns = (imports[d.module] ??= {});
    if (/(^|_)now(_|$)/.test(d.name)) fns[d.name] = () => Date.now();
    else if (/getTimezoneOffset/.test(d.name)) fns[d.name] = () => new Date().getTimezoneOffset();
    else
      fns[d.name] = () => {
        if (!warned.has(d.name)) {
          warned.add(d.name);
          warn(d.name);
        }
        return 0;
      };
  }
  return imports;
}
```

- [ ] **Step 4: 시험이 통과하는지 보기**

Run: `npx vitest run tests/artcraftLogic.test.ts`
Expected: PASS.

- [ ] **Step 5: 타입 — `src/client/view/types.ts`·`pagesBody.ts`**

`types.ts`의 `ViewInfo` 앞에 더한다:

```ts
/** /api/view/info의 artcraft 블록(서버 src/artcraftEngine.ts ArtcraftInfo와 같은 모양) */
export interface ArtcraftInfo {
  readonly kind: 'word' | 'deck';
  readonly ext: string;
  readonly experimental: boolean;
  readonly originalUrl: string;
  readonly assets: { readonly wasm: string; readonly fonts: readonly string[] };
}
```

`ViewInfo`의 `processing` 다음에 더한다:

```ts
  /** 여는 엔진 — artcraft면 브라우저 wasm 엔진(artcraft 블록이 함께 온다) */
  readonly engine?: 'libreoffice' | 'artcraft';
  /** 아트크래프트를 요청했지만 서버 변환으로 여는 이유(한 줄 안내로 보인다) */
  readonly engineReason?: string;
  readonly artcraft?: ArtcraftInfo;
```

`BodyKind`를 바꾼다:

```ts
export type BodyKind = 'pages-png' | 'pages-pdfjs' | 'pages-artcraft' | 'grid' | 'cad' | 'frame';
```

`pagesBody.ts`의 `PageSource.kind` 줄을 바꾼다(그 밖은 그대로):

```ts
  readonly kind: 'pages-png' | 'pages-pdfjs' | 'pages-artcraft';
```

- [ ] **Step 6: Worker `src/client/view/artcraftWorker.ts`**

```ts
/**
 * 아트크래프트 엔진 Worker(문서 하나 = wasm 인스턴스 하나) — DOM·Canvas 없음, RGBA 버퍼를 넘겨준다(spike §7.2).
 * 받는 것은 같은 출처의 정적 자산(/artcraft/*)과 원본(/api/view/download)뿐이다(설계 §6). 작업은 artcraftSource.ts가 하나씩 보낸다.
 */
import { stubImports, type PageText, type WorkerReply, type WorkerRequest } from './artcraftLogic.js';

interface Engine {
  readonly memory: WebAssembly.Memory;
  sp_alloc(n: number): number;
  sp_free(ptr: number, n: number): void;
  sp_add_font(ptr: number, len: number, aliasPtr: number, aliasLen: number): number;
  sp_open(namePtr: number, nameLen: number, ptr: number, len: number): number;
  sp_count(h: number): number;
  sp_page_w(h: number, i: number): number;
  sp_page_h(h: number, i: number): number;
  sp_render(h: number, i: number, scaleMilli: number): number;
  sp_text(h: number, i: number): number;
  sp_out_ptr(): number;
  sp_out_len(): number;
  sp_out_w(): number;
  sp_out_h(): number;
  sp_out_clear(): void;
  sp_msg_ptr(): number;
  sp_msg_len(): number;
}

// tsconfig는 DOM 타입만 싣는다 — Worker 전역에서 쓰는 두 가지만 좁게 적는다
const scope = self as unknown as {
  onmessage: ((event: MessageEvent<WorkerRequest>) => void) | null;
  postMessage(message: WorkerReply, transfer?: Transferable[]): void;
};
const encoder = new TextEncoder();
const decoder = new TextDecoder();
let engine: Engine | null = null;
let docHandle = 0;

async function bytesOf(url: string): Promise<Uint8Array> {
  const res = await fetch(url, { credentials: 'same-origin' });
  if (!res.ok) throw new Error(`받지 못했습니다(${res.status}): ${url.split('?')[0]}`);
  return new Uint8Array(await res.arrayBuffer());
}

function put(e: Engine, bytes: Uint8Array): [number, number] {
  const ptr = e.sp_alloc(bytes.length);
  new Uint8Array(e.memory.buffer, ptr, bytes.length).set(bytes);
  return [ptr, bytes.length];
}

function lastMessage(e: Engine): string {
  return decoder.decode(new Uint8Array(e.memory.buffer, e.sp_msg_ptr(), e.sp_msg_len()));
}

function opened(): Engine {
  if (!engine || !docHandle) throw new Error('문서가 열려 있지 않습니다');
  return engine;
}

async function init(m: Extract<WorkerRequest, { cmd: 'init' }>): Promise<WorkerReply> {
  const wasm = fetch(m.wasm, { credentials: 'same-origin' }).then((res) => {
    if (!res.ok) throw new Error(`엔진을 받지 못했습니다(${res.status})`);
    return res;
  });
  // 컴파일은 받는 동안 함께(compileStreaming — 크롬은 결과를 캐시해 두 번째부터 수십 ms), 글꼴·원본도 동시에 받는다
  const [module, fonts, original] = await Promise.all([WebAssembly.compileStreaming(wasm), Promise.all(m.fonts.map(bytesOf)), bytesOf(m.original)]);
  const instance = await WebAssembly.instantiate(
    module,
    stubImports(WebAssembly.Module.imports(module), (name) => console.warn(`[artcraft] 채우지 않은 바깥 함수가 불렸습니다: ${name}`)),
  );
  const e = instance.exports as unknown as Engine;
  for (const font of fonts) {
    const [ptr, n] = put(e, font);
    e.sp_add_font(ptr, n, 0, 0);
    e.sp_free(ptr, n);
  }
  const [namePtr, nameLen] = put(e, encoder.encode(m.name));
  const [ptr, n] = put(e, original);
  const h = e.sp_open(namePtr, nameLen, ptr, n);
  e.sp_free(ptr, n);
  e.sp_free(namePtr, nameLen);
  if (!h) throw new Error(lastMessage(e) || '문서를 열지 못했습니다');
  engine = e;
  docHandle = h;
  const count = e.sp_count(h);
  const sizes = Array.from({ length: count }, (_, i) => ({ w: e.sp_page_w(h, i), h: e.sp_page_h(h, i) }));
  return { id: m.id, count, sizes };
}

function render(m: Extract<WorkerRequest, { cmd: 'render' }>): { id: number; w: number; h: number; rgba: ArrayBuffer } {
  const e = opened();
  if (e.sp_render(docHandle, m.page - 1, Math.round(m.scale * 1000))) throw new Error(lastMessage(e) || '쪽을 그리지 못했습니다');
  const w = e.sp_out_w();
  const h = e.sp_out_h();
  const ptr = e.sp_out_ptr();
  const rgba = e.memory.buffer.slice(ptr, ptr + w * h * 4);
  e.sp_out_clear(); // 큰 RGBA를 엔진 쪽에서 바로 놓는다
  return { id: m.id, w, h, rgba };
}

function text(page: number): PageText {
  const e = opened();
  if (e.sp_text(docHandle, page - 1)) throw new Error(lastMessage(e) || '글자를 꺼내지 못했습니다');
  const json = decoder.decode(new Uint8Array(e.memory.buffer, e.sp_out_ptr(), e.sp_out_len()));
  e.sp_out_clear();
  return JSON.parse(json) as PageText;
}

async function answer(m: WorkerRequest): Promise<void> {
  try {
    if (m.cmd === 'init') scope.postMessage(await init(m));
    else if (m.cmd === 'render') {
      const reply = render(m);
      scope.postMessage(reply, [reply.rgba]);
    } else scope.postMessage({ id: m.id, text: text(m.page) });
  } catch (error) {
    scope.postMessage({ id: m.id, error: error instanceof Error ? error.message : String(error) });
  }
}

scope.onmessage = (event) => void answer(event.data);
```

- [ ] **Step 7: 원천 `src/client/view/artcraftSource.ts`**

```ts
/**
 * 아트크래프트 원천(설계 §4.1) — 원본을 브라우저 Worker 안 ArtCraft wasm 엔진이 그린다. 서버 변환을 부르지 않는다.
 * 열기에 실패하면 ArtcraftUnavailable — view.ts가 리브레오피스 원천으로 바꾼다. 열린 뒤 그리기 오류는 PagesBody.onDrawError로 간다.
 * view.ts가 동적 import로만 부른다.
 */
import {
  ByteLru,
  JobQueue,
  drawPlan,
  pageCountSuspicious,
  quickPixelWidth,
  runHits,
  runSpanGeom,
  targetPixelWidth,
  type InitReply,
  type PageText,
  type RenderReply,
  type TextReply,
  type WorkerCall,
  type WorkerReply,
  type WorkerRequest,
} from './artcraftLogic.js';
import { SEARCH_LIMIT } from './logic.js';
import type { DrawOpts, PageSource } from './pagesBody.js';
import type { PageSize, SearchHit, ViewInfo } from './types.js';

/** 브라우저 엔진으로 열 수 없음 — view.ts가 리브레오피스로 넘어간다 */
export class ArtcraftUnavailable extends Error {}

/** 열기 = wasm·글꼴·원본 받기 + 엔진 open(설계에 값이 없어 이 계획이 정함 — 느린 회선의 4 MB 받기를 감안) */
export const OPEN_TIMEOUT_MS = 60_000;
/** 첫 쪽(과 그 뒤 모든 요청)이 이 안에 끝나지 않으면 Worker를 끄고 리브레오피스로(설계 §4.2·§6) */
export const RENDER_TIMEOUT_MS = 20_000;
/** 쪽 그림 보관 상한 — 1600×2071 한 장이 약 13 MB라 보이는 쪽 ±2 정도 */
const IMAGE_CACHE_BYTES = 64 * 1024 * 1024;
const WORKER_URL = '/js/artcraftWorker.bundle.js';

const abortError = () => new DOMException('그리기를 멈췄습니다.', 'AbortError');

interface Job {
  readonly msg: WorkerCall;
  readonly timeoutMs: number;
  readonly signal?: AbortSignal;
  readonly resolve: (reply: WorkerReply) => void;
  readonly reject: (error: unknown) => void;
  id: number;
}

/** Worker에 한 번에 하나씩 보낸다(엔진은 한 줄) — 급한 것 먼저, 취소된 것은 줄에서 뺀다, 시간 넘기면 Worker를 끈다 */
export class WorkerClient {
  private readonly queue = new JobQueue<Job>();
  private busy: Job | null = null;
  private nextId = 1;
  private dead: Error | null = null;
  private timer: ReturnType<typeof setTimeout> | null = null;

  constructor(private readonly worker: Worker) {
    worker.onmessage = (event: MessageEvent<WorkerReply>) => this.finish(event.data);
    worker.onerror = (event) => {
      event.preventDefault();
      this.die(new ArtcraftUnavailable(event.message || '브라우저 엔진이 멈췄습니다'));
    };
  }

  call<R extends WorkerReply>(msg: WorkerCall, timeoutMs: number, urgent = true, signal?: AbortSignal): Promise<R> {
    if (this.dead) return Promise.reject(this.dead);
    if (signal?.aborted) return Promise.reject(abortError());
    return new Promise<R>((resolve, reject) => {
      const job: Job = { msg, timeoutMs, signal, resolve: resolve as (reply: WorkerReply) => void, reject, id: 0 };
      signal?.addEventListener(
        'abort',
        () => {
          if (this.queue.remove(job)) reject(abortError());
        },
        { once: true },
      );
      this.queue.push(job, urgent);
      this.pump();
    });
  }

  terminate(): void {
    this.die(new ArtcraftUnavailable('브라우저 엔진을 닫았습니다'));
  }

  private pump(): void {
    if (this.busy || this.dead) return;
    const job = this.queue.shift();
    if (!job) return;
    job.id = this.nextId++;
    this.busy = job;
    this.timer = setTimeout(() => this.die(new ArtcraftUnavailable('브라우저 엔진이 제때 답하지 않았습니다')), job.timeoutMs);
    const request: WorkerRequest = { ...job.msg, id: job.id };
    this.worker.postMessage(request);
  }

  private finish(reply: WorkerReply): void {
    const job = this.busy;
    if (!job || reply.id !== job.id) return;
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = null;
    this.busy = null;
    if (job.signal?.aborted) job.reject(abortError());
    else if ('error' in reply) job.reject(new Error(reply.error));
    else job.resolve(reply);
    this.pump();
  }

  private die(error: Error): void {
    if (this.dead) return;
    this.dead = error;
    if (this.timer !== null) clearTimeout(this.timer);
    this.timer = null;
    this.worker.terminate();
    this.busy?.reject(error);
    this.busy = null;
    for (let job = this.queue.shift(); job; job = this.queue.shift()) job.reject(error);
  }
}

function paint(canvas: HTMLCanvasElement, img: ImageData): void {
  canvas.width = img.width;
  canvas.height = img.height;
  const ctx = canvas.getContext('2d');
  if (!ctx) throw new Error('쪽을 그릴 캔버스를 만들지 못했습니다.');
  ctx.putImageData(img, 0, 0);
}

let measureCtx: CanvasRenderingContext2D | null = null;
function naturalWidth(text: string, fontPx: number): number {
  measureCtx ??= document.createElement('canvas').getContext('2d');
  if (!measureCtx) return 0;
  measureCtx.font = `${fontPx}px sans-serif`;
  return measureCtx.measureText(text).width;
}

export class ArtcraftSource implements PageSource {
  readonly kind = 'pages-artcraft' as const;
  private readonly images = new ByteLru<ImageData>(IMAGE_CACHE_BYTES, (img) => img.data.byteLength);
  private readonly texts = new Map<number, Promise<PageText>>();

  constructor(
    private readonly client: WorkerClient,
    private readonly sizes: readonly PageSize[],
    /** 글자 층은 복사 권한(canCopy)이 있을 때만(설계 §4.1) */
    readonly textSelectable: boolean,
  ) {}

  get pageCount(): number {
    return this.sizes.length;
  }

  sizeOf(page: number): PageSize | null {
    return this.sizes[page - 1] ?? null;
  }

  qualityKey(_page: number, opts: DrawOpts): number {
    return Math.ceil(opts.cssWidth * opts.dpr);
  }

  /** 열 때 1쪽 1배를 미리 — 첫 쪽 20초 제한을 여기서 재고, 첫 draw는 보관한 그림을 바로 쓴다 */
  async warmUp(): Promise<void> {
    await this.render(1, quickPixelWidth(this.sizes[0]), true);
  }

  async draw(page: number, canvas: HTMLCanvasElement, opts: DrawOpts, signal: AbortSignal): Promise<void> {
    const size = this.sizes[page - 1];
    if (!size) throw new Error(`${page}쪽이 없습니다.`);
    const px = targetPixelWidth(opts.cssWidth, opts.dpr, size);
    if (drawPlan(this.pixelWidthsOf(page), px, size.w) !== 'two-stage') {
      paint(canvas, await this.render(page, px, true, signal));
      return;
    }
    paint(canvas, await this.render(page, quickPixelWidth(size), true, signal));
    // 선명한 그림은 급하지 않은 줄로 — 다 되면 같은 캔버스를 바꿔 칠한다(PagesBody가 버린 캔버스는 폭이 0)
    // shortcut: a failed sharpen leaves the 1x picture; the next zoom or scroll redraws it.
    this.render(page, px, false, signal).then(
      (img) => {
        if (canvas.width > 0) paint(canvas, img);
      },
      () => undefined,
    );
  }

  async drawText(page: number, layer: HTMLElement, cssWidth: number): Promise<void> {
    const text = await this.textOf(page);
    layer.replaceChildren();
    for (const run of text.runs) {
      const g = runSpanGeom(run, text, cssWidth);
      const span = document.createElement('span');
      span.textContent = run.t;
      span.style.left = `${g.left}%`;
      span.style.top = `${g.top}%`;
      span.style.fontSize = `${g.fontPx}px`;
      span.style.fontFamily = 'sans-serif';
      const natural = naturalWidth(run.t, g.fontPx);
      if (natural > 0) span.style.transform = `scaleX(${g.widthPx / natural})`;
      layer.append(span);
    }
  }

  async search(q: string, signal: AbortSignal): Promise<SearchHit[]> {
    const hits: SearchHit[] = [];
    for (let page = 1; page <= this.sizes.length && hits.length < SEARCH_LIMIT; page += 1) {
      if (signal.aborted) throw new DOMException('찾기를 멈췄습니다.', 'AbortError');
      hits.push(...runHits(page, await this.textOf(page), q));
    }
    return hits.slice(0, SEARCH_LIMIT);
  }

  async thumbnail(page: number, width: number): Promise<HTMLCanvasElement> {
    const img = await this.render(page, Math.ceil(width * Math.min(2, window.devicePixelRatio || 1)), false);
    const canvas = document.createElement('canvas');
    paint(canvas, img);
    return canvas;
  }

  destroy(): void {
    this.client.terminate();
  }

  private pixelWidthsOf(page: number): number[] {
    const prefix = `${page}:`;
    return this.images
      .keys()
      .filter((key) => key.startsWith(prefix))
      .map((key) => Number(key.slice(prefix.length)));
  }

  private async render(page: number, px: number, urgent: boolean, signal?: AbortSignal): Promise<ImageData> {
    const key = `${page}:${px}`;
    const hit = this.images.get(key);
    if (hit) return hit;
    const size = this.sizes[page - 1];
    const reply = await this.client.call<RenderReply>({ cmd: 'render', page, scale: px / size.w }, RENDER_TIMEOUT_MS, urgent, signal);
    const img = new ImageData(new Uint8ClampedArray(reply.rgba), reply.w, reply.h);
    this.images.set(key, img);
    return img;
  }

  private textOf(page: number): Promise<PageText> {
    let hit = this.texts.get(page);
    if (!hit) {
      hit = this.client.call<TextReply>({ cmd: 'text', page }, RENDER_TIMEOUT_MS, false).then((reply) => reply.text);
      hit.catch(() => this.texts.delete(page)); // 실패한 것은 다음에 다시 묻는다
      this.texts.set(page, hit);
    }
    return hit;
  }
}

/** Worker를 띄워 엔진·글꼴·원본을 받고 열어 1쪽 1배까지 그린다. 어떤 실패든 ArtcraftUnavailable(설계 §4.2 전환 규칙) */
export async function openArtcraftSource(info: ViewInfo): Promise<ArtcraftSource> {
  const a = info.artcraft;
  if (!a) throw new ArtcraftUnavailable('브라우저 엔진 정보가 없습니다');
  let client: WorkerClient | null = null;
  try {
    client = new WorkerClient(new Worker(WORKER_URL, { type: 'module' }));
    const opened = await client.call<InitReply>(
      { cmd: 'init', wasm: a.assets.wasm, fonts: a.assets.fonts, original: a.originalUrl, name: `document.${a.ext}` },
      OPEN_TIMEOUT_MS,
    );
    if (opened.count === 0) throw new ArtcraftUnavailable('쪽이 없습니다');
    if (pageCountSuspicious(a.experimental, opened.count, info.pageCount)) {
      throw new ArtcraftUnavailable(`쪽 수가 서버 변환(${info.pageCount})과 많이 다릅니다(${opened.count})`);
    }
    const source = new ArtcraftSource(client, opened.sizes, info.permissions.canCopy);
    await source.warmUp();
    return source;
  } catch (error) {
    client?.terminate();
    throw error instanceof ArtcraftUnavailable ? error : new ArtcraftUnavailable(error instanceof Error ? error.message : String(error));
  }
}
```

- [ ] **Step 8: Worker 진입점을 빌드에 넣기**

`scripts/build-editor.mjs`의 `entryPoints` 배열 끝(`'src/client/view.ts',` 다음)에 더한다:

```js
    // 아트크래프트 브라우저 엔진 Worker(module worker) — view.ts의 artcraftSource.ts가 /js/artcraftWorker.bundle.js로 띄운다
    'src/client/view/artcraftWorker.ts',
```

`.gitignore`의 `public/js/view.bundle.js.map` 다음에 더한다:

```gitignore
public/js/artcraftWorker.bundle.js
public/js/artcraftWorker.bundle.js.map
```

- [ ] **Step 9: 빌드·타입 검사·시험**

Run: `npm run typecheck && npm run build:editor && ls -la public/js/artcraftWorker.bundle.js && npx vitest run tests/artcraftLogic.test.ts tests/viewShellLogic.test.ts`
Expected: typecheck 오류 0, esbuild가 `public/js/artcraftWorker.bundle.js`를 만든다, 두 시험 파일 PASS. `git status --short public/js`는 빈 출력.

- [ ] **Step 10: 커밋**

```bash
git add .gitignore scripts/build-editor.mjs src/client/view/artcraftLogic.ts src/client/view/artcraftWorker.ts src/client/view/artcraftSource.ts src/client/view/types.ts src/client/view/pagesBody.ts tests/artcraftLogic.test.ts
git commit -m "아트크래프트 3: Worker(wasm·글꼴·원본·open·render·text)와 ArtcraftSource(두 단계 그리기·글자 층·찾기·섬네일·시간 제한)"
```

---

### Task 4: 연결 — `view.ts` 원천 고르기·전환·표지, 메인 화면 단추·공유 라디오

**Files:**
- Modify: `src/client/view.ts:105-185` (openSource 아래 ~ boot 끝)
- Modify: `public/view.html:49` (doc-title 옆 표지), `public/css/view.css` (표지 모양)
- Modify: `public/js/main.js` (issueToken·openViewer·renderDocItem), `public/css/style.css` (라디오·이유 줄)

**Interfaces:**
- Consumes: `openArtcraftSource(info: ViewInfo): Promise<ArtcraftSource>`·`ArtcraftUnavailable`(Task 3, `src/client/view/artcraftSource.ts`), `ViewInfo.engine/engineReason/artcraft`(Task 3 types), `PagesBody.kind === 'pages-artcraft'`, 토큰 응답 `engine`/`engineReason`(Task 2).
- Produces: DOM — `#engine-badge`(아트크래프트로 그릴 때만 보임, 글 `브라우저 엔진(아트크래프트)`(+` · 실험`)), `#view-notice`에 전환·이유 문구, 메인 화면 각 줄의 `아트크래프트로 열기` 단추(안 되는 형식은 `disabled` + 감싼 span의 `title`), 공유 폼 `input[name=engine]` 라디오 두 개(값 `libreoffice`·`artcraft`). Task 5 e2e가 이 선택자를 쓴다.

- [ ] **Step 1: 표지 자리 — `public/view.html`·`public/css/view.css`**

`public/view.html`의 `<span id="doc-title" class="doc-title">문서 불러오는 중…</span>` 바로 다음 줄에 더한다:

```html
        <span id="engine-badge" class="engine-badge" hidden title="브라우저 안에서 그린 화면입니다 — 쪽 수가 Word·서버 변환과 다를 수 있습니다. 인쇄·내려받기는 서버 PDF를 씁니다.">브라우저 엔진(아트크래프트)</span>
```

`public/css/view.css` 끝에 더한다:

```css
/* 아트크래프트(브라우저 엔진)로 그릴 때만 — 쪽 수가 Word와 다를 수 있음을 알린다(설계 §4.2) */
.engine-badge { margin-left: 8px; padding: 2px 8px; border-radius: 999px; background: #e6f4f1; color: #1f4d47; font-size: 12px; line-height: 18px; white-space: nowrap; }
.engine-badge[hidden] { display: none; }
```

- [ ] **Step 2: `src/client/view.ts` — 원천 고르기와 전환**

`const LOCKED_NOTICE = …;` 줄 다음에 더한다:

```ts
/** 아트크래프트(브라우저 엔진)가 열지 못했거나 그리다 멈췄다(설계 §4.2) */
const ARTCRAFT_FALLBACK_NOTICE = '브라우저 엔진이 이 문서를 열지 못해 서버 변환으로 보여 줍니다';

function showEngineBadge(info: ViewInfo, on: boolean): void {
  const badge = document.getElementById('engine-badge') as HTMLElement;
  badge.hidden = !on;
  if (on) badge.textContent = info.artcraft?.experimental ? '브라우저 엔진(아트크래프트) · 실험' : '브라우저 엔진(아트크래프트)';
}

/** 처음 원천 — 아트크래프트 토큰이면 브라우저 엔진(동적 import), 실패하면 안내하고 서버 변환. 서버가 거절 이유를 줬으면 그 안내 */
async function firstSource(info: ViewInfo, libreSource: () => Promise<PageSource>): Promise<PageSource> {
  if (info.engine === 'artcraft' && info.artcraft) {
    try {
      const { openArtcraftSource } = await import('./view/artcraftSource.js');
      const source = await openArtcraftSource(info);
      showEngineBadge(info, true);
      return source;
    } catch (error) {
      console.warn('브라우저 엔진으로 열지 못해 서버 변환으로 보여 줍니다:', error);
      notice(ARTCRAFT_FALLBACK_NOTICE);
    }
  } else if (info.engineReason) {
    notice(info.engineReason);
  }
  return libreSource();
}
```

`boot()`에서 `setupEditButton({ info });` 다음부터 `const body = new PagesBody(await openSource(info, png), info.pages);`까지(`if (info.processing) {…}` 블록 포함)를 아래로 바꾼다:

```ts
  const png = () => new PngSource(tokenQuery, info.pageCount, info.pages);
  /** 서버 변환 원천 — 첫 변환 중(editReady)이면 "쪽을 준비하는 중"을 기다린다. 아트크래프트는 이것을 부르지 않으므로 변환을 기다리지 않는다 */
  const libreSource = async (): Promise<PageSource> => {
    if (info.processing) {
      postReady();
      await waitUntilPrepared(tokenQuery);
      info = await fetchInfo();
    }
    return openSource(info, png);
  };
  let first: PageSource;
  try {
    first = await firstSource(info, libreSource);
  } catch (error) {
    showError(error instanceof Error ? error.message : '문서를 열 수 없습니다.', info.editorUrl);
    return;
  }
  const body = new PagesBody(first, info.pages);
```

`body.onDrawError(() => { … });` 블록 전체를 아래로 바꾼다:

```ts
  // pdf.js로 그리다 실패하면(깨진 파일·휴대폰 메모리) 같은 자리에 서버 PNG로(설계 §4·§7).
  // 아트크래프트가 그리다 실패·멈추면(Worker 오류·20초 초과) 서버 변환 원천으로(설계 2026-11-21 §4.2) — 여러 쪽이 함께 실패해도 한 번만
  let leavingArtcraft = false;
  body.onDrawError(() => {
    if (body.kind === 'pages-pdfjs') {
      notice(FALLBACK_NOTICE);
      body.replaceSource(png());
      return;
    }
    if (body.kind !== 'pages-artcraft' || leavingArtcraft) return;
    leavingArtcraft = true;
    notice(ARTCRAFT_FALLBACK_NOTICE);
    showEngineBadge(info, false);
    libreSource().then(
      (next) => body.replaceSource(next),
      (error) => showError(error instanceof Error ? error.message : '문서를 열 수 없습니다.', info.editorUrl),
    );
  });
```

(나머지 — `body.mount`·`pagehide`·`setupToolbar` 등 — 는 그대로.)

- [ ] **Step 3: 메인 화면 — `public/js/main.js`**

`issueToken`을 바꾼다:

```js
async function issueToken(docId, viewer, canPrint, canDownload, canCopy, canEdit, pdfEditor = false, engine = 'libreoffice') {
  return fetchJson(`/api/documents/${docId}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    // pdfEditor: PDF 를 PDF 편집기(editor-pdf.html)로 열기 — 내려받기 권한이 있을 때만 editorUrl 이 온다(서버 editorUrlFor)
    // engine: 'artcraft' = 브라우저 엔진으로 열기 — 안 되면 서버가 거절하지 않고 libreoffice + engineReason(보기 화면이 안내)
    body: JSON.stringify({
      viewer,
      canPrint,
      canDownload,
      canCopy,
      canEdit: canEdit === true,
      ...(pdfEditor ? { pdfEditor: true } : {}),
      ...(engine === 'artcraft' ? { engine } : {}),
    }),
  });
}
```

`openViewer`를 바꾼다(주석 블록은 그대로 두고 함수만):

```js
async function openViewer(docId, viewer, ext, engine = 'libreoffice') {
  const canEdit = !supportsEdit({ ext });
  // PDF 는 향기랩처럼 PDF 편집기로 연다 — 보기·찾기·주석·쪽 정리·서명이 한 화면에 있다(올린 사람 본인 화면이라 전체 권한)
  const pdf = ext === 'pdf' && engine !== 'artcraft';
  const issued = await issueToken(docId, viewer, true, true, true, canEdit, pdf, engine);
  location.href = (pdf && issued.editorUrl) || issued.viewerUrl;
}
```

`supportsMarkup` 함수 다음에 더한다:

```js
/** 브라우저 엔진(아트크래프트)이 그리는 형식 — 서버 src/artcraftEngine.ts ARTCRAFT_FORMATS와 같다(향기랩 옛 형식은 올린 형식 기준) */
const ARTCRAFT_EXTS = new Set(['docx', 'pptx', 'doc']);
const ARTCRAFT_UNSUPPORTED = '이 형식은 브라우저 엔진이 지원하지 않습니다';
function supportsArtcraft(doc) {
  return ARTCRAFT_EXTS.has(doc.legacy?.ext ?? doc.ext);
}
```

`renderDocItem`에서 `info.append(quickView);` 줄을 아래로 바꾼다:

```js
  const artOk = supportsArtcraft(doc);
  if (artOk) quickView.textContent = '리브레오피스로 열기';
  info.append(quickView);

  // 아트크래프트(브라우저 엔진)로 열기 — 안 되는 형식은 흐리게 + 이유 말풍선(흐린 단추는 마우스를 받지 않아 감싼 span에 title)
  const artWrap = document.createElement('span');
  const artView = document.createElement('button');
  artView.type = 'button';
  artView.className = 'btn';
  artView.textContent = '아트크래프트로 열기';
  if (artOk) {
    artWrap.title = '브라우저 안에서 그립니다 — 서버 변환 없이 열리며, 쪽 수가 Word와 다를 수 있습니다';
    artView.addEventListener('click', () => {
      artView.disabled = true;
      openViewer(doc.id, '업로더', doc.ext, 'artcraft').catch((error) => {
        artView.disabled = false;
        setStatus(error.message, 'error');
      });
    });
  } else {
    artWrap.title = ARTCRAFT_UNSUPPORTED;
    artView.disabled = true;
    artView.style.pointerEvents = 'none';
  }
  artWrap.append(artView);
  info.append(artWrap);
```

공유 폼 `form.innerHTML` 템플릿에서 `${doc.ext === 'pdf' ? … : ''}` 줄 다음, `<button type="submit"` 앞에 더한다:

```js
    <span class="engine-choice" role="radiogroup" aria-label="여는 엔진">
      <label><input type="radio" name="engine" value="libreoffice" checked /> 리브레오피스</label>
      <label title="${artOk ? '브라우저 안에서 그립니다 — 저장 권한이 있어야 합니다' : ARTCRAFT_UNSUPPORTED}"><input type="radio" name="engine" value="artcraft" ${artOk ? '' : 'disabled'} /> 아트크래프트</label>
    </span>
```

`form.addEventListener('submit', …)` 안의 `issueToken(` 호출 인자 끝(`data.get('pdfEditor') === 'on',` 다음)에 더한다:

```js
        data.get('engine') === 'artcraft' ? 'artcraft' : 'libreoffice',
```

같은 핸들러에서 `open.textContent = '뷰어 열기';` 줄을 아래로 바꾸고, `result.append(open);` 다음에 이유 줄을 더한다:

```js
      open.textContent = issued.engine === 'artcraft' ? '아트크래프트로 열기' : artOk ? '리브레오피스로 열기' : '뷰어 열기';
```

```js
      if (issued.engineReason) {
        const reason = document.createElement('span');
        reason.className = 'engine-reason';
        reason.textContent = issued.engineReason;
        result.append(reason);
      }
```

`public/css/style.css` 끝에 더한다:

```css
/* 메인 화면 공유 폼 — 여는 엔진 고르기(리브레오피스/아트크래프트) */
.engine-choice { display: inline-flex; gap: 10px; align-items: center; }
.engine-choice label[title] { cursor: help; }
.engine-reason { margin-left: 8px; color: #8a5a00; font-size: 0.85rem; }
```

- [ ] **Step 4: 타입 검사·빌드·단위 시험**

Run: `npm run typecheck && npm run build:editor && node --check public/js/main.js && npx vitest run tests/viewShellLogic.test.ts tests/artcraftLogic.test.ts`
Expected: 오류 0, 빌드 성공, `node --check` 출력 없음, PASS. (화면 동작은 Task 5 e2e가 실제 브라우저로 본다.)

- [ ] **Step 5: 커밋**

```bash
git add src/client/view.ts public/view.html public/css/view.css public/js/main.js public/css/style.css
git commit -m "아트크래프트 4: 보기 화면 원천 고르기·리브레오피스 자동 전환·표지, 메인 화면 '아트크래프트로 열기'·공유 폼 엔진 라디오"
```

---

### Task 5: 브라우저 검사 `verify-artcraft-e2e.mjs` · 쪽 수 기록 · 회귀 묶음

**Files:**
- Create: `scripts/verify-artcraft-e2e.mjs`
- Modify: `package.json` (scripts `verify:artcraft`)
- 산출(커밋 안 함): `docs/quality-audit/out/artcraft-pages.md`, `.audit-run/artcraft-e2e-<포트>/`

**Interfaces:**
- Consumes: Task 1~4 전부 — `/artcraft/*`, 토큰 `engine`/`engineReason`, `#engine-badge`, `#view-notice`, `.page-frame[data-page="N"].page-ready`, `#page-jump-total`("의 N"), `#zoom-in`, `#zoom-level`, `#rotate-btn`, `#pages.pages-rotated`, `#page-next`, `#page-jump-input`, `#thumbs-panel canvas`, `#search-btn`, `#search-input`, `#search-count`, `.search-hit-overlay`, `.pdf-text-layer span`, 메인 화면 `li.doc-item` 안 `아트크래프트로 열기` 단추·`input[name=engine]`.
- Produces: 통과/실패 줄과 종료 코드(실패 있으면 1), `docs/quality-audit/out/artcraft-pages.md`.

- [ ] **Step 1: e2e 스크립트 쓰기**

`scripts/verify-artcraft-e2e.mjs`:

```js
// 아트크래프트(브라우저 엔진) 보기 — 설계 2026-11-21 §7의 브라우저 검사 일곱 가지 + 메인 화면 단추, 쪽 수 비교 기록.
//  (1) 아트크래프트로 열려 쪽 수·첫 쪽 그림(잉크 비율·LibreOffice 첫 쪽과의 유사도·한글 글자 층)
//  (2) 확대·회전·쪽 이동·섬네일·찾기  (3) 서버 변환 요청 0건  (4) wasm을 막으면 리브레오피스로 자동 전환 + 안내
//  (5) 내려받기 권한 없는 토큰은 서버 변환으로  (6) CPU×4 첫 쪽 시간(로컬 서버라 받기 시간은 사실상 0 — 목표 2초, 고객 pptx 3초)
//  (7) 두 번째 열기에서 wasm·글꼴을 서버에서 다시 받지 않음(서버 요청 기록으로 센다)
// 이 스크립트가 시험 서버를 직접 띄운다(npm run build:editor, npm run build:artcraft 뒤, LibreOffice 필요):
//   ARTCRAFT_E2E_ORIGINALS=<고객형 원본 폴더> node scripts/verify-artcraft-e2e.mjs [포트=8177]
// 고객형 문서는 그 폴더에서 .audit-run/으로 복사해 쓴다(저장소에 넣지 않는다). 폴더를 안 주면 그 둘은 건너뛴다(SKIP 줄).
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import sharp from 'sharp';
import { chromium } from 'playwright';

const PORT = Number(process.argv[2] ?? 8177);
const FORBIDDEN_PORTS = [8000, 8010, 8020, 8100, 8110, 8130, 8148];
if (FORBIDDEN_PORTS.includes(PORT)) throw new Error(`운영 포트 ${PORT}는 쓰지 않는다`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `artcraft-e2e-${PORT}`);
const ORIGINALS = process.env.ARTCRAFT_E2E_ORIGINALS ?? '';
const CUSTOMER = [
  { src: '62c3881f-247e-48ad-aacb-977010a23ef1.docx', name: '고객형.docx', limitMs: 2000 },
  { src: '8c76f45a-e15a-40ca-b6be-d9ca5ed63137.pptx', name: '고객형.pptx', limitMs: 3000 },
];
const REPORT = path.resolve('docs/quality-audit/out/artcraft-pages.md');
const SERVER_CONVERT = /^\/api\/view\/(pages|derived-pdf|status|search|pdf-)/;
const results = [];
const rows = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

for (const [file, how] of [['public/artcraft/manifest.json', 'npm run build:artcraft'], ['public/js/artcraftWorker.bundle.js', 'npm run build:editor']]) {
  if (!fs.existsSync(file)) {
    console.error(`${file} 없음 — ${how} 먼저`);
    process.exit(1);
  }
}

let server = null;
async function startServer() {
  fs.rmSync(RUN, { recursive: true, force: true });
  fs.mkdirSync(path.join(RUN, 'tmp'), { recursive: true });
  const log = fs.openSync(path.join(RUN, 'server.log'), 'a');
  server = spawn(process.execPath, ['node_modules/tsx/dist/cli.mjs', 'src/server.ts'], {
    env: { ...process.env, DOCVIEW_PORT: String(PORT), DOCVIEW_HOST: '127.0.0.1', DOCVIEW_SECRET: randomBytes(24).toString('hex'), DOCVIEW_DATA_DIR: path.join(RUN, 'data'), TMP: path.join(RUN, 'tmp'), TEMP: path.join(RUN, 'tmp'), DOCVIEW_LO_PROFILE_ROOT: path.join(RUN, 'lo') },
    stdio: ['ignore', log, log],
  });
  for (let i = 0; i < 120; i += 1) {
    if (await fetch(`${BASE}/api/formats`).then((r) => r.ok, () => false)) return;
    await sleep(500);
  }
  throw new Error('시험 서버가 뜨지 않음');
}
function stopServer() {
  // 내가 띄운 서버만(PID, 자식 포함) — 이름으로 node를 끄지 않는다
  if (server?.pid) execFileSync('taskkill', ['/PID', String(server.pid), '/T', '/F'], { stdio: 'ignore' });
  server = null;
}
/** 서버 기록에서 /artcraft/ 요청 줄 수(브라우저 캐시에서 오면 서버에 요청이 오지 않는다) */
async function artcraftHits() {
  await sleep(500);
  return fs.readFileSync(path.join(RUN, 'server.log'), 'utf8').split('\n').filter((line) => /"url":"\/artcraft\//.test(line)).length;
}

async function upload(file, name) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), name);
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패 ${name}: ${JSON.stringify(r)}`);
  for (let i = 0; i < 1200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready') return st;
    if (st.status === 'failed') throw new Error(`변환 실패 ${name}: ${st.error}`);
    await sleep(250);
  }
  throw new Error(`대기 시간 초과 ${name}`);
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: 'artcraft-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);

async function gray(png) {
  const { data, info } = await sharp(png).flatten({ background: '#ffffff' }).resize({ width: 128 }).grayscale().raw().toBuffer({ resolveWithObject: true });
  return { data, w: info.width, h: info.height };
}
/** 1 − 평균 밝기 차/255(겹치는 높이만) */
function similarity(a, b) {
  const h = Math.min(a.h, b.h);
  let sum = 0;
  for (let i = 0; i < h * a.w; i += 1) sum += Math.abs(a.data[i] - b.data[i]);
  return 1 - sum / (h * a.w * 255);
}
const inkRatio = (g) => g.data.filter((v) => v < 200).length / g.data.length;
const pct = (v) => `${(v * 100).toFixed(1)}%`;

async function openPage(context, url, { throttle = false, block = false } = {}) {
  const page = await context.newPage();
  const errors = [];
  const paths = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  page.on('request', (r) => paths.push(new URL(r.url()).pathname));
  if (block) await page.route('**/artcraft/**', (r) => r.abort());
  let cdp = null;
  if (throttle) {
    cdp = await context.newCDPSession(page);
    await cdp.send('Emulation.setCPUThrottlingRate', { rate: 4 });
  }
  const t0 = Date.now();
  await page.goto(`${BASE}${url}`);
  await page.waitForSelector('.page-frame[data-page="1"].page-ready', { timeout: 90000 });
  const firstMs = Date.now() - t0;
  if (cdp) await cdp.send('Emulation.setCPUThrottlingRate', { rate: 1 });
  return { page, errors, paths, firstMs };
}

const cases = [
  { file: 'audit/office/corpus/docx/a1-korean-report.docx', name: 'a1-korean-report.docx', limitMs: 2000, hangul: true },
  { file: 'audit/office/corpus/pptx/a2-korean-deck.pptx', name: 'a2-korean-deck.pptx', limitMs: 2000, hangul: true },
  { file: 'tests/fixtures/legacy/SimpleMacro.doc', name: 'SimpleMacro.doc', limitMs: 2000, hangul: false },
];

await startServer();
const browser = await chromium.launch();
try {
  if (ORIGINALS) {
    fs.mkdirSync(path.join(RUN, 'cust'), { recursive: true });
    for (const c of CUSTOMER) {
      const to = path.join(RUN, 'cust', c.src);
      fs.copyFileSync(path.join(ORIGINALS, c.src), to);
      cases.push({ file: to, name: c.name, limitMs: c.limitMs, hangul: true });
    }
  } else {
    console.log('SKIP 고객형 docx·pptx — ARTCRAFT_E2E_ORIGINALS를 주지 않음');
  }

  const ids = {};
  for (const c of cases) {
    const meta = await upload(c.file, c.name);
    ids[c.name] = meta.id;
    const tok = await token(meta.id, { engine: 'artcraft' });
    rec(`${c.name}: 토큰 engine=artcraft`, tok.engine === 'artcraft', tok.engineReason ?? '');
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const before = await artcraftHits();
    const o = await openPage(context, tok.viewerUrl, { throttle: true });

    // (6) CPU×4 첫 쪽
    rec(`${c.name}: (6) CPU×4 첫 쪽 ${o.firstMs} ms ≤ ${c.limitMs}`, o.firstMs <= c.limitMs);

    // (1) 아트크래프트 표지·쪽 수·첫 쪽 그림
    const badge = await o.page.locator('#engine-badge').isVisible();
    const total = Number((await o.page.textContent('#page-jump-total')).replace(/\D/g, ''));
    rec(`${c.name}: (1) 아트크래프트로 열림(표지) · 쪽 수 ${total} (LibreOffice ${meta.pageCount})`, badge && total > 0);
    await sleep(3000); // 두 단계 그리기의 선명한 그림까지
    const dataUrl = await o.page.$eval('.page-frame[data-page="1"] canvas', (cv) => cv.toDataURL('image/png'));
    const mine = await gray(Buffer.from(dataUrl.split(',')[1], 'base64'));
    const loTok = await token(meta.id, {});
    const loPng = Buffer.from(await (await fetch(`${BASE}/api/view/pages/1?token=${encodeURIComponent(loTok.token)}&scale=2`)).arrayBuffer());
    const sim = similarity(mine, await gray(loPng));
    const ink = inkRatio(mine);
    const hangul = await o.page.$$eval('.pdf-text-layer span', (ss) => ss.some((s) => /[가-힣]/.test(s.textContent ?? '')));
    rec(
      `${c.name}: (1) 첫 쪽 그림 — 잉크 ${pct(ink)} · LibreOffice와 유사도 ${sim.toFixed(3)} · 한글 글자 층 ${hangul}`,
      ink > 0.003 && sim >= 0.8 && (!c.hangul || hangul),
    );

    // (2) 확대·회전·쪽 이동·섬네일·찾기
    const w0 = await o.page.$eval('.page-frame[data-page="1"] canvas', (cv) => cv.width);
    await o.page.click('#zoom-in');
    const zoomed = await o.page
      .waitForFunction((w) => (document.querySelector('.page-frame[data-page="1"] canvas')?.width ?? 0) > w * 1.15, w0, { timeout: 30000 })
      .then(() => true, () => false);
    rec(`${c.name}: (2) 확대 ${await o.page.textContent('#zoom-level')} — 쪽 그림을 새 배율로`, zoomed);
    await o.page.click('#rotate-btn');
    rec(`${c.name}: (2) 회전`, await o.page.$eval('#pages', (el) => el.classList.contains('pages-rotated')));
    for (let i = 0; i < 3; i += 1) await o.page.click('#rotate-btn');
    if (total > 1) {
      await o.page.click('#page-next');
      const moved = await o.page
        .waitForFunction(() => document.querySelector('#page-jump-input')?.value === '2' && document.querySelector('.page-frame[data-page="2"]')?.classList.contains('page-ready'), null, { timeout: 30000 })
        .then(() => true, () => false);
      rec(`${c.name}: (2) 쪽 이동 → 2쪽`, moved);
    }
    const thumbs = await o.page
      .waitForFunction(() => [...document.querySelectorAll('#thumbs-panel canvas')].some((cv) => cv.width > 0), null, { timeout: 30000 })
      .then(() => true, () => false);
    rec(`${c.name}: (2) 섬네일`, thumbs);
    const word = await o.page.$$eval('.pdf-text-layer span', (ss) => {
      for (const s of ss) {
        const m = (s.textContent ?? '').match(/[가-힣A-Za-z]{2,}/);
        if (m) return m[0];
      }
      return null;
    });
    if (word) {
      await o.page.click('#search-btn');
      await o.page.fill('#search-input', word);
      await o.page.press('#search-input', 'Enter');
      const found = await o.page
        .waitForFunction(() => /\d+ \/ /.test(document.querySelector('#search-count')?.textContent ?? '') && !!document.querySelector('.search-hit-overlay'), null, { timeout: 30000 })
        .then(() => true, () => false);
      rec(`${c.name}: (2) 찾기 "${word}" — ${await o.page.textContent('#search-count')}`, found);
    } else {
      rec(`${c.name}: (2) 찾기 — 글자 층이 비어 건너뜀`, true);
    }

    // (3) 서버 변환 요청 0건
    const converts = o.paths.filter((p) => SERVER_CONVERT.test(p));
    rec(`${c.name}: (3) 서버 변환 요청 0건`, converts.length === 0, converts.join(', '));
    rec(`${c.name}: 페이지 오류 없음`, o.errors.length === 0, o.errors.join(' | ').slice(0, 300));

    // (7) 두 번째 열기 — 같은 브라우저(캐시)
    await o.page.close();
    const mid = await artcraftHits();
    const again = await openPage(context, tok.viewerUrl);
    const after = await artcraftHits();
    rec(`${c.name}: (7) 두 번째 열기 — wasm·글꼴을 캐시에서(처음 ${mid - before}건, 두 번째 ${after - mid}건)`, mid > before && after === mid);
    await again.page.close();
    await context.close();
    rows.push({ name: c.name, lo: meta.pageCount, art: total, sim, firstMs: o.firstMs });
  }

  // (4) wasm을 막으면 리브레오피스로 자동 전환 + 안내
  {
    const tok = await token(ids['a1-korean-report.docx'], { engine: 'artcraft' });
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const o = await openPage(context, tok.viewerUrl, { block: true });
    const text = await o.page.textContent('#view-notice');
    const badge = await o.page.locator('#engine-badge').isVisible();
    rec('(4) wasm을 막으면 리브레오피스로 전환 + 안내', /브라우저 엔진이 이 문서를 열지 못해/.test(text ?? '') && !badge, text ?? '');
    await context.close();
  }

  // (5) 내려받기 권한 없는 토큰은 서버 변환
  {
    const tok = await token(ids['a1-korean-report.docx'], { engine: 'artcraft', canDownload: false });
    rec('(5) 내려받기 권한 없음 → 토큰 engine=libreoffice + 이유', tok.engine === 'libreoffice' && tok.engineReason === '내려받기 권한이 없어 서버 변환으로 엽니다', tok.engineReason ?? '');
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const o = await openPage(context, tok.viewerUrl);
    const asked = o.paths.filter((p) => p.startsWith('/artcraft/') || p.includes('artcraftWorker'));
    const text = await o.page.textContent('#view-notice');
    rec('(5) 서버 변환으로 열림 — 엔진 자산 요청 없음 + 이유 안내', asked.length === 0 && /내려받기 권한이 없어/.test(text ?? ''), asked.join(', '));
    await context.close();
  }

  // 메인 화면 — 단추·공유 폼 라디오
  {
    await upload('tests/fixtures/legacy/SimpleMacro.ppt', 'SimpleMacro.ppt');
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } });
    const page = await context.newPage();
    await page.goto(`${BASE}/`);
    const docRow = page.locator('li.doc-item', { hasText: 'a1-korean-report.docx' });
    const pptRow = page.locator('li.doc-item', { hasText: 'SimpleMacro.ppt' });
    await docRow.waitFor({ timeout: 30000 });
    const pptButton = pptRow.getByRole('button', { name: '아트크래프트로 열기' });
    const pptTitle = await pptButton.locator('xpath=..').getAttribute('title');
    rec('메인: .ppt 줄의 "아트크래프트로 열기"는 흐림 + 이유', (await pptButton.isDisabled()) && pptTitle === '이 형식은 브라우저 엔진이 지원하지 않습니다', pptTitle ?? '');
    rec('메인: docx 줄에 "리브레오피스로 열기"', await docRow.getByRole('button', { name: '리브레오피스로 열기' }).isVisible());
    const form = docRow.locator('form.token-form');
    await form.locator('input[name=viewer]').fill('공유받는 사람');
    await form.locator('input[name=canDownload]').check();
    await form.locator('input[name=engine][value=artcraft]').check();
    await form.getByRole('button', { name: '열람 링크 발급' }).click();
    const link = docRow.locator('.token-result a', { hasText: '아트크래프트로 열기' });
    await link.waitFor({ timeout: 15000 });
    const shared = await openPage(context, await link.getAttribute('href'));
    rec('메인: 공유 폼 아트크래프트 링크 → 표지 보임', await shared.page.locator('#engine-badge').isVisible());
    await shared.page.close();
    await docRow.getByRole('button', { name: '아트크래프트로 열기' }).click();
    await page.waitForURL(/\/view\.html\?token=/, { timeout: 15000 });
    await page.waitForSelector('.page-frame[data-page="1"].page-ready', { timeout: 90000 });
    rec('메인: "아트크래프트로 열기" 단추 → 표지 보임', await page.locator('#engine-badge').isVisible());
    await context.close();
  }

  fs.mkdirSync(path.dirname(REPORT), { recursive: true });
  fs.writeFileSync(
    REPORT,
    [
      '# 아트크래프트 쪽 수 비교',
      '',
      `만든 때 ${new Date().toISOString()} · \`scripts/verify-artcraft-e2e.mjs\` · 고객형 문서는 이름 대신 "고객형"으로 적음`,
      '',
      '| 문서 | LibreOffice 쪽 | 아트크래프트 쪽 | 차이 | 첫 쪽 유사도 | 첫 쪽 ms(CPU×4) |',
      '|---|---|---|---|---|---|',
      ...rows.map((r) => `| ${r.name} | ${r.lo} | ${r.art} | ${r.art - r.lo} | ${r.sim.toFixed(3)} | ${r.firstMs} |`),
      '',
    ].join('\n'),
  );
  console.log(`쪽 수 기록 → ${REPORT}`);
} finally {
  await browser.close();
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed ? 1 : 0);
```

`package.json` `scripts`의 `"verify:docx"` 다음에 더한다:

```json
    "verify:artcraft": "node scripts/verify-artcraft-e2e.mjs",
```

- [ ] **Step 2: 문법 확인**

Run: `node --check scripts/verify-artcraft-e2e.mjs`
Expected: 출력 없음.

- [ ] **Step 3: 전체 단위 시험·빌드**

Run: `npm run typecheck && npm test && npm run build:editor && npm run build:artcraft`
Expected: typecheck 0, vitest 전부 PASS, 두 빌드 성공(build:artcraft는 캐시가 있어 cargo가 빨리 끝난다).

- [ ] **Step 4: 아트크래프트 e2e 실행(포트 8177 — 비어 있는지 먼저)**

```bash
netstat -ano | grep ':8177 ' ; echo "포트 확인 끝(위에 LISTENING이 없어야 함)"
ARTCRAFT_E2E_ORIGINALS='C:/Users/OpenMinds/projects/webviewer-q-all/.audit-run/test-server/originals' npm run verify:artcraft
cat docs/quality-audit/out/artcraft-pages.md
```
Expected: 마지막 줄 `N/N 통과`, 종료 코드 0. `artcraft-pages.md`에 5개 문서 표(a1 5/5, a2 3/3, SimpleMacro.doc 1/1, 고객형.docx LibreOffice 6 · 아트크래프트 5, 고객형.pptx 2/2 — spike 값과 같아야 함). FAIL이 있으면 고치지 말고 실패 줄과 `.audit-run/artcraft-e2e-8177/server.log`를 보고 원인을 찾는다(superpowers:systematic-debugging). (4)가 실패하고 `page.route`가 Worker 요청을 잡지 못한 것이면 그 사실을 보고한다 — 화면 코드를 바꿔 맞추지 않는다.

- [ ] **Step 5: 회귀 묶음 — 리브레오피스 경로(시험 서버 8124를 직접 띄우고 PID로 끈다)**

PowerShell에서(포트가 비어 있는지 먼저 보고, 서버는 `Start-Process`가 준 PID로만 끈다):

```powershell
if (netstat -ano | Select-String ':8124 .*LISTENING') { throw '8124를 이미 누가 쓰고 있다 — 다른 빈 포트로 바꿔 아래 BASE도 같이 바꾼다' }
New-Item -ItemType Directory -Force .audit-run\regress-8124\tmp | Out-Null
$env:DOCVIEW_PORT='8124'; $env:DOCVIEW_HOST='127.0.0.1'; $env:DOCVIEW_DATA_DIR="$PWD\.audit-run\regress-8124\data"; $env:DOCVIEW_LO_PROFILE_ROOT="$PWD\.audit-run\regress-8124\lo"; $env:TMP="$PWD\.audit-run\regress-8124\tmp"; $env:TEMP=$env:TMP
$p = Start-Process -FilePath node -ArgumentList 'node_modules/tsx/dist/cli.mjs','src/server.ts' -PassThru -WindowStyle Hidden -RedirectStandardOutput .audit-run\regress-8124\server.log -RedirectStandardError .audit-run\regress-8124\server.err
$p.Id
for ($i = 0; $i -lt 60; $i++) { try { Invoke-WebRequest -UseBasicParsing http://127.0.0.1:8124/api/formats | Out-Null; break } catch { Start-Sleep -Milliseconds 500 } }
node scripts/verify-docx-edit-e2e.mjs http://127.0.0.1:8124
node scripts/verify-pptx-fidelity-e2e.mjs http://127.0.0.1:8124
node scripts/verify-view-only-e2e.mjs http://127.0.0.1:8124
node scripts/verify-view-shell-e2e.mjs http://127.0.0.1:8124
taskkill /PID $p.Id /T /F
```
Expected: 네 스크립트 모두 실패 0(각자 마지막 요약 줄이 전부 통과). 이 브랜치 전의 결과와 다른 실패가 있으면 그 스크립트를 `git worktree add` 로 만든 8103231(이 브랜치 시작점) 작업 폴더에서 같은 방법으로(다른 빈 포트) 돌려 원래부터 있던 실패인지 가린 뒤 보고한다.

- [ ] **Step 6: 커밋**

```bash
git add scripts/verify-artcraft-e2e.mjs package.json
git status --short   # docs/quality-audit/out·.audit-run·public/artcraft·public/js가 목록에 없어야 한다
git commit -m "아트크래프트 5: verify:artcraft 브라우저 검사(설계 §7 일곱 가지 + 메인 화면)와 쪽 수 비교 기록"
```

---

## 자체 점검(계획 작성자)

**설계 대응:**
- §1 같은 `/view.html`·도구 막대, 원천만 다름 → Task 3(ArtcraftSource)·Task 4(view.ts). PagesBody는 kind 합집합 한 줄만.
- §2 형식표·거절하지 않음·메인 단추 비활성+이유 → Task 2(`ARTCRAFT_FORMATS`·`engineFor`), Task 4(main.js).
- §3.1 토큰 본문 `engine`·판정 3규칙·클레임·응답 모양 → Task 2 Step 7~8(클레임은 "요청"만 담고 판정은 매번 — 아래 "설계와 다른 점" 1).
- §3.2 info `engine`·`artcraft` 블록(`originalUrl`=download, `assets` 해시 주소) → Task 2 Step 9. 블록에 `ext`를 더함(Worker가 엔진에 줄 파일 이름).
- §3.3 메인 단추·공유 라디오 → Task 4 Step 3, e2e 메인 화면 검사.
- §4.1 kind·pageCount·textSelectable·sizeOf·qualityKey·draw(두 단계, 같은 쪽·배율 보관)·drawText·search·thumbnail·destroy → Task 3 Step 3·7. `release`는 만들지 않고 64 MB LRU(Global Constraints).
- §4.2 Worker 시작 순서·바깥 함수(시간대만 진짜)·표지·보이는 쪽 우선 큐·실패 규칙·.doc 2쪽 규칙 → Task 3 Step 3·6·7, Task 4 Step 2.
- §4.3 인쇄·내려받기 등 그대로 → 손대지 않음.
- §5 껍데기·ARTCRAFT.lock·build:artcraft(GNU, wasm-bindgen 선택, brotli, 해시, manifest)·산출물 git 제외·manifest 없을 때 끔·정적 .br/immutable·크기 상한 → Task 1.
- §6 canDownload·클레임만 믿음·같은 출처·20초·wasm 최대 메모리 → Task 2, Task 1 Step 12(`--max-memory`), Task 3 Step 7. CSP는 "설계와 다른 점" 3.
- §7 단위 4종 → Task 1(manifest 없을 때·.br 선택)·Task 2(판정표·info 블록). 브라우저 7종 + 회귀 + 쪽 수 기록 → Task 5.
- §8 하지 않는 것 → Global Constraints 마지막 줄.

**설계와 다른 점·빈 곳(실행 전에 사용자 확인 권장):**
1. 클레임 `engine`은 "요청한 엔진"이다. 판정을 토큰 때 한 번만 하면 거절 이유(`engineReason`)를 보기 화면이 알 길이 없어서, info가 같은 `engineFor`로 다시 판정한다. 권한(perms)과 요청 둘 다 서명된 클레임에서 오므로 보안 뜻은 같다.
2. 설계 §6의 "메인 화면 아트크래프트로 열기도 editReady 변환을 미리 돌리지 않게"는 지금 구조로는 안 된다: 변환은 **업로드 때** 엔진과 무관하게 시작되고(`ingestUpload` → `trackInitialConversion`), 엔진은 그 뒤 토큰 때 정해진다. 이 계획은 아트크래프트 화면이 변환을 기다리거나 부르지 않게만 했다(`libreSource`만 `waitUntilPrepared`를 부름, e2e (3)이 서버 변환 요청 0건 확인). 업로드 자체의 변환을 막으려면 업로드에 엔진 옵션이 필요하다 — 범위 밖.
3. CSP: 지금 서버가 내는 CSP는 `frame-ancestors`(허용 출처를 설정했을 때)뿐이라 `script-src`·`worker-src`가 없다 → wasm·Worker를 막는 것이 없어 더할 것이 없다. `script-src`를 새로 넣으면 `view.html`의 인라인 스크립트가 막혀 오히려 깨진다. 그래서 CSP는 바꾸지 않는다.
4. 열기 시간 제한 60초(`OPEN_TIMEOUT_MS`)는 설계에 값이 없어 이 계획이 정했다(조작된 문서의 `open` 무한 루프 대비). 20초는 첫 쪽과 그 뒤 모든 Worker 요청에 같이 쓴다.
5. 원본을 기존 `/api/view/download`로 받으므로 아트크래프트로 열 때마다 감사 기록에 `downloaded`가 한 줄 남는다(설계대로 "지금 있는 경로"). 구분이 필요하면 별도 작업.
6. 마크업(쪽 번호·비율 좌표로 저장)은 서버 변환 쪽 나눔 기준이라, 쪽 수가 다른 문서(고객 docx 6↔5)에서는 아트크래프트 화면의 다른 내용 위에 놓일 수 있다. 설계는 "변경 없음"이라 그대로 둔다.
7. 향기랩 옛 형식(legacy) 문서는 내려받기가 올린 .doc·.ppt 바이트를 주므로 판정·엔진 파일 이름을 `sourceExtOf`(올린 형식)로 했다 — .ppt는 안 됨, .doc는 실험.

**자리표시 점검:** "TBD/TODO/나중에/적절히" 없음. 모든 코드 단계에 실제 코드. 남긴 `shortcut:` 주석 셋(회전 글자 건너뜀 ×2는 spike 그대로, run을 걸친 찾기·선명하게 실패)은 한계와 고칠 때를 적었다.

**이름 일치:** `readArtcraftManifest`·`artcraftDir`·`parseManifest`(T1) → T2에서 같은 이름. `engineFor`·`sourceExtOf`·`artcraftInfo`·`ENGINE_REASONS`(T2). `ArtcraftInfo` 모양 서버(T2)=클라이언트(T3 types.ts). `openArtcraftSource`·`ArtcraftUnavailable`·`OPEN_TIMEOUT_MS`·`RENDER_TIMEOUT_MS`(T3) → T4. `PageSource.kind`·`BodyKind`의 `'pages-artcraft'`(T3) → T4 `body.kind` 비교. C ABI `sp_page_w/h`(T1) → Worker(T3). Worker 주소 `/js/artcraftWorker.bundle.js` = esbuild `[name].bundle`(T3 Step 8). DOM id `engine-badge`(T4) → e2e(T5).
