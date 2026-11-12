# 향기랩 문서도구 "늘 편집 모드" 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 향기랩 문서도구가 받는 모든 확장자를 처음부터 편집 모드로 연다 — `.doc`·`.xls`·`.ppt`는 webviewer가 OOXML로 바꿔 편집기로 열고 원래 형식으로 내려준다(안 고쳤으면 같은 바이트), 이미지는 cadview 브라우저 안에서 래스터 편집으로 열고 원래 형식으로 저장한다, NWD는 받지 않는다.

**Architecture:** webviewer는 업로드 조회 문자열 `legacyEdit=1`(향기랩 업로드 함수만 붙인다)이 있을 때만 옛 형식을 LibreOffice로 바꿔 원본 자리에 두고 올린 바이트를 파생 `legacy`로 보관, 내려받기 경로 하나가 "고쳤는가"(방의 `editedSinceSave` 또는 해시)를 보고 원래 바이트·다시 바꾼 파일을 준다. cadview는 부모 메시지 `imageEdit:true`일 때 묶음에 든 빈 R2000 도면(`blank.dwg`)을 이 탭 wasm 편집 세션으로 열고 그림을 원점·축척 1로 붙여 기존 래스터 명령으로 고치게 한다(저장은 그림만). calc-site는 업로드 표시·경로표(nwd 삭제)·`imageEdit`·`editorUrl`만 쓰기를 맞춘다.

**Tech Stack:** webviewer — TypeScript, Fastify 5, vitest 3, LibreOffice(soffice), Playwright. cadview — TypeScript(esbuild 0.24, vitest 3), Rust(edit-core 시험), wasm. calc-site — 정적 HTML·Cloudflare Pages Functions, `node --test`, wrangler(로컬 확인만).

**Spec:** `docs/superpowers/specs/2026-11-12-hyanggilab-always-edit-design.md`(a48ecc4) — §0(A1~A12) §1~§9 전부. 배포·병합·cadview 판 올리기·향기랩 `/cadview/` 묶음 커밋은 범위 밖(§9).

## Global Constraints

- **저장소·작업 폴더(각 작업 머리에 다시 적는다):**
  - webviewer: 이미 있는 워크트리 `C:\Users\OpenMinds\projects\webviewer-view`, 가지 `view-shell`(a96bf34 위, 함께 병합된다). 주 작업 폴더 `C:\Users\OpenMinds\projects\webviewer`의 `src/`·`public/`은 **한 글자도 고치지 않는다** — 운영 8000·8010이 그 폴더에서 `tsx watch`로 돌아 고치는 순간 모든 세션·토큰이 끊긴다.
  - cadview: 새 워크트리 `C:\Users\OpenMinds\projects\cadview-image-edit`, 새 가지 `image-edit`(`e1-editor-core`에서, 계획 작성 때 7d9a2e8). 주 폴더 `C:\Users\OpenMinds\projects\cadview`의 커밋 안 된 세 파일(`viewer/src/base.ts`·`viewer/src/main.ts`·`viewer/test/desktop.test.ts`)은 다른 작업이다 — 건드리지 않고, 새 워크트리에는 따라오지 않는다.
  - calc-site: 새 워크트리 `C:\Users\OpenMinds\calc-site-always-edit`, 새 가지 `always-edit`(`cadview-ocr` d16fbee에서). 주 폴더 `C:\Users\OpenMinds\calc-site`(master)·`calc-site-ocr`은 건드리지 않는다. 배포는 master에서만이고 이 계획에 없다.
- **시험 포트는 임시 포트만:** webviewer 8141(메모리 저장, 향기랩 흉내)·8143(디스크 저장, 회귀 verify용), wrangler 8142. 쓰기 전에 `netstat -ano | grep -E ":(8141|8142|8143) "`로 비었는지 보고, 차 있으면 8144~8149 중 빈 것. **8000·8010·8020·8100·8110·8130 금지**(8130은 운영 게임 서버). 띄운 서버는 PID를 적어 두고 **PID로만** 끈다(`taskkill //PID <pid> //T //F`) — node를 이름으로 죽이지 않는다.
- `git add`는 파일 이름으로 하나씩, 커밋은 `git commit -m "…" -- <경로들>`(새 파일은 먼저 `git add <경로>`). **push·stash·`git add -A` 금지.** 커밋 메시지는 한국어, 끝 줄은 정확히 `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- UI·오류 문구는 한국어(향기랩 페이지는 영어 사전 `i18n-en.js`도). 사용자에게 보이는 오류에 파일 경로를 넣지 않는다.
- **다른 호출자 회귀 금지:** `legacyEdit`이 없는 업로드(8000 유지보수 시스템, `WebviewerClient.java`)는 지금과 바이트 하나 다르지 않게 — .doc·.ppt는 PDF 열람(편집기 주소 없음), .xls는 지금처럼 xlsx로 정규화·이름 `.xlsx`. cadview 데스크톱·8020·browser-local 자체 host는 그대로.
- **시험 정책(사용자 규칙):** 작업마다 단위 시험·타입 검사·빌드만. 서버 e2e(`verify:*`)·실제 브라우저 조작·묶음 싣기는 **Task 10에서 한 번**.
- **UI는 실제 브라우저로 모든 입력 컨트롤을 조작해 확인한 뒤에 끝났다고 한다**(사용자 규칙) — Task 10.
- **master/e1-editor-core 병합은 운영 반영이다 — 이 계획에 없다.** Task 10 보고 뒤 사용자 승인을 받고 따로 한다.

---

## 설계와 맞춘 이름

| 설계가 쓴 것 | 실제(이 계획이 쓰는 것) |
|---|---|
| `legacyEdit=1`(A1) | `/api/documents/stream` querystring `legacyEdit: { type: 'string', enum: ['1'] }` → `ingestUpload(…, legacyEdit)` |
| `convertOfficeFile`(§2.2) | `src/converters/libreOffice.ts` `convertOfficeFile(data: Uint8Array, fromExt: string, toExt: string): Promise<Buffer>` |
| 옛 형식 도구 | 새 `src/legacyOffice.ts`: `LEGACY_EDIT_TARGET`·`isLegacyExt`·`sha256`·`hasVbaMacros`·`legacyDownload`·`LegacyDownload` |
| `meta.legacy`(A2) | `src/types.ts` `LegacyExt = 'doc' \| 'xls' \| 'ppt'`, `DocumentMeta.legacy?: { ext; sha256; hasMacros }` |
| 파생 `legacy` | `src/storage.ts` `DerivedKind`에 `'legacy'`(접미 `legacy.bin`) |
| 방 "고쳤는가" | `src/collab/room.ts` `hasUnsavedEdits(docId): Promise<boolean>` |
| 편집기 안내(§2.5) | `/api/view/info`의 `data.legacy`, `OfficeShellOptions.legacy` |
| e2e | `scripts/verify-legacy-edit-e2e.mjs`, `npm run verify:legacy-edit` |
| 빈 도면 틀(A8·A9) | cadview `viewer/src/edit/blank.dwg`(바이너리, esbuild `--loader:.dwg=binary`), 타입 `viewer/src/edit/dwg.d.ts`, 엔진 시험 `engine/crates/edit-core/tests/blank_template.rs` |
| `imageEdit`(A8) | `acceptOpenMessage(...)` 반환에 `imageEdit: boolean` |
| 저장은 그림만(A10) | `LocalBackendOptions.imageOnly?: boolean` |
| WEBP 다시 쓰기(A11) | `LocalBackendOptions.saveAs?`, 새 `viewer/src/images/webp.ts` `webpToPng`·`pngToWebp` |
| 리본 래스터 탭 | `setupWebRibbon(...)` 반환에 `show(tab: string): void` |

## 해석한 스펙 모호점(계획 작성자 결정)

1. **매크로 경고 위치:** 세 편집기 공용 `officeShell`에서 `#toolbar` 바로 뒤에 `.unsupported-warning` 모양 한 줄(`role="status"`). 닫기 단추는 두지 않는다(한 줄, 매크로 문서에서만).
2. **"고침" 판단 중 방이 시딩 중이면** `room.ready`를 기다린 뒤 본다(시딩 쓰기가 잠깐 `editedSinceSave`를 켜므로).
3. **e2e 스크립트의 편집은 워드만 화면으로 한다**(ProseMirror에 글 넣기가 안정적). 엑셀·PPT의 "고친 뒤 원래 형식"은 Task 10 실제 브라우저 조작에서 확인한다.
4. **cadview 이미지 편집에서 그림 크기**는 `localImages.size(이름)`(이미 TIF 화면용 변환 때 잰다), 없으면 `createImageBitmap`으로 잰다.
5. **시험 자료:** Apache POI test-data(Apache-2.0) 고정 커밋 `a4e8640da7c3bb906efc6fa11914d1b80f3cc9dc`의 6개 파일을 `tests/fixtures/legacy/`에 넣는다(sha256은 Task 1에 적었다). 시험 파일 머리 주석에 출처를 적는다.

## 파일 구조

webviewer(`webviewer-view`, 가지 `view-shell`):
- 새로: `src/legacyOffice.ts`, `tests/legacyOffice.test.ts`, `tests/legacyOfficeRoutes.test.ts`, `tests/fixtures/legacy/*`(6개), `scripts/verify-legacy-edit-e2e.mjs`
- 고침: `src/converters/libreOffice.ts`, `src/types.ts`, `src/storage.ts`, `src/routes/documents.ts`, `src/routes/viewer.ts`, `src/collab/room.ts`, `src/client/officeShell.ts`, `src/client/editor.ts`, `src/client/grid-editor.ts`, `src/client/editor-slides.ts`, `public/api-docs.html`, `package.json`, `docs/superpowers/specs/2026-11-12-unified-view-shell-design.md`(§8)

cadview(`cadview-image-edit`, 가지 `image-edit`):
- 새로: `viewer/src/edit/blank.dwg`, `viewer/src/edit/dwg.d.ts`, `viewer/src/images/webp.ts`, `engine/crates/edit-core/tests/blank_template.rs`
- 고침: `viewer/src/local.ts`, `viewer/test/local.test.ts`, `viewer/src/edit/backend-local.ts`, `viewer/src/edit/backend-local.test.ts`, `viewer/src/edit/ribbon-web.ts`, `viewer/src/main.ts`, `viewer/package.json`

calc-site(`calc-site-always-edit`, 가지 `always-edit`):
- 고침: `functions/api/docview/upload.js`, `public/tools/index.html`, `public/assets/i18n-en.js`

---

### Task 1: webviewer — 옛 형식 바꾸기와 매크로 찾기

**작업 폴더:** `C:\Users\OpenMinds\projects\webviewer-view`(가지 `view-shell`).

**Files:**
- Modify: `src/converters/libreOffice.ts`(파일 끝 근처, `convertLegacySpreadsheetToXlsx` 다음)
- Create: `src/legacyOffice.ts`(이 작업에서는 순수 부분만 — `legacyDownload`는 Task 3), `tests/legacyOffice.test.ts`, `tests/fixtures/legacy/`(6개)

**Interfaces:**
- Produces: `convertOfficeFile(data, fromExt, toExt): Promise<Buffer>`(실패는 `ConversionError`), `LEGACY_EDIT_TARGET`, `isLegacyExt(ext): ext is LegacyExt`, `sha256(bytes): string`, `hasVbaMacros(bytes: Buffer, ext: LegacyExt): boolean`
- Consumes: 같은 파일의 `sofficePath`·`withProfileSlot`·`execFileAsync`·`CONVERT_TIMEOUT_MS`·`ConversionError`(이미 import됨 — `convertPdfToOffice`가 쓴다)

- [ ] **Step 1: 시험 자료 받기** — POI 고정 커밋에서 6개(Git Bash):

```bash
cd /c/Users/OpenMinds/projects/webviewer-view
mkdir -p tests/fixtures/legacy
POI=https://raw.githubusercontent.com/apache/poi/a4e8640da7c3bb906efc6fa11914d1b80f3cc9dc/test-data
for p in document/SimpleMacro.doc spreadsheet/SimpleMacro.xls slideshow/SimpleMacro.ppt document/table-merges.doc spreadsheet/chinese-provinces.xls slideshow/SampleShow.ppt; do
  curl -sSfL -o tests/fixtures/legacy/$(basename $p) "$POI/$p"
done
sha256sum tests/fixtures/legacy/*
```

기대 sha256(앞 16자리): `SimpleMacro.doc 39e9608c711d38f2`, `SimpleMacro.xls 0e92c9bb018abd8a`, `SimpleMacro.ppt c6755e4c53635608`, `table-merges.doc bceea684bbf492fa`, `chinese-provinces.xls 0c4f1a10e9586ecd`, `SampleShow.ppt 64275a6685b8f267`. 다르면 멈추고 보고.

- [ ] **Step 2: 실패하는 시험 쓰기** — `tests/legacyOffice.test.ts`

```ts
// 시험 자료: Apache POI test-data(Apache License 2.0), 커밋 a4e8640da7c3bb906efc6fa11914d1b80f3cc9dc
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { describe, expect, test } from 'vitest';
import { convertOfficeFile, sofficePath } from '../src/converters/libreOffice.js';
import { hasVbaMacros, isLegacyExt, LEGACY_EDIT_TARGET, sha256 } from '../src/legacyOffice.js';

const fx = (name: string) => readFileSync(path.resolve('tests', 'fixtures', 'legacy', name));

describe('옛 형식 표', () => {
  test('doc·xls·ppt만 옛 형식, 바꿀 곳은 OOXML', () => {
    expect(['doc', 'xls', 'ppt'].every(isLegacyExt)).toBe(true);
    expect(['docx', 'odt', 'hwp', ''].some(isLegacyExt)).toBe(false);
    expect(LEGACY_EDIT_TARGET).toEqual({ doc: 'docx', xls: 'xlsx', ppt: 'pptx' });
    expect(sha256(Buffer.from('a'))).toBe('ca978112ca1bbdcafac231b39a23dc4da786eff8147c4e72b9807785afee48bb');
  });
});

describe('매크로(VBA) 찾기', () => {
  test('매크로가 든 파일', () => {
    expect(hasVbaMacros(fx('SimpleMacro.doc'), 'doc')).toBe(true);
    expect(hasVbaMacros(fx('SimpleMacro.xls'), 'xls')).toBe(true);
    expect(hasVbaMacros(fx('SimpleMacro.ppt'), 'ppt')).toBe(true);
  });
  test('매크로 없는 파일·복합 파일이 아닌 것', () => {
    expect(hasVbaMacros(fx('table-merges.doc'), 'doc')).toBe(false);
    expect(hasVbaMacros(fx('chinese-provinces.xls'), 'xls')).toBe(false);
    expect(hasVbaMacros(fx('SampleShow.ppt'), 'ppt')).toBe(false);
    expect(hasVbaMacros(Buffer.from('PK\u0003\u0004 not ole'), 'doc')).toBe(false);
  });
});

describe.skipIf(!sofficePath)('LibreOffice로 바꾸기', () => {
  test('doc → docx → doc: OOXML(zip)과 OLE 머리, 글자 유지', async () => {
    const docx = await convertOfficeFile(fx('table-merges.doc'), 'doc', 'docx');
    expect(docx.subarray(0, 2).toString()).toBe('PK');
    const back = await convertOfficeFile(docx, 'docx', 'doc');
    expect(back.readUInt32LE(0)).toBe(0xe011cfd0);
    const again = await convertOfficeFile(back, 'doc', 'docx');
    const text = (b: Buffer) => strFromU8(unzipSync(b, { filter: (f) => f.name === 'word/document.xml' })['word/document.xml']!).replace(/<[^>]+>/g, '');
    expect(text(again)).toBe(text(docx));
  }, 120_000);
  test('xls·ppt도 OOXML로', async () => {
    expect((await convertOfficeFile(fx('chinese-provinces.xls'), 'xls', 'xlsx')).subarray(0, 2).toString()).toBe('PK');
    expect((await convertOfficeFile(fx('SampleShow.ppt'), 'ppt', 'pptx')).subarray(0, 2).toString()).toBe('PK');
  }, 120_000);
  test('모르는 형식은 거절', async () => {
    await expect(convertOfficeFile(Buffer.from('x'), 'doc', 'hwp')).rejects.toThrow('바꿀 수 없는 형식');
  });
});
```

- [ ] **Step 3: 실패 확인** — `npx vitest run tests/legacyOffice.test.ts` → `legacyOffice.js`·`convertOfficeFile`이 없어 실패.

- [ ] **Step 4: `convertOfficeFile` 쓰기** — `src/converters/libreOffice.ts`, `convertLegacySpreadsheetToXlsx` 함수 바로 다음:

```ts
/** 옛 오피스 ↔ OOXML(향기랩 옛 형식 편집, 설계 2026-11-12-hyanggilab-always-edit §2.2) */
const OFFICE_FILTERS: Readonly<Record<string, string>> = {
  docx: 'docx:MS Word 2007 XML',
  doc: 'doc:MS Word 97',
  xlsx: 'xlsx:Calc MS Excel 2007 XML',
  xls: 'xls:MS Excel 97',
  pptx: 'pptx:Impress MS PowerPoint 2007 XML',
  ppt: 'ppt:MS PowerPoint 97',
};

// shortcut: 바꿀 때마다 soffice를 새로 띄운다(수 초) — 느리다는 말이 나오면 loWorker.py에 storeToURL 필터를 더한다
export async function convertOfficeFile(data: Uint8Array, fromExt: string, toExt: string): Promise<Buffer> {
  const filter = OFFICE_FILTERS[toExt];
  if (!filter || !OFFICE_FILTERS[fromExt] || fromExt === toExt) throw new ConversionError(`바꿀 수 없는 형식입니다: .${fromExt} → .${toExt}`);
  if (!sofficePath) throw new ConversionError('이 서버에는 LibreOffice가 없어 변환할 수 없습니다.');
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-office-'));
  try {
    const source = path.join(workDir, `source.${fromExt}`);
    await fs.writeFile(source, data);
    await withProfileSlot(async (profileDir) => {
      await execFileAsync(
        sofficePath!,
        ['--headless', '--norestore', `-env:UserInstallation=${pathToFileURL(profileDir).href}`, '--convert-to', filter, '--outdir', workDir, source],
        { timeout: CONVERT_TIMEOUT_MS },
      );
    });
    const out = await fs.readFile(path.join(workDir, `source.${toExt}`)).catch(() => null);
    if (!out || out.length === 0) throw new ConversionError('변환하지 못했습니다(LibreOffice가 결과를 만들지 않았습니다).');
    return out;
  } finally {
    await fs.rm(workDir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 }).catch(() => undefined);
  }
}
```

- [ ] **Step 5: `src/legacyOffice.ts` 쓰기(순수 부분)**

```ts
/**
 * 향기랩 옛 오피스 형식 편집(설계 2026-11-12-hyanggilab-always-edit §2): .doc·.xls·.ppt를 업로드 때 OOXML로 바꿔 편집기로 열고,
 * 내려받기는 원래 형식(안 고쳤으면 올린 바이트 그대로). 업로드에 legacyEdit=1이 있을 때만(향기랩) — 다른 호출자는 지금 그대로.
 */
import { createHash } from 'node:crypto';
import type { LegacyExt } from './types.js';

export const LEGACY_EDIT_TARGET = { doc: 'docx', xls: 'xlsx', ppt: 'pptx' } as const;

export function isLegacyExt(ext: string): ext is LegacyExt {
  return Object.hasOwn(LEGACY_EDIT_TARGET, ext);
}

export const sha256 = (bytes: Uint8Array): string => createHash('sha256').update(bytes).digest('hex');

const VBA_STORAGES = new Set(['Macros', '_VBA_PROJECT_CUR', 'VBA']);
/** VBAInfoContainer(0x03FF) 머리의 recType·길이 20 — 바로 뒤에 VBAInfoAtom(0x0400, 길이 12)이 와야 한다 */
const PPT_VBA = Buffer.from([0xff, 0x03, 0x14, 0x00, 0x00, 0x00]);

/**
 * 매크로(VBA)가 든 옛 오피스 파일인가(설계 §2.3). 워드·엑셀은 복합 파일 디렉터리 항목(128바이트: 이름 UTF-16LE, +64 이름 길이,
 * +66 종류 1 = 저장소) 중 VBA 저장소 이름, PPT는 PowerPoint Document 흐름의 VBAInfo 레코드. 틀려도 경고 한 줄이 더 뜨거나 안 뜰 뿐이다.
 */
export function hasVbaMacros(b: Buffer, ext: LegacyExt): boolean {
  if (b.length < 512 || b.readUInt32LE(0) !== 0xe011cfd0) return false;
  if (ext === 'ppt') {
    for (let i = b.indexOf(PPT_VBA, 2); i >= 2; i = b.indexOf(PPT_VBA, i + 1)) {
      const h = i - 2;
      if ((b[h]! & 0xf) === 0xf && h + 16 <= b.length && b.readUInt16LE(h + 10) === 0x0400 && b.readUInt32LE(h + 12) === 12) return true;
    }
    return false;
  }
  for (let o = 512; o + 128 <= b.length; o += 128) {
    if (b[o + 66] !== 1) continue;
    const len = b.readUInt16LE(o + 64);
    if (len < 4 || len > 64 || len % 2 !== 0) continue;
    if (b.readUInt16LE(o + len - 2) === 0 && VBA_STORAGES.has(b.toString('utf16le', o, o + len - 2))) return true;
  }
  return false;
}
```

`src/types.ts`에 이 작업에서 타입 하나만 먼저 더한다(Task 2가 쓴다): `export type LegacyExt = 'doc' | 'xls' | 'ppt';`(`DocumentMeta` 바로 위).

- [ ] **Step 6: 통과 확인** — `npx vitest run tests/legacyOffice.test.ts` 전부 통과(LibreOffice 묶음은 이 PC에서 실제로 돈다 — 건너뛰면 `sofficePath`를 확인하고 보고). `npx tsc --noEmit` 오류 없음.

- [ ] **Step 7: 커밋**

```bash
git add src/legacyOffice.ts tests/legacyOffice.test.ts tests/fixtures/legacy/SimpleMacro.doc tests/fixtures/legacy/SimpleMacro.xls tests/fixtures/legacy/SimpleMacro.ppt tests/fixtures/legacy/table-merges.doc tests/fixtures/legacy/chinese-provinces.xls tests/fixtures/legacy/SampleShow.ppt
git commit -m "옛 오피스 형식: LibreOffice로 doc·xls·ppt ↔ OOXML 바꾸기, 매크로(VBA) 찾기(POI 시험 자료)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" -- src/converters/libreOffice.ts src/types.ts src/legacyOffice.ts tests/legacyOffice.test.ts tests/fixtures/legacy
```

---

### Task 2: webviewer — 업로드 `legacyEdit=1`

**작업 폴더:** `C:\Users\OpenMinds\projects\webviewer-view`(가지 `view-shell`).

**Files:**
- Modify: `src/types.ts`(`DocumentMeta.legacy`), `src/storage.ts`(`DerivedKind`), `src/routes/documents.ts`(`ingestUpload`, `/api/documents/stream` 스키마)
- Create: `tests/legacyOfficeRoutes.test.ts`(업로드 묶음 — Task 3이 내려받기 묶음을 더한다)

**Interfaces:**
- Consumes: Task 1의 `convertOfficeFile`·`isLegacyExt`·`LEGACY_EDIT_TARGET`·`hasVbaMacros`·`sha256`
- Produces: `DocumentMeta.legacy?: { readonly ext: LegacyExt; readonly sha256: string; readonly hasMacros: boolean }`, 파생 `'legacy'`, 업로드 응답 `data`의 `legacy`·`editReady`

- [ ] **Step 1: 실패하는 시험 쓰기** — `tests/legacyOfficeRoutes.test.ts`

```ts
// 시험 자료: Apache POI test-data(Apache License 2.0), 커밋 a4e8640da7c3bb906efc6fa11914d1b80f3cc9dc
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { afterAll, beforeAll, describe, expect, test, vi } from 'vitest';
import type { FastifyInstance } from 'fastify';
import { sofficePath } from '../src/converters/libreOffice.js';

const fx = (name: string) => readFileSync(path.resolve('tests', 'fixtures', 'legacy', name));

describe.skipIf(!sofficePath)('향기랩 옛 형식 편집(legacyEdit)', () => {
  let app: FastifyInstance;
  let store: (typeof import('../src/storage.js'))['store'];
  let issueViewToken: (typeof import('../src/security.js'))['issueViewToken'];

  // 저장소는 메모리(향기랩과 같다) — 모듈을 읽기 전에 환경을 바꿔야 해서 동적 import
  beforeAll(async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    ({ store } = await import('../src/storage.js'));
    ({ issueViewToken } = await import('../src/security.js'));
    const { buildApp } = await import('../src/app.js');
    app = await buildApp();
    await app.ready();
  }, 60_000);
  afterAll(async () => {
    await app.close();
    vi.unstubAllEnvs();
  });

  const upload = async (name: string, legacyEdit: boolean) => {
    const res = await app.inject({
      method: 'POST',
      url: `/api/documents/stream?filename=${encodeURIComponent(name)}${legacyEdit ? '&legacyEdit=1' : ''}`,
      headers: { 'content-type': 'application/octet-stream' },
      payload: fx(name),
    });
    return { status: res.statusCode, body: res.json() as { success: boolean; data?: Record<string, any>; error?: string } };
  };
  const tokenFor = (docId: string) =>
    app.inject({ method: 'POST', url: `/api/documents/${docId}/tokens`, payload: { viewer: '시험', canEdit: true, canDownload: true } });

  test('doc: docx로 바꿔 편집기 주소, 이름은 그대로, 올린 바이트는 파생 legacy', async () => {
    const { status, body } = await upload('table-merges.doc', true);
    expect(status).toBe(201);
    const meta = body.data!;
    expect(meta).toMatchObject({ name: 'table-merges.doc', ext: 'docx', editReady: true, legacy: { ext: 'doc', hasMacros: false } });
    expect(meta.legacy.sha256).toMatch(/^[0-9a-f]{64}$/);
    expect((await store.readDerived(meta.id, 'legacy')).equals(fx('table-merges.doc'))).toBe(true);
    expect((await store.readOriginal(meta.id, 'docx')).subarray(0, 2).toString()).toBe('PK');
    const tok = (await tokenFor(meta.id)).json();
    expect(tok.data.editorUrl).toMatch(/^\/editor\.html\?token=/);
  }, 120_000);

  test('xls·ppt: xlsx·pptx로, 매크로 표시', async () => {
    const xls = (await upload('SimpleMacro.xls', true)).body.data!;
    expect(xls).toMatchObject({ name: 'SimpleMacro.xls', ext: 'xlsx', legacy: { ext: 'xls', hasMacros: true } });
    const ppt = (await upload('SimpleMacro.ppt', true)).body.data!;
    expect(ppt).toMatchObject({ name: 'SimpleMacro.ppt', ext: 'pptx', editReady: true, legacy: { ext: 'ppt', hasMacros: true } });
    expect((await tokenFor(ppt.id)).json().data.editorUrl).toMatch(/^\/editor-slides\.html\?token=/);
  }, 120_000);

  test('legacyEdit 없으면 지금 그대로: doc은 편집기 없음, xls는 xlsx로 이름까지', async () => {
    const doc = (await upload('table-merges.doc', false)).body.data!;
    expect(doc).toMatchObject({ name: 'table-merges.doc', ext: 'doc' });
    expect(doc.legacy).toBeUndefined();
    expect(doc.editReady).toBeUndefined();
    // 변환 중(processing)인 doc은 토큰을 아직 못 받으므로 주소 규칙으로 본다 — 편집기 없음(지금 그대로)
    const { editorUrlFor } = await import('../src/routes/documents.js');
    expect(editorUrlFor('doc', { effectiveCanEdit: true, canDownload: true, pdfEditor: false }, 't')).toBeUndefined();
    const xls = (await upload('chinese-provinces.xls', false)).body.data!;
    expect(xls).toMatchObject({ name: 'chinese-provinces.xlsx', ext: 'xlsx' });
    expect(xls.legacy).toBeUndefined();
  }, 120_000);

  test('바꾸지 못하는 옛 파일은 422(열람 화면으로 내려가지 않는다)', async () => {
    const res = await app.inject({
      method: 'POST',
      url: '/api/documents/stream?filename=broken.doc&legacyEdit=1',
      headers: { 'content-type': 'application/octet-stream' },
      payload: Buffer.from('this is not a word file'),
    });
    expect(res.statusCode).toBe(422);
    expect(res.json().error).toContain('편집할 수 있게 바꾸지 못했습니다');
  }, 120_000);

  test('legacyEdit 값은 1만', async () => {
    const res = await app.inject({ method: 'POST', url: '/api/documents/stream?filename=a.doc&legacyEdit=yes', headers: { 'content-type': 'application/octet-stream' }, payload: fx('table-merges.doc') });
    expect(res.statusCode).toBe(400);
  });

  // Task 3이 이 아래에 내려받기 시험을 더한다(issueViewToken을 쓴다)
  void issueViewToken;
});
```

주: LibreOffice가 깨진 .doc를 "글자 파일"로 읽어 docx를 만들어 버리면 422 시험이 실패한다. 그때는 `ingestUpload`가 바꾸기 전에 OLE 머리(`D0 CF 11 E0`)를 확인해 아니면 같은 422를 돌려주게 하고(한 줄, `b.readUInt32LE(0) !== 0xe011cfd0`), 시험은 그대로 둔다.

- [ ] **Step 2: 실패 확인** — `npx vitest run tests/legacyOfficeRoutes.test.ts` → 업로드가 doc 그대로(ext `doc`)라 실패.

- [ ] **Step 3: 타입·저장소**

`src/types.ts` `DocumentMeta`의 `editReady` 다음:

```ts
  /** 향기랩 옛 형식 편집(설계 2026-11-12-hyanggilab-always-edit A2): 올린 .doc·.xls·.ppt를 OOXML로 바꿔 ext에 두고, 원래 형식·바꾼 직후
   * OOXML의 sha256(고쳤는지 판단)·매크로 유무를 적는다. 올린 바이트는 파생 'legacy'. 업로드에 legacyEdit=1이 있을 때만 생긴다 */
  readonly legacy?: { readonly ext: LegacyExt; readonly sha256: string; readonly hasMacros: boolean };
```

`src/storage.ts`: `export type DerivedKind = 'pdf' | 'grid' | 'markup' | 'pagesizes' | 'gridindex' | 'svg' | 'legacy';`, `DERIVED_SUFFIX`에 `legacy: 'legacy.bin',`, 위 주석 끝에 `/ legacy(향기랩 옛 형식 편집 — 올린 .doc·.xls·.ppt 바이트)`.

- [ ] **Step 4: `ingestUpload`** — `src/routes/documents.ts`

import 두 줄 더하기:

```ts
import { convertDxfToSvg, convertLegacySpreadsheetToXlsx, convertOfficeFile } from '../converters/libreOffice.js';
import { hasVbaMacros, isLegacyExt, LEGACY_EDIT_TARGET, sha256 } from '../legacyOffice.js';
import type { DocumentMeta, LegacyExt, ViewKind } from '../types.js';
```

시그니처에 매개변수 하나:

```ts
async function ingestUpload(
  app: FastifyInstance,
  originalFilename: string,
  data: Buffer,
  callerProject: string | undefined,
  legacyEdit = false,
): Promise<UploadResult> {
```

zip 검사 다음, 지금 `if (ext === 'xls' || ext === 'ods') {`를 `else if`로 바꾸고 그 앞에:

```ts
  // 향기랩(legacyEdit=1)만: 옛 오피스를 OOXML로 바꿔 편집기로 연다 — 표시 이름은 그대로, 올린 바이트는 파생 'legacy'(설계 §2.1)
  let legacy: { ext: LegacyExt; bytes: Buffer; hasMacros: boolean } | null = null;
  if (legacyEdit && isLegacyExt(ext)) {
    const legacyExt = ext;
    try {
      const converted = await convertOfficeFile(data, legacyExt, LEGACY_EDIT_TARGET[legacyExt]);
      legacy = { ext: legacyExt, bytes: data, hasMacros: hasVbaMacros(data, legacyExt) };
      data = converted;
      ext = LEGACY_EDIT_TARGET[legacyExt];
    } catch (error) {
      app.log.warn(error, 'legacy office conversion failed');
      return {
        code: 422,
        body: { success: false, error: `이 .${legacyExt} 파일을 편집할 수 있게 바꾸지 못했습니다. 파일이 손상되었거나 암호가 걸려 있을 수 있습니다.` },
      };
    }
  } else if (ext === 'xls' || ext === 'ods') {
```

저장 부분 — `await store.saveOriginal(id, ext, data);` 다음 줄:

```ts
    if (legacy) await store.saveDerived(id, 'legacy', legacy.bytes);
```

`meta` 객체의 `editReady` 펼침 다음 줄:

```ts
      ...(legacy ? { legacy: { ext: legacy.ext, sha256: sha256(data), hasMacros: legacy.hasMacros } } : {}),
```

- [ ] **Step 5: 경로 스키마** — 같은 파일 `/api/documents/stream`

`app.post<{ Querystring: { filename: string; uploadId?: string; index?: number; total?: number; legacyEdit?: '1' } }>(`, querystring `properties`에

```ts
            // 향기랩만 보낸다 — 옛 오피스(.doc·.xls·.ppt)를 편집용 OOXML로 바꿔 받는다(설계 2026-11-12-hyanggilab-always-edit A1)
            legacyEdit: { type: 'string', enum: ['1'] },
```

처리부 마지막 `ingestUpload(app, request.query.filename, data, callerProject)`를 `ingestUpload(app, request.query.filename, data, callerProject, request.query.legacyEdit === '1')`로. multipart `/api/documents`는 그대로.

- [ ] **Step 6: 통과 확인** — `npx vitest run tests/legacyOfficeRoutes.test.ts tests/storage-derivedkinds.test.ts tests/routes-documents.test.ts`, `npx tsc --noEmit`.

- [ ] **Step 7: 커밋**

```bash
git add tests/legacyOfficeRoutes.test.ts
git commit -m "업로드 legacyEdit=1(향기랩만): .doc·.xls·.ppt를 OOXML로 바꿔 편집기로, 올린 바이트·해시·매크로 표시 보관 — 없으면 지금 그대로

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" -- src/types.ts src/storage.ts src/routes/documents.ts tests/legacyOfficeRoutes.test.ts
```

---

### Task 3: webviewer — 원래 형식으로 내려받기

**작업 폴더:** `C:\Users\OpenMinds\projects\webviewer-view`(가지 `view-shell`).

**Files:**
- Modify: `src/collab/room.ts`(`hasUnsavedEdits` 추가 — `getCurrentBytes` 바로 위), `src/legacyOffice.ts`(`legacyDownload`), `src/routes/viewer.ts`(`/api/view/download`), `tests/legacyOfficeRoutes.test.ts`(내려받기 시험), `tests/legacyOffice.test.ts`(실패 갈래)

**Interfaces:**
- Produces: `hasUnsavedEdits(docId: string): Promise<boolean>`, `legacyDownload(meta: DocumentMeta, convert?: typeof convertOfficeFile): Promise<LegacyDownload>`, `LegacyDownload = { bytes: Buffer; name: string; how: 'original' | 'converted' | 'fallback' }`
- Consumes: `getCurrentBytes`(room.ts), `store`, `convertOfficeFile`, `sha256`

- [ ] **Step 1: 실패하는 시험 더하기** — `tests/legacyOfficeRoutes.test.ts`의 `void issueViewToken;` 줄을 지우고 그 자리에:

```ts
  const download = (docId: string) => {
    const t = issueViewToken({ docId, viewer: '시험', perms: { canPrint: false, canDownload: true, canCopy: false, canEdit: true } }, 600);
    return app.inject({ method: 'GET', url: `/api/view/download?token=${encodeURIComponent(t)}` });
  };
  const fileName = (cd: string | string[] | undefined) => decodeURIComponent(String(cd).split("''")[1] ?? '');

  test('안 고친 내려받기는 올린 바이트 그대로·원래 이름', async () => {
    for (const name of ['table-merges.doc', 'chinese-provinces.xls', 'SampleShow.ppt']) {
      const meta = (await upload(name, true)).body.data!;
      const res = await download(meta.id);
      expect(res.statusCode).toBe(200);
      expect(res.rawPayload.equals(fx(name))).toBe(true);
      expect(fileName(res.headers['content-disposition'])).toBe(name);
    }
  }, 180_000);

  test('고친(저장된 OOXML이 바뀐) 내려받기는 원래 형식으로 다시 바꾼다', async () => {
    const { makeSampleDocx } = await import('./helpers.js');
    const { convertOfficeFile } = await import('../src/converters/libreOffice.js');
    const meta = (await upload('table-merges.doc', true)).body.data!;
    await store.saveOriginal(meta.id, 'docx', makeSampleDocx(['고친글자LEGACY']));
    const res = await download(meta.id);
    expect(res.statusCode).toBe(200);
    expect(res.rawPayload.readUInt32LE(0)).toBe(0xe011cfd0);
    expect(fileName(res.headers['content-disposition'])).toBe('table-merges.doc');
    const back = await convertOfficeFile(res.rawPayload, 'doc', 'docx');
    const { unzipSync, strFromU8 } = await import('fflate');
    expect(strFromU8(unzipSync(back, { filter: (f) => f.name === 'word/document.xml' })['word/document.xml']!)).toContain('고친글자LEGACY');
  }, 180_000);
```

`tests/legacyOffice.test.ts` 끝에 실패 갈래(LibreOffice 없이 도는 순수 시험 — 저장소를 메모리로):

```ts
describe('legacyDownload: 다시 바꾸기 실패면 편집본을 OOXML 이름으로', () => {
  test('fallback', async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.resetModules();
    const { store, newDocumentId } = await import('../src/storage.js');
    const { legacyDownload } = await import('../src/legacyOffice.js');
    const id = newDocumentId();
    await store.saveOriginal(id, 'docx', Buffer.from('편집본'));
    await store.saveDerived(id, 'legacy', Buffer.from('올린 원본'));
    const meta = { id, name: '보고서.doc', ext: 'docx', status: 'ready', pageCount: 1, createdAt: '', legacy: { ext: 'doc', sha256: '다름', hasMacros: false } } as const;
    const out = await legacyDownload(meta, async () => {
      throw new Error('soffice 없음');
    });
    expect(out).toEqual({ bytes: Buffer.from('편집본'), name: '보고서.docx', how: 'fallback' });
    const same = await legacyDownload({ ...meta, legacy: { ...meta.legacy, sha256: (await import('../src/legacyOffice.js')).sha256(Buffer.from('편집본')) } });
    expect(same).toEqual({ bytes: Buffer.from('올린 원본'), name: '보고서.doc', how: 'original' });
    vi.unstubAllEnvs();
  });
});
```

(같은 파일 맨 위 import에 `vi`를 더한다: `import { describe, expect, test, vi } from 'vitest';`)

- [ ] **Step 2: 실패 확인** — 두 시험 파일을 돌려 내려받기가 docx 바이트·`legacyDownload` 없음으로 실패.

- [ ] **Step 3: `hasUnsavedEdits`** — `src/collab/room.ts`, `getCurrentBytes` 위 문서 주석 블록 바로 앞:

```ts
/** 열린 방에 아직 저장하지 않은 편집이 있는가(방이 없으면 false). 시딩 쓰기가 잠깐 표시를 켜므로 시딩이 끝난 뒤에 본다 —
 * 향기랩 옛 형식 내려받기가 "고쳤는가"를 정할 때 쓴다(src/legacyOffice.ts) */
export async function hasUnsavedEdits(docId: string): Promise<boolean> {
  const room = rooms.get(docId);
  if (!room) return false;
  await room.ready;
  return room.editedSinceSave;
}
```

- [ ] **Step 4: `legacyDownload`** — `src/legacyOffice.ts` import 보태고 끝에 더하기:

```ts
import { getCurrentBytes, hasUnsavedEdits } from './collab/room.js';
import { convertOfficeFile } from './converters/libreOffice.js';
import { store } from './storage.js';
import type { DocumentMeta, LegacyExt } from './types.js';
```

```ts
export interface LegacyDownload {
  readonly bytes: Buffer;
  readonly name: string;
  readonly how: 'original' | 'converted' | 'fallback';
}

/**
 * 옛 형식 문서 내려받기(설계 §2.4). 안 고쳤으면(방에 저장 안 된 편집 없음 + 저장된 OOXML 해시가 바꾼 직후와 같음) 올린 바이트 그대로,
 * 고쳤으면 지금 내용을 LibreOffice로 원래 형식으로. 그게 실패하면 고친 것을 잃지 않게 편집본을 OOXML 이름으로(A4).
 */
export async function legacyDownload(meta: DocumentMeta, convert = convertOfficeFile): Promise<LegacyDownload> {
  const legacy = meta.legacy!;
  const stored = await store.readOriginal(meta.id, meta.ext);
  const edited = (await hasUnsavedEdits(meta.id)) || sha256(stored) !== legacy.sha256;
  if (!edited) return { bytes: await store.readDerived(meta.id, 'legacy'), name: meta.name, how: 'original' };
  const current = Buffer.from((await getCurrentBytes(meta.id)) ?? stored);
  try {
    return { bytes: await convert(current, meta.ext, legacy.ext), name: meta.name, how: 'converted' };
  } catch (error) {
    console.error(`옛 형식으로 다시 바꾸지 못해 편집본을 .${meta.ext}로 내려준다(문서 ${meta.id}):`, error);
    return { bytes: current, name: `${meta.name.replace(/\.[^./]+$/, '')}.${meta.ext}`, how: 'fallback' };
  }
}
```

(`LegacyExt` import는 Task 1 것과 합친다.)

- [ ] **Step 5: 내려받기 경로** — `src/routes/viewer.ts` `/api/view/download`, 마크업 분기(`if (request.query.includeMarkup === '1') { … }`) 바로 뒤, `const data = (await getCurrentBytes(…` 앞:

```ts
      // 향기랩 옛 형식(.doc·.xls·.ppt): 원래 형식으로(설계 2026-11-12-hyanggilab-always-edit §2.4)
      if (view.meta.legacy) {
        const out = await legacyDownload(view.meta);
        await recordAudit({ type: 'downloaded', docId: view.meta.id, viewer: view.claims.viewer, detail: `legacy ${out.how}` });
        return reply
          .type('application/octet-stream')
          .header('content-disposition', `attachment; filename*=UTF-8''${encodeURIComponent(out.name)}`)
          .send(out.bytes);
      }
```

import `import { legacyDownload } from '../legacyOffice.js';`.

- [ ] **Step 6: 통과 확인** — `npx vitest run tests/legacyOffice.test.ts tests/legacyOfficeRoutes.test.ts`, 전체 `npx vitest run`, `npx tsc --noEmit`.

- [ ] **Step 7: 커밋**

```bash
git commit -m "옛 형식 내려받기: 안 고쳤으면 올린 바이트 그대로, 고쳤으면 LibreOffice로 원래 형식, 실패하면 편집본(OOXML 이름)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" -- src/collab/room.ts src/legacyOffice.ts src/routes/viewer.ts tests/legacyOffice.test.ts tests/legacyOfficeRoutes.test.ts
```

---

### Task 4: webviewer — 편집기 안내(형식 이름·내려받기 문구·매크로 경고)

**작업 폴더:** `C:\Users\OpenMinds\projects\webviewer-view`(가지 `view-shell`).

**Files:**
- Modify: `src/routes/viewer.ts`(`/api/view/info` 응답), `src/client/officeShell.ts`, `src/client/editor.ts`·`src/client/grid-editor.ts`·`src/client/editor-slides.ts`(각 `ViewInfo`에 `legacy`, `initOfficeShell`에 넘김), `tests/viewShellRoutes.test.ts`(info 시험 하나)

**Interfaces:**
- Produces: `/api/view/info` `data.legacy?: { ext: 'doc' | 'xls' | 'ppt'; hasMacros: boolean }`, `OfficeShellOptions.legacy?: { readonly ext: string; readonly hasMacros: boolean }`

- [ ] **Step 1: 실패하는 시험** — `tests/viewShellRoutes.test.ts`의 `describe` 안 끝에:

```ts
  test('info: 옛 형식이면 legacy(형식·매크로), 아니면 없음', async () => {
    const id = await addDoc('docx');
    const meta = (await store.readMeta(id))!;
    await store.writeMeta({ ...meta, name: '시험.doc', legacy: { ext: 'doc', sha256: 'x', hasMacros: true } });
    const res = await get(`/api/view/info?token=${encodeURIComponent(tok(id, { canDownload: true }))}`);
    expect(res.json().data.legacy).toEqual({ ext: 'doc', hasMacros: true });
    const plain = await addDoc('docx');
    expect((await get(`/api/view/info?token=${encodeURIComponent(tok(plain, { canDownload: true }))}`)).json().data.legacy).toBeUndefined();
  });
```

`npx vitest run tests/viewShellRoutes.test.ts` → 실패 확인.

- [ ] **Step 2: info 응답** — `src/routes/viewer.ts` `/api/view/info`의 `data`에 `processing` 다음 줄:

```ts
          // 향기랩 옛 형식(.doc·.xls·.ppt) — 편집기가 형식 이름·내려받기 문구·매크로 경고를 보인다(설계 §2.5)
          legacy: view.meta.legacy ? { ext: view.meta.legacy.ext, hasMacros: view.meta.legacy.hasMacros } : undefined,
```

- [ ] **Step 3: officeShell** — `src/client/officeShell.ts`

옵션(`viewerUrl` 다음):

```ts
  /** 향기랩 옛 형식(.doc·.xls·.ppt)을 바꿔 연 문서 — 내려받기는 원래 형식(서버 /api/view/download), 매크로가 있으면 경고 한 줄 */
  readonly legacy?: { readonly ext: string; readonly hasMacros: boolean };
```

`EXT_LABELS`에 세 줄(형식 칸은 이미 `docName` 확장자로 고르므로 `보고서.doc`이면 이 이름이 나온다):

```ts
  doc: 'Word 97-2003 문서',
  xls: 'Excel 97-2003 통합 문서',
  ppt: 'PowerPoint 97-2003 프레젠테이션',
```

다운로드 칸 — 지금 `options.canDownload ? '<p>지금 편집 중인 내용 그대로 …</p><a … data-action="download">원본 형식으로 다운로드</a>' : '<p class="backstage-denied">…</p>'` 식을 아래로 바꾼다(`legacy.ext`는 서버가 정한 세 값 중 하나라 그대로 넣어도 된다):

```ts
          options.canDownload
            ? options.legacy
              ? `<p>원래 형식(.${options.legacy.ext})으로 내려받습니다. 고치지 않았으면 올린 파일 그대로이고, 고쳤으면 LibreOffice로 다시 바꾸므로 모양이 조금 달라질 수 있습니다.</p><a class="backstage-action" data-action="download">${options.legacy.ext.toUpperCase()} 파일로 다운로드</a>`
              : '<p>지금 편집 중인 내용 그대로 원본 형식 파일을 내려받습니다.</p><a class="backstage-action" data-action="download">원본 형식으로 다운로드</a>'
            : '<p class="backstage-denied">이 문서는 다운로드가 허용되지 않았습니다.</p>'
```

매크로 경고 — `initOfficeShell` 안, `const toolbar = …` 바로 다음:

```ts
  // 매크로는 고치지 않고 내려받으면 남고(같은 바이트), 고친 뒤 내려받으면 빠진다 — 고치기 전에 알린다(설계 A5)
  if (options.legacy?.hasMacros) {
    const note = document.createElement('div');
    note.className = 'unsupported-warning legacy-macro-note';
    note.setAttribute('role', 'status');
    note.textContent = '이 문서에는 매크로(VBA)가 있습니다. 고치지 않고 내려받으면 원본 그대로이지만, 고친 뒤 내려받으면 매크로가 빠집니다.';
    toolbar.after(note);
  }
```

(세 편집기 HTML 모두 `/css/style.css`를 싣고 `.unsupported-warning` 규칙이 거기 있다 — CSS는 고치지 않는다.)

- [ ] **Step 4: 세 편집기** — `src/client/editor.ts`·`src/client/grid-editor.ts`·`src/client/editor-slides.ts` 각각:
  - `interface ViewInfo`에 `readonly legacy?: { readonly ext: string; readonly hasMacros: boolean };`
  - `initOfficeShell({ … viewerUrl: info.viewerUrl,` 다음 줄에 `legacy: info.legacy,`

- [ ] **Step 5: 확인** — `npx vitest run`(전체), `npx tsc --noEmit`, `npm run build:editor`.

- [ ] **Step 6: 커밋**

```bash
git commit -m "편집기: 옛 형식 문서는 형식 이름(97-2003)·원래 형식 내려받기 문구, 매크로가 있으면 리본 아래 경고 한 줄

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" -- src/routes/viewer.ts src/client/officeShell.ts src/client/editor.ts src/client/grid-editor.ts src/client/editor-slides.ts tests/viewShellRoutes.test.ts
```

---

### Task 5: webviewer — e2e 스크립트·API 문서·통합 보기 설계 §8

**작업 폴더:** `C:\Users\OpenMinds\projects\webviewer-view`(가지 `view-shell`).

**Files:**
- Create: `scripts/verify-legacy-edit-e2e.mjs`
- Modify: `package.json`(scripts: `"verify:legacy-edit": "node scripts/verify-legacy-edit-e2e.mjs"`, `verify:view-shell` 다음 줄), `public/api-docs.html`(stream 쿼리 표), `docs/superpowers/specs/2026-11-12-unified-view-shell-design.md`(§8 둘째 문단)

- [ ] **Step 1: e2e 스크립트** — `scripts/verify-legacy-edit-e2e.mjs`(돌리는 것은 Task 10):

```js
// 향기랩 옛 형식 편집(legacyEdit) 확인 — 시험 서버(기본 http://127.0.0.1:8141, 메모리 저장 권장).
//   node scripts/verify-legacy-edit-e2e.mjs [BASE] [스크린샷 폴더]
// doc·xls·ppt를 legacyEdit=1로 올리면 편집기 주소, 안 고친 내려받기는 같은 바이트, 워드는 화면에서 고친 뒤 .doc로 받아 다시 올리면
// 고친 글자가 남는다. legacyEdit 없이 올린 doc은 지금처럼 편집기가 없다(다른 호출자 회귀). 시험 자료: tests/fixtures/legacy(Apache POI).
import crypto from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8141';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'legacy-edit-'));
fs.mkdirSync(shotDir, { recursive: true });
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};
const fx = (name) => fs.readFileSync(path.resolve('tests', 'fixtures', 'legacy', name));
const hash = (b) => crypto.createHash('sha256').update(b).digest('hex');

async function upload(name, legacyEdit = true, bytes = fx(name)) {
  const q = `filename=${encodeURIComponent(name)}${legacyEdit ? '&legacyEdit=1' : ''}`;
  const r = await fetch(`${BASE}/api/documents/stream?${q}`, { method: 'POST', headers: { 'content-type': 'application/octet-stream' }, body: bytes });
  return { status: r.status, json: await r.json() };
}
/** 변환 중(processing)이면 토큰이 404 — 열릴 때까지 다시 묻는다 */
async function token(id) {
  for (let i = 0; i < 240; i++) {
    const r = await fetch(`${BASE}/api/documents/${id}/tokens`, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ viewer: '옛형식시험', canEdit: true, canDownload: true, canPrint: true, canCopy: true }),
    });
    if (r.status === 200) return (await r.json()).data;
    await new Promise((res) => setTimeout(res, 500));
  }
  throw new Error(`토큰을 받지 못함: ${id}`);
}
async function download(tok) {
  const r = await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(tok)}`);
  const name = decodeURIComponent((r.headers.get('content-disposition') ?? '').split("''")[1] ?? '');
  return { status: r.status, name, bytes: Buffer.from(await r.arrayBuffer()) };
}

// 1) 안 고친 내려받기 = 같은 바이트·원래 이름, 편집기 주소, 매크로 표시
for (const [name, editor, macros] of [
  ['table-merges.doc', '/editor.html', false],
  ['chinese-provinces.xls', '/editor-grid.html', false],
  ['SampleShow.ppt', '/editor-slides.html', false],
  ['SimpleMacro.xls', '/editor-grid.html', true],
]) {
  const up = await upload(name);
  check(up.status === 201 && up.json.data.legacy?.ext === name.split('.').pop(), `${name}: 올림(legacy 표시)`, `status ${up.status}`);
  check(up.json.data.legacy?.hasMacros === macros, `${name}: 매크로 표시 ${macros}`);
  const t = await token(up.json.data.id);
  check(t.editorUrl?.startsWith(editor), `${name}: 편집기 주소 ${editor}`, t.editorUrl);
  const d = await download(t.token);
  check(d.status === 200 && hash(d.bytes) === hash(fx(name)), `${name}: 안 고친 내려받기는 같은 바이트`);
  check(d.name === name, `${name}: 내려받기 이름`, d.name);
}

// 2) 다른 호출자 회귀: legacyEdit 없는 doc은 편집기 없음
{
  const up = await upload('table-merges.doc', false);
  const t = await token(up.json.data.id);
  check(up.json.data.ext === 'doc' && !up.json.data.legacy && !t.editorUrl, 'legacyEdit 없는 doc: 지금처럼 열람(편집기 주소 없음)');
}

// 3) 워드: 화면에서 고친 뒤 .doc로 받고, 그 .doc를 다시 올려 편집기에서 고친 글자를 본다 + 매크로 경고 줄
const browser = await chromium.launch();
try {
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  const up = await upload('SimpleMacro.doc');
  const t = await token(up.json.data.id);
  await page.goto(`${BASE}${t.editorUrl}&home=hidden&back=parent`);
  await page.waitForSelector('.ProseMirror', { timeout: 60000 });
  check(await page.locator('.legacy-macro-note').isVisible(), '매크로 경고 줄이 보인다');
  await page.click('.ProseMirror');
  await page.keyboard.press('Control+End');
  await page.keyboard.type(' 옛형식편집LEGACY');
  await page.waitForTimeout(1500);
  await fetch(`${BASE}/api/view/save?token=${encodeURIComponent(t.token)}`, { method: 'POST' });
  await page.screenshot({ path: path.join(shotDir, 'doc-edited.png') });
  const d = await download(t.token);
  check(d.bytes.readUInt32LE(0) === 0xe011cfd0 && d.name === 'SimpleMacro.doc', '고친 뒤 내려받기: .doc(OLE)', d.name);
  check(hash(d.bytes) !== hash(fx('SimpleMacro.doc')), '고친 뒤 내려받기: 올린 바이트와 다름');
  const again = await upload('SimpleMacro.doc', true, d.bytes);
  const t2 = await token(again.json.data.id);
  await page.goto(`${BASE}${t2.editorUrl}&home=hidden&back=parent`);
  await page.waitForSelector('.ProseMirror', { timeout: 60000 });
  check((await page.locator('.ProseMirror').innerText()).includes('옛형식편집LEGACY'), '다시 올린 .doc에 고친 글자가 남았다');
  check(errors.length === 0, '편집기 페이지 오류 없음', errors.join(' | '));
} finally {
  await browser.close();
}

console.log(`\n${results.filter(Boolean).length}/${results.length} 통과 — 스크린샷 ${shotDir}`);
process.exit(results.every(Boolean) ? 0 : 1);
```

- [ ] **Step 2: package.json** — `"verify:view-shell": …` 줄 끝에 쉼표, 다음 줄 `"verify:legacy-edit": "node scripts/verify-legacy-edit-e2e.mjs"`.

- [ ] **Step 3: api-docs** — `public/api-docs.html` stream "쿼리 파라미터" 표의 `filename` 줄 다음:

```html
              <tr><td><code>legacyEdit</code></td><td><code>1</code></td><td>(향기랩 전용) <code>.doc</code>·<code>.xls</code>·<code>.ppt</code>를 LibreOffice로 <code>.docx</code>·<code>.xlsx</code>·<code>.pptx</code>로 바꿔 편집기로 연다. 표시 이름은 그대로, 응답 <code>data.ext</code>는 바꾼 형식, <code>data.legacy</code>에 원래 형식. 내려받기는 원래 형식(안 고쳤으면 올린 바이트 그대로). 바꾸지 못하면 <code>422</code>. 없으면 지금과 같다(<code>.doc</code>·<code>.ppt</code>는 열람만, <code>.xls</code>는 <code>.xlsx</code>로).</td></tr>
```

- [ ] **Step 4: 통합 보기 설계 §8** — `docs/superpowers/specs/2026-11-12-unified-view-shell-design.md` §8 둘째 문단(「단 향기랩에서도 **편집기가 없는 .doc·.ppt는 새 보기 화면으로 열린다**」로 시작해 「(최종 검토 I1, 2026-11-12 결정).」으로 끝나는 문단)을 통째로:

```markdown
향기랩의 .doc·.xls·.ppt는 이제 업로드 때 OOXML로 바꿔 편집기로 연다(`legacyEdit=1`, `2026-11-12-hyanggilab-always-edit-design.md` — 이 자리의 예전 결정 I1을 대체). 새 보기 화면으로 열리는 .doc·.ppt는 `legacyEdit`을 보내지 않는 연동 제품뿐이다.
```

- [ ] **Step 5: 확인** — `node --check scripts/verify-legacy-edit-e2e.mjs`, `npx vitest run`, `npx tsc --noEmit`.

- [ ] **Step 6: 커밋**

```bash
git add scripts/verify-legacy-edit-e2e.mjs
git commit -m "옛 형식 편집: e2e(verify:legacy-edit)·API 문서 legacyEdit·통합 보기 설계 §8을 새 설계로 대체

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" -- scripts/verify-legacy-edit-e2e.mjs package.json public/api-docs.html docs/superpowers/specs/2026-11-12-unified-view-shell-design.md
```

---

### Task 6: cadview — 워크트리와 빈 도면 틀(`blank.dwg`)

**작업 폴더:** 새 워크트리 `C:\Users\OpenMinds\projects\cadview-image-edit`, 가지 `image-edit`.

**Files:**
- Create: `viewer/src/edit/blank.dwg`(바이너리, 수십 KB), `engine/crates/edit-core/tests/blank_template.rs`

**Interfaces:**
- Produces: 편집 가능한 R2000 빈 도면 바이트(Task 8이 묶음에 싣는다)

- [ ] **Step 1: 워크트리**

```bash
cd /c/Users/OpenMinds/projects/cadview
git worktree add -b image-edit ../cadview-image-edit e1-editor-core
cd ../cadview-image-edit/viewer && npm ci
```

주 폴더의 커밋 안 된 세 파일은 따라오지 않는다(정상). `testdata/`는 git 밖이므로 필요하면 `node tools/testdata/fetch.mjs`(이 작업의 시험은 testdata를 쓰지 않는다).

- [ ] **Step 2: 빈 DXF 만들기(scratchpad, 저장소 밖)** — `S="<세션 scratchpad>/blank"`:

```bash
mkdir -p "$S" && python -m venv "$S/venv" && "$S/venv/Scripts/python" -m pip install -q ezdxf==1.4.5
"$S/venv/Scripts/python" -I -c "import ezdxf; d = ezdxf.new('R2000', setup=False); d.header['\$INSUNITS'] = 0; d.saveas(r'$S/blank.dxf')"
```

- [ ] **Step 3: DWG로 바꾸기 — LibreDWG 0.14 `dxf2dwg`**

```bash
mkdir -p "$S/dl" "$S/libredwg"
gh release download 0.14 -R LibreDWG/libredwg -p 'libredwg-0.14-win64.zip' -D "$S/dl" --clobber
unzip -o -q "$S/dl/libredwg-0.14-win64.zip" -d "$S/libredwg"
"$S/libredwg/dxf2dwg.exe" -y -o "$S/blank.dwg" "$S/blank.dxf"
cp "$S/blank.dwg" viewer/src/edit/blank.dwg   # 워크트리 뿌리에서
```

**멈춤 조건:** 내려받은 실행 파일 실행이 권한 규칙으로 막히면(계획 작성 때 실제로 막혔다) 우회하지 말고 멈춰 사용자에게 위 `dxf2dwg` 한 줄을 대신 돌려 달라고(또는 허용해 달라고) 묻는다. 시험 자료 도면(`testdata/example_*.dwg`, LibreDWG 저장소 것)을 지워서 틀을 만들지 않는다(설계 A9 — 출처).

- [ ] **Step 4: 엔진 시험** — `engine/crates/edit-core/tests/blank_template.rs`

```rust
//! 빈 도면 틀(viewer/src/edit/blank.dwg — 향기랩 이미지 편집, webviewer 설계 2026-11-12-hyanggilab-always-edit A8·A9):
//! 편집 가능한 판이고, 그림을 원점·축척 1로 붙여 래스터로 열고 고치고 저장할 수 있다
use cad_raster::{encode, fixture, Format};
use edit_core::EditCore;
use serde_json::{json, Map, Value};

const BLANK: &[u8] = include_bytes!("../../../../viewer/src/edit/blank.dwg");

fn obj(v: Value) -> Map<String, Value> {
    v.as_object().unwrap().clone()
}

#[test]
fn blank_template_takes_an_image_and_saves_it() {
    let mut c = EditCore::new(Box::new(|| scene::FontSet::from_memory(Vec::new())), 0);
    let o = c.open(BLANK.to_vec()).unwrap();
    assert_eq!(o["editable"], json!(true), "틀은 편집할 수 있는 판이어야 한다: {o:?}");
    let (px, _) = fixture::scan_sized(300, 200, 7);
    // 뷰어 addImageOp([0,0], 이름, [w,h], 1)와 같은 op
    let a = c
        .handle("apply", &obj(json!({"batch":{"seq":1,"ops":[{"op":"add_entity","kind":"image","layer":"0","geometry":{
            "at":[0,0,0],"u":[1,0,0],"v":[0,1,0],"size":[300,200],"path":"사진.tif"}}]}})))
        .unwrap();
    let image = a["change"]["added"][0].as_str().unwrap().to_string();
    let r = c.raster_open(&obj(json!({"image":image,"maxTexture":1024})), encode(&px, Format::Tiff, None).unwrap().bytes).unwrap();
    assert_eq!((r["width"].as_u64(), r["height"].as_u64()), (Some(300), Some(200)));
    c.handle("apply", &obj(json!({"batch":{"seq":2,"ops":[{"op":"raster","image":image,"do":"erase","shape":{"rect":[[10,10],[20,20]]}}]}})))
        .unwrap();
    let src = c.handle("raster_query", &obj(json!({"what":"source","image":image}))).unwrap();
    let (bytes, _) = c.raster_save_bytes(&obj(json!({"def": src["def"]}))).unwrap();
    assert!(bytes.starts_with(b"II*\0") || bytes.starts_with(b"MM\0*"), "원래 형식(TIF)으로 저장");
}
```

(`scan_sized`·`raster_save_bytes`·`raster_query` 이름과 `def` 모양은 `tests/raster.rs`의 `opened_in`·`new_image_is_one_step_with_add_image`와 같다 — 다르면 그 파일을 따른다.)

`cd engine && cargo test -p edit-core --test blank_template` 통과. `editable`이 false거나 열기 오류면: (1) Step 2를 `setup=True`로 다시, (2) 그래도 안 되면 멈추고 보고(대안: 사용자가 설치한 ODA File Converter로 같은 DXF를 ACAD2018 DWG로 — 사용자 결정).

- [ ] **Step 5: 커밋**

```bash
git add viewer/src/edit/blank.dwg engine/crates/edit-core/tests/blank_template.rs
git commit -m "빈 도면 틀(R2000, ezdxf 빈 DXF → LibreDWG dxf2dwg): 이미지 편집용 그릇, 엔진 시험으로 편집·그림 붙이기·래스터 저장 확인

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" -- viewer/src/edit/blank.dwg engine/crates/edit-core/tests/blank_template.rs
```

---

### Task 7: cadview — `imageEdit` 메시지와 "그림만 저장" 백엔드

**작업 폴더:** `C:\Users\OpenMinds\projects\cadview-image-edit`(가지 `image-edit`), 명령은 `viewer/`에서.

**Files:**
- Modify: `viewer/src/local.ts`, `viewer/test/local.test.ts`, `viewer/src/edit/backend-local.ts`, `viewer/src/edit/backend-local.test.ts`

**Interfaces:**
- Produces: `acceptOpenMessage(...)` → `{ file; images; hidden; imageEdit: boolean } | null`, `LocalBackendOptions.imageOnly?: boolean`, `LocalBackendOptions.saveAs?: (name: string, bytes: Uint8Array<ArrayBuffer>) => Promise<{ name: string; bytes: Uint8Array<ArrayBuffer> }>`

- [ ] **Step 1: 실패하는 시험**

`viewer/test/local.test.ts`: 첫 `it`의 `toEqual({ file, images: [], hidden: false })`를 `toEqual({ file, images: [], hidden: false, imageEdit: false })`로(같은 모양의 다른 `toEqual`도 모두 `imageEdit: false`를 더한다), `describe('acceptOpenMessage')` 안에 더하기:

```ts
  it('imageEdit: true일 때만 켠다(향기랩 이미지 편집)', () => {
    const open = (imageEdit: unknown) => acceptOpenMessage({ source: parent, origin: env.origin, data: { type: 'cadview:open', file, imageEdit } }, env);
    expect(open(true)?.imageEdit).toBe(true);
    expect(open('true')?.imageEdit).toBe(false);
    expect(open(undefined)?.imageEdit).toBe(false);
  });
```

`viewer/src/edit/backend-local.test.ts` 끝에:

```ts
describe('이미지 편집 모드(imageOnly·saveAs)', () => {
  function imageSetup(saveAs?: ConstructorParameters<typeof LocalBackend>[0]['saveAs']) {
    const w = new FakeWorker();
    const picked: string[] = [];
    const written: number[][] = [];
    const b = new LocalBackend({
      worker: w,
      input: async () => new Blob([]),
      fileName: () => '사진.png.dwg',
      imageOnly: true,
      saveAs,
      pickTarget: async (name) => {
        picked.push(name);
        return { name, write: async (bytes) => void written.push([...bytes]) };
      },
    });
    // raster_query dirty → 그림 하나, raster_saved → ok
    w.replies.push(JSON.stringify({ ok: true, items: [{ def: '2A', name: '사진.png' }] }), JSON.stringify({ ok: true }));
    w.rasterSaveReply = { json: '{"ok":true}', bytes: new Uint8Array([1, 2, 3]) };
    return { w, b, picked, written };
  }
  it('저장은 그림만 — 도면 저장(editSave)·도면 이름 고르기 없음, "도면 옆" 꼬리 없음', async () => {
    const { w, b, picked, written } = imageSetup();
    const r = await b.save();
    expect(picked).toEqual(['사진.png']);
    expect(written).toEqual([[1, 2, 3]]);
    expect(w.order).not.toContain('save');
    expect(r.canceled).toBeFalsy();
    expect(r.text).toContain('사진.png');
    expect(r.text).not.toContain('도면 옆');
  });
  it('saveAs가 이름·바이트를 바꾼다(WEBP로 다시 쓰기)', async () => {
    const { b, picked, written } = imageSetup(async (name) => ({ name: name.replace(/\.png$/, '.webp'), bytes: new Uint8Array([9]) }));
    const r = await b.save();
    expect(picked).toEqual(['사진.webp']);
    expect(written).toEqual([[9]]);
    expect(r.text).toContain('사진.webp');
  });
});
```

`npx vitest run test/local.test.ts src/edit/backend-local.test.ts` → 실패 확인.

- [ ] **Step 2: `local.ts`**

```ts
export function acceptOpenMessage(ev: { source: unknown; origin: string; data: unknown }, env: FrameEnv): { file: Blob; images: File[]; hidden: boolean; imageEdit: boolean } | null {
  if (env.parent === env.self || ev.source !== env.parent || ev.origin !== env.origin) return null;
  const d = ev.data as { type?: unknown; file?: unknown; images?: unknown; hidden?: unknown; imageEdit?: unknown } | null;
  if (!d || typeof d !== 'object' || d.type !== 'cadview:open' || !(d.file instanceof Blob)) return null;
  return { file: d.file, images: imageFiles(d.images), hidden: d.hidden === true, imageEdit: d.imageEdit === true };
}
```

아래 계약 주석의 `hidden?: true …` 줄 다음에: ` *    imageEdit?: true — 이미지 파일을 빈 도면 위 래스터 편집으로 연다(저장은 그림만, 원래 형식). 없으면 지금처럼 문자 인식 이미지 보기. 향기랩이 보낸다.`

- [ ] **Step 3: `backend-local.ts`**

`LocalBackendOptions`에(`onPicked` 다음):

```ts
  /** 이미지 편집 모드(향기랩 imageEdit): 저장(Ctrl+S·SAVE)은 그림만, 빈 도면(.dwg)은 쓰지 않는다 */
  imageOnly?: boolean;
  /** 그림을 쓰기 직전 이름·바이트를 바꾼다(WEBP로 다시 쓰기 — 엔진은 WEBP를 못 쓴다). 다른 이름으로(as)에는 쓰지 않는다 */
  saveAs?: (name: string, bytes: Uint8Array<ArrayBuffer>) => Promise<{ name: string; bytes: Uint8Array<ArrayBuffer> }>;
```

`rasterSave`의 그림 고리(`for (const it of items) {`) 앞 세 줄을 이렇게:

```ts
      const { r, bytes } = await this.rasterBytes({ def: it.def, ...(it.format ? { as: true, format: it.format } : {}) });
      const out = this.o.saveAs && o.as === undefined ? await this.o.saveAs(it.name, bytes) : { name: it.name, bytes };
      const t = await this.o.pickTarget(out.name);
      if (!t) return { ...canceled, saved, warn };
      await this.writeTo(t, out.bytes);
```

같은 고리의 `if (imageKey(t.name) !== imageKey(it.name) && …)`는 `imageKey(t.name) !== imageKey(out.name)`로(다시 쓴 이름은 "다른 이름으로 고름"이 아니다), `saved.push({ name: nm, size: bytes.length })`는 `saved.push({ name: out.name !== it.name ? out.name : nm, size: out.bytes.length })`로. `replaceOriginal(nm, bytes)`는 그대로(탭 안 저장소는 엔진이 읽는 PNG 그대로 둔다).

고리 뒤 꼬리: `if (saved.length) text += ' — 도면 옆의 같은 이름 파일을 이것으로 바꿔 두세요';` → `if (saved.length && !this.o.imageOnly) text += …`.

`save()` 맨 앞:

```ts
    // 이미지 편집 모드: 그림만 내려받는다(빈 도면은 보이지 않는 그릇 — 설계 A10)
    if (this.o.imageOnly) {
      const ri = await this.rasterSave({});
      if (ri.canceled) return { canceled: true, text: ri.text };
      return { text: ri.text, ...(ri.unconfirmed ? { unconfirmed: true } : {}), ...(ri.warn.length ? { warn: ri.warn.join(' · ') } : {}) };
    }
```

- [ ] **Step 4: 확인** — `npx vitest run`(전체), `npm run typecheck`.

- [ ] **Step 5: 커밋**

```bash
git commit -m "browser-local: 부모 메시지 imageEdit, 편집 백엔드 imageOnly(저장은 그림만)·saveAs(쓰기 직전 형식 바꾸기)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" -- viewer/src/local.ts viewer/test/local.test.ts viewer/src/edit/backend-local.ts viewer/src/edit/backend-local.test.ts
```

---

### Task 8: cadview — 이미지를 빈 도면 위 래스터 편집으로 열기

**작업 폴더:** `C:\Users\OpenMinds\projects\cadview-image-edit`(가지 `image-edit`), 명령은 `viewer/`에서.

**Files:**
- Create: `viewer/src/edit/dwg.d.ts`, `viewer/src/images/webp.ts`
- Modify: `viewer/src/main.ts`, `viewer/src/edit/ribbon-web.ts`, `viewer/package.json`(`build`의 main.ts esbuild에 `--loader:.dwg=binary`)

**Interfaces:**
- Consumes: `blank.dwg`(Task 6), `acceptOpenMessage().imageEdit`·`imageOnly`·`saveAs`(Task 7), `addImageOp`(`edit/tools.ts`), `fitCamera`(`camera.ts`), `localImages.size`
- Produces: `setupWebRibbon(...)` → `{ refresh(): void; show(tab: string): void }`, `webpToPng(file: Blob, name: string): Promise<File>`, `pngToWebp(pngName, bytes, webpName)`

- [ ] **Step 1: 바이너리 가져오기**

`viewer/src/edit/dwg.d.ts`:

```ts
// esbuild --loader:.dwg=binary(package.json build) — 빈 도면 틀(blank.dwg) 바이트
declare module '*.dwg' {
  const bytes: Uint8Array;
  export default bytes;
}
```

`viewer/package.json` `build`: `esbuild src/main.ts --bundle --format=iife …` 부분에 `--loader:.dwg=binary`를 더한다(demo.ts 쪽은 그대로). `grep -l "src/main" test/*.ts src/**/*.test.ts`가 비어 있는지 확인한다 — main.ts를 읽는 단위 시험이 있으면 vitest가 `.dwg`를 못 읽으므로 멈추고 보고.

- [ ] **Step 2: `viewer/src/images/webp.ts`**

```ts
/** WEBP는 엔진이 읽고 쓰지 못한다(설계 A11): 열 때 브라우저로 PNG로 풀고, 저장 때 브라우저가 WEBP를 쓸 수 있으면 다시 WEBP로 */
async function draw(src: Blob): Promise<OffscreenCanvas> {
  const bmp = await createImageBitmap(src);
  const c = new OffscreenCanvas(bmp.width, bmp.height);
  c.getContext('2d')!.drawImage(bmp, 0, 0);
  bmp.close();
  return c;
}

export async function webpToPng(file: Blob, name: string): Promise<File> {
  const png = await (await draw(file)).convertToBlob({ type: 'image/png' });
  return new File([png], name.replace(/\.webp$/i, '.png'), { type: 'image/png' });
}

/** 사파리는 WEBP를 쓰지 못해 PNG를 돌려준다 — 그때는 .png 이름 그대로(저장 문구에 그 이름이 나온다) */
export async function pngToWebp(pngName: string, bytes: Uint8Array<ArrayBuffer>, webpName: string): Promise<{ name: string; bytes: Uint8Array<ArrayBuffer> }> {
  const out = await (await draw(new Blob([bytes], { type: 'image/png' }))).convertToBlob({ type: 'image/webp', quality: 1 });
  return out.type === 'image/webp' ? { name: webpName, bytes: new Uint8Array(await out.arrayBuffer()) } : { name: pngName, bytes };
}
```

- [ ] **Step 3: 리본 탭** — `viewer/src/edit/ribbon-web.ts` `setupWebRibbon`의 반환 타입 `{ refresh(): void; show(tab: string): void }`, 마지막 `return { refresh … }`에 `show: (tab) => r.show(tab),`를 더한다(지금 반환문 모양을 따른다).

- [ ] **Step 4: `main.ts` — 모듈 수준 상태**(`let imagePicked …` 선언 근처)

```ts
import blankDwg from './edit/blank.dwg';
import { addImageOp } from './edit/tools';
import { pngToWebp, webpToPng } from './images/webp';
```

```ts
/** 향기랩 이미지 편집(설계 A8): 빈 도면을 그린 뒤 그림을 붙인다 — setup이 잇는다(세션·카메라·리본이 그 안에 있다) */
let imageEditHook: ((img: File) => Promise<void>) | null = null;
/** 이미지 편집 모드: 편집 백엔드가 그림만 저장, WEBP면 저장 직전 다시 WEBP로 */
let imageOnly = false;
let imageSaveAs: ((name: string, bytes: Uint8Array<ArrayBuffer>) => Promise<{ name: string; bytes: Uint8Array<ArrayBuffer> }>) | undefined;
const IMAGE_EDIT_RE = /\.(png|jpe?g|bmp|tiff?|webp)$/i;
```

`makeEditBackend`의 `new LocalBackend({ … onPicked: … })`에 `imageOnly, saveAs: imageSaveAs,` 두 항목(백엔드는 setup에서 만들어지므로 그 전에 값이 정해진다).

- [ ] **Step 5: `startLocal`**

`offer` 시그니처에 `imageEdit = false`를 마지막으로 더하고, `act === 'open'` 줄을:

```ts
    if (act === 'open') {
      if (imageEdit && IMAGE_EDIT_RE.test(name)) return void openImageForEdit(file, name);
      return void (DOC_RE.test(name) ? ((body.dataset.state = 'loading'), openDoc(file, name)) : open(file, name, images));
    }
```

`open` 정의 다음에:

```ts
  /** 향기랩 이미지 편집: 빈 도면 틀을 열고(그림은 이 탭 메모리 저장소로), 그린 뒤 setup이 그림을 붙인다(설계 §4.2) */
  const openImageForEdit = async (file: Blob, name: string) => {
    imageOnly = true;
    const webp = /\.webp$/i.test(name);
    let img: File;
    try {
      img = webp ? await webpToPng(file, name) : new File([file], name);
    } catch {
      errorText.textContent = '이미지를 열지 못했습니다';
      return notify({ type: 'cadview:error', kind: 'load', message: '이미지를 열지 못했습니다' });
    }
    if (webp) imageSaveAs = (n, bytes) => pngToWebp(n, bytes, name);
    await open(new Blob([blankDwg]), name, [img]);
    if (body.dataset.state === 'ready') await imageEditHook?.(img);
  };
```

부모 메시지 처리의 `offer(m.file, m.file instanceof File ? m.file.name : 'drawing.dwg', true, m.images);`를 `offer(m.file, m.file instanceof File ? m.file.name : 'drawing.dwg', true, m.images, m.imageEdit);`로.

- [ ] **Step 6: `setup` 안 — 그림 붙이기**

웹 리본 블록의 `const ribbon = setupWebRibbon(…)`을 바깥에서 쓸 수 있게 `let ribbon: ReturnType<typeof setupWebRibbon> | null = null;`을 블록 앞에 두고 블록 안에서 `ribbon = setupWebRibbon(…)`으로(아래 `bridge.subscribe(ribbon.refresh)`는 `bridge.subscribe(ribbon.refresh)` 그대로 — 블록 안이라 null 아님). 그 블록 다음에:

```ts
  // 향기랩 이미지 편집(설계 §4.2): 그림을 원점·축척 1(1 px = 1 단위)로 붙이고, 그림에 맞춰 보고, 래스터 탭을 연다.
  // 래스터 명령은 이미지가 하나뿐이면 그것을 바로 연다(rasterCmd → onlyImage)
  imageEditHook = async (img) => {
    let size = localImages?.size(img.name) ?? null;
    if (!size) {
      const bmp = await createImageBitmap(img);
      size = { width: bmp.width, height: bmp.height };
      bmp.close();
    }
    if (!session || !(await session.ensure())) return;
    if (!(await session.apply([addImageOp([0, 0], img.name, [size.width, size.height], 1)], `그림을 열었습니다: ${img.name}`))) return;
    const [cx, cy] = root.center;
    cam = fitCamera({ min: [-cx, -cy], max: [size.width - cx, size.height - cy] }, viewW(), stage.clientHeight);
    redraw();
    ribbon?.show('raster');
  };
```

(`fitCamera`의 상자 모양·`viewW`·`root`·`cam`·`redraw` 이름은 같은 함수 안 `fit`(1195줄 근처)이 쓰는 것과 같다 — 다르면 그쪽을 따른다. `fitCamera`가 받는 타입 이름이 `Bbox`면 `{min, max}` 그대로 맞는다.)

- [ ] **Step 7: 확인** — `npm run typecheck`, `npx vitest run`, `npm run build`, `npm run build:local`(wasm 빌드 포함 — 몇 분). 브라우저 확인은 Task 10.

- [ ] **Step 8: 커밋**

```bash
git add viewer/src/edit/dwg.d.ts viewer/src/images/webp.ts
git commit -m "browser-local 이미지 편집(imageEdit): 빈 도면 틀 위에 그림을 원점·축척 1로 붙여 래스터 탭으로, 저장은 그림만·원래 형식(WEBP는 브라우저가 되면 WEBP)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" -- viewer/src/edit/dwg.d.ts viewer/src/images/webp.ts viewer/src/main.ts viewer/src/edit/ribbon-web.ts viewer/package.json
```

---

### Task 9: calc-site — 업로드 표시·경로표(nwd 삭제)·imageEdit·편집기만

**작업 폴더:** 새 워크트리 `C:\Users\OpenMinds\calc-site-always-edit`, 가지 `always-edit`(`cadview-ocr`에서).

**Files:**
- Modify: `functions/api/docview/upload.js`, `public/tools/index.html`, `public/assets/i18n-en.js`

- [ ] **Step 1: 워크트리**

```bash
cd /c/Users/OpenMinds/calc-site
git worktree add -b always-edit ../calc-site-always-edit cadview-ocr
cd ../calc-site-always-edit
```

- [ ] **Step 2: `upload.js`** — `for (const key of ['uploadId', 'index', 'total']) { … }` 다음 줄:

```js
  // 옛 오피스(.doc·.xls·.ppt)를 편집용 OOXML로 바꿔 받는다 — 이 함수는 향기랩만 쓴다(webviewer 설계 2026-11-12-hyanggilab-always-edit A1).
  // 내려받기는 원래 형식. 조각 업로드는 조각마다 붙인다(마지막 조각이 처리)
  upstream.searchParams.set('legacyEdit', '1');
```

머리 주석의 "요청:" 줄 아래에 `//   위로 보낼 때 늘 legacyEdit=1을 붙인다(옛 오피스 편집).`

- [ ] **Step 3: `tools/index.html`**

- `accept=".docx,.doc,.xlsx,.xls,.pptx,.ppt,.pdf,.hwp,.hwpx,.dwg,.nwd,.png,…"`에서 `.nwd,` 삭제.
- 지원 형식 줄(`data-i18n="tools.dz.formats"`)의 글을: `지원 형식: 워드(docx·doc) · 엑셀(xlsx·xls) · 파워포인트(pptx·ppt) · PDF · 한글(hwp/hwpx) · 도면(dwg) · 이미지(png·jpg·tif 등 — 편집·글자 인식)`
- 안내 표: 워드 줄 확장자 칸 `docx` → `docx, doc`, 엑셀 `xlsx` → `xlsx, xls`, PPT `pptx` → `pptx, ppt`. 도면 줄 확장자 `dwg, nwd` → `dwg`, 할 수 있는 일 `보기, 길이 재기, 인쇄 · nwd(Navisworks)는 3D 보기만` → `보기·편집, 길이 재기, 인쇄`. 도면 줄 다음에 새 줄:

```html
        <tr><td data-i18n="tools.guide.img.kind">이미지</td><td>png, jpg, bmp, tif, webp</td><td data-i18n="tools.guide.img.can">그림 고치기(지우개·선·글자 등), 같은 형식으로 다시 받기, 글자 인식</td><td data-i18n="tools.guide.img.where">내 브라우저 안(서버로 보내지 않음)</td></tr>
```

- `ROUTES`: `// Navisworks NWD도 …` 주석과 `nwd: 'cadview-local',` 줄 삭제. 이미지 주석을 `// 이미지(사진·스캔)는 같은 뷰어에서 그림 편집으로(브라우저 안 — 서버로 보내지 않는다, 글자 인식은 리본). PDF는 위 PDF 편집기 그대로`로.
- `openViaCadviewLocal`: 주석 `// 이미지는 변환하지 않고 문자 인식 보기로 연다 — 문구도 그에 맞게` → `// 이미지는 빈 도면 위 그림 편집으로 연다(imageEdit) — 문구도 그에 맞게`, `postMessage({ type: 'cadview:open', file: file }, …)` → `postMessage({ type: 'cadview:open', file: file, imageEdit: image }, location.origin)`.
- `showWebviewerFrame`: `const path = tokenJson.data.editorUrl || tokenJson.data.viewerUrl;` →

```js
    // 향기랩은 늘 편집 모드로 연다 — 편집기 주소가 없으면 열람 화면 대신 오류(webviewer 설계 2026-11-12-hyanggilab-always-edit A7)
    const path = tokenJson.data.editorUrl;
    if (!path) throw new Error(T('tools.err.noEditor', '이 문서는 편집기로 열 수 없습니다.'));
```

- message 처리: `if (… event.data.type === 'cadview:replace' && event.data.file instanceof Blob && frameOrigin === location.origin) openViaCadviewLocal(event.data.file);` → 끝을 `openFile(event.data.file);`로(확장자 경로표를 다시 탄다 — 뷰어 안에 nwd를 놓아도 받지 않는다, 설계 A12).

- [ ] **Step 4: `i18n-en.js`**

- `'tools.dz.formats'`: `'Supported formats: Word (docx/doc) · Excel (xlsx/xls) · PowerPoint (pptx/ppt) · PDF · Hangul (hwp/hwpx) · Drawing (dwg) · Image (png/jpg/tif — edit and text recognition)'`
- `'tools.guide.dwg.can'`: `"View and edit, measure distances, print"`
- 새 키(dwg 키들 다음): `'tools.guide.img.kind': "Image",`, `'tools.guide.img.can': "Edit the picture (eraser, lines, text…), download in the same format, text recognition",`, `'tools.guide.img.where': "Inside your browser (not sent to a server)",`
- 새 키(`tools.err.open` 다음): `'tools.err.noEditor': 'This document cannot be opened in an editor.',`

- [ ] **Step 5: 확인** — `npm test`(`i18n-keys.test.js`가 새 키를 본다), `grep -n -i "nwd\|navisworks" public/tools/index.html public/assets/i18n-en.js` 결과 없음.

- [ ] **Step 6: 커밋**

```bash
git commit -m "문서도구: 늘 편집 모드 — 옛 오피스는 legacyEdit로 편집기, 이미지는 그림 편집(imageEdit), NWD 제외, 편집기 주소가 없으면 오류

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" -- functions/api/docview/upload.js public/tools/index.html public/assets/i18n-en.js
```

---

### Task 10: 전체 검증 — 단위·e2e·회귀·실제 브라우저(병합 전 멈춤)

**작업 폴더:** 세 워크트리. 운영 포트·주 작업 폴더는 건드리지 않는다.

- [ ] **Step 1: 단위·빌드 한 번 더** — webviewer `npx vitest run && npx tsc --noEmit && npm run build:editor`; cadview `viewer/`에서 `npx vitest run && npm run typecheck`, `engine/`에서 `cargo test -p edit-core`; calc-site `npm test`.

- [ ] **Step 2: 서버 띄우기(run_in_background, PID 적기)** — 포트 확인 뒤:

```bash
S="<세션 scratchpad>"
cd /c/Users/OpenMinds/projects/webviewer-view && npm run build:editor
# 향기랩 흉내(메모리)
DOCVIEW_PORT=8141 DOCVIEW_HOST=127.0.0.1 DOCVIEW_STORAGE=memory DOCVIEW_SECRET=e2e-test-secret npx tsx src/server.ts
# 회귀 verify용(디스크)
DOCVIEW_PORT=8143 DOCVIEW_HOST=127.0.0.1 DOCVIEW_DATA_DIR="$S/legacy-data" DOCVIEW_SECRET=e2e-test-secret npx tsx src/server.ts
```

- [ ] **Step 3: webviewer e2e** — `npm run verify:legacy-edit -- http://127.0.0.1:8141`, 회귀 `npm run verify:view-shell -- http://127.0.0.1:8143`, `npm run verify:view-only -- http://127.0.0.1:8143`, `npm run verify:xlsx -- http://127.0.0.1:8143`, `npm run verify:pptx -- http://127.0.0.1:8143`(각 스크립트의 BASE 인자 모양은 머리 주석을 따른다). 전부 통과.

- [ ] **Step 4: 향기랩 로컬** — cadview 묶음을 calc-site 워크트리에 싣고(커밋하지 않는다) wrangler로:

```bash
cd /c/Users/OpenMinds/projects/cadview-image-edit/viewer && npm run publish:hyanggilab -- C:/Users/OpenMinds/calc-site-always-edit
cd /c/Users/OpenMinds/calc-site-always-edit
npx wrangler pages dev public --port 8142 -b DOCVIEW_ORIGIN=http://127.0.0.1:8141 -b DOCVIEW_API_KEY=local-test   # run_in_background, PID 기록
```

(webviewer에 API 키를 두지 않았으므로 아무 키나 지난다.)

- [ ] **Step 5: 실제 브라우저 조작**(Playwright, 헤드리스면 UA와 `navigator.webdriver`를 바꿔 봇 거르기를 피한다 — calc-site README) `http://127.0.0.1:8142/tools/`에서, 단추·입력은 모두 실제로 눌러 본다:
  1. `table-merges.doc`: 편집기가 바로 열림(열람 화면 아님) → 글 입력 → 파일 → 다운로드 → `.doc`, OLE 머리, 다시 열면 입력한 글. 고치지 않은 새 문서는 내려받은 바이트 sha256 = 원본.
  2. `chinese-provinces.xls`: 칸 값 바꾸기 → `.xls` 내려받기 → 다시 올려 값 확인. 안 고친 것은 같은 바이트.
  3. `SampleShow.ppt`: 글상자 글 바꾸기 → `.ppt` 내려받기 → 다시 올려 확인. 안 고친 것은 같은 바이트.
  4. `SimpleMacro.doc`: 매크로 경고 줄 보임, 파일 → 정보 "Word 97-2003 문서", 다운로드 칸 문구.
  5. 이미지 png·jpg·tif·webp(직접 만든 시험 그림): 래스터 탭이 열린 채 뜸, 지우개·선 그리기 → Ctrl+S → 같은 이름·같은 형식으로 저장(webp는 크롬에서 `.webp`), 저장 문구에 "도면 옆" 없음, 리본 OCR로 글자 인식 결과가 나옴, 실행 취소.
  6. DWG(`testdata/sample_2018.dwg`): 편집 켜기 없이 선 긋기 → Ctrl+S 내려받기.
  7. NWD: 고르기에서 `.nwd` 안 보임(accept), 끌어 놓기면 "지원하지 않는 파일 형식" 안내, 도면이 떠 있는 뷰어 안에 `.nwd`를 놓아도 받지 않음.
  8. hwp(rhwp)·pdf(PDF 편집기) 여전히 편집기로.
  9. 휴대폰 폭(390px)에서 1·5번을 한 번씩.

  스크린샷과 sha256 비교 결과를 scratchpad에 남긴다.

- [ ] **Step 6: 정리** — 띄운 서버를 PID로 끄고(`taskkill //PID <pid> //T //F`), calc-site 워크트리의 실은 묶음을 되돌린다: `git -C /c/Users/OpenMinds/calc-site-always-edit checkout -- public/cadview` 뒤 `git -C /c/Users/OpenMinds/calc-site-always-edit status --short public/cadview`에 남은 새 파일만 그 경로에서 지운다(이 워크트리 안, 내가 만든 것만). 세 워크트리는 병합 때 쓰므로 남긴다.

- [ ] **Step 7: 보고하고 멈춤** — 세 가지·커밋 sha, 시험 결과, 실제 브라우저 확인 목록, 알려진 한계(설계 §3, webp 사파리 PNG, 옛 형식 바꾸기 수 초). **병합(webviewer master = 8000·8010 운영 반영, cadview e1-editor-core, calc-site master 배포)은 사용자 승인 뒤 따로** — cadview 판 올리기와 향기랩 `/cadview/` 묶음 커밋도 그때.

---

## 자체 점검(계획 작성자)

- **설계 범위:** A1 업로드 표시(Task 2·9), A2 메타·파생(Task 2), A3 고침 판단(Task 3), A4 실패 갈래(Task 3), A5 매크로 경고(Task 1·4), A6 422(Task 2), A7 편집기만(Task 9), A8 빈 도면 + 그림(Task 6·8), A9 틀 출처(Task 6), A10 그림만 저장(Task 7), A11 WEBP(Task 7·8), A12 replace → openFile(Task 9). §5 DWG·§7 통합 보기 §8(Task 5·10). §3 한계는 설계·보고에만.
- **빠진 것 없음 확인:** 8000 회귀는 Task 2 시험(legacyEdit 없음)·Task 5 e2e 2)·Task 10 회귀 verify. cadview 데스크톱·8020은 메시지를 받지 않아 바뀌지 않음(Task 7 기본값 false).
- **이름 일관:** `convertOfficeFile`·`legacyDownload`·`hasUnsavedEdits`·`hasVbaMacros`·`imageOnly`·`saveAs`·`imageEdit`·`imageEditHook`·`show(tab)` — 정의한 작업과 쓰는 작업이 같다.
- **위험:** Task 6 `dxf2dwg` 실행이 권한으로 막힐 수 있고(계획 작성 때 막혔다), LibreDWG가 쓴 R2000을 엔진이 편집 가능으로 열지 아직 확인하지 못했다 — Task 6이 멈춤 조건을 둔다.
