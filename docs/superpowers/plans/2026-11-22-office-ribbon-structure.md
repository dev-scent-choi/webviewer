# 리본을 실제 오피스 구조로(워드·엑셀·PPT 전체 탭·묶음·단추) 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 세 편집기(워드 `editor.html`·엑셀 `editor-grid.html`·PPT `editor-slides.html`)의 리본을 Microsoft 365 한국어판의 탭·묶음·단추 구조 그대로 갖추고, 이미 있는 기능은 오피스의 같은 자리로 옮기며(id·처리기·단축키 그대로), 없는 기능은 보이되 꺼진 "준비 중 — <이름>" 단추로 둔다.

**Architecture:** 전체 목록은 `src/client/ribbon/inventory.json` 한 파일(탭 → 묶음 → 단추, 이름·크기·종류·기존 선택자 또는 준비 중)이고, `scripts/gen-ribbon.mjs`가 이 목록으로 각 편집기 HTML의 두 표시(`<!-- ribbon:tabs:… -->`·`<!-- ribbon:panels:… -->`) 사이를 **정적 마크업으로 다시 만든다**. 기존 단추는 첫 속성 `data-rb="<탭>.<묶음>.<키>"`로 찾아 마크업을 글자 그대로 옮기고, 준비 중 단추·대리 단추·접힌 묶음 단추·"자세히"는 새로 쓴다. 결과 HTML은 결정적이라(두 번 돌려도 같음, `--check`) 그대로 커밋하고 빌드 없이 열린다. 브라우저에서 하는 일은 화면 크기에 따라야 하는 두 가지뿐 — 고른 탭의 준비 중 아이콘 달기(탭마다 늦게), 좁을 때 오피스처럼 묶음을 큰 단추 하나로 접기(`src/client/ribbonRuntime.ts`).

**왜 정적 생성인가(런타임 렌더러 대신):** 세 편집기는 시작할 때 `getElementById`·`querySelectorAll('[data-insert-chart]')` 등으로 정적 마크업에 처리기를 단다. 마크업을 정적으로 두면 (1) id·처리기·단축키가 손대지 않고 그대로 붙고(초기화 순서 문제 없음), (2) 모든 e2e 선택자(`#btn-…`, `.ribbon-tab[data-tab="…"]`)가 그대로 맞으며, (3) 첫 화면에서 리본이 바뀌어 보이는 깜빡임이 없고, (4) 결과가 커밋된 HTML이라 리뷰어가 diff로 본다. 런타임에 DOM을 옮기는 방식은 초기화 순서·깜빡임·마크업과 실제 화면의 불일치를 새로 만든다.

**Tech Stack:** Node 20 · Fastify 5 · TypeScript · esbuild(`npm run build:editor`) · vitest · Playwright · PowerShell 5.1 + UI 자동화(M365 대조용) · 손으로 쓴 CSS·SVG(새 의존성 없음)

**Spec:** `docs/superpowers/specs/2026-11-22-office-ribbon-structure-design.md` (모양은 계속 `docs/superpowers/specs/2026-11-22-artcraft-ribbon-skin-design.md`, 사용자 규칙 `feedback_artcraft_ui_rules.md` 6번이 이 계획)

## Global Constraints

- 리본 구조는 **Microsoft 365(한국어) 워드·엑셀·파워포인트**의 탭 이름·순서, 묶음 이름·순서, 묶음 안 단추의 종류·이름·큰/작은 구분·드롭다운 여부까지 그대로(설계 §1-1). 기본 탭 목록은 설계 §2 그대로(도움말 탭 없음). 묶음·단추는 Task 1의 M365 UIA 대조 결과가 이 계획의 목록보다 우선.
- 이미 있는 기능은 오피스의 같은 자리에 둔다 — **id·처리기·단축키 유지**(설계 §1-2). 기존 id를 지우거나 바꾸지 않는다(`tests/fixtures/ribbon-ids-before.json` 대조). 새 id를 만들지 않는다.
- 아직 없는 기능: 단추는 보이되 `disabled` + `data-pending` + 말풍선 `title="준비 중 — <기능 이름>"`(설계 §1-2). 기능 구현은 하지 않는다(설계 §3).
- 읽기 전용(ui=view): 쓰는 단추는 `.ro-disabled`(0.45), **준비 중 단추는 `disabled`가 우선**이라 `.ro-disabled`를 붙이지 않는다. 접힌 묶음 단추(둘러보기용)도 흐리지 않는다(설계 §1-3).
- 콘텐츠만 보기(`ui=content`)는 리본 전체 숨김 — 건드리지 않는다. 보기 화면(`view.html`)은 범위 밖(설계 §1-4·5). 파일 탭(백스테이지) 화면은 그대로, 오피스 항목은 목록 `backstage`에만(설계 §3).
- 글자는 전부 한글(사용자 규칙 4, `verify:ribbon-skin` (2)). 오피스 이름이 영문만인 것(SmartArt·WordArt·SharePoint)은 한글 표기("스마트아트"·"워드아트"·"셰어포인트"), 목록 `note`에 적는다. 한글이 섞인 이름("3D 모델", "SmartArt 디자인", "100%")은 그대로.
- 꾸밈 규격은 그대로: 제목 줄 24 · 탭 줄 22 · 리본 62 · 상태 줄 18px(±2), 큰 단추 48px·아이콘 32px, 작은 단추 16px·아이콘 16px, 묶음 이름 줄 14px. 900·800px 판 안 스크롤바·이름 줄 겹침 없음, 600px 가로 스크롤 없음(`verify:ribbon-skin` 그대로 통과).
- 넘침: 오피스처럼 좁으면 **묶음을 큰 단추 하나(묶음 이름 ▾)로 접고**, 누르면 원래 단추들이 그 아래 떠 있는 판으로 펼쳐진다. 접는 순서는 목록의 묶음별 `collapse`(1~9, 작은 수 먼저, 같으면 오른쪽 묶음 먼저). 601px 이상은 판이 넘치지 않을 때까지 접고, 600px 이하는 모든 묶음을 접는다(기존 600px 줄 바꿈 규칙 위에서 아이콘 한 줄).
- 다른 에이전트가 같은 파일에서 한 일(아이콘 32/16px·이름 표기, "읽기 전용"·"편집 열기"·"연결됨" 숨김, `ui=content`)은 이미 들어와 있다 — 다시 하지 않는다. 그 위에서 작업한다.
- 포트 8000·8010·8020·8100·8110·8130·8148은 쓰지도 건드리지도 않는다. 이 계획의 시험 서버는 **8180~8189**(새 e2e 기본 8180, 회귀용 수동 서버 8181). 프로세스는 직접 띄운 것만 PID로 끈다(`taskkill /PID <번호> /T /F`, 이름으로 node·Office를 끄지 않음).
- 커밋하는 것: `src/client/ribbon/inventory.json`, 생성된 `public/editor*.html`, `public/ribbon/icons.svg`, 스크립트·시험·TS·CSS. 커밋하지 않는 것: `public/js/*.bundle.js`, `.audit-run/`(M365 덤프·캡처·시험 결과).
- 커밋은 경로를 지정해서(`git add <경로>`), 커밋 전에 `git diff --cached --stat`으로 내 변경만인지 본다(다른 에이전트의 변경이 섞였으면 멈추고 보고). 메시지는 한국어, 끝줄 `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`. 구현자는 push 하지 않는다. 가지 `office-artcraft`.
- 시험 순서(사용자 규칙): 태스크마다 단위 시험·타입 검사·빌드·`gen-ribbon --check`만, 브라우저 e2e·회귀 묶음·실제 브라우저 조작·M365 나란히 보기는 Task 7에서 한 번에.

---

## 파일 구조

| 파일 | 할 일 | 태스크 |
|---|---|---|
| `src/client/ribbon/inventory.json` | **전체 목록**(이 계획과 함께 이미 작성됨 — 워드 324·엑셀 298·PPT 333 단추). 앱 → `tabs`·`contextual`(상황별, `trigger`·`wired`) → `groups`(`collapse`·`layout`·`launcher`) → `controls`. `summary`는 수 요약, `schema`는 필드 설명 | 1(대조·고침) |
| `scripts/ribbonGenLib.mjs` + `.d.mts` | 목록 읽기·검사·요약·쓰기(Part A), 마크업 만들기·기대 모양(Part B). node 전용 | 1, 2 |
| `scripts/dump-m365-ribbon.ps1` | 이 PC의 한국어 M365(16.0.19127, 언어 ko-kr)를 COM으로 열고 UI 자동화로 탭·묶음·단추 이름·높이를 JSON으로, 탭마다 리본 캡처 PNG(`.audit-run/m365-ribbon/`) | 1 |
| `scripts/compare-ribbon-inventory.mjs` | 덤프 ↔ 목록 차이를 한 줄씩 | 1 |
| `scripts/gen-ribbon.mjs` | 목록 → 편집기 HTML 두 표시 사이. `--check`면 쓰지 않고 다르면 실패 | 2 |
| `src/client/ribbonRuntime.ts` | `hydrateIcons`(고른 탭만 아이콘), `fitPanel`·`collapseOrder`(접기), `installGroupToggles`·`closeGroupPopup`(접힌 묶음 펼침 판) | 2 |
| `src/client/officeShell.ts` | 렌더·창 크기 바뀔 때 위 함수 부르기, ui=view 흐림에서 준비 중·접힘 단추 빼기, `?rbctx=1`(상황별 탭 모두 보이기 — 사람 확인·e2e용) | 2 |
| `public/css/style.css` | 파일 끝 리본 꾸밈 뒤에 "오피스 리본 구조" 묶음(준비 중·빈 갤러리·확인란·입력칸·접힌 묶음·아이콘 크기), 스타일 갤러리 폭 상한 | 2 |
| `tests/ribbonInventory.test.ts` · `tests/ribbonGen.test.ts` · `tests/ribbonRuntime.test.ts` · `tests/ribbonPages.test.ts` · `tests/fixtures/ribbon-ids-before.json` | 목록 규칙, 생성기(꺼내기·순서·결정성·CRLF·오류), 접는 순서, 페이지 = 목록·기존 id 보존 | 1, 2(3~5에서 한 줄씩) |
| `public/editor.html` · `editor-grid.html` · `editor-slides.html` | 표시 넣기, 기존 단추에 `data-rb`, 크기 바꾸기·이름 고치기, 생성 | 3, 4, 5 |
| `src/client/editor.ts` · `grid-editor.ts` · `editor-slides.ts` | 상황별 탭 연결(워드 테이블 디자인·엑셀 차트 서식·PPT 도형 서식) — 각 1~4줄 | 3, 4, 5 |
| `public/ribbon/icons.svg` · `tests/ribbonIcons.test.ts` | 준비 중·대리·접힘 단추 아이콘 441개(단색 선, 24칸, 96 KB 상한) | 6 |
| `scripts/verify-ribbon-inventory-e2e.mjs` | DOM ↔ 목록 1:1 e2e + 폭별 접힘 + ui=view + M365 나란히 보고서 | 7 |
| `scripts/verify-readonly-editor-e2e.mjs` · `scripts/verify-ribbon-skin-e2e.mjs` | 단추 수 단언을 목록에서 계산, 포트 범위에 8180~8189 더함 | 7 |
| `package.json` | `gen:ribbon`·`verify:ribbon-inventory` 스크립트(아래 조건) | 7 |

**`package.json` 주의:** 계획 작성 때 다른 에이전트의 커밋 안 된 변경이 있었다. Task 7에서 `git diff --quiet package.json`이 깨끗할 때만 두 줄을 더해 커밋하고, 아니면 더하지 않는다(모든 명령은 `node scripts/…`로 직접 부른다).

## 목록 요약(계획 작성 시점, Task 1 대조 뒤 바뀔 수 있음 — 수는 `inventory.json`의 `summary`가 진짜)

| 앱 | 기본 탭(파일 포함) | 상황별 탭 | 묶음 | 단추 | 기존 | 대리(기존 단추를 다른 탭에서 대신 누름) | 준비 중 | 준비 중 "자세히" |
|---|---|---|---|---|---|---|---|---|
| 워드 | 10 | 5 (테이블 디자인·레이아웃·머리글/바닥글·그림 서식·도형 서식) | 84 | 324 | 64 | 1 | 259 | 16 |
| 엑셀 | 9 | 4 (차트 디자인·서식·테이블 디자인·그림 서식) | 75 | 298 | 63 | 0 | 235 | 14 |
| PPT | 11 | 8 (도형 서식·테이블 디자인·레이아웃·그림 서식·차트 디자인·서식·SmartArt 디자인·서식) | 96 | 333 | 33 | 3 | 297 | 16 |

기본 탭만(상황별 제외): 워드 기존 56(대리 포함)·준비 중 152, 엑셀 57·165, PPT 35·143. 전체 표는 **부록 A**.

**자리 결정 규칙(목록 `map`):** `same` = 오피스의 같은 기능·같은 자리 / `moved` = 다른 탭·묶음에서 오피스 자리로 옮김 / `nearest` = 오피스에 딱 맞는 것이 없어 가장 가까운 자리(이유는 `note`) / `proxy` = 오피스에서 두 군데 있는 기능의 두 번째 자리(기존 단추를 대신 누름). 기존 단추의 크기가 오피스와 다르면 `convert`(Task 3~5가 마크업을 바꿈), 이름·말풍선이 오피스와 다르면 `fix`, 그대로 둔 종류가 오피스와 다르면 `officeKind`.

**지워지는 것(기능 아님):** 오피스에 없는 "자세히" 두 개(워드 검토›언어 교정 → 단어 개수 대리, 엑셀 데이터›정렬 및 필터 → 사용자 지정 정렬 대리 — 같은 기능 단추가 같은 묶음에 있음), PPT 삽입 탭의 사각형·둥근 사각형·타원 대리 단추(같은 단추가 홈›그리기에 있고 오피스 삽입›도형 갤러리는 준비 중), 판 사이 HTML 주석. 모두 id가 없다.

---

### Task 1: 목록 검사와 M365 실물 대조

**Files:**
- Create: `scripts/ribbonGenLib.mjs` (Part A), `scripts/ribbonGenLib.d.mts`, `scripts/dump-m365-ribbon.ps1`, `scripts/compare-ribbon-inventory.mjs`, `tests/ribbonInventory.test.ts`
- Modify: `src/client/ribbon/inventory.json` (대조 결과로 고침 — 파일은 이미 있음)

**Interfaces:**
- Consumes: `src/client/ribbon/inventory.json`(이미 있음, 모양은 파일 안 `schema`)
- Produces (`scripts/ribbonGenLib.mjs`): `INVENTORY_PATH: string`, `APP_OF_EXT: {docx:'word',xlsx:'excel',pptx:'powerpoint'}`, `KINDS: string[]`, `loadInventory(file?): RibbonInventory`, `controlKey(tabId, groupId, c): string`(= `"<탭>.<묶음>.<키>"`), `tabCount(app): number`(파일 포함 기본 탭 수), `eachControl(app): Generator<{tab, group, c, key}>`, `pendingIn(app, tabId): number`(그 탭의 `[data-pending]` 요소 수 = 준비 중 단추 + 준비 중 자세히), `launcherCount(app): number`(상황별 포함 자세히 수), `summarize(app): Record<string, number>`, `validateInventory(inv): string[]`, `serializeInventory(inv): string`, `requiredIcons(inv): string[]`

- [ ] **Step 1: 실패하는 시험 쓰기** — `tests/ribbonInventory.test.ts`

```ts
import fs from 'node:fs';
import { describe, expect, it } from 'vitest';
import { INVENTORY_PATH, loadInventory, requiredIcons, serializeInventory, summarize, validateInventory } from '../scripts/ribbonGenLib.mjs';

describe('오피스 리본 목록(src/client/ribbon/inventory.json)', () => {
  const inv = loadInventory();
  it('규칙 위반 없음(겹침·한글 이름·종류·출처 하나·자리 결정 note·줄 순서·요약 수)', () => {
    expect(validateInventory(inv)).toEqual([]);
  });
  it('파일 모양이 serializeInventory 그대로(단추 한 줄에 하나)', () => {
    expect(fs.readFileSync(INVENTORY_PATH, 'utf8')).toBe(serializeInventory(inv));
  });
  it('기본 탭 순서는 설계 §2 그대로', () => {
    expect(inv.apps.word.tabs.map((t) => t.name)).toEqual(['홈', '삽입', '그리기', '디자인', '레이아웃', '참조', '편지', '검토', '보기']);
    expect(inv.apps.excel.tabs.map((t) => t.name)).toEqual(['홈', '삽입', '그리기', '페이지 레이아웃', '수식', '데이터', '검토', '보기']);
    expect(inv.apps.powerpoint.tabs.map((t) => t.name)).toEqual(['홈', '삽입', '그리기', '디자인', '전환', '애니메이션', '슬라이드 쇼', '녹음/녹화', '검토', '보기']);
  });
  it('기존 탭 id(편집기·시험이 쓰는 것)는 그대로', () => {
    const ids = (a: 'word' | 'excel' | 'powerpoint') => [...inv.apps[a].tabs, ...inv.apps[a].contextual].map((t) => t.id);
    expect(ids('word')).toEqual(expect.arrayContaining(['home', 'insert', 'layout', 'review', 'view', 'table']));
    expect(ids('excel')).toEqual(expect.arrayContaining(['home', 'insert', 'formulas', 'data', 'view', 'chart']));
    expect(ids('powerpoint')).toEqual(expect.arrayContaining(['home', 'insert', 'design', 'transitions', 'animations']));
  });
  it('기존 기능이 줄지 않는다(계획 작성 때 워드 64·엑셀 63·PPT 33, 대리 1·0·3)', () => {
    expect(summarize(inv.apps.word).existing).toBeGreaterThanOrEqual(64);
    expect(summarize(inv.apps.excel).existing).toBeGreaterThanOrEqual(63);
    expect(summarize(inv.apps.powerpoint).existing).toBeGreaterThanOrEqual(33);
    expect(summarize(inv.apps.powerpoint).proxy).toBeGreaterThanOrEqual(3);
  });
  it('아이콘 키는 영문 소문자·숫자·하이픈', () => {
    for (const k of requiredIcons(inv)) expect(k).toMatch(/^[a-z0-9-]+$/);
  });
});
```

- [ ] **Step 2: 실패 확인**

Run: `npx vitest run tests/ribbonInventory.test.ts`
Expected: FAIL — `Failed to resolve import "../scripts/ribbonGenLib.mjs"`

- [ ] **Step 3: `scripts/ribbonGenLib.mjs` Part A 쓰기**

```js
// 오피스 리본 목록(src/client/ribbon/inventory.json)과 리본 마크업 만들기·검사(설계 2026-11-22-office-ribbon-structure).
// scripts/gen-ribbon.mjs·compare-ribbon-inventory.mjs, tests/ribbon*.test.ts, scripts/verify-ribbon-inventory-e2e.mjs·
// verify-readonly-editor-e2e.mjs·verify-ribbon-skin-e2e.mjs가 함께 쓴다. 브라우저에서는 쓰지 않는다(node 전용).
import fs from 'node:fs';

export const INVENTORY_PATH = 'src/client/ribbon/inventory.json';
export const APP_OF_EXT = Object.freeze({ docx: 'word', xlsx: 'excel', pptx: 'powerpoint' });
export const KINDS = Object.freeze(['button', 'split', 'dropdown', 'gallery', 'toggle', 'input', 'select', 'info']);
const MAPS = ['same', 'moved', 'nearest', 'proxy'];
const LAYOUTS = ['columns', 'rows', 'fields'];
const ICON_KINDS = ['button', 'split', 'dropdown'];

export const loadInventory = (file = INVENTORY_PATH) => JSON.parse(fs.readFileSync(file, 'utf8'));
export const controlKey = (tabId, groupId, c) => `${tabId}.${groupId}.${c.key}`;
const allTabs = (app) => [...app.tabs, ...app.contextual];
/** 탭 줄의 기본 탭 수(파일 포함, 상황별 제외) */
export const tabCount = (app) => app.tabs.length + 1;
export function* eachControl(app) {
  for (const tab of allTabs(app)) for (const group of tab.groups) for (const c of group.controls) yield { tab, group, c, key: controlKey(tab.id, group.id, c) };
}
const latinOnly = (s) => /[A-Za-z]/.test(s) && !/[가-힣]/.test(s);

/** 탭 하나에서 [data-pending]이 붙는 요소 수 — 준비 중 단추 + 준비 중 "자세히" */
export function pendingIn(app, tabId) {
  const tab = allTabs(app).find((t) => t.id === tabId);
  if (!tab) throw new Error(`탭 없음 ${tabId}`);
  return tab.groups.reduce((n, g) => n + g.controls.filter((c) => c.pending).length + (g.launcher?.pending ? 1 : 0), 0);
}
/** 화면 전체(상황별 탭 포함)의 "자세히" 수 */
export const launcherCount = (app) => allTabs(app).reduce((n, t) => n + t.groups.filter((g) => g.launcher).length, 0);

export function summarize(app) {
  const s = { tabs: tabCount(app), contextualTabs: app.contextual.length, groups: 0, controls: 0, existing: 0, proxy: 0, pending: 0, launchersPending: 0, mainExisting: 0, mainPending: 0, ctxExisting: 0, ctxPending: 0 };
  for (const [list, which] of [[app.tabs, 'main'], [app.contextual, 'ctx']]) {
    for (const t of list) {
      for (const g of t.groups) {
        s.groups += 1;
        if (g.launcher?.pending) s.launchersPending += 1;
        for (const c of g.controls) {
          s.controls += 1;
          if (c.pending) {
            s.pending += 1;
            s[`${which}Pending`] += 1;
          } else {
            if (c.proxy) s.proxy += 1;
            else s.existing += 1;
            s[`${which}Existing`] += 1;
          }
        }
      }
    }
  }
  return s;
}

/** 목록 규칙 위반을 한 줄씩 — 빈 배열이면 통과 */
export function validateInventory(inv) {
  const bad = [];
  for (const [appName, app] of Object.entries(inv.apps)) {
    const say = (x) => bad.push(`${appName} ${x}`);
    const tabIds = new Set(['file']);
    const keys = new Set();
    for (const tab of allTabs(app)) {
      if (tabIds.has(tab.id)) say(`탭 id 겹침 ${tab.id}`);
      tabIds.add(tab.id);
      if (latinOnly(tab.name)) say(`탭 이름에 한글 없음 ${tab.name}`);
      const groupIds = new Set();
      for (const g of tab.groups) {
        const at = `${tab.id}.${g.id}`;
        if (groupIds.has(g.id)) say(`묶음 id 겹침 ${at}`);
        groupIds.add(g.id);
        if (latinOnly(g.name)) say(`묶음 이름에 한글 없음 ${at} ${g.name}`);
        if (!Number.isInteger(g.collapse) || g.collapse < 1 || g.collapse > 9) say(`collapse가 1~9 정수 아님 ${at}`);
        if (!LAYOUTS.includes(g.layout)) say(`layout ${g.layout} ${at}`);
        if (g.launcher !== null && typeof g.launcher?.pending !== 'string' && typeof g.launcher?.proxy !== 'string') say(`launcher 모양 ${at}`);
        if (!g.controls.length) say(`빈 묶음 ${at}`);
        let lastRow = -1;
        for (const c of g.controls) {
          const k = controlKey(tab.id, g.id, c);
          if (keys.has(k)) say(`key 겹침 ${k}`);
          keys.add(k);
          if (!/^[a-z0-9-]+$/.test(c.key)) say(`key 모양 ${k}`);
          if (!c.name || latinOnly(c.name)) say(`이름 ${k} "${c.name}"`);
          if (!['big', 'small'].includes(c.size)) say(`size ${k}`);
          if (!KINDS.includes(c.kind)) say(`kind ${k}`);
          if (c.kind === 'gallery' && c.size !== 'big') say(`갤러리는 big ${k}`);
          if ([c.pending === true, typeof c.ref === 'string', typeof c.proxy === 'string'].filter(Boolean).length !== 1) say(`pending·ref·proxy 중 꼭 하나 ${k}`);
          if (c.pending && c.map !== undefined) say(`준비 중인데 map ${k}`);
          if (!c.pending && !MAPS.includes(c.map)) say(`map ${k}`);
          if ((c.map === 'nearest' || c.map === 'moved') && !c.note) say(`nearest·moved는 note 필수 ${k}`);
          if (c.row !== undefined && g.layout !== 'rows') say(`row는 rows 묶음에만 ${k}`);
          if (c.row !== undefined) {
            if (c.row < lastRow) say(`row 순서(줄 번호가 줄어듦) ${k}`);
            lastRow = c.row;
          }
          if (c.stackTitle !== undefined && g.layout !== 'fields') say(`stackTitle은 fields 묶음에만 ${k}`);
          if (c.convert !== undefined && !['big→small', 'small→big'].includes(c.convert)) say(`convert ${k}`);
        }
      }
    }
    if (JSON.stringify(inv.summary?.[appName]) !== JSON.stringify(summarize(app))) say(`summary가 실제 수와 다름 — 맞는 값 ${JSON.stringify(summarize(app))}`);
  }
  return bad;
}

/** 목록 파일 모양(단추 한 줄에 하나) — 목록을 고친 뒤 이것으로 다시 쓴다 */
export function serializeInventory(inv) {
  const ind = (n) => '  '.repeat(n);
  const ser = (v, d) => {
    if (Array.isArray(v)) {
      if (v.length && v.every((x) => x && typeof x === 'object' && 'key' in x && 'name' in x)) return `[\n${v.map((c) => ind(d + 1) + JSON.stringify(c)).join(',\n')}\n${ind(d)}]`;
      if (v.every((x) => typeof x !== 'object' || x === null)) return JSON.stringify(v);
      return `[\n${v.map((x) => ind(d + 1) + ser(x, d + 1)).join(',\n')}\n${ind(d)}]`;
    }
    if (v && typeof v === 'object') {
      const keys = Object.keys(v);
      if (keys.every((k) => typeof v[k] !== 'object' || v[k] === null)) {
        const flat = JSON.stringify(v);
        if (flat.length < 140) return flat;
      }
      return `{\n${keys.map((k) => `${ind(d + 1)}${JSON.stringify(k)}: ${ser(v[k], d + 1)}`).join(',\n')}\n${ind(d)}}`;
    }
    return JSON.stringify(v);
  };
  return `${ser(inv, 0)}\n`;
}

/** 준비 중·대리 단추가 쓰는 아이콘 키 + 접힌 묶음 단추(묶음 첫 단추의 아이콘) — public/ribbon/icons.svg 의 rbi-<키> */
export function requiredIcons(inv) {
  const keys = new Set();
  for (const app of Object.values(inv.apps)) {
    for (const tab of allTabs(app)) {
      for (const g of tab.groups) {
        keys.add(g.controls[0].icon ?? g.controls[0].key);
        for (const c of g.controls) if ((c.pending || c.proxy) && ICON_KINDS.includes(c.kind)) keys.add(c.icon ?? c.key);
      }
    }
  }
  return [...keys].sort();
}
```

`scripts/ribbonGenLib.d.mts`(시험이 TypeScript라 `tsc --noEmit`이 .mjs 타입을 찾는다 — Task 2가 아래 `// Part B` 줄들을 더한다):

```ts
// scripts/ribbonGenLib.mjs 타입 — 함수가 바뀌면 같이 고친다
export interface RibbonControl {
  key: string; name: string; size: 'big' | 'small';
  kind: 'button' | 'split' | 'dropdown' | 'gallery' | 'toggle' | 'input' | 'select' | 'info';
  pending?: true; ref?: string; unit?: string; proxy?: string; map?: 'same' | 'moved' | 'nearest' | 'proxy';
  officeKind?: string; convert?: 'big→small' | 'small→big'; fix?: string; note?: string;
  row?: number; iconOnly?: true; break?: true; stackTitle?: string; icon?: string; items?: string[];
}
export interface RibbonGroup { id: string; name: string; collapse: number; layout: 'columns' | 'rows' | 'fields'; launcher: null | { pending: string } | { proxy: string }; domId?: string; className?: string; note?: string; controls: RibbonControl[] }
export interface RibbonTab { id: string; name: string; editOnly: boolean; panelId?: string; note?: string; trigger?: string; wired?: string | null; groups: RibbonGroup[] }
export interface RibbonApp { page: string; backstage: { office: string[]; existing: string[] }; tabs: RibbonTab[]; contextual: RibbonTab[] }
export type AppName = 'word' | 'excel' | 'powerpoint';
export interface RibbonInventory { version: number; spec: string; source: string; schema: Record<string, string>; summary: Record<AppName, Record<string, number>>; apps: Record<AppName, RibbonApp> }
export declare const INVENTORY_PATH: string;
export declare const APP_OF_EXT: Readonly<Record<'docx' | 'xlsx' | 'pptx', AppName>>;
export declare const KINDS: readonly string[];
export declare function loadInventory(file?: string): RibbonInventory;
export declare function controlKey(tabId: string, groupId: string, c: RibbonControl): string;
export declare function tabCount(app: RibbonApp): number;
export declare function eachControl(app: RibbonApp): Generator<{ tab: RibbonTab; group: RibbonGroup; c: RibbonControl; key: string }>;
export declare function pendingIn(app: RibbonApp, tabId: string): number;
export declare function launcherCount(app: RibbonApp): number;
export declare function summarize(app: RibbonApp): Record<string, number>;
export declare function validateInventory(inv: RibbonInventory): string[];
export declare function serializeInventory(inv: RibbonInventory): string;
export declare function requiredIcons(inv: RibbonInventory): string[];
```

- [ ] **Step 4: 통과 확인**

Run: `npx vitest run tests/ribbonInventory.test.ts && npx tsc --noEmit`
Expected: PASS 6개, 타입 오류 없음(계획 작성 때 목록은 이미 이 규칙을 통과했다).

- [ ] **Step 5: M365 덤프 스크립트 쓰기** — `scripts/dump-m365-ribbon.ps1` (**UTF-8 BOM 있게 저장** — Windows PowerShell 5.1은 BOM 없는 파일을 CP949로 읽어 `'리본'` 같은 한글 비교가 깨진다)

```powershell
# 이 PC의 한국어 Microsoft 365 리본을 UI 자동화로 읽어 탭·묶음·단추 이름·높이를 JSON으로, 탭마다 리본 그림을 PNG로 남긴다
# (설계 2026-11-22-office-ribbon-structure, 계획 Task 1). 목록 대조용 — 결과는 .audit-run/ 아래(커밋하지 않음).
# 쓰는 법: powershell -ExecutionPolicy Bypass -File scripts/dump-m365-ribbon.ps1 -App word|excel|powerpoint [-Out .audit-run/m365-ribbon]
# 같은 앱이 이미 떠 있으면 멈춘다(사용자 문서를 닫지 않게). 띄운 Office는 이 스크립트가 받은 프로세스 번호로만 끈다.
param(
  [Parameter(Mandatory = $true)][ValidateSet('word', 'excel', 'powerpoint')][string]$App,
  [string]$Out = '.audit-run/m365-ribbon'
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Drawing
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class RibbonWin {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
}
'@
$procName = @{ word = 'WINWORD'; excel = 'EXCEL'; powerpoint = 'POWERPNT' }[$App]
if (Get-Process -Name $procName -ErrorAction SilentlyContinue) { throw "$procName 이(가) 이미 실행 중 — 사용자 창을 건드리지 않으려고 멈춤. 닫은 뒤 다시 실행." }
New-Item -ItemType Directory -Force $Out | Out-Null
$Out = (Resolve-Path $Out).Path
$A = [System.Windows.Automation.AutomationElement]
$CT = [System.Windows.Automation.ControlType]
$Scope = [System.Windows.Automation.TreeScope]
$walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
$png = Join-Path $Out 'probe.png'
$b = New-Object System.Drawing.Bitmap 40, 40; $b.Save($png, [System.Drawing.Imaging.ImageFormat]::Png); $b.Dispose()

$procId = [uint32]0
try {
  switch ($App) {
    'word' { $office = New-Object -ComObject Word.Application; $office.Visible = $true; $doc = $office.Documents.Add(); $hwnd = [IntPtr]$office.ActiveWindow.Hwnd }
    'excel' { $office = New-Object -ComObject Excel.Application; $office.Visible = $true; $doc = $office.Workbooks.Add(); $hwnd = [IntPtr]$office.Hwnd }
    'powerpoint' { $office = New-Object -ComObject PowerPoint.Application; $office.Visible = -1; $doc = $office.Presentations.Add(); $null = $doc.Slides.Add(1, 12); $hwnd = [IntPtr]$office.HWND }
  }
  [void][RibbonWin]::GetWindowThreadProcessId($hwnd, [ref]$procId)
  [void][RibbonWin]::ShowWindow($hwnd, 3) # 최대화 — 묶음이 접히지 않게
  [void][RibbonWin]::SetForegroundWindow($hwnd)
  Start-Sleep -Seconds 3
  $root = $A::FromHandle($hwnd)

  function Get-RibbonTabs {
    $cond = New-Object System.Windows.Automation.PropertyCondition($A::ControlTypeProperty, $CT::TabItem)
    $tabs = @($root.FindAll($Scope::Descendants, $cond)) | Where-Object { $walker.GetParent($_).Current.Name -match '리본' }
    if (-not $tabs) {
      $all = @($root.FindAll($Scope::Descendants, $cond)) | ForEach-Object { "$($_.Current.Name) <- $($walker.GetParent($_).Current.Name)" }
      $all | Set-Content (Join-Path $Out "$App-tabitems-debug.txt") -Encoding utf8
      throw "리본 탭을 못 찾음 — $App-tabitems-debug.txt 의 부모 이름을 보고 Get-RibbonTabs 의 '리본' 조건을 고칠 것"
    }
    $tabs
  }
  $Leaf = @('Button', 'SplitButton', 'MenuItem', 'ComboBox', 'CheckBox', 'Edit', 'List', 'Spinner', 'RadioButton', 'Hyperlink')
  function Get-Controls($el) {
    $out = @()
    $child = $walker.GetFirstChild($el)
    while ($child) {
      $type = $child.Current.ControlType.ProgrammaticName -replace '^ControlType\.', ''
      if ($Leaf -contains $type) {
        $r = $child.Current.BoundingRectangle
        if (-not $child.Current.IsOffscreen) { $out += [ordered]@{ name = $child.Current.Name; type = $type; height = [int]$r.Height; width = [int]$r.Width } }
      } else { $out += @(Get-Controls $child) }
      $child = $walker.GetNextSibling($child)
    }
    $out
  }
  function Get-RibbonGroups($tabBottom) {
    $cond = New-Object System.Windows.Automation.PropertyCondition($A::ControlTypeProperty, $CT::Group)
    $all = @($root.FindAll($Scope::Descendants, $cond) | Where-Object { $r = $_.Current.BoundingRectangle; -not $_.Current.IsOffscreen -and $r.Top -ge ($tabBottom - 2) -and $r.Bottom -le ($tabBottom + 200) -and $r.Width -gt 20 })
    $outer = $all | Where-Object {
      $p = $walker.GetParent($_); $nested = $false
      while ($p -and -not $nested) { if ($all -contains $p) { $nested = $true }; $p = $walker.GetParent($p) }
      -not $nested
    }
    $outer | Sort-Object { $_.Current.BoundingRectangle.Left }
  }
  function Save-RibbonShot($name, $top, $bottom) {
    $w = $root.Current.BoundingRectangle
    $rect = New-Object System.Drawing.Rectangle ([int]$w.Left), ([int]$top), ([int]$w.Width), ([int]($bottom - $top))
    $bmp = New-Object System.Drawing.Bitmap $rect.Width, $rect.Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.CopyFromScreen($rect.Location, [System.Drawing.Point]::Empty, $rect.Size)
    $bmp.Save((Join-Path $Out ("$App-$name.png" -replace '[\\/:*?"<>|]', '_')), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
  }
  function Dump-Tabs($only) {
    $result = @()
    foreach ($tab in Get-RibbonTabs) {
      $name = $tab.Current.Name
      if ($name -eq '파일' -or ($only -and $only -notcontains $name)) { continue }
      $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
      Start-Sleep -Milliseconds 900
      $tabBottom = $tab.Current.BoundingRectangle.Bottom
      $groups = @(); $bandBottom = $tabBottom
      foreach ($g in Get-RibbonGroups $tabBottom) {
        $groups += [ordered]@{ name = $g.Current.Name; controls = @(Get-Controls $g) }
        $bandBottom = [Math]::Max($bandBottom, $g.Current.BoundingRectangle.Bottom)
      }
      Save-RibbonShot $name ($tab.Current.BoundingRectangle.Top - 2) ($bandBottom + 2)
      $result += [ordered]@{ name = $name; groups = $groups }
    }
    $result
  }

  $base = @(Get-RibbonTabs | ForEach-Object { $_.Current.Name })
  $main = Dump-Tabs $null
  # 상황별 탭 — 개체를 넣고 골라서 새로 생긴 탭만 읽는다
  $setups = @{
    word = @(
      @{ name = '표'; run = { $t = $doc.Tables.Add($doc.Range(0, 0), 2, 2); $t.Cell(1, 1).Select() } },
      @{ name = '도형'; run = { $doc.Shapes.AddShape(1, 100, 100, 80, 60).Select() } },
      @{ name = '그림'; run = { $doc.InlineShapes.AddPicture($png).Select() } },
      @{ name = '머리글'; run = { $office.ActiveWindow.ActivePane.View.SeekView = 9 } }
    )
    excel = @(
      @{ name = '차트'; run = { $ws = $doc.ActiveSheet; $ws.Range('A1:B3').Value2 = 1; $ws.Range('A1:B3').Select(); $ws.Shapes.AddChart2(201, 51).Select() } },
      @{ name = '표'; run = { $ws = $doc.ActiveSheet; $ws.Range('D1:E3').Value2 = 2; $null = $ws.ListObjects.Add(1, $ws.Range('D1:E3'), $null, 1); $ws.Range('D2').Select() } },
      @{ name = '그림'; run = { $doc.ActiveSheet.Shapes.AddPicture($png, 0, -1, 300, 300, 60, 60).Select() } }
    )
    powerpoint = @(
      @{ name = '도형'; run = { $doc.Slides(1).Shapes.AddShape(1, 100, 100, 80, 60).Select() } },
      @{ name = '표'; run = { $doc.Slides(1).Shapes.AddTable(2, 2, 300, 100, 200, 80).Select() } },
      @{ name = '그림'; run = { $doc.Slides(1).Shapes.AddPicture($png, 0, -1, 100, 300, 60, 60).Select() } },
      @{ name = '차트'; run = { $s = $doc.Slides(1).Shapes.AddChart2(201, 51); try { $s.Chart.ChartData.Workbook.Close() } catch {}; $s.Select() } },
      @{ name = 'SmartArt'; run = { $doc.Slides(1).Shapes.AddSmartArt($office.SmartArtLayouts.Item(1)).Select() } }
    )
  }
  $contexts = @()
  foreach ($s in $setups[$App]) {
    try { & $s.run } catch { Write-Warning "상황 '$($s.name)' 준비 실패: $_"; continue }
    [void][RibbonWin]::SetForegroundWindow($hwnd); Start-Sleep -Seconds 2
    $new = @(Get-RibbonTabs | ForEach-Object { $_.Current.Name } | Where-Object { $base -notcontains $_ })
    $contexts += [ordered]@{ trigger = $s.name; tabs = @(Dump-Tabs $new) }
    if ($App -eq 'word' -and $s.name -eq '머리글') { $office.ActiveWindow.ActivePane.View.SeekView = 0 }
  }
  [ordered]@{ app = $App; version = "$($office.Version)"; tabs = $main; contextual = $contexts } | ConvertTo-Json -Depth 10 | Set-Content (Join-Path $Out "$App.json") -Encoding utf8
  Write-Output "덤프 $Out\$App.json, 캡처 $Out\$App-*.png"
}
finally {
  try { if ($App -eq 'word') { $doc.Close(0) } elseif ($App -eq 'excel') { $doc.Close($false) } else { $doc.Saved = -1; $doc.Close() } } catch {}
  try { $office.Quit() } catch {}
  Start-Sleep -Seconds 2
  if ($procId -and (Get-Process -Id $procId -ErrorAction SilentlyContinue)) { taskkill /PID $procId /T /F | Out-Null }
}
```

- [ ] **Step 6: 대조 스크립트 쓰기** — `scripts/compare-ribbon-inventory.mjs`

```js
#!/usr/bin/env node
// M365 UIA 덤프(scripts/dump-m365-ribbon.ps1)와 목록(src/client/ribbon/inventory.json)을 탭·묶음·단추 이름·크기로 맞대 본다(계획 Task 1)
// 쓰는 법: node scripts/compare-ribbon-inventory.mjs word|excel|powerpoint [.audit-run/m365-ribbon/<앱>.json]
import fs from 'node:fs';
import { loadInventory } from './ribbonGenLib.mjs';

const [appName, dumpPath = `.audit-run/m365-ribbon/${appName}.json`] = process.argv.slice(2);
const app = loadInventory().apps[appName];
if (!app) throw new Error('앱 이름은 word|excel|powerpoint');
const dump = JSON.parse(fs.readFileSync(dumpPath, 'utf8').replace(/^﻿/, ''));
// 영문만 오피스 이름은 목록에서 한글 표기(사용자 규칙 4) — 같은 것으로 본다
const ALIAS = { SmartArt: '스마트아트', WordArt: '워드아트', SharePoint: '셰어포인트', 'WordArt 스타일': '워드아트 스타일', 'SmartArt로 변환': '스마트아트로 변환', Designer: '디자이너' };
const norm = (s) => {
  const t = String(s ?? '').replace(/[.…]+$/, '').replace(/\s+/g, ' ').trim();
  return ALIAS[t] ?? t;
};
const asList = (x) => (Array.isArray(x) ? x : x ? [x] : []); // ConvertTo-Json은 한 개짜리 배열을 객체로 쓴다
let diffs = 0;
const say = (s) => {
  diffs += 1;
  console.log(s);
};
function compareTabs(ours, theirs, label) {
  const tn = theirs.map((t) => norm(t.name));
  const on = ours.map((t) => norm(t.name));
  if (tn.join('|') !== on.join('|')) say(`[${label} 탭] 오피스 ${tn.join(' · ')}\n           목록 ${on.join(' · ')}`);
  for (const t of theirs) {
    const o = ours.find((x) => norm(x.name) === norm(t.name));
    if (!o) continue;
    const tg = asList(t.groups).map((g) => norm(g.name));
    const og = o.groups.map((g) => norm(g.name));
    if (tg.join('|') !== og.join('|')) say(`[${o.name} 묶음] 오피스 ${tg.join(' · ')}\n           목록 ${og.join(' · ')}`);
    for (const g of asList(t.groups)) {
      const mine = o.groups.find((x) => norm(x.name) === norm(g.name));
      if (!mine) continue;
      const tc = asList(g.controls).map((c) => norm(c.name)).filter(Boolean);
      const oc = mine.controls.map((c) => norm(c.name));
      const onlyOffice = tc.filter((n) => !oc.includes(n));
      const onlyOurs = oc.filter((n) => !tc.includes(n));
      if (onlyOffice.length || onlyOurs.length) say(`[${o.name} > ${mine.name}] 오피스에만: ${onlyOffice.join(', ') || '-'} | 목록에만: ${onlyOurs.join(', ') || '-'}`);
      const order = oc.filter((n) => tc.includes(n));
      const theirOrder = tc.filter((n) => oc.includes(n));
      if (order.join('|') !== theirOrder.join('|')) say(`[${o.name} > ${mine.name}] 순서 오피스 ${theirOrder.join(' · ')}\n           목록 ${order.join(' · ')}`);
      for (const c of asList(g.controls)) {
        const m = mine.controls.find((x) => norm(x.name) === norm(c.name));
        // 오피스 단추 높이 50px 넘으면 큰 단추(100~150% 배율에서 큰 66~99px, 작은 22~33px)
        if (m && ['Button', 'SplitButton', 'MenuItem'].includes(c.type)) {
          const size = c.height > 50 ? 'big' : 'small';
          if (size !== m.size) say(`[${o.name} > ${mine.name} > ${m.name}] 크기 오피스 ${size} · 목록 ${m.size} (${c.height}px)`);
        }
      }
    }
  }
}
compareTabs(app.tabs, asList(dump.tabs), '기본');
compareTabs(app.contextual, asList(dump.contextual).flatMap((c) => asList(c.tabs)), '상황별');
console.log(diffs ? `\n차이 ${diffs}건 — 목록을 고친 뒤 다시` : '\n차이 없음');
process.exitCode = diffs ? 1 : 0;
```

- [ ] **Step 7: 덤프 받기(앱마다, 사용자의 Office 창이 없을 때)**

Run(PowerShell, 각 1분 안팎 — Office 창이 잠깐 떴다 닫힌다):
```powershell
foreach ($a in 'word','excel','powerpoint') { powershell -ExecutionPolicy Bypass -File scripts/dump-m365-ribbon.ps1 -App $a }
```
Expected: `.audit-run/m365-ribbon/word.json`·`excel.json`·`powerpoint.json`과 탭마다 PNG. "이미 실행 중"으로 멈추면 사용자에게 해당 Office를 닫아 달라고 하고 기다린다(직접 끄지 않는다). "리본 탭을 못 찾음"이면 debug 파일의 부모 이름으로 `Get-RibbonTabs` 조건을 고쳐 다시. 모니터가 좁아 오피스 묶음 자체가 접혀 덤프에 묶음 이름 하나짜리 단추로 나오면, 그 묶음은 차이 보고에서 빼고(목록 그대로) 보고서에 적는다.

- [ ] **Step 8: 대조하고 목록 고치기**

Run: `node scripts/compare-ribbon-inventory.mjs word` (excel·powerpoint도)
고치는 규칙(오피스가 기준, 탭 목록만은 설계 §2가 기준):
1. 탭: 설계 §2 목록 그대로 둔다(오피스의 "도움말"·추가 기능 탭은 넣지 않는다). 상황별 탭 이름은 오피스대로(예: "테이블 디자인" — 설계 §2의 "표 디자인"과 다르면 오피스대로, 묶음 `note`가 아니라 보고서에 적는다).
2. 묶음 이름·순서·단추 이름·순서·크기가 다르면 오피스대로 고친다. 기존 기능 단추의 이름을 바꾸면 그 단추에 `fix`("이름 "옛" → "새"")를 적는다(Task 3~5가 마크업을 고친다). `key`는 바꾸지 않는다.
3. 오피스에만 있는 단추는 그 자리에 준비 중으로 넣는다: UIA `type` → `kind`: Button → button(펼침 화살표가 보이면 dropdown), SplitButton → split, MenuItem → dropdown, ComboBox → select, CheckBox → toggle, Edit·Spinner → input, List → gallery. 크기는 높이 50px 기준.
4. 목록에만 있는 준비 중 단추는 지운다. 목록에만 있는 **기존** 단추는 지우지 않고 `map: "nearest"` + `note`(오피스에 없는 이유)로 둔다.
5. 새 키는 묶음 안에서 유일한 영문 소문자(예: `"reuse-slides"`).
6. 다 고친 뒤 요약과 파일 모양을 다시 쓴다:
```bash
node --input-type=module -e "import fs from 'node:fs'; import { INVENTORY_PATH, loadInventory, serializeInventory, summarize } from './scripts/ribbonGenLib.mjs'; const inv = loadInventory(); for (const [n, a] of Object.entries(inv.apps)) inv.summary[n] = summarize(a); fs.writeFileSync(INVENTORY_PATH, serializeInventory(inv));"
```
Expected: 세 앱 모두 `차이 없음`(또는 남는 차이가 2번의 "기존 단추 nearest"·1번의 탭 규칙·덤프 접힘뿐이고 보고서에 적힘).

덤프를 못 받는 경우(스크립트가 UI 자동화로 탭을 끝내 못 찾음): 목록은 그대로 두고, Task 7의 사람 확인에서 M365 화면과 나란히 보며 고친다 — 그 사실을 보고서에 적는다.

- [ ] **Step 9: 시험 다시**

Run: `npx vitest run tests/ribbonInventory.test.ts && npx tsc --noEmit`
Expected: PASS

- [ ] **Step 10: 커밋**

```bash
git add scripts/ribbonGenLib.mjs scripts/ribbonGenLib.d.mts scripts/dump-m365-ribbon.ps1 scripts/compare-ribbon-inventory.mjs tests/ribbonInventory.test.ts src/client/ribbon/inventory.json
git diff --cached --stat
git commit -m "오피스 리본 목록 검사와 M365 실물 대조: 목록 규칙 시험(겹침·한글·출처·자리 결정 note·요약 수), UIA 덤프·대조 스크립트, 이 PC 한국어 M365로 묶음·단추 이름·순서·크기 맞춤

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 2: 마크업 생성기·접기·아이콘 달기(렌더러 기반)

**Files:**
- Modify: `scripts/ribbonGenLib.mjs` (Part B를 파일 끝에 더함), `scripts/ribbonGenLib.d.mts` (Part B 타입), `src/client/officeShell.ts`, `public/css/style.css`
- Create: `scripts/gen-ribbon.mjs`, `src/client/ribbonRuntime.ts`, `tests/ribbonGen.test.ts`, `tests/ribbonRuntime.test.ts`, `tests/ribbonPages.test.ts`, `tests/fixtures/ribbon-ids-before.json`

**Interfaces:**
- Consumes: Task 1의 `loadInventory`, `controlKey`, `eachControl`, `RibbonApp`·`RibbonControl` 타입
- Produces:
  - `scripts/ribbonGenLib.mjs`: `extractUnits(html): Map<string, string>`, `pendingHtml(key, c): string`, `proxyHtml(key, c): string`, `renderPanels(app, units, pad?): string`, `renderTabs(app, pad?): string`, `generatePage(source, app): string`(표시가 없거나 기존 단추의 `data-rb`가 없으면 throw), `expectedStructure(app): {tabs, panels}`, `refsOf(app): Record<key, selector>`
  - 마크업 약속(Task 3~7이 기댐): 탭 `button.ribbon-tab[data-tab]`(상황별은 `.ribbon-tab-contextual[hidden]`), 판 `.ribbon-panel[data-panel]`, 묶음 `.ribbon-group[data-group][data-collapse][role=group][aria-label]` > `button.rb-group-toggle[data-icon]` + `.ribbon-group-buttons` + `.ribbon-group-label` + (선택) `.ribbon-launcher`, 단추마다 바깥 요소에 첫 속성 `data-rb`, 준비 중은 `data-pending`·`data-kind`·`title="준비 중 — 이름"`·disabled, 대리는 `data-kind`·`data-proxy`, 아이콘이 필요한 것은 `data-icon`
  - `src/client/ribbonRuntime.ts`: `ICON_SPRITE = '/ribbon/icons.svg'`, `NARROW_PX = 600`, `collapseOrder(groups: {priority: number; index: number}[]): number[]`, `fitPanel(panel: HTMLElement, viewportWidth: number): void`, `hydrateIcons(panel: HTMLElement): void`, `installGroupToggles(toolbar: HTMLElement): void`, `closeGroupPopup(): void`
  - CLI: `node scripts/gen-ribbon.mjs [--check] [word|excel|powerpoint …]`(목록 검사 실패·표시 없음·다름이면 종료 코드 1)
  - URL 매개변수 `rbctx=1`: 상황별 탭을 모두 보이고 편집기가 숨기려 해도 그대로(사람 확인·e2e)

- [ ] **Step 1: 바꾸기 전 id 목록 고정**(HTML을 아직 안 바꾼 지금 만든다)

```bash
node -e "const fs=require('fs');const out={};for(const p of ['public/editor.html','public/editor-grid.html','public/editor-slides.html']){const h=fs.readFileSync(p,'utf8');const s=h.slice(h.indexOf('<header'),h.indexOf('</header>'));out[p]=[...new Set([...s.matchAll(/\sid=\"([^\"]+)\"/g)].map(m=>m[1]))].sort();}fs.writeFileSync('tests/fixtures/ribbon-ids-before.json',JSON.stringify(out,null,1)+'\n')"
```
Expected: 계획 작성 때 기준 워드 82·엑셀 69·PPT 46개 id(다른 에이전트 커밋으로 늘었으면 그 수).

- [ ] **Step 2: 실패하는 시험 쓰기** — `tests/ribbonGen.test.ts`, `tests/ribbonRuntime.test.ts`, `tests/ribbonPages.test.ts`

`tests/ribbonGen.test.ts`:
```ts
import { describe, expect, it } from 'vitest';
import { extractUnits, generatePage, pendingHtml, type RibbonApp } from '../scripts/ribbonGenLib.mjs';

const app: RibbonApp = {
  page: 'x.html',
  backstage: { office: [], existing: [] },
  tabs: [
    {
      id: 'home', name: '홈', editOnly: true, groups: [
        { id: 'clipboard', name: '클립보드', collapse: 9, layout: 'columns', launcher: { pending: '클립보드 창' }, controls: [
          { key: 'paste', name: '붙여넣기', size: 'big', kind: 'button', ref: '#btn-paste', map: 'same' },
          { key: 'cut', name: '잘라내기', size: 'small', kind: 'button', pending: true },
          { key: 'copy', name: '복사', size: 'small', kind: 'split', pending: true },
        ] },
        { id: 'font', name: '글꼴', collapse: 8, layout: 'rows', launcher: null, controls: [
          { key: 'bold', name: '굵게', size: 'small', kind: 'button', ref: '#btn-bold', map: 'same', row: 0, iconOnly: true },
          { key: 'shadow', name: '텍스트 그림자', size: 'small', kind: 'button', pending: true, row: 1, iconOnly: true },
        ] },
      ],
    },
  ],
  contextual: [
    { id: 'table', name: '레이아웃', editOnly: true, trigger: '표 안', wired: null, groups: [
      { id: 'cell-size', name: '셀 크기', collapse: 3, layout: 'fields', launcher: null, controls: [{ key: 'width', name: '너비', size: 'small', kind: 'input', pending: true, stackTitle: '크기' }] },
    ] },
  ],
};
const page = (panels: string) =>
  `<header>\n    <div class="ribbon-tabs">\n      <!-- ribbon:tabs:begin -->\n      <button>옛 탭</button>\n      <!-- ribbon:tabs:end -->\n      <span class="ribbon-tabs-spacer"></span>\n    </div>\n    <!-- ribbon:panels:begin -->\n${panels}\n    <!-- ribbon:panels:end -->\n  </header>\n`;
const OLD = page(`    <div class="ribbon-panel" data-panel="home">
      <button data-rb="home.clipboard.paste" id="btn-paste" class="icon-btn icon-btn-big" title="붙여넣기 (Ctrl+V)">
        <svg viewBox="0 0 24 24"><path d="M1 1"/></svg>
        <span>붙여넣기</span>
      </button>
      <span><button data-rb="home.font.bold" id="btn-bold" class="icon-btn" title="굵게 (Ctrl+Shift+>)"><strong>가</strong></button></span>
      <!-- 옛 주석 -->
      <button type="button" class="ribbon-launcher" data-proxy="btn-x">옛 자세히</button>
    </div>`);

describe('extractUnits', () => {
  it('data-rb 요소를 닫는 태그까지(속성 안의 > 포함) 꺼낸다', () => {
    const units = extractUnits(OLD);
    expect([...units.keys()]).toEqual(['home.clipboard.paste', 'home.font.bold']);
    expect(units.get('home.font.bold')).toBe('<button data-rb="home.font.bold" id="btn-bold" class="icon-btn" title="굵게 (Ctrl+Shift+>)"><strong>가</strong></button>');
    expect(units.get('home.clipboard.paste')).toMatch(/^<button data-rb="home\.clipboard\.paste"[^]*<\/button>$/);
  });
  it('안쪽에 같은 이름 태그가 겹쳐도 짝을 맞춘다', () => {
    const units = extractUnits('<span data-rb="a.b.c" class="dropdown-btn"><button>x</button><span>y<span>z</span></span></span><span>밖</span>');
    expect(units.get('a.b.c')).toBe('<span data-rb="a.b.c" class="dropdown-btn"><button>x</button><span>y<span>z</span></span></span>');
  });
  it('data-rb가 첫 속성이 아니면 멈춘다', () => {
    expect(() => extractUnits('<button id="x" data-rb="a.b.c">x</button>')).toThrow(/첫 속성/);
  });
  it('같은 data-rb가 두 번이면 멈춘다', () => {
    expect(() => extractUnits('<i data-rb="a.b.c"></i><i data-rb="a.b.c"></i>')).toThrow(/두 번/);
  });
});

describe('generatePage', () => {
  it('목록 순서대로 탭·판·묶음·단추를 만들고 기존 단추는 마크업 그대로 옮긴다', () => {
    const out = generatePage(OLD, app);
    expect(out).toContain('<button class="ribbon-tab ribbon-tab-file" role="tab" data-tab="file">파일</button>');
    expect(out).toContain('<button class="ribbon-tab" role="tab" data-tab="home" data-edit-only>홈</button>');
    expect(out).toContain('<button class="ribbon-tab ribbon-tab-contextual" role="tab" data-tab="table" data-edit-only hidden>레이아웃</button>');
    expect(out).not.toMatch(/옛 탭|옛 주석|옛 자세히/);
    expect(out.indexOf('data-rb="home.clipboard.paste"')).toBeLessThan(out.indexOf('data-rb="home.clipboard.cut"'));
    expect(out).toContain('title="굵게 (Ctrl+Shift+>)"');
    expect(out).toContain('<button data-rb="home.clipboard.cut" data-pending data-kind="button" title="준비 중 — 잘라내기" type="button" class="icon-btn ribbon-btn-labeled rb-pending" data-icon="cut" aria-label="잘라내기" disabled><span class="ribbon-btn-label">잘라내기</span></button>');
    expect(out).toContain('<button type="button" class="ribbon-launcher" data-pending disabled title="준비 중 — 클립보드 창" aria-label="클립보드 자세히">');
    expect(out).toContain('<div class="ribbon-group" data-group="font" data-collapse="8" role="group" aria-label="글꼴">');
    expect(out).toContain('class="icon-btn icon-btn-big rb-group-toggle" data-icon="paste" aria-expanded="false" title="클립보드"');
    expect(out).toContain('<div class="ribbon-group-buttons-wrap rb-rows">');
    expect(out).toContain('<span class="layout-field-title">크기</span>');
    expect(out).toMatch(/data-panel="home" role="tabpanel">/);
    expect(out).toMatch(/data-panel="table" role="tabpanel" hidden>/);
  });
  it('두 번 돌려도 같다(결정적)', () => {
    const once = generatePage(OLD, app);
    expect(generatePage(once, app)).toBe(once);
  });
  it('CRLF 파일은 CRLF로 둔다', () => {
    const once = generatePage(OLD.replace(/\n/g, '\r\n'), app);
    expect(once.includes('\r\n')).toBe(true);
    expect(/[^\r]\n/.test(once)).toBe(false);
  });
  it('기존 단추의 data-rb가 없으면·목록에 없는 data-rb가 있으면 멈춘다', () => {
    expect(() => generatePage(OLD.replace('data-rb="home.font.bold" ', ''), app)).toThrow(/home\.font\.bold/);
    expect(() => generatePage(OLD.replace('home.font.bold', 'home.font.zzz'), app)).toThrow(/목록에 없는 data-rb/);
  });
  it('기존 기능 자리에 준비 중 마크업이 있으면 멈춘다', () => {
    const fake = OLD.replace(/<span><button data-rb="home\.font\.bold"[^]*?<\/button><\/span>/, pendingHtml('home.font.bold', { key: 'bold', name: '굵게', size: 'small', kind: 'button' }));
    expect(() => generatePage(fake, app)).toThrow(/준비 중/);
  });
  it('표시가 없으면 멈춘다', () => {
    expect(() => generatePage('<header></header>', app)).toThrow(/표시 없음/);
  });
});
```

`tests/ribbonRuntime.test.ts`:
```ts
import { describe, expect, it } from 'vitest';
import { collapseOrder } from '../src/client/ribbonRuntime.js';

describe('collapseOrder — 오피스처럼 접는 순서', () => {
  it('collapse 작은 수 먼저, 같으면 오른쪽 묶음 먼저', () => {
    const groups = [
      { priority: 9, index: 0 },
      { priority: 2, index: 1 },
      { priority: 5, index: 2 },
      { priority: 2, index: 3 },
      { priority: 1, index: 4 },
    ];
    expect(collapseOrder(groups)).toEqual([4, 3, 1, 2, 0]);
  });
  it('빈 목록', () => expect(collapseOrder([])).toEqual([]));
});
```

`tests/ribbonPages.test.ts`:
```ts
import fs from 'node:fs';
import { describe, expect, it } from 'vitest';
import { generatePage, loadInventory, type AppName } from '../scripts/ribbonGenLib.mjs';

// 목록대로 바꾼 앱 — Task 3(word)·4(excel)·5(powerpoint)가 하나씩 더한다
const DONE: readonly AppName[] = [];
const PAGES = ['public/editor.html', 'public/editor-grid.html', 'public/editor-slides.html'];
const before = JSON.parse(fs.readFileSync('tests/fixtures/ribbon-ids-before.json', 'utf8')) as Record<string, string[]>;
const headerIds = (html: string) => [...html.slice(html.indexOf('<header'), html.indexOf('</header>')).matchAll(/\sid="([^"]+)"/g)].map((m) => m[1]);

describe('리본 페이지 마크업', () => {
  const inv = loadInventory();
  for (const page of PAGES) {
    it(`${page} — 바꾸기 전 id가 모두 그대로, 겹치는 id 없음`, () => {
      const ids = headerIds(fs.readFileSync(page, 'utf8'));
      expect(before[page].filter((id) => !ids.includes(id))).toEqual([]);
      expect(ids.filter((id, i) => ids.indexOf(id) !== i)).toEqual([]);
    });
  }
  for (const name of DONE) {
    it(`${name} — 페이지가 목록과 같다(node scripts/gen-ribbon.mjs ${name} 결과와 한 글자도 다르지 않음)`, () => {
      const app = inv.apps[name];
      const html = fs.readFileSync(app.page, 'utf8');
      expect(generatePage(html, app)).toBe(html);
    });
  }
});
```

- [ ] **Step 3: 실패 확인**

Run: `npx vitest run tests/ribbonGen.test.ts tests/ribbonRuntime.test.ts tests/ribbonPages.test.ts`
Expected: ribbonGen·ribbonRuntime FAIL(`extractUnits is not a function` / `Failed to resolve import "../src/client/ribbonRuntime.js"`), ribbonPages는 id 3개 PASS.

- [ ] **Step 4: `scripts/ribbonGenLib.mjs` Part B를 파일 끝에 더하기**

```js
// ---------------- 마크업(Part B) ----------------
const VOID_TAGS = new Set(['input', 'img', 'br', 'hr', 'meta', 'link', 'source', 'col', 'wbr', 'area', 'base', 'embed', 'param', 'track']);
const LAUNCHER_SVG = '<svg viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"><path d="M1.5 5.5v-4h4M4 4l4.5 4.5M8.5 5v3.5H5"/></svg>';
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

/** 여는 태그의 끝 '>' 다음 위치 — 따옴표 안의 '>'(말풍선 "Ctrl+Shift+>")는 건너뛴다 */
function tagEnd(html, start) {
  let quote = '';
  for (let i = start; i < html.length; i += 1) {
    const ch = html[i];
    if (quote) {
      if (ch === quote) quote = '';
    } else if (ch === '"' || ch === "'") quote = ch;
    else if (ch === '>') return i + 1;
  }
  throw new Error('여는 태그가 닫히지 않음');
}

/** data-rb="키"가 붙은 요소를 닫는 태그까지 통째로 꺼낸다. data-rb는 태그 이름 바로 뒤 첫 속성이어야 한다(앞 속성의 '<'에 속지 않게) */
export function extractUnits(html) {
  const units = new Map();
  const re = /data-rb="([^"]+)"/g;
  let m;
  while ((m = re.exec(html))) {
    const key = m[1];
    const start = html.lastIndexOf('<', m.index);
    const head = /^<([a-zA-Z][\w-]*)\s+$/.exec(html.slice(start, m.index));
    if (!head) throw new Error(`data-rb="${key}"는 태그 이름 바로 뒤 첫 속성이어야 함`);
    const tag = head[1].toLowerCase();
    const openEnd = tagEnd(html, m.index);
    let end = openEnd;
    if (!VOID_TAGS.has(tag)) {
      const tagRe = new RegExp(`<(/?)${tag}(?=[\\s>/])`, 'gi');
      tagRe.lastIndex = openEnd;
      let depth = 1;
      let t = null;
      while (depth > 0 && (t = tagRe.exec(html))) {
        if (t[1]) depth -= 1;
        else if (!html.slice(t.index, tagEnd(html, t.index)).endsWith('/>')) depth += 1;
      }
      if (depth > 0 || !t) throw new Error(`data-rb="${key}"의 </${tag}> 없음`);
      end = tagEnd(html, t.index);
    }
    if (units.has(key)) throw new Error(`data-rb="${key}"가 두 번 나옴`);
    units.set(key, html.slice(start, end));
    re.lastIndex = end; // 단추 안에 또 data-rb를 두지 않는다
  }
  return units;
}

/** 여러 줄 조각을 들여쓰기 pad에 맞춘다 — 마지막 줄(닫는 태그) 들여쓰기를 기준으로 옮겨서 두 번 돌려도 같다 */
function reindent(snippet, pad) {
  const lines = snippet.split('\n');
  if (lines.length === 1) return pad + snippet;
  const lead = (s) => s.length - s.trimStart().length;
  const shift = pad.length - lead(lines[lines.length - 1]);
  return lines
    .map((line, i) => {
      if (i === 0) return pad + line;
      if (!line.trim()) return '';
      return shift >= 0 ? ' '.repeat(shift) + line : line.slice(Math.min(-shift, lead(line)));
    })
    .join('\n');
}

const iconAttr = (c) => (ICON_KINDS.includes(c.kind) ? ` data-icon="${esc(c.icon ?? c.key)}"` : '');

/** 준비 중 단추 — 보이되 꺼짐, 말풍선 "준비 중 — 이름"(설계 §1-2) */
export function pendingHtml(key, c) {
  const name = esc(c.name);
  const head = `data-rb="${esc(key)}" data-pending data-kind="${c.kind}" title="준비 중 — ${name}"`;
  switch (c.kind) {
    case 'gallery':
      return `<div ${head} class="rb-gallery" role="group" aria-label="${name}"><span>${name}</span></div>`;
    case 'toggle':
      return `<label ${head} class="rb-check"><input type="checkbox" disabled /><span>${name}</span></label>`;
    case 'input':
      return `<label ${head} class="layout-field rb-field"><span>${name}</span><input type="text" disabled aria-label="${name}" /></label>`;
    case 'select':
      return `<select ${head} class="toolbar-select rb-select" aria-label="${name}" disabled><option>${name}</option></select>`;
    default:
      if (c.size === 'big') return `<button ${head} type="button" class="icon-btn icon-btn-big rb-pending"${iconAttr(c)} aria-label="${name}" disabled><span>${name}</span></button>`;
      return `<button ${head} type="button" class="icon-btn${c.iconOnly ? '' : ' ribbon-btn-labeled'} rb-pending"${iconAttr(c)} aria-label="${name}" disabled>${c.iconOnly ? '' : `<span class="ribbon-btn-label">${name}</span>`}</button>`;
  }
}

/** 다른 탭의 기존 단추를 대신 누르는 단추(officeShell.ts의 data-proxy 처리) */
export function proxyHtml(key, c) {
  const name = esc(c.name);
  const head = `data-rb="${esc(key)}" data-kind="${c.kind}" data-proxy="${esc(c.proxy)}"`;
  if (c.size === 'big') return `<button ${head} type="button" class="icon-btn icon-btn-big"${iconAttr(c)} title="${name}"><span>${name}</span></button>`;
  return `<button ${head} type="button" class="icon-btn ribbon-btn-labeled"${iconAttr(c)} title="${name}"><span class="ribbon-btn-label">${name}</span></button>`;
}

function groupBody(tab, g, units, pad) {
  const p1 = `${pad}  `;
  const p2 = `${pad}    `;
  const p3 = `${pad}      `;
  const html = (c) => {
    const key = controlKey(tab.id, g.id, c);
    if (c.pending) return pendingHtml(key, c);
    if (c.proxy) return proxyHtml(key, c);
    const unit = units.get(key);
    if (!unit) throw new Error(`기존 단추 마크업 없음: data-rb="${key}" (ref ${c.ref}) — 그 요소에 data-rb를 첫 속성으로 다세요`);
    if (/^<[a-zA-Z][\w-]*\s+data-rb="[^"]+"\s+data-pending/.test(unit)) throw new Error(`data-rb="${key}"는 기존 기능인데 마크업이 준비 중 단추 — 실제 단추에 data-rb를 다세요`);
    return unit;
  };
  const out = [];
  if (g.layout === 'fields') {
    let col = null;
    let count = 0;
    const flush = () => {
      if (col) out.push(`${p1}<div class="layout-field-col">\n${col.join('\n')}\n${p1}</div>`);
      col = null;
      count = 0;
    };
    for (const c of g.controls) {
      if (!col || c.stackTitle !== undefined || count === 3) {
        flush();
        col = c.stackTitle !== undefined ? [`${p2}<span class="layout-field-title">${esc(c.stackTitle)}</span>`] : [];
        count = col.length; // 제목 줄도 한 칸(16px 세 줄 = 48px 안)
      }
      col.push(reindent(html(c), p2));
      count += 1;
    }
    flush();
    return `${pad}<div class="ribbon-group-buttons layout-fields">\n${out.join('\n')}\n${pad}</div>`;
  }
  let stack = [];
  let rows = null;
  const flushStack = () => {
    if (stack.length) out.push(`${p1}<div class="ribbon-group-buttons-wrap">\n${stack.map((c) => `${p2}<div class="ribbon-row">\n${reindent(html(c), p3)}\n${p2}</div>`).join('\n')}\n${p1}</div>`);
    stack = [];
  };
  const flushRows = () => {
    if (!rows) return;
    const count = Math.max(...rows.map((c) => c.row)) + 1;
    const lines = Array.from({ length: count }, (_, i) => `${p2}<div class="ribbon-row">\n${rows.filter((c) => c.row === i).map((c) => reindent(html(c), p3)).join('\n')}\n${p2}</div>`);
    out.push(`${p1}<div class="ribbon-group-buttons-wrap rb-rows">\n${lines.join('\n')}\n${p1}</div>`);
    rows = null;
  };
  for (const c of g.controls) {
    if (c.row !== undefined) {
      flushStack();
      (rows ??= []).push(c);
      continue;
    }
    flushRows();
    if (c.size === 'big' || c.kind === 'gallery') {
      flushStack();
      out.push(reindent(html(c), p1));
      continue;
    }
    if (c.break || stack.length === 3) flushStack();
    stack.push(c);
  }
  flushStack();
  flushRows();
  return `${pad}<div class="ribbon-group-buttons">\n${out.join('\n')}\n${pad}</div>`;
}

function groupHtml(tab, g, units, pad) {
  const p1 = `${pad}  `;
  const first = g.controls[0];
  const name = esc(g.name);
  const attrs = [`class="ribbon-group${g.className ? ` ${esc(g.className)}` : ''}"`, g.domId ? `id="${esc(g.domId)}"` : '', `data-group="${esc(g.id)}"`, `data-collapse="${g.collapse}"`, 'role="group"', `aria-label="${name}"`].filter(Boolean).join(' ');
  const lines = [
    `${pad}<div ${attrs}>`,
    `${p1}<button type="button" class="icon-btn icon-btn-big rb-group-toggle" data-icon="${esc(first.icon ?? first.key)}" aria-expanded="false" title="${name}"><span>${name}</span></button>`,
    groupBody(tab, g, units, p1),
    `${p1}<span class="ribbon-group-label">${name}</span>`,
  ];
  if (g.launcher?.pending) lines.push(`${p1}<button type="button" class="ribbon-launcher" data-pending disabled title="준비 중 — ${esc(g.launcher.pending)}" aria-label="${name} 자세히">${LAUNCHER_SVG}</button>`);
  if (g.launcher?.proxy) lines.push(`${p1}<button type="button" class="ribbon-launcher" data-proxy="${esc(g.launcher.proxy)}" title="${name} 자세히" aria-label="${name} 자세히">${LAUNCHER_SVG}</button>`);
  lines.push(`${pad}</div>`);
  return lines.join('\n');
}

export function renderPanels(app, units, pad = '    ') {
  return allTabs(app)
    .map((tab, i) => {
      const attrs = `class="toolbar-group toolbar-format ribbon-panel"${tab.panelId ? ` id="${esc(tab.panelId)}"` : ''} data-panel="${esc(tab.id)}" role="tabpanel"${i === 0 ? '' : ' hidden'}`;
      return `${pad}<div ${attrs}>\n${tab.groups.map((g) => groupHtml(tab, g, units, `${pad}  `)).join('\n')}\n${pad}</div>`;
    })
    .join('\n');
}

export function renderTabs(app, pad = '      ') {
  const one = (t, cls, extra = '') => `${pad}<button class="${cls}" role="tab" data-tab="${esc(t.id)}"${t.editOnly ? ' data-edit-only' : ''}${extra}>${esc(t.name)}</button>`;
  return [`${pad}<button class="ribbon-tab ribbon-tab-file" role="tab" data-tab="file">파일</button>`, ...app.tabs.map((t) => one(t, 'ribbon-tab')), ...app.contextual.map((t) => one(t, 'ribbon-tab ribbon-tab-contextual', ' hidden'))].join('\n');
}

function region(html, name) {
  const begin = new RegExp(`([ \\t]*)<!-- ribbon:${name}:begin[^>]*-->\\n`).exec(html);
  if (!begin) throw new Error(`표시 없음: <!-- ribbon:${name}:begin -->`);
  const start = begin.index + begin[0].length;
  const end = new RegExp(`[ \\t]*<!-- ribbon:${name}:end -->`).exec(html.slice(start));
  if (!end) throw new Error(`표시 없음: <!-- ribbon:${name}:end -->`);
  return { start, stop: start + end.index, pad: begin[1] };
}

/** 페이지 HTML의 두 표시(탭·판) 사이를 목록대로 다시 만든다. 기존 단추는 data-rb로 찾아 마크업을 글자 그대로 옮긴다 */
export function generatePage(source, app) {
  const crlf = source.includes('\r\n');
  let html = crlf ? source.replace(/\r\n/g, '\n') : source;
  const panels = region(html, 'panels');
  const units = extractUnits(html.slice(panels.start, panels.stop));
  const known = new Set([...eachControl(app)].map(({ key }) => key));
  const unknown = [...units.keys()].filter((k) => !known.has(k));
  if (unknown.length) throw new Error(`목록에 없는 data-rb: ${unknown.join(', ')}`);
  html = `${html.slice(0, panels.start)}${renderPanels(app, units, panels.pad)}\n${html.slice(panels.stop)}`;
  const tabs = region(html, 'tabs');
  html = `${html.slice(0, tabs.start)}${renderTabs(app, tabs.pad)}\n${html.slice(tabs.stop)}`;
  return crlf ? html.replace(/\n/g, '\r\n') : html;
}

// ---------------- e2e가 쓰는 기대 모양 ----------------

/** DOM에서 읽은 것과 같은 모양 — scripts/verify-ribbon-inventory-e2e.mjs가 1:1로 맞댄다 */
export function expectedStructure(app) {
  const tabs = [
    { id: 'file', name: '파일', editOnly: false, contextual: false },
    ...app.tabs.map((t) => ({ id: t.id, name: t.name, editOnly: t.editOnly, contextual: false })),
    ...app.contextual.map((t) => ({ id: t.id, name: t.name, editOnly: t.editOnly, contextual: true })),
  ];
  const panels = Object.fromEntries(
    allTabs(app).map((t) => [
      t.id,
      t.groups.map((g) => ({
        id: g.id,
        name: g.name,
        collapse: g.collapse,
        launcher: g.launcher ? (g.launcher.pending ? 'pending' : `proxy:${g.launcher.proxy}`) : null,
        controls: g.controls.map((c) => ({ key: controlKey(t.id, g.id, c), name: c.name, size: c.size, kind: c.kind, pending: !!c.pending })),
      })),
    ]),
  );
  return { tabs, panels };
}

/** 기존 기능(대리 단추 포함)의 data-rb → 누르는 요소 선택자 */
export function refsOf(app) {
  return Object.fromEntries([...eachControl(app)].filter(({ c }) => !c.pending).map(({ key, c }) => [key, c.ref ?? `[data-rb="${key}"]`]));
}
```

`scripts/ribbonGenLib.d.mts` 끝에 더하기:
```ts
// Part B
export declare function extractUnits(html: string): Map<string, string>;
export declare function pendingHtml(key: string, c: RibbonControl): string;
export declare function proxyHtml(key: string, c: RibbonControl): string;
export declare function renderPanels(app: RibbonApp, units: Map<string, string>, pad?: string): string;
export declare function renderTabs(app: RibbonApp, pad?: string): string;
export declare function generatePage(source: string, app: RibbonApp): string;
export interface ExpectedControl { key: string; name: string; size: string; kind: string; pending: boolean }
export interface ExpectedGroup { id: string; name: string; collapse: number; launcher: string | null; controls: ExpectedControl[] }
export declare function expectedStructure(app: RibbonApp): { tabs: { id: string; name: string; editOnly: boolean; contextual: boolean }[]; panels: Record<string, ExpectedGroup[]> };
export declare function refsOf(app: RibbonApp): Record<string, string>;
```

- [ ] **Step 5: `scripts/gen-ribbon.mjs` 쓰기**

```js
#!/usr/bin/env node
// 리본 마크업 만들기(설계 2026-11-22-office-ribbon-structure) — src/client/ribbon/inventory.json → public/editor*.html 의 두 표시 사이.
// 쓰는 법: node scripts/gen-ribbon.mjs [--check] [word|excel|powerpoint …]   (--check: 쓰지 않고 다르면 종료 코드 1)
import fs from 'node:fs';
import { generatePage, loadInventory, validateInventory } from './ribbonGenLib.mjs';

const args = process.argv.slice(2);
const check = args.includes('--check');
const only = args.filter((a) => !a.startsWith('--'));
const inv = loadInventory();
const problems = validateInventory(inv);
if (problems.length) {
  console.error(`목록 규칙 위반 ${problems.length}건:\n${problems.join('\n')}`);
  process.exit(1);
}
let failed = 0;
for (const [name, app] of Object.entries(inv.apps)) {
  if (only.length && !only.includes(name)) continue;
  const before = fs.readFileSync(app.page, 'utf8');
  let after;
  try {
    after = generatePage(before, app);
  } catch (error) {
    console.error(`${app.page}: ${error.message}`);
    failed += 1;
    continue;
  }
  if (after === before) console.log(`그대로 ${app.page}`);
  else if (check) {
    console.error(`다름 ${app.page} — node scripts/gen-ribbon.mjs ${name}`);
    failed += 1;
  } else {
    fs.writeFileSync(app.page, after);
    console.log(`만듦 ${app.page}`);
  }
}
process.exit(failed ? 1 : 0);
```

- [ ] **Step 6: `src/client/ribbonRuntime.ts` 쓰기**

```ts
/**
 * 오피스 리본 구조(설계 2026-11-22-office-ribbon-structure)의 브라우저 쪽 두 가지 — 마크업은 scripts/gen-ribbon.mjs가 정적으로 만든다.
 *  1) 준비 중·대리·접힌 묶음 단추의 아이콘(data-icon)을 고른 탭에서만 단다(스프라이트 /ribbon/icons.svg 하나, 탭마다 늦게).
 *  2) 좁으면 오피스처럼 묶음을 큰 단추 하나로 접는다 — data-collapse 작은 수부터(같으면 오른쪽부터), 600px 이하는 모두.
 *     접힌 묶음 단추를 누르면 원래 단추들(.ribbon-group-buttons)이 그 아래 떠 있는 판으로 보인다(DOM은 그대로 — id·처리기 유지).
 */
export const ICON_SPRITE = '/ribbon/icons.svg';
export const NARROW_PX = 600;
const SVG_NS = 'http://www.w3.org/2000/svg';

export function collapseOrder(groups: readonly { priority: number; index: number }[]): number[] {
  return [...groups].sort((a, b) => a.priority - b.priority || b.index - a.index).map((g) => g.index);
}

export function hydrateIcons(panel: HTMLElement): void {
  for (const el of panel.querySelectorAll<HTMLElement>('[data-icon]')) {
    if (el.querySelector(':scope > svg')) continue;
    const svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('class', 'rb-ico');
    svg.setAttribute('aria-hidden', 'true');
    const use = document.createElementNS(SVG_NS, 'use');
    use.setAttribute('href', `${ICON_SPRITE}#rbi-${el.dataset.icon ?? ''}`);
    svg.append(use);
    el.prepend(svg);
  }
}

let openGroup: HTMLElement | null = null;

export function closeGroupPopup(): void {
  if (!openGroup) return;
  openGroup.classList.remove('rb-open');
  openGroup.querySelector(':scope > .rb-group-toggle')?.setAttribute('aria-expanded', 'false');
  openGroup = null;
}

export function fitPanel(panel: HTMLElement, viewportWidth: number): void {
  const groups = [...panel.querySelectorAll<HTMLElement>(':scope > .ribbon-group[data-collapse]')];
  closeGroupPopup();
  for (const g of groups) g.classList.remove('rb-collapsed');
  if (viewportWidth <= NARROW_PX) {
    for (const g of groups) g.classList.add('rb-collapsed');
    return;
  }
  for (const index of collapseOrder(groups.map((g, i) => ({ priority: Number(g.dataset.collapse), index: i })))) {
    if (panel.scrollWidth <= panel.clientWidth + 1) break;
    groups[index].classList.add('rb-collapsed');
  }
}

export function installGroupToggles(toolbar: HTMLElement): void {
  toolbar.addEventListener('click', (event) => {
    const toggle = (event.target as Element).closest<HTMLElement>('.rb-group-toggle');
    const group = toggle?.closest<HTMLElement>('.ribbon-group');
    if (!toggle || !group) return;
    const wasOpen = openGroup === group;
    closeGroupPopup();
    if (wasOpen) return;
    group.classList.add('rb-open');
    toggle.setAttribute('aria-expanded', 'true');
    openGroup = group;
    const body = group.querySelector<HTMLElement>(':scope > .ribbon-group-buttons');
    if (body) {
      const r = toggle.getBoundingClientRect();
      body.style.left = `${Math.max(4, Math.min(r.left, window.innerWidth - body.offsetWidth - 4))}px`;
      body.style.top = `${r.bottom + 2}px`;
    }
  });
  document.addEventListener('mousedown', (event) => {
    const target = event.target as Element;
    if (openGroup && !openGroup.contains(target) && !target.closest('.dropdown-menu')) closeGroupPopup();
  });
  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') closeGroupPopup();
  });
}
```

- [ ] **Step 7: `src/client/officeShell.ts` 고치기** (네 군데)

(a) import 줄 바로 아래(`import { EDIT_DOC_KEY, … } from './viewStrip.js';` 다음 줄)에:
```ts
import { closeGroupPopup, fitPanel, hydrateIcons, installGroupToggles } from './ribbonRuntime.js';
```

(b) `render()` 안 `for (const panel of panels) panel.hidden = !activeVisible || panel.dataset.panel !== activeTab;` 바로 아래에:
```ts
    // 고른 탭만 아이콘을 달고(탭마다 늦게) 판이 넘치면 묶음을 접는다(오피스처럼) — 축소된 리본은 펼쳐 볼 때 다시 부른다
    const shown = panels.find((panel) => !panel.hidden);
    if (shown && !(collapsed && !peeking)) {
      hydrateIcons(shown);
      fitPanel(shown, window.innerWidth);
    }
```
그리고 `selectTab()` 안 `activeTab = tab;` 바로 위에 `closeGroupPopup();`.

(c) `if (!options.canEdit && !viewUi) { for (const tab of tabs) if (tab.dataset.editOnly !== undefined) tab.hidden = true; }` 블록 바로 아래에:
```ts
  // 사람 확인·목록 대조 시험용 — ?rbctx=1 이면 상황별 탭을 모두 보이고, 편집기가 선택에 따라 숨기려 해도 그대로 둔다
  const showAllContextual = new URLSearchParams(location.search).get('rbctx') === '1';
  if (showAllContextual) for (const tab of tabs) if (tab.classList.contains('ribbon-tab-contextual')) tab.hidden = false;
```
`document.getElementById('ribbon-collapse')?.addEventListener('click', toggleCollapsed);` 바로 아래에:
```ts
  installGroupToggles(toolbar);
  let fitFrame = 0;
  window.addEventListener('resize', () => {
    cancelAnimationFrame(fitFrame);
    fitFrame = requestAnimationFrame(() => {
      const shown = panels.find((panel) => !panel.hidden);
      if (shown && !(collapsed && !peeking)) fitPanel(shown, window.innerWidth);
    });
  });
  // 메뉴 글꼴(Outfit)이 늦게 오면 폭이 바뀐다 — 받은 뒤 한 번 더 맞춘다
  void document.fonts?.ready.then(() => render());
```
`return { … setContextualTab(tab, visible) {` 의 첫 줄로:
```ts
      if (showAllContextual) return;
```

(d) 보기 꾸밈 `greyOut` 안 `for (const el of document.querySelectorAll<HTMLElement>(VIEW_EDIT_CONTROLS)) {` 바로 다음 줄에:
```ts
        // 준비 중 단추는 disabled가 우선(흐림 클래스 안 붙임), 접힌 묶음 단추는 둘러보기용 — 설계 2026-11-22-office-ribbon-structure §1-3
        if (el.closest('[data-pending]') || el.classList.contains('rb-group-toggle')) continue;
```

- [ ] **Step 8: `public/css/style.css` 고치기**

(a) 리본 꾸밈 묶음의 스타일 갤러리 줄 `.ribbon-skin .style-gallery { flex-wrap: wrap; align-content: flex-start; overflow: hidden auto; max-width: min(760px, 44vw); }`에서 `max-width: min(760px, 44vw)` → `max-width: min(520px, 30vw)` (홈 탭에 묶음이 늘어 갤러리가 1280px에서 홈 탭 절반을 먹지 않게 — 오피스도 좁으면 갤러리 칸 수를 줄인다).

(b) 파일 맨 끝에 더하기:
```css
/* ==================== 오피스 리본 구조(설계 2026-11-22-office-ribbon-structure) ====================
   scripts/gen-ribbon.mjs가 만드는 준비 중 단추(보이되 꺼짐 — disabled가 읽기 전용 흐림보다 우선)·빈 갤러리·확인란·입력칸·목록,
   좁을 때 묶음을 큰 단추 하나로 접기(src/client/ribbonRuntime.ts). 62px·큰 32/작은 16px 규격은 위 리본 꾸밈 그대로 — 높이를 늘리는 규칙을 넣지 말 것 */
.ribbon-group > .rb-group-toggle { display: none; }
.ribbon-skin [data-pending] { opacity: 0.45; cursor: default; }
.ribbon-skin [data-pending]:disabled,
.ribbon-skin [data-pending] :disabled { cursor: default; }
/* 준비 중 아이콘(스프라이트 단색 선) — 선 굵기·색은 여기서 정해 <use> 안으로 물려준다 */
.ribbon-skin .rb-ico { flex: none; width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.5; stroke-linecap: round; stroke-linejoin: round; }
.ribbon-skin .icon-btn-big > .rb-ico { width: var(--skin-big-icon); height: var(--skin-big-icon); stroke-width: 1.25; }
/* 드롭다운·분할 단추 표시(▾) — 글자(textContent)에는 넣지 않는다 */
.ribbon-skin :is(.rb-pending, [data-proxy])[data-kind="dropdown"] > span::after,
.ribbon-skin :is(.rb-pending, [data-proxy])[data-kind="split"] > span::after { content: " ▾"; }
.ribbon-skin .rb-pending:is([data-kind="dropdown"], [data-kind="split"]):not(:has(span))::after { content: "▾"; font-size: 8px; line-height: 1; }
.ribbon-skin .rb-gallery { display: inline-flex; align-items: center; justify-content: center; box-sizing: border-box; width: 112px; height: 44px; border: 1px solid var(--skin-line); border-radius: 3px; color: var(--skin-text-dim); font: 400 10px/12px var(--skin-font); }
.ribbon-skin .rb-check { display: inline-flex; align-items: center; gap: 3px; height: var(--skin-small-h); color: var(--skin-text); font: 400 11px/14px var(--skin-font); white-space: nowrap; }
.ribbon-skin .rb-check input { width: 12px; height: 12px; margin: 0; }
.ribbon-skin .rb-field input { width: 44px; }
.ribbon-skin .rb-select { max-width: 9em; }
/* 접힌 묶음 — 묶음 이름 큰 단추(아이콘 32 + 이름 ▾)가 판 높이 전체, 누르면 원래 단추들이 그 아래 떠 있는 판으로 */
.ribbon-skin .ribbon-group.rb-collapsed > .rb-group-toggle { display: inline-flex; grid-row: 1 / -1; }
.ribbon-skin .rb-group-toggle > span::after { content: " ▾"; }
.ribbon-skin .rb-group-toggle[aria-expanded="true"] { background: var(--skin-hover); }
.ribbon-skin .ribbon-group.rb-collapsed > :is(.ribbon-group-buttons, .ribbon-group-label, .ribbon-launcher) { display: none; }
.ribbon-skin .ribbon-group.rb-collapsed.rb-open > .ribbon-group-buttons { display: flex; position: fixed; z-index: 70; box-sizing: border-box; height: calc(var(--skin-ribbon-h) - var(--skin-label-h) + 8px); padding: 4px 6px; border-radius: 6px; background: #ffffff; box-shadow: 0 0 0 1px var(--skin-line), 0 4px 12px rgba(0, 0, 0, 0.14); }
@media (max-width: 600px) {
  .ribbon-skin .ribbon-group.rb-collapsed.rb-open > .ribbon-group-buttons { height: auto; max-width: calc(100vw - 8px); flex-wrap: wrap; }
}
```

- [ ] **Step 9: 통과 확인**

Run: `npx vitest run tests/ribbonInventory.test.ts tests/ribbonGen.test.ts tests/ribbonRuntime.test.ts tests/ribbonPages.test.ts && npx tsc --noEmit && npm run build:editor`
Expected: 전부 PASS, 타입 오류 없음, 빌드 성공. (`node scripts/gen-ribbon.mjs --check`는 아직 표시가 없어 세 페이지 모두 "표시 없음"으로 실패하는 것이 맞다 — Task 3~5가 하나씩 고친다.)

- [ ] **Step 10: 커밋**

```bash
git add scripts/ribbonGenLib.mjs scripts/ribbonGenLib.d.mts scripts/gen-ribbon.mjs src/client/ribbonRuntime.ts src/client/officeShell.ts public/css/style.css tests/ribbonGen.test.ts tests/ribbonRuntime.test.ts tests/ribbonPages.test.ts tests/fixtures/ribbon-ids-before.json
git diff --cached --stat
git commit -m "오피스 리본 구조 1: 목록으로 리본 마크업을 만드는 생성기(gen-ribbon, 기존 단추는 data-rb로 글자 그대로 옮김·결정적·--check), 좁으면 오피스처럼 묶음 접기·고른 탭만 아이콘 달기(ribbonRuntime), 준비 중·접힘 단추는 읽기 전용 흐림에서 뺌, ?rbctx=1 상황별 탭 보기

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 3: 워드 — 전체 탭·묶음·단추

**Files:**
- Modify: `public/editor.html`, `src/client/editor.ts:747`(근처 — `shell.setContextualTab('table', …)` 줄), `tests/ribbonPages.test.ts`(`DONE`에 `'word'`), `src/client/ribbon/inventory.json`(배치 확인 뒤 `collapse`만 고칠 수 있음)

**Interfaces:**
- Consumes: Task 2의 `generatePage`·`gen-ribbon.mjs`·마크업 약속, 목록 `apps.word`
- Produces: 목록과 1:1인 `public/editor.html`(탭 10 + 상황별 5), 상황별 탭 `table-design`이 표 안 커서에서 보임

- [ ] **Step 1: 실패하는 시험** — `tests/ribbonPages.test.ts`의 `const DONE: readonly AppName[] = [];` → `['word']`

Run: `npx vitest run tests/ribbonPages.test.ts`
Expected: FAIL — `표시 없음: <!-- ribbon:panels:begin -->`

- [ ] **Step 2: 표시 넣기** — `public/editor.html`
  - 탭: `<div class="ribbon-tabs" role="tablist" aria-label="리본">` 안, `<button class="ribbon-tab ribbon-tab-file" …>파일</button>` 줄 바로 위에 `      <!-- ribbon:tabs:begin -->`, `<span class="ribbon-tabs-spacer"></span>` 줄 바로 위에 `      <!-- ribbon:tabs:end -->`
  - 판: `    <!-- ===== 홈 ===== -->` 줄 바로 위에 `    <!-- ribbon:panels:begin — scripts/gen-ribbon.mjs가 src/client/ribbon/inventory.json으로 만든다(손으로 고치지 말 것) -->`, `  </header>` 줄 바로 위에 `    <!-- ribbon:panels:end -->`

- [ ] **Step 3: 기존 단추에 `data-rb` 달기** — 아래 표의 "다는 요소"의 **여는 태그 이름 바로 뒤 첫 속성**으로 `data-rb="<키>"`를 넣는다(예: `<span class="dropdown-btn" id="change-case-dropdown">` → `<span data-rb="home.font.change-case" class="dropdown-btn" id="change-case-dropdown">`, `<label class="layout-field"><span>왼쪽</span><input id="indent-left" …` → `<label data-rb="layout.paragraph.indent-left" class="layout-field">…`). "대리"는 생성기가 새로 쓰므로 달지 않는다. 같은 줄의 "고칠 것"도 이때 한다 — `title`·`aria-label`·이름 고치기는 **누르는 요소**(목록의 `ref`, 예: `#highlight-dropdown`이면 안의 `#btn-highlight-apply`)에, 이름이 보이는 글자(`<span>`·`.ribbon-btn-label`)면 그 글자를 고친다. 말풍선 뒤의 단축키 `(Ctrl+…)`와 ` — 설명`은 그대로 둔다.

| 키(`data-rb` 값) | 다는 요소 | 자리 | 고칠 것(convert·fix) |
|---|---|---|---|
| `home.clipboard.paste` | `#btn-paste` | same |  |
| `home.clipboard.cut` | `#btn-cut` | same |  |
| `home.clipboard.copy` | `#btn-copy` | same |  |
| `home.clipboard.format-painter` | `#btn-format-painter` | same |  |
| `home.font.font-family` | `#font-family` | same |  |
| `home.font.font-size` | `#font-size` | same |  |
| `home.font.grow` | `#btn-font-grow` | same |  |
| `home.font.shrink` | `#btn-font-shrink` | same |  |
| `home.font.change-case` | `#change-case-dropdown` | same |  |
| `home.font.clear-format` | `#btn-clear-format` | same |  |
| `home.font.bold` | `#btn-bold` | same |  |
| `home.font.italic` | `#btn-italic` | same |  |
| `home.font.underline` | `#btn-underline` | same |  |
| `home.font.strike` | `#btn-strike` | same |  |
| `home.font.subscript` | `#btn-subscript` | same |  |
| `home.font.superscript` | `#btn-superscript` | same |  |
| `home.font.highlight` | `#highlight-dropdown` | same |  |
| `home.font.font-color` | `#font-color-dropdown` | same |  |
| `home.paragraph.bullets` | `#btn-bullet-list` | same |  |
| `home.paragraph.numbering` | `#btn-ordered-list` | same |  |
| `home.paragraph.multilevel` | `#btn-multilevel-list` | same |  |
| `home.paragraph.outdent` | `#btn-outdent` | same |  |
| `home.paragraph.indent` | `#btn-indent` | same |  |
| `home.paragraph.align-left` | `#btn-align-left` | same |  |
| `home.paragraph.align-center` | `#btn-align-center` | same |  |
| `home.paragraph.align-right` | `#btn-align-right` | same |  |
| `home.paragraph.justify` | `#btn-align-justify` | same |  |
| `home.paragraph.line-spacing` | `#line-spacing-dropdown` | same |  |
| `home.paragraph.shading` | `#shading-dropdown` | same |  |
| `home.styles.style-gallery` | `.style-gallery` | same | div.style-gallery 에 aria-label="스타일" 더함 |
| `home.editing.find` | `#btn-ribbon-find` | same |  |
| `home.editing.replace` | `#btn-ribbon-replace` | same |  |
| `home.editing.select` | `#btn-ribbon-select` | nearest | 이름 "모두 선택" → "선택"(말풍선 "모두 선택 (Ctrl+A)" 그대로) |
| `insert.pages.page-break` | `#btn-page-break` | moved | big→small |
| `insert.pages.line-break` | `#btn-insert-break` | nearest | big→small |
| `insert.tables.table` | `#table-insert-dropdown` | same |  |
| `insert.comments.comment` | (생성 — 대리 단추 btn-new-comment) | proxy |  |
| `layout.page-setup.margins` | `#margins-dropdown` | same |  |
| `layout.page-setup.orientation` | `#orientation-dropdown` | same |  |
| `layout.page-setup.size` | `#paper-size-dropdown` | same |  |
| `layout.paragraph.indent-left` | `label.layout-field:has(> #indent-left)` | same |  |
| `layout.paragraph.indent-right` | `label.layout-field:has(> #indent-right)` | same |  |
| `layout.paragraph.spacing-before` | `label.layout-field:has(> #spacing-before)` | same |  |
| `layout.paragraph.spacing-after` | `label.layout-field:has(> #spacing-after)` | same |  |
| `review.proofing.spelling` | `#btn-spellcheck` | nearest | 이름 "맞춤법 검사" → "맞춤법 및 문법 검사" |
| `review.proofing.word-count` | `#btn-word-count` | same | big→small |
| `review.comments.new-comment` | `#btn-new-comment` | same |  |
| `review.comments.delete-comment` | `#btn-delete-comment` | same |  |
| `review.comments.previous-comment` | `#btn-prev-comment` | same |  |
| `review.comments.next-comment` | `#btn-next-comment` | same |  |
| `review.comments.show-comments` | `#btn-show-comments` | same | small→big |
| `review.tracking.reviewing-pane` | `#btn-show-changes` | nearest | big→small; 이름 "변경 내용" → "검토 창" |
| `view.zoom.zoom-100` | `#btn-zoom-100` | same |  |
| `view.zoom.page-width` | `#btn-zoom-page-width` | same | big→small |
| `view.zoom.zoom-in` | `#btn-zoom-in-ribbon` | nearest |  |
| `view.zoom.zoom-out` | `#btn-zoom-out-ribbon` | nearest |  |
| `table.rows-columns.delete-table` | `#btn-del-table` | nearest | big→small |
| `table.rows-columns.delete-row` | `#btn-del-row` | nearest |  |
| `table.rows-columns.delete-column` | `#btn-del-col` | nearest |  |
| `table.rows-columns.insert-above` | `#btn-add-row-before` | same | small→big |
| `table.rows-columns.insert-below` | `#btn-add-row` | same | small→big |
| `table.rows-columns.insert-left` | `#btn-add-col-before` | same | small→big |
| `table.rows-columns.insert-right` | `#btn-add-col` | same | small→big |
| `table.merge.merge-cells` | `#btn-merge-cells` | same |  |
| `table.merge.split-cells` | `#btn-split-cell` | same |  |

크기 바꾸기 방법(같은 단추의 클래스와 이름 칸만):
- `big→small`: `class="icon-btn icon-btn-big"` → `class="icon-btn ribbon-btn-labeled"`, 안의 `<span>이름</span>` → `<span class="ribbon-btn-label">이름</span>`
- `small→big`: 거꾸로(`ribbon-btn-labeled` → `icon-btn-big`, `<span class="ribbon-btn-label">` → `<span>`)

- [ ] **Step 4: 생성**

Run: `node scripts/gen-ribbon.mjs word && node scripts/gen-ribbon.mjs --check word`
Expected: `만듦 public/editor.html` 다음 `그대로 public/editor.html`. 오류가 나면 메시지의 `data-rb`를 표와 맞춰 고치고 다시(생성기는 기존 단추를 하나도 잃지 않으면 진행하지 않는다).

- [ ] **Step 5: 상황별 탭 연결** — `src/client/editor.ts`, `dispatchTransaction` 안 `shell.setContextualTab('table', canEdit && isInTable(newState));` 바로 위에:

```ts
      shell.setContextualTab('table-design', canEdit && isInTable(newState)); // 오피스처럼 표 안에서는 "테이블 디자인"·"레이아웃" 두 탭
```

- [ ] **Step 6: 확인**

Run: `npx vitest run tests/ribbonPages.test.ts tests/ribbonInventory.test.ts && npx tsc --noEmit && npm run build:editor`
Expected: PASS(워드 페이지 = 목록, 바꾸기 전 id 82개 모두 있음, 겹침 없음). 눈으로: `git diff --stat public/editor.html`에서 판 밖(찾기 줄·메모 창·상태 줄) 줄이 바뀌지 않았는지.

- [ ] **Step 7: 커밋**

```bash
git add public/editor.html src/client/editor.ts tests/ribbonPages.test.ts
git diff --cached --stat
git commit -m "오피스 리본 구조 2: 워드 — M365 탭 10(홈·삽입·그리기·디자인·레이아웃·참조·편지·검토·보기)·상황별 5, 기존 단추 64개를 오피스 자리로(id·처리기 그대로), 없는 기능은 준비 중 단추, 표 안에서 테이블 디자인 탭

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 4: 엑셀 — 전체 탭·묶음·단추

**Files:**
- Modify: `public/editor-grid.html`, `src/client/grid-editor.ts:4060`(근처 — `setChartTabVisible:` 줄), `tests/ribbonPages.test.ts`(`DONE`에 `'excel'`)

**Interfaces:**
- Consumes: Task 2의 생성기·마크업 약속, 목록 `apps.excel`
- Produces: 목록과 1:1인 `public/editor-grid.html`(탭 9 + 상황별 4), 차트를 고르면 `chart`(차트 디자인)와 `chart-format`(서식) 두 탭

- [ ] **Step 1: 실패하는 시험** — `DONE`을 `['word', 'excel']`로

Run: `npx vitest run tests/ribbonPages.test.ts`
Expected: FAIL — excel `표시 없음`

- [ ] **Step 2: 표시 넣기** — `public/editor-grid.html`
  - 탭: 워드와 같게(`파일` 탭 줄 위 `<!-- ribbon:tabs:begin -->`, `ribbon-tabs-spacer` 줄 위 `<!-- ribbon:tabs:end -->`, 6칸 들여쓰기)
  - 판: 첫 `<div class="toolbar-group toolbar-format ribbon-panel" data-panel="home" role="tabpanel">` 줄 바로 위에 `    <!-- ribbon:panels:begin — scripts/gen-ribbon.mjs가 src/client/ribbon/inventory.json으로 만든다(손으로 고치지 말 것) -->`(바로 위가 `<!-- ===== … ===== -->` 주석 줄이면 그 주석 위에), `  </header>` 줄 바로 위에 `    <!-- ribbon:panels:end -->`

- [ ] **Step 3: 기존 단추에 `data-rb` 달기·고치기** — 방법은 Task 3 Step 3과 같다(첫 속성, 크기 바꾸기 규칙 동일).

| 키(`data-rb` 값) | 다는 요소 | 자리 | 고칠 것(convert·fix) |
|---|---|---|---|
| `home.clipboard.paste` | `#btn-paste` | same |  |
| `home.clipboard.cut` | `#btn-cut` | same |  |
| `home.clipboard.copy` | `#btn-copy` | same |  |
| `home.font.font-family` | `#cell-font-family` | same |  |
| `home.font.font-size` | `#cell-font-size` | same | aria-label·title "글자 크기" → "글꼴 크기" |
| `home.font.grow` | `#btn-font-size-up` | same | title "글자 크기 크게" → "글꼴 크기 크게" |
| `home.font.shrink` | `#btn-font-size-down` | same | title "글자 크기 작게" → "글꼴 크기 작게" |
| `home.font.bold` | `#btn-cell-bold` | same |  |
| `home.font.italic` | `#btn-cell-italic` | same | title "기울임 (Ctrl+I)" → "기울임꼴 (Ctrl+I)" |
| `home.font.underline` | `#btn-cell-underline` | same |  |
| `home.font.borders` | `#border-dropdown` | same |  |
| `home.font.fill-color` | `#cell-bg-color-btn` | same | title "배경 색 채우기" → "채우기 색" |
| `home.font.font-color` | `#cell-font-color-btn` | same | title "글자 색 채우기" → "글꼴 색" |
| `home.font.strikethrough` | `#btn-cell-strike` | nearest |  |
| `home.alignment.top-align` | `#btn-valign-top` | same |  |
| `home.alignment.middle-align` | `#btn-valign-middle` | same |  |
| `home.alignment.bottom-align` | `#btn-valign-bottom` | same |  |
| `home.alignment.wrap-text` | `#btn-wrap-text` | same |  |
| `home.alignment.align-left` | `#btn-cell-align-l` | same | title "왼쪽 정렬" → "왼쪽 맞춤" |
| `home.alignment.center` | `#btn-cell-align-c` | same | title "가운데 정렬" → "가운데 맞춤" |
| `home.alignment.align-right` | `#btn-cell-align-r` | same | title "오른쪽 정렬" → "오른쪽 맞춤" |
| `home.alignment.merge-center` | `#merge-dropdown` | same |  |
| `home.number.number-format` | `#cell-numfmt-select` | same |  |
| `home.number.accounting` | `#btn-numfmt-currency` | same |  |
| `home.number.percent` | `#btn-numfmt-percent` | same |  |
| `home.number.comma` | `#btn-numfmt-comma` | same |  |
| `home.number.increase-decimal` | `#btn-numfmt-decimal-up` | same |  |
| `home.number.decrease-decimal` | `#btn-numfmt-decimal-down` | same |  |
| `home.styles.conditional-formatting` | `#cf-dropdown` | same |  |
| `home.styles.format-as-table` | `#home-table-style-dropdown` | same |  |
| `home.cells.insert-row` | `#btn-add-row` | nearest |  |
| `home.cells.insert-column` | `#btn-add-col` | nearest |  |
| `home.cells.delete-row` | `#btn-del-row` | nearest |  |
| `home.cells.delete-column` | `#btn-del-col` | nearest |  |
| `home.editing.autosum` | `#autosum-dropdown-home` | same |  |
| `home.editing.clear` | `#btn-cell-clear-format` | moved | title "서식 지우기" → "지우기 — 서식 지우기" |
| `home.editing.sort-filter` | `#home-sort-filter-dropdown` | same |  |
| `home.editing.find-select` | `#btn-find-replace` | same |  |
| `insert.tables.table` | `#table-style-dropdown` | nearest | 이름 "표 서식 ▾" → "표 ▾"(말풍선 그대로) |
| `insert.charts.column` | `[data-insert-chart="column"]` | nearest | big→small |
| `insert.charts.bar` | `[data-insert-chart="bar"]` | nearest | big→small |
| `insert.charts.line` | `[data-insert-chart="line"]` | nearest | big→small |
| `insert.charts.pie` | `[data-insert-chart="pie"]` | nearest | big→small |
| `formulas.function-library.autosum` | `#autosum-dropdown-formulas` | same |  |
| `formulas.function-library.average` | `.autosum-apply[data-fn="AVERAGE"]` | nearest |  |
| `formulas.function-library.max` | `.autosum-apply[data-fn="MAX"]` | nearest |  |
| `formulas.function-library.count` | `.autosum-apply[data-fn="COUNT"]` | nearest |  |
| `formulas.function-library.min` | `.autosum-apply[data-fn="MIN"]` | nearest |  |
| `data.sort-filter.sort-asc` | `#btn-sort-asc` | same | 이름 "오름차순" → "오름차순 정렬" |
| `data.sort-filter.sort-desc` | `#btn-sort-desc` | same | 이름 "내림차순" → "내림차순 정렬" |
| `data.sort-filter.sort` | `#btn-sort-custom` | same |  |
| `data.sort-filter.filter` | `#btn-filter` | same |  |
| `data.sort-filter.clear` | `#btn-filter-clear` | same |  |
| `view.zoom.zoom-100` | `#btn-zoom-100` | same |  |
| `view.zoom.zoom-in` | `#btn-zoom-in-ribbon` | nearest |  |
| `view.zoom.zoom-out` | `#btn-zoom-out-ribbon` | nearest |  |
| `view.window.freeze-panes` | `#btn-freeze-panes` | same | big→small |
| `chart.chart-layouts.chart-title` | `label.chart-title-field` | nearest | 이름 "제목" → "차트 제목" |
| `chart.type.column` | `[data-chart-kind="column"]` | nearest |  |
| `chart.type.bar` | `[data-chart-kind="bar"]` | nearest |  |
| `chart.type.line` | `[data-chart-kind="line"]` | nearest |  |
| `chart.type.pie` | `[data-chart-kind="pie"]` | nearest |  |
| `chart.location.delete` | `#btn-chart-delete` | nearest |  |

주의: `label.chart-title-field`는 `<label class="layout-field chart-title-field">` — `data-rb`를 `<label` 바로 뒤에. `.autosum-apply[data-fn="…"]` 넷은 수식 탭의 작은 단추(`<button class="icon-btn ribbon-btn-labeled autosum-apply" data-fn="AVERAGE"` → `<button data-rb="formulas.function-library.average" class=…`). `[data-insert-chart="…"]` 넷은 큰 단추를 작은 단추로 바꾼다(`big→small`).

- [ ] **Step 4: 생성**

Run: `node scripts/gen-ribbon.mjs excel && node scripts/gen-ribbon.mjs --check excel`
Expected: `만듦` 다음 `그대로`.

- [ ] **Step 5: 상황별 탭 연결** — `src/client/grid-editor.ts`의 `setChartTabVisible: (visible) => shell.setContextualTab('chart', visible),` 를 아래로 바꾼다:

```ts
    setChartTabVisible: (visible) => {
      // 오피스처럼 차트를 고르면 "차트 디자인"·"서식" 두 탭
      shell.setContextualTab('chart', visible);
      shell.setContextualTab('chart-format', visible);
    },
```

- [ ] **Step 6: 확인**

Run: `npx vitest run tests/ribbonPages.test.ts tests/ribbonInventory.test.ts && npx tsc --noEmit && npm run build:editor`
Expected: PASS(엑셀 바꾸기 전 id 69개 모두 있음). `git diff --stat public/editor-grid.html`에서 판 밖(찾기 줄·정렬/조건부 서식 대화상자·수식 줄·시트 탭) 줄이 그대로인지.

- [ ] **Step 7: 커밋**

```bash
git add public/editor-grid.html src/client/grid-editor.ts tests/ribbonPages.test.ts
git diff --cached --stat
git commit -m "오피스 리본 구조 3: 엑셀 — M365 탭 9(홈·삽입·그리기·페이지 레이아웃·수식·데이터·검토·보기)·상황별 4, 기존 단추 63개를 오피스 자리로(id·처리기 그대로), 없는 기능은 준비 중, 차트를 고르면 차트 디자인·서식 두 탭

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 5: 파워포인트 — 전체 탭·묶음·단추

**Files:**
- Modify: `public/editor-slides.html`, `src/client/editor-slides.ts`(import 1줄, 변수 1줄, `initOfficeShell` 호출 1곳, focus 처리기 1줄, `renderCanvasForSlideInner` 첫 줄), `tests/ribbonPages.test.ts`(`DONE`에 `'powerpoint'`)

**Interfaces:**
- Consumes: Task 2의 생성기·마크업 약속, 목록 `apps.powerpoint`
- Produces: 목록과 1:1인 `public/editor-slides.html`(탭 11 + 상황별 8), 텍스트 상자·도형을 클릭하면 `shape-format`(도형 서식) 탭, 슬라이드를 바꾸면 숨김

- [ ] **Step 1: 실패하는 시험** — `DONE`을 `['word', 'excel', 'powerpoint']`로

Run: `npx vitest run tests/ribbonPages.test.ts`
Expected: FAIL — powerpoint `표시 없음`

- [ ] **Step 2: 표시 넣기** — `public/editor-slides.html`
  - 탭: 워드와 같게
  - 판: `<!-- 프레젠테이션 프로그램 "홈" 탭과 같은 사용 흐름의 리본 …` 주석 시작 줄 바로 위에 `    <!-- ribbon:panels:begin — scripts/gen-ribbon.mjs가 src/client/ribbon/inventory.json으로 만든다(손으로 고치지 말 것) -->`, `  </header>` 줄 바로 위에 `    <!-- ribbon:panels:end -->`. (판 id `slides-ribbon`, 묶음 id `ribbon-slides`·`format-toolbar`·`ribbon-paragraph`·`ribbon-draw`, 클래스 `slides-text-group`은 목록 `panelId`·`domId`·`className`으로 다시 나온다.)

- [ ] **Step 3: 기존 단추에 `data-rb` 달기·고치기** — 방법은 Task 3 Step 3과 같다.

| 키(`data-rb` 값) | 다는 요소 | 자리 | 고칠 것(convert·fix) |
|---|---|---|---|
| `home.slides.new-slide` | `#btn-add-blank` | same |  |
| `home.slides.duplicate` | `#btn-dup-slide` | nearest |  |
| `home.slides.delete` | `#btn-del-slide` | nearest |  |
| `home.font.font-family` | `#font-family` | same |  |
| `home.font.font-size` | `#font-size` | same | aria-label·title "글자 크기" → "글꼴 크기" |
| `home.font.grow` | `#btn-font-grow` | same |  |
| `home.font.shrink` | `#btn-font-shrink` | same |  |
| `home.font.clear-format` | `#btn-clear-format` | same |  |
| `home.font.bold` | `#btn-bold` | same |  |
| `home.font.italic` | `#btn-italic` | same |  |
| `home.font.underline` | `#btn-underline` | same |  |
| `home.font.font-color` | `#text-color-btn` | same |  |
| `home.paragraph.align-left` | `#btn-align-left` | same |  |
| `home.paragraph.align-center` | `#btn-align-center` | same |  |
| `home.paragraph.align-right` | `#btn-align-right` | same |  |
| `home.paragraph.justify` | `#btn-align-justify` | same |  |
| `home.drawing.text-box` | `#btn-add-textbox` | nearest |  |
| `home.drawing.rect` | `#btn-shape-rect` | nearest |  |
| `home.drawing.roundrect` | `#btn-shape-roundrect` | nearest |  |
| `home.drawing.ellipse` | `#btn-shape-ellipse` | nearest |  |
| `home.drawing.shape-fill` | `#shape-fill-btn` | same | input title·aria-label "도형 채우기 색" → "도형 채우기" |
| `insert.slides.new-slide` | (생성 — 대리 단추 btn-add-blank) | proxy |  |
| `insert.text.text-box` | (생성 — 대리 단추 btn-add-textbox) | proxy |  |
| `design.customize.background-colors` | `.slide-bg-swatches` | nearest |  |
| `design.customize.background-custom` | `label.slide-bg-custom` | nearest |  |
| `design.customize.apply-all` | `#btn-bg-apply-all` | nearest |  |
| `design.customize.reset-background` | `#btn-bg-reset` | nearest |  |
| `transitions.preview.preview` | `#btn-transition-preview` | same |  |
| `transitions.transition-to-this-slide.transition-gallery` | `.effect-gallery[aria-label="전환 효과"]` | same | 전환 갤러리 div에 aria-label="전환 효과" 더함 |
| `transitions.timing.duration` | `label.layout-field:has(> #transition-speed)` | nearest |  |
| `transitions.timing.apply-all` | `#btn-transition-apply-all` | same |  |
| `animations.preview.preview` | `#btn-animation-preview` | same |  |
| `animations.animation.animation-gallery` | `.effect-gallery[aria-label="애니메이션 효과"]` | same | 애니메이션 갤러리 div에 aria-label="애니메이션 효과" 더함 |
| `animations.timing.start` | `label.anim-trigger-field` | same |  |
| `animations.timing.target` | `#animation-target` | nearest | span#animation-target 에 aria-label="애니메이션 대상" 더함 |
| `shape-format.insert-shapes.text-box` | (생성 — 대리 단추 btn-add-textbox) | proxy |  |

주의: 디자인 탭 `label.slide-bg-custom`은 `<label class="color-btn slide-bg-custom" title="다른 색">`. 전환·애니메이션 갤러리는 `<div class="ribbon-group-buttons effect-gallery">` 두 개 — 먼저 `aria-label`(전환 쪽 `"전환 효과"`, 애니메이션 쪽 `"애니메이션 효과"`)을 더하고 `data-rb`를 첫 속성으로(`<div data-rb="transitions.transition-to-this-slide.transition-gallery" class="ribbon-group-buttons effect-gallery" aria-label="전환 효과">`). 생성 뒤 이 div는 묶음의 `.ribbon-group-buttons` 안에 들어가므로 클래스가 겹쳐도 CSS(`.effect-gallery`)는 그대로 맞는다. `#animation-target`은 `<span id="animation-target" class="animation-target">` → `<span data-rb="animations.timing.target" id="animation-target" class="animation-target" aria-label="애니메이션 대상">`.

- [ ] **Step 4: 생성**

Run: `node scripts/gen-ribbon.mjs powerpoint && node scripts/gen-ribbon.mjs --check`
Expected: `만듦 public/editor-slides.html`, 다음 줄은 세 페이지 모두 `그대로`.

- [ ] **Step 5: 상황별 탭 연결** — `src/client/editor-slides.ts`
  - `import { initOfficeShell } from './officeShell.js';` → `import { initOfficeShell, type OfficeShell } from './officeShell.js';`
  - `let lastFocusedShapeId: string | null = null;` 바로 아래에 `let shell: OfficeShell | null = null; // 도형 서식 상황별 탭`
  - `initOfficeShell({` 호출 앞에 `shell = ` 를 붙인다(`shell = initOfficeShell({ … });`)
  - 상자 focus 처리기의 `refreshRibbonState(view);` 바로 아래에 `shell?.setContextualTab('shape-format', canEdit);`
  - `function renderCanvasForSlideInner(entryId: string): void {` 바로 다음 줄에 `  shell?.setContextualTab('shape-format', false); // 슬라이드를 다시 그리면 고른 상자가 없다`

- [ ] **Step 6: 확인**

Run: `npx vitest run && npx tsc --noEmit && npm run build:editor && node scripts/gen-ribbon.mjs --check`
Expected: 전체 vitest PASS(PPT 바꾸기 전 id 46개 모두 있음), 타입·빌드 성공, 세 페이지 `그대로`.

- [ ] **Step 7: 커밋**

```bash
git add public/editor-slides.html src/client/editor-slides.ts tests/ribbonPages.test.ts
git diff --cached --stat
git commit -m "오피스 리본 구조 4: 파워포인트 — M365 탭 11(홈·삽입·그리기·디자인·전환·애니메이션·슬라이드 쇼·녹음/녹화·검토·보기)·상황별 8, 기존 단추 33개와 대리 3개를 오피스 자리로(id·처리기 그대로), 없는 기능은 준비 중, 상자를 고르면 도형 서식 탭

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 6: 준비 중 아이콘(단색 선 SVG 스프라이트)

**Files:**
- Create: `public/ribbon/icons.svg`, `tests/ribbonIcons.test.ts`

**Interfaces:**
- Consumes: Task 1의 `requiredIcons(inv)`(계획 작성 때 441개), Task 2의 `hydrateIcons`(`<use href="/ribbon/icons.svg#rbi-<키>">`), CSS `.rb-ico`(선 굵기·색을 여기서 물려줌)
- Produces: `public/ribbon/icons.svg` — `<symbol id="rbi-<키>" viewBox="0 0 24 24">` 441개, **96 KB 상한**(평균 220바이트), 서버 기본 정적 경로(`no-cache` + ETag)로 나가므로 한 번 받고 재검증만

**크기 예산과 늦게 받기:** 스프라이트는 파일 하나(≤ 96 KB, gzip 약 20 KB)다. 페이지 HTML에는 아이콘이 없고, 탭을 처음 고를 때 그 탭의 준비 중 단추에만 `<use>`가 붙는다(`hydrateIcons`) — 첫 화면(홈 탭)에서 스프라이트를 한 번 받고, 안 연 탭의 아이콘 요소는 만들지 않는다. 기존 단추의 인라인 아이콘은 그대로(설계 §3).

**그리는 규칙:** 24×24 칸, 선만(채움은 꼭 필요할 때 `fill="currentColor"` 한 조각), 색 지정 금지(색·굵기는 `.rb-ico`가 준다: 16px 1.5, 32px 1.25), 요소 3개 안팎, 좌표는 정수 또는 .5. 오피스 아이콘을 베끼지 않고 **뜻**을 그린다(설계 §3 "오피스 의미에 맞는 단색 선 아이콘"). 뜻이 같은 키는 다른 기호를 그대로 쓴다: `<symbol id="rbi-a" viewBox="0 0 24 24"><use href="#rbi-b"/></symbol>`(예: `bring-forward` 여러 탭, `alt-text` 여러 탭은 이미 키가 같아 한 번만).

- [ ] **Step 1: 실패하는 시험** — `tests/ribbonIcons.test.ts`

```ts
import fs from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadInventory, requiredIcons } from '../scripts/ribbonGenLib.mjs';

const SPRITE = 'public/ribbon/icons.svg';
const BUDGET_BYTES = 96 * 1024;

describe('리본 준비 중 아이콘 스프라이트(public/ribbon/icons.svg)', () => {
  const svg = fs.existsSync(SPRITE) ? fs.readFileSync(SPRITE, 'utf8') : '';
  const ids = [...svg.matchAll(/<symbol id="rbi-([a-z0-9-]+)"/g)].map((m) => m[1]);
  it(`크기 상한 ${BUDGET_BYTES} 바이트`, () => {
    expect(svg.length).toBeGreaterThan(0);
    expect(Buffer.byteLength(svg)).toBeLessThanOrEqual(BUDGET_BYTES);
  });
  it('목록이 쓰는 아이콘이 모두 있고 남는 것이 없다', () => {
    const want = requiredIcons(loadInventory());
    expect(want.filter((id) => !ids.includes(id))).toEqual([]);
    expect(ids.filter((id) => !want.includes(id))).toEqual([]);
  });
  it('id가 겹치지 않는다', () => expect(new Set(ids).size).toBe(ids.length));
  it('단색 선 — 24칸, 색·그림·글자·스크립트·바깥 주소 없음', () => {
    for (const m of svg.matchAll(/<symbol [^>]*>/g)) expect(m[0]).toMatch(/viewBox="0 0 24 24"/);
    expect(svg).not.toMatch(/#[0-9a-fA-F]{3,8}\b|rgb\(|<image|<text|<script|<style|<foreignObject|href="(?!#rbi-)/);
  });
});
```

Run: `npx vitest run tests/ribbonIcons.test.ts`
Expected: FAIL — 크기 0

- [ ] **Step 2: 그릴 키 목록 뽑기**

Run: `node --input-type=module -e "import { loadInventory, requiredIcons } from './scripts/ribbonGenLib.mjs'; console.log(requiredIcons(loadInventory()).join('\n'))"`
Expected: 441줄(Task 1 대조로 바뀌었으면 그 수). 키마다 목록에서 이름을 찾아(`grep -n '"key":"<키>"' src/client/ribbon/inventory.json`) 뜻을 확인한다.

- [ ] **Step 3: 스프라이트 쓰기** — `public/ribbon/icons.svg`, 키 알파벳 순, 한 줄에 기호 하나. 머리와 앞부분(나머지 키도 같은 모양으로 모두):

```xml
<svg xmlns="http://www.w3.org/2000/svg" style="display:none">
<!-- 리본 준비 중·대리·접힌 묶음 단추 아이콘(설계 2026-11-22-office-ribbon-structure §3) — 24칸 단색 선, 색·굵기는 CSS .rb-ico. 키 목록 = scripts/ribbonGenLib.mjs requiredIcons() -->
<symbol id="rbi-3d-models" viewBox="0 0 24 24"><path d="M12 3 20 7.5v9L12 21l-8-4.5v-9zM12 12l8-4.5M12 12v9M12 12 4 7.5"/></symbol>
<symbol id="rbi-accept" viewBox="0 0 24 24"><rect x="4" y="4" width="16" height="16" rx="2"/><path d="m8 12 3 3 5-6"/></symbol>
<symbol id="rbi-addins" viewBox="0 0 24 24"><path d="M5 5h6v3a2 2 0 1 0 4 0V5h4v14H5zM11 12h3a2 2 0 1 1 0 4h-3"/></symbol>
<symbol id="rbi-alt-text" viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="M7 15l2.5-6 2.5 6M8 13h3M14 9v6h3"/></symbol>
<symbol id="rbi-chart" viewBox="0 0 24 24"><path d="M4 20h16M7 17v-6M12 17V7M17 17v-9"/></symbol>
<symbol id="rbi-copy" viewBox="0 0 24 24"><rect x="8" y="8" width="12" height="13" rx="1.5"/><path d="M16 8V4.5A1.5 1.5 0 0 0 14.5 3h-9A1.5 1.5 0 0 0 4 4.5V16h4"/></symbol>
<symbol id="rbi-cut" viewBox="0 0 24 24"><circle cx="7" cy="17" r="3"/><circle cx="17" cy="17" r="3"/><path d="M9 15 18 3M15 15 6 3"/></symbol>
<symbol id="rbi-find" viewBox="0 0 24 24"><circle cx="10.5" cy="10.5" r="5.5"/><path d="m15 15 5 5"/></symbol>
<symbol id="rbi-icons" viewBox="0 0 24 24"><circle cx="8" cy="8" r="3.5"/><path d="M14 4.5h6v6h-6zM8 14l3.5 6h-7zM17 14v6M14 17h6"/></symbol>
<symbol id="rbi-macros" viewBox="0 0 24 24"><rect x="4" y="5" width="16" height="14" rx="1.5"/><path d="m8 10 2 2-2 2M12 14h4"/></symbol>
<symbol id="rbi-paste" viewBox="0 0 24 24"><rect x="5" y="5" width="14" height="16" rx="1.5"/><path d="M9 3h6v4H9zM9 12h6M9 16h6"/></symbol>
<symbol id="rbi-pictures" viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="14" rx="1.5"/><circle cx="9" cy="10" r="1.6"/><path d="m4 18 5-5 4 4 3-3 4 4"/></symbol>
<symbol id="rbi-shapes" viewBox="0 0 24 24"><circle cx="8.5" cy="8.5" r="4.5"/><rect x="11" y="11" width="9" height="9" rx="1"/></symbol>
<symbol id="rbi-themes" viewBox="0 0 24 24"><circle cx="12" cy="12" r="8"/><path d="M12 4a8 8 0 0 0 0 16z" fill="currentColor"/></symbol>
</svg>
```
(위 14개는 견본이자 실제 기호 — 나머지 427개를 같은 규칙으로 더한다. 아직 다 그리지 못한 키를 임시 도형으로 채우지 않는다: 시험이 빠진 키를 알려 주는 대로 끝까지 그린다.)

- [ ] **Step 4: 통과 확인**

Run: `npx vitest run tests/ribbonIcons.test.ts && node -e "console.log(require('fs').statSync('public/ribbon/icons.svg').size)"`
Expected: PASS 4개, 크기 ≤ 98304.

- [ ] **Step 5: 커밋**

```bash
git add public/ribbon/icons.svg tests/ribbonIcons.test.ts
git diff --cached --stat
git commit -m "오피스 리본 구조 5: 준비 중·대리·접힌 묶음 단추 아이콘 — 단색 선 SVG 스프라이트 하나(24칸, 96 KB 상한, 탭을 고를 때만 붙음), 목록이 쓰는 키와 1:1 시험

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 7: 목록 대조 e2e·기존 단언 갱신·전체 회귀·사람 확인

**Files:**
- Create: `scripts/verify-ribbon-inventory-e2e.mjs`
- Modify: `scripts/verify-readonly-editor-e2e.mjs`, `scripts/verify-ribbon-skin-e2e.mjs`, `package.json`(조건부), `src/client/ribbon/inventory.json`(폭 확인 결과로 `collapse`만 고칠 수 있음 — 고치면 `gen-ribbon` 다시)

**Interfaces:**
- Consumes: `loadInventory`, `APP_OF_EXT`, `controlKey`, `expectedStructure`, `refsOf`, `tabCount`, `pendingIn`, `launcherCount`(Task 1·2), 마크업 약속(Task 2), `?rbctx=1`
- Produces: `node scripts/verify-ribbon-inventory-e2e.mjs [8180~8189]` — 항목별 OK/FAIL, 실패 수만큼 종료 코드, `.audit-run/ribbon-inventory-shots/report.html`

- [ ] **Step 1: 새 e2e 쓰기** — `scripts/verify-ribbon-inventory-e2e.mjs`

```js
#!/usr/bin/env node
// 오피스 리본 목록 대조(설계 2026-11-22-office-ribbon-structure §4) — 세 편집기의 탭·묶음·단추가 src/client/ribbon/inventory.json과 1:1:
//  (1) 탭 이름·순서·편집 전용·상황별  (2) 묶음 이름·순서·접힘 순위·자세히  (3) 단추 data-rb·이름·크기·종류·준비 중 여부
//  (4) 준비 중은 꺼짐·말풍선 "준비 중 — 이름"  (5) 기존 기능 선택자가 하나씩, 그 단추 안에 있고 가려지지 않음(접힌 묶음은 펼쳐서)
//  (6) 연 탭마다 아이콘이 붙고 스프라이트에 모양이 있음  (7) 1920·1400·1280·1200·900·800px 모든 탭 가로 넘침 없음·접힘은 순위대로,
//      1200px에서 기존 시험이 누르는 홈 묶음은 펼친 채  (8) 600px 모든 묶음 접힘·가로 넘침 없음·펼친 판이 화면 안
//  (9) ui=view — 준비 중·접힌 묶음 단추에 .ro-disabled 없음  (10) gen-ribbon --check  + M365 캡처와 나란히 보고서
// 쓰는 법: npm run build:editor && node scripts/verify-ribbon-inventory-e2e.mjs [포트 8180~8189, 기본 8180]
import { execFileSync, spawn, spawnSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';
import { APP_OF_EXT, controlKey, expectedStructure, loadInventory, refsOf } from './ribbonGenLib.mjs';

const PORT = Number(process.argv[2] ?? 8180);
// 운영 포트(8000·8010·8020·8100·8110·8130·8148)와 다른 시험 범위를 피해 이 계획 범위만
if (!(PORT >= 8180 && PORT <= 8189)) throw new Error(`시험 포트는 8180~8189만 쓴다(받은 값 ${PORT})`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `ribbon-inventory-e2e-${PORT}`);
const SHOTS = path.resolve('.audit-run', 'ribbon-inventory-shots');
const M365 = path.resolve('.audit-run', 'm365-ribbon');
const FILES = {
  docx: 'audit/office/corpus/docx/a1-korean-report.docx',
  xlsx: 'audit/office/corpus/xlsx/x07-multisheet-print.xlsx',
  pptx: 'audit/office/corpus/pptx/a2-korean-deck.pptx',
};
const READY = { docx: '.ProseMirror', xlsx: '#grid-edit-table td[data-row]', pptx: '#slides-rail .slide-card' };
// 기존 e2e(verify-xlsx-view 1200px, browser-check 1280px, readonly 강제 클릭)가 누르는 홈 묶음 — 1200px에서 접히면 그 시험이 깨진다
const LEGACY_OPEN = { docx: ['clipboard', 'font', 'paragraph'], xlsx: ['clipboard', 'font', 'alignment', 'cells'], pptx: ['slides', 'font'] };
const WIDTHS = [1920, 1400, 1280, 1200, 900, 800];
const INV = loadInventory();
const results = [];
const rec = (name, ok, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

if (!fs.existsSync('public/js/editor.bundle.js')) {
  console.error('public/js/editor.bundle.js 없음 — npm run build:editor 먼저');
  process.exit(1);
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
  const pid = server?.pid;
  server = null;
  if (pid) {
    try {
      execFileSync('taskkill', ['/PID', String(pid), '/T', '/F'], { stdio: 'ignore' });
    } catch {
      // 이미 끝났다
    }
  }
}
process.on('SIGINT', () => {
  stopServer();
  process.exit(130);
});
process.on('exit', stopServer);

async function upload(file) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패 ${file}: ${JSON.stringify(r)}`);
  for (let i = 0; i < 1200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready') return st.id;
    if (st.status === 'failed') throw new Error(`변환 실패 ${file}: ${st.error}`);
    await sleep(250);
  }
  throw new Error(`대기 시간 초과 ${file}`);
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: 'ribbon-inventory-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);
async function open(context, url, ready) {
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.goto(`${BASE}${url}`);
  await page.waitForSelector(ready, { timeout: 90000 });
  await page.waitForFunction(() => !document.querySelector('#conn-status') || document.querySelector('#conn-status').getAttribute('data-status') === 'synced', null, { timeout: 30000 }).catch(() => {});
  await sleep(800);
  return { page, errors };
}
const withCtx = (url) => `${url}${url.includes('?') ? '&' : '?'}rbctx=1`;

/** 브라우저 안 — DOM을 목록과 같은 모양으로 읽는다(이름·종류·크기 규칙은 계획 Task 7에 적힌 대로) */
function readRibbon(refs) {
  const text = (el) => (el?.textContent ?? '').replace(/\s+/g, ' ').trim();
  const clean = (s) => s.replace(/\s*▾\s*$/, '').trim();
  const nameOf = (unit, ref) => {
    if (unit.matches('label')) {
      const s = [...unit.querySelectorAll(':scope > span')].map(text).find(Boolean);
      if (s) return clean(s);
    }
    const lab = ref.querySelector('.ribbon-btn-label');
    if (lab) return clean(text(lab));
    if (ref.classList.contains('icon-btn-big') || ref.classList.contains('rb-gallery')) {
      const s = ref.querySelector(':scope > span');
      if (s && text(s)) return clean(text(s));
    }
    const aria = ref.getAttribute('aria-label');
    if (aria) return clean(aria);
    return clean((ref.getAttribute('title') ?? '').split(/ \(| — /)[0]);
  };
  const kindOf = (unit, ref) => {
    if (unit.dataset.kind) return unit.dataset.kind;
    if (unit.matches('.style-gallery, .effect-gallery, .slide-bg-swatches')) return 'gallery';
    if (ref.matches('select')) return 'select';
    if (ref.matches('input[type="checkbox"]')) return 'toggle';
    if (ref.matches('input[type="color"]')) return 'split';
    if (ref.matches('input')) return 'input';
    if (unit.matches('.dropdown-caret') || unit.querySelector('.dropdown-caret')) return ref.matches('.dropdown-caret') ? 'dropdown' : 'split';
    if (!ref.matches('button')) return 'info';
    return 'button';
  };
  const tabs = [...document.querySelectorAll('.ribbon-tabs .ribbon-tab[data-tab]')].map((t) => ({ id: t.dataset.tab, name: text(t), editOnly: t.hasAttribute('data-edit-only'), contextual: t.classList.contains('ribbon-tab-contextual') }));
  const panels = {};
  for (const p of document.querySelectorAll('#toolbar .ribbon-panel[data-panel]')) {
    panels[p.dataset.panel] = [...p.querySelectorAll(':scope > .ribbon-group')].map((g) => {
      const l = g.querySelector(':scope > .ribbon-launcher');
      return {
        id: g.dataset.group ?? '',
        name: text(g.querySelector(':scope > .ribbon-group-label')),
        collapse: Number(g.dataset.collapse),
        launcher: l ? (l.hasAttribute('data-pending') ? 'pending' : `proxy:${l.dataset.proxy}`) : null,
        controls: [...g.querySelectorAll('[data-rb]')].map((unit) => {
          const key = unit.dataset.rb;
          const ref = (refs[key] && document.querySelector(refs[key])) || unit;
          const kind = kindOf(unit, ref);
          return { key, name: nameOf(unit, ref), size: kind === 'gallery' || ref.classList.contains('icon-btn-big') ? 'big' : 'small', kind, pending: unit.hasAttribute('data-pending') };
        }),
      };
    });
  }
  return { tabs, panels };
}
/** 같은 모양 둘의 차이를 경로와 함께(처음 20개) */
function diff(want, got, at = '', out = []) {
  if (out.length >= 20) return out;
  if (Array.isArray(want)) {
    if (!Array.isArray(got) || got.length !== want.length) out.push(`${at} 길이 목록 ${want.length} · 화면 ${Array.isArray(got) ? got.length : typeof got}${Array.isArray(got) ? ` (${JSON.stringify(got.map((x) => x.key ?? x.id ?? x)).slice(0, 300)})` : ''}`);
    want.forEach((w, i) => diff(w, got?.[i], `${at}[${i}${w?.key ? ` ${w.key}` : w?.id ? ` ${w.id}` : ''}]`, out));
  } else if (want && typeof want === 'object') {
    for (const k of Object.keys(want)) diff(want[k], got?.[k], `${at}.${k}`, out);
  } else if (want !== got) out.push(`${at} 목록 ${JSON.stringify(want)} · 화면 ${JSON.stringify(got)}`);
  return out;
}
/** 판 하나의 접힘 상태 — 넘침, 접힌 묶음이 순위 앞쪽부터인지 */
const fitState = (page, tabId) =>
  page.evaluate((t) => {
    const p = document.querySelector(`.ribbon-panel[data-panel="${t}"]`);
    const groups = [...p.querySelectorAll(':scope > .ribbon-group[data-collapse]')].map((g, index) => ({ id: g.dataset.group, priority: Number(g.dataset.collapse), index, collapsed: g.classList.contains('rb-collapsed') }));
    const order = [...groups].sort((a, b) => a.priority - b.priority || b.index - a.index);
    const k = groups.filter((g) => g.collapsed).length;
    const prefix = order.slice(0, k).every((g) => g.collapsed);
    return { over: p.scrollWidth - p.clientWidth, collapsed: groups.filter((g) => g.collapsed).map((g) => g.id), prefix, doc: document.documentElement.scrollWidth };
  }, tabId);

function writeReport(rows) {
  const body = rows
    .map(([ours, ref]) => {
      const has = ref && fs.existsSync(path.join(M365, ref));
      if (has) fs.copyFileSync(path.join(M365, ref), path.join(SHOTS, `m365-${ref}`));
      return `<tr><td><img src="${ours}" alt="${ours}"><p>${ours}</p></td><td>${has ? `<img src="m365-${ref}" alt="${ref}">` : '(M365 캡처 없음 — Task 1 덤프)'}<p>${ref ?? ''}</p></td></tr>`;
    })
    .join('');
  fs.writeFileSync(path.join(SHOTS, 'report.html'), `<!doctype html><html lang="ko"><meta charset="utf-8"><title>리본 구조 비교</title><style>body{font-family:"맑은 고딕",sans-serif}img{width:900px;border:1px solid #ccc}td{vertical-align:top;padding:8px}</style><h1>우리 리본(1920px, ?rbctx=1)과 M365 나란히</h1><table><tr><th>우리</th><th>M365</th></tr>${body}</table></html>`);
}

const check = spawnSync(process.execPath, ['scripts/gen-ribbon.mjs', '--check'], { encoding: 'utf8' });
rec('(10) gen-ribbon --check — 세 페이지가 목록과 같음', check.status === 0, (check.stderr || '').trim().split('\n').slice(0, 3).join(' | '));

await startServer();
fs.rmSync(SHOTS, { recursive: true, force: true });
fs.mkdirSync(SHOTS, { recursive: true });
const report = [];
let browser;
try {
  browser = await chromium.launch({ ignoreDefaultArgs: ['--hide-scrollbars'] });
  for (const [ext, file] of Object.entries(FILES)) {
    const app = INV.apps[APP_OF_EXT[ext]];
    const refs = refsOf(app);
    const id = await upload(file);
    const editorUrl = (await token(id, { canEdit: true })).editorUrl;
    const ctx = await browser.newContext({ viewport: { width: 1920, height: 1000 } });
    const { page, errors } = await open(ctx, withCtx(editorUrl), READY[ext]);

    // (1)~(3) 구조 1:1
    const want = expectedStructure(app);
    const got = await page.evaluate(readRibbon, refs);
    const tabDiff = diff(want.tabs, got.tabs, 'tabs');
    rec(`(1) ${ext} 탭 ${want.tabs.length}개 — 이름·순서·편집 전용·상황별`, tabDiff.length === 0, tabDiff.join(' | '));
    const panelDiff = diff(want.panels, got.panels, 'panels');
    rec(`(2)(3) ${ext} 묶음·단추 — 이름·순서·접힘 순위·자세히·크기·종류·준비 중`, panelDiff.length === 0 && Object.keys(got.panels).length === Object.keys(want.panels).length, panelDiff.join(' | '));

    // (4)(5) 준비 중 꺼짐, 기존 선택자
    const units = await page.evaluate((r) => {
      const badPending = [...document.querySelectorAll('#toolbar [data-pending]')]
        .filter((u) => !(u.matches(':disabled') || u.querySelector(':disabled') || u.classList.contains('rb-gallery')) || !(u.getAttribute('title') ?? '').startsWith('준비 중 — '))
        .map((u) => u.dataset.rb ?? u.getAttribute('aria-label'));
      const badRefs = Object.entries(r)
        .filter(([key, sel]) => {
          const all = document.querySelectorAll(sel);
          const unit = document.querySelector(`[data-rb="${key}"]`);
          return all.length !== 1 || !unit || !(unit === all[0] || unit.contains(all[0])) || unit.hasAttribute('data-pending');
        })
        .map(([key, sel]) => `${key} ${sel}`);
      return { pending: document.querySelectorAll('#toolbar [data-pending]').length, badPending, badRefs };
    }, refs);
    rec(`(4) ${ext} 준비 중 ${units.pending}개 모두 꺼짐·말풍선 "준비 중 — 이름"`, units.pending > 0 && units.badPending.length === 0, units.badPending.slice(0, 5).join(', '));
    rec(`(5) ${ext} 기존 기능 ${Object.keys(refs).length}개 — 선택자가 하나씩, 그 단추(data-rb) 안`, units.badRefs.length === 0, units.badRefs.slice(0, 5).join(' | '));

    // (5) 눌림 — 탭을 하나씩 고르고, 접힌 묶음은 펼쳐서 가운데가 그 단추에 닿는지 + (6) 아이콘
    const covered = [];
    for (const tab of [...app.tabs, ...app.contextual]) {
      await page.click(`.ribbon-tab[data-tab="${tab.id}"]`);
      await sleep(150);
      await page.locator('#toolbar').screenshot({ path: path.join(SHOTS, `${ext}-${tab.id}.png`) });
      report.push([`${ext}-${tab.id}.png`, `${APP_OF_EXT[ext]}-${tab.name}.png`.replace(/[\\/:*?"<>|]/g, '_')]);
      for (const g of tab.groups) {
        for (const c of g.controls) {
          if (c.pending || c.kind === 'info') continue;
          const groupSel = `.ribbon-panel[data-panel="${tab.id}"] .ribbon-group[data-group="${g.id}"]`;
          const folded = await page.$eval(groupSel, (el) => el.classList.contains('rb-collapsed'));
          if (folded) await page.click(`${groupSel} > .rb-group-toggle`);
          const hit = await page.$eval(refs[controlKey(tab.id, g.id, c)], (el) => {
            const r = el.getBoundingClientRect();
            if (!r.width || !r.height) return 'hidden';
            const at = document.elementFromPoint(r.left + r.width / 2, r.top + r.height / 2);
            return at && (at === el || el.contains(at) || at.contains(el)) ? 'ok' : `가림 ${at?.className ?? at?.tagName}`;
          });
          if (hit !== 'ok') covered.push(`${controlKey(tab.id, g.id, c)} ${hit}`);
          if (folded) await page.keyboard.press('Escape');
        }
      }
    }
    await page.click('.ribbon-tab[data-tab="home"]');
    rec(`(5) ${ext} 기존 기능이 모두 보이고 가려지지 않음(접힌 묶음은 펼쳐서)`, covered.length === 0, covered.slice(0, 5).join(' | '));
    const icons = await page.evaluate(() => {
      const els = [...document.querySelectorAll('#toolbar .ribbon-panel [data-icon]')];
      const missing = els.filter((el) => !el.querySelector(':scope > svg.rb-ico > use')?.getAttribute('href')?.endsWith(`#rbi-${el.dataset.icon}`)).map((el) => el.dataset.rb ?? el.dataset.icon);
      return { count: els.length, missing, used: [...new Set(els.map((el) => el.dataset.icon))] };
    });
    const sprite = await (await fetch(`${BASE}/ribbon/icons.svg`)).text();
    const noShape = icons.used.filter((k) => !sprite.includes(`id="rbi-${k}"`));
    rec(`(6) ${ext} 연 탭의 아이콘 ${icons.count}개가 붙고 스프라이트에 모양이 있음`, icons.missing.length === 0 && noShape.length === 0, JSON.stringify({ missing: icons.missing.slice(0, 5), noShape: noShape.slice(0, 5) }));
    rec(`${ext} 화면 오류 없음`, errors.length === 0, errors.slice(0, 2).join(' | '));

    // (7) 폭별 접힘
    for (const width of WIDTHS) {
      await page.setViewportSize({ width, height: 1000 });
      await sleep(300);
      const bad = [];
      for (const tab of app.tabs) {
        await page.click(`.ribbon-tab[data-tab="${tab.id}"]`);
        await sleep(80);
        const s = await fitState(page, tab.id);
        if (s.over > 1 || !s.prefix || s.doc > width) bad.push(`${tab.id} ${JSON.stringify(s)}`);
        if (width === 1200 && tab.id === 'home') {
          const closed = LEGACY_OPEN[ext].filter((g) => s.collapsed.includes(g));
          if (closed.length) bad.push(`1200px 홈에서 기존 시험이 누르는 묶음이 접힘: ${closed.join(',')} — 목록 collapse를 높일 것`);
          console.log(`     ${ext} 1200px 홈 접힘: ${s.collapsed.join(',') || '없음'}`);
        }
      }
      await page.click('.ribbon-tab[data-tab="home"]');
      rec(`(7) ${ext} ${width}px — 기본 탭 ${app.tabs.length}개 가로 넘침 없음·접힘은 순위 순서`, bad.length === 0, bad.join(' | '));
    }
    await ctx.close();

    // (8) 600px
    const narrow = await browser.newContext({ viewport: { width: 600, height: 900 } });
    const n = await open(narrow, editorUrl, READY[ext]);
    const ns = await n.page.evaluate(() => {
      const groups = [...document.querySelectorAll('.ribbon-panel:not([hidden]) > .ribbon-group[data-collapse]')];
      return { all: groups.length, folded: groups.filter((g) => g.classList.contains('rb-collapsed')).length, doc: document.documentElement.scrollWidth };
    });
    await n.page.click('.ribbon-panel:not([hidden]) > .ribbon-group[data-collapse] > .rb-group-toggle');
    const pop = await n.page.evaluate(() => {
      const b = document.querySelector('.ribbon-group.rb-open > .ribbon-group-buttons');
      if (!b) return null;
      const r = b.getBoundingClientRect();
      return { left: Math.round(r.left), right: Math.round(r.right), w: Math.round(r.width) };
    });
    rec(`(8) ${ext} 600px — 홈 묶음 ${ns.all}개 모두 접힘·가로 넘침 없음·펼친 판이 화면 안`, ns.all > 0 && ns.folded === ns.all && ns.doc <= 600 && !!pop && pop.left >= 0 && pop.right <= 600 && pop.w > 0, JSON.stringify({ ...ns, pop }));
    await narrow.close();

    // (9) ui=view
    const ro = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const v = await open(ro, (await token(id, { engine: 'editor' })).viewerUrl, READY[ext]);
    const vs = await v.page.evaluate(() => ({
      body: document.body.classList.contains('ui-view'),
      pendingGrey: [...document.querySelectorAll('#toolbar [data-pending], #toolbar [data-pending] *')].filter((el) => el.classList.contains('ro-disabled')).length,
      toggleGrey: document.querySelectorAll('#toolbar .rb-group-toggle.ro-disabled').length,
      greyed: document.querySelectorAll('#toolbar .ro-disabled').length,
    }));
    rec(`(9) ${ext} ui=view — 쓰는 단추 흐림 ${vs.greyed}개, 준비 중·접힌 묶음 단추엔 흐림 없음`, vs.body && vs.greyed > 0 && vs.pendingGrey === 0 && vs.toggleGrey === 0, JSON.stringify(vs));
    await ro.close();
  }
  writeReport(report);
  console.log(`\n나란히 보기: ${path.join(SHOTS, 'report.html')}`);
} catch (error) {
  rec('스크립트 오류', false, String(error?.stack ?? error));
} finally {
  await browser?.close();
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${results.length - failed}/${results.length} 통과`);
process.exit(failed);
```

DOM 읽기 규칙(위 `readRibbon`, 목록 `name`·`size`·`kind`가 이것과 맞아야 한다):
- 이름: 단추가 `label`이면 안의 첫 글자 있는 `span` → `.ribbon-btn-label` → 큰 단추(`icon-btn-big`)·빈 갤러리의 바로 아래 `span` → `aria-label` → `title`의 ` (`·` — ` 앞. 끝의 `▾`는 뗀다.
- 종류: `data-kind`(준비 중·대리) → 갤러리 클래스 → select/확인란/색 입력(분할)/입력 → 안에 `.dropdown-caret`가 있으면 누르는 요소 자신이 caret이면 dropdown 아니면 split → 단추가 아니면 info → button.
- 크기: 갤러리이거나 누르는 요소가 `icon-btn-big`이면 big.

- [ ] **Step 2: 기존 e2e 단언을 목록에서 계산하게** — `scripts/verify-readonly-editor-e2e.mjs`
  - import 줄들 끝에 `import { APP_OF_EXT, loadInventory, pendingIn, tabCount } from './ribbonGenLib.mjs';`
  - 포트 줄 `if (!(PORT >= 8160 && PORT <= 8169)) throw new Error(\`시험 포트는 8160~8169만 쓴다(받은 값 ${PORT})\`);` → `if (!((PORT >= 8160 && PORT <= 8169) || (PORT >= 8180 && PORT <= 8189))) throw new Error(\`시험 포트는 8160~8169·8180~8189만 쓴다(받은 값 ${PORT})\`);`
  - `const MIN_TABS = { docx: 6, xlsx: 6, pptx: 6 };` → 
```js
// 탭 수는 오피스 리본 목록에서(설계 2026-11-22-office-ribbon-structure — 사용자 지시로 탭이 늘었다)
const INV = loadInventory();
const MIN_TABS = Object.fromEntries(Object.entries(APP_OF_EXT).map(([ext, name]) => [ext, tabCount(INV.apps[name])]));
```
  - `(1) … 탭 눌러 보기` 블록의 `page.evaluate`와 `rec`을 아래로(지금 줄: `const panel = await page.evaluate((t) => { … return { visible: …, controls: controls.length, greyed: … }; }, tab);` 와 그 아래 `const wantActive …`, `rec(…)`):
```js
      const panel = await page.evaluate((t) => {
        const el = document.querySelector(`.ribbon-panel[data-panel="${t}"]`);
        const controls = [...el.querySelectorAll('button, select, input')].filter((c) => c.getClientRects().length > 0);
        // 준비 중(disabled가 우선)과 접힌 묶음 단추(둘러보기용)는 흐림 대상이 아니다 — 설계 2026-11-22-office-ribbon-structure §1-3
        const pending = controls.filter((c) => c.closest('[data-pending]')).length;
        const toggles = controls.filter((c) => c.classList.contains('rb-group-toggle')).length;
        return { visible: !el.hidden && el.getClientRects().length > 0, controls: controls.length, greyed: controls.filter((c) => c.classList.contains('ro-disabled')).length, pending, toggles, pendingAll: el.querySelectorAll('[data-pending]').length };
      }, tab);
      const wantActive = ext === 'pptx' ? 1 : 0; // 전환 미리 보기
      const wantPending = pendingIn(INV.apps[APP_OF_EXT[ext]], tab);
      rec(`(1) ${ext} "${tab}" 탭 눌러 보기 — 판이 열리고 쓰는 단추는 모두 흐림, 준비 중 ${wantPending}개(목록 기준)는 꺼진 채`, panel.visible && panel.controls > 0 && panel.greyed === panel.controls - wantActive - panel.pending - panel.toggles && panel.pendingAll === wantPending, JSON.stringify({ ...panel, wantPending }));
```

  `scripts/verify-ribbon-skin-e2e.mjs`:
  - import 줄들 끝에 `import { APP_OF_EXT, launcherCount, loadInventory } from './ribbonGenLib.mjs';`
  - 포트 줄을 위와 같이 8180~8189까지(메시지 포함)
  - `const LAUNCHERS = { docx: 1, xlsx: 1, pptx: 0 };` → 
```js
// "자세히" 수는 오피스 리본 목록에서(상황별 탭 판 포함, 준비 중 포함)
const INV = loadInventory();
const LAUNCHERS = Object.fromEntries(Object.entries(APP_OF_EXT).map(([ext, name]) => [ext, launcherCount(INV.apps[name])]));
```
  - (5) 두 줄(`const launchers = await edit.page.$$eval('.ribbon-launcher', …)` 와 그 `rec`)을 아래로:
```js
    const launchers = await edit.page.$$eval('.ribbon-launcher', (ls) => ls.map((l) => ({ proxy: l.dataset.proxy ?? '', pending: l.hasAttribute('data-pending'), ok: !l.id && (l.hasAttribute('data-pending') ? l.disabled : !!document.getElementById(l.dataset.proxy ?? '')) })));
    rec(`(5) ${ext} "자세히" ${LAUNCHERS[ext]}개(목록 기준) — 준비 중은 꺼짐, 나머지는 기존 단추를 가리킴, id 없음`, launchers.length === LAUNCHERS[ext] && launchers.every((l) => l.ok), JSON.stringify(launchers.filter((l) => !l.ok)));
```
  - 파일 머리 주석의 `(5) "자세히"는 있는 단추를 가리킴(워드 1·엑셀 1·PPT 0)` → `(5) "자세히" 수는 목록 기준, 준비 중은 꺼짐·나머지는 있는 단추를 가리킴`

- [ ] **Step 3: 단위 시험·빌드**

Run: `npx vitest run && npx tsc --noEmit && npm run build:editor && node scripts/gen-ribbon.mjs --check`
Expected: 전부 PASS / 성공 / `그대로` 셋.

- [ ] **Step 4: 새 e2e 돌리기**

Run: `node scripts/verify-ribbon-inventory-e2e.mjs 8180`
Expected: 모든 항목 `OK`, 마지막 `N/N 통과`. 실패 처리:
- (2)(3) 이름·종류·크기 차이 → 기존 단추면 마크업(Task 3~5 표의 `fix`·`convert`)이나 목록 `kind`·`size`를 실제와 맞춘다(오피스 쪽이 맞으면 마크업을 고침). 준비 중이면 생성기가 만든 것이라 목록·생성기 쪽 문제.
- (7) 1200px 홈 묶음 접힘 → 목록에서 그 묶음 `collapse`를 올리거나 다른 묶음(오른쪽·덜 쓰는 것)을 낮추고 `node scripts/gen-ribbon.mjs` → 다시. 넘침이 남으면(모두 접어도 넘침) 그 탭 묶음 수가 너무 많은 것 — 보고서에 적고 사용자에게 묻는다.
- 고쳤으면 Step 3부터 다시.

- [ ] **Step 5: 회귀 묶음** — 시험 서버(8181)를 PowerShell로 띄우고(PID 저장), 서버를 받는 e2e를 돌린 뒤, 그 PID만 끈다.

```powershell
$run = '.audit-run/ribbon-structure-8181'; Remove-Item -Recurse -Force $run -ErrorAction SilentlyContinue; New-Item -ItemType Directory -Force "$run/tmp" | Out-Null
$env:DOCVIEW_PORT = '8181'; $env:DOCVIEW_HOST = '127.0.0.1'; $env:DOCVIEW_SECRET = [guid]::NewGuid().ToString('N'); $env:DOCVIEW_DATA_DIR = "$run/data"; $env:TMP = "$run/tmp"; $env:TEMP = "$run/tmp"; $env:DOCVIEW_LO_PROFILE_ROOT = "$run/lo"
$p = Start-Process node -ArgumentList 'node_modules/tsx/dist/cli.mjs', 'src/server.ts' -RedirectStandardOutput "$run/server.log" -RedirectStandardError "$run/server.err" -PassThru -WindowStyle Hidden
$p.Id | Set-Content "$run/pid"; "PID $($p.Id)"
```
그다음(각 명령 따로, 서버 준비는 `curl -s http://127.0.0.1:8181/api/formats`가 응답할 때까지):
```bash
node scripts/verify-ribbon-skin-e2e.mjs 8182
node scripts/verify-readonly-editor-e2e.mjs 8183
node scripts/verify-content-only-e2e.mjs http://127.0.0.1:8181
node scripts/verify-view-shell-e2e.mjs http://127.0.0.1:8181
node scripts/verify-docx-edit-e2e.mjs http://127.0.0.1:8181
node scripts/verify-xlsx-fidelity-e2e.mjs http://127.0.0.1:8181
node scripts/verify-pptx-fidelity-e2e.mjs http://127.0.0.1:8181
node scripts/verify-xlsx-view.mjs http://127.0.0.1:8181
```
끝나면: `taskkill /PID <저장한 번호> /T /F` (PowerShell: `taskkill /PID (Get-Content .audit-run/ribbon-structure-8181/pid) /T /F`).
Expected: 모두 실패 0. `verify:ribbon-skin`은 62px·900/800px·600px·흐림 0.45·아이콘 32/16·한글 라벨·"자세히" 수(목록 기준) 그대로 통과해야 한다. `verify-xlsx-view`(1200px에서 `#btn-add-row`·`#btn-del-row` 누름)·`browser-check`류가 접힌 묶음 때문에 실패하면 Step 4의 (7) 규칙대로 `collapse`를 고친다(시험을 고치지 않는다). `verify-artcraft`(8150~8159 전용, 아트크래프트 앱 화면 — 리본 마크업 무관)는 이 계획에서 돌리지 않는다.

- [ ] **Step 6: 사람 확인(실제 브라우저, 사용자 규칙 "입력 컨트롤 전부 조작")** — 8181 서버에서 세 문서를 편집 토큰으로 열어(메인 화면 `http://127.0.0.1:8181/` 업로드 → 동시편집 열기) Playwright `headless: false` 또는 크롬으로 직접:
  1. 세 편집기 모든 기본 탭을 차례로 눌러 묶음 이름·순서가 `report.html`의 M365 캡처와 같은지 나란히 본다(`.audit-run/ribbon-inventory-shots/report.html`). 다른 곳을 목록 키와 함께 적는다.
  2. 준비 중 단추 아무거나 5개에 마우스를 올려 "준비 중 — <이름>" 말풍선, 눌러도 아무 일 없음.
  3. 기존 기능은 옮긴 자리에서 실제로 눌러 본다: 워드 — 붙여넣기·굵게·글꼴/크기 고르기·대소문자·강조 색·줄 간격·스타일 갤러리 한 칸·찾기·바꾸기·선택(모두 선택)·삽입›페이지 나누기·줄 바꿈·표(격자)·메모(대리)·여백/용지 방향/크기·들여쓰기/간격 숫자 칸·검토›맞춤법 및 문법 검사·단어 개수·새 메모·메모 표시·검토 창·보기›100%·페이지 너비·확대·축소, 표 안에 커서 → 테이블 디자인·레이아웃 탭이 나타나고 위/아래/왼쪽/오른쪽에 삽입·표/행/열 삭제·셀 병합/분할. 엑셀 — 홈 전부(테두리 메뉴·채우기 색·글꼴 색·맞춤 여섯·자동 줄 바꿈·병합 메뉴·표시 형식 고르기·회계·백분율·쉼표·자릿수·조건부 서식 메뉴·표 서식·행/열 삽입·삭제·자동 합계 메뉴·지우기(서식)·정렬 및 필터 메뉴·찾기 및 선택), 삽입›표·세로/가로 막대·꺾은선·원형(차트를 넣어 차트 디자인·서식 탭이 나타남, 종류 넷·차트 제목 입력·삭제), 수식›자동 합계·평균·최대값·숫자 개수·최소값, 데이터›오름차순/내림차순 정렬·정렬·필터·지우기, 보기›100%·확대·축소·틀 고정. PPT — 새 슬라이드·복제·삭제·글꼴/크기·굵게·기울임·밑줄·글꼴 색·모든 서식 지우기·맞춤 넷·텍스트 상자·사각형·둥근 사각형·타원·도형 채우기, 상자를 클릭 → 도형 서식 탭, 그 탭의 텍스트 상자(대리), 삽입›새 슬라이드·텍스트 상자(대리), 디자인›배경 색 견본·직접 고르기·모두 적용·배경 원래대로, 전환›미리 보기·효과 한 칸·속도·모두 적용, 애니메이션›미리 보기·효과 한 칸·시작.
  4. 창 폭을 1920 → 1200 → 900 → 800 → 600으로 줄이며 묶음이 오른쪽·덜 쓰는 것부터 "묶음 이름 ▾" 단추로 접히는지, 접힌 단추를 눌러 떠 있는 판에서 단추를 눌러 보고(드롭다운 메뉴도 열어 봄) Esc·바깥 클릭으로 닫히는지.
  5. Ctrl+F1(리본 축소) → 탭 눌러 잠깐 펼치기 → 접힌 묶음도 정상.
  6. ui=view(메인 화면 "열기")에서 쓰는 단추는 흐리고, 준비 중은 원래 꺼진 모양 그대로, 접힌 묶음 단추는 눌러 펼쳐 둘러볼 수 있음. `ui=content`(콘텐츠만 보기)에서 리본·탭이 하나도 안 보임.
  7. `?rbctx=1`을 붙여 상황별 탭(머리글/바닥글·그림 서식·SmartArt 디자인 등 연결 없는 것)까지 모두 M365 캡처와 나란히.
  사람 확인에서 고친 것이 있으면 Step 3·4를 다시 돌린다. 끝나면 8181 서버 PID를 끈다.

- [ ] **Step 7: `package.json`(조건부)**

Run: `git diff --quiet package.json && echo clean || echo dirty`
- `clean`이면 `"scripts"`의 `"verify:ribbon-skin": …` 줄 아래에 두 줄:
```json
    "gen:ribbon": "node scripts/gen-ribbon.mjs",
    "verify:ribbon-inventory": "node scripts/verify-ribbon-inventory-e2e.mjs",
```
- `dirty`면 더하지 않는다(다른 에이전트 변경을 같이 커밋하지 않게) — 보고서에 적는다.

- [ ] **Step 8: 커밋**

```bash
git add scripts/verify-ribbon-inventory-e2e.mjs scripts/verify-readonly-editor-e2e.mjs scripts/verify-ribbon-skin-e2e.mjs src/client/ribbon/inventory.json public/editor.html public/editor-grid.html public/editor-slides.html
# Step 7이 clean이었을 때만: git add package.json
git diff --cached --stat
git commit -m "오피스 리본 구조 6: 목록 대조 e2e(verify-ribbon-inventory — 탭·묶음·단추 1:1, 준비 중 꺼짐, 기존 기능 눌림, 폭별 접힘 순서, 600px, ui=view, M365 나란히 보고서), readonly·ribbon-skin의 탭·자세히·흐림 수를 목록에서 계산, 회귀 묶음 통과

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

## 자체 점검(계획 작성자가 한 것)

- **설계 대조:** §1-1 구조 → 목록(부록 A)·Task 1 대조·Task 3~5. §1-2 기존 자리·준비 중 → `map`·`pendingHtml`·`ribbon-ids-before` 시험. §1-3 ui=view → Task 2 Step 7(d)·Task 7 (9)·readonly 단언. §1-4 ui=content → 건드리지 않음, Task 7 Step 6-6에서 확인. §1-5 view.html → 범위 밖(손대는 파일 없음). §2 목록 표 → 부록 A + `inventory.json`(함께 커밋). 상황별 탭 → 목록 `contextual`(`trigger`·`wired`), 연결은 Task 3·4·5 Step 5. §3 하지 않는 것 → 기능 구현 없음, 백스테이지는 목록에만, 아이콘은 준비 중만 새로(Task 6). §4 검사 → Task 7 Step 1(1:1)·Step 2(단언 갱신)·Step 5(기존 e2e·62px·600px·흐림·ribbon-skin)·Step 6(M365 나란히).
- **빈칸 점검:** 단추 441개 아이콘 중 427개는 그림을 계획에 다 적지 않았다 — 대신 키 목록을 뽑는 명령·그리는 규칙·견본 14개·빠진 키를 알려 주는 시험을 두었다(그림 자체가 이 태스크의 일). 그 밖의 단계는 코드·명령·기대 결과가 다 있다.
- **이름 일치:** `controlKey`·`expectedStructure`·`refsOf`·`pendingIn`·`launcherCount`·`tabCount`·`requiredIcons`·`generatePage`(lib), `hydrateIcons`·`fitPanel`·`collapseOrder`·`installGroupToggles`·`closeGroupPopup`(runtime)이 정의한 태스크와 쓰는 태스크에서 같다. 마크업 클래스 `rb-group-toggle`·`rb-collapsed`·`rb-open`·`rb-pending`·`rb-gallery`·`rb-check`·`rb-field`·`rb-select`·`rb-ico`·`rb-rows`, 속성 `data-rb`·`data-pending`·`data-kind`·`data-icon`·`data-group`·`data-collapse`가 생성기·CSS·runtime·e2e에서 같다.
- **계획 작성 중 확인한 것:** 이 목록(`inventory.json`)은 `validateInventory` 위반 0, `serializeInventory` 왕복 같음. 생성기(Part A+B와 같은 코드)를 현재 세 HTML 사본에 `data-rb`를 단 채로 돌려 기존 단추 64·63·33개를 모두 찾고, 두 번 돌려도 같은 결과(결정적)임을 확인했다.

---

## 부록 A — 전체 목록 표(계획 작성 시점 `inventory.json`에서 뽑음 — Task 1 대조 뒤에는 JSON이 진짜)

표기: ●기존 기능 ◇대리 단추(다른 탭의 기존 단추를 대신 누름) ○준비 중 · `[L]` 큰 단추 `[S]` 작은 단추 · `▾` 드롭다운 `+▾` 분할 · 갤러리·확인란·입력·목록·글. 묶음 옆 괄호는 접힘 순위(작을수록 먼저 접힘)와 "자세히". 드롭다운 고정 항목·자리 결정 이유는 JSON의 `items`·`note`.

### A.1 워드 — 기존 64 · 대리 1 · 준비 중 259 (준비 중 자세히 16)

파일(백스테이지, 화면은 기존 그대로): 오피스 홈 · 새로 만들기 · 열기 · 정보 · 저장 · 다른 이름으로 저장 · 인쇄 · 공유 · 내보내기 · 변환 · 닫기 · 계정 · 피드백 · 옵션 / 지금 있는 것 정보 · 다운로드 · 인쇄 · 닫기

**탭 「홈」** (`home`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 클립보드 (9 · 자세히(준비 중: 클립보드 창)) | ●붙여넣기[L] · ●잘라내기[S] · ●복사[S] · ●서식 복사[S] |
| 글꼴 (8 · 자세히(준비 중: 글꼴)) | ●글꼴[S 목록] · ●글꼴 크기[S 목록] · ●글꼴 크기 크게[S] · ●글꼴 크기 작게[S] · ●대/소문자 바꾸기[S ▾] · ●모든 서식 지우기[S] · ○윗주 달기[S] · ○문자 테두리[S] · ●굵게[S] · ●기울임꼴[S] · ●밑줄[S] · ●취소선[S] · ●아래 첨자[S] · ●위 첨자[S] · ○텍스트 효과 및 타이포그래피[S ▾] · ●텍스트 강조 색[S +▾] · ●글꼴 색[S +▾] · ○문자 음영[S] · ○원 문자[S ▾] |
| 단락 (7 · 자세히(준비 중: 단락)) | ●글머리 기호[S] · ●번호 매기기[S] · ●다단계 목록[S] · ●내어쓰기[S] · ●들여쓰기[S] · ○정렬[S] · ○편집 기호 표시/숨기기[S] · ●왼쪽 맞춤[S] · ●가운데 맞춤[S] · ●오른쪽 맞춤[S] · ●양쪽 맞춤[S] · ○균등 분할[S] · ●줄 및 단락 간격[S ▾] · ●음영[S +▾] · ○테두리[S +▾] |
| 스타일 (4 · 자세히(준비 중: 스타일 창)) | ●스타일[L 갤러리] |
| 편집 (5) | ●찾기[S] · ●바꾸기[S] · ●선택[S] |
| 음성 (2) | ○받아쓰기[L +▾] |
| 편집기 (1) | ○편집기[L] |

**탭 「삽입」** (`insert`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 페이지 (3) | ○표지[S ▾] · ○빈 페이지[S] · ●페이지 나누기[S] · ●줄 바꿈[S] |
| 표 (8) | ●표[L ▾] |
| 일러스트레이션 (6) | ○그림[L ▾] · ○도형[L ▾] · ○아이콘[L] · ○3D 모델[L +▾] · ○스마트아트[S] · ○차트[S] · ○스크린샷[S ▾] |
| 추가 기능 (2) | ○추가 기능[L] |
| 미디어 (2) | ○온라인 비디오[L] |
| 링크 (4) | ○링크[L +▾] · ○책갈피[S] · ○상호 참조[S] |
| 메모 (5) | ◇메모[L] |
| 머리글/바닥글 (5) | ○머리글[L ▾] · ○바닥글[L ▾] · ○페이지 번호[L ▾] |
| 텍스트 (4) | ○텍스트 상자[L ▾] · ○빠른 문서 요소[S ▾] · ○워드아트[S ▾] · ○단락의 첫 문자 장식[S ▾] · ○서명란[S +▾] · ○날짜 및 시간[S] · ○개체[S +▾] |
| 기호 (3) | ○수식[L +▾] · ○기호[L ▾] |

**탭 「그리기」** (`draw`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 그리기 도구 (6) | ○선택[L ▾] · ○올가미 선택[L] · ○지우개[L ▾] · ○펜[L 갤러리] · ○펜 추가[L ▾] |
| 변환 (4) | ○잉크를 도형으로[L] · ○잉크를 수식으로[L] |
| 재생 (2) | ○잉크 재생[L] |

**탭 「디자인」** (`design`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 문서 서식 (5) | ○테마[L ▾] · ○스타일 세트[L 갤러리] · ○색[L ▾] · ○글꼴[L ▾] · ○단락 간격[S ▾] · ○효과[S ▾] · ○기본값으로 설정[S] |
| 페이지 배경 (6) | ○워터마크[L ▾] · ○페이지 색[L ▾] · ○페이지 테두리[L] |

**탭 「레이아웃」** (`layout`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 페이지 설정 (8 · 자세히(준비 중: 페이지 설정)) | ○텍스트 방향[L ▾] · ●여백[L ▾] · ●용지 방향[L ▾] · ●크기[L ▾] · ○단[L ▾] · ○나누기[S ▾] · ○줄 번호[S ▾] · ○하이픈 넣기[S ▾] |
| 원고지 (2) | ○원고지 설정[L] |
| 단락 (7 · 자세히(준비 중: 단락)) | ●왼쪽[S 입력] · ●오른쪽[S 입력] · ●앞[S 입력] · ●뒤[S 입력] |
| 정렬 (4) | ○위치[L ▾] · ○텍스트 줄 바꿈[L ▾] · ○앞으로 가져오기[S +▾] · ○뒤로 보내기[S +▾] · ○선택 창[S] · ○맞춤[S ▾] · ○그룹화[S ▾] · ○회전[S ▾] |

**탭 「참조」** (`references`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 목차 (5) | ○목차[L ▾] · ○텍스트 추가[S ▾] · ○목차 업데이트[S] |
| 각주 (5 · 자세히(준비 중: 각주 및 미주)) | ○각주 삽입[L] · ○미주 삽입[S] · ○다음 각주[S +▾] · ○주석 표시[S] |
| 조사 (2) | ○검색[L] |
| 인용 및 참고 문헌 (4) | ○인용 삽입[L ▾] · ○원본 관리[S] · ○스타일[S 목록] · ○참고 문헌[S ▾] |
| 캡션 (3) | ○캡션 삽입[L] · ○그림 목차 삽입[S] · ○표 업데이트[S] · ○상호 참조[S] |
| 색인 (3) | ○항목 표시[L] · ○색인 삽입[S] · ○색인 업데이트[S] |
| 인용 목차 (2) | ○인용 표시[L] · ○인용 목차 삽입[S] · ○인용 목차 업데이트[S] |

**탭 「편지」** (`mailings`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 만들기 (3) | ○봉투[L] · ○레이블[L] |
| 편지 병합 시작 (5) | ○편지 병합 시작[L ▾] · ○받는 사람 선택[L ▾] · ○받는 사람 목록 편집[L] |
| 필드 쓰기 및 삽입 (4) | ○병합 필드 강조 표시[L] · ○주소 블록[L] · ○인사말[L] · ○병합 필드 삽입[L +▾] · ○규칙[S ▾] · ○필드 일치[S] · ○레이블 업데이트[S] |
| 결과 미리 보기 (3) | ○결과 미리 보기[L] · ○처음 레코드[S] · ○이전 레코드[S] · ○레코드[S 입력] · ○다음 레코드[S] · ○마지막 레코드[S] · ○받는 사람 찾기[S] · ○오류 자동 확인[S] |
| 마침 (6) | ○완료 및 병합[L ▾] |

**탭 「검토」** (`review`)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 언어 교정 (6) | ●맞춤법 및 문법 검사[L] · ○동의어 사전[S] · ●단어 개수[S] |
| 음성 (2) | ○소리 내어 읽기[L] |
| 접근성 (2) | ○접근성 검사[L +▾] |
| 언어 (3) | ○번역[L ▾] · ○언어[L ▾] |
| 메모 (7) | ●새 메모[L] · ●삭제[S] · ●이전[S] · ●다음[S] · ●메모 표시[L] |
| 추적 (5 · 자세히(준비 중: 변경 내용 추적 옵션)) | ○변경 내용 추적[L +▾] · ○검토할 표시[S 목록] · ○태그 표시[S ▾] · ●검토 창[S] |
| 변경 내용 (4) | ○적용[L +▾] · ○거부[S +▾] · ○이전[S] · ○다음[S] |
| 비교 (3) | ○비교[L ▾] |
| 보호 (3) | ○작성자 차단[L] · ○편집 제한[L] |
| 잉크 (1) | ○잉크 숨기기[L ▾] |

**탭 「보기」** (`view`)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 보기 (6) | ○읽기 모드[L] · ○인쇄 모양[L] · ○웹 모양[L] · ○개요[S] · ○초안[S] |
| 몰입형 (3) | ○포커스[L] · ○몰입형 리더[L] |
| 페이지 이동 (3) | ○세로[L] · ○나란히[L] |
| 표시 (4) | ○눈금자[S 확인란] · ○눈금선[S 확인란] · ○탐색 창[S 확인란] |
| 확대/축소 (7) | ○확대/축소[L] · ●100%[L] · ○한 페이지[S] · ○여러 페이지[S] · ●페이지 너비[S] · ●확대[S] · ●축소[S] |
| 창 (5) | ○새 창[L] · ○모두 정렬[L] · ○나누기[L] · ○나란히 보기[S] · ○동시 스크롤[S] · ○창 위치 다시 설정[S] · ○창 전환[L ▾] |
| 매크로 (2) | ○매크로[L +▾] |
| 셰어포인트 (1) | ○속성[L] |

**상황별 탭 「테이블 디자인」** (`table-design`, 편집 전용, 나타나는 때: 표 안에 커서 — 연결: editor.ts dispatchTransaction의 isInTable(newState) — setContextualTab('table-design', …) 한 줄 더함)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 표 스타일 옵션 (3) | ○머리글 행[S 확인란] · ○첫째 열[S 확인란] · ○요약 행[S 확인란] · ○마지막 열[S 확인란] · ○줄무늬 행[S 확인란] · ○줄무늬 열[S 확인란] |
| 표 스타일 (5) | ○표 스타일[L 갤러리] · ○음영[L +▾] |
| 테두리 (4 · 자세히(준비 중: 테두리 및 음영)) | ○테두리 스타일[L ▾] · ○선 스타일[S 목록] · ○펜 두께[S 목록] · ○펜 색[S +▾] · ○테두리[L +▾] · ○테두리 복사[L] |

**상황별 탭 「레이아웃」** (`table`, 편집 전용, 나타나는 때: 표 안에 커서 — 연결: 기존 — editor.ts setContextualTab('table', canEdit && isInTable(newState)))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 표 (3) | ○선택[S ▾] · ○눈금선 보기[S] · ○속성[S] |
| 그리기 (2) | ○표 그리기[L] · ○지우개[L] |
| 행 및 열 (8 · 자세히(준비 중: 셀 삽입)) | ●표 삭제[S] · ●행 삭제[S] · ●열 삭제[S] · ●위에 삽입[L] · ●아래에 삽입[L] · ●왼쪽에 삽입[L] · ●오른쪽에 삽입[L] |
| 병합 (7) | ●셀 병합[S] · ●셀 분할[S] · ○표 분할[S] |
| 셀 크기 (4 · 자세히(준비 중: 표 속성)) | ○자동 맞춤[L ▾] · ○높이[S 입력] · ○너비[S 입력] · ○행 높이를 같게[S] · ○열 너비를 같게[S] |
| 맞춤 (3) | ○위쪽 왼쪽 맞춤[S] · ○위쪽 가운데 맞춤[S] · ○위쪽 오른쪽 맞춤[S] · ○가운데 왼쪽 맞춤[S] · ○가운데 맞춤[S] · ○가운데 오른쪽 맞춤[S] · ○아래쪽 왼쪽 맞춤[S] · ○아래쪽 가운데 맞춤[S] · ○아래쪽 오른쪽 맞춤[S] · ○텍스트 방향[L] · ○셀 여백[L] |
| 데이터 (2) | ○정렬[L] · ○제목 행 반복[S] · ○텍스트로 변환[S] · ○수식[S] |

**상황별 탭 「머리글/바닥글」** (`header-footer`, 편집 전용, 나타나는 때: 머리글·바닥글 편집 중(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 머리글/바닥글 (5) | ○머리글[L ▾] · ○바닥글[L ▾] · ○페이지 번호[L ▾] |
| 삽입 (4) | ○날짜 및 시간[L] · ○문서 정보[L ▾] · ○빠른 문서 요소[L ▾] · ○그림[L] · ○온라인 그림[L] |
| 탐색 (3) | ○머리글로 이동[L] · ○바닥글로 이동[L] · ○이전[S] · ○다음[S] · ○이전 머리글에 연결[S] |
| 옵션 (3) | ○첫 페이지를 다르게 지정[S 확인란] · ○홀수/짝수 페이지를 다르게 지정[S 확인란] · ○문서 텍스트 표시[S 확인란] |
| 위치 (2) | ○위쪽 머리글[S 입력] · ○아래쪽 바닥글[S 입력] · ○맞춤 탭 삽입[S] |
| 닫기 (9) | ○머리글/바닥글 닫기[L] |

**상황별 탭 「그림 서식」** (`picture-format`, 편집 전용, 나타나는 때: 그림 선택(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 조정 (6) | ○배경 제거[L] · ○수정[L ▾] · ○색[L ▾] · ○꾸밈 효과[L ▾] · ○투명도[L ▾] · ○그림 압축[S] · ○그림 바꾸기[S ▾] · ○그림 원래대로[S +▾] |
| 그림 스타일 (5 · 자세히(준비 중: 그림 서식)) | ○그림 스타일[L 갤러리] · ○그림 테두리[S +▾] · ○그림 효과[S ▾] · ○그림 레이아웃[S ▾] |
| 접근성 (2) | ○대체 텍스트[L] |
| 정렬 (3) | ○앞으로 가져오기[S +▾] · ○뒤로 보내기[S +▾] · ○선택 창[S] · ○맞춤[S ▾] · ○그룹화[S ▾] · ○회전[S ▾] |
| 크기 (4 · 자세히(준비 중: 크기 및 위치)) | ○자르기[L +▾] · ○높이[S 입력] · ○너비[S 입력] |

**상황별 탭 「도형 서식」** (`shape-format`, 편집 전용, 나타나는 때: 도형·텍스트 상자 선택(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 도형 삽입 (3) | ○도형[L 갤러리] · ○도형 편집[S ▾] · ○텍스트 상자 그리기[S +▾] |
| 도형 스타일 (5 · 자세히(준비 중: 도형 서식)) | ○도형 스타일[L 갤러리] · ○도형 채우기[S +▾] · ○도형 윤곽선[S +▾] · ○도형 효과[S ▾] |
| 워드아트 스타일 (4 · 자세히(준비 중: 텍스트 효과 서식)) | ○빠른 스타일[L 갤러리] · ○텍스트 채우기[S +▾] · ○텍스트 윤곽선[S +▾] · ○텍스트 효과[S ▾] |
| 텍스트 (3) | ○텍스트 방향[S ▾] · ○텍스트 맞춤[S ▾] · ○연결 만들기[S] |
| 접근성 (2) | ○대체 텍스트[L] |
| 정렬 (3) | ○위치[L ▾] · ○텍스트 줄 바꿈[L ▾] · ○앞으로 가져오기[S +▾] · ○뒤로 보내기[S +▾] · ○선택 창[S] · ○맞춤[S ▾] · ○그룹화[S ▾] · ○회전[S ▾] |
| 크기 (4 · 자세히(준비 중: 크기 및 위치)) | ○높이[S 입력] · ○너비[S 입력] |

### A.2 엑셀 — 기존 63 · 대리 0 · 준비 중 235 (준비 중 자세히 14)

파일(백스테이지, 화면은 기존 그대로): 오피스 홈 · 새로 만들기 · 열기 · 정보 · 저장 · 다른 이름으로 저장 · 인쇄 · 공유 · 내보내기 · 게시 · 닫기 · 계정 · 피드백 · 옵션 / 지금 있는 것 정보 · 다운로드 · 인쇄 · 닫기

**탭 「홈」** (`home`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 클립보드 (9 · 자세히(준비 중: 클립보드 창)) | ●붙여넣기[L] · ●잘라내기[S] · ●복사[S] · ○서식 복사[S] |
| 글꼴 (8 · 자세히(준비 중: 셀 서식: 글꼴)) | ●글꼴[S 목록] · ●글꼴 크기[S 입력] · ●글꼴 크기 크게[S] · ●글꼴 크기 작게[S] · ●굵게[S] · ●기울임꼴[S] · ●밑줄[S] · ●테두리[S +▾] · ●채우기 색[S +▾] · ●글꼴 색[S +▾] · ●취소선[S] · ○윗주 필드 표시/숨기기[S +▾] |
| 맞춤 (7 · 자세히(준비 중: 셀 서식: 맞춤)) | ●위쪽 맞춤[S] · ●가운데 맞춤[S] · ●아래쪽 맞춤[S] · ○방향[S ▾] · ●자동 줄 바꿈[S] · ●왼쪽 맞춤[S] · ●가운데 맞춤[S] · ●오른쪽 맞춤[S] · ○내어쓰기[S] · ○들여쓰기[S] · ●병합하고 가운데 맞춤[S +▾] |
| 표시 형식 (5 · 자세히(준비 중: 셀 서식: 표시 형식)) | ●표시 형식[S 목록] · ●회계 표시 형식[S] · ●백분율 스타일[S] · ●쉼표 스타일[S] · ●자릿수 늘림[S] · ●자릿수 줄임[S] |
| 스타일 (2) | ●조건부 서식[L ▾] · ●표 서식[L ▾] · ○셀 스타일[L ▾] |
| 셀 (6) | ●행 삽입[S] · ●열 삽입[S] · ●행 삭제[S] · ●열 삭제[S] · ○서식[L ▾] |
| 편집 (3) | ●자동 합계[S +▾] · ○채우기[S ▾] · ●지우기[S] · ●정렬 및 필터[L ▾] · ●찾기 및 선택[L] |
| 분석 (1) | ○데이터 분석[L] |

**탭 「삽입」** (`insert`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 표 (6) | ○피벗 테이블[L +▾] · ○추천 피벗 테이블[L] · ●표[L ▾] |
| 일러스트레이션 (4) | ○그림[L ▾] · ○도형[S ▾] · ○아이콘[S] · ○3D 모델[S +▾] · ○스마트아트[S] · ○스크린샷[S ▾] |
| 추가 기능 (2) | ○추가 기능[L] |
| 차트 (8 · 자세히(준비 중: 모든 차트 보기)) | ○추천 차트[L] · ●세로 막대형[S] · ●가로 막대형[S] · ●꺾은선형[S] · ●원형[S] · ○계층 구조 차트[S ▾] · ○통계 차트[S ▾] · ○분산형 또는 거품형 차트[S ▾] · ○폭포, 깔때기형, 주식형, 표면형 또는 방사형 차트[S ▾] · ○콤보 차트[S ▾] · ○지도[L ▾] · ○피벗 차트[L +▾] |
| 투어 (1) | ○3D 맵[L ▾] |
| 스파크라인 (3) | ○꺾은선형[L] · ○열[L] · ○승패[L] |
| 필터 (3) | ○슬라이서[L] · ○시간 표시 막대[L] |
| 링크 (3) | ○링크[L] |
| 메모 (4) | ○메모[L] |
| 텍스트 (2) | ○텍스트 상자[L] · ○머리글/바닥글[L] · ○워드아트[S ▾] · ○서명란[S +▾] · ○개체[S] |
| 기호 (2) | ○수식[L +▾] · ○기호[L] |

**탭 「그리기」** (`draw`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 그리기 도구 (6) | ○선택[L ▾] · ○올가미 선택[L] · ○지우개[L ▾] · ○펜[L 갤러리] · ○펜 추가[L ▾] |
| 변환 (4) | ○잉크를 도형으로[L] · ○잉크를 수식으로[L] |
| 재생 (2) | ○잉크 재생[L] |

**탭 「페이지 레이아웃」** (`page-layout`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 테마 (3) | ○테마[L ▾] · ○색[S ▾] · ○글꼴[S ▾] · ○효과[S ▾] |
| 페이지 설정 (7 · 자세히(준비 중: 페이지 설정)) | ○여백[L ▾] · ○용지 방향[L ▾] · ○크기[L ▾] · ○인쇄 영역[L ▾] · ○나누기[L ▾] · ○배경[L] · ○인쇄 제목[L] |
| 크기 조정 (5 · 자세히(준비 중: 페이지 설정: 페이지)) | ○너비[S 목록] · ○높이[S 목록] · ○배율[S 입력] |
| 시트 옵션 (4 · 자세히(준비 중: 페이지 설정: 시트)) | ○보기[S 확인란] · ○인쇄[S 확인란] · ○보기[S 확인란] · ○인쇄[S 확인란] |
| 정렬 (2) | ○앞으로 가져오기[S +▾] · ○뒤로 보내기[S +▾] · ○선택 창[S] · ○맞춤[S ▾] · ○그룹화[S ▾] · ○회전[S ▾] |

**탭 「수식」** (`formulas`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 함수 라이브러리 (8) | ○함수 삽입[L] · ●자동 합계[L +▾] · ●평균[S] · ●최대값[S] · ●숫자 개수[S] · ●최소값[S] · ○최근 사용 항목[L ▾] · ○재무[L ▾] · ○논리[L ▾] · ○텍스트[L ▾] · ○날짜 및 시간[L ▾] · ○찾기/참조 영역[L ▾] · ○수학/삼각[L ▾] · ○함수 더 보기[L ▾] |
| 정의된 이름 (4) | ○이름 관리자[L] · ○이름 정의[S +▾] · ○수식에서 사용[S ▾] · ○선택 영역에서 만들기[S] |
| 수식 분석 (3) | ○참조되는 셀 추적[S] · ○참조하는 셀 추적[S] · ○연결선 제거[S +▾] · ○수식 표시[S] · ○오류 검사[S +▾] · ○수식 계산[S] · ○조사식 창[L] |
| 계산 (5) | ○계산 옵션[L ▾] · ○지금 계산[S] · ○시트 계산[S] |

**탭 「데이터」** (`data`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 데이터 가져오기 및 변환 (4) | ○데이터 가져오기[L ▾] · ○텍스트/CSV에서[S] · ○웹[S] · ○테이블/범위에서[S] · ○그림에서[S ▾] · ○최근에 사용한 원본[S] · ○기존 연결[S] |
| 쿼리 및 연결 (3) | ○모두 새로 고침[L +▾] · ○쿼리 및 연결[S] · ○속성[S] · ○통합 문서 링크[S] |
| 데이터 형식 (2) | ○데이터 형식[L 갤러리] |
| 정렬 및 필터 (8) | ●오름차순 정렬[S] · ●내림차순 정렬[S] · ●정렬[L] · ●필터[L] · ●지우기[S] · ○다시 적용[S] · ○고급[S] |
| 데이터 도구 (5) | ○텍스트 나누기[L] · ○빠른 채우기[S] · ○중복된 항목 제거[S] · ○데이터 유효성 검사[S +▾] · ○통합[S] · ○관계[S] · ○데이터 모델 관리[S] |
| 예측 (4) | ○가상 분석[L ▾] · ○예측 시트[L] |
| 개요 (3 · 자세히(준비 중: 설정)) | ○그룹[L +▾] · ○그룹 해제[L +▾] · ○부분합[L] |

**탭 「검토」** (`review`)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 언어 교정 (5) | ○맞춤법 검사[L] · ○동의어 사전[S] · ○통합 문서 통계[S] |
| 접근성 (2) | ○접근성 검사[L +▾] |
| 정보 활용 (1) | ○스마트 조회[L] |
| 언어 (2) | ○번역[L] |
| 변경 내용 (3) | ○변경 내용 표시[L] |
| 메모 (5) | ○새 메모[L] · ○삭제[L] · ○이전 메모[L] · ○다음 메모[L] · ○메모 표시[L] |
| 노트 (3) | ○노트[L ▾] |
| 보호 (4) | ○시트 보호[L] · ○통합 문서 보호[L] · ○범위 편집 허용[L] · ○통합 문서 공유 취소[L] |
| 잉크 (1) | ○잉크 숨기기[L ▾] |

**탭 「보기」** (`view`)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 시트 보기 (3) | ○시트 보기[S 목록] · ○유지[S] · ○종료[S] · ○새로 만들기[S] · ○옵션[S] |
| 통합 문서 보기 (4) | ○기본[L] · ○페이지 나누기 미리 보기[L] · ○페이지 레이아웃[L] · ○사용자 지정 보기[L] |
| 표시 (3) | ○눈금자[S 확인란] · ○수식 입력줄[S 확인란] · ○눈금선[S 확인란] · ○머리글[S 확인란] |
| 확대/축소 (6) | ○확대/축소[L] · ●100%[L] · ○선택 영역 확대/축소[L] · ●확대[S] · ●축소[S] |
| 창 (6) | ○새 창[S] · ○모두 정렬[S] · ●틀 고정[S] · ○나누기[S] · ○숨기기[S] · ○숨기기 취소[S] · ○나란히 보기[S] · ○동시 스크롤[S] · ○창 위치 다시 설정[S] · ○창 전환[L ▾] |
| 매크로 (2) | ○매크로[L +▾] |

**상황별 탭 「차트 디자인」** (`chart`, 편집 전용, 나타나는 때: 넣은 차트를 고를 때 — 연결: 기존 — grid-editor.ts setChartTabVisible → setContextualTab('chart', …))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 차트 레이아웃 (6) | ○차트 요소 추가[L ▾] · ○빠른 레이아웃[L ▾] · ●차트 제목[S 입력] |
| 차트 스타일 (5) | ○색 변경[L ▾] · ○차트 스타일[L 갤러리] |
| 데이터 (4) | ○행/열 전환[L] · ○데이터 선택[L] |
| 종류 (7) | ●세로 막대형[L] · ●가로 막대형[L] · ●꺾은선형[L] · ●원형[L] |
| 위치 (3) | ○차트 이동[L] · ●삭제[L] |

**상황별 탭 「서식」** (`chart-format`, 편집 전용, 나타나는 때: 넣은 차트를 고를 때 — 연결: grid-editor.ts setChartTabVisible에 setContextualTab('chart-format', visible) 한 줄 더함)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 현재 선택 영역 (4) | ○차트 요소[S 목록] · ○선택 영역 서식[S] · ○스타일에 맞게 다시 설정[S] |
| 도형 삽입 (3) | ○도형[L 갤러리] · ○도형 모양 변경[S ▾] |
| 도형 스타일 (5 · 자세히(준비 중: 도형 서식)) | ○도형 스타일[L 갤러리] · ○도형 채우기[S +▾] · ○도형 윤곽선[S +▾] · ○도형 효과[S ▾] |
| 워드아트 스타일 (4 · 자세히(준비 중: 텍스트 효과 서식)) | ○빠른 스타일[L 갤러리] · ○텍스트 채우기[S +▾] · ○텍스트 윤곽선[S +▾] · ○텍스트 효과[S ▾] |
| 접근성 (2) | ○대체 텍스트[L] |
| 정렬 (3) | ○앞으로 가져오기[S +▾] · ○뒤로 보내기[S +▾] · ○선택 창[S] · ○맞춤[S ▾] · ○그룹화[S ▾] · ○회전[S ▾] |
| 크기 (4 · 자세히(준비 중: 크기 및 위치)) | ○높이[S 입력] · ○너비[S 입력] |

**상황별 탭 「테이블 디자인」** (`table-design`, 편집 전용, 나타나는 때: 표(테이블) 범위 안 칸(우리 "표 서식"은 서식만 입혀 표 개체가 없음 — 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 속성 (3) | ○표 이름[S 입력] · ○표 크기 조정[S] |
| 도구 (4) | ○피벗 테이블로 요약[S] · ○중복된 항목 제거[S] · ○범위로 변환[S] · ○슬라이서 삽입[L] |
| 외부 표 데이터 (2) | ○내보내기[L ▾] · ○새로 고침[L +▾] · ○속성[S] · ○브라우저에서 열기[S] · ○연결 끊기[S] |
| 표 스타일 옵션 (5) | ○머리글 행[S 확인란] · ○요약 행[S 확인란] · ○줄무늬 행[S 확인란] · ○첫째 열[S 확인란] · ○마지막 열[S 확인란] · ○줄무늬 열[S 확인란] · ○필터 단추[S 확인란] |
| 표 스타일 (6) | ○표 스타일[L 갤러리] |

**상황별 탭 「그림 서식」** (`picture-format`, 편집 전용, 나타나는 때: 그림 선택(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 조정 (6) | ○배경 제거[L] · ○수정[L ▾] · ○색[L ▾] · ○꾸밈 효과[L ▾] · ○투명도[L ▾] · ○그림 압축[S] · ○그림 바꾸기[S ▾] · ○그림 원래대로[S +▾] |
| 그림 스타일 (5 · 자세히(준비 중: 그림 서식)) | ○그림 스타일[L 갤러리] · ○그림 테두리[S +▾] · ○그림 효과[S ▾] · ○그림 레이아웃[S ▾] |
| 접근성 (2) | ○대체 텍스트[L] |
| 정렬 (3) | ○앞으로 가져오기[S +▾] · ○뒤로 보내기[S +▾] · ○선택 창[S] · ○맞춤[S ▾] · ○그룹화[S ▾] · ○회전[S ▾] |
| 크기 (4 · 자세히(준비 중: 크기 및 위치)) | ○자르기[L +▾] · ○높이[S 입력] · ○너비[S 입력] |

### A.3 파워포인트 — 기존 33 · 대리 3 · 준비 중 297 (준비 중 자세히 16)

파일(백스테이지, 화면은 기존 그대로): 오피스 홈 · 새로 만들기 · 열기 · 정보 · 저장 · 다른 이름으로 저장 · 인쇄 · 공유 · 내보내기 · 닫기 · 계정 · 피드백 · 옵션 / 지금 있는 것 정보 · 다운로드 · 인쇄 · 닫기

**탭 「홈」** (`home`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 클립보드 (9 · 자세히(준비 중: 클립보드 창)) | ○붙여넣기[L +▾] · ○잘라내기[S] · ○복사[S +▾] · ○서식 복사[S] |
| 슬라이드 (8) | ●새 슬라이드[L] · ○레이아웃[S ▾] · ○다시 설정[S] · ○구역[S ▾] · ●복제[S] · ●삭제[S] |
| 글꼴 (7 · 자세히(준비 중: 글꼴)) | ●글꼴[S 목록] · ●글꼴 크기[S 목록] · ●글꼴 크기 크게[S] · ●글꼴 크기 작게[S] · ●모든 서식 지우기[S] · ●굵게[S] · ●기울임꼴[S] · ●밑줄[S] · ○텍스트 그림자[S] · ○취소선[S] · ○문자 간격[S ▾] · ○대/소문자 바꾸기[S ▾] · ○텍스트 강조 색[S +▾] · ●글꼴 색[S +▾] |
| 단락 (6 · 자세히(준비 중: 단락)) | ○글머리 기호[S +▾] · ○번호 매기기[S +▾] · ○목록 수준 줄임[S] · ○목록 수준 늘림[S] · ○줄 간격[S ▾] · ●왼쪽 맞춤[S] · ●가운데 맞춤[S] · ●오른쪽 맞춤[S] · ●양쪽 맞춤[S] · ○균등 분할[S] · ○열 추가 또는 제거[S ▾] · ○텍스트 방향[S ▾] · ○텍스트 맞춤[S ▾] · ○스마트아트로 변환[S ▾] |
| 그리기 (5) | ●텍스트 상자[L] · ●사각형 삽입[S] · ●둥근 사각형 삽입[S] · ●타원 삽입[S] · ○정렬[L ▾] · ○빠른 스타일[L ▾] · ●도형 채우기[S +▾] · ○도형 윤곽선[S +▾] · ○도형 효과[S ▾] |
| 편집 (4) | ○찾기[S] · ○바꾸기[S +▾] · ○선택[S ▾] |
| 추가 기능 (2) | ○추가 기능[L] |
| 디자이너 (1) | ○디자이너[L] |

**탭 「삽입」** (`insert`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 슬라이드 (7) | ◇새 슬라이드[L] · ○슬라이드 다시 사용[L] |
| 표 (6) | ○표[L ▾] |
| 이미지 (6) | ○그림[L ▾] · ○스크린샷[L ▾] · ○사진 앨범[L +▾] |
| 일러스트레이션 (5) | ○도형[L ▾] · ○아이콘[L] · ○3D 모델[L +▾] · ○스마트아트[L] · ○차트[L] |
| 추가 기능 (2) | ○추가 기능[L] |
| 링크 (4) | ○확대/축소[L ▾] · ○링크[L] · ○실행[L] |
| 메모 (4) | ○메모[L] |
| 텍스트 (6) | ◇텍스트 상자[L] · ○머리글/바닥글[L] · ○워드아트[L ▾] · ○날짜 및 시간[S] · ○슬라이드 번호[S] · ○개체[S] |
| 기호 (3) | ○수식[L +▾] · ○기호[L] |
| 미디어 (3) | ○비디오[L ▾] · ○오디오[L ▾] · ○화면 녹화[L] |

**탭 「그리기」** (`draw`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 그리기 도구 (6) | ○선택[L ▾] · ○올가미 선택[L] · ○지우개[L ▾] · ○펜[L 갤러리] · ○펜 추가[L ▾] |
| 스텐실 (2) | ○눈금자[L] |
| 변환 (4) | ○잉크를 도형으로[L] · ○잉크를 수식으로[L] |
| 재생 (2) | ○잉크 재생[L] |

**탭 「디자인」** (`design`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 테마 (6) | ○테마[L 갤러리] |
| 적용 (5) | ○적용[L 갤러리] |
| 사용자 지정 (7) | ○슬라이드 크기[L ▾] · ●배경 색[L 갤러리] · ●배경 색 직접 고르기[S +▾] · ●모두 적용[S] · ●배경 원래대로[S] |
| 디자이너 (4) | ○디자이너[L] |

**탭 「전환」** (`transitions`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 미리 보기 (9) | ●미리 보기[L] |
| 이 슬라이드로 전환 (8) | ●전환 효과[L 갤러리] · ○효과 옵션[L ▾] |
| 타이밍 (6) | ○소리[S 목록] · ●속도[S 목록] · ●모두 적용[S] · ○마우스를 클릭할 때[S 확인란] · ○다음 시간 후[S 확인란] · ○시간[S 입력] |

**탭 「애니메이션」** (`animations`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 미리 보기 (9) | ●미리 보기[L] |
| 애니메이션 (8 · 자세히(준비 중: 효과 옵션)) | ●애니메이션 효과[L 갤러리] · ○효과 옵션[L ▾] |
| 고급 애니메이션 (5) | ○애니메이션 추가[L ▾] · ○애니메이션 창[S] · ○트리거[S ▾] · ○애니메이션 복사[S] |
| 타이밍 (6) | ●시작[S 목록] · ○재생 시간[S 입력] · ○지연[S 입력] · ●애니메이션 대상[S 글] · ○앞으로 이동[S] · ○뒤로 이동[S] |

**탭 「슬라이드 쇼」** (`slide-show`)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 슬라이드 쇼 시작 (5) | ○처음부터[L] · ○현재 슬라이드부터[L] · ○온라인 프레젠테이션[L ▾] · ○슬라이드 쇼 재구성[L ▾] |
| 설정 (4) | ○슬라이드 쇼 설정[L] · ○슬라이드 숨기기[L] · ○예행 연습[L] · ○녹화[L ▾] · ○내레이션 재생[S 확인란] · ○시간 사용[S 확인란] · ○미디어 컨트롤 표시[S 확인란] |
| 모니터 (2) | ○모니터[S 목록] · ○발표자 보기 사용[S 확인란] |

**탭 「녹음/녹화」** (`record`, 편집 전용)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 녹음/녹화 (5) | ○녹음/녹화[L +▾] |
| 콘텐츠 (4) | ○화면 녹화[L] · ○오디오[L ▾] |
| 저장 (3) | ○쇼로 저장[L] · ○비디오로 내보내기[L] |

**탭 「검토」** (`review`)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 언어 교정 (5) | ○맞춤법 검사[L] · ○동의어 사전[S] |
| 접근성 (2) | ○접근성 검사[L +▾] |
| 정보 활용 (1) | ○스마트 조회[L] |
| 언어 (2) | ○번역[L] · ○언어[L ▾] |
| 메모 (5) | ○새 메모[L] · ○삭제[L +▾] · ○이전[L] · ○다음[L] · ○메모 표시[L ▾] |
| 비교 (3) | ○비교[L] · ○적용[L +▾] · ○거부[L +▾] · ○이전[S] · ○다음[S] · ○검토 창[S] · ○검토 끝내기[L] |
| 잉크 (1) | ○잉크 숨기기[L ▾] |

**탭 「보기」** (`view`)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 프레젠테이션 보기 (6) | ○기본[L] · ○개요 보기[L] · ○여러 슬라이드[L] · ○슬라이드 노트[L] · ○읽기용 보기[L] |
| 마스터 보기 (4) | ○슬라이드 마스터[L] · ○유인물 마스터[L] · ○슬라이드 노트 마스터[L] |
| 표시 (4 · 자세히(준비 중: 눈금 및 안내선)) | ○눈금자[S 확인란] · ○눈금선[S 확인란] · ○안내선[S 확인란] · ○슬라이드 노트[L] |
| 확대/축소 (5) | ○확대/축소[L] · ○창에 맞춤[L] |
| 컬러/회색조 (3) | ○컬러[S] · ○회색조[S] · ○흑백[S] |
| 창 (4) | ○새 창[L] · ○모두 정렬[S] · ○계단식 배열[S] · ○창 나누기 이동[S] · ○창 전환[L ▾] |
| 매크로 (2) | ○매크로[L] |

**상황별 탭 「도형 서식」** (`shape-format`, 편집 전용, 나타나는 때: 텍스트 상자·도형을 클릭할 때 — 연결: editor-slides.ts 상자 focus 처리기 → setContextualTab('shape-format', canEdit), 슬라이드를 다시 그릴 때(renderCanvasForSlideInner) 숨김)

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 도형 삽입 (3) | ○도형[L 갤러리] · ○도형 편집[S ▾] · ◇텍스트 상자[S] · ○도형 병합[S ▾] |
| 도형 스타일 (5 · 자세히(준비 중: 도형 서식)) | ○도형 스타일[L 갤러리] · ○도형 채우기[S +▾] · ○도형 윤곽선[S +▾] · ○도형 효과[S ▾] |
| 워드아트 스타일 (4 · 자세히(준비 중: 텍스트 효과 서식)) | ○빠른 스타일[L 갤러리] · ○텍스트 채우기[S +▾] · ○텍스트 윤곽선[S +▾] · ○텍스트 효과[S ▾] |
| 접근성 (2) | ○대체 텍스트[L] |
| 정렬 (3) | ○앞으로 가져오기[L +▾] · ○뒤로 보내기[L +▾] · ○선택 창[L] · ○맞춤[S ▾] · ○그룹화[S ▾] · ○회전[S ▾] |
| 크기 (4 · 자세히(준비 중: 크기 및 위치)) | ○높이[S 입력] · ○너비[S 입력] |

**상황별 탭 「테이블 디자인」** (`table-design`, 편집 전용, 나타나는 때: 표 선택(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 표 스타일 옵션 (3) | ○머리글 행[S 확인란] · ○요약 행[S 확인란] · ○줄무늬 행[S 확인란] · ○첫째 열[S 확인란] · ○마지막 열[S 확인란] · ○줄무늬 열[S 확인란] |
| 표 스타일 (5) | ○표 스타일[L 갤러리] · ○음영[S +▾] · ○테두리[S +▾] · ○효과[S ▾] |
| 워드아트 스타일 (3) | ○빠른 스타일[L 갤러리] · ○텍스트 채우기[S +▾] · ○텍스트 윤곽선[S +▾] · ○텍스트 효과[S ▾] |
| 테두리 그리기 (4) | ○펜 스타일[S 목록] · ○펜 두께[S 목록] · ○펜 색[S +▾] · ○표 그리기[L] · ○지우개[L] |

**상황별 탭 「레이아웃」** (`table-layout`, 편집 전용, 나타나는 때: 표 선택(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 표 (3) | ○선택[S ▾] · ○눈금선 보기[S] |
| 행 및 열 (7) | ○삭제[L ▾] · ○위에 삽입[L] · ○아래에 삽입[L] · ○왼쪽에 삽입[L] · ○오른쪽에 삽입[L] |
| 병합 (6) | ○셀 병합[L] · ○셀 분할[L] |
| 셀 크기 (4) | ○높이[S 입력] · ○너비[S 입력] · ○행 높이를 같게[S] · ○열 너비를 같게[S] |
| 맞춤 (3) | ○텍스트 왼쪽 맞춤[S] · ○가운데 맞춤[S] · ○텍스트 오른쪽 맞춤[S] · ○위쪽 맞춤[S] · ○세로 가운데 맞춤[S] · ○아래쪽 맞춤[S] · ○텍스트 방향[L ▾] · ○셀 여백[L ▾] |
| 표 크기 (2) | ○높이[S 입력] · ○너비[S 입력] · ○가로 세로 비율 고정[S 확인란] |
| 정렬 (2) | ○앞으로 가져오기[S +▾] · ○뒤로 보내기[S +▾] · ○선택 창[S] · ○맞춤[S ▾] |

**상황별 탭 「그림 서식」** (`picture-format`, 편집 전용, 나타나는 때: 그림 선택(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 조정 (6) | ○배경 제거[L] · ○수정[L ▾] · ○색[L ▾] · ○꾸밈 효과[L ▾] · ○투명도[L ▾] · ○그림 압축[S] · ○그림 바꾸기[S ▾] · ○그림 원래대로[S +▾] |
| 그림 스타일 (5 · 자세히(준비 중: 그림 서식)) | ○그림 스타일[L 갤러리] · ○그림 테두리[S +▾] · ○그림 효과[S ▾] · ○그림 레이아웃[S ▾] |
| 접근성 (2) | ○대체 텍스트[L] |
| 정렬 (3) | ○앞으로 가져오기[S +▾] · ○뒤로 보내기[S +▾] · ○선택 창[S] · ○맞춤[S ▾] · ○그룹화[S ▾] · ○회전[S ▾] |
| 크기 (4 · 자세히(준비 중: 크기 및 위치)) | ○자르기[L +▾] · ○높이[S 입력] · ○너비[S 입력] |

**상황별 탭 「차트 디자인」** (`chart-design`, 편집 전용, 나타나는 때: 차트 선택(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 차트 레이아웃 (5) | ○차트 요소 추가[L ▾] · ○빠른 레이아웃[L ▾] |
| 차트 스타일 (4) | ○색 변경[L ▾] · ○차트 스타일[L 갤러리] |
| 데이터 (3) | ○행/열 전환[L] · ○데이터 선택[L] · ○데이터 편집[L +▾] · ○데이터 새로 고침[L] |
| 종류 (6) | ○차트 종류 변경[L] |

**상황별 탭 「서식」** (`chart-format`, 편집 전용, 나타나는 때: 차트 선택(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 현재 선택 영역 (4) | ○차트 요소[S 목록] · ○선택 영역 서식[S] · ○스타일에 맞게 다시 설정[S] |
| 도형 삽입 (3) | ○도형[L 갤러리] · ○도형 모양 변경[S ▾] |
| 도형 스타일 (5 · 자세히(준비 중: 도형 서식)) | ○도형 스타일[L 갤러리] · ○도형 채우기[S +▾] · ○도형 윤곽선[S +▾] · ○도형 효과[S ▾] |
| 워드아트 스타일 (4 · 자세히(준비 중: 텍스트 효과 서식)) | ○빠른 스타일[L 갤러리] · ○텍스트 채우기[S +▾] · ○텍스트 윤곽선[S +▾] · ○텍스트 효과[S ▾] |
| 접근성 (2) | ○대체 텍스트[L] |
| 정렬 (3) | ○앞으로 가져오기[S +▾] · ○뒤로 보내기[S +▾] · ○선택 창[S] · ○맞춤[S ▾] · ○그룹화[S ▾] · ○회전[S ▾] |
| 크기 (4 · 자세히(준비 중: 크기 및 위치)) | ○높이[S 입력] · ○너비[S 입력] |

**상황별 탭 「SmartArt 디자인」** (`smartart-design`, 편집 전용, 나타나는 때: SmartArt 선택(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 그래픽 만들기 (4) | ○도형 추가[S +▾] · ○글머리 기호 추가[S] · ○텍스트 창[S] · ○수준 올리기[S] · ○수준 내리기[S] · ○오른쪽에서 왼쪽[S] · ○위로 이동[S] · ○아래로 이동[S] · ○레이아웃[S ▾] |
| 레이아웃 (5) | ○레이아웃[L 갤러리] |
| SmartArt 스타일 (3) | ○색 변경[L ▾] · ○SmartArt 스타일[L 갤러리] |
| 원래대로 (2) | ○그래픽 원래대로[L] · ○변환[L ▾] |

**상황별 탭 「서식」** (`smartart-format`, 편집 전용, 나타나는 때: SmartArt 선택(우리 편집기엔 아직 없음) — 연결 없음(숨은 채))

| 묶음 (접힘 순위 · 자세히) | 단추 (●기존 ◇대리 ○준비 중, [크기 종류]) |
|---|---|
| 도형 (3) | ○도형 모양 변경[S ▾] · ○크게[S] · ○작게[S] |
| 도형 스타일 (5 · 자세히(준비 중: 도형 서식)) | ○도형 스타일[L 갤러리] · ○도형 채우기[S +▾] · ○도형 윤곽선[S +▾] · ○도형 효과[S ▾] |
| 워드아트 스타일 (4 · 자세히(준비 중: 텍스트 효과 서식)) | ○빠른 스타일[L 갤러리] · ○텍스트 채우기[S +▾] · ○텍스트 윤곽선[S +▾] · ○텍스트 효과[S ▾] |
| 접근성 (2) | ○대체 텍스트[L] |
| 정렬 (3) | ○앞으로 가져오기[S +▾] · ○뒤로 보내기[S +▾] · ○선택 창[S] · ○맞춤[S ▾] · ○그룹화[S ▾] · ○회전[S ▾] |
| 크기 (4 · 자세히(준비 중: 크기 및 위치)) | ○높이[S 입력] · ○너비[S 입력] |

