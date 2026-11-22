# 편집기·보기 화면 리본을 아트크래프트 디자인으로(기능 동일, 한글) 구현 계획

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 세 편집기(워드·엑셀·PPT)와 보기 화면(`/view.html`)의 제목 줄·탭 줄·리본·상태 줄을 아트크래프트 웹 앱의 데스크톱 오피스형 리본 모양(28/24/90/22px, 밝은 회색 틀, 큰 단추·작은 단추 세 줄, 묶음 이름 줄, "자세히" 화살표)으로 바꾸되 단추·id·동작·단축키는 그대로 둔다.

**Architecture:** 모양은 `public/css/style.css` 끝에 붙이는 한 묶음("아트크래프트 리본 꾸밈")이 맡고, `#toolbar`에 `ribbon-skin` 클래스가 붙은 화면만 바뀐다(PDF·한글 편집기·옛 `viewer.html`은 그대로). 마크업은 기존 묶음을 아트크래프트 순서에 맞춰 옮기기만 하고, "자세히"는 기존 `data-proxy`(officeShell이 이미 처리하는 대리 클릭) 단추로 기존 대화상자를 연다. 검사는 순수 규칙 모듈(`scripts/ribbonSkinLib.mjs`, vitest)과 새 브라우저 검사(`scripts/verify-ribbon-skin-e2e.mjs`) + 기존 e2e 전부.

**Tech Stack:** Node 20 · Fastify 5 · TypeScript · esbuild(`npm run build:editor`) · vitest · Playwright · 손으로 쓴 CSS(프레임워크 없음)

**Spec:** `docs/superpowers/specs/2026-11-22-artcraft-ribbon-skin-design.md`

## Global Constraints

- 단추를 더하거나 빼지 않는다. 유일한 예외는 설계 §2-1의 "자세히" 화살표(`button.ribbon-launcher`) — id 없음, 처리기 없음, `data-proxy="<기존 단추 id>"`로 기존 단추를 대신 누를 뿐이며 기존 대화상자가 있는 묶음에만 단다(워드 검토›언어 교정 → `btn-word-count`, 엑셀 데이터›정렬 및 필터 → `btn-sort-custom`). 보기 화면의 `.toolbar-sep` 구분선 `<div>`는 묶음 구분선으로 대신하므로 지운다(단추 아님).
- **id·처리기·단축키·탭 구성은 그대로.** 기존 id를 지우거나 바꾸지 않고 새 id도 만들지 않는다. `src/client/editor.ts`·`grid-editor.ts`·`editor-slides.ts`·`src/client/view/*.ts`는 고치지 않는다(이 계획에서 고치는 TS는 `src/client/officeShell.ts` 한 줄뿐).
- 글자는 전부 한글(탭·묶음 이름·단추 이름·말풍선 `title`·`aria-label`·상태 줄). 아트크래프트 영어 이름은 참고만. 글꼴 이름·용지 크기(A4·Letter)·단위(cm·pt)·글자 모양 단추(가, Aa, x²)는 라벨이 아니라 값·모양이라 그대로.
- 읽기 전용(ui=view) 규칙은 그대로: 같은 화면, 쓰는 조작만 `.ro-disabled`(투명도 0.45). 새 디자인에서도 흐림이 고르다(보이는 흐린 단추 모두 0.45). "자세히"는 가리키는 단추와 같이 흐리거나 같이 눌린다.
- 높이 규격(설계 §4, ±2px): 제목 줄 28 · 탭 줄 24 · 리본 본체 90 · 상태 줄 22px. 색: 틀 `#f3f3f3`, 리본 카드 흰색, 선 `#e0e0e0`, 강조는 기존 앱 색(`--app-accent`: 워드 `#1e5aa8`, 엑셀 `#1d6f42`, PPT `#c0462a`). 아이콘 세트는 바꾸지 않고 크기만(큰 단추 22→32px).
- 600px: 리본은 가로 스크롤 없이 접힘(아이콘만)·줄 바꿈, 제목 줄은 한 줄.
- 기존 e2e는 **고치지 않고** 그대로 통과해야 한다: `verify:readonly`(85항목)·`verify:docx`·`verify:pptx`·`verify:xlsx:browser`·`verify:artcraft`·`verify:view-only`·`verify:view-shell`.
- 향기랩 경로의 동작은 바꾸지 않는다(모양만 바뀌는 것은 의도된 변경). 새 기능·아트크래프트 앱 임베드·아이콘 세트 교체·워드 눈금자·제목 줄 자동 저장/메모/공유/사용자 원은 만들지 않는다(설계 §3).
- 운영 규칙: 포트 8000·8010·8020·8100·8110·8130·8148은 쓰지도 건드리지도 않는다(시험 서버: 새 검사 8160~8169, `verify:artcraft` 8150~8159, 수동·회귀용 서버 8165). 프로세스를 이름으로 끄지 않는다(`taskkill /IM node.exe` 금지) — 직접 띄운 것만 PID로 끈다. `public/js/*.bundle.js`·`.audit-run/` 같은 빌드·시험 산출물은 커밋하지 않는다. push 하지 않는다. 커밋은 경로를 지정해서(`git add <경로>`), 메시지는 한국어, 꼬리말 `Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>`. 가지는 `office-artcraft`.
- 시험 순서: 태스크마다 단위 시험·타입 검사·빌드만, 브라우저 e2e·회귀 묶음·실제 브라우저 조작은 Task 6에서 한 번에.

---

## 파일 구조

| 파일 | 할 일 |
|---|---|
| `scripts/ribbonSkinLib.mjs` (새) | 규격 숫자와 순수 판정(높이·한글 라벨·큰 단추 배치) — e2e가 쓴다 |
| `tests/ribbonSkinLib.test.mjs` (새) | 위 모듈 vitest(vitest 기본 include가 `.test.mjs`를 잡는다, tsc 대상 아님) |
| `public/css/style.css` | 끝에 "아트크래프트 리본 꾸밈" 묶음(토큰·제목 줄·탭·리본·묶음·단추·자세히·상태 줄·600px), Task 2·4에서 워드 레이아웃 칸·PPT 레일 몇 줄 |
| `src/client/officeShell.ts` | 보기 꾸밈 흐림 판정을 `data-proxy` 대상 기준으로(한 줄) |
| `public/editor.html` | `#toolbar`에 `ribbon-skin`, 레이아웃·삽입 묶음 순서, "자세히" 1개 |
| `public/editor-grid.html` | `#toolbar`에 `ribbon-skin`, "자세히" 1개 |
| `public/editor-slides.html` | `#toolbar`에 `ribbon-skin`, 삽입 묶음 순서 |
| `public/view.html` | `#toolbar`에 `ribbon-skin`, 탭 줄, 도구 줄을 묶음(표시·보기·쪽·찾기·인쇄·내려받기)으로 |
| `public/css/view.css` | 막대 높이 변수 `--view-chrome-h`로 쪽·썸네일·안내 위치 맞춤, 빈 묶음 숨김 |
| `scripts/verify-ribbon-skin-e2e.mjs` (새) | 설계 §4 브라우저 검사 + 나란히 비교 보고서 |
| `package.json` | `verify:ribbon-skin` |

참고로 지금 마크업은 이미 `.ribbon-group`(묶음)·`.ribbon-group-label`(묶음 이름)·`.icon-btn-big`(큰 단추: `> svg` + `> span`)·`.ribbon-row`(작은 단추 줄) 구조이고 홈 탭 묶음 순서도 아트크래프트와 같다 — 그래서 일의 대부분은 CSS다.

---

### Task 1: 꾸밈 토큰·공용 틀 CSS + 규격 판정 모듈

**Files:**
- Create: `scripts/ribbonSkinLib.mjs`
- Create: `tests/ribbonSkinLib.test.mjs`
- Modify: `public/css/style.css` (파일 끝에 덧붙임 — 같은 세기의 앞 규칙을 뒤에서 덮는다)
- Modify: `public/editor.html:46`, `public/editor-grid.html:50`, `public/editor-slides.html:48` (`<header … id="toolbar" …>` 줄) — `class`에 `ribbon-skin` 하나 추가
- Modify: `src/client/officeShell.ts:317-319` (`greyOut` 안 `viewControlAllowed` 호출)

**Interfaces:**
- Consumes: 없음
- Produces:
  - `scripts/ribbonSkinLib.mjs`: `SKIN_PX = { title: 28, tabs: 24, ribbon: 90, status: 22 }`, `SKIN_TOLERANCE_PX = 2`, `heightProblems(measured: { title?: number|null, tabs?: number|null, ribbon?: number|null, status?: number|null }): string[]`, `latinOnlyLabels(labels: string[]): string[]`, `bigButtonLayoutOk(icon: Rect, label: Rect): boolean` (`Rect = { left, top, width, height }`)
  - CSS 클래스 약속(Task 2~5가 쓴다): `#toolbar.ribbon-skin`(꾸밈 켜기), `.ribbon-group`(묶음, 마지막 칸 = `.ribbon-group-label`), `.ribbon-group-buttons`, `.ribbon-row`, `.icon-btn-big`(`> svg` 아이콘 + `> span` 이름), `button.ribbon-launcher[data-proxy]`(묶음의 마지막 자식), `.toolbar-row-title`·`.toolbar-row-tools`(보기 화면의 제목 줄·리본)
  - CSS 변수: `--skin-title-h`·`--skin-tabs-h`·`--skin-ribbon-h`·`--skin-status-h`·`--skin-label-h`·`--skin-small-h`·`--skin-big-icon`·`--skin-frame`·`--skin-line`·`--skin-hover`·`--skin-text`·`--skin-text-dim`·`--skin-font`

- [ ] **Step 1: 판정 시험을 쓴다** — `tests/ribbonSkinLib.test.mjs`

```js
import { describe, expect, test } from 'vitest';
import { SKIN_PX, SKIN_TOLERANCE_PX, bigButtonLayoutOk, heightProblems, latinOnlyLabels } from '../scripts/ribbonSkinLib.mjs';

describe('heightProblems — 설계 §4 높이 규격 ±2px', () => {
  test('규격 값과 경계(±2)는 통과', () => {
    expect(heightProblems({ title: 28, tabs: 24, ribbon: 90, status: 22 })).toEqual([]);
    expect(heightProblems({ title: 30, tabs: 22, ribbon: 88, status: 24 })).toEqual([]);
  });
  test('3px 넘게 벗어나면 그 칸 이름과 값을 알린다', () => {
    expect(heightProblems({ title: 31, tabs: 24, ribbon: 101, status: 22 })).toEqual(['title 31px (규격 28±2)', 'ribbon 101px (규격 90±2)']);
  });
  test('없는 칸(null·undefined — 보기 화면의 상태 줄)은 건너뛴다', () => {
    expect(heightProblems({ title: 28, tabs: 24, ribbon: 90, status: null })).toEqual([]);
    expect(heightProblems({ ribbon: 90 })).toEqual([]);
  });
  test('규격 숫자는 설계 값 그대로', () => {
    expect(SKIN_PX).toEqual({ title: 28, tabs: 24, ribbon: 90, status: 22 });
    expect(SKIN_TOLERANCE_PX).toBe(2);
  });
});

describe('latinOnlyLabels — 한글이 없고 영문자만 있는 라벨', () => {
  test.each([
    [['홈', '붙여넣기', '글꼴'], []],
    [['Home', '홈'], ['Home']],
    [['  Paste ', '클립보드'], ['Paste']],
    [['100%', '1.15', '', '▲'], []],
    [['PDF로 저장', 'A4 용지'], []],
  ])('%j → %j', (labels, expected) => {
    expect(latinOnlyLabels(labels)).toEqual(expected);
  });
});

describe('bigButtonLayoutOk — 큰 단추는 아이콘 위·이름 아래, 가로 가운데', () => {
  const icon = { left: 10, top: 4, width: 32, height: 32 };
  test('아이콘 아래에 이름, 가운데 맞음', () => {
    expect(bigButtonLayoutOk(icon, { left: 4, top: 38, width: 44, height: 14 })).toBe(true);
  });
  test('이름이 아이콘 옆(같은 높이)이면 실패', () => {
    expect(bigButtonLayoutOk(icon, { left: 44, top: 12, width: 40, height: 14 })).toBe(false);
  });
  test('가운데에서 3px 넘게 비켜나면 실패', () => {
    expect(bigButtonLayoutOk(icon, { left: 12, top: 38, width: 44, height: 14 })).toBe(false);
  });
  test('1px 겹침은 반올림 오차로 본다', () => {
    expect(bigButtonLayoutOk(icon, { left: 4, top: 35, width: 44, height: 14 })).toBe(true);
  });
});
```

- [ ] **Step 2: 시험이 실패하는지 본다**

Run: `npx vitest run tests/ribbonSkinLib.test.mjs`
Expected: FAIL — `../scripts/ribbonSkinLib.mjs`를 찾지 못함(파일 없음)

- [ ] **Step 3: 판정 모듈을 쓴다** — `scripts/ribbonSkinLib.mjs`

```js
// 아트크래프트 리본 꾸밈 규격(설계 2026-11-22-artcraft-ribbon-skin §4)과 순수 판정 — scripts/verify-ribbon-skin-e2e.mjs가 쓴다.
// 숫자는 public/css/style.css 의 --skin-title-h·--skin-tabs-h·--skin-ribbon-h·--skin-status-h 와 같은 값이어야 한다.

export const SKIN_PX = Object.freeze({ title: 28, tabs: 24, ribbon: 90, status: 22 });
export const SKIN_TOLERANCE_PX = 2;

/** 잰 높이 중 규격 ±2px 밖인 칸 — 값이 없는 칸(null·undefined)은 그 화면에 없는 줄이라 건너뛴다 */
export function heightProblems(measured) {
  const out = [];
  for (const [key, want] of Object.entries(SKIN_PX)) {
    const got = measured[key];
    if (got === null || got === undefined) continue;
    if (Math.abs(got - want) > SKIN_TOLERANCE_PX) out.push(`${key} ${got}px (규격 ${want}±${SKIN_TOLERANCE_PX})`);
  }
  return out;
}

/** 영문자는 있는데 한글이 하나도 없는 라벨 — 숫자·기호만 있는 것(100%, 1.15, ▲)은 라벨 문제가 아니다 */
export function latinOnlyLabels(labels) {
  return labels.map((label) => label.trim()).filter((label) => /[A-Za-z]/.test(label) && !/[가-힣]/.test(label));
}

/** 큰 단추: 아이콘 아래 끝이 이름 위 끝보다 위(1px 반올림 허용), 둘의 가로 가운데가 3px 안 */
export function bigButtonLayoutOk(icon, label) {
  const iconMid = icon.left + icon.width / 2;
  const labelMid = label.left + label.width / 2;
  return icon.top + icon.height <= label.top + 1 && Math.abs(iconMid - labelMid) <= 3;
}
```

- [ ] **Step 4: 시험이 통과하는지 본다**

Run: `npx vitest run tests/ribbonSkinLib.test.mjs`
Expected: PASS (13 tests)

- [ ] **Step 5: 꾸밈 CSS 묶음을 `public/css/style.css` 맨 끝에 덧붙인다**

```css

/* ==================== 아트크래프트 리본 꾸밈(설계 2026-11-22-artcraft-ribbon-skin) ====================
   #toolbar에 .ribbon-skin이 붙은 화면(워드·엑셀·PPT 편집기, 보기 화면 view.html)만 바뀐다 — 모양만, 단추·id·동작은 그대로.
   높이 규격(scripts/ribbonSkinLib.mjs SKIN_PX와 같은 값): 제목 줄 28 · 탭 줄 24 · 리본 90 · 상태 줄 22px.
   리본 90px = 위 여백 3 + 단추 칸 71(작은 단추 22px 세 줄 = 66, 큰 단추 68) + 묶음 이름 줄 16. 이 묶음은 파일 끝에 있어서
   같은 세기의 앞 규칙(:is(.grid-editor-page, …) …, .office-shell …)을 뒤에서 덮는다 — 앞으로 옮기지 말 것. */
:root {
  --skin-title-h: 28px;
  --skin-tabs-h: 24px;
  --skin-ribbon-h: 90px;
  --skin-status-h: 22px;
  --skin-label-h: 16px;
  --skin-small-h: 22px;
  --skin-big-icon: 32px;
  --skin-frame: #f3f3f3;
  --skin-line: #e0e0e0;
  --skin-hover: #e6e6e6;
  --skin-text: #242424;
  --skin-text-dim: #616161;
  --skin-font: "Segoe UI", "맑은 고딕", "Malgun Gothic", -apple-system, "Apple SD Gothic Neo", system-ui, sans-serif;
}
#toolbar.ribbon-skin { background: var(--skin-frame); border-bottom: 1px solid var(--skin-line); box-shadow: none; font-family: var(--skin-font); }

/* 제목 줄 — 밝은 회색(앱 색 띠 아님), 빠른 실행·문서 이름·오른쪽 단추 */
.ribbon-skin .office-titlebar,
.ribbon-skin .toolbar-row-title { box-sizing: border-box; height: var(--skin-title-h); padding: 0 8px; gap: 8px; background: var(--skin-frame); color: var(--skin-text); }
.ribbon-skin :is(.office-titlebar, .toolbar-row-title) .doc-title { color: var(--skin-text); font: 400 12px/var(--skin-title-h) var(--skin-font); }
.ribbon-skin :is(.office-titlebar, .toolbar-row-title) :is(.icon-btn, .qa-btn) { width: 24px; min-width: 24px; height: 24px; border-radius: 4px; background: transparent; color: var(--skin-text); }
.ribbon-skin :is(.office-titlebar, .toolbar-row-title) :is(.icon-btn, .qa-btn):hover:not(:disabled) { background: var(--skin-hover); color: var(--skin-text); }
.ribbon-skin :is(.office-titlebar, .toolbar-row-title) :is(.icon-btn, .qa-btn):focus-visible { outline: 2px solid var(--app-accent, var(--v-accent)); outline-offset: -2px; }
.ribbon-skin :is(.office-titlebar, .toolbar-row-title) :is(.icon-btn, .qa-btn) svg { width: 16px; height: 16px; }
.ribbon-skin .quick-access { border-left-color: var(--skin-line); }
.ribbon-skin .office-titlebar .save-status { color: var(--skin-text-dim); }
.ribbon-skin .office-titlebar .conn-status { padding: 1px 8px; font-size: 11px; line-height: 16px; background: var(--skin-hover); color: var(--skin-text-dim); }
.ribbon-skin .office-titlebar .readonly-badge { background: var(--app-accent); color: #ffffff; }
.ribbon-skin :is(.office-titlebar, .toolbar-row-title) .presence-list { --v-presence-border: var(--skin-frame); }
.ribbon-skin :is(.office-titlebar, .toolbar-row-title) .presence-badge { width: 22px; height: 22px; font-size: 10px; }
.ribbon-skin :is(.office-titlebar, .toolbar-row-title) :is(.back-to-view, .btn-pill) { padding: 0 10px; border-color: var(--skin-line); background: #ffffff; color: var(--skin-text); font-size: 12px; line-height: 20px; }
.ribbon-skin :is(.office-titlebar, .toolbar-row-title) :is(.back-to-view, .btn-pill):hover:not(:disabled) { background: var(--skin-hover); }
/* 편집/보기 모드 단추(편집 열기 · 보기 화면의 편집) — 아트크래프트의 모드 단추처럼 앱 색 알약 */
.ribbon-skin :is(.office-titlebar, .toolbar-row-title) :is(.view-open-edit, .edit-btn) { border-color: transparent; background: var(--app-accent, var(--v-accent)); color: #ffffff; }

/* 탭 줄 — 24px, 고른 탭은 앱 색 글자 + 밑줄 */
.ribbon-skin .ribbon-tabs { box-sizing: border-box; height: var(--skin-tabs-h); align-items: stretch; gap: 2px; padding: 0 8px; background: var(--skin-frame); }
.ribbon-skin .ribbon-tab { padding: 0 8px; border-radius: 0; color: var(--skin-text); font: 400 12px/var(--skin-tabs-h) var(--skin-font); }
.ribbon-skin .ribbon-tab:hover { background: var(--skin-hover); }
.ribbon-skin .ribbon-tab[aria-selected="true"] { color: var(--app-accent, var(--v-accent)); font-weight: 600; }
.ribbon-skin .ribbon-tab[aria-selected="true"]::after { left: 8px; right: 8px; bottom: 1px; height: 2px; border-radius: 1px; background: var(--app-accent, var(--v-accent)); }
.ribbon-skin .ribbon-tab-contextual { color: var(--app-accent); }
.ribbon-skin .ribbon-collapse { width: 22px; height: 22px; }

/* 리본 본체 — 회색 틀 안 흰 카드 90px. 탭마다 높이가 같아 아래 작업 영역이 움직이지 않는다 */
.ribbon-skin .toolbar-format.ribbon-panel,
.ribbon-skin .toolbar-row-tools {
  box-sizing: border-box;
  height: var(--skin-ribbon-h);
  min-height: 0;
  margin: 0 6px 6px;
  padding: 3px 4px 0;
  gap: 0;
  align-items: stretch;
  border: 0;
  border-radius: 6px;
  background: #ffffff;
  box-shadow: 0 0 0 1px var(--skin-line), 0 1px 2px rgba(0, 0, 0, 0.06);
  overflow-y: hidden;
}
.ribbon-skin.ribbon-peeking .ribbon-panel:not([hidden]) { top: calc(var(--skin-title-h) + var(--skin-tabs-h)); }

/* 묶음 — 위는 단추 칸, 아래 16px은 묶음 이름, 사이는 세로 구분선 */
.ribbon-skin .ribbon-group { position: relative; display: grid; grid-template-rows: minmax(0, 1fr) var(--skin-label-h); gap: 0; padding: 0 6px; border-right: 1px solid var(--skin-line); }
.ribbon-skin .ribbon-group:last-child { border-right: none; }
.ribbon-skin .ribbon-group-buttons { align-items: center; gap: 2px; }
.ribbon-skin .ribbon-group-buttons-wrap { flex-direction: column; justify-content: center; align-items: stretch; gap: 0; }
.ribbon-skin .ribbon-group-label { height: var(--skin-label-h); color: var(--skin-text-dim); font: 400 11px/var(--skin-label-h) var(--skin-font); letter-spacing: 0; text-align: center; white-space: nowrap; }

/* 작은 단추 — 22px 한 줄(세 줄 쌓기), 아이콘+이름 단추도 같은 높이. 아이콘 크기(17px)는 그대로 */
.ribbon-skin .ribbon-row,
.ribbon-skin .ribbon-row:has(.ribbon-btn-labeled) { min-height: var(--skin-small-h); gap: 1px; }
.ribbon-skin .ribbon-panel .icon-btn:not(.icon-btn-big) { width: auto; min-width: var(--skin-small-h); height: var(--skin-small-h); padding: 0 3px; border-radius: 3px; }
.ribbon-skin .ribbon-btn-labeled { height: var(--skin-small-h); padding: 0 4px; gap: 4px; }
.ribbon-skin .ribbon-btn-label { color: var(--skin-text); font: 400 12px/16px var(--skin-font); }
.ribbon-skin .ribbon-panel :is(.toolbar-select, .toolbar-number) { height: var(--skin-small-h); font: 400 12px var(--skin-font); }

/* 큰 단추 — 아이콘(32px) 위, 이름 아래 */
.ribbon-skin .icon-btn-big {
  box-sizing: border-box;
  width: auto;
  min-width: 44px;
  height: calc(var(--skin-ribbon-h) - var(--skin-label-h) - 6px);
  padding: 4px 4px 0;
  justify-content: flex-start;
  gap: 2px;
  color: var(--skin-text);
  font: 400 12px/14px var(--skin-font);
}
.ribbon-skin .icon-btn-big > svg { flex: none; width: var(--skin-big-icon); height: var(--skin-big-icon); }

/* "자세히" — 묶음 오른쪽 아래 작은 화살표, data-proxy로 기존 단추(대화상자)를 대신 누른다(officeShell.ts) */
.ribbon-launcher { position: absolute; right: 1px; bottom: 1px; display: inline-flex; align-items: center; justify-content: center; width: 14px; height: 14px; padding: 0; border: 0; border-radius: 2px; background: transparent; color: var(--skin-text-dim); cursor: pointer; }
.ribbon-launcher:hover { background: var(--skin-hover); color: var(--skin-text); }
.ribbon-launcher:focus-visible { outline: 2px solid var(--app-accent, var(--v-accent)); outline-offset: 0; }
.ribbon-launcher svg { width: 10px; height: 10px; }
.ribbon-group:has(> .ribbon-launcher) .ribbon-group-label { padding-right: 16px; }

/* 상태 줄 — 22px(워드·PPT는 #statusbar, 엑셀은 격자 아래 .grid-status-row) */
:is(.word-editor-page, .slides-editor-page) .office-statusbar { box-sizing: border-box; height: var(--skin-status-h); padding: 0 8px; gap: 16px; border-top-color: var(--skin-line); background: var(--skin-frame); font: 400 11px var(--skin-font); }
.grid-editor-page .grid-status-row { box-sizing: border-box; height: var(--skin-status-h); padding-top: 0; font-size: 11px; }

/* 600px — 리본은 가로 스크롤 없이 아이콘만 남기고 줄 바꿈, 묶음 이름·자세히는 접는다. 제목 줄은 위 규칙 그대로 한 줄 */
@media (max-width: 600px) {
  .ribbon-skin .toolbar-format.ribbon-panel,
  .ribbon-skin .toolbar-row-tools { height: auto; flex-wrap: wrap; row-gap: 2px; padding-bottom: 3px; overflow: visible; }
  .ribbon-skin .ribbon-group { flex: 0 1 auto; max-width: 100%; grid-template-rows: auto; padding: 0 3px; }
  .ribbon-skin .ribbon-group-buttons,
  .ribbon-skin .ribbon-row { flex-wrap: wrap; }
  .ribbon-skin .ribbon-group-label,
  .ribbon-skin .ribbon-launcher,
  .ribbon-skin .icon-btn-big > span,
  .ribbon-skin .ribbon-btn-label { display: none; }
  .ribbon-skin .icon-btn-big { flex-direction: row; min-width: var(--skin-small-h); height: var(--skin-small-h); padding: 0 3px; }
  .ribbon-skin .icon-btn-big > svg { width: 18px; height: 18px; }
}
```

- [ ] **Step 6: 세 편집기 머리에 `ribbon-skin`을 단다** — 세 파일 모두 `class="editor-toolbar office-shell"`을 `class="editor-toolbar office-shell ribbon-skin"`으로(id·data-app 그대로).

Run: `grep -n 'id="toolbar"' public/editor.html public/editor-grid.html public/editor-slides.html`
Expected: 세 줄 모두 `class="editor-toolbar office-shell ribbon-skin" id="toolbar" data-app="…"`

- [ ] **Step 7: 보기 꾸밈 흐림을 "자세히"의 대상 기준으로** — `src/client/officeShell.ts`의 `greyOut` 안 첫 `for` 문을 바꾼다. "자세히"는 id가 없어 지금 판정으로는 늘 흐려지는데, 대상이 보기 조작(단어 개수)이면 함께 눌려야 한다. 기존 `data-proxy` 단추(PPT 삽입 탭 5개)의 대상은 모두 쓰는 조작이라 결과가 같다.

```ts
      for (const el of document.querySelectorAll<HTMLElement>(VIEW_EDIT_CONTROLS)) {
        // "자세히"·삽입 탭 대리 단추(data-proxy)는 대신 누르는 단추와 같은 판정
        if (!viewControlAllowed(el.dataset.proxy ?? el.id, options.canCopy)) markReadOnly(el);
      }
```

- [ ] **Step 8: 타입·단위 시험·빌드**

Run: `npm run typecheck && npx vitest run tests/ribbonSkinLib.test.mjs tests/viewStrip.test.ts && npm run build:editor`
Expected: 타입 오류 0, 두 파일 PASS, 빌드 완료(`public/js/*.bundle.js`는 gitignore — 커밋하지 않는다)

- [ ] **Step 9: 커밋**

```bash
git add scripts/ribbonSkinLib.mjs tests/ribbonSkinLib.test.mjs public/css/style.css public/editor.html public/editor-grid.html public/editor-slides.html src/client/officeShell.ts
git commit -m "리본 꾸밈 1: 아트크래프트 규격 토큰·공용 틀 CSS(28/24/90/22px, 큰·작은 단추, 묶음 이름, 자세히, 600px)와 규격 판정 모듈

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 2: 워드 편집기 — 묶음 배치·자세히·레이아웃 칸

**Files:**
- Modify: `public/editor.html:354-378` (삽입 탭), `:382-433` (레이아웃 탭), `:453-464` (검토 › 언어 교정)
- Modify: `public/css/style.css` (Task 1 묶음 뒤에 워드 몇 줄)

**Interfaces:**
- Consumes: Task 1의 `.ribbon-skin` 규칙, `button.ribbon-launcher[data-proxy]` 약속, officeShell의 `[data-proxy]` 대리 클릭(기존 `initOfficeShell` 안 `for (const proxy of toolbar.querySelectorAll('[data-proxy]'))`)
- Produces: 워드 "자세히" 1개(`data-proxy="btn-word-count"`) — Task 6 e2e가 개수 1을 기대한다

아트크래프트 워드 순서: 홈(클립보드·글꼴·단락·스타일·편집) — 이미 같다. 삽입은 쪽 묶음이 표보다 앞, 레이아웃의 나누기는 페이지 설정 묶음 안. 검토(언어 교정·메모·추적)·보기(확대/축소)·표 레이아웃은 그대로.

- [ ] **Step 1: 라벨이 한글인지 먼저 확인한다(영문만 있는 것이 나오면 이 태스크에서 한글로 바꾼다)**

Run: `node -e "const s=require('fs').readFileSync(process.argv[1],'utf8');const bad=[...s.matchAll(/(?:title|aria-label)=\"([^\"]*)\"|<span(?: class=\"ribbon-(?:group-label|btn-label)\")?>([^<]+)<\/span>|role=\"tab\"[^>]*>([^<]+)</g)].map((m)=>m[1]??m[2]??m[3]).filter((t)=>/[A-Za-z]/.test(t)&&!/[가-힣]/.test(t));console.log(bad.length?bad:'OK')" public/editor.html`
Expected: `OK` (계획 작성 때 확인함)

- [ ] **Step 2: 삽입 탭 — 나누기 묶음을 표 묶음 앞으로 옮긴다.** `data-panel="insert"` 판 안의 두 `<div class="ribbon-group">` 블록 순서만 바꾼다(안의 마크업은 한 글자도 바꾸지 않는다). 결과 순서:

```html
    <div class="toolbar-group toolbar-format ribbon-panel" data-panel="insert" role="tabpanel" hidden>
      <div class="ribbon-group">
        <div class="ribbon-group-buttons">
          <button id="btn-insert-break" class="icon-btn icon-btn-big" title="줄 바꿈 (Shift+Enter)">…기존 내용 그대로…</button>
        </div>
        <span class="ribbon-group-label">나누기</span>
      </div>
      <div class="ribbon-group">
        <div class="ribbon-group-buttons">
          <span class="dropdown-btn" id="table-insert-dropdown">…기존 내용 그대로…</span>
        </div>
        <span class="ribbon-group-label">표</span>
      </div>
    </div>
```

Run: `node -e "const s=require('fs').readFileSync('public/editor.html','utf8');const p=s.slice(s.indexOf('data-panel=\"insert\"'),s.indexOf('data-panel=\"layout\"'));console.log([...p.matchAll(/ribbon-group-label\">([^<]+)/g)].map(m=>m[1]).join(','))"`
Expected: `나누기,표`

- [ ] **Step 3: 레이아웃 탭 — 페이지 나누기를 페이지 설정 묶음 안으로.** `<button id="btn-page-break" …>…</button>` 블록 전체를 잘라 `#paper-size-dropdown`의 닫는 `</span>` 바로 뒤(같은 `.ribbon-group-buttons` 안)에 붙이고, 비게 된 레이아웃 "나누기" 묶음(`<div class="ribbon-group">` … `<span class="ribbon-group-label">나누기</span>` … `</div>`)을 지운다. 결과 순서: 페이지 설정(여백·용지 방향·크기·페이지 나누기) · 단락(들여쓰기·간격 칸).

Run: `node -e "const s=require('fs').readFileSync('public/editor.html','utf8');const p=s.slice(s.indexOf('data-panel=\"layout\"'),s.indexOf('data-panel=\"review\"'));console.log([...p.matchAll(/ribbon-group-label\">([^<]+)/g)].map(m=>m[1]).join(','), p.includes('id=\"btn-page-break\"'))"`
Expected: `페이지 설정,단락 true`

- [ ] **Step 4: 검토 › 언어 교정 묶음에 "자세히"를 단다** — `<span class="ribbon-group-label">언어 교정</span>` 바로 뒤(그 묶음의 닫는 `</div>` 앞)에:

```html
        <button type="button" class="ribbon-launcher" data-proxy="btn-word-count" title="단어 개수 대화상자 열기" aria-label="언어 교정 자세히">
          <svg viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"><path d="M1.5 5.5v-4h4M4 4l4.5 4.5M8.5 5v3.5H5"/></svg>
        </button>
```

- [ ] **Step 5: 레이아웃 들여쓰기·간격 칸이 90px 리본 안에 들게** — `public/css/style.css` 끝에 덧붙인다. 칸 한 열 = 제목 16 + 1 + 입력 20 + 1 + 입력 20 = 58px(큰 단추 칸 68px 안).

```css
/* 리본 꾸밈 — 워드 레이아웃 탭 들여쓰기·간격 칸(58px 높이로) */
.ribbon-skin .layout-fields { gap: 12px; align-items: center; }
.ribbon-skin .layout-field-col { gap: 1px; }
.ribbon-skin .layout-field-title { font: 600 11px/16px var(--skin-font); }
.ribbon-skin .layout-field { font: 400 12px/20px var(--skin-font); }
.ribbon-skin .layout-field input { box-sizing: border-box; height: 20px; padding: 0 4px; }
```

- [ ] **Step 6: id가 그대로인지 확인한다**

Run: `git show HEAD:public/editor.html | grep -o 'id="[^"]*"' | sort > .audit-run-ids-before.txt; grep -o 'id="[^"]*"' public/editor.html | sort > .audit-run-ids-after.txt; diff .audit-run-ids-before.txt .audit-run-ids-after.txt && echo SAME; rm -f .audit-run-ids-*.txt`
Expected: `SAME`(id 추가·삭제 없음 — "자세히"에는 id가 없다)

- [ ] **Step 7: 빌드·단위 시험**

Run: `npm run build:editor && npx vitest run tests/ribbonSkinLib.test.mjs tests/viewStrip.test.ts`
Expected: 빌드 완료, PASS

- [ ] **Step 8: 커밋**

```bash
git add public/editor.html public/css/style.css
git commit -m "리본 꾸밈 2: 워드 — 삽입 탭 나누기를 표 앞으로, 페이지 나누기를 페이지 설정 묶음 안으로, 언어 교정 자세히(단어 개수), 레이아웃 칸 90px 안

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 3: 엑셀 편집기 — 자세히, 이름 상자·수식 줄 유지

**Files:**
- Modify: `public/editor-grid.html:564` (데이터 › 정렬 및 필터 묶음 이름 줄 뒤)

**Interfaces:**
- Consumes: Task 1의 `.ribbon-skin`, `.grid-editor-page .grid-status-row` 22px 규칙, `button.ribbon-launcher[data-proxy]`
- Produces: 엑셀 "자세히" 1개(`data-proxy="btn-sort-custom"`) — Task 6 e2e가 개수 1을 기대한다

아트크래프트 엑셀 순서: 홈(클립보드·글꼴·맞춤·표시 형식·스타일·셀·편집) — 이미 같다. 삽입(표·차트)·수식·데이터·보기·차트 디자인도 그대로. 이름 상자(`#cell-ref-box`)와 수식 줄(`.formula-bar` · `#formula-input`)은 리본 아래 그대로 둔다(마크업·CSS 손대지 않음). 1627/1447/1396/1270px에서 홈 탭 글자를 접는 기존 미디어 규칙(id 선택자라 꾸밈 규칙보다 세다)도 그대로 둔다.

- [ ] **Step 1: 라벨이 한글인지 확인**

Run: `node -e "const s=require('fs').readFileSync(process.argv[1],'utf8');const bad=[...s.matchAll(/(?:title|aria-label)=\"([^\"]*)\"|<span(?: class=\"ribbon-(?:group-label|btn-label)\")?>([^<]+)<\/span>|role=\"tab\"[^>]*>([^<]+)</g)].map((m)=>m[1]??m[2]??m[3]).filter((t)=>/[A-Za-z]/.test(t)&&!/[가-힣]/.test(t));console.log(bad.length?bad:'OK')" public/editor-grid.html`
Expected: `OK`

- [ ] **Step 2: 데이터 › 정렬 및 필터에 "자세히"** — `<span class="ribbon-group-label">정렬 및 필터</span>` 바로 뒤(그 묶음의 닫는 `</div>` 앞)에:

```html
        <button type="button" class="ribbon-launcher" data-proxy="btn-sort-custom" title="사용자 지정 정렬 대화상자 열기" aria-label="정렬 및 필터 자세히">
          <svg viewBox="0 0 10 10" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"><path d="M1.5 5.5v-4h4M4 4l4.5 4.5M8.5 5v3.5H5"/></svg>
        </button>
```

- [ ] **Step 3: id가 그대로이고 이름 상자·수식 줄·상태 줄이 남아 있는지**

Run: `git show HEAD:public/editor-grid.html | grep -o 'id="[^"]*"' | sort > .audit-run-ids-before.txt; grep -o 'id="[^"]*"' public/editor-grid.html | sort > .audit-run-ids-after.txt; diff .audit-run-ids-before.txt .audit-run-ids-after.txt && echo SAME; rm -f .audit-run-ids-*.txt; grep -c 'id="cell-ref-box"\|id="formula-input"\|class="grid-status-row"' public/editor-grid.html`
Expected: `SAME`, 그다음 `3`

- [ ] **Step 4: 빌드·단위 시험**

Run: `npm run build:editor && npx vitest run tests/ribbonSkinLib.test.mjs tests/viewStrip.test.ts`
Expected: 빌드 완료, PASS

- [ ] **Step 5: 커밋**

```bash
git add public/editor-grid.html
git commit -m "리본 꾸밈 3: 엑셀 — 정렬 및 필터 자세히(사용자 지정 정렬 대화상자), 이름 상자·수식 줄은 그대로

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 4: PPT 편집기 — 삽입 묶음 순서·섬네일 레일

**Files:**
- Modify: `public/editor-slides.html:215-252` (삽입 탭)
- Modify: `public/css/style.css` (끝에 레일 몇 줄)

**Interfaces:**
- Consumes: Task 1의 `.ribbon-skin`, 상태 줄 22px 규칙(`.slides-editor-page .office-statusbar`)
- Produces: PPT "자세히" 0개(기존 대화상자 없음) — Task 6 e2e가 개수 0을 기대한다

아트크래프트 PPT 순서: 홈(슬라이드·글꼴·단락·그리기) — 이미 같다. 삽입은 슬라이드 · 도형 · 텍스트 순. 디자인·전환·애니메이션은 그대로. 섬네일 레일(`#slides-rail`)은 폭(220px)·안쪽 여백을 바꾸지 않는다(섬네일 크기 계산·`verify:pptx` 영향 없게) — 색과 번호 글자만.

- [ ] **Step 1: 라벨이 한글인지 확인**

Run: `node -e "const s=require('fs').readFileSync(process.argv[1],'utf8');const bad=[...s.matchAll(/(?:title|aria-label)=\"([^\"]*)\"|<span(?: class=\"ribbon-(?:group-label|btn-label)\")?>([^<]+)<\/span>|role=\"tab\"[^>]*>([^<]+)</g)].map((m)=>m[1]??m[2]??m[3]).filter((t)=>/[A-Za-z]/.test(t)&&!/[가-힣]/.test(t));console.log(bad.length?bad:'OK')" public/editor-slides.html`
Expected: `OK`

- [ ] **Step 2: 삽입 탭 — "텍스트" 묶음을 "도형" 묶음 뒤로.** `data-panel="insert"` 판 안의 `<div class="ribbon-group">` … `<span class="ribbon-group-label">텍스트</span>` … `</div>` 블록을 통째로 잘라 "도형" 묶음의 닫는 `</div>` 뒤에 붙인다(안의 `data-proxy` 단추는 그대로).

Run: `node -e "const s=require('fs').readFileSync('public/editor-slides.html','utf8');const p=s.slice(s.indexOf('data-panel=\"insert\"'),s.indexOf('data-panel=\"design\"'));console.log([...p.matchAll(/ribbon-group-label\">([^<]+)/g)].map(m=>m[1]).join(','))"`
Expected: `슬라이드,도형,텍스트`

- [ ] **Step 3: 레일 색·번호 글자** — `public/css/style.css` 끝에:

```css
/* 리본 꾸밈 — PPT 왼쪽 섬네일 레일(폭·여백은 그대로: 섬네일 크기 계산이 엮여 있다) */
.slides-editor-page .slides-rail { background: var(--skin-frame); border-right-color: var(--skin-line); }
.slides-editor-page .slide-card-number { color: var(--skin-text-dim); font: 400 11px var(--skin-font); }
```

- [ ] **Step 4: id 그대로**

Run: `git show HEAD:public/editor-slides.html | grep -o 'id="[^"]*"' | sort > .audit-run-ids-before.txt; grep -o 'id="[^"]*"' public/editor-slides.html | sort > .audit-run-ids-after.txt; diff .audit-run-ids-before.txt .audit-run-ids-after.txt && echo SAME; rm -f .audit-run-ids-*.txt`
Expected: `SAME`

- [ ] **Step 5: 빌드·단위 시험**

Run: `npm run build:editor && npx vitest run tests/ribbonSkinLib.test.mjs`
Expected: 빌드 완료, PASS

- [ ] **Step 6: 커밋**

```bash
git add public/editor-slides.html public/css/style.css
git commit -m "리본 꾸밈 4: PPT — 삽입 탭 슬라이드·도형·텍스트 순, 섬네일 레일 회색 틀

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 5: 보기 화면(view.html) 도구 막대를 같은 규격으로

**Files:**
- Modify: `public/view.html:44-149` (`<header class="viewer-toolbar" id="toolbar">` 전체)
- Modify: `public/css/view.css` (위치 상수 `6.2rem`·`6.4rem` → 변수, 빈 묶음 숨김, 도구 줄 단추 크기)

**Interfaces:**
- Consumes: Task 1의 `.ribbon-skin` 규칙(`.toolbar-row-title` 28px, `.ribbon-tabs` 24px, `.toolbar-row-tools` 90px 카드, `.ribbon-group` 묶음, 600px 접기)
- Produces: 보기 화면 구조 약속(Task 6 e2e가 쓴다): `#toolbar.ribbon-skin` 안에 `.toolbar-row-title` · `.ribbon-tabs` · `.toolbar-row-tools`; 묶음 이름이 문서 순서로 `표시,보기,쪽,찾기,인쇄,내려받기`; CSS 변수 `--view-chrome-h`(막대 전체 높이, 넓은 화면 150px / 600px 이하 96px)

지켜야 할 기존 약속(`src/client/view/toolbar.ts`·`logic.ts`의 `menuIdsFor`·`verify:view-shell`): 세 칸 `#toolbar-tools-left.toolbar-tools-left`·`.toolbar-tools-center`·`.toolbar-tools-right`는 그대로 두고 그 **안에** 묶음을 넣는다. `#print-btn`은 넓은 화면에서 `.toolbar-tools-right` 안이어야 한다(view-shell 477행). 도구▾ 메뉴로 옮겨지는 단추(`TOOL_MENU_IDS`·`NARROW_MENU_IDS`)는 메뉴에서 돌아올 때 기억해 둔 원래 부모(묶음)로 돌아오므로 그대로 산다. `#edit-btn`은 제목 줄에 그대로 둔다(편집기의 "편집 열기"와 같은 자리, 390px 배치 유지). 640px 미만에서 `#toolbar-tools-left`를 숨기는 view.css 규칙도 그대로 — 그래서 확대(`#zoom-*`)는 가운데 칸에 둔다.

- [ ] **Step 1: 머리에 `ribbon-skin`, 탭 줄을 넣는다** — `<header class="viewer-toolbar" id="toolbar">`를 `<header class="viewer-toolbar ribbon-skin" id="toolbar">`로 바꾸고, `.toolbar-row-title`의 닫는 `</div>`와 `<div class="toolbar-row toolbar-row-tools">` 사이에 탭 줄(단추 아님 — 탭이 하나라 고를 것이 없다):

```html
    <div class="ribbon-tabs" role="tablist" aria-label="리본">
      <span class="ribbon-tab" role="tab" aria-selected="true">보기</span>
    </div>
```

- [ ] **Step 2: 도구 줄을 묶음으로 다시 배치한다.** 아래 뼈대대로 쓰되, `[#id]` 자리에는 지금 `view.html`에 있는 그 요소의 마크업 **전체를 한 글자도 바꾸지 않고** 옮겨 붙인다(바로 앞 HTML 주석이 있으면 주석째: `#zoom-fit-btn`·`#zoom-level`·`#fullscreen-btn`). 대상은 `<button …>…</button>`·`<label id="markup-include-wrap" …>…</label>`·`<span id="zoom-level" …>…</span>`·`<div class="page-jump-group" …>…</div>`·`<div id="tools-wrap" …>…</div>`. 도구 줄 안의 `<div class="toolbar-sep" aria-hidden="true"></div>`(모두)는 지운다 — 묶음 구분선이 대신한다.

```html
    <div class="toolbar-row toolbar-row-tools">
      <div id="toolbar-tools-left" class="toolbar-group toolbar-tools toolbar-tools-left">
        <div class="ribbon-group">
          <div class="ribbon-group-buttons">[#measure-btn] [#markup-add-btn] [#text-tool-btn] [#markup-include-wrap]</div>
          <span class="ribbon-group-label">표시</span>
        </div>
      </div>

      <div class="toolbar-group toolbar-tools toolbar-tools-center">
        <div class="ribbon-group">
          <div class="ribbon-group-buttons">[#thumbs-btn] [#zoom-out] [#zoom-in] [#zoom-fit-btn] [#zoom-level] [#rotate-btn] [#invert-btn] [#fullscreen-btn]</div>
          <span class="ribbon-group-label">보기</span>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-buttons">[#page-prev] [.page-jump-group] [#page-next]</div>
          <span class="ribbon-group-label">쪽</span>
        </div>
      </div>

      <div class="toolbar-group toolbar-tools toolbar-tools-right">
        <div class="ribbon-group">
          <div class="ribbon-group-buttons">[#search-btn]</div>
          <span class="ribbon-group-label">찾기</span>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-buttons">[#print-btn]</div>
          <span class="ribbon-group-label">인쇄</span>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-buttons">[#download-btn] [#download-as-btn]</div>
          <span class="ribbon-group-label">내려받기</span>
        </div>
        [#tools-wrap]
      </div>
    </div>
```

Run: `git show HEAD:public/view.html | grep -o 'id="[^"]*"' | sort > .audit-run-ids-before.txt; grep -o 'id="[^"]*"' public/view.html | sort > .audit-run-ids-after.txt; diff .audit-run-ids-before.txt .audit-run-ids-after.txt && echo SAME; rm -f .audit-run-ids-*.txt; grep -c 'toolbar-sep' public/view.html`
Expected: `SAME`, 그다음 `0`

- [ ] **Step 3: view.css — 막대 높이 변수와 위치, 빈 묶음 숨김.** 막대 = 제목 줄 28 + 탭 줄 24 + 리본 90 + 아래 여백 6 + 선 1 = 149px → 150px. 600px 이하는 리본이 한 줄(위아래 여백 3+3 + 단추 28 = 34px)로 접혀 28 + 24 + 34 + 6 + 1 = 93px → 96px. `public/css/view.css`에서 세 값을 바꾼다:
  - `.view-notice`의 `top: 6.4rem;` → `top: calc(var(--view-chrome-h) + 0.3rem);`
  - `.thumbs-panel`의 `top: 6.2rem;` → `top: var(--view-chrome-h);`
  - `.prep-screen`의 `inset: 6.2rem 0 0;` → `inset: var(--view-chrome-h) 0 0;`

그리고 파일 끝에 덧붙인다:

```css
/* 리본 꾸밈(설계 2026-11-22-artcraft-ribbon-skin §2-5) — 막대 전체 높이. 쪽·도면 칸이 막대 밑에 깔리지 않게 같은 값을 쓴다 */
:root { --view-chrome-h: 150px; }
@media (max-width: 600px) { :root { --view-chrome-h: 96px; } }
.pages-container,
.svg-scroll { padding-top: calc(var(--view-chrome-h) + 0.6rem); }
/* 도구 줄 — 세 칸 안의 묶음이 카드 높이를 채우고, 단추는 한 줄(28px) */
.ribbon-skin .toolbar-tools { gap: 0; align-items: stretch; }
.ribbon-skin .toolbar-row-tools .icon-btn { width: 28px; height: 28px; }
.ribbon-skin .toolbar-row-tools .tools-btn { width: auto; align-self: center; }
.ribbon-skin .ribbon-tab:hover { background: transparent; cursor: default; }
/* 단추가 모두 도구▾로 옮겨졌거나 숨은 묶음은 이름만 남지 않게 통째로 숨긴다 */
.ribbon-skin .ribbon-group:not(:has(> .ribbon-group-buttons > :not([hidden]))) { display: none; }
```

- [ ] **Step 4: 순수 규칙 시험(도구▾ 메뉴 판정은 그대로)·빌드**

Run: `npx vitest run tests/viewShellLogic.test.ts tests/ribbonSkinLib.test.mjs && npm run build:editor`
Expected: PASS, 빌드 완료(`menuIdsFor`·`TOOL_MENU_IDS`·`NARROW_MENU_IDS`는 바꾸지 않았다)

- [ ] **Step 5: 커밋**

```bash
git add public/view.html public/css/view.css
git commit -m "리본 꾸밈 5: 보기 화면 — 제목 줄·탭 줄·90px 리본 카드, 도구를 표시·보기·쪽·찾기·인쇄·내려받기 묶음으로(id·도구▾ 동작 그대로), 쪽 위치를 막대 높이 변수로

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

---

### Task 6: 리본 꾸밈 브라우저 검사·나란히 비교 보고서·전체 회귀

**Files:**
- Create: `scripts/verify-ribbon-skin-e2e.mjs`
- Modify: `package.json` (`"verify:artcraft"` 줄 뒤에 `"verify:ribbon-skin"`)

**Interfaces:**
- Consumes: `scripts/ribbonSkinLib.mjs`(`heightProblems`, `latinOnlyLabels`, `bigButtonLayoutOk`), Task 1~5의 마크업·클래스 약속, "자세히" 개수(워드 1·엑셀 1·PPT 0), 보기 화면 묶음 이름 `표시,보기,쪽,찾기,인쇄,내려받기`, 서버 API(`POST /api/documents`, `POST /api/documents/:id/tokens` — `engine: 'editor'`면 `viewerUrl`이 `<편집기>?token=…&ui=view`, `canEdit: true`면 `editorUrl`, 기본 `viewerUrl`은 docx·pdf가 `/view.html`)
- Produces: `npm run verify:ribbon-skin`(포트 8160~8169, 기본 8166), 캡처·보고서 `.audit-run/ribbon-skin-shots/report.html`(gitignore)

- [ ] **Step 1: 검사 스크립트를 쓴다** — `scripts/verify-ribbon-skin-e2e.mjs`

```js
#!/usr/bin/env node
// 아트크래프트 리본 꾸밈 검사(설계 2026-11-22-artcraft-ribbon-skin §4) — 세 편집기(편집·읽기 전용 ui=view)와 보기 화면(view.html):
//  (1) 제목 줄 28·탭 줄 24·리본 90·상태 줄 22px(±2)  (2) 탭·묶음·단추 이름·말풍선이 한글(영문자만 있는 라벨 0)
//  (3) 큰 단추 = 아이콘 위·이름 아래  (4) 모든 탭의 리본 내용이 90px 안(잘림 없음)  (5) "자세히"는 있는 단추를 가리킴(워드 1·엑셀 1·PPT 0)
//  (6) 읽기 전용 흐림이 고르다(0.45), "자세히" 흐림 = 가리키는 단추 흐림  (7) 600px 가로 스크롤 없음·제목 줄 한 줄
//  (8) 보기 화면 같은 규격·id 그대로·첫 쪽이 막대 밑에 깔리지 않음  + 아트크래프트 캡처와 나란히 놓은 보고서
// 쓰는 법: npm run build:editor && node scripts/verify-ribbon-skin-e2e.mjs [포트 8160~8169, 기본 8166]
import { spawn, execFileSync } from 'node:child_process';
import { randomBytes } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { chromium } from 'playwright';
import { bigButtonLayoutOk, heightProblems, latinOnlyLabels } from './ribbonSkinLib.mjs';

const PORT = Number(process.argv[2] ?? 8166);
// 운영 포트(8000·8010·8020·8100·8110·8130·8148)와 다른 시험(8150~8159) 범위와 겹치지 않게
if (!(PORT >= 8160 && PORT <= 8169)) throw new Error(`시험 포트는 8160~8169만 쓴다(받은 값 ${PORT})`);
const BASE = `http://127.0.0.1:${PORT}`;
const RUN = path.resolve('.audit-run', `ribbon-skin-e2e-${PORT}`);
const SHOTS = path.resolve('.audit-run', 'ribbon-skin-shots');
const REF_DIR = process.env.ARTCRAFT_SHOTS_DIR ?? 'C:\\Users\\OpenMinds\\projects\\artcraft-eval\\spike-webapp\\shots';
const FILES = {
  docx: 'audit/office/corpus/docx/a1-korean-report.docx',
  xlsx: 'audit/office/corpus/xlsx/x07-multisheet-print.xlsx',
  pptx: 'audit/office/corpus/pptx/a2-korean-deck.pptx',
};
const READY = { docx: '.ProseMirror', xlsx: '#grid-edit-table td[data-row]', pptx: '#slides-rail .slide-card' };
const REF = { docx: 'word-report.png', xlsx: 'grid-plain.png', pptx: 'deck-korean.png' };
const LAUNCHERS = { docx: 1, xlsx: 1, pptx: 0 };
const EDITOR_SEL = { title: '.office-titlebar', tabs: '.ribbon-tabs', ribbon: '.ribbon-panel:not([hidden])', status: '#statusbar, .grid-status-row' };
const VIEW_SEL = { title: '.toolbar-row-title', tabs: '.ribbon-tabs', ribbon: '.toolbar-row-tools', status: null };
const VIEW_PAGE_IDS = ['back-btn', 'doc-title', 'edit-btn', 'thumbs-btn', 'measure-btn', 'invert-btn', 'markup-add-btn', 'text-tool-btn', 'markup-include-wrap', 'zoom-out', 'zoom-in', 'zoom-fit-btn', 'zoom-level', 'page-prev', 'page-jump-input', 'page-next', 'rotate-btn', 'search-btn', 'print-btn', 'download-btn', 'download-as-btn', 'fullscreen-btn', 'tools-btn', 'tools-menu'];
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

async function upload(bytes, name) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
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
    body: JSON.stringify({ viewer: 'ribbon-skin-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);
async function smallPdf() {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  doc.addPage([400, 300]).drawText('ribbon skin e2e pdf', { x: 40, y: 150, size: 20, font });
  return Buffer.from(await doc.save());
}

async function open(context, url, ready) {
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  await page.goto(`${BASE}${url}`);
  await page.waitForSelector(ready, { timeout: 90000 });
  await page.waitForFunction(() => !document.querySelector('#conn-status') || document.querySelector('#conn-status').getAttribute('data-status') === 'synced', null, { timeout: 30000 }).catch(() => {});
  await sleep(1000);
  return { page, errors };
}
const heights = (page, sel) =>
  page.evaluate((s) => {
    const h = (q) => {
      const el = q && document.querySelector(q);
      return el && el.getClientRects().length > 0 ? Math.round(el.getBoundingClientRect().height) : null;
    };
    return { title: h(s.title), tabs: h(s.tabs), ribbon: h(s.ribbon), status: h(s.status) };
  }, sel);
/** 숨은 탭·판의 것까지 — 글자는 보이는지와 상관없이 한글이어야 한다 */
const labels = (page) =>
  page.evaluate(() => {
    const text = [...document.querySelectorAll('.ribbon-tab, .ribbon-group-label, .icon-btn-big > span, .ribbon-btn-label, .statusbar-item')].map((el) => el.textContent ?? '');
    const tips = [...document.querySelectorAll('#toolbar button, #toolbar select, #toolbar input, #toolbar a')].flatMap((el) => [el.getAttribute('title') ?? '', el.getAttribute('aria-label') ?? '']);
    return [...text, ...tips];
  });
const bigButtons = (page) =>
  page.evaluate(() => {
    const r = (el) => {
      const x = el.getBoundingClientRect();
      return { left: x.left, top: x.top, width: x.width, height: x.height };
    };
    return [...document.querySelectorAll('.ribbon-panel:not([hidden]) .icon-btn-big')]
      .filter((b) => b.getClientRects().length > 0)
      .map((b) => {
        const svg = b.querySelector(':scope > svg');
        const span = b.querySelector(':scope > span');
        return svg && span && span.getClientRects().length > 0 ? { name: b.id || b.title, icon: r(svg), label: r(span) } : null;
      })
      .filter(Boolean);
  });
/** 보이는 탭을 하나씩 눌러 판 높이(90±2)와 잘림(scrollHeight > clientHeight)을 본다 */
async function tabPanels(page) {
  const tabs = await page.$$eval('.ribbon-tab:not(.ribbon-tab-file)', (ts) => ts.filter((t) => !t.hidden && t.dataset.tab).map((t) => t.dataset.tab));
  const bad = [];
  for (const tab of tabs) {
    await page.click(`.ribbon-tab[data-tab="${tab}"]`);
    const m = await page.evaluate(() => {
      const p = document.querySelector('.ribbon-panel:not([hidden])');
      return p ? { h: Math.round(p.getBoundingClientRect().height), sh: p.scrollHeight, ch: p.clientHeight } : null;
    });
    if (!m || Math.abs(m.h - 90) > 2 || m.sh > m.ch + 1) bad.push(`${tab} ${JSON.stringify(m)}`);
  }
  await page.click('.ribbon-tab[data-tab="home"]');
  return { tabs, bad };
}
/** 600px — 문서·리본 가로 넘침, 제목 줄 높이(한 줄) */
const narrowState = (page) =>
  page.evaluate(() => {
    const p = document.querySelector('.ribbon-panel:not([hidden]), .toolbar-row-tools');
    const t = document.querySelector('.office-titlebar, .toolbar-row-title');
    return { doc: document.documentElement.scrollWidth, body: document.body.scrollWidth, panel: p ? p.scrollWidth - p.clientWidth : 0, title: Math.round(t.getBoundingClientRect().height) };
  });
const narrowOk = (n) => n.doc <= 600 && n.body <= 600 && n.panel <= 1 && Math.abs(n.title - 28) <= 2;
const underToolbar = (page) => page.evaluate(() => document.querySelector('.page-frame').getBoundingClientRect().top < document.getElementById('toolbar').getBoundingClientRect().bottom);
async function shot(page, name) {
  await page.screenshot({ path: path.join(SHOTS, `${name}.png`) });
}
function writeReport(pairs) {
  const rows = pairs.map(([name, ref]) => {
    const src = path.join(REF_DIR, ref);
    const has = fs.existsSync(src);
    if (has) fs.copyFileSync(src, path.join(SHOTS, `ref-${ref}`));
    return `<tr><td><img src="${name}.png" alt="${name}"><p>${name}</p></td><td>${has ? `<img src="ref-${ref}" alt="${ref}">` : '(참고 캡처 없음)'}<p>${ref}</p></td></tr>`;
  });
  fs.writeFileSync(
    path.join(SHOTS, 'report.html'),
    `<!doctype html><html lang="ko"><meta charset="utf-8"><title>리본 꾸밈 비교</title><style>body{font-family:"맑은 고딕",sans-serif}img{width:680px;border:1px solid #ccc}td{vertical-align:top;padding:8px}</style><h1>리본 꾸밈 — 우리 화면과 아트크래프트 나란히</h1><table><tr><th>우리 화면(1400×900)</th><th>아트크래프트</th></tr>${rows.join('')}</table></html>`,
  );
}

await startServer();
fs.rmSync(SHOTS, { recursive: true, force: true });
fs.mkdirSync(SHOTS, { recursive: true });
const pairs = [];
let browser;
try {
  const ids = {};
  for (const [ext, file] of Object.entries(FILES)) ids[ext] = (await upload(fs.readFileSync(file), path.basename(file))).id;
  ids.pdf = (await upload(await smallPdf(), 'ribbon-skin.pdf')).id;
  browser = await chromium.launch();

  for (const ext of ['docx', 'xlsx', 'pptx']) {
    // ---- 편집 화면 ----
    const wide = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const edit = await open(wide, (await token(ids[ext], { canEdit: true })).editorUrl, READY[ext]);
    const h = await heights(edit.page, EDITOR_SEL);
    rec(`(1) ${ext} 편집 — 높이 28/24/90/22(±2)`, heightProblems(h).length === 0 && h.status !== null, `${JSON.stringify(h)} ${heightProblems(h).join(', ')}`);
    const latin = latinOnlyLabels(await labels(edit.page));
    rec(`(2) ${ext} 탭·묶음·단추 이름·말풍선이 한글`, latin.length === 0, latin.slice(0, 5).join(' | '));
    const big = await bigButtons(edit.page);
    const bigBad = big.filter((b) => !bigButtonLayoutOk(b.icon, b.label)).map((b) => b.name);
    rec(`(3) ${ext} 홈 탭 큰 단추 ${big.length}개 — 아이콘 위·이름 아래`, big.length > 0 && bigBad.length === 0, bigBad.join(', '));
    const panels = await tabPanels(edit.page);
    rec(`(4) ${ext} 탭 ${panels.tabs.length}개 모두 리본 90px 안(잘림 없음)`, panels.tabs.length >= 5 && panels.bad.length === 0, panels.bad.join(' | '));
    const launchers = await edit.page.$$eval('.ribbon-launcher', (ls) => ls.map((l) => ({ proxy: l.dataset.proxy ?? '', ok: !!document.getElementById(l.dataset.proxy ?? '') && !l.id })));
    rec(`(5) ${ext} "자세히" ${LAUNCHERS[ext]}개 — 기존 단추를 가리키고 id 없음`, launchers.length === LAUNCHERS[ext] && launchers.every((l) => l.ok), JSON.stringify(launchers));
    rec(`${ext} 편집 — 화면 오류 없음`, edit.errors.length === 0, edit.errors.slice(0, 2).join(' | '));
    await shot(edit.page, `editor-${ext}`);
    pairs.push([`editor-${ext}`, REF[ext]]);
    await wide.close();

    // ---- 읽기 전용(ui=view) ----
    const roCtx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const ro = await open(roCtx, (await token(ids[ext], { engine: 'editor' })).viewerUrl, READY[ext]);
    const rh = await heights(ro.page, EDITOR_SEL);
    rec(`(1) ${ext} 읽기 전용 — 같은 높이 규격`, heightProblems(rh).length === 0, `${JSON.stringify(rh)} ${heightProblems(rh).join(', ')}`);
    const grey = await ro.page.evaluate(() => {
      const vis = [...document.querySelectorAll('.ribbon-panel:not([hidden]) button.ro-disabled, .quick-access button.ro-disabled')].filter((b) => b.getClientRects().length > 0);
      const odd = vis.filter((b) => getComputedStyle(b).opacity !== '0.45').map((b) => b.id || b.title);
      const launch = [...document.querySelectorAll('.ribbon-launcher')]
        .filter((l) => l.classList.contains('ro-disabled') !== !!document.getElementById(l.dataset.proxy ?? '')?.classList.contains('ro-disabled'))
        .map((l) => l.dataset.proxy);
      return { count: vis.length, odd, launch, body: document.body.classList.contains('ui-view') };
    });
    rec(`(6) ${ext} 읽기 전용 — 흐린 단추 ${grey.count}개 모두 0.45, "자세히" 흐림 = 대상 흐림`, grey.body && grey.count > 0 && grey.odd.length === 0 && grey.launch.length === 0, JSON.stringify(grey));
    await shot(ro.page, `readonly-${ext}`);
    pairs.push([`readonly-${ext}`, REF[ext]]);
    await roCtx.close();

    // ---- 600px (편집·읽기 전용) ----
    for (const [mode, body, key] of [
      ['편집', { canEdit: true }, 'editorUrl'],
      ['읽기 전용', { engine: 'editor' }, 'viewerUrl'],
    ]) {
      const narrow = await browser.newContext({ viewport: { width: 600, height: 900 } });
      const n = await open(narrow, (await token(ids[ext], body))[key], READY[ext]);
      const tabs = await n.page.$$eval('.ribbon-tab:not(.ribbon-tab-file)', (ts) => ts.filter((t) => !t.hidden && t.dataset.tab).map((t) => t.dataset.tab));
      const bad = [];
      for (const tab of tabs) {
        await n.page.click(`.ribbon-tab[data-tab="${tab}"]`);
        const s = await narrowState(n.page);
        if (!narrowOk(s)) bad.push(`${tab} ${JSON.stringify(s)}`);
      }
      rec(`(7) ${ext} ${mode} 600px — 탭 ${tabs.length}개 모두 가로 넘침 없음·제목 줄 한 줄`, tabs.length >= 5 && bad.length === 0, bad.join(' | '));
      await narrow.close();
    }
  }

  // ---- 보기 화면(view.html): 리브레오피스로 그린 docx, pdf ----
  for (const ext of ['docx', 'pdf']) {
    const tok = await token(ids[ext], {});
    const viewPath = new URL(tok.viewerUrl, BASE).pathname;
    const wide = await browser.newContext({ viewport: { width: 1400, height: 900 } });
    const v = await open(wide, tok.viewerUrl, '.page-frame.page-ready');
    const vh = await heights(v.page, VIEW_SEL);
    rec(`(8) 보기 ${ext} — ${viewPath}, 높이 28/24/90(상태 줄 없음)`, viewPath === '/view.html' && heightProblems(vh).length === 0 && vh.title !== null && vh.tabs !== null && vh.ribbon !== null, `${JSON.stringify(vh)} ${heightProblems(vh).join(', ')}`);
    const vl = latinOnlyLabels(await labels(v.page));
    const groups = await v.page.$$eval('#toolbar .ribbon-group-label', (ls) => ls.map((l) => l.textContent));
    rec(`(8) 보기 ${ext} — 묶음 표시·보기·쪽·찾기·인쇄·내려받기, 라벨 한글`, vl.length === 0 && groups.join(',') === '표시,보기,쪽,찾기,인쇄,내려받기', `${groups.join(',')} ${vl.slice(0, 3).join(' | ')}`);
    const vs = await v.page.evaluate((list) => ({
      missing: list.filter((id) => !document.getElementById(id)),
      printRight: !!document.getElementById('print-btn')?.closest('.toolbar-tools-right'),
    }), VIEW_PAGE_IDS);
    const under = await underToolbar(v.page);
    rec(`(8) 보기 ${ext} — id 그대로·인쇄는 오른쪽 칸·첫 쪽이 막대 밑에 깔리지 않음`, vs.missing.length === 0 && vs.printRight && !under, JSON.stringify({ ...vs, under }));
    rec(`보기 ${ext} — 화면 오류 없음`, v.errors.length === 0, v.errors.slice(0, 2).join(' | '));
    await shot(v.page, `view-${ext}`);
    pairs.push([`view-${ext}`, 'word-report.png']);
    await wide.close();

    const narrow = await browser.newContext({ viewport: { width: 600, height: 900 } });
    const n = await open(narrow, tok.viewerUrl, '.page-frame.page-ready');
    const ns = await narrowState(n.page);
    const nUnder = await underToolbar(n.page);
    rec(`(7) 보기 ${ext} 600px — 가로 넘침 없음·제목 줄 한 줄·첫 쪽이 막대 밑에 깔리지 않음`, narrowOk(ns) && !nUnder, JSON.stringify({ ...ns, under: nUnder }));
    await narrow.close();
  }
  writeReport(pairs);
  console.log(`나란히 비교 보고서: ${path.join(SHOTS, 'report.html')}`);
} catch (error) {
  rec('스크립트 오류', false, String(error?.stack ?? error));
} finally {
  await browser?.close().catch(() => {});
  stopServer();
}
const failed = results.filter((ok) => !ok).length;
console.log(`\n${results.length - failed}/${results.length} 통과`);
if (!failed) fs.rmSync(RUN, { recursive: true, force: true, maxRetries: 5, retryDelay: 500 });
else console.log(`실패 — 서버 기록·데이터는 ${RUN}에 남겨 둔다`);
process.exit(failed ? 1 : 0);
```

- [ ] **Step 2: package.json에 명령을 단다** — `"verify:artcraft": "node scripts/verify-artcraft-e2e.mjs",` 줄 바로 뒤에:

```json
    "verify:ribbon-skin": "node scripts/verify-ribbon-skin-e2e.mjs",
```

- [ ] **Step 3: 단위 시험 전체·타입·빌드**

Run: `npm test && npm run typecheck && npm run build:editor`
Expected: 전부 통과(새 `tests/ribbonSkinLib.test.mjs` 포함), 타입 오류 0

- [ ] **Step 4: 새 검사를 돌린다**

Run: `npm run verify:ribbon-skin -- 8166`
Expected: 마지막 줄 `N/N 통과`(약 40항목)와 `나란히 비교 보고서: …\.audit-run\ribbon-skin-shots\report.html`. 실패하면 해당 태스크의 CSS·마크업을 고친다 — 검사와 규격 숫자는 고치지 않는다(설계 값). 흔한 원인: 어떤 탭의 판이 90px을 넘거나 잘림(→ 그 묶음 칸 높이를 `--skin-small-h` 세 줄·`--skin-big-icon` 안으로), 600px 넘침(→ 그 묶음의 넓은 `select`·`input`에 `.ribbon-skin` 미디어 쿼리 안 `max-width`), 보기 화면 첫 쪽이 막대 밑(→ `--view-chrome-h` 값).

- [ ] **Step 5: 기존 e2e — 스스로 서버를 띄우는 것(스크립트는 고치지 않는다)**

Run: `npm run verify:readonly -- 8162`
Expected: `85/85 통과`(`ARTCRAFT_E2E_PRIVATE_DIR`가 있으면 고객형 2건 더해 87)

Run: `npm run verify:artcraft -- 8156`
Expected: 마지막 줄 전부 통과(스크립트가 선행 조건 — 아트크래프트 엔진 산출물 — 이 없다고 하면 그 안내대로 만든 뒤 다시)

- [ ] **Step 6: 기존 e2e — 띄워 둔 서버가 필요한 것.** 수동·회귀용 서버를 8165에 띄우고 PID를 적어 둔다(PowerShell):

```powershell
New-Item -ItemType Directory -Force .audit-run\skin-regress | Out-Null
$env:DOCVIEW_PORT = '8165'; $env:DOCVIEW_HOST = '127.0.0.1'; $env:DOCVIEW_SECRET = [guid]::NewGuid().ToString('N'); $env:DOCVIEW_DATA_DIR = "$PWD\.audit-run\skin-regress\data"
$p = Start-Process -FilePath node -ArgumentList 'node_modules/tsx/dist/cli.mjs', 'src/server.ts' -PassThru -WindowStyle Hidden -RedirectStandardOutput .audit-run\skin-regress\server.log -RedirectStandardError .audit-run\skin-regress\server.err
$p.Id | Out-File -Encoding ascii .audit-run\skin-regress\pid.txt
```

그다음 차례로(모두 같은 서버):

Run: `node scripts/verify-docx-edit-e2e.mjs http://127.0.0.1:8165`
Expected: 전부 통과

Run: `node scripts/verify-pptx-fidelity-e2e.mjs http://127.0.0.1:8165`
Expected: `ALL CHECKS PASSED`

Run: `node scripts/verify-xlsx-view.mjs http://127.0.0.1:8165`
Expected: 전부 통과

Run: `node scripts/verify-view-only-e2e.mjs http://127.0.0.1:8165`
Expected: `ALL CHECKS PASSED`

Run: `node scripts/verify-view-shell-e2e.mjs http://127.0.0.1:8165`
Expected: `ALL CHECKS PASSED`

- [ ] **Step 7: 실제 브라우저로 조작해 본다(같은 8165 서버, 사람 눈 확인).** 메인 화면(`http://127.0.0.1:8165/`)에서 docx·xlsx·pptx·pdf를 올리고 각각 "열기"(읽기 전용)·"동시편집 열기"·"리브레오피스로 열기"로 연다. 편집기마다: 모든 탭을 누른다 → 모든 ▾ 드롭다운(글꼴·크기·색·줄 간격·음영·표 삽입·여백·용지 방향·크기·표 서식·조건부 서식·정렬·차트)을 열어 단추 바로 아래 제자리에 뜨는지, 큰 단추·작은 단추·입력칸을 하나씩 눌러 동작이 전과 같은지, "자세히"가 단어 개수·사용자 지정 정렬 대화상자를 여는지, 리본 축소(Ctrl+F1·탭 두 번 클릭)와 축소 중 탭을 눌러 잠깐 펼치기(펼친 판이 탭 줄 바로 아래), 파일 화면, 상태 줄 확대 단추·슬라이더. 읽기 전용: 같은 화면에서 쓰는 단추만 고르게 흐리고 "편집 열기"가 앱 색 알약. 보기 화면: 확대·맞춤·쪽 이동·회전·찾기·인쇄·내려받기·썸네일·전체 화면, 창 폭 1024·800·600·390에서 단추가 도구▾ 메뉴로 갔다가 제 묶음으로 돌아오는지. 600px에서 편집기 리본이 아이콘만 남아 줄 바꿈되고 가로 스크롤이 없는지. 마지막으로 `.audit-run\ribbon-skin-shots\report.html`을 열어 아트크래프트 캡처와 나란히 비교한다.

- [ ] **Step 8: 8165 서버를 PID로만 끈다**

```powershell
taskkill /PID (Get-Content .audit-run\skin-regress\pid.txt) /T /F
```

- [ ] **Step 9: 커밋(보고서·캡처·`.audit-run/`·번들은 커밋하지 않는다)**

```bash
git add scripts/verify-ribbon-skin-e2e.mjs package.json
git commit -m "리본 꾸밈 6: verify:ribbon-skin — 높이 28/24/90/22·한글 라벨·큰 단추 배치·탭별 잘림·자세히·읽기 전용 흐림·600px·보기 화면 검사와 아트크래프트 나란히 비교 보고서

Co-Authored-By: Claude Fable 5.1 <noreply@anthropic.com>"
```

- [ ] **Step 10: 보고** — 새 검사 결과(N/N), 기존 e2e 7개 결과, `report.html`의 캡처 쌍 8개(편집·읽기 전용 × 워드·엑셀·PPT, 보기 docx·pdf)를 사용자에게 보인다. 운영(8000·8010) 반영은 사용자 지시가 있을 때만.

---

## 설계와 맞춰 본 결과(자체 검토)

- §1 기준 디자인: 제목 줄·탭 줄·리본·묶음·상태 줄 규격 → Task 1. 워드 눈금자는 지금 없는 기능이라 만들지 않음(§3). 엑셀 이름 상자+수식 줄 → Task 3에서 유지 확인. PPT 섬네일 레일 → Task 4.
- §2-1 모양 → Task 1(공용)·2·3·4. §2-2 배치만(없는 묶음 안 만듦) → Task 2(워드 삽입·레이아웃 순서), Task 4(PPT 삽입 순서), 엑셀은 이미 같은 순서. §2-3 한글 → 각 태스크 Step 1 검사 + Task 6 (2). §2-4 읽기 전용 → Task 1 Step 7(자세히 판정) + Task 6 (6) + `verify:readonly`. §2-5 보기 화면 → Task 5 + Task 6 (8). §2-6 600px → Task 1 미디어 쿼리 + Task 6 (7).
- §4 검사 → Task 6 Step 4~7(새 검사·기존 e2e 7개·실제 브라우저·나란히 비교 보고서).
- 계획이 정한 것(설계에 값이 없어서): "자세히"는 기존 대화상자가 있는 두 묶음에만(워드 언어 교정 → 단어 개수, 엑셀 정렬 및 필터 → 사용자 지정 정렬; PPT 없음). 보기 화면의 측정·강조·글 넣기는 설계 묶음 목록에 없어 "표시" 묶음으로, `#edit-btn`은 제목 줄에 그대로(390px 배치·view-shell 검사 유지) — 그래서 "편집 열기" 묶음은 만들지 않는다. 보기 화면에는 상태 줄이 없어 상태 줄 높이는 편집기에서만 잰다. PPT 앱 색은 기존 `#c0462a` 그대로.
- 참고 캡처(1400×900)를 재 보면 제목 줄 약 24·탭 줄 약 22·리본 약 62·상태 줄 약 18px이다 — 설계 §1·§4의 28/24/90/22와 다르지만 설계가 구속력이 있으므로 설계 값을 따른다(사용자 확인 대상).
