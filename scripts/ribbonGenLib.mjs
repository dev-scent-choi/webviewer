// 오피스 리본 목록(src/client/ribbon/inventory.json)과 리본 마크업 만들기·검사(설계 2026-11-22-office-ribbon-structure).
// scripts/gen-ribbon.mjs, tests/ribbon*.test.ts, scripts/verify-ribbon-inventory-e2e.mjs·
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
      if (tab.unverified !== undefined && tab.unverified !== true) say(`unverified는 true만 ${tab.id}`);
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
        if (g.unverified !== undefined && g.unverified !== true) say(`unverified는 true만 ${at}`);
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
          if (c.unverified !== undefined && c.unverified !== true) say(`unverified는 true만 ${k}`);
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
