// 오피스 리본 목록(src/client/ribbon/inventory.json)과 리본 마크업 만들기·검사(설계 2026-11-22-office-ribbon-structure).
// scripts/gen-ribbon.mjs, tests/ribbon*.test.ts, scripts/verify-ribbon-inventory-e2e.mjs·
// verify-readonly-editor-e2e.mjs·verify-ribbon-skin-e2e.mjs가 함께 쓴다. 브라우저에서는 쓰지 않는다(node 전용).
import fs from 'node:fs';

export const INVENTORY_PATH = 'src/client/ribbon/inventory.json';
export const APP_OF_EXT = Object.freeze({ docx: 'word', xlsx: 'excel', pptx: 'powerpoint' });
/** 리본을 목록으로 만든 편집기 확장자 — e2e 단추 수 단언이 목록을 따르는 앱(Task 3 워드·4 엑셀·5 PPT가 하나씩 더함) */
export const LISTED_EXTS = new Set(['docx', 'xlsx']);
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
