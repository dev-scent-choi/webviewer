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
