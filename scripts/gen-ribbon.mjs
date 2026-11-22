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
