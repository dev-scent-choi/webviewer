// PDF 찾기 패널의 제목·본문 판별(src/client/pdf/textStructure.ts)을 실제 PDF에 돌려 쪽마다 결과를 뽑는다 — 판별 기준을
// 실제 문서로 검증·조정할 때 쓴다(편집기와 같은 pdf.js로 글 조각·글꼴·태그 구조를 읽는다).
// 사용: npx tsx scripts/analyze-pdf-structure.ts <pdf> [쪽 범위 예: 1-40] [--all-lines] [--eval] [--find 검색어 [--within 검색어]]
//   --find: 찾기 패널과 같은 엔진으로 전체·제목·본문에서 찾은 결과를 뽑는다(--within: 그 결과가 속한 절 안에서 다시 찾기)
//   --eval: 책갈피가 있는 문서에서 책갈피 제목을 정답으로 삼아(판별에는 책갈피를 쓰지 않고) 찾은 비율·정확도를 잰다
import fs from 'node:fs';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { normalizeForFind } from '../src/client/pdf/findText.js';
import { searchPage, sectionsOf, type SearchResult, type SearchTarget } from '../src/client/pdf/searchEngine.js';
import { analyzePage, buildProfile, isBoldFontName, type PageStructure, type StructItem } from '../src/client/pdf/textStructure.js';

const [file, range = '1-20', ...flags] = process.argv.slice(2);
const allLines = flags.includes('--all-lines');
const evaluate = flags.includes('--eval');
const flagValue = (name: string) => (flags.includes(name) ? flags[flags.indexOf(name) + 1] : undefined);
const findQuery = flagValue('--find');
const withinQuery = flagValue('--within');
const analyzed: { items: StructItem[]; structure: PageStructure }[] = [];
const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(file)), cMapUrl: 'node_modules/pdfjs-dist/cmaps/', cMapPacked: true, standardFontDataUrl: 'node_modules/pdfjs-dist/standard_fonts/', isEvalSupported: false, verbosity: 0 }).promise;
const [from, to] = range === 'all' ? [1, doc.numPages] : range.split('-').map(Number);

async function pageItems(n: number): Promise<{ items: StructItem[]; pageHeight: number }> {
  const p = await doc.getPage(n);
  await p.getOperatorList();
  const content = await p.getTextContent({ includeMarkedContent: true });
  const roleOf = new Map<string, string>();
  const tree = await p.getStructTree().catch(() => null);
  const walk = (node: unknown, role: string | null) => {
    const nd = node as { role?: string; type?: string; id?: string; children?: unknown[] } | null;
    if (!nd) return;
    if (nd.type === 'content' && nd.id && role) roleOf.set(nd.id, role);
    const next = nd.role && !['Root', 'Document', 'Part', 'Sect', 'Div'].includes(nd.role) ? nd.role : role;
    for (const c of nd.children ?? []) walk(c, next);
  };
  walk(tree, null);
  const items: StructItem[] = [];
  const stack: string[] = [];
  for (const item of content.items) {
    const it = item as { type?: string; id?: string; str?: string; fontName?: string; transform?: number[]; width?: number; height?: number };
    if (it.type === 'beginMarkedContentProps' || it.type === 'beginMarkedContent') stack.push(it.id ?? '');
    else if (it.type === 'endMarkedContent') stack.pop();
    else if (typeof it.str === 'string') {
      let role: string | null = null;
      for (let k = stack.length - 1; k >= 0 && !role; k -= 1) role = roleOf.get(stack[k]) ?? null;
      let bold = false;
      try {
        const f = p.commonObjs.get(it.fontName!) as { name?: string; bold?: boolean; black?: boolean };
        bold = !!(f?.bold || f?.black) || isBoldFontName(f?.name);
      } catch {
        /* 없음 */
      }
      items.push({ str: it.str, transform: it.transform!, width: it.width!, height: it.height!, fontName: it.fontName, bold, role });
    }
  }
  return { items, pageHeight: p.getViewport({ scale: 1 }).height };
}

// 문서 기준: 고르게 30쪽 표본
const sampleN = Math.min(30, doc.numPages);
const sample = [];
for (let k = 0; k < sampleN; k += 1) sample.push(await pageItems(1 + Math.floor((k * doc.numPages) / sampleN)));
const profile = buildProfile(sample);
console.log(`본문 서식 ${profile.bodyStyle} 크기 ${profile.bodySize} 줄간격 ${profile.lineGap.toFixed(2)} 반복 머리말 ${[...profile.repeated].slice(0, 5).join(' / ')}`);

// 책갈피(정답지)
const truth = new Map<number, string[]>();
if (evaluate) {
  const outline = (await doc.getOutline()) ?? [];
  const walk = async (nodes: typeof outline) => {
    for (const n of nodes) {
      let dest = n.dest;
      if (typeof dest === 'string') dest = await doc.getDestination(dest);
      if (Array.isArray(dest) && dest[0]) {
        const page = (await doc.getPageIndex(dest[0])) + 1;
        const list = truth.get(page) ?? [];
        list.push(normalizeForFind(n.title).trim());
        truth.set(page, list);
      }
      if (n.items?.length) await walk(n.items);
    }
  };
  await walk(outline);
}

let prevHeading = '';
let prevSection = '';
let tp = 0;
let fp = 0;
let fn = 0;
const fps: string[] = [];
const fns: string[] = [];
for (let n = from; n <= Math.min(to, doc.numPages); n += 1) {
  const { items, pageHeight } = await pageItems(n);
  const s = analyzePage(n, items, prevHeading, new Set(), { pageHeight, profile }, prevSection);
  prevHeading = s.lastHeading;
  prevSection = s.lastSection;
  if (findQuery) {
    analyzed.push({ items, structure: s });
    continue;
  }
  if (evaluate) {
    const expected = truth.get(n) ?? [];
    const found = s.blocks.filter((b) => b.isHeading).map((b) => normalizeForFind(b.text.replace(/\n/g, ' ')).trim());
    const strip = (t: string) => t.replace(/\s+/g, '');
    for (const e of expected) {
      if (found.some((f) => strip(f) === strip(e) || strip(f).endsWith(strip(e)) || strip(e).endsWith(strip(f)))) tp += 1;
      else {
        fn += 1;
        fns.push(`${n}쪽: ${e}`);
      }
    }
    for (const f of found) if (!expected.some((e) => strip(f) === strip(e) || strip(f).endsWith(strip(e)) || strip(e).endsWith(strip(f)))) {
      fp += 1;
      fps.push(`${n}쪽: ${f.slice(0, 60)}`);
    }
    continue;
  }
  console.log(`\n=== ${n}쪽 (블록 ${s.blocks.length})`);
  for (const b of s.blocks) {
    const first = b.text.split('\n')[0];
    const lines = b.text.split('\n').length;
    if (b.toc) {
      if (allLines) console.log(`      #목차 ${first.slice(0, 50)}`);
    } else if (b.furniture) {
      if (allLines) console.log(`      ~머리말 ${first.slice(0, 40)}`);
    } else if (b.isHeading) console.log(`  [${b.level === 1 ? '제목' : '소제'} ${b.score}] ${first.slice(0, 70)}${lines > 1 ? ` …(${lines}줄)` : ''}`);
    else if (allLines) console.log(`      (${b.score}) ${first.slice(0, 60)}${lines > 1 ? ` …(${lines}줄)` : ''}`);
  }
}
if (evaluate) {
  console.log(`\n책갈피 제목 ${tp + fn}개 중 찾음 ${tp} (찾은 비율 ${((tp / Math.max(1, tp + fn)) * 100).toFixed(0)}%), 책갈피에 없는 제목 판정 ${fp}개 (정확도 ${((tp / Math.max(1, tp + fp)) * 100).toFixed(0)}%)`);
  console.log('놓친 예:', fns.slice(0, 15).join(' | '));
  console.log('책갈피에 없는 제목 예:', fps.slice(0, Number(process.env.FP_SHOW ?? 25)).join(' | '));
}
if (findQuery) {
  const find = (q: string, target: SearchTarget, allowed: ReadonlySet<string> | null) => analyzed.flatMap(({ items, structure }) => searchPage(structure, items, q, target, allowed));
  const show = (label: string, rs: SearchResult[]) => {
    console.log(`
--- ${label}: ${rs.length}건`);
    for (const r of rs.slice(0, Number(process.env.SHOW ?? 40))) console.log(`  ${r.page}쪽 ${r.isHeading ? '[제목]' : '      '} ‹${r.heading.slice(0, 30)}› ${r.before}[${r.hit}]${r.after}`);
  };
  const all = find(findQuery, 'all', null);
  const headings = find(findQuery, 'headings', null);
  const body = find(findQuery, 'body', null);
  show(`전체 '${findQuery}'`, all);
  show('제목만', headings);
  show('본문만', body);
  console.log(`
전체 ${all.length} = 제목 ${headings.length} + 본문 ${body.length} + 목차 ${all.length - headings.length - body.length}`);
  if (withinQuery) show(`'${findQuery}' 결과가 속한 절(${sectionsOf(all).size}개) 안에서 '${withinQuery}'`, find(withinQuery, 'all', sectionsOf(all)));
}
