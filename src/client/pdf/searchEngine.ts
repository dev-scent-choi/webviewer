/**
 * PDF 찾기 패널의 3단 검색 — (1) 이 페이지 (2) 전체 문서 (3) 결과 안에서(앞 단계 결과가 든 블록만 다시 찾기).
 * 찾을 곳은 모든 글 / 제목만. 결과마다 쪽·속한 제목·앞뒤 문장을 붙인다. DOM 없는 순수 함수(tests/pdfSearchEngine.test.ts).
 */
import { findInPage, normalizeForFind, type FindMatch } from './findText.js';
import type { PageStructure, StructItem } from './textStructure.js';

export type SearchScope = 'page' | 'doc' | 'within';
export type SearchTarget = 'all' | 'headings';

export interface SearchResult {
  readonly match: FindMatch;
  readonly page: number;
  readonly block: number;
  readonly heading: string;
  readonly isHeading: boolean;
  readonly before: string;
  readonly hit: string;
  readonly after: string;
}

/** 블록을 가리키는 열쇠 — "결과 안에서" 단계가 이 집합 안에서만 찾는다 */
export const blockKey = (page: number, block: number) => `${page}:${block}`;

const SNIPPET = 28;

/** 블록 글에서 n번째 일치를 앞뒤 문장과 함께 — 줄바꿈은 한 칸으로 */
function snippetOf(text: string, query: string, nth: number): { before: string; hit: string; after: string } {
  const flat = text.replace(/\s+/g, ' ');
  const norm = normalizeForFind(flat);
  const q = normalizeForFind(query).trim();
  let at = -1;
  for (let k = 0, from = 0; k <= nth; k += 1) {
    const found = norm.indexOf(q, from);
    if (found < 0) break;
    at = found;
    from = found + q.length;
  }
  // 정규화로 글자 수가 달라지는 드문 경우(호환 문자)엔 위치가 어긋날 수 있다 — 그래도 앞뒤 문맥은 보인다
  if (at < 0 || norm.length !== flat.length) {
    const i = norm.indexOf(q);
    if (i < 0) return { before: '', hit: flat.slice(0, SNIPPET * 2), after: '' };
    at = i;
  }
  const start = Math.max(0, at - SNIPPET);
  const end = Math.min(flat.length, at + q.length + SNIPPET);
  return {
    before: (start > 0 ? '…' : '') + flat.slice(start, at),
    hit: flat.slice(at, at + q.length),
    after: flat.slice(at + q.length, end) + (end < flat.length ? '…' : ''),
  };
}

/** 쪽 하나에서 찾기 — 블록 정보를 붙이고 찾을 곳·허용 블록으로 거른다 */
export function searchPage(
  structure: PageStructure,
  items: readonly StructItem[],
  query: string,
  target: SearchTarget,
  allowed: ReadonlySet<string> | null,
): SearchResult[] {
  const matches = findInPage(structure.page, items, query);
  const out: SearchResult[] = [];
  const seenInBlock = new Map<number, number>();
  for (const m of matches) {
    const block = structure.blockOfItem[m.itemStart] ?? -1;
    const b = structure.blocks[block];
    if (!b) continue;
    const nth = seenInBlock.get(block) ?? 0;
    seenInBlock.set(block, nth + 1);
    if (target === 'headings' && !b.isHeading) continue;
    if (allowed && !allowed.has(blockKey(structure.page, block))) continue;
    out.push({ match: m, page: structure.page, block, heading: b.heading, isHeading: b.isHeading, ...snippetOf(b.text, query, nth) });
  }
  return out;
}

/** "결과 안에서" 단계가 쓸 블록 집합 — 지금 결과가 든 블록들 */
export function blocksOf(results: readonly SearchResult[]): Set<string> {
  return new Set(results.map((r) => blockKey(r.page, r.block)));
}

/** 결과를 쪽별로 묶는다(목록 머리에 "3쪽 · 2건") */
export function groupByPage(results: readonly SearchResult[]): { page: number; results: SearchResult[] }[] {
  const groups: { page: number; results: SearchResult[] }[] = [];
  for (const r of results) {
    const last = groups[groups.length - 1];
    if (last && last.page === r.page) last.results.push(r);
    else groups.push({ page: r.page, results: [r] });
  }
  return groups;
}
