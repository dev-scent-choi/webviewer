/**
 * PDF 찾기 패널의 3단 검색 — (1) 이 페이지 (2) 전체 문서 (3) 결과 안에서(앞 단계 결과가 속한 절에서만 다시 찾기).
 * 찾을 곳은 전체 / 제목만 / 본문만(목차는 '전체'에서만). 결과마다 쪽·속한 제목(절)·앞뒤 문장을 붙인다.
 * 머리말·꼬리말·쪽 번호는 결과에서 뺀다.
 * DOM 없는 순수 함수(tests/pdfSearchEngine.test.ts).
 */
import { findInPage, normalizeForFind, type FindMatch } from './findText.js';
import type { PageStructure, StructItem } from './textStructure.js';

export type SearchScope = 'page' | 'doc' | 'within';
export type SearchTarget = 'all' | 'headings' | 'body';

export interface SearchResult {
  readonly match: FindMatch;
  readonly page: number;
  readonly block: number;
  readonly heading: string;
  /** 속한 절 열쇠 — 제목 블록의 "쪽:블록"(첫 제목 앞의 글이면 "b:쪽:블록") */
  readonly section: string;
  readonly isHeading: boolean;
  /** 제목 단계 — 1 큰 제목, 2 소제목, 본문 0 */
  readonly level: 0 | 1 | 2;
  readonly before: string;
  readonly hit: string;
  readonly after: string;
}

/** 결과가 속한 절 열쇠 — 제목 앞의 글은 그 블록만 */
export function sectionKeyOf(structure: PageStructure, block: number): string {
  const b = structure.blocks[block];
  return b?.section || `b:${structure.page}:${block}`;
}

const SNIPPET = 28;

/** 블록 글에서 n번째 일치를 앞뒤 문장과 함께 — 줄바꿈은 한 칸으로 */
function snippetOf(text: string, query: string, nth: number): { at: number; before: string; hit: string; after: string } {
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
    if (i < 0) return { at: -1, before: '', hit: flat.slice(0, SNIPPET * 2), after: '' };
    at = i;
  }
  const start = Math.max(0, at - SNIPPET);
  const end = Math.min(flat.length, at + q.length + SNIPPET);
  return {
    at,
    before: (start > 0 ? '…' : '') + flat.slice(start, at),
    hit: flat.slice(at, at + q.length),
    after: flat.slice(at + q.length, end) + (end < flat.length ? '…' : ''),
  };
}

/** 쪽 하나에서 찾기 — 블록·절 정보를 붙이고 찾을 곳·허용 절로 거른다 */
export function searchPage(
  structure: PageStructure,
  items: readonly StructItem[],
  query: string,
  target: SearchTarget,
  allowedSections: ReadonlySet<string> | null,
): SearchResult[] {
  const matches = findInPage(structure.page, items, query);
  const out: SearchResult[] = [];
  const seenInBlock = new Map<number, number>();
  for (const m of matches) {
    const block = structure.blockOfItem[m.itemStart] ?? -1;
    const b = structure.blocks[block];
    if (!b || b.furniture) continue;
    const nth = seenInBlock.get(block) ?? 0;
    seenInBlock.set(block, nth + 1);
    // 목차는 제목도 본문도 아니다 — '전체'에서만
    if (target !== 'all' && b.toc) continue;
    const { at, ...snippet } = snippetOf(b.text, query, nth);
    // 본문 첫머리의 조 제목('제3조(목적) 이 규정은…')에 걸린 일치는 제목으로 친다
    const inTitle = b.isHeading || (at >= 0 && at < b.inlineTitle);
    if (target === 'headings' && !inTitle) continue;
    if (target === 'body' && inTitle) continue;
    const section = sectionKeyOf(structure, block);
    if (allowedSections && !allowedSections.has(section)) continue;
    out.push({ match: m, page: structure.page, block, heading: b.heading, section, isHeading: inTitle, level: b.isHeading ? b.level : inTitle ? 2 : 0, ...snippet });
  }
  return out;
}

/** "결과 안에서" 단계가 쓸 절 집합 — 지금 결과가 속한 절들 */
export function sectionsOf(results: readonly SearchResult[]): Set<string> {
  return new Set(results.map((r) => r.section));
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

/**
 * 절 하나의 글 — 제목 블록부터 다음 절이 시작하기 전까지(쪽을 넘어가도 이어서). structures 는 쪽 번호 → 구조.
 * 제목 앞의 글(b:쪽:블록)은 그 블록만.
 */
export function sectionText(section: string, structures: ReadonlyMap<number, PageStructure>): { heading: string; text: string; pages: number[] } {
  if (section.startsWith('b:')) {
    const [, p, b] = section.split(':').map(Number);
    const block = structures.get(p)?.blocks[b];
    // 목차 줄은 heading 이 '목차', 첫 제목 앞의 글은 ''
    return { heading: block?.heading ?? '', text: block?.text ?? '', pages: [p] };
  }
  const [startPage] = section.split(':').map(Number);
  const parts: string[] = [];
  const pages: number[] = [];
  let heading = '';
  for (let page = startPage; structures.has(page); page += 1) {
    const s = structures.get(page)!;
    let any = false;
    let ended = false;
    for (const b of s.blocks) {
      if (b.furniture) continue;
      if (b.section === section) {
        if (!heading) heading = b.heading;
        parts.push(b.text);
        any = true;
      } else if (any || page > startPage) {
        ended = true;
        break;
      }
    }
    if (any) pages.push(page);
    if (ended) break;
  }
  return { heading, text: parts.join('\n\n'), pages };
}
