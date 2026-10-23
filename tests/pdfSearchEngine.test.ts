import { describe, expect, test } from 'vitest';
import { analyzePage, type StructItem } from '../src/client/pdf/textStructure.js';
import { blocksOf, groupByPage, searchPage } from '../src/client/pdf/searchEngine.js';

/** 한 줄 = 조각 하나. y 는 위에서부터 내려가며(PDF 좌표는 위로 커지므로 800 - n) */
const line = (str: string, y: number, size = 10, x = 50): StructItem => ({ str, transform: [size, 0, 0, size, x, y], width: str.length * size * 0.5, height: size });

// 1쪽: 큰 제목 → 본문 두 줄(한 문단) → 번호 제목 → 본문
const page1: StructItem[] = [
  line('설비 운전 지침', 780, 16),
  line('냉각수 펌프는 매일 점검합니다.', 750),
  line('펌프 소음이 크면 정지합니다.', 738),
  line('1. 배관 관리', 700, 10.5),
  line('배관 누수는 즉시 보고합니다.', 680),
];
// 2쪽: 제목 없이 본문만 — 앞 쪽 제목을 이어받는다
const page2: StructItem[] = [line('예비 펌프로 바꿔 가동합니다.', 780), line('배관 압력을 기록합니다.', 740)];

describe('analyzePage', () => {
  test('줄을 문단으로 묶고 큰 글자·번호 붙은 짧은 줄을 제목으로 본다', () => {
    const s = analyzePage(1, page1);
    expect(s.blocks.map((b) => [b.isHeading, b.text])).toEqual([
      [true, '설비 운전 지침'],
      [false, '냉각수 펌프는 매일 점검합니다.\n펌프 소음이 크면 정지합니다.'],
      [true, '1. 배관 관리'],
      [false, '배관 누수는 즉시 보고합니다.'],
    ]);
    expect(s.blocks.map((b) => b.heading)).toEqual(['설비 운전 지침', '설비 운전 지침', '1. 배관 관리', '1. 배관 관리']);
    expect(s.blockOfItem).toEqual([0, 1, 1, 2, 3]);
    expect(s.lastHeading).toBe('1. 배관 관리');
    const [x0, y0, x1, y1] = s.blocks[1].bbox;
    expect(x0).toBe(50);
    expect(y0).toBeLessThan(738);
    expect(y1).toBeGreaterThan(750);
    expect(x1).toBeGreaterThan(x0);
  });

  test('앞 쪽 제목을 이어받고, 책갈피 제목과 같은 줄은 크기가 같아도 제목으로 본다', () => {
    const s = analyzePage(2, page2, '1. 배관 관리');
    expect(s.blocks.every((b) => !b.isHeading && b.heading === '1. 배관 관리')).toBe(true);
    const withOutline = analyzePage(2, page2, '', new Set(['배관 압력을 기록합니다.']));
    expect(withOutline.blocks.find((b) => b.text === '배관 압력을 기록합니다.')?.isHeading).toBe(true);
  });

  test('같은 높이의 조각은 한 줄 — 떨어져 있으면 한 칸 띄운다', () => {
    const s = analyzePage(1, [
      { str: 'P-101', transform: [10, 0, 0, 10, 50, 500], width: 25, height: 10 },
      { str: '(원심펌프)', transform: [10, 0, 0, 10, 90, 500.5], width: 40, height: 10 },
    ]);
    expect(s.blocks).toHaveLength(1);
    expect(s.blocks[0].text).toBe('P-101 (원심펌프)');
  });
});

describe('searchPage — 3단 검색', () => {
  const s1 = analyzePage(1, page1);
  const s2 = analyzePage(2, page2, s1.lastHeading);

  test('모든 글: 쪽·제목·앞뒤 문장', () => {
    const r = [...searchPage(s1, page1, '펌프', 'all', null), ...searchPage(s2, page2, '펌프', 'all', null)];
    expect(r.map((x) => [x.page, x.heading])).toEqual([
      [1, '설비 운전 지침'],
      [1, '설비 운전 지침'],
      [2, '1. 배관 관리'],
    ]);
    expect(r[1].before.endsWith('점검합니다. ')).toBe(true);
    expect(r[1].hit).toBe('펌프');
    expect(r[1].after.startsWith(' 소음이')).toBe(true);
    expect(groupByPage(r).map((g) => [g.page, g.results.length])).toEqual([[1, 2], [2, 1]]);
  });

  test('제목만: 제목 블록 안의 결과만', () => {
    expect(searchPage(s1, page1, '배관', 'headings', null).map((x) => x.isHeading)).toEqual([true]);
    expect(searchPage(s1, page1, '배관', 'all', null)).toHaveLength(2);
  });

  test('결과 안에서: 앞 단계 결과가 든 블록에서만 다시 찾는다', () => {
    const first = [...searchPage(s1, page1, '펌프', 'all', null), ...searchPage(s2, page2, '펌프', 'all', null)];
    const allowed = blocksOf(first);
    // '합니다'는 다섯 블록 모두에 있지만 '펌프' 결과가 든 블록(1쪽 문단, 2쪽 첫 문단)만 남는다
    const narrowed = [...searchPage(s1, page1, '합니다', 'all', allowed), ...searchPage(s2, page2, '합니다', 'all', allowed)];
    expect(narrowed.map((x) => `${x.page}:${x.block}`)).toEqual(['1:1', '1:1', '2:0']);
  });
});
