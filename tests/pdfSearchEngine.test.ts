import { describe, expect, test } from 'vitest';
import { analyzePage, buildProfile, isBoldFontName, type StructItem } from '../src/client/pdf/textStructure.js';
import { groupByPage, searchPage, sectionsOf, sectionText } from '../src/client/pdf/searchEngine.js';

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

  test('본문만: 제목 블록의 일치는 빼고 본문만', () => {
    const r = searchPage(s1, page1, '배관', 'body', null);
    expect(r.map((x) => [x.isHeading, x.heading])).toEqual([[false, '1. 배관 관리']]);
    // 전체 = 제목 + 본문
    expect(searchPage(s1, page1, '배관', 'all', null)).toHaveLength(searchPage(s1, page1, '배관', 'headings', null).length + r.length);
  });

  test('결과 안에서: 앞 단계 결과가 속한 절(제목~다음 제목, 쪽을 넘어서도)에서만 다시 찾는다', () => {
    const s2c = analyzePage(2, page2, s1.lastHeading, new Set(), {}, s1.lastSection);
    // '누수'는 1쪽 '1. 배관 관리' 절에만 있다 — 그 절은 2쪽 글까지 이어진다
    const first = searchPage(s1, page1, '누수', 'all', null);
    const allowed = sectionsOf(first);
    expect([...allowed]).toEqual(['1:2']);
    const narrowed = [...searchPage(s1, page1, '합니다', 'all', allowed), ...searchPage(s2c, page2, '합니다', 'all', allowed)];
    // 1쪽 첫 절(설비 운전 지침)의 '합니다' 두 건은 빠지고, 배관 절의 1쪽 1건 + 이어진 2쪽 2건만
    expect(narrowed.map((x) => `${x.page}:${x.block}`)).toEqual(['1:3', '2:0', '2:1']);
  });

  test('절 글: 제목부터 다음 절 전까지 쪽을 넘어 잇는다', () => {
    const s2c = analyzePage(2, page2, s1.lastHeading, new Set(), {}, s1.lastSection);
    const map = new Map([[1, s1], [2, s2c]]);
    const sec = sectionText('1:2', map);
    expect(sec.heading).toBe('1. 배관 관리');
    expect(sec.pages).toEqual([1, 2]);
    expect(sec.text.split('\n\n')).toEqual(['1. 배관 관리', '배관 누수는 즉시 보고합니다.', '예비 펌프로 바꿔 가동합니다.', '배관 압력을 기록합니다.']);
    expect(sectionText('1:0', map).text.split('\n\n')).toHaveLength(2);
  });
});

describe('제목 점수(붙여 받은 기준표)', () => {
  const at = (str: string, y: number, extra: Partial<StructItem> = {}, size = 10, x = 50): StructItem => ({ ...line(str, y, size, x), ...extra });
  const body = (y: number) => at('운전원은 매일 아침 점검표를 작성하고 이상이 있으면 보고합니다.', y, { fontName: 'R' });

  test('본문 크기라도 굵은 짧은 줄은 제목, 같은 글꼴의 본문은 아님', () => {
    const s = analyzePage(1, [at('점검 순서', 700, { fontName: 'B', bold: true }), body(680), body(666), body(652)]);
    expect(s.blocks[0].isHeading).toBe(true);
    expect(s.blocks[1].isHeading).toBe(false);
    // 알림 상자 표지('주의 사항', '참고!')는 굵어도 제목이 아니다
    expect(analyzePage(1, [at('주의 사항', 700, { fontName: 'B', bold: true }), body(680), body(666)]).blocks[0].isHeading).toBe(false);
  });

  test('글자가 커도 도면 치수(숫자뿐)·긴 문장은 제목이 아니다', () => {
    const s = analyzePage(1, [
      at('300', 700, {}, 18),
      body(660),
      at('이 도면은 1층 기계실 냉각수 배관의 배치와 밸브 위치를 보여 주며 시공 전에 현장 치수를 다시 확인해야 합니다.', 620, {}, 14),
      body(580),
    ]);
    expect(s.blocks.filter((b) => b.isHeading)).toHaveLength(0);
  });

  test('태그 PDF: H 역할은 크기와 상관없이 제목, P 역할의 번호 줄은 본문', () => {
    const s = analyzePage(1, [at('설치 조건', 700, { role: 'H2' }), at('1. 전원은 220V 입니다', 680, { role: 'P' }), body(660)]);
    expect(s.blocks[0].isHeading).toBe(true);
    expect(s.blocks.find((b) => b.text.startsWith('1. 전원'))?.isHeading).toBe(false);
  });

  test('큰 제목(1)과 소제목(2) 단계', () => {
    const s = analyzePage(1, page1);
    expect(s.blocks.filter((b) => b.isHeading).map((b) => b.level)).toEqual([1, 2]);
    expect(s.blocks.filter((b) => !b.isHeading).every((b) => b.level === 0)).toBe(true);
  });

  test('여백으로 떨어진 짧은 줄은 점수가 오르고, 칸이 크게 벌어진 표 줄은 내려간다', () => {
    const apart = analyzePage(1, [body(720), body(706), at('운전 기록', 670, { fontName: 'R' }), body(650), body(636)]);
    const tight = analyzePage(1, [body(720), body(706), at('운전 기록', 692, { fontName: 'R' }), body(678), body(664)]);
    expect(apart.blocks.find((b) => b.text === '운전 기록')!.isHeading).toBe(true);
    // 줄 간격이 보통이면 본문 흐름의 한 줄 — 따로 제목이 되지 않는다
    expect(tight.blocks.some((b) => b.isHeading)).toBe(false);
    const row = analyzePage(1, [at('1.', 700), at('펌프', 700, {}, 10, 150), at('정상', 700, {}, 10, 300), body(680)]);
    expect(row.blocks[0].text).toBe('1. 펌프 정상');
    expect(row.blocks[0].score).toBeLessThan(analyzePage(1, [at('1. 펌프 정상', 700), body(680)]).blocks[0].score);
  });

  test('글꼴 이름으로 굵기 짐작', () => {
    expect(isBoldFontName('ABCDEF+NanumGothicBold')).toBe(true);
    expect(isBoldFontName('HYHeadLine-Medium')).toBe(true);
    expect(isBoldFontName('MalgunGothic')).toBe(false);
  });
});

describe('본문만 찾기가 실제 문서에서 걸렸던 경우', () => {
  const at = (str: string, y: number, extra: Partial<StructItem> = {}, size = 10, x = 50): StructItem => ({ ...line(str, y, size, x), ...extra });

  test('목차(점선 뒤 쪽 번호) 쪽은 제목·본문 어디에도 들지 않고 전체에서만 찾는다', () => {
    const items = [
      at('목 차', 780, {}, 16),
      at('제1장 총칙', 740, { bold: true }),
      at('제1조 (목적) ·································· 3', 720),
      at('제2조 (정의) ·································· 4', 705),
      at('제3조 (적용 범위) ······························ 6', 690),
    ];
    const s = analyzePage(1, items);
    const chapter = s.blocks.find((b) => b.text === '제1장 총칙')!;
    expect(chapter.toc).toBe(true);
    expect(chapter.isHeading).toBe(false);
    expect(searchPage(s, items, '목적', 'body', null)).toHaveLength(0);
    expect(searchPage(s, items, '목적', 'headings', null)).toHaveLength(0);
    const all = searchPage(s, items, '목적', 'all', null);
    expect(all.map((r) => r.heading)).toEqual(['목차']);
    // 목차 줄은 다른 절에 끼지 않는다 — '결과 안에서'가 엉뚱한 절을 끌어오지 않게
    expect(all[0].section.startsWith('b:')).toBe(true);
  });

  test("본문 첫머리의 조 제목('제3조(목적) 이 규정은…'): 조 제목의 일치는 제목, 뒤 문장의 일치는 본문", () => {
    const items = [
      at('제3조(목적) 이 규정은 정보화 업무에 관한 사항을 정함을 목적으로 한다.', 700),
      at('제4조(정의) 이 규정에서 쓰는 말의 뜻은 다음과 같다.', 680),
    ];
    const s = analyzePage(1, items);
    expect(s.blocks.every((b) => !b.isHeading)).toBe(true);
    expect(s.blocks.map((b) => b.heading)).toEqual(['제3조(목적)', '제4조(정의)']);
    expect(searchPage(s, items, '목적', 'headings', null).map((r) => r.before + r.hit)).toEqual(['제3조(목적']);
    const body = searchPage(s, items, '목적', 'body', null);
    expect(body).toHaveLength(1);
    expect(body[0].after.startsWith('으로 한다')).toBe(true);
    // 두 조는 서로 다른 절
    expect(new Set(s.blocks.map((b) => b.section)).size).toBe(2);
  });

  test('머리말·쪽 번호는 어느 찾을 곳에서도 결과가 아니다', () => {
    const pageOf = (n: number) => [at('설비 운전 지침서', 830), at(`본문 ${n}쪽: 펌프 점검을 합니다.`, 500), at(String(n), 20)];
    const pages = [1, 2, 3, 4].map((n) => ({ items: pageOf(n), pageHeight: 842 }));
    const profile = buildProfile(pages);
    const s = analyzePage(2, pages[1].items, '', new Set(), { pageHeight: 842, profile });
    expect(s.blocks.filter((b) => b.furniture).map((b) => b.text)).toEqual(['설비 운전 지침서', '2']);
    for (const target of ['all', 'headings', 'body'] as const) expect(searchPage(s, pages[1].items, '지침', target, null)).toHaveLength(0);
    expect(searchPage(s, pages[1].items, '펌프', 'body', null)).toHaveLength(1);
  });

  test("조문 안의 '1.', '가.' 항목은 번호가 붙어도 본문", () => {
    const items = [
      at('제2조 (정의)', 760, { bold: true }, 11),
      at('이 영에서 사용하는 용어의 뜻은 다음과 같다.', 740),
      at('1. “직무관련자”란 공무원의 소관 업무와 관련되는 자를 말한다.', 725),
      at('가. 민원을 신청하는 중인 개인', 710),
      at('나. 계약을 체결하려는 법인', 695),
      at('2. “금품등”이란 다음 각 목의 것을 말한다.', 680),
    ];
    const s = analyzePage(1, items);
    expect(s.blocks.filter((b) => b.isHeading).map((b) => b.text)).toEqual(['제2조 (정의)']);
    expect(searchPage(s, items, '민원', 'body', null).map((r) => r.heading)).toEqual(['제2조 (정의)']);
  });
});
