import { describe, expect, test } from 'vitest';
import { CAPACITY_NOTICE, blockedViewKey, capacityFallback, isViewUi, viewControlAllowed } from '../src/client/viewStrip.js';

describe('isViewUi — 두 겹(토큰 읽기 전용 + ui=view)', () => {
  test.each([
    ['?token=t&ui=view', false, true],
    ['?ui=view&token=t', false, true],
    ['?token=t&ui=view', true, false],
    ['?token=t', false, false],
    ['?token=t&ui=edit', false, false],
    ['?token=t&ui=VIEW', false, false],
  ])('%s · canEdit %s → %s', (search, canEdit, expected) => {
    expect(isViewUi(search, canEdit)).toBe(expected);
  });
});

describe('viewControlAllowed — 보기 꾸밈에서 눌리는 리본 단추(나머지는 흐리게)', () => {
  test.each([
    ['btn-zoom-in-ribbon', true, true],
    ['btn-ribbon-find', true, true],
    ['btn-show-comments', true, true],
    ['btn-transition-preview', true, true],
    ['btn-copy', true, true],
    ['btn-copy', false, false],
    ['btn-bold', true, false],
    ['btn-ribbon-replace', true, false],
    ['btn-new-comment', true, false],
    ['btn-freeze-panes', true, false],
    ['btn-find-replace', true, false],
    ['qa-save', true, false],
    ['', true, false],
  ])('%s · 복사 %s → %s', (id, canCopy, expected) => {
    expect(viewControlAllowed(id, canCopy)).toBe(expected);
  });
});

describe('blockedViewKey — 편집 단축키는 무시, 찾기·이동은 그대로', () => {
  test.each([
    ['s', true, false, true, true],
    ['S', true, false, true, true],
    ['h', true, false, true, true],
    ['v', true, false, true, true],
    ['x', true, false, true, true],
    ['z', true, false, true, true],
    ['y', true, false, true, true],
    ['c', true, false, true, false],
    ['c', true, false, false, true],
    ['f', true, false, false, false],
    ['a', true, false, false, false],
    ['Delete', false, false, true, true],
    ['Backspace', false, false, true, true],
    ['F2', false, false, true, true],
    ['ArrowDown', false, false, true, false],
    ['PageDown', false, false, true, false],
    ['a', false, false, true, false],
    ['Backspace', false, true, true, false],
    ['v', true, true, true, false],
  ])('%s · ctrl %s · 입력칸 %s · 복사 %s → %s', (key, ctrl, inField, canCopy, expected) => {
    expect(blockedViewKey({ key, ctrl, inField, canCopy })).toBe(expected);
  });
});

describe('blockedViewKey — Ctrl+글자는 허용 목록(F·A·C(복사)·P(인쇄))과 확대 키만', () => {
  test.each<[string, string | undefined, boolean, boolean, boolean]>([
    // 허용 목록 밖의 글자는 모두 막는다(굵게·밑줄·기울임·메모·채우기 …)
    ...['b', 'd', 'i', 'u', 'm', 'k', 'e', 'g', 'j'].map((key): [string, undefined, boolean, boolean, boolean] => [key, undefined, true, true, true]),
    // 브라우저 단축키(새로 고침·주소창)는 보기 화면처럼 막지 않는다
    ['r', undefined, true, true, false],
    ['l', undefined, true, true, false],
    ['p', undefined, true, true, false],
    ['p', undefined, true, false, true],
    ['P', undefined, true, false, true],
    ['c', undefined, false, true, true],
    ['c', undefined, true, false, false],
    ['f', undefined, false, false, false],
    ['a', undefined, false, false, false],
    ['+', undefined, false, false, false],
    ['=', undefined, false, false, false],
    ['-', undefined, false, false, false],
    ['0', undefined, false, false, false],
    ['ArrowDown', undefined, false, false, false],
    ['Home', undefined, false, false, false],
    // 한글 입력 상태 — key는 한글이지만 code로 글자를 안다
    ['ㅠ', 'KeyB', true, true, true],
    ['ㄹ', 'KeyF', false, false, false],
    ['ㅔ', 'KeyP', true, true, false],
  ])('Ctrl+%s (code %s) · 복사 %s · 인쇄 %s → %s', (key, code, canCopy, canPrint, expected) => {
    expect(blockedViewKey({ key, code, ctrl: true, inField: false, canCopy, canPrint })).toBe(expected);
  });
  test('입력칸 안은 Ctrl+글자도 막지 않는다', () => {
    expect(blockedViewKey({ key: 'b', ctrl: true, inField: true, canCopy: false, canPrint: false })).toBe(false);
  });
});

describe('capacityFallback — 큰 엑셀 정원 초과(4429)', () => {
  const url = '/viewer.html?token=t';
  test('보기 꾸밈 + 4429 + 보기 주소 → 그 주소', () => expect(capacityFallback(4429, true, url)).toBe(url));
  test('편집 화면·다른 거부·주소 없음 → null(지금처럼 오류 화면)', () => {
    expect(capacityFallback(4429, false, url)).toBeNull();
    expect(capacityFallback(4403, true, url)).toBeNull();
    expect(capacityFallback(4429, true, undefined)).toBeNull();
  });
  test('안내 문구는 설계 그대로', () => expect(CAPACITY_NOTICE).toBe('지금은 보는 사람이 많습니다. 서버 변환으로 보여 줍니다'));
});
