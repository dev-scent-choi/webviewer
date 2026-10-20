import { describe, expect, test } from 'vitest';
import { findInPage, normalizeForFind } from '../src/client/pdf/findText.js';

describe('PDF 찾기', () => {
  test('대소문자·전각·연속 공백을 무시하고, 글자 조각 경계를 넘는 일치도 찾는다', () => {
    const items = [{ str: '분기 보' }, { str: '고서 ' }, { str: 'Report  2026' }];
    const m = findInPage(3, items, '보고서');
    expect(m).toEqual([{ page: 3, itemStart: 0, offsetStart: 3, itemEnd: 1, offsetEnd: 2 }]);
    expect(findInPage(3, items, 'report 2026')).toHaveLength(1);
    expect(normalizeForFind('ＡＢ  c')).toBe('ab c');
  });
  test('빈 검색어는 결과 없음, 여러 번 나오면 모두', () => {
    expect(findInPage(1, [{ str: 'aa aa' }], '')).toEqual([]);
    expect(findInPage(1, [{ str: 'aa aa' }], 'aa')).toHaveLength(2);
  });
  test('조각 사이에 공백이 없어도(줄 바꿈으로 나뉜 단어) 이어서 찾는다', () => {
    expect(findInPage(1, [{ str: '가동' }, { str: '률' }], '가동률')).toEqual([{ page: 1, itemStart: 0, offsetStart: 0, itemEnd: 1, offsetEnd: 1 }]);
  });
});
