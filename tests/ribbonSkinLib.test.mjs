import { describe, expect, test } from 'vitest';
import { SKIN_PX, SKIN_TOLERANCE_PX, bigButtonLayoutOk, heightProblems, latinOnlyLabels } from '../scripts/ribbonSkinLib.mjs';

describe('heightProblems — 설계 §4 높이 규격 ±2px', () => {
  test('규격 값과 경계(±2)는 통과', () => {
    expect(heightProblems({ title: 24, tabs: 22, ribbon: 62, status: 18 })).toEqual([]);
    expect(heightProblems({ title: 26, tabs: 20, ribbon: 60, status: 20 })).toEqual([]);
  });
  test('3px 넘게 벗어나면 그 칸 이름과 값을 알린다', () => {
    expect(heightProblems({ title: 27, tabs: 22, ribbon: 73, status: 18 })).toEqual(['title 27px (규격 24±2)', 'ribbon 73px (규격 62±2)']);
  });
  test('없는 칸(null·undefined — 보기 화면의 상태 줄)은 건너뛴다', () => {
    expect(heightProblems({ title: 24, tabs: 22, ribbon: 62, status: null })).toEqual([]);
    expect(heightProblems({ ribbon: 62 })).toEqual([]);
  });
  test('규격 숫자는 설계 값(캡처 실측) 그대로', () => {
    expect(SKIN_PX).toEqual({ title: 24, tabs: 22, ribbon: 62, status: 18 });
    expect(SKIN_TOLERANCE_PX).toBe(2);
  });
});

describe('latinOnlyLabels — 한글이 없고 영문자만 있는 라벨', () => {
  test.each([
    [['홈', '붙여넣기', '글꼴'], []],
    [['Home', '홈'], ['Home']],
    [['  Paste ', '클립보드'], ['Paste']],
    [['100%', '1.15', '', '▲'], []],
    [['PDF로 저장', 'A4 용지'], []],
  ])('%j → %j', (labels, expected) => {
    expect(latinOnlyLabels(labels)).toEqual(expected);
  });
});

describe('bigButtonLayoutOk — 큰 단추는 아이콘 위·이름 아래, 가로 가운데', () => {
  const icon = { left: 10, top: 2, width: 24, height: 24 };
  test('아이콘 아래에 이름, 가운데 맞음', () => {
    expect(bigButtonLayoutOk(icon, { left: 2, top: 28, width: 40, height: 13 })).toBe(true);
  });
  test('이름이 아이콘 옆(같은 높이)이면 실패', () => {
    expect(bigButtonLayoutOk(icon, { left: 36, top: 8, width: 40, height: 13 })).toBe(false);
  });
  test('가운데에서 3px 넘게 비켜나면 실패', () => {
    expect(bigButtonLayoutOk(icon, { left: 6, top: 28, width: 40, height: 13 })).toBe(false);
  });
  test('1px 겹침은 반올림 오차로 본다', () => {
    expect(bigButtonLayoutOk(icon, { left: 2, top: 25, width: 40, height: 13 })).toBe(true);
  });
});
