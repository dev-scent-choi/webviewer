import { describe, expect, test } from 'vitest';
import { CAPACITY_NOTICE, blockedViewKey, capacityFallback, isViewUi, pickBoundary, stepIndex, stripActions } from '../src/client/viewStrip.js';

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

describe('stripActions — 얇은 줄 단추', () => {
  test('권한 다 있고 확대 있고 편집 열기 가능 → 정해진 순서로 전부', () => {
    expect(stripActions({ canPrint: true, canDownload: true, hasZoom: true, canOpenEdit: true })).toEqual(['zoom-out', 'zoom-in', 'print', 'download', 'edit']);
  });
  test('PPT(확대 없음)·인쇄 권한 없음·공유 링크(편집 열기 없음) → 내려받기만', () => {
    expect(stripActions({ canPrint: false, canDownload: true, hasZoom: false, canOpenEdit: false })).toEqual(['download']);
  });
  test('인쇄만', () => {
    expect(stripActions({ canPrint: true, canDownload: false, hasZoom: false, canOpenEdit: false })).toEqual(['print']);
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

describe('stepIndex — 시트·슬라이드 한 칸 앞/뒤(끝에서 멈춤)', () => {
  test.each([
    [3, 0, 1, 1],
    [3, 2, 1, 2],
    [3, 0, -1, 0],
    [3, 2, -1, 1],
    [3, -1, 1, 0],
    [0, -1, 1, -1],
  ] as const)('길이 %s · 지금 %s · 방향 %s → %s', (length, current, dir, expected) => {
    expect(stepIndex(length, current, dir)).toBe(expected);
  });
});

describe('pickBoundary — 워드 쪽 경계(도구 줄 아래 기준 px)', () => {
  test('다음: 기준보다 아래인 첫 경계, 이전: 기준보다 위인 마지막 경계', () => {
    const tops = [-500, -2, 0.5, 300];
    expect(pickBoundary(tops, 1)).toBe(3);
    expect(pickBoundary(tops, -1)).toBe(1);
  });
  test('없으면 -1', () => {
    expect(pickBoundary([], 1)).toBe(-1);
    expect(pickBoundary([100], -1)).toBe(-1);
    expect(pickBoundary([-100], 1)).toBe(-1);
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
