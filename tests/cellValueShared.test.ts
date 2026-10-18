import { describe, expect, test } from 'vitest';
import {
  dateTextToSerial,
  isDateNumFmt,
  parseTypedNumber,
  serialToDateText,
} from '../src/collab/cellValueShared.js';

describe('cellValueShared: 엑셀 날짜 일련번호', () => {
  test('알려진 기준값과 일치한다 (2026-07-05 = 46208, 1900-01-01 = 2)', () => {
    expect(dateTextToSerial('2026-07-05')).toBe(46208);
    expect(dateTextToSerial('1900-01-01')).toBe(2);
  });
  test('시각은 소수부로 들어가고 왕복이 정확하다', () => {
    const serial = dateTextToSerial('2026-07-05 18:00')!;
    expect(serial).toBe(46208.75);
    expect(serialToDateText(serial)).toBe('2026-07-05 18:00');
    expect(serialToDateText(46215)).toBe('2026-07-12');
  });
  test('날짜 모양이 아니거나 존재하지 않는 날짜는 null (텍스트로 남는다)', () => {
    expect(dateTextToSerial('2026-02-31')).toBeNull();
    expect(dateTextToSerial('010-1234-5678')).toBeNull();
    expect(dateTextToSerial('2026/07/05')).toBeNull();
    expect(dateTextToSerial('본사')).toBeNull();
  });
});

describe('cellValueShared: isDateNumFmt', () => {
  test.each([
    ['yyyy-mm-dd', true],
    ['mm-dd-yy', true],
    ['yyyy"년" m"월" d"일"', true],
    ['[$-409]mmm d, yyyy', true],
    ['h:mm', false],
    ['mm:ss', false],
    ['[h]:mm:ss', false],
    ['#,##0', false],
    ['0.0%', false],
    ['"₩"#,##0', false],
    ['General', false],
    [undefined, false],
  ])('%s → %s', (fmt, expected) => {
    expect(isDateNumFmt(fmt)).toBe(expected);
  });
});

describe('cellValueShared: parseTypedNumber (엑셀처럼 입력값 해석)', () => {
  test('천 단위 쉼표', () => {
    expect(parseTypedNumber('1,000')).toEqual({ value: 1000, kind: 'number', decimals: 0 });
    expect(parseTypedNumber('-12,345.50')).toEqual({ value: -12345.5, kind: 'number', decimals: 2 });
  });
  test('백분율', () => {
    expect(parseTypedNumber('50%')).toEqual({ value: 0.5, kind: 'percent', decimals: 0 });
    expect(parseTypedNumber('12.5%')).toEqual({ value: 0.125, kind: 'percent', decimals: 1 });
  });
  test('원화', () => {
    expect(parseTypedNumber('₩1,000')).toEqual({ value: 1000, kind: 'currency', decimals: 0 });
    expect(parseTypedNumber('₩ 2500')).toEqual({ value: 2500, kind: 'currency', decimals: 0 });
  });
  test('엑셀도 텍스트로 두는 입력은 null', () => {
    expect(parseTypedNumber('1,00')).toBeNull(); // 천 단위 규칙 위반
    expect(parseTypedNumber('123')).toBeNull(); // 일반 숫자는 원래 경로가 처리
    expect(parseTypedNumber('A,B')).toBeNull();
    expect(parseTypedNumber('%')).toBeNull();
  });
});
