import { describe, expect, test } from 'vitest';
import { singleFlightWithBackoff } from '../src/xlsx/singleFlight.js';

describe('검토 N4 — 보기 그리드 재생성: 문서마다 하나, 실패하면 쉰다', () => {
  test('동시에 온 요청은 진행 중인 재생성 하나를 함께 기다린다', async () => {
    const sf = singleFlightWithBackoff(60_000);
    let runs = 0;
    let release!: () => void;
    const fn = () => {
      runs += 1;
      return new Promise<boolean>((r) => (release = () => r(true)));
    };
    const all = Promise.all(Array.from({ length: 20 }, () => sf.run('doc', fn)));
    release();
    expect(await all).toEqual(Array(20).fill(true));
    expect(runs).toBe(1);
  });
  test('실패하면 쉬는 동안 다시 하지 않고, 시간이 지나면 다시 한다', async () => {
    let t = 0;
    const sf = singleFlightWithBackoff(10_000, () => t);
    let runs = 0;
    const failing = async () => {
      runs += 1;
      throw new Error('원본 없음');
    };
    expect(await sf.run('doc', failing)).toBe(false);
    for (let i = 0; i < 50; i += 1) expect(await sf.run('doc', failing)).toBe(false);
    expect(runs).toBe(1);
    expect(sf.backingOff('doc')).toBe(true);
    t = 10_001;
    expect(await sf.run('doc', async () => true)).toBe(true);
    expect(sf.backingOff('doc')).toBe(false);
  });
});
