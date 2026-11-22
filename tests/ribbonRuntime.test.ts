import { describe, expect, it } from 'vitest';
import { collapseOrder } from '../src/client/ribbonRuntime.js';

describe('collapseOrder — 오피스처럼 접는 순서', () => {
  it('collapse 작은 수 먼저, 같으면 오른쪽 묶음 먼저', () => {
    const groups = [
      { priority: 9, index: 0 },
      { priority: 2, index: 1 },
      { priority: 5, index: 2 },
      { priority: 2, index: 3 },
      { priority: 1, index: 4 },
    ];
    expect(collapseOrder(groups)).toEqual([4, 3, 1, 2, 0]);
  });
  it('빈 목록', () => expect(collapseOrder([])).toEqual([]));
});
