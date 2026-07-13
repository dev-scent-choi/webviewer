import { describe, expect, test } from 'vitest';
import { SizedLruCache } from '../src/cache.js';

describe('SizedLruCache', () => {
  test('caches by key and reuses the in-flight promise', async () => {
    const cache = new SizedLruCache<number>({ maxBytes: 1000, sizeOf: () => 10 });
    let calls = 0;
    const loader = () => {
      calls += 1;
      return Promise.resolve(42);
    };
    const a = cache.getOrLoad('k', loader);
    const b = cache.getOrLoad('k', loader);
    expect(await a).toBe(42);
    expect(await b).toBe(42);
    expect(calls).toBe(1);
  });

  test('evicts the oldest entry when the byte budget is exceeded', async () => {
    const cache = new SizedLruCache<string>({ maxBytes: 100, sizeOf: (v) => v.length });
    await cache.getOrLoad('a', () => Promise.resolve('x'.repeat(60)));
    await cache.getOrLoad('b', () => Promise.resolve('y'.repeat(60)));
    // 총 120B > 100B 예산 → 가장 오래된 'a'가 제거된다
    expect(cache.get('a')).toBeUndefined();
    expect(await cache.get('b')).toBe('y'.repeat(60));
  });

  test('get() refreshes recency so the least-recently-used entry is evicted', async () => {
    const cache = new SizedLruCache<string>({ maxBytes: 100, sizeOf: (v) => v.length });
    await cache.getOrLoad('a', () => Promise.resolve('a'.repeat(45)));
    await cache.getOrLoad('b', () => Promise.resolve('b'.repeat(45)));
    // 'a'를 최근 사용으로 갱신 → 다음 삽입 시 'b'가 제거 대상이 된다
    await cache.get('a');
    await cache.getOrLoad('c', () => Promise.resolve('c'.repeat(45)));
    expect(await cache.get('a')).toBe('a'.repeat(45));
    expect(cache.get('b')).toBeUndefined();
  });

  test('enforces the entry-count limit independently of bytes', async () => {
    const cache = new SizedLruCache<number>({ maxBytes: 1e9, maxEntries: 2, sizeOf: () => 1 });
    await cache.getOrLoad('a', () => Promise.resolve(1));
    await cache.getOrLoad('b', () => Promise.resolve(2));
    await cache.getOrLoad('c', () => Promise.resolve(3));
    expect(cache.get('a')).toBeUndefined();
    expect(await cache.get('c')).toBe(3);
  });

  test('drops entries whose loader rejects', async () => {
    const cache = new SizedLruCache<number>({ maxBytes: 1000, sizeOf: () => 10 });
    await expect(cache.getOrLoad('x', () => Promise.reject(new Error('fail')))).rejects.toThrow('fail');
    expect(cache.get('x')).toBeUndefined();
  });
});
