import { randomUUID } from 'node:crypto';
import { describe, expect, test } from 'vitest';
import { acceptChunk } from '../src/chunkedUpload.js';

describe('조각 업로드(chunkedUpload.ts)', () => {
  test('조각이 순서와 상관없이 모이면 원래 바이트로 합친다', async () => {
    const id = randomUUID();
    const whole = Buffer.from(Array.from({ length: 300 }, (_, i) => i % 256));
    const parts = [whole.subarray(0, 100), whole.subarray(100, 200), whole.subarray(200)];
    expect(await acceptChunk(id, 2, 3, parts[2], 10_000)).toEqual({ done: false, received: 1 });
    expect(await acceptChunk(id, 0, 3, parts[0], 10_000)).toEqual({ done: false, received: 2 });
    const last = await acceptChunk(id, 1, 3, parts[1], 10_000);
    expect(last.done && last.bytes.equals(whole)).toBe(true);
  });

  test('합친 크기가 상한을 넘으면 413, 잘못된 값은 400', async () => {
    const id = randomUUID();
    await acceptChunk(id, 0, 2, Buffer.alloc(60), 100);
    await expect(acceptChunk(id, 1, 2, Buffer.alloc(60), 100)).rejects.toMatchObject({ status: 413 });
    await expect(acceptChunk('../../etc', 0, 1, Buffer.alloc(1), 100)).rejects.toMatchObject({ status: 400 });
    await expect(acceptChunk(randomUUID(), 3, 2, Buffer.alloc(1), 100)).rejects.toMatchObject({ status: 400 });
    await expect(acceptChunk(randomUUID(), 0, 999, Buffer.alloc(1), 100)).rejects.toMatchObject({ status: 400 });
  });
});
