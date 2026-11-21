import { brotliDecompressSync } from 'node:zlib';
import { describe, expect, test } from 'vitest';
import { MiB, SIZE_LIMITS, brotli, checkLock, hashedName, sizeProblems, wasmMaxMemory } from '../scripts/artcraftBuildLib.js';

describe('build:artcraft 도우미', () => {
  test('해시 이름: 같은 내용은 같은 이름, 내용이 바뀌면 이름도 바뀐다(10자리 16진수)', () => {
    const a = hashedName('artcraft_word', '.wasm', Buffer.from('하나'));
    expect(a).toMatch(/^artcraft_word-[0-9a-f]{10}\.wasm$/);
    expect(hashedName('artcraft_word', '.wasm', Buffer.from('하나'))).toBe(a);
    expect(hashedName('artcraft_word', '.wasm', Buffer.from('둘'))).not.toBe(a);
  });

  test('크기 상한: word 5 MB, deck 2.5 MB, 글꼴 합 0.6 MB(brotli)', () => {
    expect(SIZE_LIMITS).toEqual({ word: 5 * MiB, deck: 2.5 * MiB, fonts: 0.6 * MiB });
    expect(sizeProblems({ word: 5 * MiB, deck: 2.5 * MiB, fonts: [0.3 * MiB, 0.3 * MiB] })).toEqual([]);
    const over = sizeProblems({ word: 5 * MiB + 1, deck: 2.5 * MiB + 1, fonts: [0.3 * MiB, 0.3 * MiB + 1] });
    expect(over).toHaveLength(3);
    expect(over[0]).toContain('word');
    expect(over[1]).toContain('deck');
    expect(over[2]).toContain('글꼴');
  });

  test('잠금 검사: HEAD가 다르거나 고친 파일이 있으면 이유, 같고 깨끗하면 null', () => {
    const want = 'f8ec3cc0fa884369c4f3ebbd9ca3085398f5ec31';
    expect(checkLock('wordcraft', want, want, '')).toBeNull();
    expect(checkLock('wordcraft', want, 'b'.repeat(40), '')).toContain('ARTCRAFT.lock');
    expect(checkLock('wordcraft', want, want, ' M crates/doc/src/lib.rs\n')).toContain('고친 파일');
  });

  test('brotli: 풀면 원래 바이트', () => {
    const raw = Buffer.from('아트크래프트 '.repeat(500));
    const br = brotli(raw);
    expect(br.length).toBeLessThan(raw.length);
    expect(Buffer.compare(brotliDecompressSync(br), raw)).toBe(0);
  });

  test('wasm 최대 메모리: 메모리 절의 max(바이트), 상한이 없거나 메모리 절이 없으면 null', () => {
    const head = [0x00, 0x61, 0x73, 0x6d, 0x01, 0x00, 0x00, 0x00];
    // 형식 절(빈 함수 형식 1개) 뒤 메모리 절: min 1쪽, max 8192쪽(LEB128 0x80 0x40) = 512 MiB
    const capped = Uint8Array.from([...head, 0x01, 0x04, 0x01, 0x60, 0x00, 0x00, 0x05, 0x05, 0x01, 0x01, 0x01, 0x80, 0x40]);
    expect(WebAssembly.validate(capped)).toBe(true);
    expect(wasmMaxMemory(capped)).toBe(512 * MiB);
    const uncapped = Uint8Array.from([...head, 0x05, 0x03, 0x01, 0x00, 0x01]);
    expect(WebAssembly.validate(uncapped)).toBe(true);
    expect(wasmMaxMemory(uncapped)).toBeNull();
    expect(wasmMaxMemory(Uint8Array.from(head))).toBeNull();
  });
});
