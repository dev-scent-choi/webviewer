import { brotliDecompressSync } from 'node:zlib';
import { describe, expect, test } from 'vitest';
import { APP_SIZE_LIMITS, MiB, SIZE_LIMITS, appSizeProblems, brotli, checkLock, containsBytes, craftFontsManifest, hashedName, sizeProblems, treeHash, unservedAppAssets, wasmMaxMemory } from '../scripts/artcraftBuildLib.js';

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

  test('앱 화면 크기 상한(brotli wasm): word 8.5 · deck 5.5 · grid 5.5 MB — 시험 실측 7.9/4.9/4.5 MB + 여유(격자는 Task 6의 한글 글꼴 약 0.3 MB 포함)', () => {
    expect(APP_SIZE_LIMITS).toEqual({ word: 8.5 * MiB, deck: 5.5 * MiB, grid: 5.5 * MiB });
    expect(appSizeProblems({ word: 8.5 * MiB, deck: 5.5 * MiB, grid: 5.5 * MiB })).toEqual([]);
    const over = appSizeProblems({ word: 8.5 * MiB + 1, deck: 1, grid: 5.5 * MiB + 1 });
    expect(over).toHaveLength(2);
    expect(over[0]).toContain('word');
    expect(over[1]).toContain('grid');
  });

  test('폴더 해시: 같은 파일 묶음은 같은 값(순서 무관), 이름이나 내용이 바뀌면 바뀐다', () => {
    const a = [
      { rel: 'wordcraft-web.js', bytes: Buffer.from('js') },
      { rel: 'snippets/eframe-1/inline0.js', bytes: Buffer.from('s') },
    ];
    const h = treeHash(a);
    expect(h).toMatch(/^[0-9a-f]{10}$/);
    expect(treeHash([...a].reverse())).toBe(h);
    expect(treeHash([{ ...a[0], rel: 'other.js' }, a[1]])).not.toBe(h);
    expect(treeHash([{ ...a[0], bytes: Buffer.from('js2') }, a[1]])).not.toBe(h);
  });

  test('격자 앱 한글 글꼴 manifest: gridcraft 형식(8칸, " | "), NanumGothic Regular 자리, 문자 Hang,Kore, sha256 64자', () => {
    const sha = 'a'.repeat(64);
    const rows = craftFontsManifest('fonts/SCDream4.otf', sha).split('\n').filter((l) => l.trim() !== '' && !l.startsWith('#'));
    expect(rows).toHaveLength(1);
    const fields = rows[0].split(' | ');
    expect(fields).toHaveLength(8);
    expect(fields.slice(0, 4)).toEqual(['NanumGothic', 'Regular', 'fonts/SCDream4.otf', 'Hang,Kore']);
    expect(fields[6]).toBe(sha);
    expect(fields.every((f) => f.trim() !== '')).toBe(true);
  });

  test('containsBytes: 글꼴 조각이 wasm에 들어갔는지', () => {
    const hay = Buffer.from('앞부분-글꼴조각-뒷부분');
    expect(containsBytes(hay, Buffer.from('글꼴조각'))).toBe(true);
    expect(containsBytes(hay, Buffer.from('없는조각'))).toBe(false);
  });

  test('산출 파일 허용 목록: 서버 APP_ASSET_PATH에 안 맞는 파일만 돌려준다(build:artcraft-app이 쓰기 전에 멈춤)', () => {
    const dir = 'gridcraft-0123456789';
    expect(unservedAppAssets(dir, ['gridcraft-web.js', 'gridcraft-web_bg.wasm', 'snippets/egui-abc_1/inline0.js'])).toEqual([]);
    expect(unservedAppAssets(dir, ['gridcraft-web.js', 'snippets/egui/inline0.mjs', 'snippets/a/b/inline1.js', 'gridcraft-web.d.ts'])).toEqual([
      'snippets/egui/inline0.mjs',
      'snippets/a/b/inline1.js',
      'gridcraft-web.d.ts',
    ]);
  });
});
