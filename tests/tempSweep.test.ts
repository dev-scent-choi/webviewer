import { mkdtempSync, mkdirSync, utimesSync, writeFileSync, existsSync, rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterEach, describe, expect, test } from 'vitest';
import { STALE_AFTER_MS, isSweepableTempName, sweepStaleTempDirs } from '../src/tempSweep.js';

describe('변환 임시 폴더 청소', () => {
  const roots: string[] = [];
  afterEach(() => roots.splice(0).forEach((r) => rmSync(r, { recursive: true, force: true })));

  test('이 앱의 변환 폴더·내보내기 파일만 대상 — LibreOffice 프로필·다른 프로그램 폴더는 아님', () => {
    expect(isSweepableTempName('docview-lo-QuupxQ')).toBe(true);
    expect(isSweepableTempName('docview-convert-5Htzcl')).toBe(true);
    expect(isSweepableTempName('docview-warmup-BoZfW5')).toBe(true);
    expect(isSweepableTempName('docview-xlsx-stream-123e4567-e89b-12d3-a456-426614174000.xlsx')).toBe(true);
    expect(isSweepableTempName('docview-lo-profiles')).toBe(false);
    expect(isSweepableTempName('docview-lazygrid-abc.jsonl')).toBe(false);
    expect(isSweepableTempName('other-lo-QuupxQ')).toBe(false);
  });

  test('오래된 사본은 지우고, 방금 만든(변환 중일 수 있는) 폴더와 프로필은 남긴다', async () => {
    const root = mkdtempSync(path.join(os.tmpdir(), 'sweep-test-'));
    roots.push(root);
    const old = path.join(root, 'docview-lo-AAAAAA');
    const fresh = path.join(root, 'docview-lo-BBBBBB');
    const profiles = path.join(root, 'docview-lo-profiles');
    for (const d of [old, fresh, profiles]) {
      mkdirSync(d);
      writeFileSync(path.join(d, 'source.pptx'), 'x');
    }
    const now = Date.now();
    const past = new Date(now - STALE_AFTER_MS - 60_000);
    utimesSync(old, past, past);
    utimesSync(profiles, past, past);
    expect(await sweepStaleTempDirs(root, now)).toBe(1);
    expect(existsSync(old)).toBe(false);
    expect(existsSync(fresh)).toBe(true);
    expect(existsSync(profiles)).toBe(true);
  });
});
