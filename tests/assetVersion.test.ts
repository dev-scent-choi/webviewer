import { describe, expect, test } from 'vitest';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { versionAssetUrls } from '../src/assetVersion.js';

describe('화면 HTML의 JS·CSS 주소에 내용 해시', () => {
  test('있는 파일은 ?v=해시, 내용이 바뀌면 해시도 바뀐다, 없는 파일·외부 주소는 그대로', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'asset-version-'));
    fs.mkdirSync(path.join(dir, 'js'));
    fs.writeFileSync(path.join(dir, 'js', 'a.js'), 'one');
    const html = '<script src="/js/a.js"></script><script src="/js/missing.js"></script><link href="https://x/y.css">';
    const first = versionAssetUrls(html, dir);
    expect(first).toMatch(/src="\/js\/a\.js\?v=[0-9a-f]{10}"/);
    expect(first).toContain('src="/js/missing.js"');
    expect(first).toContain('href="https://x/y.css"');
    // 수정 시각이 바뀌어야 다시 계산한다
    fs.writeFileSync(path.join(dir, 'js', 'a.js'), 'two');
    const later = new Date(Date.now() + 5000);
    fs.utimesSync(path.join(dir, 'js', 'a.js'), later, later);
    const second = versionAssetUrls(html, dir);
    expect(second).not.toBe(first);
    expect(versionAssetUrls(html, dir)).toBe(second);
  });
});
