import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'vitest';
import { fontStack } from '../src/docx/schema.js';

describe('나눔고딕 자체 호스팅', () => {
  test('글꼴 목록이 나눔고딕과 영문 이름을 함께 싣는다', () => {
    expect(fontStack(['나눔고딕'])).toMatch(/^"나눔고딕", "NanumGothic"/);
  });
  test('CSS가 세 이름 × 굵기 400·700 @font-face 를 갖고 파일이 있다', () => {
    const css = readFileSync('public/css/style.css', 'utf8');
    for (const fam of ['NanumGothic', '나눔고딕', 'Nanum Gothic']) {
      for (const w of [400, 700]) expect(css).toContain(`font-family: '${fam}'; font-weight: ${w}; src: url('/fonts/nanum/`);
    }
    for (const f of ['NanumGothic-Regular.woff', 'NanumGothic-Bold.woff', 'OFL.txt']) expect(readFileSync(`public/fonts/nanum/${f}`).length).toBeGreaterThan(1000);
  });
});
