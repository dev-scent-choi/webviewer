import { mkdtempSync, readFileSync, rmSync, writeFileSync, mkdirSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterAll, describe, expect, test } from 'vitest';
import { ensureFontSubstitutions, FONT_SUBSTITUTIONS } from '../src/converters/loFontSubstitution.js';

describe('LibreOffice 글꼴 대체표', () => {
  const root = mkdtempSync(path.join(os.tmpdir(), 'fontsub-test-'));
  afterAll(() => rmSync(root, { recursive: true, force: true }));
  const fileOf = (profile: string) => path.join(profile, 'user', 'registrymodifications.xcu');
  const count = (xml: string) => (xml.match(/oor:name="docview_font_\d+"/g) ?? []).length;

  test('아직 없는 프로필에는 새 파일로, KoPub돋움체 → 맑은 고딕', () => {
    const profile = path.join(root, 'fresh');
    ensureFontSubstitutions(profile);
    const xml = readFileSync(fileOf(profile), 'utf8');
    expect(xml).toContain('<value>KoPub돋움체 Medium</value>');
    expect(xml).toContain('<value>Malgun Gothic</value>');
    expect(xml).toContain('<prop oor:name="Replacement" oor:op="fuse"><value>true</value></prop>');
    expect(count(xml)).toBe(FONT_SUBSTITUTIONS.length);
  });

  test('이미 쓰던 프로필에는 기존 설정을 지키며 덧붙이고, 두 번 불러도 겹치지 않는다', () => {
    const profile = path.join(root, 'used');
    mkdirSync(path.join(profile, 'user'), { recursive: true });
    writeFileSync(
      fileOf(profile),
      '<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry">\n<item oor:path="/org.openoffice.Office.Common/Misc"><prop oor:name="X" oor:op="fuse"><value>1</value></prop></item>\n</oor:items>\n',
    );
    ensureFontSubstitutions(profile);
    ensureFontSubstitutions(profile);
    const xml = readFileSync(fileOf(profile), 'utf8');
    expect(xml).toContain('oor:path="/org.openoffice.Office.Common/Misc"');
    expect(count(xml)).toBe(FONT_SUBSTITUTIONS.length);
    expect(xml.trim().endsWith('</oor:items>')).toBe(true);
  });

  test('외부 링크 차단(BlockUntrustedRefererLinks)도 넣고, 그것만 없던 예전 프로필에도 한 번만 덧붙인다', () => {
    const blocks = (xml: string) => (xml.match(/BlockUntrustedRefererLinks/g) ?? []).length;
    const fresh = path.join(root, 'block-fresh');
    ensureFontSubstitutions(fresh);
    expect(readFileSync(fileOf(fresh), 'utf8')).toMatch(/BlockUntrustedRefererLinks" oor:op="fuse"><value>true</);
    // 글꼴 대체표만 있던 예전 판 프로필
    const old = path.join(root, 'block-old');
    ensureFontSubstitutions(old);
    writeFileSync(fileOf(old), readFileSync(fileOf(old), 'utf8').replace(/<item[^\n]*BlockUntrustedRefererLinks[^\n]*\n/, ''));
    expect(blocks(readFileSync(fileOf(old), 'utf8'))).toBe(0);
    ensureFontSubstitutions(old);
    ensureFontSubstitutions(old);
    const xml = readFileSync(fileOf(old), 'utf8');
    expect(blocks(xml)).toBe(1);
    expect(count(xml)).toBe(FONT_SUBSTITUTIONS.length);
  });
});
