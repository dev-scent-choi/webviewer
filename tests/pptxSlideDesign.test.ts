import { describe, expect, test } from 'vitest';
import { applySlideDesign, buildTimingXml, readSlideDesign } from '../src/pptx/slideDesign.js';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const slide = (extra = '', bg = '') =>
  `<p:sld ${NS}><p:cSld>${bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
  `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${extra}</p:sld>`;

describe('PowerPoint 디자인·전환·애니메이션 저장', () => {
  test('배경 색: cSld 첫 자식으로 넣고, 다시 바꾸면 교체, null이면 지운다', () => {
    const once = applySlideDesign(slide(), { background: '#1f4e79' });
    expect(once).toContain('<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="1F4E79"/>');
    const twice = applySlideDesign(once, { background: '#ffffff' });
    expect(twice.match(/<p:bg>/g)).toHaveLength(1);
    expect(readSlideDesign(twice).background).toBe('#ffffff');
    expect(applySlideDesign(twice, { background: null })).not.toContain('<p:bg>');
  });

  test('전환: clrMapOvr 뒤에, 기존 mc:AlternateContent 전환은 지우고 하나만', () => {
    const withOld = slide('<mc:AlternateContent xmlns:mc="m"><mc:Choice Requires="p14"><p:transition spd="slow" p14:dur="2000"><p14:vortex/></p:transition></mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>');
    const out = applySlideDesign(withOld, { transition: { kind: 'push', speed: 'fast' } });
    expect(out).not.toContain('AlternateContent');
    expect(out).toContain('</p:clrMapOvr><p:transition spd="fast"><p:push dir="u"/></p:transition></p:sld>');
    expect(readSlideDesign(out).transition).toEqual({ kind: 'push', speed: 'fast' });
    expect(applySlideDesign(out, { transition: null })).not.toContain('p:transition');
  });

  test('애니메이션: 전환 뒤에 timing, 클릭/함께/다음 묶음, 다시 읽으면 같은 목록', () => {
    const animations = [
      { shapeId: '4', effect: 'fade' as const, trigger: 'click' as const },
      { shapeId: '5', effect: 'fly' as const, trigger: 'with' as const },
      { shapeId: '6', effect: 'wipe' as const, trigger: 'after' as const },
      { shapeId: '7', effect: 'appear' as const, trigger: 'click' as const },
    ];
    const out = applySlideDesign(slide('<p:transition><p:fade/></p:transition>'), { animations });
    expect(out.indexOf('<p:timing>')).toBeGreaterThan(out.indexOf('</p:transition>'));
    expect(readSlideDesign(out).animations).toEqual(animations);
    const timing = buildTimingXml(animations);
    // 클릭 그룹 2개(4·5·6 / 7), 6은 앞 효과(500ms) 뒤에 시작
    expect(timing.match(/<p:cond delay="indefinite"\/>/g)).toHaveLength(2);
    expect(timing).toContain('<p:cond delay="500"/>');
    expect(timing).toContain('<p:bldP spid="7" grpId="0" animBg="1"/>');
    // cTn id는 겹치지 않는다
    const ids = [...timing.matchAll(/<p:cTn id="(\d+)"/g)].map((m) => m[1]);
    expect(new Set(ids).size).toBe(ids.length);
    expect(applySlideDesign(out, { animations: [] })).not.toContain('<p:timing>');
  });

  test('바꾸지 않은 항목은 원본 그대로', () => {
    const original = slide('<p:transition spd="slow"><p:wipe dir="d"/></p:transition><p:timing><p:tnLst/></p:timing>', '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>');
    expect(applySlideDesign(original, {})).toBe(original);
  });
});
