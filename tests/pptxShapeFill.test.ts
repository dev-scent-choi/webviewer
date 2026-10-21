import { describe, expect, test } from 'vitest';
import { setShapeSolidFill } from '../src/pptx/slideContent.js';

const sp = (id: string, spPr: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>${spPr}<p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>`;

describe('기존 도형 채우기 바꾸기(setShapeSolidFill)', () => {
  test('기존 단색 채우기를 바꾸고 선(a:ln) 안의 색은 그대로', () => {
    const xml = sp('3', '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="10" cy="10"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></p:spPr>');
    const out = setShapeSolidFill(xml, '3', '#ff8800');
    expect(out).toContain('</a:prstGeom><a:solidFill><a:srgbClr val="FF8800"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>');
    expect(out).not.toContain('accent1');
  });

  test('채우기 없음·그라데이션도 단색으로, 다른 도형은 그대로', () => {
    const xml = sp('2', '<p:spPr><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="111111"/></a:gs></a:gsLst></a:gradFill></p:spPr>') +
      sp('5', '<p:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>');
    const out = setShapeSolidFill(setShapeSolidFill(xml, '5', '#00FF00'), '2', '#123456');
    expect(out).toContain('<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:spPr>');
    expect(out).toContain('<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></p:spPr>');
    expect(out).not.toContain('gradFill');
    expect(out).not.toContain('noFill');
  });

  test('빈 spPr(레이아웃 상속)에도 넣고, 형식이 틀린 색·없는 도형은 그대로', () => {
    const xml = sp('7', '<p:spPr/>');
    expect(setShapeSolidFill(xml, '7', '#ABCDEF')).toContain('<p:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></p:spPr>');
    expect(setShapeSolidFill(xml, '7', 'red')).toBe(xml);
    expect(setShapeSolidFill(xml, '99', '#ABCDEF')).toBe(xml);
  });
});
