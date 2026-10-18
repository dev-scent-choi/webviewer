import { describe, expect, test } from 'vitest';
import { resolveSlideShapeLooks } from '../src/pptx/inheritance.js';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const tree = (inner: string) =>
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
  inner +
  '</p:spTree></p:cSld>';
const sp = (id: number, ph: string, spPr: string, txBody: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr>${ph}</p:nvPr></p:nvSpPr>` +
  `<p:spPr>${spPr}</p:spPr><p:txBody>${txBody}<a:p/></p:txBody></p:sp>`;
const xfrm = (x: number, y: number, w: number, h: number) =>
  `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`;
const rels = (type: string, target: string) =>
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>` +
  '</Relationships>';
const enc = (s: string) => new TextEncoder().encode(s);

function deck(): Record<string, Uint8Array> {
  const slide =
    `<p:sld ${NS}>` +
    tree(
      sp(2, '<p:ph type="ctrTitle"/>', '', '<a:bodyPr/>') + // 위치·크기 전부 상속
      sp(3, '<p:ph type="subTitle" idx="1"/>', '', '<a:bodyPr/>') +
      sp(4, '', xfrm(10, 20, 30, 40), '<a:bodyPr anchor="ctr"/><a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1400"/></a:lvl1pPr></a:lstStyle>'),
    ) +
    '</p:sld>';
  const layout =
    `<p:sldLayout ${NS}>` +
    tree(
      sp(2, '<p:ph type="ctrTitle"/>', xfrm(685800, 2130425, 7772400, 1470025), '<a:bodyPr anchor="b"/>') +
      sp(3, '<p:ph type="subTitle" idx="1"/>', xfrm(1371600, 3886200, 6400800, 1752600),
        '<a:bodyPr/><a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>'),
    ) +
    '</p:sldLayout>';
  const master =
    `<p:sldMaster ${NS}>` +
    tree(sp(2, '<p:ph type="title"/>', xfrm(1, 2, 3, 4), '<a:bodyPr anchor="ctr"/>')) +
    '<p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"/></a:lvl1pPr></p:titleStyle>' +
    '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:bodyStyle>' +
    '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles>' +
    '</p:sldMaster>';
  return {
    'ppt/slides/slide1.xml': enc(slide),
    'ppt/slides/_rels/slide1.xml.rels': enc(rels('slideLayout', '../slideLayouts/slideLayout1.xml')),
    'ppt/slideLayouts/slideLayout1.xml': enc(layout),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': enc(rels('slideMaster', '../slideMasters/slideMaster1.xml')),
    'ppt/slideMasters/slideMaster1.xml': enc(master),
  };
}

describe('resolveSlideShapeLooks: 레이아웃·마스터 상속', () => {
  const looks = resolveSlideShapeLooks(deck(), 'slide1.xml');

  test('가운데 제목 개체 틀: 위치·크기·세로 맞춤은 레이아웃, 글자 크기·정렬은 마스터 titleStyle', () => {
    expect(looks.get("2")).toMatchObject({
      x: 685800, y: 2130425, w: 7772400, h: 1470025, anchor: 'b', defaultSizePt: 44, defaultAlign: 'ctr',
    });
  });

  test('부제목(idx로 짝 찾기): 레이아웃 lstStyle의 크기·정렬이 마스터 bodyStyle보다 우선', () => {
    expect(looks.get('3')).toMatchObject({ x: 1371600, y: 3886200, defaultSizePt: 24, defaultAlign: 'ctr', anchor: 't' });
  });

  test('일반 글상자: 자기 값(세로 가운데·14pt·오른쪽)이 그대로', () => {
    expect(looks.get("4")).toMatchObject({ x: 10, y: 20, w: 30, h: 40, anchor: 'ctr', defaultSizePt: 14, defaultAlign: 'r' });
  });

  test('레이아웃·마스터가 없어도 던지지 않고 기본값(위쪽/18pt/왼쪽)', () => {
    const only = { 'ppt/slides/slide1.xml': deck()['ppt/slides/slide1.xml'] };
    expect(resolveSlideShapeLooks(only, "slide1.xml").get("2")).toMatchObject({
      x: null, y: null, w: null, h: null, anchor: 't', defaultSizePt: 18, defaultAlign: 'l',
    });
  });
});

describe('resolveSlideShapeLooks: 목록 수준(글머리·들여쓰기)', () => {
  test('마스터 bodyStyle의 수준별 글머리·여백을 개체 틀이 물려받고, 레이아웃 buNone이 이긴다', () => {
    const zip = deck();
    const master = new TextDecoder().decode(zip['ppt/slideMasters/slideMaster1.xml']).replace(
      '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="3200"/></a:lvl1pPr></p:bodyStyle>',
      '<p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900"><a:buChar char="•"/><a:defRPr sz="3200"/></a:lvl1pPr>' +
        '<a:lvl2pPr marL="742950" indent="-285750"><a:buChar char="–"/><a:defRPr sz="2800"/></a:lvl2pPr></p:bodyStyle>',
    );
    zip['ppt/slideMasters/slideMaster1.xml'] = new TextEncoder().encode(master);
    const looks = resolveSlideShapeLooks(zip, 'slide1.xml');
    // 부제목(idx=1)은 레이아웃 lstStyle이 lvl1을 다시 정의하지만 글머리는 언급 안 함 → 마스터 bodyStyle의 •
    const sub = looks.get('3')!;
    expect(sub.levels[0]).toEqual({ marLPt: 27, indentPt: -27, bullet: '•', sizePt: 24 });
    expect(sub.levels[1]).toEqual({ marLPt: 58.5, indentPt: -22.5, bullet: '–', sizePt: 28 });
  });
});
