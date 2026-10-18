import { describe, expect, test } from 'vitest';
import { loadThemeColors } from '../src/pptx/theme.js';
import { parseSlideShapes, parseSlideVisuals, serializeTxBodyXml } from '../src/pptx/slideContent.js';

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const enc = (s: string) => new TextEncoder().encode(s);
const tree = (inner: string, bg = '') =>
  `<p:cSld>${bg}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${inner}</p:spTree></p:cSld>`;
const xfrm = (x: number, y: number, w: number, h: number) =>
  `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`;
const rels = (entries: Array<[string, string, string]>) =>
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  entries
    .map(([id, type, target]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`)
    .join('') +
  '</Relationships>';

const THEME =
  `<a:theme ${NS} name="T"><a:themeElements><a:clrScheme name="c">` +
  '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>' +
  '</a:clrScheme></a:themeElements></a:theme>';

/** 슬라이드 1장 + 레이아웃 + 마스터(로고 그림·색 띠·개체 틀·배경) + 테마 */
function deck(opts: { slideShowMasterSp?: string; slideInner?: string } = {}): Record<string, Uint8Array> {
  const showAttr = opts.slideShowMasterSp ? ` showMasterSp="${opts.slideShowMasterSp}"` : '';
  const slide = `<p:sld ${NS}${showAttr}>${tree(opts.slideInner ?? '')}</p:sld>`;
  const layout =
    `<p:sldLayout ${NS}>` +
    tree(
      // 레이아웃 개체 틀(보이면 안 됨) + 레이아웃 고유 장식 도형(스타일 fillRef → accent2)
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="t"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
        `<p:spPr>${xfrm(1, 1, 1, 1)}<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="3" name="deco"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(10, 20, 30, 40)}<a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom></p:spPr>` +
        '<p:style><a:lnRef idx="2"><a:schemeClr val="accent2"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef>' +
        '<a:effectRef idx="0"><a:schemeClr val="accent2"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>',
    ) +
    '</p:sldLayout>';
  const master =
    `<p:sldMaster ${NS}>` +
    tree(
      '<p:pic><p:nvPicPr><p:cNvPr id="4" name="logo"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId9"/></p:blipFill>' +
        `<p:spPr>${xfrm(100, 200, 300, 400)}</p:spPr></p:pic>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="5" name="bar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(0, 0, 9144000, 50)}` +
        '<a:solidFill><a:schemeClr val="tx2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></p:spPr></p:sp>',
      '<p:bg><p:bgRef idx="1001"><a:schemeClr val="accent1"/></p:bgRef></p:bg>',
    ) +
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
    '</p:sldMaster>';
  return {
    'ppt/presentation.xml': enc(`<p:presentation ${NS}><p:sldSz cx="9144000" cy="6858000"/></p:presentation>`),
    'ppt/slides/slide1.xml': enc(slide),
    'ppt/slides/_rels/slide1.xml.rels': enc(rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml']])),
    'ppt/slideLayouts/slideLayout1.xml': enc(layout),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': enc(rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']])),
    'ppt/slideMasters/slideMaster1.xml': enc(master),
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': enc(
      rels([
        ['rId1', 'theme', '../theme/theme1.xml'],
        ['rId9', 'image', '../media/logo.png'],
      ]),
    ),
    'ppt/theme/theme1.xml': enc(THEME),
    'ppt/media/logo.png': new Uint8Array([137, 80, 78, 71]),
  };
}

describe('loadThemeColors', () => {
  const theme = loadThemeColors(deck(), 'slide1.xml');
  test('clrMap 별칭(tx1→dk1, sysClr lastClr)과 직접 RGB', () => {
    expect(theme.resolve({ 'a:schemeClr': { '@_val': 'tx1' } })).toBe('#000000');
    expect(theme.resolve({ 'a:schemeClr': { '@_val': 'accent1' } })).toBe('#4472C4');
    expect(theme.resolve({ 'a:srgbClr': { '@_val': 'c00000' } })).toBe('#C00000');
  });
  test('밝기 조정: lumMod/lumOff(HSL 밝기), tint(흰색과 혼합)', () => {
    // 검정(tx1) + lumMod 75% + lumOff 25% → 밝기 0.25 → #404040 (PowerPoint "검정, 텍스트 1, 25% 밝게")
    expect(theme.resolve({ 'a:schemeClr': { '@_val': 'tx1', 'a:lumMod': { '@_val': '75000' }, 'a:lumOff': { '@_val': '25000' } } })).toBe('#404040');
    expect(theme.resolve({ 'a:schemeClr': { '@_val': 'tx1', 'a:tint': { '@_val': '75000' } } })).toBe('#404040');
  });
  test('테마 파트가 없으면 테마 색을 추측하지 않는다(undefined), 직접 RGB는 그대로', () => {
    const bare = loadThemeColors({ 'ppt/slides/slide1.xml': enc('<p:sld/>') }, 'slide1.xml');
    expect(bare.resolve({ 'a:schemeClr': { '@_val': 'accent1' } })).toBeUndefined();
    expect(bare.resolve({ 'a:srgbClr': { '@_val': '123456' } })).toBe('#123456');
  });
});

describe('parseSlideVisuals: 레이아웃·마스터 그래픽과 배경', () => {
  test('마스터 배경(bgRef 테마 색) → 마스터 로고·색 띠 → 레이아웃 장식 순서, 개체 틀은 제외', () => {
    const v = parseSlideVisuals(deck(), 'slide1.xml');
    expect(v[0]).toEqual({ kind: 'shape', x: 0, y: 0, w: 9144000, h: 6858000, fill: '#4472C4', background: true });
    expect(v).toContainEqual({ kind: 'image', x: 100, y: 200, w: 300, h: 400, mediaPath: 'ppt/media/logo.png' });
    expect(v).toContainEqual(expect.objectContaining({ kind: 'shape', w: 9144000, h: 50, fill: '#17375E' })); // tx2(1F497D) 75%
    expect(v).toContainEqual(expect.objectContaining({ kind: 'shape', x: 10, y: 20, fill: '#ED7D31', geom: 'rightArrow' })); // fillRef → accent2
    expect(v.some((x) => x.fill === '#FF0000')).toBe(false); // 레이아웃 개체 틀은 그리지 않음
  });

  test('슬라이드 "배경 그래픽 숨기기"(showMasterSp=0)면 마스터·레이아웃 그래픽은 빼고 배경만', () => {
    const v = parseSlideVisuals(deck({ slideShowMasterSp: '0' }), 'slide1.xml');
    expect(v).toEqual([{ kind: 'shape', x: 0, y: 0, w: 9144000, h: 6858000, fill: '#4472C4', background: true }]);
  });

  test('명시적 noFill은 스타일 fillRef보다 우선(투명)', () => {
    const inner =
      `<p:sp><p:nvSpPr><p:cNvPr id="9" name="x"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(1, 2, 3, 4)}<a:noFill/></p:spPr>` +
      '<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></p:style></p:sp>';
    const v = parseSlideVisuals(deck({ slideInner: inner }), 'slide1.xml');
    expect(v.find((x) => x.x === 1 && x.y === 2)?.fill).toBeUndefined();
  });
});

describe('글자의 테마 색: 화면엔 풀어서, 저장엔 원본 그대로', () => {
  test('schemeClr+lumMod 글자 → 표시 색은 테마로 해석, 재직렬화는 원본 XML 유지', () => {
    const slideXml =
      `<p:sld ${NS}>` +
      tree(
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/>' +
          '<a:p><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:rPr><a:t>테마색</a:t></a:r></a:p>' +
          '</p:txBody></p:sp>',
      ) +
      '</p:sld>';
    const [shape] = parseSlideShapes(slideXml, loadThemeColors(deck(), 'slide1.xml'));
    const mark = shape.doc.firstChild!.firstChild!.marks.find((m) => m.type.name === 'textColor')!;
    expect(mark.attrs.color).toBe('#2F5597'); // 4472C4의 75% 밝기
    expect(serializeTxBodyXml(shape.doc)).toContain(
      '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>',
    );
  });
});

describe('parseSlideVisuals: 그라데이션 채우기와 레이아웃·마스터 글자 도형', () => {
  test('gradFill은 위치(pos) 순으로 정렬한 CSS linear-gradient — 첫 번째로 적힌 정지점 색 하나로 칠하지 않는다', () => {
    const inner =
      `<p:sp><p:nvSpPr><p:cNvPr id="9" name="g"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(1, 2, 3, 4)}` +
      '<a:gradFill><a:gsLst><a:gs pos="100000"><a:srgbClr val="4C5E78"/></a:gs><a:gs pos="38000"><a:srgbClr val="0339A6"/></a:gs></a:gsLst>' +
      '<a:lin ang="0" scaled="0"/></a:gradFill></p:spPr></p:sp>';
    const v = parseSlideVisuals(deck({ slideInner: inner }), 'slide1.xml');
    expect(v.find((x) => x.x === 1 && x.y === 2)?.fill).toBe('linear-gradient(90deg, #0339A6 38%, #4C5E78 100%)');
  });

  test('레이아웃의 글자 도형(쪽 번호·꼬리말 문구 등, 개체 틀 아님)은 원본 렌더링에서 잘라 그린다', () => {
    const zipEntries = deck();
    const layoutXml = new TextDecoder().decode(zipEntries['ppt/slideLayouts/slideLayout1.xml']);
    zipEntries['ppt/slideLayouts/slideLayout1.xml'] = enc(
      layoutXml.replace(
        '</p:spTree>',
        `<p:sp><p:nvSpPr><p:cNvPr id="7" name="footer"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(500, 600, 700, 80)}</p:spPr>` +
          '<p:txBody><a:bodyPr/><a:p><a:r><a:t>사용자 매뉴얼</a:t></a:r></a:p></p:txBody></p:sp>' +
          `<p:sp><p:nvSpPr><p:cNvPr id="8" name="num"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(900, 600, 50, 80)}</p:spPr>` +
          '<p:txBody><a:bodyPr/><a:p><a:fld id="{1}" type="slidenum"><a:t>‹#›</a:t></a:fld></a:p></p:txBody></p:sp>' +
          `<p:sp><p:nvSpPr><p:cNvPr id="10" name="empty"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(950, 600, 50, 80)}</p:spPr>` +
          '<p:txBody><a:bodyPr/><a:p><a:r><a:t> </a:t></a:r></a:p></p:txBody></p:sp></p:spTree>',
      ),
    );
    const v = parseSlideVisuals(zipEntries, 'slide1.xml');
    expect(v).toContainEqual({ kind: 'snapshot', x: 500, y: 600, w: 700, h: 80, snapshot: true });
    expect(v).toContainEqual({ kind: 'snapshot', x: 900, y: 600, w: 50, h: 80, snapshot: true });
    expect(v.some((x) => x.x === 950)).toBe(false); // 공백뿐인 글자 도형(채우기 없음)은 그릴 것이 없다
  });
});
