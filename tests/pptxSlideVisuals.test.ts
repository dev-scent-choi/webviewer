import { describe, expect, test } from 'vitest';
import { parseSlideVisuals } from '../src/pptx/slideContent.js';
import { SLIDE_WIDTH_EMU, SLIDE_HEIGHT_EMU } from '../src/collab/pptxShared.js';

const XML_HEAD =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
  '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr/>';
const XML_TAIL = '</p:spTree></p:cSld></p:sld>';

function zip(entries: Record<string, string | Uint8Array>): Record<string, Uint8Array> {
  const out: Record<string, Uint8Array> = {};
  for (const [key, value] of Object.entries(entries)) {
    out[key] = typeof value === 'string' ? new TextEncoder().encode(value) : value;
  }
  return out;
}

describe('parseSlideVisuals', () => {
  test('extracts an image with its position and its resolved zip media path', () => {
    const slideXml =
      XML_HEAD +
      '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></a:xfrm></p:spPr></p:pic>' +
      XML_TAIL;
    const relsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId2" Type=".../image" Target="../media/image1.png"/>' +
      '</Relationships>';
    const entries = zip({
      'ppt/slides/slide1.xml': slideXml,
      'ppt/slides/_rels/slide1.xml.rels': relsXml,
      'ppt/media/image1.png': new Uint8Array([137, 80, 78, 71]), // 진짜 PNG일 필요는 없다 — 경로 해석 경로만 검증
    });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');

    expect(visuals).toEqual([
      { kind: 'image', x: 1000, y: 2000, w: 3000, h: 4000, mediaPath: 'ppt/media/image1.png' },
    ]);
  });

  test('an image the browser cannot render (emf) is cropped from the original rendering (snapshot)', () => {
    const slideXml =
      XML_HEAD +
      '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></p:spPr></p:pic>' +
      XML_TAIL;
    const relsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId2" Type=".../image" Target="../media/image1.emf"/>' +
      '</Relationships>';
    const entries = zip({
      'ppt/slides/slide1.xml': slideXml,
      'ppt/slides/_rels/slide1.xml.rels': relsXml,
      'ppt/media/image1.emf': new Uint8Array([1, 2, 3]),
    });

    expect(parseSlideVisuals(entries, 'slide1.xml')).toEqual([
      { kind: 'snapshot', x: 0, y: 0, w: 1, h: 1, snapshot: true },
    ]);
  });

  test('an image whose media part is missing is also left to the original rendering (snapshot)', () => {
    const slideXml =
      XML_HEAD +
      '<p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture"/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm></p:spPr></p:pic>' +
      XML_TAIL;
    const relsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId2" Type=".../image" Target="../media/missing.png"/>' +
      '</Relationships>';
    const entries = zip({
      'ppt/slides/slide1.xml': slideXml,
      'ppt/slides/_rels/slide1.xml.rels': relsXml,
      // ppt/media/missing.png는 의도적으로 없다.
    });

    expect(parseSlideVisuals(entries, 'slide1.xml')).toEqual([
      { kind: 'snapshot', x: 0, y: 0, w: 1, h: 1, snapshot: true },
    ]);
  });

  test('extracts a textless shape as a filled rectangle', () => {
    const slideXml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="5" name="Rectangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr></p:sp>' +
      XML_TAIL;
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');

    expect(visuals).toEqual([{ kind: 'shape', x: 100, y: 200, w: 300, h: 400, fill: '#4472C4' }]);
  });

  test('a shape with a <p:txBody> is left for parseSlideShapes — not duplicated here', () => {
    const slideXml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="6" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    expect(parseSlideVisuals(entries, 'slide1.xml')).toHaveLength(0);
  });

  test('extracts a table with real cell text content, row by row', () => {
    const slideXml =
      XML_HEAD +
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="7" name="Table"/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
      '<a:tbl><a:tr h="1"><a:tc><a:txBody><a:p><a:r><a:t>A1</a:t></a:r></a:p></a:txBody></a:tc>' +
      '<a:tc><a:txBody><a:p><a:r><a:t>B1</a:t></a:r></a:p></a:txBody></a:tc></a:tr>' +
      '<a:tr h="1"><a:tc><a:txBody><a:p><a:r><a:t>A2</a:t></a:r></a:p></a:txBody></a:tc>' +
      '<a:tc><a:txBody><a:p><a:r><a:t>B2</a:t></a:r></a:p></a:txBody></a:tc></a:tr>' +
      '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>' +
      XML_TAIL;
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');

    expect(visuals).toEqual([
      {
        kind: 'table',
        snapshot: true,
        x: 10,
        y: 20,
        w: 30,
        h: 40,
        rows: [
          ['A1', 'B1'],
          ['A2', 'B2'],
        ],
      },
    ]);
  });

  test('a chart graphicFrame becomes a labeled placeholder at the correct position', () => {
    const slideXml =
      XML_HEAD +
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="8" name="Chart"/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart r:id="rId3"/></a:graphicData></a:graphic></p:graphicFrame>' +
      XML_TAIL;
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    expect(parseSlideVisuals(entries, 'slide1.xml')).toEqual([
      { kind: 'placeholder', x: 1, y: 2, w: 3, h: 4, label: '차트', snapshot: true },
    ]);
  });

  test('a diagram (SmartArt) graphicFrame becomes a labeled placeholder', () => {
    const slideXml =
      XML_HEAD +
      '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="9" name="SmartArt"/></p:nvGraphicFramePr>' +
      '<p:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></p:xfrm>' +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
      '<dgm:relIds/></a:graphicData></a:graphic></p:graphicFrame>' +
      XML_TAIL;
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');
    expect(visuals[0]).toMatchObject({ kind: 'placeholder', label: '스마트아트' });
  });

  test('returns an empty array when the slide part is missing from the zip', () => {
    expect(parseSlideVisuals(zip({}), 'slide1.xml')).toEqual([]);
  });

  test('extracts a slide background image (<p:bg>, sibling of <p:spTree>) as a full-slide image visual, listed before other visuals', () => {
    const slideXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld>' +
      '<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rId9"/></a:blipFill><a:effectLst/></p:bgPr></p:bg>' +
      '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="5" name="Rectangle"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>' +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr></p:sp>' +
      '</p:spTree></p:cSld></p:sld>';
    const relsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId9" Type=".../image" Target="../media/bg.png"/>' +
      '</Relationships>';
    const entries = zip({
      'ppt/slides/slide1.xml': slideXml,
      'ppt/slides/_rels/slide1.xml.rels': relsXml,
      'ppt/media/bg.png': new Uint8Array([137, 80, 78, 71]),
    });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');

    expect(visuals).toEqual([
      { kind: 'image', x: 0, y: 0, w: SLIDE_WIDTH_EMU, h: SLIDE_HEIGHT_EMU, mediaPath: 'ppt/media/bg.png', background: true },
      { kind: 'shape', x: 100, y: 200, w: 300, h: 400, fill: '#4472C4' },
    ]);
  });

  test('extracts a slide background solid fill as a full-slide shape visual when there is no background image', () => {
    const slideXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld>' +
      '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>' +
      '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '</p:spTree></p:cSld></p:sld>';
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');

    expect(visuals).toEqual([
      { kind: 'shape', x: 0, y: 0, w: SLIDE_WIDTH_EMU, h: SLIDE_HEIGHT_EMU, fill: '#112233', background: true },
    ]);
  });

  test('finds an image nested inside a <p:grpSp> and converts its local coordinates to absolute slide coordinates', () => {
    // 그룹 off=(1000,2000) ext=(2000,4000), chOff=(0,0) chExt=(1000,1000) → 자식 좌표를 가로 2배,
    // 세로 4배로 확대. 자식 로컬 (x=100,y=100,w=200,h=200) → 절대 (1200, 2400, 400, 800).
    const slideXml =
      XML_HEAD +
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group 1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="2000" cy="4000"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></p:grpSpPr>' +
      '<p:pic><p:nvPicPr><p:cNvPr id="11" name="Picture"/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="200" cy="200"/></a:xfrm></p:spPr></p:pic>' +
      '</p:grpSp>' +
      XML_TAIL;
    const relsXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId2" Type=".../image" Target="../media/image1.png"/>' +
      '</Relationships>';
    const entries = zip({
      'ppt/slides/slide1.xml': slideXml,
      'ppt/slides/_rels/slide1.xml.rels': relsXml,
      'ppt/media/image1.png': new Uint8Array([137, 80, 78, 71]),
    });

    const visuals = parseSlideVisuals(entries, 'slide1.xml');

    expect(visuals).toEqual([
      { kind: 'image', x: 1200, y: 2400, w: 400, h: 800, mediaPath: 'ppt/media/image1.png' },
    ]);
  });

  test('skips a theme-referenced background (<p:bgRef>) rather than guessing a color', () => {
    const slideXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld>' +
      '<p:bg><p:bgRef idx="1001"><a:schemeClr val="accent1"/></p:bgRef></p:bg>' +
      '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>' +
      '</p:spTree></p:cSld></p:sld>';
    const entries = zip({ 'ppt/slides/slide1.xml': slideXml });

    expect(parseSlideVisuals(entries, 'slide1.xml')).toEqual([]);
  });
});

describe('편집 캔버스 충실도: 실제 슬라이드 크기 / 글자 있는 도형의 채우기 / 둥근 도형', () => {
  test('readSlideSizeEmu: presentation.xml의 4:3 크기를 읽고, 없으면 16:9로 대체한다', async () => {
    const { readSlideSizeEmu } = await import('../src/pptx/slideContent.js');
    const pres = '<p:presentation xmlns:p="p"><p:sldSz cy="6858000" type="screen4x3" cx="9144000"/></p:presentation>';
    expect(readSlideSizeEmu(zip({ 'ppt/presentation.xml': pres }))).toEqual({ cx: 9144000, cy: 6858000 });
    expect(readSlideSizeEmu(zip({}))).toEqual({ cx: SLIDE_WIDTH_EMU, cy: SLIDE_HEIGHT_EMU });
  });

  test('배경 그림/색은 실제 슬라이드 크기(4:3)로 깔린다', () => {
    const slideXml =
      XML_HEAD.replace('<p:cSld>', '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></p:bgPr></p:bg>') +
      XML_TAIL;
    const entries = zip({
      'ppt/slides/slide1.xml': slideXml,
      'ppt/presentation.xml': '<p:presentation xmlns:p="p"><p:sldSz cx="9144000" cy="6858000"/></p:presentation>',
    });
    expect(parseSlideVisuals(entries, 'slide1.xml')[0]).toMatchObject({ kind: 'shape', w: 9144000, h: 6858000, fill: '#112233' });
  });

  test('글자가 든 도형도 채우기 색이 있으면 색 판(+둥근 모양)을 깔고, 채우기 없으면 건너뛴다', () => {
    const slideXml =
      XML_HEAD +
      '<p:sp><p:nvSpPr><p:cNvPr id="5" name="Badge"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>' +
      '<a:solidFill><a:srgbClr val="FFC000"/></a:solidFill></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>뱃지</a:t></a:r></a:p></p:txBody></p:sp>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="6" name="Plain"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:t>그냥 글자</a:t></a:r></a:p></p:txBody></p:sp>' +
      XML_TAIL;
    const visuals = parseSlideVisuals(zip({ 'ppt/slides/slide1.xml': slideXml }), 'slide1.xml');
    expect(visuals).toHaveLength(1);
    expect(visuals[0]).toMatchObject({ kind: 'shape', x: 10, y: 20, w: 300, h: 400, fill: '#FFC000', geom: 'ellipse' });
  });
});

describe('parseSlideVisuals: 테두리·겹침 순서·회전된 선 (매뉴얼 PPT에서 빠지던 강조 표시)', () => {
  const rels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId2" Type=".../image" Target="../media/image1.png"/></Relationships>';
  const pic = (id: string, ln = '') =>
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="스크린샷"/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9000" cy="9000"/></a:xfrm>${ln}</p:spPr></p:pic>`;
  const redDotRect = (id: string) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="강조"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="sysDot"/></a:ln></p:spPr></p:sp>';
  const entriesFor = (body: string) =>
    zip({ 'ppt/slides/slide1.xml': XML_HEAD + body + XML_TAIL, 'ppt/slides/_rels/slide1.xml.rels': rels, 'ppt/media/image1.png': new Uint8Array([137, 80, 78, 71]) });

  test('채우기 없이 빨간 점선 테두리만 있는 사각형은 테두리로 그리고, 스크린샷보다 위에 놓인다', () => {
    const visuals = parseSlideVisuals(entriesFor(pic('2') + redDotRect('3')), 'slide1.xml');
    expect(visuals.map((v) => v.kind)).toEqual(['image', 'shape']);
    expect(visuals[1]).toMatchObject({ kind: 'shape', x: 100, y: 200, stroke: { color: '#FF0000', width: 25400, dash: 'dotted', prstDash: 'sysDot' } });
    expect(visuals[1].fill).toBeUndefined();
  });

  test('둥근 사각형 모서리는 원본 조정값으로 — adj 0은 직각, 없으면 기본 16667(짧은 변의 1/6)', () => {
    const rr = (id: string, av: string) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="버튼"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="600000" cy="300000"/></a:xfrm>` +
      `<a:prstGeom prst="roundRect"><a:avLst>${av}</a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp>`;
    const visuals = parseSlideVisuals(entriesFor(rr('2', '<a:gd name="adj" fmla="val 0"/>') + rr('3', '') + rr('4', '<a:gd name="adj" fmla="val 50000"/>')), 'slide1.xml');
    expect(visuals.map((v) => v.corners)).toEqual([[0, 0, 0, 0], [50001, 50001, 50001, 50001], [150000, 150000, 150000, 150000]]);
  });

  test('실선 테두리는 점선 이름을 싣지 않는다', () => {
    const visuals = parseSlideVisuals(entriesFor(pic('2', '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>')), 'slide1.xml');
    expect(visuals[0].stroke?.prstDash).toBeUndefined();
  });

  test('겹침 순서는 원본 순서 — 테두리 상자가 그림보다 앞(아래)에 있으면 아래에 둔다', () => {
    const visuals = parseSlideVisuals(entriesFor(redDotRect('3') + pic('2')), 'slide1.xml');
    expect(visuals.map((v) => v.kind)).toEqual(['shape', 'image']);
  });

  test('그림에 직접 준 테두리도 읽는다', () => {
    const visuals = parseSlideVisuals(entriesFor(pic('2', '<a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>')), 'slide1.xml');
    expect(visuals[0]).toMatchObject({ kind: 'image', stroke: { color: '#FF0000', dash: 'solid' } });
  });

  test('90° 돌린 가로 연결선(= 실제로는 세로선)은 돌린 뒤 영역을 넉넉히 잘라 온다', () => {
    const cxn =
      '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="5" name="연결선"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>' +
      '<a:xfrm rot="5400000"><a:off x="0" y="1000000"/><a:ext cx="2000000" cy="0"/></a:xfrm><a:prstGeom prst="bentConnector3"><a:avLst/></a:prstGeom>' +
      '<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></p:spPr></p:cxnSp>';
    const [line] = parseSlideVisuals(entriesFor(cxn), 'slide1.xml');
    // 중심 (1000000, 1000000) 기준 세로 2000000 — 사방으로 약 6px(57150 EMU) 넓힌다
    expect(line).toMatchObject({ kind: 'snapshot', snapshot: true });
    expect(line.x).toBeCloseTo(1000000 - 57150);
    expect(line.y).toBeCloseTo(0 - 57150);
    expect(line.w).toBeCloseTo(57150 * 2);
    expect(line.h).toBeCloseTo(2000000 + 57150 * 2);
  });
});
