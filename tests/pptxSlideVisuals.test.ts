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
