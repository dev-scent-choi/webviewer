import { describe, expect, test } from 'vitest';
import { loadThemeColors } from '../src/pptx/theme.js';
import { parseSlideShapes, parseSlideVisuals, serializeTxBodyXml, spliceShapeTxBody } from '../src/pptx/slideContent.js';
import { resolveSlideShapeLooks } from '../src/pptx/inheritance.js';
import { autoNumberPlan, autoNumberText, pPrLook } from '../src/pptx/paragraphLook.js';
import { pptxTextSchema as schema } from '../src/pptx/textSchema.js';

// 서버가 원본에서 읽은 조각 목록(검토 I2) — 이 파일의 파싱·저장 시험이 함께 쓴다
const trusted = new Set<string>();
// 품질 점검(docs/quality-audit/REPORT.md 4.3 P1~P6) 고친 것들의 단위 시험 — 합성 최소 덱
const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const enc = (s: string) => new TextEncoder().encode(s);
const tree = (inner: string) =>
  `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${inner}</p:spTree></p:cSld>`;
const xfrm = (x: number, y: number, w: number, h: number, attrs = '') => `<a:xfrm${attrs}><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`;
const rels = (entries: Array<[string, string, string]>) =>
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  entries.map(([id, type, target]) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="${target}"/>`).join('') +
  '</Relationships>';

const THEME =
  `<a:theme ${NS} name="T"><a:themeElements><a:clrScheme name="c">` +
  '<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3>' +
  '<a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>' +
  '<a:fontScheme name="f"><a:majorFont><a:latin typeface="Georgia"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hang" typeface="HY견고딕"/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Segoe UI"/><a:ea typeface="맑은 고딕"/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="m"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill>' +
  '<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="50000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000"/></a:gradFill></a:fillStyleLst>' +
  '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
  '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>' +
  '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';

const STYLE =
  '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';

function deck(slideInner: string, layoutInner = '', masterExtra = ''): Record<string, Uint8Array> {
  return {
    'ppt/slides/slide1.xml': enc(`<p:sld ${NS}>${tree(slideInner)}</p:sld>`),
    'ppt/slides/_rels/slide1.xml.rels': enc(rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ['rId2', 'image', '../media/image1.png']])),
    'ppt/slideLayouts/slideLayout1.xml': enc(`<p:sldLayout ${NS}>${tree(layoutInner)}</p:sldLayout>`),
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': enc(rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']])),
    'ppt/slideMasters/slideMaster1.xml': enc(
      `<p:sldMaster ${NS}>${tree('')}<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
        `<p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></p:titleStyle>` +
        `<p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial"/><a:buChar char="•"/><a:defRPr sz="3200"><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr>` +
        `<a:lvl2pPr marL="742950" indent="-285750"><a:buAutoNum type="alphaLcParenR"/><a:defRPr sz="2800"/></a:lvl2pPr></p:bodyStyle>` +
        `<p:otherStyle><a:lvl1pPr algn="l"><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle></p:txStyles>${masterExtra}</p:sldMaster>`,
    ),
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': enc(rels([['rId1', 'theme', '../theme/theme1.xml']])),
    'ppt/theme/theme1.xml': enc(THEME),
    'ppt/media/image1.png': new Uint8Array([137, 80, 78, 71]),
  };
}

const sp = (id: number, nv: string, spPr: string, txBody: string, style = '') =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s${id}"/><p:cNvSpPr/><p:nvPr>${nv}</p:nvPr></p:nvSpPr><p:spPr>${spPr}</p:spPr>${style}<p:txBody>${txBody}</p:txBody></p:sp>`;

describe('P1 자동 맞춤·세로 쓰기(bodyPr)', () => {
  test('normAutofit fontScale/lnSpcReduction을 비율로, 레이아웃에선 맞춤 종류만 물려받고(배율은 1), spAutoFit은 맞춤 없음', () => {
    const zip = deck(
      sp(2, '<p:ph idx="1"/>', '', '<a:bodyPr><a:normAutofit fontScale="47500" lnSpcReduction="20000"/></a:bodyPr><a:p/>') +
        sp(3, '<p:ph type="body" idx="2"/>', '', '<a:bodyPr/><a:p/>') +
        sp(4, '', xfrm(0, 0, 10, 10), '<a:bodyPr vert="eaVert"><a:spAutoFit/></a:bodyPr><a:p/>'),
      sp(3, '<p:ph type="body" idx="2"/>', xfrm(1, 1, 1, 1), '<a:bodyPr><a:normAutofit fontScale="90000"/></a:bodyPr><a:p/>'),
    );
    const looks = resolveSlideShapeLooks(zip, 'slide1.xml');
    expect(looks.get('2')?.autofit).toEqual({ fontScale: 0.475, lnSpcReduction: 0.2 });
    expect(looks.get('3')?.autofit).toEqual({ fontScale: 1, lnSpcReduction: 0 }); // 검토 Minor 3: 배율은 상속 안 함
    expect(looks.get('4')?.autofit).toBeUndefined();
    expect(looks.get('4')?.vert).toBe('eaVert');
    expect(looks.get('2')?.hlinkColor).toBe('#0000FF');
  });
});

describe('P2 그림 자르기·회전·뒤집기, 그룹 회전', () => {
  const pic = (id: number, extra: string, x = 0, attrs = '') =>
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="p"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId2"/>${extra}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrm(x, 0, 1000, 1000, attrs)}</p:spPr></p:pic>`;
  test('srcRect는 [왼, 위, 오른, 아래] 비율, rot는 도, flipH는 true', () => {
    const v = parseSlideVisuals(deck(pic(2, '<a:srcRect l="25000" t="10000" r="5000"/>') + pic(3, '', 2000, ' rot="900000" flipH="1"')), 'slide1.xml');
    expect(v[0]).toMatchObject({ kind: 'image', crop: [0.25, 0.1, 0.05, 0] });
    expect(v[1]).toMatchObject({ kind: 'image', rot: 15, flipH: true });
    expect(v[1].crop).toBeUndefined();
  });
  test('90° 돌린 그룹 안 도형: 중심이 그룹 중심을 기준으로 돌고, 회전각이 더해진다', () => {
    // 그룹 0,0 2000x1000(회전 90°), 자식 도형은 왼쪽 절반(0,0 1000x1000) — 돌면 중심 (500,500) → (1000,0)
    const group =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="2000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="1000"/></a:xfrm></p:grpSpPr>` +
      sp(5, '', xfrm(0, 0, 1000, 1000, ' rot="600000"'), '<a:bodyPr/><a:p><a:r><a:t>A</a:t></a:r></a:p>') +
      '</p:grpSp>';
    const [shape] = parseSlideShapes(`<p:sld ${NS}>${tree(group)}</p:sld>`);
    expect(shape).toMatchObject({ x: 500, y: -500, w: 1000, h: 1000, rot: 100 });
  });
});

describe('P3·P4 수준 상속: 간격·글꼴·굵기·글머리·자동 번호·한글 줄 바꿈', () => {
  const zip = deck(
    sp(2, '<p:ph type="title"/>', '', '<a:bodyPr/><a:p/>') +
      sp(3, '<p:ph idx="1"/>', '', '<a:bodyPr/><a:lstStyle><a:lvl1pPr eaLnBrk="0"><a:lnSpc><a:spcPct val="130000"/></a:lnSpc><a:spcAft><a:spcPts val="600"/></a:spcAft><a:buClr><a:schemeClr val="accent2"/></a:buClr><a:buSzPct val="80000"/><a:defRPr b="1" cap="all"/></a:lvl1pPr></a:lstStyle><a:p/>'),
    sp(2, '<p:ph type="title"/>', xfrm(1, 1, 1, 1), '<a:bodyPr/><a:p/>'),
  );
  const looks = resolveSlideShapeLooks(zip, 'slide1.xml');
  test('제목은 테마 major 글꼴(라틴 Georgia, 한글 ea가 비면 Hang 글꼴)', () => {
    expect(looks.get('2')?.levels[0]).toMatchObject({ fontLatin: 'Georgia', fontEa: 'HY견고딕', sizePt: 44 });
  });
  test('본문 1수준: 직접 적힌 간격·굵게·대문자·글머리 꾸밈 + 마스터 spcBef·글머리 글꼴, eaLnBrk=0이면 어절 단위', () => {
    expect(looks.get('3')?.levels[0]).toMatchObject({
      spacing: { line: { pct: 1.3 }, before: { pct: 0.2 }, after: { pts: 6 } },
      bold: true,
      caps: true,
      keepWords: true,
      bullet: '•',
      bulletFont: 'Arial',
      bulletColor: '#C0504D',
      bulletSize: { pct: 0.8 },
      fontLatin: 'Segoe UI',
      fontEa: '맑은 고딕',
    });
  });
  test('2수준은 마스터의 자동 번호(buAutoNum)', () => {
    expect(looks.get('3')?.levels[1].autoNum).toEqual({ scheme: 'alphaLcParenR', startAt: 1 });
  });
});

describe('paragraphLook: 문단 직접 속성·자동 번호', () => {
  test('pPrLook은 줄·단락 간격, 글머리 재정의, eaLnBrk를 읽는다', () => {
    expect(pPrLook('<a:pPr marL="342900" indent="-342900" eaLnBrk="0"><a:lnSpc><a:spcPts val="2400"/></a:lnSpc><a:spcBef><a:spcPct val="50000"/></a:spcBef><a:buClr><a:srgbClr val="ff0000"/></a:buClr><a:buAutoNum type="romanUcPeriod" startAt="3"/></a:pPr>')).toEqual({
      line: { pts: 24 },
      before: { pct: 0.5 },
      bullet: { kind: 'auto', scheme: 'romanUcPeriod', startAt: 3 },
      keepWords: true,
      marLPt: 27,
      indentPt: -27,
    });
    expect(pPrLook('<a:pPr><a:buChar char="&amp;"/></a:pPr>').bullet).toEqual({ kind: 'char', char: '&' });
    expect(pPrLook(null)).toEqual({});
  });
  test('번호 체계', () => {
    expect([autoNumberText('arabicPeriod', 3), autoNumberText('alphaLcParenR', 28), autoNumberText('romanUcPeriod', 9), autoNumberText('arabicParenBoth', 2)]).toEqual(['3.', 'ab)', 'IX.', '(2)']);
    expect([autoNumberText('circleNumDbPlain', 2), autoNumberText('ea1JpnKorPeriod', 3), autoNumberText('모름', 4)]).toEqual(['②', '三.', '4.']);
  });
  test('번호는 같은 수준에서 이어지고, 번호 아닌 문단에서 끊기며, 얕은 수준이 나오면 깊은 수준은 다시 시작', () => {
    const para = (lvl: number, text: string, pPr: string | null = null) => schema.node('paragraph', { lvl, pPr }, text ? [schema.text(text)] : []);
    const auto = '<a:pPr><a:buAutoNum type="arabicPeriod"/></a:pPr>';
    const doc = schema.node('doc', null, [
      para(0, 'a', auto), para(0, 'b', auto), para(1, 'c'), para(1, 'd'), para(0, 'e', auto), para(0, ''), para(0, 'f', auto), para(0, 'g'), para(0, 'h', auto),
    ]);
    const levels = [undefined, { marLPt: 0, indentPt: 0, bullet: '•', sizePt: 18, autoNum: { scheme: 'alphaLcPeriod', startAt: 1 } }];
    expect(autoNumberPlan(doc, levels as never)).toEqual(['1.', '2.', 'a.', 'b.', '3.', null, '4.', null, '1.']);
  });
});

describe('P5 테마 서식(fmtScheme): 채우기 그라데이션·선 두께·그림자', () => {
  test('fillRef 3 → 테마 그라데이션(phClr = accent1), lnRef 2 → 2pt 선, effectRef 2 → 그림자', () => {
    const v = parseSlideVisuals(deck(sp(2, '', `${xfrm(0, 0, 1000, 1000)}<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>`, '<a:bodyPr/><a:p><a:r><a:t>1</a:t></a:r></a:p>', STYLE)), 'slide1.xml');
    expect(v[0]).toMatchObject({
      kind: 'shape',
      geom: 'ellipse',
      fill: 'linear-gradient(0deg, #4F81BD 0%, #385D8A 100%)',
      stroke: { color: '#4F81BD', width: 25400 },
      shadow: { color: 'rgba(0, 0, 0, 0.35)', dx: 0, dy: 23000, blur: 40000 },
    });
  });
  test('직접 <a:effectLst/>(빈 목록)는 테마 그림자를 끈다', () => {
    const v = parseSlideVisuals(deck(sp(2, '', `${xfrm(0, 0, 1, 1)}<a:effectLst/>`, '<a:bodyPr/><a:p/>', STYLE)), 'slide1.xml');
    expect(v[0].shadow).toBeUndefined();
  });
  test('테마 글꼴 참조 풀기', () => {
    const theme = loadThemeColors(deck(''), 'slide1.xml');
    expect([theme.font('+mj-lt'), theme.font('+mj-ea'), theme.font('+mn-ea'), theme.font('Arial'), theme.font('+mn-cs')]).toEqual(['Georgia', 'HY견고딕', '맑은 고딕', 'Arial', undefined]);
  });
});

describe('저장 보존: 줄 바꿈·필드·모델링하지 않는 글자 속성(하이퍼링크·lang)', () => {
  const slide =
    `<p:sld ${NS}>` +
    tree(
      sp(2, '', xfrm(0, 0, 10, 10), '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" b="1" spc="-50"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="Georgia"/><a:hlinkClick r:id="rId9"/></a:rPr><a:t>링크</a:t></a:r>' +
        '<a:br><a:rPr lang="ko-KR"/></a:br><a:r><a:t>둘째 줄 </a:t></a:r><a:fld id="{X}" type="slidenum"><a:rPr lang="en-US"/><a:t>3</a:t></a:fld><a:endParaRPr lang="ko-KR" sz="1200"/></a:p>'),
    ) +
    '</p:sld>';
  test('글자 순서대로 br·필드가 읽히고, 편집 안 한 그대로 저장하면 원문의 속성이 모두 남는다', () => {
    const [shape] = parseSlideShapes(slide, undefined, trusted);
    const para = shape.doc.child(0);
    expect(para.textContent).toBe('링크둘째 줄 ');
    expect(para.child(3).attrs.text).toBe('3');
    expect([...Array(para.childCount).keys()].map((i) => para.child(i).type.name)).toEqual(['text', 'br', 'text', 'field']);
    const xml = serializeTxBodyXml(shape.doc, trusted);
    expect(xml).toContain('<a:rPr lang="ko-KR" b="1" spc="-50"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="Georgia"/><a:hlinkClick r:id="rId9"/></a:rPr>');
    expect(xml).toContain('<a:br><a:rPr lang="ko-KR"/></a:br>');
    expect(xml).toContain('<a:fld id="{X}" type="slidenum"><a:rPr lang="en-US"/><a:t>3</a:t></a:fld>');
    expect(xml).toContain('<a:endParaRPr lang="ko-KR" sz="1200"/></a:p>');
    expect(xml.indexOf('<a:br>')).toBeLessThan(xml.indexOf('둘째'));
  });
  test('굵게를 끄고 글자를 고쳐도 하이퍼링크·lang·자간·테마 색은 남고 b만 빠진다', () => {
    const [shape] = parseSlideShapes(slide, undefined, trusted);
    const first = shape.doc.child(0).child(0);
    const marks = first.marks.filter((m) => m.type.name !== 'strong');
    const para = shape.doc.child(0).copy(shape.doc.child(0).content.replaceChild(0, schema.text('새 링크', marks)));
    const xml = serializeTxBodyXml(schema.node('doc', null, [para]), trusted);
    expect(xml).toContain('<a:r><a:rPr lang="ko-KR" spc="-50"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="Georgia"/><a:hlinkClick r:id="rId9"/></a:rPr><a:t>새 링크</a:t></a:r>');
    const spliced = spliceShapeTxBody(slide, '2', xml);
    expect(parseSlideShapes(spliced)[0].doc.child(0).textContent).toBe('새 링크둘째 줄 ');
    expect(spliced).toContain('<a:fld id="{X}" type="slidenum">');
  });
  test('글꼴을 바꾸면 latin·ea를 새 이름으로 쓴다(나머지는 유지)', () => {
    const [shape] = parseSlideShapes(slide, undefined, trusted);
    const first = shape.doc.child(0).child(0);
    const marks = first.marks.filter((m) => m.type.name !== 'fontFamily').concat(schema.marks.fontFamily.create({ family: '바탕' }));
    const xml = serializeTxBodyXml(schema.node('doc', null, [schema.node('paragraph', null, [schema.text('x', marks)])]), trusted);
    expect(xml).toContain('<a:latin typeface="바탕"/><a:ea typeface="바탕"/><a:hlinkClick r:id="rId9"/>');
  });
});

describe('P-v2 LibreOffice 전처리: 표 머리 줄 굵게', () => {
  const tbl = (style: string, firstRow = '1') =>
    `<a:tbl><a:tblPr firstRow="${firstRow}" bandRow="1"><a:tableStyleId>${style}</a:tableStyleId></a:tblPr>` +
    '<a:tr h="1"><a:tc><a:txBody><a:p><a:r><a:rPr lang="ko-KR"/><a:t>구분</a:t></a:r><a:r><a:t>X</a:t></a:r><a:r><a:rPr b="0"/><a:t>Y</a:t></a:r></a:p></a:txBody></a:tc></a:tr>' +
    '<a:tr h="1"><a:tc><a:txBody><a:p><a:r><a:rPr lang="ko-KR"/><a:t>1단계</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl>';
  test('내장 스타일 + 머리글 행: 첫 줄만 b="1"(직접 적힌 b는 그대로)', async () => {
    const { boldTableHeaderXml } = await import('../src/pptx/loPreprocess.js');
    const out = boldTableHeaderXml(tbl('{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'), '');
    expect(out).toContain('<a:rPr b="1" lang="ko-KR"/><a:t>구분');
    expect(out).toContain('<a:r><a:rPr b="1"/><a:t>X');
    expect(out).toContain('<a:rPr b="0"/><a:t>Y');
    expect(out).toContain('<a:rPr lang="ko-KR"/><a:t>1단계');
  });
  test('머리글 행 꺼짐·서식 없는 내장 스타일·정의상 굵게 아님은 그대로', async () => {
    const { boldTableHeaderXml } = await import('../src/pptx/loPreprocess.js');
    expect(boldTableHeaderXml(tbl('{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}', '0'), '')).toBe(tbl('{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}', '0'));
    expect(boldTableHeaderXml(tbl('{2D5ABB26-0587-4C30-8999-92F81FD0307C}'), '')).toBe(tbl('{2D5ABB26-0587-4C30-8999-92F81FD0307C}'));
    const id = '{11111111-2222-3333-4444-555555555555}';
    const styles = `<a:tblStyleLst><a:tblStyle styleId="${id}" styleName="x"><a:firstRow><a:tcTxStyle b="off"/></a:firstRow></a:tblStyle></a:tblStyleLst>`;
    expect(boldTableHeaderXml(tbl(id), styles)).toBe(tbl(id));
    expect(boldTableHeaderXml(tbl(id), styles.replace('b="off"', 'b="on"'))).toContain('<a:rPr b="1" lang="ko-KR"/><a:t>구분');
  });
  test('검토 C1: 스타일 id는 정규식에 넣지 않는다 — 역추적 폭발·잘못된 정규식 값도 즉시 끝나고 굵게 안 함', async () => {
    const { boldTableHeaderXml } = await import('../src/pptx/loPreprocess.js');
    const styles = `<a:tblStyleLst><a:tblStyle styleId="${'a'.repeat(40)}!"></a:tblStyle></a:tblStyleLst>`;
    const t0 = Date.now();
    for (const id of ['(a|a)+$', '{X(}', `${'a'.repeat(40)}!`]) expect(boldTableHeaderXml(tbl(id), styles)).toBe(tbl(id));
    expect(Date.now() - t0).toBeLessThan(100);
  });
});

describe('슬라이드 배경 bgRef → 테마 배경 채우기 목록(bgFillStyleLst)', () => {
  test('idx 1001 = 첫 배경 채우기(phClr = bgRef 색)', () => {
    const zip = deck('');
    zip['ppt/slides/slide1.xml'] = enc(`<p:sld ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="accent2"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>`);
    expect(parseSlideVisuals(zip, 'slide1.xml')[0]).toMatchObject({ kind: 'shape', background: true, fill: '#C0504D' });
  });
});

describe('검토 I1: 표 머리 줄 전처리는 바뀐 슬라이드만 다시 묶고 무압축 미디어는 그대로', () => {
  test('30MB 무압축 미디어가 있어도 빠르게(재압축 없음), 이벤트 루프를 길게 막지 않음', async () => {
    const { zipSync, unzipSync } = await import('fflate');
    const { readFileSync } = await import('node:fs');
    const { randomBytes } = await import('node:crypto');
    const { boldPptxTableHeaders } = await import('../src/pptx/loPreprocess.js');
    const entries: Record<string, unknown> = { ...unzipSync(readFileSync('audit/office/corpus/pptx/p04-table.pptx')) };
    const media = randomBytes(30 * 1024 * 1024);
    entries['ppt/media/video1.mp4'] = [media, { level: 0 }];
    const big = zipSync(entries as never);
    let maxGap = 0;
    let last = Date.now();
    const tick = setInterval(() => { const n = Date.now(); maxGap = Math.max(maxGap, n - last); last = n; }, 5);
    const t0 = Date.now();
    const out = await boldPptxTableHeaders(big);
    const took = Date.now() - t0;
    clearInterval(tick);
    expect(took).toBeLessThan(1500);
    expect(maxGap).toBeLessThan(400);
    expect(Math.abs(out.length - big.length)).toBeLessThan(64 * 1024); // 미디어를 다시 압축하지 않았다
    const back = unzipSync(out);
    expect(Buffer.from(back['ppt/media/video1.mp4']).equals(media)).toBe(true);
    expect(new TextDecoder().decode(back['ppt/slides/slide1.xml'])).toContain('b="1"');
  }, 20000);
});

describe('검토 I3: 숫자 문자 참조(&#xAC00;)는 글자로 읽고, 편집 후 저장에서 망가지지 않는다', () => {
  const xml =
    `<p:sld ${NS}>` +
    tree(sp(2, '', xfrm(0, 0, 1, 1), '<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR"><a:latin typeface="&#xB9D1;&#xC740; &#xACE0;&#xB515;"/><a:ea typeface="Evil&#10;Font"/></a:rPr><a:t>&#xAC00;&#xB098;&#xB2E4; &amp; &lt;x&gt;</a:t></a:r></a:p>')) +
    '</p:sld>';
  test('화면 글자는 "가나다", 한 글자 고친 저장본은 글자·글꼴 이름을 올바르게 다시 쓴다', () => {
    const doc = parseSlideShapes(xml, undefined, trusted)[0].doc;
    expect(doc.textContent).toBe('가나다 & <x>');
    const run = doc.child(0).child(0);
    expect(run.marks.find((m) => m.type.name === 'fontFamily')?.attrs.family).toBe('Evil\nFont');
    const edited = schema.node('doc', null, [schema.node('paragraph', null, [schema.text('가나다! & <x>', run.marks)])]);
    const out = serializeTxBodyXml(edited, trusted);
    expect(out).toContain('<a:t>가나다! &amp; &lt;x&gt;</a:t>');
    expect(out).toContain('typeface="맑은 고딕"');
    expect(out).toContain('typeface="Evil&#10;Font"');
    expect(out).not.toContain('&amp;#');
  });
});

describe('검토 I2: 협업자가 쓴 원문 문자열은 저장 XML에 붙지 않는다(원본에서 읽은 조각만)', () => {
  test('끝 속성·필드·줄 바꿈·글자 속성·글자 색에 넣은 주입 문자열이 나오지 않고, 결과는 올바른 XML', async () => {
    const { XMLValidator } = await import('fast-xml-parser');
    const evil = '</a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="999" name="injected"/>';
    const marks = [schema.marks.runProps.create({ raw: `<a:rPr/>${evil}` }), schema.marks.textColor.create({ color: '#FF0000', raw: `<a:srgbClr val="00FF00"/>${evil}` })];
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { endRPr: evil, pPr: `<a:pPr/>${evil}` }, [
        schema.text('a', marks),
        schema.node('field', { raw: '<a:fld id="x"><a:t>1</a:t></a:fld><broken', text: '1' }),
        schema.node('br', { rPr: '<a:rPr/><zz' }),
      ]),
    ]);
    const xml = serializeTxBodyXml(doc, new Set());
    expect(xml).not.toContain('injected');
    expect(xml).not.toContain('<broken');
    expect(xml).not.toContain('<zz');
    expect(xml).toBe('<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr><a:t>a</a:t></a:r><a:r><a:t>1</a:t></a:r><a:br/></a:p></p:txBody>');
    const ns = 'xmlns:a="urn:a" xmlns:p="urn:p"';
    expect(XMLValidator.validate(xml.replace('<p:txBody>', `<p:txBody ${ns}>`))).toBe(true);
  });
});

describe('검토 I4: 문단 안 수식(mc:AlternateContent/a14:m)은 원문째 보존하고 대체 글자를 그 자리에 보인다', () => {
  const MATH =
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>mc2</m:t></m:r></m:oMath></m:oMathPara></a14:m></mc:Choice><mc:Fallback><a:r><a:rPr lang="en-US"/><a:t>mc²</a:t></a:r></mc:Fallback></mc:AlternateContent>';
  const slide = `<p:sld ${NS}>${tree(sp(2, '', xfrm(0, 0, 1, 1), `<a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>E=</a:t></a:r>${MATH}<a:r><a:rPr lang="en-US"/><a:t> 끝</a:t></a:r></a:p>`))}</p:sld>`;
  test('순서대로 읽고(대체 글자 mc²), 한 글자 고쳐 저장해도 수식 원문이 그대로', () => {
    const set = new Set<string>();
    const [shape] = parseSlideShapes(slide, undefined, set);
    const para = shape.doc.child(0);
    expect([...Array(para.childCount).keys()].map((i) => para.child(i).type.name)).toEqual(['text', 'field', 'text']);
    expect(para.child(1).attrs.text).toBe('mc²');
    const edited = schema.node('doc', null, [para.copy(para.content.replaceChild(0, schema.text('E =', para.child(0).marks)))]);
    const out = spliceShapeTxBody(slide, '2', serializeTxBodyXml(edited, set));
    expect(out).toContain(`<a:t>E =</a:t></a:r>${MATH}<a:r>`);
    expect(out).not.toContain('_o=');
  });
});

describe('검토 I5: 슬라이드 최상위 mc:AlternateContent 개체는 대체 표시 자리에 원본 렌더링 조각으로 보이고, 편집 대상이 아니다', () => {
  test('Fallback 도형 자리 → snapshot, 글상자 목록엔 없음(저장은 원문 그대로)', () => {
    const box = (id: number, x: number, text: string) =>
      `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(x, 0, 100000, 50000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
    const inner = `${box(2, 0, '보통')}<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">${box(3, 500000, '수식(Choice)')}</mc:Choice><mc:Fallback>${box(3, 500000, '수식 그림(Fallback)')}</mc:Fallback></mc:AlternateContent>`;
    const zip = deck(inner);
    expect(parseSlideShapes(new TextDecoder().decode(zip['ppt/slides/slide1.xml'])).map((s) => s.shapeId)).toEqual(['2']);
    expect(parseSlideVisuals(zip, 'slide1.xml')).toEqual([{ kind: 'snapshot', x: 500000, y: 0, w: 100000, h: 50000, snapshot: true }]);
  });
});

describe('검토 Minor 1: 저장 zip은 시각과 무관하게 같은 내용이면 같은 바이트', () => {
  test('두 번 내보내는 사이 시계가 1분 지나도 바이트 동일', async () => {
    const { vi } = await import('vitest');
    const Y = await import('yjs');
    const { exportPptxFromSlides, seedSlidesFromPptx } = await import('../src/pptx/collab.js');
    const { makeSamplePptx } = await import('./helpers.js');
    const { PPTX_SLIDES_ARRAY } = await import('../src/collab/pptxShared.js');
    const original = makeSamplePptx(['A', 'B']);
    const slides = new Y.Doc().getArray(PPTX_SLIDES_ARRAY) as never;
    seedSlidesFromPptx(original, slides);
    vi.useFakeTimers();
    try {
      vi.setSystemTime(new Date(2026, 10, 20, 12, 0, 0));
      const a = exportPptxFromSlides(original, slides).bytes;
      vi.setSystemTime(new Date(2026, 10, 20, 12, 1, 1));
      const b = exportPptxFromSlides(original, slides).bytes;
      expect(Buffer.from(a).equals(Buffer.from(b))).toBe(true);
    } finally {
      vi.useRealTimers();
    }
  });
});

describe('검토 Minor 3: 다시 맞춘 자동 맞춤 값을 저장(normAutofit 도형만, 값은 범위 검사)', () => {
  test('setShapeAutofit', async () => {
    const { setShapeAutofit } = await import('../src/pptx/slideContent.js');
    const slide = (bodyPr: string) => `<p:sld ${NS}>${tree(sp(2, '', xfrm(0, 0, 1, 1), `${bodyPr}<a:lstStyle/><a:p/>`))}</p:sld>`;
    const norm = slide('<a:bodyPr><a:normAutofit fontScale="62500" lnSpcReduction="20000"/></a:bodyPr>');
    expect(setShapeAutofit(norm, '2', { fontScale: 0.85, lnSpcReduction: 0.2 })).toContain('<a:normAutofit fontScale="85000" lnSpcReduction="20000"/>');
    expect(setShapeAutofit(norm, '2', { fontScale: 1, lnSpcReduction: 0 })).toContain('<a:bodyPr><a:normAutofit/></a:bodyPr>');
    expect(setShapeAutofit(norm, '2', { fontScale: 99, lnSpcReduction: -5 })).toContain('<a:normAutofit/>');
    expect(setShapeAutofit(norm, '2', { fontScale: '"/><x', lnSpcReduction: 0 })).toBe(norm);
    const spFit = slide('<a:bodyPr><a:spAutoFit/></a:bodyPr>');
    expect(setShapeAutofit(spFit, '2', { fontScale: 0.5, lnSpcReduction: 0.2 })).toBe(spFit);
  });
});

describe('검토 Minor 5: 색 변환은 파일 순서대로, hue·comp·inv·gray 등도', () => {
  const theme = loadThemeColors(deck(''), 'slide1.xml');
  test('순서가 다르면 결과가 다르다(lumOff 뒤 shade ≠ shade 뒤 lumOff)', () => {
    const a = theme.resolve({ 'a:srgbClr': { '@_val': '808080', 'a:lumOff': { '@_val': '20000' }, 'a:shade': { '@_val': '50000' } } });
    const b = theme.resolve({ 'a:srgbClr': { '@_val': '808080', 'a:shade': { '@_val': '50000' }, 'a:lumOff': { '@_val': '20000' } } });
    expect(a).not.toBe(b);
  });
  test('comp·inv·gray·hueOff', () => {
    expect(theme.resolve({ 'a:srgbClr': { '@_val': 'FF0000', 'a:comp': {} } })).toBe('#00FFFF');
    expect(theme.resolve({ 'a:srgbClr': { '@_val': '102030', 'a:inv': {} } })).toBe('#EFDFCF');
    expect(theme.resolve({ 'a:srgbClr': { '@_val': 'FF0000', 'a:gray': {} } })).toBe('#4D4D4D');
    expect(theme.resolve({ 'a:srgbClr': { '@_val': 'FF0000', 'a:hueOff': { '@_val': '7200000' } } })).toBe('#00FF00');
  });
});

describe('검토 Minor 7: 자동 번호 체계 전부(41가지)·한자 수 11 이상', () => {
  test('대표 값', () => {
    const cases: Array<[string, number, string]> = [
      ['ea1JpnKorPlain', 11, '十一'], ['ea1JpnKorPeriod', 20, '二十.'], ['ea1ChsPlain', 105, '百五'], ['ea1JpnChsDbPeriod', 3, '三．'],
      ['arabic1Minus', 2, '2-'], ['arabic2Minus', 2, '-2-'], ['hebrew2Minus', 1, '-א-'], ['thaiAlphaPeriod', 2, 'ข.'],
      ['thaiNumParenBoth', 12, '(๑๒)'], ['hindiNumPeriod', 3, '३.'], ['hindiAlphaPeriod', 1, 'अ.'], ['hindiAlpha1Period', 2, 'ख.'],
      ['circleNumWdBlackPlain', 3, '❸'], ['alphaUcParenBoth', 27, '(AA)'], ['romanLcParenR', 14, 'xiv)'],
    ];
    for (const [scheme, n, want] of cases) expect([scheme, autoNumberText(scheme, n)]).toEqual([scheme, want]);
  });
});

describe('검토 Minor 4·6·8: 빈 줄 높이, "링크도 글자 색", 테마 색·그림 글머리', () => {
  const toDOM = (node: ReturnType<typeof schema.node>) => schema.nodes.paragraph.spec.toDOM!(node) as [string, Record<string, string>, number];
  test('빈 문단은 문단 끝 글자 크기(endParaRPr sz)로 높이를 잡는다', () => {
    const [, attrs] = toDOM(schema.node('paragraph', { endRPr: '<a:endParaRPr lang="ko-KR" sz="4000"/>' }, []));
    expect(attrs.style).toContain('font-size: calc(round(down, 40 *');
  });
  test('hlinkClr val="tx" 하이퍼링크는 글자 색 그대로(클래스)', () => {
    const raw = '<a:rPr><a:hlinkClick r:id="rId2"><a:extLst><a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}"><ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/></a:ext></a:extLst></a:hlinkClick></a:rPr>';
    expect(schema.marks.runProps.spec.toDOM!(schema.marks.runProps.create({ raw }), true)).toEqual(['span', { class: 'slide-hlink slide-hlink-tx' }, 0]);
  });
  test('문단 직접 글머리 색(schemeClr)은 테마로 풀고, 그림 글머리는 슬라이드 관계로 그림 경로를 찾는다', async () => {
    const { mediaResolverFor } = await import('../src/pptx/theme.js');
    const zip = deck(sp(2, '', xfrm(0, 0, 1, 1), '<a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:buClr><a:schemeClr val="accent2"/></a:buClr><a:buChar char="■"/></a:pPr><a:r><a:t>색</a:t></a:r></a:p><a:p><a:pPr><a:buBlip><a:blip r:embed="rId2"/></a:buBlip></a:pPr><a:r><a:t>그림</a:t></a:r></a:p>'));
    const theme = loadThemeColors(zip, 'slide1.xml');
    const [shape] = parseSlideShapes(new TextDecoder().decode(zip['ppt/slides/slide1.xml']), theme, new Set(), mediaResolverFor(zip, 'ppt/slides/slide1.xml'));
    expect(shape.doc.child(0).attrs.buColor).toBe('#C0504D');
    expect(shape.doc.child(1).attrs.buImage).toBe('ppt/media/image1.png');
    const [, attrs] = toDOM(shape.doc.child(1));
    expect(attrs['data-bu']).toBe('');
    expect(attrs.style).toContain('--p-bu-img: url(');
  });
  test('레이아웃 목록 스타일의 그림 글머리는 레이아웃 관계로 찾는다', () => {
    const zip = deck(
      sp(2, '<p:ph idx="1"/>', '', '<a:bodyPr/><a:p/>'),
      sp(3, '<p:ph idx="1"/>', xfrm(1, 1, 1, 1), '<a:bodyPr/><a:lstStyle><a:lvl1pPr><a:buBlip><a:blip r:embed="rIdPic"/></a:buBlip></a:lvl1pPr></a:lstStyle><a:p/>'),
    );
    zip['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = enc(rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml'], ['rIdPic', 'image', '../media/image1.png']]));
    expect(resolveSlideShapeLooks(zip, 'slide1.xml').get('2')?.levels[0]).toMatchObject({ bullet: '•', bulletImage: 'ppt/media/image1.png' });
  });
});

describe('검토 Minor 9: 뒤집힌 도형의 그라데이션은 뒤집지 않는다(보기 화면 LibreOffice와 같게)', () => {
  test('가로·세로·둘 다 뒤집기 보정, 뒤집기 없거나 방사형이면 그대로', async () => {
    const { unmirroredFill } = await import('../src/client/slideTextLook.js');
    const g = 'linear-gradient(30deg, #000 0%, #fff 100%)';
    expect(unmirroredFill(g, { flipH: true })).toBe('linear-gradient(330deg, #000 0%, #fff 100%)');
    expect(unmirroredFill(g, { flipV: true })).toBe('linear-gradient(150deg, #000 0%, #fff 100%)');
    expect(unmirroredFill(g, { flipH: true, flipV: true })).toBe('linear-gradient(210deg, #000 0%, #fff 100%)');
    expect(unmirroredFill(g, {})).toBe(g);
    expect(unmirroredFill('radial-gradient(#000 0%, #fff 100%)', { flipH: true })).toBe('radial-gradient(#000 0%, #fff 100%)');
  });
});

describe('스마트아트 그림 대체본(dsp:drawing) — 합성 표본(실제 PowerPoint 스마트아트 파일로는 미확인)', () => {
  test('그래픽 틀 → 데이터 파트의 dataModelExt → 그림 파트 도형(틀 기준 좌표·채우기·글자)', () => {
    const gf =
      `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="다이어그램"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="1000" y="2000"/><a:ext cx="9000" cy="3000"/></p:xfrm>` +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rId3" r:lo="rId4" r:qs="rId5" r:cs="rId6"/></a:graphicData></a:graphic></p:graphicFrame>';
    const zip = deck(gf);
    zip['ppt/slides/_rels/slide1.xml.rels'] = enc(rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ['rId3', 'diagramData', '../diagrams/data1.xml'], ['rId7', 'diagramDrawing', '../diagrams/drawing1.xml']]));
    zip['ppt/diagrams/data1.xml'] = enc('<dgm:dataModel xmlns:dgm="d" xmlns:dsp="s"><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt relId="rId7" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>');
    zip['ppt/diagrams/drawing1.xml'] = enc(
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" ${NS}><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>` +
        '<dsp:sp modelId="{A}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4000" cy="3000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></dsp:spPr>' +
        '<dsp:style><a:lnRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:lnRef><a:fillRef idx="1"><a:scrgbClr r="0" g="0" b="0"/></a:fillRef><a:effectRef idx="0"><a:scrgbClr r="0" g="0" b="0"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>' +
        '<dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>계획</a:t></a:r></a:p><a:p><a:r><a:t>1단계</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>',
    );
    const [v] = parseSlideVisuals(zip, 'slide1.xml');
    expect(v).toMatchObject({ kind: 'placeholder', label: '스마트아트', x: 1000, y: 2000, snapshot: true });
    expect(v.parts).toEqual([expect.objectContaining({ kind: 'shape', x: 0, y: 0, w: 4000, h: 3000, fill: '#4F81BD', geom: 'roundRect', text: '계획\n1단계', textColor: '#FFFFFF' })]);
  });
});

describe('2차 검토 R2-I1: 문단 자식 순서는 속성 값 안의 > 나 주석·CDATA에 흔들리지 않는다', () => {
  const order = (para: string) => {
    const xml = `<p:sld ${NS}>${tree(sp(2, '', xfrm(0, 0, 1, 1), `<a:bodyPr/><a:lstStyle/>${para}`))}</p:sld>`;
    const set = new Set<string>();
    const doc = parseSlideShapes(xml, undefined, set)[0].doc;
    const seq: string[] = [];
    doc.firstChild!.forEach((c) => seq.push(c.isText ? c.text! : c.type.name));
    return { seq, set, doc };
  };
  test('r10: tooltip="a>b" — ONE | br | TWO | field | FOUR 순서 그대로, 고쳐 저장해도 순서 유지', () => {
    const P = '<a:p><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId2" tooltip="a>b"/></a:rPr><a:t>ONE</a:t></a:r><a:br/><a:r><a:rPr lang="en-US"/><a:t>TWO</a:t></a:r><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:t>3</a:t></a:fld><a:r><a:rPr lang="en-US"/><a:t>FOUR</a:t></a:r></a:p>';
    const { seq, set, doc } = order(P);
    expect(seq).toEqual(['ONE', 'br', 'TWO', 'field', 'FOUR']);
    const xml = serializeTxBodyXml(doc, set);
    expect(xml.indexOf('ONE')).toBeLessThan(xml.indexOf('<a:br'));
    expect(xml.indexOf('<a:br')).toBeLessThan(xml.indexOf('TWO'));
    expect(xml.indexOf('TWO')).toBeLessThan(xml.indexOf('<a:fld'));
    expect(xml.indexOf('<a:fld')).toBeLessThan(xml.indexOf('FOUR'));
  });
  test('흔들기: 속성 값에 > · 따옴표 · </a:p>처럼 생긴 글자, 주석·CDATA가 섞여도 순서 유지', () => {
    const values = ['a>b', "it's > that", '>>>', '/>', 'x" y', '&gt;&lt;', 'a>/b>c'];
    for (const v of values) {
      const attr = v.includes('"') ? `tooltip='${v}'` : `tooltip="${v}"`;
      const P = `<a:p><!-- 주석 <a:r> > --><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId2" ${attr}/></a:rPr><a:t>A</a:t></a:r><a:br/><a:r><a:rPr lang="en-US" altLang="${v.includes('"') ? 'x' : v}"/><a:t>B</a:t></a:r><a:r><a:t><![CDATA[C<>]]></a:t></a:r></a:p>`;
      expect([v, order(P).seq]).toEqual([v, ['A', 'br', 'B', 'C<>']]);
    }
  });
});

describe('2차 검토 R2-m2: 표 머리 줄 전처리도 압축 폭탄 상한을 지킨다', () => {
  test('적힌 풀린 크기가 상한을 넘으면 풀기 전에 거절(ZipTooLargeError)', async () => {
    const { zipSync } = await import('fflate');
    const { boldPptxTableHeaders } = await import('../src/pptx/loPreprocess.js');
    // 0으로 가득 찬 300MB 슬라이드는 수백 KB로 압축된다 → 파일 크기의 20배·1GB 상한 중 큰 값(1GB)보다 작으니, 상한을 넘기려면
    // 여러 개가 필요하다 — 대신 중앙 목록의 풀린 크기 칸을 크게 고쳐 "적힌 크기" 검사를 확인한다
    const zip = zipSync({ 'ppt/slides/slide1.xml': new TextEncoder().encode('<a:tbl></a:tbl>') });
    const view = new DataView(zip.buffer, zip.byteOffset, zip.byteLength);
    const cd = zip.length - 22 - 46 - 'ppt/slides/slide1.xml'.length; // 중앙 목록 시작(엔트리 1개)
    view.setUint32(cd + 24, 0xfff00000, true); // 풀린 크기 ≈ 4GB
    await expect(boldPptxTableHeaders(zip)).rejects.toThrow(/너무 큽니다/);
  });
});

describe('2차 검토 R2-I2·R2-m1: 글자 속성은 바이트 신뢰 대신 구조 검증(관계는 이 슬라이드의 맞는 종류만)', () => {
  const run = (raw: string, text = '링크') => schema.text(text, [schema.marks.strong.create(), schema.marks.runProps.create({ raw })]);
  const docOf = (...nodes: ReturnType<typeof run>[]) => schema.node('doc', null, [schema.node('paragraph', null, nodes)]);
  test('방이 새로 만들어져 원문이 믿는 목록에 없어도(빈 목록) 하이퍼링크·lang은 남고 굵게만 반영', () => {
    const notices = new Set<string>();
    const xml = serializeTxBodyXml(docOf(run('<a:rPr lang="ko-KR"><a:hlinkClick r:id="rId9"/></a:rPr>')), new Set(), { relTypes: new Map([['rId9', 'hyperlink']]), notices });
    expect(xml).toContain('<a:rPr lang="ko-KR" b="1"><a:hlinkClick r:id="rId9"/></a:rPr>');
    expect([...notices]).toEqual([]);
  });
  test('다른 슬라이드의 링크 rPr를 붙여도 이 슬라이드의 rId2가 그림 관계면 링크를 빼고 알린다', () => {
    const notices = new Set<string>();
    const xml = serializeTxBodyXml(docOf(run('<a:rPr lang="ko-KR"><a:hlinkClick r:id="rId2"/></a:rPr>')), new Set(), { relTypes: new Map([['rId2', 'image']]), notices });
    expect(xml).not.toContain('hlinkClick');
    expect([...notices]).toEqual(['연결 대상이 없는 하이퍼링크를 링크 없이 저장했습니다.']);
    const missing = serializeTxBodyXml(docOf(run('<a:rPr><a:hlinkClick r:id="rId77"/></a:rPr>')), new Set(), { relTypes: new Map(), notices });
    expect(missing).not.toContain('hlinkClick');
  });
  test('허용 안 된 속성 값·요소는 빠진다(구조가 깨지지 않음)', () => {
    const xml = serializeTxBodyXml(docOf(run('<a:rPr lang="ko-KR" sz="12a" foo="1"><p:sp/><a:cs typeface="X"/></a:rPr>')), new Set(), { relTypes: new Map() });
    expect(xml).toContain('<a:rPr lang="ko-KR" b="1"><a:cs typeface="X"/></a:rPr>');
  });
  test('3차 R3-m1: rPr의 이름공간 선언은 extLst가 그 접두어를 쓰므로 남기고, 모르는 속성을 버리면 알린다', () => {
    const notices = new Set<string>();
    const raw = '<a:rPr lang="ko-KR" xmlns:x="urn:custom" x:tag="v"><a:extLst><a:ext uri="{X}"><x:foo/></a:ext></a:extLst></a:rPr>';
    const xml = serializeTxBodyXml(docOf(run(raw)), new Set(), { relTypes: new Map(), notices });
    expect(xml).toContain('<a:rPr lang="ko-KR" xmlns:x="urn:custom" b="1"><a:extLst><a:ext uri="{X}"><x:foo/></a:ext></a:extLst></a:rPr>');
    expect([...notices]).toEqual(['알 수 없는 글자 속성을 빼고 저장했습니다.']);
    const bad = serializeTxBodyXml(docOf(run('<a:rPr xmlns:x="a b"/>')), new Set(), { relTypes: new Map() });
    expect(bad).not.toContain('xmlns:x');
  });
  test('믿는 필드라도 이 슬라이드에 없는 관계를 가리키면 글자로 저장하고 알린다', () => {
    const notices = new Set<string>();
    const raw = '<a:fld id="{X}" type="slidenum"><a:rPr><a:hlinkClick r:id="rId5"/></a:rPr><a:t>3</a:t></a:fld>';
    const xml = serializeTxBodyXml(schema.node('doc', null, [schema.node('paragraph', null, [schema.node('field', { raw, text: '3' })])]), new Set([raw]), { relTypes: new Map(), notices });
    expect(xml).toContain('<a:r><a:t>3</a:t></a:r>');
    expect(notices.size).toBe(1);
  });
});

describe('2차 검토 R2-m3: 레이아웃에서만 맞춤을 물려받은 도형도 다시 맞춘 값을 저장', () => {
  test('bodyPr에 맞춤 요소가 없으면 넣고(100%면 안 넣음), 다른 맞춤 종류면 그대로', async () => {
    const { setShapeAutofit } = await import('../src/pptx/slideContent.js');
    const slide = (bodyPr: string) => `<p:sld ${NS}>${tree(sp(2, '<p:ph idx="1"/>', '', `${bodyPr}<a:lstStyle/><a:p/>`))}</p:sld>`;
    expect(setShapeAutofit(slide('<a:bodyPr/>'), '2', { fontScale: 0.85, lnSpcReduction: 0.2 })).toContain('<a:bodyPr><a:normAutofit fontScale="85000" lnSpcReduction="20000"/></a:bodyPr>');
    expect(setShapeAutofit(slide('<a:bodyPr wrap="square"><a:prstTxWarp prst="textNoShape"/><a:scene3d/></a:bodyPr>'), '2', { fontScale: 0.85, lnSpcReduction: 0 })).toContain('<a:prstTxWarp prst="textNoShape"/><a:normAutofit fontScale="85000"/><a:scene3d/>');
    expect(setShapeAutofit(slide('<a:bodyPr/>'), '2', { fontScale: 1, lnSpcReduction: 0 })).toBe(slide('<a:bodyPr/>'));
    expect(setShapeAutofit(slide('<a:bodyPr><a:noAutofit/></a:bodyPr>'), '2', { fontScale: 0.5, lnSpcReduction: 0.2 })).toBe(slide('<a:bodyPr><a:noAutofit/></a:bodyPr>'));
  });
});

describe('2차 검토 R2-m4: 감싼 도형은 Choice·Fallback 자리를 합친 영역을 잘라 온다', () => {
  test('두 자리가 다르면 합친 상자', () => {
    const box = (x: number) => `<p:sp><p:nvSpPr><p:cNvPr id="3" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(x, 0, 100000, 50000)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;
    const zip = deck(`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="a14">${box(500000)}</mc:Choice><mc:Fallback>${box(400000)}</mc:Fallback></mc:AlternateContent>`);
    expect(parseSlideVisuals(zip, 'slide1.xml')).toEqual([{ kind: 'snapshot', x: 400000, y: 0, w: 200000, h: 50000, snapshot: true }]);
  });
});

describe('RX-C1 저장 직전 안전망: 같은 원본 슬라이드 항목이 둘이면 하나만 저장(sldId 중복 손상 방지)', () => {
  test('합쳐진 슬라이드 목록(원본 3장 × 2)을 내보내도 sldId·슬라이드는 3개, 알림 1개', async () => {
    const Y = await import('yjs');
    const { exportPptxFromSlides, seedSlidesFromPptx } = await import('../src/pptx/collab.js');
    const { makeSamplePptx } = await import('./helpers.js');
    const { PPTX_SLIDES_ARRAY } = await import('../src/collab/pptxShared.js');
    const original = makeSamplePptx(['A', 'B', 'C']);
    const doc = new Y.Doc();
    const slides = doc.getArray(PPTX_SLIDES_ARRAY) as never as import('yjs').Array<import('../src/collab/pptxShared.js').SlideEntry>;
    seedSlidesFromPptx(original, slides as never);
    const again = new Y.Doc().getArray(PPTX_SLIDES_ARRAY) as never as import('yjs').Array<import('../src/collab/pptxShared.js').SlideEntry>;
    seedSlidesFromPptx(original, again as never);
    slides.push(again.toArray().map((e) => ({ ...e })));
    expect(slides.length).toBe(6);
    const out = exportPptxFromSlides(original, slides as never);
    const { unzipSync, strFromU8 } = await import('fflate');
    const pres = strFromU8(unzipSync(out.bytes)['ppt/presentation.xml']);
    const ids = [...pres.matchAll(/<p:sldId\b[^>]*\bid="(\d+)"/g)].map((m) => m[1]);
    expect(ids.length).toBe(3);
    expect(new Set(ids).size).toBe(3);
    expect(out.finalParts.length).toBe(3);
    expect(out.notices).toEqual(['같은 슬라이드가 두 번 들어 있어(연결이 끊겼다 이어질 때 생길 수 있음) 한 번만 저장했습니다.']);
  });
});

describe('자간(rPr spc) — 원본 글자 간격대로 그려 PowerPoint·아트크래프트보다 글이 넓어져 줄이 달리 끊기던 것', () => {
  const toDOM = (raw: string) => schema.marks.runProps.spec.toDOM!(schema.marks.runProps.create({ raw }), true);
  test('spc="-80"(1/100pt) → letter-spacing -0.8pt(슬라이드 축척)', () => {
    expect(toDOM('<a:rPr lang="ko-KR" sz="2000" b="1" spc="-80" dirty="0"/>')).toEqual(['span', { style: 'letter-spacing: calc(-0.8 * var(--slide-pt, 1pt))' }, 0]);
  });
  test('spc 없음·0이면 그대로, 하이퍼링크와 함께면 둘 다', () => {
    expect(toDOM('<a:rPr lang="ko-KR" spc="0"/>')).toEqual(['span', 0]);
    expect(toDOM('<a:rPr lang="ko-KR"/>')).toEqual(['span', 0]);
    expect(toDOM('<a:rPr spc="150"><a:hlinkClick r:id="rId2"/></a:rPr>')).toEqual(['span', { class: 'slide-hlink', style: 'letter-spacing: calc(1.5 * var(--slide-pt, 1pt))' }, 0]);
  });
  test('하위 요소(<a:hlinkClick spc=…> 같은 것)의 속성은 자간으로 읽지 않는다', () => {
    expect(toDOM('<a:rPr lang="ko-KR"><a:latin typeface="x" spc="-80"/></a:rPr>')).toEqual(['span', 0]);
  });
  test('글자에 직접 b="0"이면 수준 굵게를 끄는 slide-b0(b="1"·없음은 그대로)', () => {
    expect(toDOM('<a:rPr lang="en-US" sz="1000" b="0" dirty="0"/>')).toEqual(['span', { class: 'slide-b0' }, 0]);
    expect(toDOM('<a:rPr lang="en-US" b="1"/>')).toEqual(['span', 0]);
    expect(toDOM('<a:rPr b="0" spc="-50"><a:hlinkClick r:id="rId2"/></a:rPr>')).toEqual(['span', { class: 'slide-hlink slide-b0', style: 'letter-spacing: calc(-0.5 * var(--slide-pt, 1pt))' }, 0]);
  });
});
