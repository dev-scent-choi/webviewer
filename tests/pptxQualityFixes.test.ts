import { describe, expect, test } from 'vitest';
import { loadThemeColors } from '../src/pptx/theme.js';
import { parseSlideShapes, parseSlideVisuals, serializeTxBodyXml, spliceShapeTxBody } from '../src/pptx/slideContent.js';
import { resolveSlideShapeLooks } from '../src/pptx/inheritance.js';
import { autoNumberPlan, autoNumberText, pPrLook } from '../src/pptx/paragraphLook.js';
import { pptxTextSchema as schema } from '../src/pptx/textSchema.js';

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
  test('normAutofit fontScale/lnSpcReduction을 비율로, 레이아웃에서도 물려받고, spAutoFit은 맞춤 없음', () => {
    const zip = deck(
      sp(2, '<p:ph idx="1"/>', '', '<a:bodyPr><a:normAutofit fontScale="47500" lnSpcReduction="20000"/></a:bodyPr><a:p/>') +
        sp(3, '<p:ph type="body" idx="2"/>', '', '<a:bodyPr/><a:p/>') +
        sp(4, '', xfrm(0, 0, 10, 10), '<a:bodyPr vert="eaVert"><a:spAutoFit/></a:bodyPr><a:p/>'),
      sp(3, '<p:ph type="body" idx="2"/>', xfrm(1, 1, 1, 1), '<a:bodyPr><a:normAutofit fontScale="90000"/></a:bodyPr><a:p/>'),
    );
    const looks = resolveSlideShapeLooks(zip, 'slide1.xml');
    expect(looks.get('2')?.autofit).toEqual({ fontScale: 0.475, lnSpcReduction: 0.2 });
    expect(looks.get('3')?.autofit).toEqual({ fontScale: 0.9, lnSpcReduction: 0 });
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
      bulletColor: '#FF0000',
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
    const [shape] = parseSlideShapes(slide);
    const para = shape.doc.child(0);
    expect(para.textContent).toBe('링크둘째 줄 ');
    expect(para.child(3).attrs.text).toBe('3');
    expect([...Array(para.childCount).keys()].map((i) => para.child(i).type.name)).toEqual(['text', 'br', 'text', 'field']);
    const xml = serializeTxBodyXml(shape.doc);
    expect(xml).toContain('<a:rPr lang="ko-KR" b="1" spc="-50"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="Georgia"/><a:hlinkClick r:id="rId9"/></a:rPr>');
    expect(xml).toContain('<a:br><a:rPr lang="ko-KR"/></a:br>');
    expect(xml).toContain('<a:fld id="{X}" type="slidenum"><a:rPr lang="en-US"/><a:t>3</a:t></a:fld>');
    expect(xml).toContain('<a:endParaRPr lang="ko-KR" sz="1200"/></a:p>');
    expect(xml.indexOf('<a:br>')).toBeLessThan(xml.indexOf('둘째'));
  });
  test('굵게를 끄고 글자를 고쳐도 하이퍼링크·lang·자간·테마 색은 남고 b만 빠진다', () => {
    const [shape] = parseSlideShapes(slide);
    const first = shape.doc.child(0).child(0);
    const marks = first.marks.filter((m) => m.type.name !== 'strong');
    const para = shape.doc.child(0).copy(shape.doc.child(0).content.replaceChild(0, schema.text('새 링크', marks)));
    const xml = serializeTxBodyXml(schema.node('doc', null, [para]));
    expect(xml).toContain('<a:r><a:rPr lang="ko-KR" spc="-50"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="Georgia"/><a:hlinkClick r:id="rId9"/></a:rPr><a:t>새 링크</a:t></a:r>');
    const spliced = spliceShapeTxBody(slide, '2', xml);
    expect(parseSlideShapes(spliced)[0].doc.child(0).textContent).toBe('새 링크둘째 줄 ');
    expect(spliced).toContain('<a:fld id="{X}" type="slidenum">');
  });
  test('글꼴을 바꾸면 latin·ea를 새 이름으로 쓴다(나머지는 유지)', () => {
    const [shape] = parseSlideShapes(slide);
    const first = shape.doc.child(0).child(0);
    const marks = first.marks.filter((m) => m.type.name !== 'fontFamily').concat(schema.marks.fontFamily.create({ family: '바탕' }));
    const xml = serializeTxBodyXml(schema.node('doc', null, [schema.node('paragraph', null, [schema.text('x', marks)])]));
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
    const styles = '<a:tblStyleLst><a:tblStyle styleId="{AAA}" styleName="x"><a:firstRow><a:tcTxStyle b="off"/></a:firstRow></a:tblStyle></a:tblStyleLst>';
    expect(boldTableHeaderXml(tbl('{AAA}'), styles)).toBe(tbl('{AAA}'));
  });
});

describe('슬라이드 배경 bgRef → 테마 배경 채우기 목록(bgFillStyleLst)', () => {
  test('idx 1001 = 첫 배경 채우기(phClr = bgRef 색)', () => {
    const zip = deck('');
    zip['ppt/slides/slide1.xml'] = enc(`<p:sld ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="accent2"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>`);
    expect(parseSlideVisuals(zip, 'slide1.xml')[0]).toMatchObject({ kind: 'shape', background: true, fill: '#C0504D' });
  });
});
