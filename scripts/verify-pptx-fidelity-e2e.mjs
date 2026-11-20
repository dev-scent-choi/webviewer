// PPT 편집 화면 표시 충실도 회귀 검사 — 편집기(editor-slides.html)를 실제 브라우저로 열어
// 글상자 안쪽 여백(bodyPr lIns/tIns/rIns/bIns), 둥근 사각형 반지름(adj 0 = 직각), 점선 모양(sysDash 등)을 잰다.
// 운영에 반영하기 전에 verify:xlsx와 함께 돌린다.
//
// 사용법(서버를 먼저 띄워 둘 것 — 업로드·토큰 API가 키 없이 열려 있는 로컬 서버):
//   node scripts/verify-pptx-fidelity-e2e.mjs [http://127.0.0.1:8124] [스크린샷 폴더]
// 시험 파일은 매번 최소 구성 pptx를 직접 만든다(고객 문서를 쓰지 않는다).
import JSZip from 'jszip';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-fidelity-'));
fs.mkdirSync(shotDir, { recursive: true });

const results = [];
const rec = (name, ok, detail = '') => {
  results.push({ name, ok, detail });
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};

const MC_NS = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"';
const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (items) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${items
    .map(([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`)
    .join('')}</Relationships>`;
const tree = (inner) =>
  `<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${inner}</p:spTree></p:cSld>`;
const xfrm = (x, y, w, h) => `<a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>`;
const text = (t) => `<a:p><a:r><a:rPr lang="ko-KR" sz="2000"/><a:t>${t}</a:t></a:r></a:p>`;

// 1인치 = 914400 EMU, 슬라이드 10x7.5인치
const SHAPES =
  // id 2: 여백 0 글상자(왼쪽·위 붙음)
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="여백없음"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(914400, 457200, 4572000, 685800)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
  `<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"/><a:lstStyle/>${text('Inset zero')}</p:txBody></p:sp>` +
  // id 3: 왼쪽만 0(나머지 기본)
  `<p:sp><p:nvSpPr><p:cNvPr id="3" name="왼쪽0"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(914400, 1371600, 4572000, 685800)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
  `<p:txBody><a:bodyPr wrap="square" lIns="0"/><a:lstStyle/>${text('Left zero')}</p:txBody></p:sp>` +
  // id 4: 빨간 sysDash 테두리 사각형(채우기 없음)
  `<p:sp><p:nvSpPr><p:cNvPr id="4" name="점선"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(914400, 2743200, 2743200, 1371600)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>` +
  `<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></p:spPr></p:sp>` +
  // id 5: adj 0 둥근 사각형(= 직각)
  `<p:sp><p:nvSpPr><p:cNvPr id="5" name="직각"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(4572000, 2743200, 2743200, 914400)}<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 0"/></a:avLst></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="0070C0"/></a:solidFill></p:spPr></p:sp>` +
  // id 6: 기본 둥근 사각형(짧은 변의 1/6 반지름)
  `<p:sp><p:nvSpPr><p:cNvPr id="6" name="둥근"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(4572000, 4114800, 2743200, 914400)}<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="00B050"/></a:solidFill></p:spPr></p:sp>` +
  // id 7: 하이퍼링크·문단 안 줄 바꿈(a:br)·자동 번호 2개·어절 단위 줄 바꿈(eaLnBrk=0) — 품질 점검 말뭉치에 없는 기능
  `<p:sp><p:nvSpPr><p:cNvPr id="7" name="목록"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(914400, 5257800, 3657600, 1371600)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
  '<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>' +
  '<a:p><a:pPr eaLnBrk="0"/><a:r><a:rPr lang="ko-KR" sz="1400"><a:hlinkClick r:id="rId9"/></a:rPr><a:t>링크글자</a:t></a:r><a:br><a:rPr lang="ko-KR" sz="1400"/></a:br><a:r><a:rPr lang="ko-KR" sz="1400"/><a:t>둘째줄</a:t></a:r></a:p>' +
  '<a:p><a:pPr marL="342900" indent="-342900"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:rPr lang="ko-KR" sz="1400"/><a:t>첫 항목</a:t></a:r></a:p>' +
  '<a:p><a:pPr marL="342900" indent="-342900"><a:buAutoNum type="arabicPeriod"/></a:pPr><a:r><a:rPr lang="ko-KR" sz="1400"/><a:t>둘째 항목</a:t></a:r></a:p></p:txBody></p:sp>' +
  // id 8: 세로 쓰기(eaVert)
  `<p:sp><p:nvSpPr><p:cNvPr id="8" name="세로"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(7772400, 457200, 914400, 2743200)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
  `<p:txBody><a:bodyPr vert="eaVert"/><a:lstStyle/>${text('세로글자')}</p:txBody></p:sp>` +
  // id 9: 30° 돌린 오른쪽 화살표 도형(글자도 같이 돈다)
  `<p:sp><p:nvSpPr><p:cNvPr id="9" name="화살표"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(4572000, 5257800, 2743200, 914400).replace('<a:xfrm>', '<a:xfrm rot="1800000">')}<a:prstGeom prst="rightArrow"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/>${text('돌린 글')}</p:txBody></p:sp>` +
  // id 10: 문단 안 수식(mc:AlternateContent/a14:m) + 숫자 문자 참조 글자·글꼴 이름(검토 I3·I4)
  `<p:sp><p:nvSpPr><p:cNvPr id="10" name="수식"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(457200, 3886200, 3200400, 685800)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
  '<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" sz="1400"><a:latin typeface="Evil&#10;Font"/></a:rPr><a:t>&#xAC00;&#xB098;&#xB2E4; E=</a:t></a:r>' +
  `<mc:AlternateContent><mc:Choice Requires="a14"><a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath><m:r><m:t>mc2</m:t></m:r></m:oMath></m:oMathPara></a14:m></mc:Choice><mc:Fallback><a:r><a:rPr lang="en-US" sz="1400"/><a:t>mc²</a:t></a:r></mc:Fallback></mc:AlternateContent></a:p></p:txBody></p:sp>` +
  // id 11: 슬라이드 최상위 mc:AlternateContent로 감싼 도형(검토 I5) — Fallback의 빨간 상자
  `<mc:AlternateContent><mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="11" name="감싼 도형"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(7315200, 3886200, 914400, 457200)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp></mc:Choice>` +
  `<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="11" name="감싼 도형"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(7315200, 3886200, 914400, 457200)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp></mc:Fallback></mc:AlternateContent>` +
  // id 12: 넘치면 글자 줄이기(normAutofit) 작은 상자 — 글을 늘리면 다시 맞춰 작아진다(검토 Minor 3)
  `<p:sp><p:nvSpPr><p:cNvPr id="12" name="맞춤"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(457200, 4800600, 2286000, 457200)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
  `<p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/>${text('맞춤')}</p:txBody></p:sp>`;

const THEME =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme ${NS} name="T"><a:themeElements>` +
  '<a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>' +
  '<a:fontScheme name="T"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
  '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
  '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
  '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>' +
  '</a:themeElements></a:theme>';

async function buildFixture() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
      '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' +
      '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>',
  );
  zip.file('_rels/.rels', rels([['rId1', 'officeDocument', 'ppt/presentation.xml']]));
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
      '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000" type="screen4x3"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
  );
  zip.file('ppt/_rels/presentation.xml.rels', rels([['rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'], ['rId2', 'slide', 'slides/slide1.xml'], ['rId3', 'theme', 'theme/theme1.xml']]));
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster ${NS}>${tree('')}<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>',
  );
  zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ['rId2', 'theme', '../theme/theme1.xml']]));
  zip.file('ppt/slideLayouts/slideLayout1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout ${NS} type="blank">${tree('')}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`);
  zip.file('ppt/slideLayouts/_rels/slideLayout1.xml.rels', rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']]));
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS} ${MC_NS}>${tree(SHAPES)}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  zip.file(
    'ppt/slides/_rels/slide1.xml.rels',
    rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml']]).replace(
      '</Relationships>',
      `<Relationship Id="rId9" Type="${REL}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`,
    ),
  );
  zip.file('ppt/theme/theme1.xml', THEME);
  return zip.generateAsync({ type: 'nodebuffer' });
}

async function upload(bytes) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), 'fidelity.pptx');
  const id = (await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json()).data.id;
  for (;;) {
    const st = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data;
    if (st.status === 'ready') break;
    if (st.status === 'failed') throw new Error('변환 실패');
    await new Promise((r) => setTimeout(r, 400));
  }
  const token = (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: 'fidelity', canPrint: true, canDownload: true, canCopy: true, canEdit: true }) })).json()).data;
  return { id, token };
}

const { id, token } = await upload(await buildFixture());
const browser = await chromium.launch();
const errors = [];
const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
page.on('pageerror', (e) => errors.push(String(e)));
await page.goto(BASE + token.editorUrl);
await page.waitForSelector('#slide-canvas .slide-textbox[data-shape-id="2"]', { timeout: 60000 });
await page.waitForFunction(() => document.querySelectorAll('#slide-canvas .slide-visual-shape').length >= 3, null, { timeout: 30000 }).catch(() => {});
await page.waitForTimeout(800);

const m = await page.evaluate(() => {
  const pad = (id) => {
    const el = document.querySelector(`#slide-canvas .slide-textbox[data-shape-id="${id}"]`);
    if (!el) return null;
    const cs = getComputedStyle(el);
    return { l: parseFloat(cs.paddingLeft), t: parseFloat(cs.paddingTop), r: parseFloat(cs.paddingRight), b: parseFloat(cs.paddingBottom) };
  };
  const pt = parseFloat(getComputedStyle(document.getElementById('slide-canvas')).getPropertyValue('--slide-pt'));
  const shapes = [...document.querySelectorAll('#slide-canvas .slide-visual-shape')];
  const dashed = [...document.querySelectorAll('#slide-canvas .slide-visual-stroke path')].map((p) => ({ dash: p.getAttribute('stroke-dasharray'), d: p.getAttribute('d') }));
  const radius = (color) => {
    const el = shapes.find((s) => getComputedStyle(s).backgroundColor === color);
    return el ? getComputedStyle(el).borderTopLeftRadius : null;
  };
  return { pt, p2: pad('2'), p3: pad('3'), dashed, sharp: radius('rgb(0, 112, 192)'), round: radius('rgb(0, 176, 80)') };
});
const near = (a, b, tol = 0.6) => typeof a === 'number' && Math.abs(a - b) <= tol;
rec('글상자 여백 0(lIns·tIns·rIns·bIns=0)은 네 방향 모두 0', m.p2 && m.p2.l === 0 && m.p2.t === 0 && m.p2.r === 0 && m.p2.b === 0, JSON.stringify(m.p2));
rec(
  '왼쪽만 0인 글상자 — 왼 0, 위·아래 3.6pt, 오른 7.2pt(방향이 뒤바뀌지 않음)',
  m.p3 && m.p3.l === 0 && near(m.p3.t, 3.6 * m.pt) && near(m.p3.b, 3.6 * m.pt) && near(m.p3.r, 7.2 * m.pt),
  `${JSON.stringify(m.p3)} pt=${m.pt}`,
);
rec('sysDash 테두리는 SVG 점선(두께×3, 두께×1)', m.dashed.length === 1 && /^(\d+(\.\d+)?) (\d+(\.\d+)?)$/.test(m.dashed[0].dash ?? '') && (() => { const [a, b] = m.dashed[0].dash.split(' ').map(Number); return near(a, b * 3, 0.01); })(), JSON.stringify(m.dashed));
rec('점선 사각형 모서리는 직각(호 없음)', m.dashed.length === 1 && !/A/.test(m.dashed[0].d), m.dashed[0]?.d);
rec('adj 0 둥근 사각형은 직각', m.sharp === '0px', String(m.sharp));
rec('기본 둥근 사각형은 원형 반지름(짧은 변의 1/6, 가로·세로 같음)', /^[\d.]+px$/.test(m.round ?? '') && parseFloat(m.round) > 2, String(m.round));
// 왼쪽 슬라이드 창 — PowerPoint처럼 각진 썸네일, 번호는 썸네일 왼쪽, 선택은 굵은(3px 이상) 테두리
const rail = await page.evaluate(() => {
  const card = document.querySelector('#slides-rail .slide-card.active');
  const thumb = card?.querySelector('.slide-card-thumb');
  const num = card?.querySelector('.slide-card-number');
  if (!thumb || !num) return null;
  const cs = getComputedStyle(thumb);
  return { radius: cs.borderTopLeftRadius, outline: parseFloat(cs.outlineWidth), numLeft: num.getBoundingClientRect().right <= thumb.getBoundingClientRect().left, bold: Number(getComputedStyle(num).fontWeight) >= 600 };
});
rec('슬라이드 창: 각진 썸네일·번호 왼쪽·선택 굵은 테두리', rail && rail.radius === '0px' && rail.numLeft && rail.outline >= 3 && rail.bold, JSON.stringify(rail));
// 미리보기 그림이 실제로 뜨고(작은 그림 요청이 거부돼 빈 칸으로 남은 적이 있다), 본문용보다 작은지
await page.waitForFunction(() => [...document.querySelectorAll('#slides-rail .slide-card-thumb img')].every((i) => i.complete), null, { timeout: 30000 }).catch(() => {});
const thumb = await page.evaluate(async () => {
  const img = document.querySelector('#slides-rail .slide-card-thumb img');
  if (!img) return null;
  const res = await fetch(img.src);
  return { loaded: img.naturalWidth > 0, small: img.src.includes('thumb=true'), status: res.status, width: img.naturalWidth };
});
rec('슬라이드 창 미리보기가 작은 그림으로 뜬다', thumb && thumb.loaded && thumb.small && thumb.status === 200 && thumb.width <= 600, JSON.stringify(thumb));
// 하이퍼링크·줄 바꿈·자동 번호·어절 줄 바꿈·세로 쓰기·회전(품질 점검 P3·P4·P6 보강 — 말뭉치에 없는 기능)
const extra = await page.evaluate(() => {
  const box = (id) => document.querySelector(`#slide-canvas .slide-textbox[data-shape-id="${id}"]`);
  const ps = [...(box('7')?.querySelectorAll('p') ?? [])];
  const link = box('7')?.querySelector('.slide-hlink');
  return {
    link: link ? { color: getComputedStyle(link).color, underline: getComputedStyle(link).textDecorationLine } : null,
    br: box('7')?.querySelectorAll('p br:not(.ProseMirror-trailingBreak)').length ?? 0,
    nums: ps.slice(1).map((p) => getComputedStyle(p, '::before').content),
    keep: ps[0] ? getComputedStyle(ps[0]).wordBreak : null,
    vert: box('8') ? getComputedStyle(box('8').querySelector('.ProseMirror')).writingMode : null,
    turn: box('9') ? getComputedStyle(box('9')).transform : null,
  };
});
rec('하이퍼링크 글자: 테마 하이퍼링크 색 + 밑줄', extra.link?.color === 'rgb(0, 0, 255)' && extra.link.underline.includes('underline'), JSON.stringify(extra.link));
rec('문단 안 줄 바꿈(a:br)이 보인다', extra.br === 1, String(extra.br));
rec('자동 번호 1. 2.', JSON.stringify(extra.nums) === JSON.stringify(['"1."', '"2."']), JSON.stringify(extra.nums));
rec('eaLnBrk=0 문단은 어절 단위 줄 바꿈(keep-all)', extra.keep === 'keep-all', String(extra.keep));
rec('세로 쓰기(eaVert)', extra.vert === 'vertical-rl', String(extra.vert));
rec('30° 돌린 도형의 글상자도 30° 회전', (() => { const m = /matrix\(([^,]+), ([^,]+)/.exec(extra.turn ?? ''); return m ? Math.abs((Math.atan2(Number(m[2]), Number(m[1])) * 180) / Math.PI - 30) < 0.5 : false; })(), String(extra.turn));
// 저장 보존: 링크 글자 뒤에 타이핑하고 저장 → 내려받은 파일에 하이퍼링크·줄 바꿈·자동 번호·eaLnBrk·lang이 남는다
await page.click('#slide-canvas .slide-textbox[data-shape-id="7"] .slide-hlink');
await page.keyboard.press('End');
await page.keyboard.type('추가');
await page.waitForTimeout(1500);
await fetch(`${BASE}/api/view/save?token=${encodeURIComponent(token.token)}`, { method: 'POST' });
const saved = await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(token.token)}`)).arrayBuffer();
const savedSlide = await (await JSZip.loadAsync(saved)).file('ppt/slides/slide1.xml').async('string');
const sp7 = /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?id="7"[\s\S]*?<\/p:sp>/.exec(savedSlide)?.[0] ?? '';
rec(
  '고친 글상자 저장: 글자 반영 + 하이퍼링크·a:br·buAutoNum·eaLnBrk·lang 유지',
  /둘째줄추가|링크글자추가/.test(sp7) && sp7.includes('<a:hlinkClick r:id="rId9"/>') && /<a:br>/.test(sp7) && (sp7.match(/<a:buAutoNum type="arabicPeriod"\/>/g) ?? []).length === 2 && sp7.includes('eaLnBrk="0"') && sp7.includes('lang="ko-KR"'),
  sp7.slice(0, 400),
);
// 검토 I3·I4·I5·Minor 2·3·I2 — 실제 편집기에서 고치고 저장한 파일을 본다
const box = (id) => `#slide-canvas .slide-textbox[data-shape-id="${id}"]`;
const shown = await page.evaluate(() => ({
  text10: document.querySelector('#slide-canvas .slide-textbox[data-shape-id="10"]')?.textContent ?? '',
  snapshots: [...document.querySelectorAll('#slide-canvas .slide-visual-snapshot')].map((el) => [parseFloat(el.style.left), parseFloat(el.style.top)]),
}));
rec('숫자 문자 참조는 글자로, 수식은 대체 글자로 그 자리에 보인다', shown.text10.includes('가나다 E=') && shown.text10.includes('mc²') && !shown.text10.includes('&#'), shown.text10);
// 감싼 도형 자리(8인치, 4.25인치 = 576px·306px 근처 — 캔버스 960px, 10인치 슬라이드)에 원본 렌더링 조각
rec('mc:AlternateContent로 감싼 도형 자리에 원본 렌더링 조각', shown.snapshots.some(([x, y]) => Math.abs(x - 768) < 12 && Math.abs(y - 408) < 12), JSON.stringify(shown.snapshots));
// 수식 상자: 첫 글자 뒤에 입력 + Shift+Enter 줄 바꿈
const t10 = await page.locator(`${box('10')} p`).first().boundingBox();
await page.mouse.click(t10.x + 4, t10.y + t10.height / 2);
await page.keyboard.press('Home');
await page.keyboard.press('ArrowRight');
await page.keyboard.type('!');
await page.keyboard.press('Shift+Enter');
await page.keyboard.type('줄');
// 자동 맞춤 상자: 글을 늘리면 다시 맞춰 글자가 작아진다
const before12 = await page.locator(`${box('12')} .ProseMirror span, ${box('12')} .ProseMirror p`).first().evaluate((el) => parseFloat(getComputedStyle(el).fontSize));
await page.click(`${box('12')} .ProseMirror`);
await page.keyboard.press('End');
for (let i = 0; i < 4; i += 1) { await page.keyboard.press('Enter'); await page.keyboard.type('더 많은 글'); }
await page.waitForTimeout(600);
const after12 = await page.locator(`${box('12')} .ProseMirror span, ${box('12')} .ProseMirror p`).first().evaluate((el) => parseFloat(getComputedStyle(el).fontSize));
rec('자동 맞춤 상자: 글을 늘리면 글자가 작아진다', after12 < before12 * 0.95, `${before12} → ${after12}`);
await page.waitForTimeout(1200);
// 협업자 위조(검토 I2): 별도 Yjs 클라이언트가 글상자 7 첫 문단의 끝 속성에 다른 도형을 끼워 넣으려 한다
{
  const Y = await import('yjs');
  const { WebsocketProvider } = await import('y-websocket');
  const { default: WebSocket } = await import('ws');
  const ydoc = new Y.Doc();
  const prov = new WebsocketProvider(`${BASE.replace('http', 'ws')}/ws/edit`, id, ydoc, { params: { token: token.token }, WebSocketPolyfill: WebSocket });
  await new Promise((resolve) => (prov.synced ? resolve() : prov.once('synced', resolve)));
  const contents = ydoc.getMap('slideContents');
  for (const per of contents.values()) {
    const frag = per.get('7')?.get('content');
    const para = frag?.get(0);
    if (para) ydoc.transact(() => para.setAttribute('endRPr', '</a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="999" name="injected"/>'));
  }
  await new Promise((r) => setTimeout(r, 800));
  prov.destroy();
}
await page.waitForTimeout(1500);
await fetch(`${BASE}/api/view/save?token=${encodeURIComponent(token.token)}`, { method: 'POST' });
const saved2 = await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(token.token)}`)).arrayBuffer();
const slide2 = await (await JSZip.loadAsync(saved2)).file('ppt/slides/slide1.xml').async('string');
const shapeXml = (n) => new RegExp(`<p:sp>(?:(?!</p:sp>)[\\s\\S])*?<p:cNvPr id="${n}"[\\s\\S]*?</p:sp>`).exec(slide2)?.[0] ?? '';
const sp10 = shapeXml(10);
rec('수식 상자 고쳐 저장: 수식 원문·대체 글자 유지, 글자·글꼴 이름 바르게(&amp;# 없음), Shift+Enter = a:br', sp10.includes('<a14:m>') && sp10.includes('<a:t>mc²</a:t>') && sp10.includes('typeface="Evil&#10;Font"') && !sp10.includes('&amp;#') && sp10.includes('<a:t>가!</a:t></a:r><a:br/>') && sp10.includes('<a:t>줄나다 E=</a:t>'), sp10.slice(0, 600));
const fit12 = /<a:normAutofit fontScale="(\d+)"/.exec(shapeXml(12))?.[1];
rec('자동 맞춤 상자 저장: 새 fontScale(100% 미만)', Boolean(fit12) && Number(fit12) < 100000, String(fit12));
const { XMLValidator } = await import('fast-xml-parser');
rec('위조한 끝 속성이 저장 XML에 들어가지 않고 슬라이드 XML이 올바르다', !slide2.includes('injected') && XMLValidator.validate(slide2) === true, slide2.includes('injected') ? 'injected 있음' : String(XMLValidator.validate(slide2)));
rec('감싼 도형(mc:AlternateContent)은 원문 그대로 저장', slide2.includes('<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="11"'), '');
await page.screenshot({ path: path.join(shotDir, 'pptx-editor.png') });
rec('페이지 오류 없음', errors.length === 0, errors.slice(0, 2).join(' | '));
await browser.close();
await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' }).catch(() => {});
const fails = results.filter((r) => !r.ok);
console.log(`\n${fails.length === 0 ? 'ALL CHECKS PASSED' : `${fails.length} FAILED`} (${results.length - fails.length}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(fails.length === 0 ? 0 : 1);
