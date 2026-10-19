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
  `<a:solidFill><a:srgbClr val="00B050"/></a:solidFill></p:spPr></p:sp>`;

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
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NS}>${tree(SHAPES)}<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  zip.file('ppt/slides/_rels/slide1.xml.rels', rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml']]));
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
await page.screenshot({ path: path.join(shotDir, 'pptx-editor.png') });
rec('페이지 오류 없음', errors.length === 0, errors.slice(0, 2).join(' | '));
await browser.close();
await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' }).catch(() => {});
const fails = results.filter((r) => !r.ok);
console.log(`\n${fails.length === 0 ? 'ALL CHECKS PASSED' : `${fails.length} FAILED`} (${results.length - fails.length}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(fails.length === 0 ? 0 : 1);
