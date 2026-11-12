// 통합 보기 껍데기(view.html) 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-view-shell-e2e.mjs [BASE] [스크린샷 폴더]
// 워드·PPT·PDF × 열람만(서버 PNG)·내려받기(pdf.js)·편집 토큰. 열람만은 PDF가 브라우저로 가지 않는다(derived-pdf·pdf-source 0건).
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { strToU8, zipSync } from 'fflate';
import { PDFDocument, StandardFonts } from 'pdf-lib';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'view-shell-'));
fs.mkdirSync(shotDir, { recursive: true });
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};
const wait = (ms) => new Promise((r) => setTimeout(r, ms));

// ---------- 시험 문서(매번 만든다) ----------
async function makePdf() {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  for (let i = 1; i <= 3; i += 1) doc.addPage([612, 792]).drawText(`ViewShell Target page ${i}`, { x: 60, y: 760, size: 18, font });
  return doc.save();
}

function makeDocx() {
  const files = {
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
    '_rels/.rels': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
    'word/document.xml': strToU8(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>보기 껍데기 시험 문서</w:t></w:r></w:p><w:p><w:r><w:t>두 번째 문단</w:t></w:r></w:p></w:body></w:document>',
    ),
  };
  return zipSync(files);
}

async function makeXlsx() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('보이는 시트');
  ws.getCell('A1').value = '품목';
  ws.getCell('B1').value = '단가';
  ws.getCell('A2').value = '사과';
  ws.getCell('B2').value = 1200;
  return new Uint8Array(await wb.xlsx.writeBuffer());
}

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
const SHAPES = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="글"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(914400, 914400, 6400800, 1828800)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>${text('보기 껍데기 슬라이드')}</p:txBody></p:sp>`;
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

const uploaded = [];
async function upload(bytes, name, { wait: waitReady = true } = {}) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
  const r = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  if (!r.data) throw new Error(`업로드 실패: ${JSON.stringify(r)}`);
  uploaded.push(r.data.id);
  if (!waitReady) return r.data.id;
  for (let i = 0; i < 400; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${r.data.id}`)).json()).data;
    if (st.status === 'ready') return r.data.id;
    if (st.status === 'failed' || st.status === 'error') throw new Error(`변환 실패: ${JSON.stringify(st)}`);
    await wait(250);
  }
  throw new Error('대기 시간 초과');
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '보기시험', canPrint: false, canDownload: false, canCopy: false, canEdit: false, ...body }),
  })
    .then((r) => r.json())
    .then((j) => j.data);

// ---------- 브라우저 도구 ----------
const browser = await chromium.launch();
async function openView(url, { viewport = { width: 1400, height: 900 }, setup, ctxOpts = {} } = {}) {
  const ctx = await browser.newContext({ viewport, ...ctxOpts });
  const page = await ctx.newPage();
  const errors = [];
  const requests = [];
  page.on('pageerror', (e) => errors.push(String(e)));
  page.on('console', (m) => m.type() === 'error' && errors.push(m.text()));
  page.on('request', (r) => requests.push(new URL(r.url()).pathname));
  if (setup) await setup(page);
  await page.goto(BASE + url);
  await page.waitForSelector('.page-frame.page-ready', { timeout: 90000 });
  await page.waitForTimeout(800);
  return { ctx, page, errors, requests };
}
const pdfCalls = (reqs) => reqs.filter((p) => p === '/api/view/derived-pdf' || p === '/api/view/pdf-source').length;
const pngCalls = (reqs) => reqs.filter((p) => p.startsWith('/api/view/pages/')).length;
const shot = (page, name) => page.screenshot({ path: path.join(shotDir, `${name}.png`) });
const inputValue = (page) => page.$eval('#page-jump-input', (el) => el.value);
const frameBox = (page, n = 1) => page.locator(`.page-frame[data-page="${n}"]`).boundingBox();

const pdfId = await upload(await makePdf(), '보기시험.pdf');
const docxId = await upload(makeDocx(), '보기시험.docx');
const pptxId = await upload(await buildFixture(), '보기시험.pptx');
const FULL = { canDownload: true, canCopy: true, canPrint: true };

// ---------- 1. 권한별 안쪽 ----------
const docs = [
  ['pdf', pdfId, 3],
  ['docx', docxId, null],
  ['pptx', pptxId, 1],
];
for (const [fmt, id, pages] of docs) {
  for (const [mode, body] of [
    ['열람만', {}],
    ['내려받기', FULL],
  ]) {
    const t = await token(id, body);
    check(t.viewerUrl.startsWith('/view.html?'), `${fmt} ${mode}: viewerUrl이 /view.html`, t.viewerUrl.split('?')[0]);
    const { ctx, page, errors, requests } = await openView(t.viewerUrl);
    const spans = await page.locator('.pdf-text-layer span').count();
    const st = await page.evaluate(() => ({
      edit: document.getElementById('edit-btn').hidden,
      dl: document.getElementById('download-btn').disabled,
      dlTitle: document.getElementById('download-btn').title,
      print: document.getElementById('print-btn').disabled,
      total: document.getElementById('page-jump-total').textContent,
      count: document.querySelectorAll('.page-frame').length,
      presence: document.querySelectorAll('#presence-list .presence-badge').length,
      wm: getComputedStyle(document.documentElement).getPropertyValue('--watermark-bg').trim(),
    }));
    if (mode === '열람만') {
      check(pdfCalls(requests) === 0 && pngCalls(requests) > 0, `${fmt} 열람만: PDF 요청 0건·서버 PNG`, `pdf ${pdfCalls(requests)} png ${pngCalls(requests)}`);
      check(spans === 0, `${fmt} 열람만: 글자 층 없음`, String(spans));
      check(st.edit && st.dl && st.dlTitle === '내려받기 권한이 없습니다' && st.print, `${fmt} 열람만: 편집 숨김·내려받기/인쇄 꺼짐`, JSON.stringify({ ...st, wm: undefined }));
    } else {
      check(pdfCalls(requests) === 1 && pngCalls(requests) === 0, `${fmt} 내려받기: derived-pdf 1건·PNG 0건`, `pdf ${pdfCalls(requests)} png ${pngCalls(requests)}`);
      check(spans > 0, `${fmt} 내려받기: 글자 층 있음`, String(spans));
      check(!st.dl, `${fmt} 내려받기: 내려받기 켜짐`);
    }
    const expectTotal = `의 ${pages ?? st.count}`;
    check(st.total === expectTotal && st.presence >= 1 && st.wm !== '', `${fmt} ${mode}: 쪽 수·함께 보는 사람·워터마크`, `${st.total} / ${st.presence} / ${st.wm.slice(0, 30)}`);
    check(errors.length === 0, `${fmt} ${mode}: 오류 없음`, errors.join(' | '));
    await shot(page, `${fmt}-${mode}`);
    await ctx.close();
  }
}

// ---------- 2·3·4·5. 막대·썸네일·찾기·측정(PDF 내려받기) ----------
const pdfFull = await token(pdfId, FULL);
const pdfVo = await token(pdfId, {});
{
  const { ctx, page, errors } = await openView(pdfFull.viewerUrl);
  await page.click('#zoom-in');
  check((await page.textContent('#zoom-level')) === '125%', '확대 → 125%', await page.textContent('#zoom-level'));
  await page.click('#zoom-fit-btn');
  await page.click('#zoom-fit-btn');
  await page.waitForTimeout(400);
  const fitH = (await frameBox(page)).height;
  check(fitH <= 900, '맞춤 두 번(너비 → 쪽) 뒤 쪽 높이 ≤ 창', `${Math.round(fitH)}px`);
  await page.click('#page-next');
  await page.waitForTimeout(300);
  check((await inputValue(page)) === '2', '다음 쪽 → 2', await inputValue(page));
  await page.fill('#page-jump-input', '3');
  await page.press('#page-jump-input', 'Enter');
  await page.waitForTimeout(300);
  check((await inputValue(page)) === '3', '쪽 번호 3 + Enter → 3', await inputValue(page));
  check(await page.$eval('#page-next', (b) => b.disabled), '끝 쪽에서 다음 단추 꺼짐');
  await page.click('#page-prev');
  await page.waitForTimeout(300);
  check((await inputValue(page)) === '2', '이전 쪽 → 2', await inputValue(page));
  await page.click('#rotate-btn');
  check(await page.$eval('#pages', (p) => p.classList.contains('pages-rotated')), '회전 → .pages-rotated');
  await page.click('#rotate-btn');
  await page.click('#rotate-btn');
  await page.click('#rotate-btn');
  const before = parseInt(await page.textContent('#zoom-level'), 10);
  const fb = await frameBox(page, 2);
  await page.mouse.move(fb.x + fb.width / 2, fb.y + fb.height / 2);
  await page.keyboard.down('Control');
  await page.mouse.wheel(0, -100);
  await page.keyboard.up('Control');
  await page.waitForTimeout(300);
  const after = parseInt(await page.textContent('#zoom-level'), 10);
  check(after > before, 'Ctrl+휠(위) → 확대', `${before}% → ${after}%`);

  // 썸네일
  const th = await page.evaluate(() => ({
    open: !document.getElementById('thumbs-panel').hidden,
    items: document.querySelectorAll('.thumb-item').length,
    pics: document.querySelectorAll('.thumb-box canvas, .thumb-box img').length,
  }));
  check(th.open && th.items === 3 && th.pics >= 1, '썸네일: 1400px에서 열림·3개·그림', JSON.stringify(th));
  await page.click('.thumb-item:nth-child(3)');
  await page.waitForTimeout(400);
  check((await inputValue(page)) === '3' && (await page.$eval('.thumb-item.current', (e) => e.dataset.page)) === '3', '썸네일 3 누름 → 3쪽·현재 표시');

  // 측정·반전
  await page.click('#measure-btn');
  check(await page.$eval('#measure-bar', (e) => !e.hidden), '측정: 막대 보임');
  const mb = await frameBox(page, 3);
  await page.mouse.click(mb.x + mb.width * 0.2, mb.y + mb.height * 0.3);
  await page.mouse.click(mb.x + mb.width * 0.6, mb.y + mb.height * 0.3);
  const label = await page.locator('.measure-label').first().textContent().catch(() => '');
  check(label.includes('mm'), '측정: 두 점 → mm 표시', label);
  await page.keyboard.press('Escape');
  check(await page.$eval('#measure-bar', (e) => e.hidden), '측정: Esc → 막대 숨김');
  await page.click('#invert-btn');
  check(await page.$eval('#pages', (p) => p.classList.contains('pages-invert')), '반전 → .pages-invert');
  await page.click('#invert-btn');
  check(errors.length === 0, 'PDF 막대 조작 중 오류 없음', errors.join(' | '));
  await ctx.close();
}

// 찾기 — 두 모드
for (const [mode, t] of [
  ['열람만(PNG)', pdfVo],
  ['내려받기(pdf.js)', pdfFull],
]) {
  const { ctx, page, errors } = await openView(t.viewerUrl);
  await page.keyboard.press('Control+f');
  check(await page.$eval('#search-bar', (e) => !e.hidden), `찾기 ${mode}: Ctrl+F → 막대`);
  await page.fill('#search-input', 'Target');
  await page.press('#search-input', 'Enter');
  await page.waitForFunction(() => /\d+ \/ \d+/.test(document.getElementById('search-count').textContent), null, { timeout: 30000 });
  const c1 = await page.textContent('#search-count');
  const cur = await page.locator('.search-hit-overlay.current').count();
  check(c1 === '1 / 3' && cur >= 1, `찾기 ${mode}: 1 / 3·강조`, `${c1} / ${cur}`);
  await page.press('#search-input', 'Enter');
  check((await page.textContent('#search-count')) === '2 / 3', `찾기 ${mode}: Enter → 2 / 3`, await page.textContent('#search-count'));
  await page.press('#search-input', 'Shift+Enter');
  check((await page.textContent('#search-count')) === '1 / 3', `찾기 ${mode}: Shift+Enter → 1 / 3`, await page.textContent('#search-count'));
  await page.keyboard.press('F3');
  check((await page.textContent('#search-count')) === '2 / 3', `찾기 ${mode}: F3 → 2 / 3`, await page.textContent('#search-count'));
  await page.fill('#search-input', 'zzzz없는말');
  await page.press('#search-input', 'Enter');
  await page.waitForFunction(() => document.getElementById('search-count').textContent === '결과 없음', null, { timeout: 30000 }).catch(() => undefined);
  check((await page.textContent('#search-count')) === '결과 없음', `찾기 ${mode}: 없는 말 → 결과 없음`, await page.textContent('#search-count'));
  await page.keyboard.press('Escape');
  const closed = await page.evaluate(() => ({ hidden: document.getElementById('search-bar').hidden, hits: document.querySelectorAll('.search-hit-overlay').length }));
  check(closed.hidden && closed.hits === 0, `찾기 ${mode}: Esc → 닫힘·강조 0`, JSON.stringify(closed));
  check(errors.length === 0, `찾기 ${mode}: 오류 없음`, errors.join(' | '));
  await ctx.close();
}

// 썸네일 — 열람만은 서버 작은 그림(img)
{
  const { ctx, page } = await openView(pdfVo.viewerUrl);
  await page.waitForSelector('.thumb-box img', { timeout: 15000 }).catch(() => undefined);
  check((await page.locator('.thumb-box img').count()) >= 1, '썸네일 열람만: 서버 그림(img)');
  await ctx.close();
}

// 워드: 측정·반전 꺼짐
{
  const t = await token(docxId, FULL);
  const { ctx, page } = await openView(t.viewerUrl);
  const st = await page.evaluate(() => ['measure-btn', 'invert-btn'].map((id) => [document.getElementById(id).disabled, document.getElementById(id).title]));
  check(st.every(([d, title]) => d && title === 'PDF 문서에서만 쓸 수 있습니다'), '워드: 측정·반전 꺼짐 + 이유', JSON.stringify(st));
  await ctx.close();
}

// ---------- 6. 표시(markup) ----------
{
  const vo = await openView(pdfVo.viewerUrl);
  const st = await vo.page.$eval('#markup-add-btn', (b) => [b.disabled, b.title]);
  check(st[0] && st[1] === '표시를 넣을 권한이 없습니다', '표시: 열람만은 꺼짐 + 이유', JSON.stringify(st));
  await vo.ctx.close();
  const ed = await token(pdfId, { ...FULL, canEdit: true });
  const { ctx, page, errors } = await openView(ed.viewerUrl);
  check(await page.$eval('#markup-add-btn', (b) => !b.disabled), '표시: 편집 토큰은 켜짐');
  await page.click('#markup-add-btn');
  const b = await frameBox(page, 1);
  await page.mouse.move(b.x + b.width * 0.2, b.y + b.height * 0.2);
  await page.mouse.down();
  await page.mouse.move(b.x + b.width * 0.4, b.y + b.height * 0.3, { steps: 5 });
  await page.mouse.up();
  await page.waitForSelector('.markup-popover textarea', { timeout: 5000 });
  await page.fill('.markup-popover textarea', '보기 껍데기 표시');
  await page.click('.markup-popover .btn-pill-primary');
  await page.waitForSelector('.markup-rect:not(.markup-rect-draft)', { timeout: 5000 }).catch(() => undefined);
  const list = (await (await fetch(`${BASE}/api/view/markup?token=${encodeURIComponent(ed.token)}`)).json()).data ?? [];
  check(list.length === 1 && list[0].text === '보기 껍데기 표시', '표시: 끌기 → 팝오버 저장 → 서버 1건', String(list.length));
  check(errors.length === 0, '표시: 오류 없음', errors.join(' | '));
  await ctx.close();
}

// ---------- 7. pdf.js 실패 → PNG ----------
{
  const { ctx, page, requests } = await openView(pdfFull.viewerUrl, {
    setup: (p) => p.route('**/api/view/derived-pdf*', (r) => r.fulfill({ status: 200, contentType: 'application/pdf', body: 'not a pdf' })),
  });
  const note = await page.textContent('#view-notice');
  check(note.includes('서버 그림') && pngCalls(requests) > 0, 'pdf.js 실패 → 서버 PNG + 안내', `${note} / png ${pngCalls(requests)}`);
  await ctx.close();
}

// ---------- 8. 편집 → 보기로 왕복(워드) ----------
const docxEdit = await token(docxId, { ...FULL, canEdit: true });
{
  const { ctx, page } = await openView(`${docxEdit.viewerUrl}&home=hidden`);
  check(await page.$eval('#edit-btn', (b) => !b.hidden), '워드 편집 토큰: 편집 단추 보임');
  await page.click('#edit-btn');
  await page.waitForURL(/\/editor\.html/, { timeout: 30000 });
  const u = new URL(page.url());
  check(u.searchParams.get('from') === 'view' && u.searchParams.get('home') === 'hidden', '편집 단추 → editor.html?from=view&home=hidden', u.search.replace(/token=[^&]+/, 'token=…'));
  await page.waitForSelector('.ProseMirror[contenteditable="true"]', { timeout: 60000 });
  await page.click('.ProseMirror p >> nth=0');
  await page.keyboard.press('End');
  await page.keyboard.insertText(' 껍데기왕복');
  await page.waitForTimeout(800);
  await page.click('#back-to-view');
  await page.waitForURL(/\/view\.html/, { timeout: 120000 });
  await page.waitForSelector('.page-frame.page-ready', { timeout: 120000 });
  await page.waitForFunction(() => document.querySelector('.pdf-text-layer')?.textContent.includes('껍데기왕복'), null, { timeout: 30000 }).catch(() => undefined);
  const txt = await page.$eval('.pdf-text-layer', (e) => e.textContent).catch(() => '');
  check(txt.includes('껍데기왕복'), '보기로 → view.html에 고친 글', txt.slice(0, 60));
  check(new URL(page.url()).searchParams.get('home') === 'hidden', '돌아온 주소에 home=hidden 유지');
  await ctx.close();
}

// ---------- 9. 편집기 바로 열기(향기랩 방식) ----------
{
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await ctx.newPage();
  await page.goto(BASE + docxEdit.editorUrl);
  await page.waitForSelector('.ProseMirror', { timeout: 60000 });
  await page.waitForTimeout(5000);
  check((await page.locator('#back-to-view').count()) === 0, '편집기 바로 열기: "보기로" 없음');
  await ctx.close();
}

// ---------- 10. PPT·PDF 편집기 "보기로" ----------
{
  const t = await token(pptxId, FULL);
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await ctx.newPage();
  await page.goto(`${BASE}${t.editorUrl}&from=view`);
  await page.waitForSelector('#back-to-view', { state: 'visible', timeout: 60000 }).catch(() => undefined);
  check((await page.locator('#back-to-view').count()) === 1, 'PPT 편집기(from=view): "보기로" 보임');
  await page.click('#back-to-view');
  await page.waitForURL(/\/view\.html/, { timeout: 60000 }).catch(() => undefined);
  check(page.url().includes('/view.html'), 'PPT 보기로 → view.html');
  await ctx.close();

  const pt = await token(pdfId, { canDownload: true, canEdit: true, pdfEditor: true });
  const v = await openView(pt.viewerUrl);
  await v.page.click('#edit-btn');
  await v.page.waitForURL(/\/editor-pdf\.html/, { timeout: 30000 });
  check(new URL(v.page.url()).searchParams.get('from') === 'view', 'PDF 편집 단추 → editor-pdf.html?from=view');
  await v.page.waitForSelector('.pdf-titlebar #back-to-view', { timeout: 60000 }).catch(() => undefined);
  check((await v.page.locator('.pdf-titlebar #back-to-view').count()) === 1, 'PDF 편집기 머리에 "보기로"');
  await v.page.click('.pdf-titlebar #back-to-view');
  await v.page.waitForURL(/\/view\.html/, { timeout: 60000 }).catch(() => undefined);
  check(v.page.url().includes('/view.html'), 'PDF 보기로 → view.html');
  await v.ctx.close();
}

// ---------- 11. 넘기기 ----------
{
  const t = await token(docxId, {});
  const { ctx, page } = await openView(`/viewer.html?token=${encodeURIComponent(t.token)}`);
  check(new URL(page.url()).pathname === '/view.html', '예전 /viewer.html 링크(워드) → /view.html');
  await ctx.close();
  const xlsxId = await upload(await makeXlsx(), '보기시험.xlsx');
  const xt = await token(xlsxId, {});
  check(xt.viewerUrl.startsWith('/viewer.html?'), '엑셀 viewerUrl은 /viewer.html', xt.viewerUrl.split('?')[0]);
  const c2 = await browser.newContext();
  const p2 = await c2.newPage();
  await p2.goto(`${BASE}/view.html?token=${encodeURIComponent(xt.token)}`);
  await p2.waitForSelector('#grid-table', { timeout: 60000 }).catch(() => undefined);
  check(new URL(p2.url()).pathname === '/viewer.html' && (await p2.locator('#grid-table').count()) === 1, '엑셀을 /view.html로 열면 /viewer.html 격자로');
  await c2.close();
}

// ---------- 12. 첫 변환 중 ----------
{
  const id = await upload(await buildFixture(), '준비시험.pptx', { wait: false });
  const t = await token(id, FULL);
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await ctx.newPage();
  const requests = [];
  page.on('request', (r) => requests.push(new URL(r.url()).pathname));
  await page.goto(BASE + t.viewerUrl);
  let sawPrep = false;
  let editDuringPrep = true;
  const t0 = Date.now();
  while (Date.now() - t0 < 2000) {
    const st = await page.evaluate(() => ({ prep: !document.getElementById('prep-screen').hidden, edit: !document.getElementById('edit-btn').hidden, ready: !!document.querySelector('.page-frame.page-ready') }));
    if (st.prep) {
      sawPrep = true;
      if (!st.edit) editDuringPrep = false;
    }
    if (st.ready) break;
    await wait(100);
  }
  await page.waitForSelector('.page-frame.page-ready', { timeout: 120000 });
  const infoCalls = requests.filter((p) => p === '/api/view/info').length;
  check(editDuringPrep && infoCalls <= 2, '첫 변환 중: 결국 쪽이 그려짐·준비 중 편집 단추·info ≤ 2', `${sawPrep ? '준비 화면 보임' : '준비 화면 못 봄(변환이 빠름)'}, info ${infoCalls}`);
  await ctx.close();
}

// ---------- 13. 좁은 화면 ----------
{
  const { ctx, page } = await openView(pdfFull.viewerUrl, { viewport: { width: 390, height: 844 }, ctxOpts: { hasTouch: true, isMobile: true } });
  const box = await page.locator('#tools-btn').boundingBox();
  const st = await page.evaluate(() => ({
    thumbsHidden: document.getElementById('thumbs-panel').hidden,
    inMenu: ['text-tool-btn', 'print-btn'].every((id) => document.getElementById(id).closest('#tools-menu')),
  }));
  check(!!box && box.x + box.width <= 390, '390px: 도구▾ 화면 안', box ? `${Math.round(box.x + box.width)}px` : '없음');
  check(st.thumbsHidden && st.inMenu, '390px: 썸네일 닫힘·글 넣기·인쇄가 도구▾ 안', JSON.stringify(st));
  await page.click('#tools-btn');
  check(await page.$eval('#tools-menu', (m) => !m.hidden), '도구▾ 누르면 메뉴 열림');
  await shot(page, 'narrow');
  await page.setViewportSize({ width: 1400, height: 900 });
  await page.waitForTimeout(300);
  const wide = await page.evaluate(() => ({
    empty: document.getElementById('tools-menu').children.length === 0,
    back: !document.getElementById('print-btn').closest('#tools-menu') && !!document.getElementById('print-btn').closest('.toolbar-tools-right'),
  }));
  check(wide.empty && wide.back, '1400px로 넓히면 메뉴가 비고 단추가 제자리', JSON.stringify(wide));
  await ctx.close();
}

// ---------- 14. iframe ----------
{
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  const page = await ctx.newPage();
  await page.goto(`${BASE}/api-docs.html`);
  await page.evaluate(() => {
    window.__msgs = [];
    window.addEventListener('message', (e) => window.__msgs.push(e.data?.type));
  });
  await page.evaluate((src) => {
    const f = document.createElement('iframe');
    f.src = src;
    f.style.cssText = 'width:1000px;height:700px';
    document.body.prepend(f);
  }, `${pdfVo.viewerUrl}&back=parent`);
  await page.waitForFunction(() => window.__msgs.includes('docview:ready'), null, { timeout: 60000 }).catch(() => undefined);
  check(await page.evaluate(() => window.__msgs.includes('docview:ready')), 'iframe: docview:ready 받음');
  const frame = page.frames().find((f) => f.url().includes('/view.html'));
  await frame.click('#back-btn');
  await page.waitForFunction(() => window.__msgs.includes('docview:back'), null, { timeout: 5000 }).catch(() => undefined);
  check(await page.evaluate(() => window.__msgs.includes('docview:back')), 'iframe: ← → docview:back 받음');
  check(frame.url().includes('/view.html'), 'iframe: ← 눌러도 iframe은 이동하지 않음');
  await ctx.close();
}

await browser.close();
for (const id of uploaded) await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' }).catch(() => undefined);
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
