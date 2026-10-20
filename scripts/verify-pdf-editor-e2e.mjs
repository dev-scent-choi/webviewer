// PDF 편집기(editor-pdf.html) 1단계 브라우저 확인 — 시험 서버(기본 http://127.0.0.1:8124)에 올리고 실제 화면을 조작한다.
//   node scripts/verify-pdf-editor-e2e.mjs [BASE] [스크린샷 폴더]
// 시험 PDF는 여기서 pdf-lib로 만든다: (1) 동봉 Pretendard를 넣은 한글 3쪽 + 책갈피, (2) 레이어(선택적 콘텐츠),
// (3) 글꼴이 안 들어 있는 한글(Adobe-Korea1 / UniKS-UCS2-H) — pdfcraft 웹판이 글자를 안 그리던 경우.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import fontkit from '@pdf-lib/fontkit';
import { PDFDocument, PDFHexString, PDFName, PDFNumber, PDFOperator, PDFString, rgb } from 'pdf-lib';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-editor-'));
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

// ---------- 시험 PDF ----------
async function koreanWithOutline() {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  // subset:true는 pdf-lib+fontkit에서 글자가 빠지는 알려진 문제가 있다(rendering.ts도 통째로 넣는다)
  const font = await doc.embedFont(fs.readFileSync('assets/fonts/Pretendard-Regular.ttf'), { subset: false });
  const titles = ['분기 보고서', '두 번째 장 — 매출', '세 번째 장 — 결론'];
  const pages = titles.map((title, i) => {
    const page = doc.addPage([595, 842]);
    page.drawText(title, { x: 60, y: 760, size: 26, font, color: rgb(0.1, 0.25, 0.45) });
    page.drawText(`${i + 1}쪽 본문입니다. 가동률 97.2% 한글과 English mixed text.`, { x: 60, y: 700, size: 13, font });
    return page;
  });
  // 책갈피(개요) — pdf-lib에 API가 없어 사전을 직접 만든다
  const ctx = doc.context;
  const outlinesRef = ctx.nextRef();
  const itemRefs = pages.map(() => ctx.nextRef());
  pages.forEach((page, i) => {
    const item = ctx.obj({
      Title: PDFHexString.fromText(titles[i]),
      Parent: outlinesRef,
      Dest: [page.ref, PDFName.of('Fit')],
    });
    if (i > 0) item.set(PDFName.of('Prev'), itemRefs[i - 1]);
    if (i < pages.length - 1) item.set(PDFName.of('Next'), itemRefs[i + 1]);
    ctx.assign(itemRefs[i], item);
  });
  ctx.assign(outlinesRef, ctx.obj({ Type: 'Outlines', First: itemRefs[0], Last: itemRefs[itemRefs.length - 1], Count: PDFNumber.of(pages.length) }));
  doc.catalog.set(PDFName.of('Outlines'), outlinesRef);
  return doc.save();
}

async function withLayer() {
  const doc = await PDFDocument.create();
  const page = doc.addPage([400, 400]);
  const ctx = doc.context;
  const ocg = ctx.register(ctx.obj({ Type: 'OCG', Name: PDFString.of('Blue box') }));
  doc.catalog.set(PDFName.of('OCProperties'), ctx.obj({ OCGs: [ocg], D: ctx.obj({ Order: [ocg], ON: [ocg] }) }));
  const resources = page.node.Resources();
  resources.set(PDFName.of('Properties'), ctx.obj({ OC1: ocg }));
  page.pushOperators(
    PDFOperator.of('BDC', [PDFName.of('OC'), PDFName.of('OC1')]),
    PDFOperator.of('rg', [PDFNumber.of(0), PDFNumber.of(0), PDFNumber.of(1)]),
    PDFOperator.of('re', [PDFNumber.of(50), PDFNumber.of(50), PDFNumber.of(300), PDFNumber.of(300)]),
    PDFOperator.of('f'),
    PDFOperator.of('EMC'),
  );
  return doc.save();
}

async function koreanNotEmbedded() {
  const doc = await PDFDocument.create();
  const page = doc.addPage([400, 300]);
  const ctx = doc.context;
  const descriptor = ctx.obj({
    Type: 'FontDescriptor',
    FontName: 'HYGoThic-Medium',
    Flags: 4,
    FontBBox: [-6, -145, 1003, 880],
    ItalicAngle: 0,
    Ascent: 880,
    Descent: -120,
    CapHeight: 880,
    StemV: 93,
  });
  const cid = ctx.obj({
    Type: 'Font',
    Subtype: 'CIDFontType0',
    BaseFont: 'HYGoThic-Medium',
    CIDSystemInfo: ctx.obj({ Registry: PDFString.of('Adobe'), Ordering: PDFString.of('Korea1'), Supplement: 2 }),
    FontDescriptor: ctx.register(descriptor),
    DW: 1000,
  });
  const font = ctx.register(ctx.obj({ Type: 'Font', Subtype: 'Type0', BaseFont: 'HYGoThic-Medium', Encoding: 'UniKS-UCS2-H', DescendantFonts: [ctx.register(cid)] }));
  page.node.Resources().set(PDFName.of('Font'), ctx.obj({ F1: font }));
  // "안녕하세요" (UCS-2)
  const hex = [...'안녕하세요'].map((ch) => ch.codePointAt(0).toString(16).padStart(4, '0')).join('');
  page.pushOperators(
    PDFOperator.of('BT'),
    PDFOperator.of('Tf', [PDFName.of('F1'), PDFNumber.of(28)]),
    PDFOperator.of('Td', [PDFNumber.of(50), PDFNumber.of(150)]),
    PDFOperator.of('Tj', [PDFHexString.of(hex)]),
    PDFOperator.of('ET'),
  );
  return doc.save();
}

// ---------- 서버 ----------
async function upload(bytes, name) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
  const res = await fetch(`${BASE}/api/documents`, { method: 'POST', body: form });
  const { data } = await res.json();
  for (let i = 0; i < 200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${data.id}`)).json()).data;
    if (st.status === 'ready' || st.editReady) return data.id;
    if (st.status === 'failed') throw new Error('변환 실패');
    await new Promise((r) => setTimeout(r, 150));
  }
  throw new Error('대기 시간 초과');
}
const token = (id, body) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '시험', canPrint: true, canDownload: true, canCopy: true, canEdit: true, ...body }),
  }).then((r) => r.json());

const mainBytes = await koreanWithOutline();
const layerBytes = await withLayer();
const plainBytes = await koreanNotEmbedded();
const mainId = await upload(mainBytes, '보고서.pdf');
const layerId = await upload(layerBytes, 'layer.pdf');
const plainId = await upload(plainBytes, '글꼴없음.pdf');

// 1. 토큰
const t8000 = await token(mainId, {});
const tPdf = await token(mainId, { pdfEditor: true });
const tNoDownload = await token(mainId, { pdfEditor: true, canDownload: false });
check(!t8000.data.editorUrl && /viewer\.html/.test(t8000.data.viewerUrl), '8000처럼 pdfEditor 없이 → editorUrl 없음(보기 화면 그대로)');
check(/^\/editor-pdf\.html\?token=/.test(tPdf.data.editorUrl ?? ''), 'pdfEditor → PDF 편집기 주소');
check(!tNoDownload.data.editorUrl, '내려받기 권한 없으면 PDF 편집기 주소 없음');
const srcPlain = await fetch(`${BASE}/api/view/pdf-source?token=${encodeURIComponent(t8000.data.token)}`);
check(srcPlain.status === 403, 'pdfEditor 아닌 토큰으로 원본 요청 → 403', String(srcPlain.status));

// 2~. 화면
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => {
  if (m.type() === 'error') errors.push(m.text());
});

async function open(id) {
  const t = await token(id, { pdfEditor: true });
  await page.goto(`${BASE}${t.data.editorUrl}`);
  await page.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
  await page.waitForTimeout(500);
  return t;
}
const inkOf = (n) =>
  page.evaluate((n) => {
    const c = document.querySelector(`.pdf-page[data-page="${n}"] canvas`);
    if (!c || !c.width) return 0;
    const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
    let dark = 0;
    for (let i = 0; i < d.length; i += 16) if (d[i] + d[i + 1] + d[i + 2] < 600) dark += 1;
    return dark;
  }, n);
const textOf = (n) => page.evaluate((n) => document.querySelector(`.pdf-page[data-page="${n}"] .pdf-text-layer`)?.textContent ?? '', n);

const tMain = await open(mainId);
check((await inkOf(1)) > 50 && (await textOf(1)).includes('분기 보고서'), '글꼴 포함 한글 PDF — 그려지고 글자 층에 "분기 보고서"', String(await inkOf(1)));
check((await page.textContent('#page-count')).includes('3'), '쪽 수 표시 "/ 3"');
check((await page.textContent('#doc-name')) === '보고서.pdf', '문서 탭에 파일 이름');
await page.screenshot({ path: path.join(shotDir, '1-main.png') });

await page.click('#page-next');
await page.waitForTimeout(400);
check((await page.inputValue('#page-input')) === '2', '▼ → 2쪽');
await page.fill('#page-input', '3');
await page.press('#page-input', 'Enter');
await page.waitForTimeout(400);
check((await page.inputValue('#page-input')) === '3', '쪽 번호 입력 → 3쪽');

const z0 = await page.textContent('#zoom-level');
await page.click('#zoom-in');
await page.waitForTimeout(300);
const z1 = await page.textContent('#zoom-level');
await page.click('#zoom-out');
await page.click('#zoom-out');
await page.waitForTimeout(300);
const z2 = await page.textContent('#zoom-level');
check(parseInt(z1) > parseInt(z0) && parseInt(z2) < parseInt(z1), '확대·축소 배율 표시', `${z0} → ${z1} → ${z2}`);
const before = await page.evaluate(() => {
  const el = document.querySelector('.pdf-page[data-page="1"]');
  return el.offsetWidth / el.offsetHeight;
});
await page.click('#rotate-btn');
await page.waitForTimeout(400);
const after = await page.evaluate(() => {
  const el = document.querySelector('.pdf-page[data-page="1"]');
  return el.offsetWidth / el.offsetHeight;
});
check(before < 1 && after > 1, '보기 회전 → 쪽이 가로로', `${before.toFixed(2)} → ${after.toFixed(2)}`);
await page.click('#rotate-btn');
await page.click('#rotate-btn');
await page.click('#rotate-btn');
await page.waitForTimeout(300);

// 썸네일·책갈피
await page.click('.pdf-thumb[data-page="2"]');
await page.waitForTimeout(400);
check((await page.inputValue('#page-input')) === '2', '썸네일 2 → 2쪽');
await page.click('#rail-outline');
await page.waitForSelector('#panel-outline:not([hidden]) .pdf-outline-link');
const outlineTitles = await page.$$eval('.pdf-outline-link span', (els) => els.map((e) => e.textContent));
check(outlineTitles.join('|') === '분기 보고서|두 번째 장 — 매출|세 번째 장 — 결론', '책갈피 목록(한글)', outlineTitles.join('|'));
await page.click('.pdf-outline-link >> nth=2');
await page.waitForTimeout(400);
check((await page.inputValue('#page-input')) === '3', '책갈피 눌러 3쪽');
await page.screenshot({ path: path.join(shotDir, '2-outline.png') });

// 찾기
await page.keyboard.press('Control+f');
check(await page.isVisible('#find-bar'), 'Ctrl+F → 찾기 줄');
await page.fill('#find-input', '가동률');
await page.waitForFunction(() => /\d+\/3$/.test(document.querySelector('#find-count')?.textContent ?? ''), null, { timeout: 8000 }).catch(() => {});
const count = await page.textContent('#find-count');
check(/^\d\/3$/.test(count ?? ''), '"가동률" 3쪽 모두에서 → n/3', count);
check((await page.$$('.pdf-find-hit')).length >= 1 && (await page.$$('.pdf-find-hit.current')).length === 1, '강조 상자·현재 결과 표시');
await page.keyboard.press('Enter');
await page.waitForTimeout(300);
const count2 = await page.textContent('#find-count');
check(count2 !== count, 'Enter → 다음 결과', `${count} → ${count2}`);
await page.screenshot({ path: path.join(shotDir, '3-find.png') });
await page.keyboard.press('Escape');
check(!(await page.isVisible('#find-bar')) && (await page.$$('.pdf-find-hit')).length === 0, 'Esc → 찾기 닫고 강조 지움');

// 읽기 모드
await page.click('.pdf-mode[data-mode="read"]');
check(!(await page.isVisible('#tools-panel')), '읽기 모드 → 도구 목록 숨김');
await page.click('.pdf-mode[data-mode="tools"]');
check(await page.isVisible('#tools-panel'), '모든 도구 → 도구 목록');

// 내려받기 = 원본 바이트
const dl = await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(tMain.data.token)}`);
const dlBytes = new Uint8Array(await dl.arrayBuffer());
check(dlBytes.length === mainBytes.length && dlBytes.every((b, i) => b === mainBytes[i]), '내려받기 = 원본 그대로(1단계는 원본을 바꾸지 않음)');

// 레이어
await open(layerId);
await page.click('#rail-layers');
await page.waitForSelector('#panel-layers:not([hidden]) input[type="checkbox"]');
const inkOn = await inkOf(1);
await page.click('#panel-layers input[type="checkbox"]');
await page.waitForTimeout(800);
const inkOff = await inkOf(1);
check(inkOn > 1000 && inkOff < inkOn / 10, '레이어 끄면 파란 상자가 사라짐', `${inkOn} → ${inkOff}`);
await page.screenshot({ path: path.join(shotDir, '4-layer-off.png') });

// 글꼴 없는 한글
await open(plainId);
const inkPlain = await inkOf(1);
const textPlain = await textOf(1);
check(inkPlain > 30 && textPlain.includes('안녕하세요'), '글꼴이 안 들어 있는 한글 PDF도 그려짐(cMap)', `${inkPlain} "${textPlain}"`);
await page.screenshot({ path: path.join(shotDir, '5-not-embedded.png') });

// 부모 창 준비 신호
const host = await browser.newPage();
const t2 = await token(mainId, { pdfEditor: true });
await host.setContent(`<iframe id="f" src="${BASE}${t2.data.editorUrl}&back=parent" style="width:1200px;height:800px"></iframe>`);
const ready = await host.evaluate(
  () => new Promise((resolve) => {
    const timer = setTimeout(() => resolve(false), 20000);
    window.addEventListener('message', (e) => {
      if (e.data?.type === 'docview:ready') {
        clearTimeout(timer);
        resolve(true);
      }
    });
  }),
);
check(ready, 'iframe 부모에 docview:ready');

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 3).join(' | '));
await browser.close();
for (const id of [mainId, layerId, plainId]) await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' });
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
