// PDF 편집기 3단계(페이지 정리) 브라우저 확인 — 시험 서버(기본 http://127.0.0.1:8124).
//   node scripts/verify-pdf-organize-e2e.mjs [BASE] [스크린샷 폴더]
// 6쪽 한글 PDF(3쪽에 사각형 주석)를 올려 정리 화면에서 지우기·회전·앞으로·끌어 옮기기·빈 쪽·파일에서 넣기·실행 취소를 하고,
// 본문·썸네일·메모 목록이 새 순서를 따르는지, 추출·나누기 내려받기, 저장 파일(쪽 순서·회전·주석), 다른 참여자, 다시 열기, 읽기 전용을 확인한다.
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import fontkit from '@pdf-lib/fontkit';
import { unzipSync } from 'fflate';
import { PDFDocument, PDFString, rgb } from 'pdf-lib';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8124';
const shotDir = process.argv[3] ?? fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-organize-'));
const results = [];
const check = (ok, label, detail = '') => {
  results.push(ok);
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${label}${detail ? ` — ${detail}` : ''}`);
};

async function makePdf(labels, withAnnotOn) {
  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit);
  const font = await doc.embedFont(fs.readFileSync('assets/fonts/Pretendard-Regular.ttf'), { subset: false });
  labels.forEach((label, i) => {
    const p = doc.addPage([595, 842]);
    p.drawText(label, { x: 60, y: 760, size: 30, font, color: rgb(0.1, 0.2, 0.4) });
    if (i === withAnnotOn) {
      const ctx = doc.context;
      p.node.addAnnot(ctx.register(ctx.obj({ Type: 'Annot', Subtype: 'Square', Rect: [50, 600, 300, 700], C: [0.85, 0.2, 0.15], NM: PDFString.of(`sq-${label}`) })));
    }
  });
  return doc.save();
}

async function upload(bytes, name) {
  const form = new FormData();
  form.append('file', new Blob([bytes]), name);
  const { data } = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  for (let i = 0; i < 200; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${data.id}`)).json()).data;
    if (st.status === 'ready' || st.editReady) return data.id;
    await new Promise((r) => setTimeout(r, 150));
  }
  throw new Error('대기 시간 초과');
}
const token = (id, body = {}) =>
  fetch(`${BASE}/api/documents/${id}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '시험 편집자', canPrint: true, canDownload: true, canCopy: true, canEdit: true, pdfEditor: true, ...body }),
  }).then((r) => r.json());

/** 쪽마다 글자·회전·주석 종류 */
async function describePdf(bytes) {
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise;
  const out = [];
  for (let p = 1; p <= pdf.numPages; p += 1) {
    const page = await pdf.getPage(p);
    const text = (await page.getTextContent()).items.map((it) => it.str ?? '').join('').trim();
    const annots = (await page.getAnnotations()).map((a) => a.subtype);
    out.push({ text, rotate: page.rotate, annots });
  }
  await pdf.destroy();
  return out;
}

const original = await makePdf(['1쪽', '2쪽', '3쪽', '4쪽', '5쪽', '6쪽'], 2);
const partBytes = await makePdf(['부록 1', '부록 2'], 1);
const partFile = path.join(shotDir, '부록.pdf');
fs.writeFileSync(partFile, partBytes);
const id = await upload(original, '정리시험.pdf');
const browser = await chromium.launch();
const ctx = await browser.newContext({ viewport: { width: 1500, height: 950 }, acceptDownloads: true });
const page = await ctx.newPage();
const errors = [];
page.on('pageerror', (e) => errors.push(String(e)));
page.on('console', (m) => {
  if (m.type() === 'error') errors.push(m.text());
});
const t = await token(id);
await page.goto(`${BASE}${t.data.editorUrl}`);
await page.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await page.waitForTimeout(1000);

const cardKeys = () => page.$$eval('#org-grid .pdf-org-card', (els) => els.map((e) => e.dataset.key));
const card = (key) => page.locator(`#org-grid .pdf-org-card[data-key="${key}"]`);
const settle = () => page.waitForTimeout(350);

// 1) 모든 도구 → 페이지 정리
await page.click('.pdf-tool-item[data-tool="organize"]');
await page.waitForSelector('#organize:not([hidden])');
await page.waitForFunction(() => document.querySelectorAll('#org-grid .pdf-org-card canvas').length >= 6, null, { timeout: 15000 });
check((await cardKeys()).join() === 'o:0,o:1,o:2,o:3,o:4,o:5', '정리 화면에 6쪽 카드(그림 포함)', (await cardKeys()).join());
check(await page.isDisabled('#org-delete'), '고른 쪽이 없으면 지우기 꺼짐');

// 2) 2쪽 지우기(Delete 키)
await card('o:1').click();
await page.keyboard.press('Delete');
await settle();
check((await cardKeys()).join() === 'o:0,o:2,o:3,o:4,o:5', 'Delete로 2쪽 지움', (await cardKeys()).join());

// 3) 4쪽(o:3) 오른쪽으로 돌리기 — 카드 그림이 가로로
await card('o:3').click();
await page.click('#org-rotate-right');
await page.waitForTimeout(800);
const rotBox = await card('o:3').locator('canvas').boundingBox();
check(rotBox && rotBox.width > rotBox.height, '오른쪽으로 돌린 카드가 가로', rotBox ? `${Math.round(rotBox.width)}×${Math.round(rotBox.height)}` : '');

// 4) 6쪽을 앞으로 두 번
await card('o:5').click();
await page.click('#org-move-up');
await settle();
await page.click('#org-move-up');
await settle();
check((await cardKeys()).join() === 'o:0,o:2,o:5,o:3,o:4', '앞으로 옮기기 두 번', (await cardKeys()).join());

// 5) 1쪽을 끌어 맨 뒤로
const lastBox = await card('o:4').boundingBox();
await card('o:0').dragTo(card('o:4'), { targetPosition: { x: lastBox.width - 4, y: lastBox.height / 2 } });
await settle();
check((await cardKeys()).join() === 'o:2,o:5,o:3,o:4,o:0', '끌어서 맨 뒤로', (await cardKeys()).join());

// 6) 4쪽 뒤에 빈 쪽
await card('o:3').click();
await page.click('#org-blank');
await settle();
const afterBlank = await cardKeys();
check(afterBlank.length === 6 && afterBlank[3].startsWith('b:'), '고른 쪽 뒤에 빈 쪽', afterBlank.join());
const blankKey = afterBlank[3];

// 7) 5쪽(o:4) 뒤에 파일에서 넣기(2쪽, 2쪽에 주석)
await card('o:4').click();
const [chooser] = await Promise.all([page.waitForEvent('filechooser'), page.click('#org-insert-file')]);
await chooser.setFiles(partFile);
await page.waitForFunction(() => document.querySelectorAll('#org-grid .pdf-org-card').length === 8, null, { timeout: 15000 });
const afterPart = await cardKeys();
check(/^p:[^:]+:0$/.test(afterPart[5]) && /^p:[^:]+:1$/.test(afterPart[6]) && afterPart[7] === 'o:0', '파일의 두 쪽이 고른 쪽 뒤에', afterPart.join());
await page.waitForFunction(() => document.querySelectorAll('#org-grid .pdf-org-card canvas').length >= 8, null, { timeout: 15000 });
const partCardHasAnnot = await card(afterPart[6]).locator('.pdf-thumb-annots *').count();
check(partCardHasAnnot > 0, '넣은 파일의 주석도 카드에 보임');

// 8) 실행 취소 → 넣기 전, 다시 실행 → 다시 넣음
await page.keyboard.press('Control+z');
await settle();
const undone = (await cardKeys()).length;
await page.keyboard.press('Control+y');
await settle();
check(undone === 6 && (await cardKeys()).join() === afterPart.join(), 'Ctrl+Z / Ctrl+Y', `${undone} → ${(await cardKeys()).length}`);
await page.screenshot({ path: path.join(shotDir, '1-organize.png') });

// 9) 고른 쪽 추출(2쪽) — 내려받은 PDF
await card('o:2').click();
await card('o:0').click({ modifiers: ['Control'] });
check((await page.textContent('#org-status')).includes('2쪽 고름'), '상태에 고른 쪽 수', await page.textContent('#org-status'));
const [dl1] = await Promise.all([page.waitForEvent('download'), page.click('#org-extract')]);
const extracted = await describePdf(fs.readFileSync(await dl1.path()));
check(dl1.suggestedFilename() === '정리시험-추출.pdf' && extracted.map((p) => p.text).join() === '3쪽,1쪽' && extracted[0].annots.includes('Square'), '고른 쪽 추출(순서·주석)', `${dl1.suggestedFilename()} ${JSON.stringify(extracted)}`);

// 10) 나누기 — 3쪽씩 → zip 3개
await page.click('#org-split-btn');
await page.fill('#org-split-n', '3');
const [dl2] = await Promise.all([page.waitForEvent('download'), page.click('#org-split-go')]);
const zip = unzipSync(fs.readFileSync(await dl2.path()));
const names = Object.keys(zip);
const firstPart = await describePdf(zip[names[0]]);
check(dl2.suggestedFilename() === '정리시험-나누기.zip' && names.length === 3 && firstPart.map((p) => p.text).join() === '3쪽,6쪽,4쪽', '3쪽씩 나누기(zip)', `${names.join(',')} ${firstPart.map((p) => p.text).join()}`);

// 11) 닫기 → 본문·쪽 수·썸네일·메모 목록이 새 순서
await page.keyboard.press('Escape');
await page.waitForSelector('#organize[hidden]', { state: 'attached' });
await page.waitForTimeout(800);
const mainKeys = await page.$$eval('#pdf-pages .pdf-page', (els) => els.map((e) => e.dataset.key));
check(mainKeys.join() === afterPart.join(), '본문 쪽 순서가 정리 결과와 같음', mainKeys.join());
check((await page.textContent('#page-count')).trim() === '/ 8', '쪽 수 표시 8', await page.textContent('#page-count'));
check((await page.$$('#thumbs .pdf-thumb')).length === 8, '썸네일 8개');
// 정리 전 보던 쪽(원래 1쪽)을 계속 보고 있다 — 그 쪽은 이제 8쪽
check((await page.inputValue('#page-input')) === '8', '보던 쪽(원래 1쪽 → 지금 8쪽)을 계속 보여 줌', await page.inputValue('#page-input'));
await page.click('#thumbs .pdf-thumb[data-page="1"]');
await page.waitForFunction(() => (document.querySelector('.pdf-page[data-page="1"] .pdf-text-layer')?.textContent ?? '').length > 0, null, { timeout: 10000 });
const firstText = await page.locator('.pdf-page[data-page="1"] .pdf-text-layer').textContent();
check(firstText.includes('3쪽'), '본문 1쪽은 원래 3쪽', firstText);
await page.click('.pdf-mode[data-mode="edit"]');
await page.waitForTimeout(400);
const pageHeads = await page.$$eval('#comments .pdf-comment-page', (els) => els.map((e) => e.textContent));
check(pageHeads.join() === '1쪽,7쪽', '메모 목록 쪽 번호가 새 순서(원래 3쪽 주석 → 1쪽, 넣은 파일 주석 → 7쪽)', pageHeads.join());
// 돌린 쪽은 본문에서도 가로
const rotated = await page.locator('.pdf-page[data-key="o:3"]').boundingBox();
check(rotated.width > rotated.height, '본문에서도 돌린 쪽이 가로');
await page.screenshot({ path: path.join(shotDir, '2-main.png') });

// 12) 다른 참여자 화면도 같은 순서
const other = await ctx.newPage();
await other.goto(`${BASE}${(await token(id, { viewer: '다른 사람' })).data.editorUrl}`);
await other.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await other.waitForTimeout(1500);
const otherKeys = await other.$$eval('#pdf-pages .pdf-page', (els) => els.map((e) => e.dataset.key));
check(otherKeys.join() === afterPart.join(), '다른 참여자 화면도 같은 순서', otherKeys.join());
await other.close();

// 13) 내려받은 파일(편집 중 상태) — 순서·회전·빈 쪽·주석
const dlBytes = new Uint8Array(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(t.data.token)}`)).arrayBuffer());
const saved = await describePdf(dlBytes);
const expected = ['3쪽', '6쪽', '4쪽', '', '5쪽', '부록 1', '부록 2', '1쪽'];
check(saved.map((p) => p.text).join() === expected.join(), '저장 파일 쪽 순서(빈 쪽·넣은 쪽 포함)', saved.map((p) => p.text || '□').join(','));
check(saved[2].rotate === 90 && saved.filter((p) => p.rotate !== 0).length === 1, '저장 파일 회전(4쪽만 90°)');
check(saved[0].annots.includes('Square') && saved[6].annots.includes('Square') && saved.reduce((n, p) => n + p.annots.length, 0) === 2, '저장 파일 주석이 그 쪽을 따라감');

// 14) 변경 기록
const changes = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data.changes ?? [];
const pc = changes.find((c) => c.kind === 'pages_changed');
check(pc && pc.text === '1쪽 지움 · 3쪽 넣음 · 1쪽 돌림 · 순서 바꿈' && pc.viewer === '시험 편집자', '변경 기록에 쪽 정리 요약', pc ? `${pc.text} / ${pc.viewer}` : '없음');

// 15) 모두 나가면 저장 → 다시 열면 그 순서가 원본(o:0…)
await page.goto('about:blank');
await new Promise((r) => setTimeout(r, 2500));
const page2 = await ctx.newPage();
await page2.goto(`${BASE}${(await token(id)).data.editorUrl}`);
await page2.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await page2.waitForTimeout(1200);
const reopenedKeys = await page2.$$eval('#pdf-pages .pdf-page', (els) => els.map((e) => e.dataset.key));
const reopenedText = await page2.locator('.pdf-page[data-page="1"] .pdf-text-layer').textContent();
check(reopenedKeys.length === 8 && reopenedKeys.every((k, i) => k === `o:${i}`) && reopenedText.includes('3쪽'), '저장 뒤 다시 열면 정리된 순서가 새 원본', `${reopenedKeys.length} ${reopenedText}`);
const reopenedAnnots = await page2.$$eval('#pdf-pages [data-id]', (els) => new Set(els.map((e) => e.dataset.id)).size);
check(reopenedAnnots >= 1, '다시 열어도 주석이 편집기 층에', String(reopenedAnnots));
await page2.close();

// 16) 읽기 전용 — 정리·합치기는 없고, 추출은 있음(편집 단추 숨김)
const ro = await ctx.newPage();
await ro.goto(`${BASE}${(await token(id, { viewer: '읽기만', canEdit: false })).data.editorUrl}`);
await ro.waitForSelector('#pdf-loading[hidden]', { state: 'attached', timeout: 30000 });
await ro.waitForTimeout(800);
const roOrganize = await ro.isVisible('.pdf-tool-item[data-tool="organize"]');
const roExtract = await ro.isVisible('.pdf-tool-item[data-tool="extract"]');
await ro.click('.pdf-tool-item[data-tool="extract"]');
await ro.waitForSelector('#organize:not([hidden])');
check(!roOrganize && roExtract && !(await ro.isVisible('#org-delete')) && (await ro.isVisible('#org-extract')), '읽기 전용: 정리 없음, 추출만');
const roPart = await fetch(`${BASE}/api/view/pdf-part?token=${encodeURIComponent((await token(id, { canEdit: false })).data.token)}`, { method: 'POST', headers: { 'content-type': 'application/octet-stream' }, body: partBytes });
check(roPart.status === 403, '읽기 전용 토큰으로 파일 넣기 거부', String(roPart.status));
await ro.close();

// 17) 8000(유지보수) 방식 토큰(pdfEditor 없음)은 새 경로를 못 씀
const plain = await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ viewer: 'x', canDownload: true, canEdit: true }) }).then((r) => r.json());
const plainExtract = await fetch(`${BASE}/api/view/pdf-extract?token=${encodeURIComponent(plain.data.token)}`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ groups: [['o:0']] }) });
check(plainExtract.status === 403 && !plain.data.editorUrl, 'pdfEditor 없는 토큰은 추출·편집기 없음', String(plainExtract.status));

check(errors.length === 0, '페이지 오류 없음', errors.slice(0, 3).join(' | '));
await browser.close();
await fetch(`${BASE}/api/documents/${id}`, { method: 'DELETE' });
const passed = results.filter(Boolean).length;
console.log(`\n${passed === results.length ? 'ALL CHECKS PASSED' : 'SOME CHECKS FAILED'} (${passed}/${results.length}) — 스크린샷: ${shotDir}`);
process.exit(passed === results.length ? 0 : 1);
