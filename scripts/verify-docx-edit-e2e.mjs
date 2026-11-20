// 워드 편집기 손실 없는 저장 회귀 검사(실제 브라우저) — 편집(잘라 붙이기·되돌리기·번호 단추·Enter·두 사람 동시 편집)을 하고
// Ctrl+S로 저장한 뒤 내려받은 docx 안의 요소 개수를 원본과 비교한다. 편집 없는 왕복은 audit/office/roundtrip-docx.ts가 본다.
//
// 사용법(서버를 먼저 띄워 둘 것 — 업로드·토큰 API가 키 없이 열려 있는 로컬 시험 서버):
//   node scripts/verify-docx-edit-e2e.mjs [http://127.0.0.1:8143]
// 시험 문서는 합성 말뭉치(audit/office/corpus/docx)와 검토 재현 문서(tests/fixtures/docx-review)만 쓴다.
import fs from 'node:fs';
import path from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8143';
const CORPUS = 'audit/office/corpus/docx';
const results = [];
const rec = (name, ok, detail = '') => {
  results.push({ name, ok, detail });
  console.log(`${ok ? 'OK  ' : 'FAIL'} ${name}${detail ? ` — ${detail}` : ''}`);
};

async function upload(file) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
  const res = await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json();
  const id = res.data.id;
  for (let i = 0; ; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data;
    if (st.status === 'ready') break;
    if (st.status === 'failed' || i > 600) throw new Error(`변환 실패 ${file}`);
    await new Promise((r) => setTimeout(r, 300));
  }
  const body = { viewer: 'docx-e2e', canPrint: true, canDownload: true, canCopy: true, canEdit: true };
  return (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) })).json()).data;
}

const documentXmlOf = (bytes) => strFromU8(unzipSync(new Uint8Array(bytes))['word/document.xml']);
const count = (xml, tag) => (xml.match(new RegExp(`<${tag}[\\s/>]`, 'g')) ?? []).length;

async function openEditor(context, tok) {
  const page = await context.newPage();
  page.on('pageerror', (e) => rec(`페이지 오류 없음`, false, String(e).slice(0, 200)));
  await page.goto(`${BASE}${tok.editorUrl}&home=hidden`);
  await page.waitForSelector('.ProseMirror');
  await page.waitForFunction(() => document.querySelector('#conn-status')?.getAttribute('data-status') === 'synced', null, { timeout: 30000 }).catch(() => {});
  await page.waitForTimeout(1500);
  return page;
}

/** Ctrl+S → "저장됨"이 새로 뜰 때까지 → 내려받은 document.xml */
async function saveAndDownload(page, tok) {
  await page.evaluate(() => {
    const el = document.querySelector('#save-status');
    if (el) el.dataset.e2eBefore = el.textContent ?? '';
  });
  await page.keyboard.press('Control+s');
  await page.waitForTimeout(500);
  await page.waitForFunction(() => /저장됨/.test(document.querySelector('#save-status')?.textContent ?? ''), null, { timeout: 30000 });
  await page.waitForTimeout(500);
  return documentXmlOf(await (await fetch(`${BASE}/api/view/download?token=${encodeURIComponent(tok.token)}`)).arrayBuffer());
}

/** 그 글자를 담은 본문 문단(장식 안 것 제외)을 눌러 커서를 둔다 */
async function clickParagraph(page, text, where = 'end') {
  const target = page.locator('.ProseMirror p', { hasText: text }).first();
  await target.click();
  await page.waitForTimeout(200);
  await page.keyboard.press(where === 'end' ? 'End' : 'Home');
  await page.waitForTimeout(300); // ProseMirror가 DOM 선택을 읽을 틈
}

/** 그 글자를 담은 문단을 세 번 눌러(문단 전체 선택) 잘라 문서 끝 새 문단에 붙인다 */
async function cutParagraphToEnd(page, text) {
  await page.locator('.ProseMirror p', { hasText: text }).first().click({ clickCount: 3 });
  await page.waitForTimeout(300);
  await page.keyboard.press('Control+x');
  await page.waitForTimeout(300);
  await page.keyboard.press('Control+End');
  await page.waitForTimeout(200);
  await page.keyboard.press('Enter');
  await page.keyboard.press('Control+v');
  await page.waitForTimeout(800);
}

const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1400, height: 1000 }, permissions: ['clipboard-read', 'clipboard-write'] });
try {
  // C2: 문서 전체 잘라내기 → 붙여넣기, 문단 하나 잘라 옮기기 — 보존 요소 개수가 그대로
  const kinds = [
    ['d04-images.docx', ['w:drawing'], '글자 사이 작은 아이콘'],
    ['a1-korean-report.docx', ['w:drawing'], '그림'],
    ['d02-numbering.docx', ['w:numPr'], '추진 배경'],
    ['d06-columns-footnotes.docx', ['w:footnoteReference', 'w:sectPr'], '단락 1:'],
    ['d07-tabs-textbox-breaks.docx', ['w:txbxContent', 'w:tabs'], '글상자가 오른쪽에'],
  ];
  for (const [file, tags, paragraph] of kinds) {
    const original = documentXmlOf(fs.readFileSync(path.join(CORPUS, file)));
    for (const mode of ['전체 잘라 붙이기', '문단 잘라 옮기기']) {
      const tok = await upload(path.join(CORPUS, file));
      const page = await openEditor(context, tok);
      if (mode === '전체 잘라 붙이기') {
        await page.locator('.ProseMirror').click();
        await page.keyboard.press('Control+a');
        await page.keyboard.press('Control+x');
        await page.waitForTimeout(300);
        await page.keyboard.press('Control+v');
        await page.waitForTimeout(800);
      } else {
        await cutParagraphToEnd(page, paragraph);
      }
      const saved = await saveAndDownload(page, tok);
      for (const tag of tags) rec(`C2 ${file} ${mode}: <${tag}>`, count(saved, tag) === count(original, tag), `${count(original, tag)} → ${count(saved, tag)}`);
      await page.close();
    }
  }

  // C3: 그림 지움 → 저장 → 되돌리기 → 저장 → 그림 3개
  {
    const tok = await upload(path.join(CORPUS, 'd04-images.docx'));
    const page = await openEditor(context, tok);
    await page.locator('.ProseMirror img.docx-img').first().click();
    await page.keyboard.press('Delete');
    const afterDelete = await saveAndDownload(page, tok);
    await page.keyboard.press('Control+z');
    await page.waitForTimeout(500);
    const onScreen = await page.locator('.ProseMirror img.docx-img').count();
    const saved = await saveAndDownload(page, tok);
    rec('C3 지움→저장→되돌리기→저장: 그림', count(afterDelete, 'w:drawing') === 2 && onScreen === 3 && count(saved, 'w:drawing') === 3,
      `지운 뒤 ${count(afterDelete, 'w:drawing')}, 되돌린 화면 ${onScreen}, 저장본 ${count(saved, 'w:drawing')}`);
    await page.close();
  }

  // C2/C3 두 사람: A가 그림 문단을 잘라내고 저장된 뒤 다시 붙여넣는다 → B가 저장 → 그림 3개, 같은 문서 안 복사로 둘이 된 그림 번호는 겹치지 않는다
  {
    const tok = await upload(path.join(CORPUS, 'd04-images.docx'));
    const a = await openEditor(context, tok);
    const ctxB = await browser.newContext({ viewport: { width: 1400, height: 1000 } });
    const b = await openEditor(ctxB, tok);
    await a.locator('.ProseMirror img.docx-img').first().click(); // 그림 하나 선택
    await a.waitForTimeout(300);
    await a.keyboard.press('Control+x');
    const cut = await saveAndDownload(a, tok);
    await a.keyboard.press('Control+End');
    await a.keyboard.press('Enter');
    await a.keyboard.press('Control+v');
    await a.keyboard.press('Enter');
    await a.keyboard.press('Control+v');
    await a.waitForTimeout(1500);
    const savedByB = await saveAndDownload(b, tok);
    const ids = [...savedByB.matchAll(/<wp:docPr\b[^>]*?\bid="(\d+)"/g)].map((m) => m[1]);
    rec('두 사람: A 잘라 저장 → 다시 두 번 붙임 → B 저장', count(cut, 'w:drawing') === 2 && count(savedByB, 'w:drawing') === 4 && new Set(ids).size === ids.length,
      `잘라낸 뒤 ${count(cut, 'w:drawing')}, B 저장본 ${count(savedByB, 'w:drawing')}, 그림 번호 ${ids.join(',')}`);
    await a.close();
    await ctxB.close();
  }

  // C2: 다른 문서에서 붙인 그림은 붙이지 않고 알린다(보이는 것 = 저장되는 것)
  {
    const src = await upload(path.join(CORPUS, 'd04-images.docx'));
    const dst = await upload(path.join(CORPUS, 'd01-styles-fonts.docx'));
    const pa = await openEditor(context, src);
    await pa.locator('.ProseMirror').click();
    await pa.keyboard.press('Control+a');
    await pa.keyboard.press('Control+c');
    const pb = await openEditor(context, dst);
    let alerted = false;
    pb.on('dialog', (d) => {
      alerted = true;
      void d.accept();
    });
    await pb.locator('.ProseMirror p').last().click();
    await pb.keyboard.press('Control+End');
    await pb.keyboard.press('Control+v');
    await pb.waitForTimeout(800);
    const dialogShown = alerted || (await pb.locator('.office-dialog, [role="dialog"]').filter({ hasText: '다른 문서' }).count()) > 0;
    const imgs = await pb.locator('.ProseMirror img.docx-img').count();
    const saved = await saveAndDownload(pb, dst);
    rec('다른 문서의 그림 붙여넣기: 빼고 알림', imgs === 0 && count(saved, 'w:drawing') === 0 && dialogShown && saved.includes('글자 사이 작은 아이콘'), `화면 그림 ${imgs}, 알림 ${dialogShown}`);
    await pa.close();
    await pb.close();
  }

  // I3: 가져온 번호 문단에서 번호 단추 → 번호 끔(글자 "1. " 없음), 다시 누르면 Word 번호
  {
    const tok = await upload(path.join(CORPUS, 'd02-numbering.docx'));
    const page = await openEditor(context, tok);
    await clickParagraph(page, '사업 개요');
    await page.locator('#btn-ordered-list').click();
    await page.waitForTimeout(300);
    const off = await saveAndDownload(page, tok);
    await clickParagraph(page, '사업 개요');
    await page.locator('#btn-ordered-list').click();
    await page.waitForTimeout(300);
    const on = await saveAndDownload(page, tok);
    const label = await page.locator('.ProseMirror p', { hasText: '사업 개요' }).first().innerText();
    rec('I3 번호 단추: 끄기 → 켜기', count(off, 'w:numPr') === 17 && count(on, 'w:numPr') === 18 && !/<w:t[^>]*>1\. /.test(on) && label.startsWith('1.'), `끈 뒤 numPr ${count(off, 'w:numPr')}, 켠 뒤 ${count(on, 'w:numPr')}, 화면 "${label.slice(0, 8)}"`);
    await page.close();
  }

  // I4: 구역 끝 문단에서 Enter + 글자 → 새 글이 같은 구역(첫 구역 나눔 앞)에
  {
    const tok = await upload(path.join(CORPUS, 'd05-header-footer-sections.docx'));
    const page = await openEditor(context, tok);
    // 첫 구역 나눔을 지닌 빈 문단 = "가로 방향 구역" 제목 바로 앞 문단
    await page.locator('.ProseMirror h1', { hasText: '가로 방향 구역' }).locator('xpath=preceding-sibling::p[1]').click();
    await page.waitForTimeout(200);
    await page.keyboard.press('End');
    await page.waitForTimeout(300);
    await page.keyboard.press('Enter');
    await page.keyboard.type('구역 끝 새 글', { delay: 20 });
    const saved = await saveAndDownload(page, tok);
    // 같은 구역 = 새 글이 든 문단이 첫 구역 나눔을 지닌 문단이거나 그보다 앞(구역 나눔은 그 구역 마지막 문단의 w:pPr 안에 있다)
    const at = saved.indexOf('구역 끝 새 글');
    const paragraphStart = saved.lastIndexOf('<w:p>', at);
    rec('I4 구역 끝 Enter: 같은 구역에 남음', at > 0 && saved.indexOf('<w:sectPr') > paragraphStart && count(saved, 'w:sectPr') === 2,
      `글 문단 시작 ${paragraphStart}, 첫 구역 나눔 ${saved.indexOf('<w:sectPr')}, 구역 ${count(saved, 'w:sectPr')}`);
    await page.close();
  }

  // I2: 변경 내용 추적 문서에서 글자 입력 후 저장 — w:ins·w:del 그대로
  {
    const file = 'tests/fixtures/docx-review/B-tracked.docx';
    const original = documentXmlOf(fs.readFileSync(file));
    const tok = await upload(file);
    const page = await openEditor(context, tok);
    await page.locator('.ProseMirror p').last().click();
    await page.keyboard.press('End');
    await page.keyboard.type(' 추가', { delay: 20 });
    const saved = await saveAndDownload(page, tok);
    rec('I2 변경 내용 추적 유지', count(saved, 'w:ins') === count(original, 'w:ins') && count(saved, 'w:del') === count(original, 'w:del'),
      `ins ${count(original, 'w:ins')}→${count(saved, 'w:ins')}, del ${count(original, 'w:del')}→${count(saved, 'w:del')}`);
    await page.close();
  }

  // I8: 브라우저가 못 그리는 그림(EMF)·차트는 서버가 LibreOffice로 그린 그림으로 보이고, 저장 땐 원본 조각 그대로
  {
    const file = 'tests/fixtures/docx-review/H-emf-chart.docx';
    const tok = await upload(file);
    const page = await openEditor(context, tok);
    await page.waitForFunction(() => [...document.querySelectorAll('.ProseMirror img.docx-img')].every((i) => i.complete), null, { timeout: 60000 }).catch(() => {});
    const imgs = await page.evaluate(() => [...document.querySelectorAll('.ProseMirror img.docx-img')].map((i) => ({ drawn: i.src.includes('docx-drawing'), w: i.naturalWidth })));
    const saved = await saveAndDownload(page, tok);
    const original = documentXmlOf(fs.readFileSync(file));
    rec('I8 EMF·차트 그림으로 표시 + 저장 유지', imgs.filter((i) => i.drawn).length === 2 && imgs.every((i) => i.w > 0) && count(saved, 'c:chart') === count(original, 'c:chart') && count(saved, 'w:drawing') === count(original, 'w:drawing'),
      JSON.stringify(imgs));
    await page.close();
  }

  // N4: 그림 200개(EMF·차트) 문서 — 모두 그려지고(상한에 걸려 회색으로 남지 않음), 그리는 동안 서버가 멈추지 않는다
  {
    const file = 'tests/fixtures/docx-review/J-200-drawings.docx';
    const tok = await upload(file);
    const page = await openEditor(context, tok);
    // 그림을 그리는 동안의 다른 요청 지연만 잰다(문서 열기·변환은 따로)
    let maxLag = 0;
    let pinging = true;
    const pinger = (async () => {
      while (pinging) {
        const t = Date.now();
        await fetch(`${BASE}/api/formats`).catch(() => {});
        maxLag = Math.max(maxLag, Date.now() - t);
        await new Promise((r) => setTimeout(r, 200));
      }
    })();
    const t0 = Date.now();
    const scroll = async () => page.evaluate(() => { for (const img of document.querySelectorAll('.ProseMirror img.docx-img')) img.loading = 'eager'; });
    await scroll();
    await page.waitForFunction(() => [...document.querySelectorAll('.ProseMirror img.docx-img')].every((i) => i.complete), null, { timeout: 180000 }).catch(() => {});
    const imgs = await page.evaluate(() => [...document.querySelectorAll('.ProseMirror img.docx-img')].map((i) => i.naturalWidth));
    pinging = false;
    await pinger;
    rec('N4 그림 200개: 모두 그려짐 + 서버 응답', imgs.length >= 202 && imgs.every((w) => w > 0) && maxLag < 500, `그림 ${imgs.length}개 중 그려짐 ${imgs.filter((w) => w > 0).length}, ${Math.round((Date.now() - t0) / 1000)}초, 다른 요청 최대 지연 ${maxLag}ms`);
    // 같은 그림 다시 요청: 캐시에서 바로(문서를 다시 읽지 않음)
    const ref = await page.evaluate(() => new URL(document.querySelector('.ProseMirror img.docx-img[src*="docx-drawing"]').src).searchParams.get('ref'));
    const t1 = Date.now();
    for (let i = 0; i < 20; i += 1) await fetch(`${BASE}/api/view/docx-drawing?token=${encodeURIComponent(tok.token)}&ref=${encodeURIComponent(ref)}`);
    rec('N4 캐시된 그림 20번', Date.now() - t1 < 2000, `${Date.now() - t1}ms`);
    await page.close();
  }

  // I1: 글꼴 이름으로 CSS를 넣은 문서 — 보통 문단이 빨갛게 되지 않는다
  {
    const tok = await upload('tests/fixtures/docx-review/E-css.docx');
    const page = await openEditor(context, tok);
    const colors = await page.evaluate(() => [...document.querySelectorAll('.ProseMirror p')].map((p) => getComputedStyle(p).color));
    rec('I1 글꼴 이름 CSS 주입 막힘', colors.length > 0 && colors.every((c) => c !== 'rgb(255, 0, 0)'), colors.slice(0, 3).join(' '));
    await page.close();
  }
} finally {
  await browser.close();
}
const failed = results.filter((r) => !r.ok);
console.log(`\n${results.length - failed.length}/${results.length} 통과`);
process.exitCode = failed.length ? 1 : 0;
