// 워드 편집기 입력 속도 재기(실제 브라우저) — 큰 문서(검토 재현 F-big 150장·G-big400 400장, python-docx로 만든 합성 문서)를 열고
// 본문 가운데 문단에 글자 20개를 쳐서 글자마다 걸린 시간을 잰다:
//   처리 = 키 입력 이벤트 처리 시간(PerformanceEventTiming processingEnd-processingStart, 우리 코드·ProseMirror),
//   화면까지 = 키를 누른 뒤 다음 그림까지(event duration, INP와 같은 값 — 브라우저 다시 배치·그리기 포함),
//   긴 작업 = 50ms 넘는 작업(longtask) 수·최대.
// CPU 4배 느리게(저사양 PC 흉내, 문서 열기부터 느리게)도 같이 잰다. 첫 화면 = 편집기 주소를 연 뒤 본문 첫 문단이 보일 때까지,
// 쪽 나눔 = 첫 넘침 쪽 경계(.docx-page-boundary)가 생길 때까지. 사용: node audit/office/perf-docx-typing.mjs [http://127.0.0.1:8143]
import fs from 'node:fs';
import path from 'node:path';
import { chromium } from 'playwright';

const BASE = process.argv[2] ?? 'http://127.0.0.1:8143';
const FILES = ['audit/office/corpus/docx/d02-numbering.docx', 'tests/fixtures/docx-review/F-big.docx', 'tests/fixtures/docx-review/G-big400.docx'];

async function upload(file) {
  const form = new FormData();
  form.append('file', new Blob([fs.readFileSync(file)]), path.basename(file));
  const id = (await (await fetch(`${BASE}/api/documents`, { method: 'POST', body: form })).json()).data.id;
  for (let i = 0; ; i += 1) {
    const st = (await (await fetch(`${BASE}/api/documents/${id}`)).json()).data;
    if (st.status === 'ready') break;
    if (st.status === 'failed' || i > 1200) throw new Error(`변환 실패 ${file}`);
    await new Promise((r) => setTimeout(r, 300));
  }
  const body = { viewer: 'perf', canPrint: true, canDownload: true, canCopy: true, canEdit: true };
  return (await (await fetch(`${BASE}/api/documents/${id}/tokens`, { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(body) })).json()).data;
}

const pct = (xs, p) => (xs.length ? [...xs].sort((a, b) => a - b)[Math.min(xs.length - 1, Math.floor(xs.length * p))] : 0);
const browser = await chromium.launch();
const rows = [];
try {
  for (const file of FILES) {
    const tok = await upload(file);
    for (const throttle of [1, 4]) {
      const page = await browser.newPage({ viewport: { width: 1400, height: 1000 } });
      const cdp = await page.context().newCDPSession(page);
      await cdp.send('Emulation.setCPUThrottlingRate', { rate: throttle });
      const opened = Date.now();
      await page.goto(`${BASE}${tok.editorUrl}&home=hidden`);
      await page.waitForSelector('.ProseMirror > p', { timeout: 120000 });
      const firstScreen = Date.now() - opened;
      await page.waitForSelector('.docx-page-boundary', { timeout: 120000 }).catch(() => undefined);
      const paged = Date.now() - opened;
      await page.waitForTimeout(4000 * throttle); // 첫 동기화·넘침 쪽 나눔·글꼴
      await cdp.send('Emulation.setCPUThrottlingRate', { rate: 1 });
      const stats = await page.evaluate(() => ({ dom: document.querySelectorAll('.ProseMirror *').length, tabs: document.querySelectorAll('.docx-tab').length, nums: document.querySelectorAll('.docx-num').length, pages: document.querySelectorAll('.docx-page-boundary').length + 1 }));
      const paragraphs = page.locator('.ProseMirror > p');
      const n = await paragraphs.count();
      await paragraphs.nth(Math.floor(n / 2)).click();
      await page.keyboard.press('End');
      await page.waitForTimeout(500);
      await cdp.send('Emulation.setCPUThrottlingRate', { rate: throttle });
      await page.evaluate(() => {
        window.__perf = { events: [], longtasks: [] };
        new PerformanceObserver((list) => {
          for (const e of list.getEntries()) if (e.name === 'keydown' || e.name === 'keypress' || e.name === 'input' || e.name === 'beforeinput') window.__perf.events.push({ name: e.name, duration: e.duration, processing: e.processingEnd - e.processingStart, id: e.interactionId });
        }).observe({ type: 'event', buffered: false, durationThreshold: 0 });
        new PerformanceObserver((list) => {
          for (const e of list.getEntries()) window.__perf.longtasks.push(e.duration);
        }).observe({ type: 'longtask', buffered: false });
      });
      const t0 = Date.now();
      for (const ch of '가나다라마바사아자차abcdefghij') {
        await page.keyboard.type(ch);
        await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r))));
      }
      const wall = Date.now() - t0;
      await page.waitForTimeout(1000);
      const perf = await page.evaluate(() => window.__perf);
      await cdp.send('Emulation.setCPUThrottlingRate', { rate: 1 });
      // 글자 하나 = 같은 interactionId의 이벤트들 — 처리 시간은 합, 화면까지는 최대
      const byKey = new Map();
      for (const e of perf.events) {
        const k = byKey.get(e.id) ?? { processing: 0, duration: 0 };
        k.processing += e.processing;
        k.duration = Math.max(k.duration, e.duration);
        byKey.set(e.id, k);
      }
      const keys = [...byKey.values()];
      const row = {
        file: path.basename(file), throttle: `${throttle}x`, ...stats, firstScreen, paged, keys: keys.length,
        processingMedian: Math.round(pct(keys.map((k) => k.processing), 0.5) * 10) / 10,
        processingP95: Math.round(pct(keys.map((k) => k.processing), 0.95) * 10) / 10,
        toPaintMedian: Math.round(pct(keys.map((k) => k.duration), 0.5)),
        toPaintP95: Math.round(pct(keys.map((k) => k.duration), 0.95)),
        longtasks: perf.longtasks.length, longtaskMax: Math.round(Math.max(0, ...perf.longtasks)),
        wallPer20: wall,
      };
      rows.push(row);
      console.log(JSON.stringify(row));
      await page.close();
    }
  }
} finally {
  await browser.close();
}
console.log('\n| 문서 | CPU | DOM 요소 | 탭/번호 | 쪽 | 첫 화면 / 쪽 나눔 ms | 처리(중앙/95%) ms | 화면까지(중앙/95%) ms | 긴 작업(수/최대 ms) |');
console.log('|---|---|---|---|---|---|---|---|---|');
for (const r of rows) console.log(`| ${r.file} | ${r.throttle} | ${r.dom} | ${r.tabs}/${r.nums} | ${r.pages} | ${r.firstScreen} / ${r.paged} | ${r.processingMedian} / ${r.processingP95} | ${r.toPaintMedian} / ${r.toPaintP95} | ${r.longtasks} / ${r.longtaskMax} |`);
