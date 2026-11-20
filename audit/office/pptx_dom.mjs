// PPT 편집 화면 DOM 점검 — pptx_checks.json의 항목을 슬라이드마다 실제 브라우저 DOM(계산된 스타일)으로 잰다.
// 모양 점수(잉크 겹침)는 그림 자르기 오류도 0.97로 넘겨 버려서, 결함마다 "Office에선 이 값이어야 한다"를 숫자로 확인한다.
// capture.mjs가 편집 화면을 찍을 때 부른다(pptx만). 결과는 manifest의 edit.dom → compare.py가 점검표(편집)에 합친다.
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const CHECKS = JSON.parse(fs.readFileSync(path.join(path.dirname(fileURLToPath(import.meta.url)), 'pptx_checks.json'), 'utf-8')).items;

/** 브라우저 안에서 도는 측정 함수 — 대상 하나(글자·그림·도형)의 보이는 값을 돌려준다. 슬라이드 1pt = --slide-pt px. */
function measure(spec) {
  const canvas = document.querySelector('#slide-canvas');
  const cr = canvas.getBoundingClientRect();
  const ptPx = parseFloat(getComputedStyle(canvas).getPropertyValue('--slide-pt')) || 1;
  // 캔버스까지 거슬러 올라가며 회전·뒤집기를 모은다(transform 행렬의 각도·행렬식 부호)
  const turn = (el) => {
    let deg = 0;
    let flip = false;
    for (let e = el; e && e !== canvas; e = e.parentElement) {
      const t = getComputedStyle(e).transform;
      if (!t || t === 'none') continue;
      const [a, b, c, d] = t.slice(t.indexOf('(') + 1, -1).split(',').map(Number);
      if (a * d - b * c < 0) flip = !flip;
      deg += (Math.atan2(b, a) * 180) / Math.PI;
    }
    return { rotate: Math.round((((deg % 360) + 540) % 360) - 180), flip };
  };
  if (spec.text) {
    const walker = document.createTreeWalker(canvas, NodeFilter.SHOW_TEXT);
    const hits = [];
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      for (let i = n.data.indexOf(spec.text); i >= 0; i = n.data.indexOf(spec.text, i + 1)) hits.push([n, i]);
    }
    const hit = hits[spec.nth ?? 0];
    if (!hit) return { found: false };
    const [node, i] = hit;
    const range = document.createRange();
    range.setStart(node, i);
    range.setEnd(node, i + spec.text.length);
    const rects = [...range.getClientRects()].filter((r) => r.width > 0);
    const el = node.parentElement;
    const cs = getComputedStyle(el);
    const p = el.closest('p') ?? el;
    const ps = getComputedStyle(p);
    const before = getComputedStyle(p, '::before');
    const box = el.closest('.slide-textbox');
    let underline = false;
    for (let e = el; e && e !== canvas; e = e.parentElement) if (getComputedStyle(e).textDecorationLine.includes('underline')) underline = true;
    const tops = new Set(rects.map((r) => Math.round(r.top)));
    const fs = parseFloat(ps.fontSize);
    return {
      found: true,
      fontPt: Math.round((parseFloat(cs.fontSize) / ptPx) * 10) / 10,
      family: cs.fontFamily,
      color: cs.color,
      weight: Number(cs.fontWeight),
      underline,
      lineHeightRatio: ps.lineHeight === 'normal' ? null : Math.round((parseFloat(ps.lineHeight) / fs) * 100) / 100,
      marginTopPt: Math.round((parseFloat(ps.marginTop) / ptPx) * 10) / 10,
      marginBottomPt: Math.round((parseFloat(ps.marginBottom) / ptPx) * 10) / 10,
      textAlign: ps.textAlign,
      bullet: before.content,
      bulletColor: before.color,
      lines: tops.size,
      // 글자가 슬라이드 안에 보이는가(아래·오른쪽 끝, 0~1 = 슬라이드 높이·폭 비율)
      bottom: rects.length ? Math.round(((Math.max(...rects.map((r) => r.bottom)) - cr.top) / cr.height) * 100) / 100 : null,
      // 글자 시작 위치가 상자 왼쪽에서 얼마나(상자 폭 비율) — 가운데/왼쪽 정렬 구분용
      leftInBox: box && rects.length ? Math.round(((rects[0].left - box.getBoundingClientRect().left) / box.getBoundingClientRect().width) * 100) / 100 : null,
      writingMode: getComputedStyle(el.closest('.ProseMirror') ?? p).writingMode,
      wordBreak: ps.wordBreak,
      ...turn(el),
    };
  }
  if (spec.image !== undefined) {
    const holder = [...canvas.querySelectorAll('.slide-visual-image')].filter((e) => !e.dataset.background)[spec.image];
    const img = holder?.querySelector('img');
    if (!img) return { found: false };
    // transform 영향 없는 레이아웃 값(offset*)으로 잘린 비율을 잰다
    const W = img.offsetWidth;
    const H = img.offsetHeight;
    const r3 = (v) => Math.round(v * 1000) / 1000;
    return {
      found: true,
      cropL: r3(-img.offsetLeft / W),
      cropT: r3(-img.offsetTop / H),
      cropR: r3((img.offsetLeft + W - holder.clientWidth) / W),
      cropB: r3((img.offsetTop + H - holder.clientHeight) / H),
      ...turn(img),
    };
  }
  if (spec.shape) {
    const el = [...canvas.querySelectorAll('.slide-visual-shape')].filter((e) => e.dataset.geom === spec.shape)[spec.nth ?? 0];
    if (!el) return { found: false };
    const fillEl = el.querySelector('.slide-shape-fill') ?? el;
    const cs = getComputedStyle(fillEl);
    let shadow = false;
    for (let e = fillEl; e && e !== canvas; e = e.parentElement) {
      const s = getComputedStyle(e);
      if ((s.filter && s.filter.includes('drop-shadow')) || (s.boxShadow && s.boxShadow !== 'none')) shadow = true;
    }
    return { found: true, clip: cs.clipPath !== 'none', gradient: cs.backgroundImage.includes('gradient'), shadow, ...turn(fillEl) };
  }
  if (spec.background) {
    const el = canvas.querySelector('[data-background]');
    return el ? { found: true, color: getComputedStyle(el).backgroundColor } : { found: false };
  }
  return { found: false };
}

function judge(got, expect) {
  if (!got.found) return false;
  return Object.entries(expect).every(([k, want]) => {
    const v = got[k];
    if (want && typeof want === 'object' && !Array.isArray(want)) {
      if ('min' in want && !(typeof v === 'number' && v >= want.min)) return false;
      if ('max' in want && !(typeof v === 'number' && v <= want.max)) return false;
      if ('includes' in want && !String(v).includes(want.includes)) return false;
      if ('not' in want && String(v) === String(want.not)) return false;
      return true;
    }
    return v === want;
  });
}

/** 문서 name의 slide번째 슬라이드에 걸린 점검 항목을 지금 화면에서 잰다. */
export async function runPptxDomChecks(page, name, slide) {
  const out = [];
  for (const c of CHECKS.filter((it) => it.doc === name && it.slide === slide)) {
    const got = await page.evaluate(measure, c.target).catch((e) => ({ found: false, error: String(e).slice(0, 120) }));
    out.push({ id: c.id, doc: name, slide, item: c.item, ok: judge(got, c.expect), got });
  }
  return out;
}
