/**
 * OCR 글자 층(8단계) — 편집기(브라우저 tesseract.js)가 인식한 낱말을 쪽에 **보이지 않는 글자**(렌더 모드 3)로 쓴다.
 * 그림은 그대로 보이고, 그 위의 글자로 찾기·고르기·복사·변환이 된다. 글꼴은 Pretendard(한글·영문 모두) — 낱말마다
 * 기준선 방향으로 놓고 가로 배율(Tz)로 인식한 낱말 상자 폭에 맞춘다(고르기 영역이 그림의 글자와 겹치게).
 * 원래 내용은 q … Q로 감싼 뒤 덧붙인다(원래 내용이 좌표계를 바꾼 채 끝나도 영향이 없게).
 */
import { PDFName, type PDFDocument, type PDFFont, type PDFPage } from 'pdf-lib';
import type { OcrWord } from './annotTypes.js';

const n = (v: number) => {
  const r = Math.round(v * 1000) / 1000;
  return Object.is(r, -0) ? '0' : String(r);
};

type Rect = readonly [number, number, number, number];

/** 낱말 상자(대략) — 가린 영역과 겹치는지 볼 때 */
export function wordRect(w: OcrWord): Rect {
  const len = Math.hypot(w.ex - w.x, w.ey - w.y) || 1;
  const ux = -(w.ey - w.y) / len;
  const uy = (w.ex - w.x) / len;
  const pts = [
    [w.x, w.y],
    [w.ex, w.ey],
    [w.x + ux * w.h, w.y + uy * w.h],
    [w.ex + ux * w.h, w.ey + uy * w.h],
  ];
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
}

export function writeOcrText(doc: PDFDocument, page: PDFPage, words: readonly OcrWord[], font: PDFFont): number {
  const ops: string[] = [];
  let written = 0;
  for (const w of words) {
    const text = w.t.replace(/[\r\n\t]+/g, ' ').trim();
    if (!text) continue;
    const len = Math.hypot(w.ex - w.x, w.ey - w.y);
    const size = Math.max(1, w.h);
    if (len < 0.5) continue;
    let width: number;
    let encoded: string;
    try {
      width = font.widthOfTextAtSize(text, size);
      encoded = font.encodeText(text).toString();
    } catch {
      continue;
    }
    if (width <= 0) continue;
    const tz = Math.min(1000, Math.max(10, (len / width) * 100));
    const dx = (w.ex - w.x) / len;
    const dy = (w.ey - w.y) / len;
    ops.push(`BT 3 Tr /DvOcr ${n(size)} Tf ${n(tz)} Tz ${n(dx)} ${n(dy)} ${n(-dy)} ${n(dx)} ${n(w.x)} ${n(w.y)} Tm ${encoded} Tj ET`);
    written += 1;
  }
  if (written === 0) return 0;
  const ctx = doc.context;
  page.node.setFontDictionary(PDFName.of('DvOcr'), font.ref);
  const start = ctx.register(ctx.flateStream('q\n'));
  const end = ctx.register(ctx.flateStream('\nQ\n'));
  page.node.wrapContentStreams(start, end);
  page.node.addContentStream(ctx.register(ctx.flateStream(`q\n${ops.join('\n')}\nQ\n`)));
  return written;
}
