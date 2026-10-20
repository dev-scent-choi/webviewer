/**
 * PDF 압축(설계 §3 "압축") — 그림을 쪽에 실제로 그려지는 크기에 맞는 해상도로 줄이고 JPEG로 다시 담은 뒤, 객체 스트림으로 저장한다.
 * 그림이 그려지는 크기는 내용 스트림(contentStream.ts)을 따라가며 Do 때의 변환 행렬로 잰다(Form 안까지). 어디에도 그려지지 않는 그림
 * (주석 모양 안 등)은 건드리지 않는다. 다시 담은 것이 원래보다 크면 원래 것을 둔다. 마스크(SMask)는 그대로 둔다(해상도가 달라도 된다).
 */
import sharp from 'sharp';
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFNumber, PDFRawStream, PDFRef, PDFStream, type PDFContext, type PDFObject } from 'pdf-lib';
import { IDENTITY, mul, parseContent, type Matrix } from './contentStream.js';
import { decodeImage } from './pdfImages.js';
import { streamBytes } from './textRewrite.js';

export type CompressLevel = 'high' | 'medium' | 'low';
export const COMPRESS_LEVELS: Record<CompressLevel, { dpi: number; quality: number; label: string }> = {
  high: { dpi: 200, quality: 85, label: '선명하게(200dpi)' },
  medium: { dpi: 150, quality: 75, label: '보통(150dpi)' },
  low: { dpi: 100, quality: 60, label: '작게(100dpi)' },
};
export const isCompressLevel = (v: unknown): v is CompressLevel => v === 'high' || v === 'medium' || v === 'low';

function lookup(ctx: PDFContext, o: PDFObject | undefined): PDFObject | undefined {
  return o instanceof PDFRef ? ctx.lookup(o) : o;
}

/** 그림 ref → 그려지는 가장 큰 폭·높이(pt) */
function imageUsage(doc: PDFDocument): Map<string, { ref: PDFRef; w: number; h: number }> {
  const ctx = doc.context;
  const usage = new Map<string, { ref: PDFRef; w: number; h: number }>();
  const visit = (bytes: Uint8Array, res: PDFDict | undefined, ctm0: Matrix, depth: number) => {
    let ctm = ctm0;
    const stack: Matrix[] = [];
    const xobjects = res ? lookup(ctx, res.get(PDFName.of('XObject'))) : undefined;
    for (const op of parseContent(bytes)) {
      if (op.op === 'q') stack.push(ctm);
      else if (op.op === 'Q') ctm = stack.pop() ?? ctm;
      else if (op.op === 'cm' && op.operands.length === 6 && op.operands.every((o) => typeof o === 'number')) ctm = mul(op.operands as number[], ctm);
      else if (op.op === 'Do' && xobjects instanceof PDFDict) {
        const name = op.operands[0];
        if (!name || typeof name !== 'object' || !('name' in name)) continue;
        const ref = xobjects.get(PDFName.of(name.name));
        const xo = lookup(ctx, ref);
        if (!(xo instanceof PDFRawStream || xo instanceof PDFStream)) continue;
        const sub = xo.dict.get(PDFName.of('Subtype'));
        if (sub === PDFName.of('Image') && ref instanceof PDFRef) {
          const w = Math.hypot(ctm[0], ctm[1]);
          const h = Math.hypot(ctm[2], ctm[3]);
          const prev = usage.get(ref.tag);
          usage.set(ref.tag, { ref, w: Math.max(prev?.w ?? 0, w), h: Math.max(prev?.h ?? 0, h) });
        } else if (sub === PDFName.of('Form') && depth < 6) {
          const m = lookup(ctx, xo.dict.get(PDFName.of('Matrix')));
          const fm = m instanceof PDFArray ? (m.asArray().map((v) => ((lookup(ctx, v) as PDFNumber | undefined)?.asNumber?.() ?? 0)) as Matrix) : IDENTITY;
          const inner = streamBytes(xo);
          const formRes = lookup(ctx, xo.dict.get(PDFName.of('Resources')));
          if (inner) visit(inner, formRes instanceof PDFDict ? formRes : res, mul(fm, ctm), depth + 1);
        }
      }
    }
  };
  for (const page of doc.getPages()) {
    const contents = lookup(ctx, page.node.get(PDFName.of('Contents')));
    const list = contents instanceof PDFArray ? contents.asArray().map((o) => lookup(ctx, o)) : [contents];
    const parts = list.map((s) => (s ? streamBytes(s) : null)).filter((b): b is Uint8Array => !!b);
    if (parts.length === 0) continue;
    visit(Buffer.concat(parts.flatMap((p) => [p, Uint8Array.of(0x0a)])), page.node.Resources(), IDENTITY, 0);
  }
  return usage;
}

export interface CompressResult {
  readonly bytes: Uint8Array;
  readonly images: number;
}

/** 압축한 바이트 — 그림 수(바꾼 것)도 함께 */
export async function compressPdf(input: Uint8Array, level: CompressLevel): Promise<CompressResult> {
  const { dpi, quality } = COMPRESS_LEVELS[level];
  const doc = await PDFDocument.load(input, { updateMetadata: false });
  const ctx = doc.context;
  let images = 0;
  for (const { ref, w, h } of imageUsage(doc).values()) {
    const stream = ctx.lookup(ref);
    if (!(stream instanceof PDFRawStream)) continue;
    let decoded;
    try {
      decoded = await decodeImage(ctx, stream);
    } catch {
      continue;
    }
    if (!decoded || decoded.channels === 4) continue;
    // 그려지는 크기에서 필요한 화소 수(조금 넉넉히) — 원래보다 크게 키우지는 않는다
    const needW = Math.ceil((w / 72) * dpi);
    const needH = Math.ceil((h / 72) * dpi);
    const ratio = Math.min(1, Math.max(needW / decoded.width, needH / decoded.height));
    const width = Math.max(1, Math.round(decoded.width * ratio));
    const height = Math.max(1, Math.round(decoded.height * ratio));
    if (ratio > 0.9 && decoded.wasJpeg && quality >= 85) continue;
    const jpeg = await sharp(Buffer.from(decoded.raw), { raw: { width: decoded.width, height: decoded.height, channels: decoded.channels as 1 | 3 } })
      .resize(width, height, { fit: 'fill' })
      .jpeg({ quality, mozjpeg: true })
      .toBuffer();
    if (jpeg.length >= stream.getContents().length) continue;
    const entries: Record<string, PDFObject> = {
      Type: PDFName.of('XObject'),
      Subtype: PDFName.of('Image'),
      Width: PDFNumber.of(width),
      Height: PDFNumber.of(height),
      BitsPerComponent: PDFNumber.of(8),
      ColorSpace: PDFName.of(decoded.channels === 1 ? 'DeviceGray' : 'DeviceRGB'),
      Filter: PDFName.of('DCTDecode'),
    };
    const smask = stream.dict.get(PDFName.of('SMask'));
    if (smask) entries.SMask = smask;
    const interp = stream.dict.get(PDFName.of('Interpolate'));
    if (interp) entries.Interpolate = interp;
    ctx.assign(ref, ctx.stream(jpeg, entries));
    images += 1;
  }
  return { bytes: await doc.save({ useObjectStreams: true }), images };
}
