/**
 * PDF 그림 XObject 풀기 — 가리기(검게 칠하기)와 압축(줄여 다시 담기)이 함께 쓴다.
 * 다루는 형식: DCT(JPEG, sharp로 풀기)와 Flate/무압축 8비트 회색·RGB·CMYK·ICCBased(PNG 예측자 포함).
 * 그 밖(1비트·색표·JPX·Decode 배열·마스크)은 null — 부르는 쪽이 정한다(가리기는 그림을 빼고, 압축은 그대로 둔다).
 */
import sharp from 'sharp';
import { decodePDFRawStream, PDFArray, PDFDict, PDFName, PDFNumber, PDFRawStream, PDFRef, type PDFContext, type PDFObject } from 'pdf-lib';

export interface DecodedImage {
  readonly raw: Uint8Array;
  readonly width: number;
  readonly height: number;
  /** 1 회색 · 3 RGB · 4 CMYK */
  readonly channels: number;
  /** 다시 담을 때 쓸 색 공간(JPEG를 풀면 DeviceGray/DeviceRGB) */
  readonly colorSpace: PDFObject;
  readonly wasJpeg: boolean;
}

function lookup(ctx: PDFContext, o: PDFObject | undefined): PDFObject | undefined {
  return o instanceof PDFRef ? ctx.lookup(o) : o;
}

export function filtersOf(ctx: PDFContext, stream: PDFRawStream): string[] {
  const f = lookup(ctx, stream.dict.get(PDFName.of('Filter')));
  if (f instanceof PDFArray) return f.asArray().map((x) => (lookup(ctx, x) as PDFName | undefined)?.decodeText?.() ?? '');
  return f instanceof PDFName ? [f.decodeText()] : [];
}

export function numberOf(ctx: PDFContext, dict: PDFDict, key: string): number | undefined {
  const v = lookup(ctx, dict.get(PDFName.of(key)));
  return v instanceof PDFNumber ? v.asNumber() : undefined;
}

function pngUnpredict(data: Uint8Array, columns: number, bpp: number): Uint8Array | null {
  const rowLen = columns * bpp;
  const rows = Math.floor(data.length / (rowLen + 1));
  const out = new Uint8Array(rows * rowLen);
  for (let y = 0; y < rows; y += 1) {
    const type = data[y * (rowLen + 1)];
    const src = data.subarray(y * (rowLen + 1) + 1, (y + 1) * (rowLen + 1));
    const cur = out.subarray(y * rowLen, (y + 1) * rowLen);
    const prev = y > 0 ? out.subarray((y - 1) * rowLen, y * rowLen) : new Uint8Array(rowLen);
    for (let i = 0; i < rowLen; i += 1) {
      const left = i >= bpp ? cur[i - bpp] : 0;
      const up = prev[i];
      const ul = i >= bpp ? prev[i - bpp] : 0;
      let v = src[i];
      if (type === 1) v += left;
      else if (type === 2) v += up;
      else if (type === 3) v += (left + up) >> 1;
      else if (type === 4) {
        const p = left + up - ul;
        const pa = Math.abs(p - left);
        const pb = Math.abs(p - up);
        const pc = Math.abs(p - ul);
        v += pa <= pb && pa <= pc ? left : pb <= pc ? up : ul;
      } else if (type !== 0) return null;
      cur[i] = v & 0xff;
    }
  }
  return out;
}

export async function decodeImage(ctx: PDFContext, stream: PDFObject | undefined): Promise<DecodedImage | null> {
  if (!(stream instanceof PDFRawStream)) return null;
  const d = stream.dict;
  const W = numberOf(ctx, d, 'Width') ?? 0;
  const H = numberOf(ctx, d, 'Height') ?? 0;
  if (!W || !H || d.get(PDFName.of('ImageMask')) || d.has(PDFName.of('Decode'))) return null;
  const filters = filtersOf(ctx, stream);
  if (filters.length === 1 && filters[0] === 'DCTDecode') {
    const contents = stream.getContents();
    // JPEG(FF D8)만 sharp에 — 다른 형식 바이트가 다른 해석기로 가지 않게
    if (contents[0] !== 0xff || contents[1] !== 0xd8) return null;
    const { data, info } = await sharp(Buffer.from(contents)).raw().toBuffer({ resolveWithObject: true });
    if (info.width !== W || info.height !== H || (info.channels !== 1 && info.channels !== 3)) return null;
    return { raw: new Uint8Array(data), width: W, height: H, channels: info.channels, colorSpace: PDFName.of(info.channels === 1 ? 'DeviceGray' : 'DeviceRGB'), wasJpeg: true };
  }
  if (!(filters.length === 0 || (filters.length === 1 && filters[0] === 'FlateDecode'))) return null;
  if ((numberOf(ctx, d, 'BitsPerComponent') ?? 8) !== 8) return null;
  const cs = lookup(ctx, d.get(PDFName.of('ColorSpace')));
  const csName = cs instanceof PDFName ? cs.decodeText() : cs instanceof PDFArray ? (lookup(ctx, cs.get(0)) as PDFName | undefined)?.decodeText?.() : undefined;
  let channels = 0;
  if (csName === 'DeviceGray' || csName === 'CalGray') channels = 1;
  else if (csName === 'DeviceRGB' || csName === 'CalRGB') channels = 3;
  else if (csName === 'DeviceCMYK') channels = 4;
  else if (csName === 'ICCBased' && cs instanceof PDFArray) {
    const icc = lookup(ctx, cs.get(1));
    const n = icc instanceof PDFRawStream ? lookup(ctx, icc.dict.get(PDFName.of('N'))) : undefined;
    channels = n instanceof PDFNumber ? n.asNumber() : 0;
  }
  if (![1, 3, 4].includes(channels)) return null;
  let data = filters.length ? decodePDFRawStream(stream).decode() : stream.getContents();
  const parms = lookup(ctx, d.get(PDFName.of('DecodeParms')));
  const predictor = parms instanceof PDFDict ? numberOf(ctx, parms, 'Predictor') ?? 1 : 1;
  if (predictor >= 10) {
    const un = pngUnpredict(data, W, channels);
    if (!un) return null;
    data = un;
  } else if (predictor !== 1) return null;
  if (data.length < W * H * channels) return null;
  return { raw: new Uint8Array(data.subarray(0, W * H * channels)), width: W, height: H, channels, colorSpace: d.get(PDFName.of('ColorSpace'))!, wasJpeg: false };
}
