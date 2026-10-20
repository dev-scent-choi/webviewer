/**
 * 내용 스트림의 글자 위치 계산용 글꼴 정보 — 코드 길이(1바이트/2바이트/섞임)와 글자 너비(글리프 단위, 1000 = 글자 크기).
 * 너비를 모르는 경우(표준 14 글꼴인데 Widths가 없음 등)는 표준 글꼴 치수, 그것도 없으면 500으로 어림한다.
 * 위치는 "어느 글자가 고치기·가리기 영역에 드는지"를 가리는 데만 쓰므로 약간의 오차는 괜찮다.
 */
import { Font, FontNames } from '@pdf-lib/standard-fonts';
import { decodePDFRawStream, PDFArray, PDFDict, PDFName, PDFNumber, PDFRawStream, PDFRef, type PDFContext, type PDFObject } from 'pdf-lib';

export interface FontInfo {
  /** 바이트열 → 글자 코드들(각 코드의 바이트) */
  split(bytes: Uint8Array): Uint8Array[];
  /** 코드의 너비(글리프 단위, 1000 = 글자 크기) */
  width(code: Uint8Array): number;
  /** 한 바이트 코드 32(공백)에만 낱말 간격(Tw)이 붙는다 */
  readonly simple: boolean;
  readonly ascent: number;
  readonly descent: number;
}

function lookup(ctx: PDFContext, o: PDFObject | undefined): PDFObject | undefined {
  return o instanceof PDFRef ? ctx.lookup(o) : o;
}
function numOf(ctx: PDFContext, o: PDFObject | undefined): number | undefined {
  const v = lookup(ctx, o);
  return v instanceof PDFNumber ? v.asNumber() : undefined;
}
function nameOf(ctx: PDFContext, o: PDFObject | undefined): string | undefined {
  const v = lookup(ctx, o);
  return v instanceof PDFName ? v.decodeText() : undefined;
}
function arrayOf(ctx: PDFContext, o: PDFObject | undefined): PDFObject[] | undefined {
  const v = lookup(ctx, o);
  return v instanceof PDFArray ? v.asArray() : undefined;
}

const STD: Record<string, FontNames> = {
  Helvetica: FontNames.Helvetica,
  'Helvetica-Bold': FontNames.HelveticaBold,
  'Helvetica-Oblique': FontNames.HelveticaOblique,
  'Helvetica-BoldOblique': FontNames.HelveticaBoldOblique,
  'Times-Roman': FontNames.TimesRoman,
  'Times-Bold': FontNames.TimesRomanBold,
  'Times-Italic': FontNames.TimesRomanItalic,
  'Times-BoldItalic': FontNames.TimesRomanBoldItalic,
  Courier: FontNames.Courier,
  'Courier-Bold': FontNames.CourierBold,
  'Courier-Oblique': FontNames.CourierOblique,
  'Courier-BoldOblique': FontNames.CourierBoldOblique,
  Symbol: FontNames.Symbol,
  ZapfDingbats: FontNames.ZapfDingbats,
};
/** WinAnsi 코드 → 글리프 이름(표준 글꼴 너비 찾기용) — 자주 쓰는 ASCII 범위만 */
const ASCII_GLYPHS = [
  'space', 'exclam', 'quotedbl', 'numbersign', 'dollar', 'percent', 'ampersand', 'quotesingle', 'parenleft', 'parenright', 'asterisk', 'plus', 'comma', 'hyphen', 'period', 'slash',
  'zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'colon', 'semicolon', 'less', 'equal', 'greater', 'question', 'at',
  'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
  'bracketleft', 'backslash', 'bracketright', 'asciicircum', 'underscore', 'grave',
  'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
  'braceleft', 'bar', 'braceright', 'asciitilde',
];
const stdCache = new Map<FontNames, Font>();
function stdWidth(base: string, code: number): number | undefined {
  const name = STD[base.replace(/^[A-Z]{6}\+/, '')];
  if (!name) return undefined;
  let font = stdCache.get(name);
  if (!font) {
    font = Font.load(name);
    stdCache.set(name, font);
  }
  const glyph = code >= 32 && code < 127 ? ASCII_GLYPHS[code - 32] : undefined;
  const w = glyph ? font.getWidthOfGlyph(glyph) : undefined;
  return typeof w === 'number' ? w : undefined;
}

/** CID 글꼴의 /W 배열 → CID별 너비 */
function parseW(ctx: PDFContext, w: PDFObject[] | undefined): Map<number, number> {
  const out = new Map<number, number>();
  if (!w) return out;
  let i = 0;
  while (i < w.length) {
    const first = numOf(ctx, w[i]);
    const next = lookup(ctx, w[i + 1]);
    if (first === undefined) break;
    if (next instanceof PDFArray) {
      next.asArray().forEach((v, k) => {
        const n = numOf(ctx, v);
        if (n !== undefined) out.set(first + k, n);
      });
      i += 2;
    } else {
      const last = numOf(ctx, w[i + 1]);
      const width = numOf(ctx, w[i + 2]);
      if (last === undefined || width === undefined) break;
      for (let c = first; c <= last && c - first < 65536; c += 1) out.set(c, width);
      i += 3;
    }
  }
  return out;
}

/** 내장 CMap의 codespacerange에서 1바이트 범위 — 있으면 섞인 길이 코드 */
function oneByteRanges(ctx: PDFContext, enc: PDFObject | undefined): Array<[number, number]> | null {
  const v = lookup(ctx, enc);
  if (!(v instanceof PDFRawStream)) return null;
  try {
    const text = Buffer.from(decodePDFRawStream(v).decode()).toString('latin1');
    const ranges: Array<[number, number]> = [];
    for (const block of text.matchAll(/begincodespacerange([\s\S]*?)endcodespacerange/g)) {
      for (const m of block[1].matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/g)) {
        if (m[1].length === 2) ranges.push([parseInt(m[1], 16), parseInt(m[2], 16)]);
      }
    }
    return ranges;
  } catch {
    return null;
  }
}

export function fontInfoOf(ctx: PDFContext, fontObj: PDFObject | undefined): FontInfo {
  const font = lookup(ctx, fontObj);
  const fallback: FontInfo = { split: (b) => [...b].map((c) => Uint8Array.of(c)), width: () => 500, simple: true, ascent: 800, descent: -200 };
  if (!(font instanceof PDFDict)) return fallback;
  const subtype = nameOf(ctx, font.get(PDFName.of('Subtype')));
  const base = nameOf(ctx, font.get(PDFName.of('BaseFont'))) ?? '';
  if (subtype === 'Type0') {
    const desc = lookup(ctx, arrayOf(ctx, font.get(PDFName.of('DescendantFonts')))?.[0]);
    const cid = desc instanceof PDFDict ? desc : undefined;
    const widths = cid ? parseW(ctx, arrayOf(ctx, cid.get(PDFName.of('W')))) : new Map<number, number>();
    const dw = (cid && numOf(ctx, cid.get(PDFName.of('DW')))) ?? 1000;
    const fd = cid ? lookup(ctx, cid.get(PDFName.of('FontDescriptor'))) : undefined;
    const encObj = font.get(PDFName.of('Encoding'));
    const encName = nameOf(ctx, encObj) ?? '';
    const identity = /Identity|UCS2|UTF16/.test(encName);
    const singles = identity ? [] : (oneByteRanges(ctx, encObj) ?? (encName && !/UCS2|UTF16/.test(encName) ? [[0, 0x7f] as [number, number]] : []));
    const isSingle = (c: number) => singles.some(([a, b]) => c >= a && c <= b);
    return {
      split(bytes) {
        const out: Uint8Array[] = [];
        let i = 0;
        while (i < bytes.length) {
          const len = !identity && isSingle(bytes[i]) ? 1 : 2;
          out.push(bytes.subarray(i, Math.min(bytes.length, i + len)));
          i += len;
        }
        return out;
      },
      // Identity가 아니면 코드 → CID 표를 모르므로 코드를 그대로 CID로 본다(어림)
      width: (code) => widths.get(code.length === 2 ? (code[0] << 8) | code[1] : code[0]) ?? dw,
      simple: false,
      ascent: (fd instanceof PDFDict && numOf(ctx, fd.get(PDFName.of('Ascent')))) || 880,
      descent: (fd instanceof PDFDict && numOf(ctx, fd.get(PDFName.of('Descent')))) || -120,
    };
  }
  // 단순 글꼴(Type1·TrueType·Type3·MMType1)
  const first = numOf(ctx, font.get(PDFName.of('FirstChar'))) ?? 0;
  const ws = arrayOf(ctx, font.get(PDFName.of('Widths')))?.map((w) => numOf(ctx, w) ?? 0);
  const fd = lookup(ctx, font.get(PDFName.of('FontDescriptor')));
  const missing = fd instanceof PDFDict ? numOf(ctx, fd.get(PDFName.of('MissingWidth'))) : undefined;
  // Type3은 너비가 글리프 공간 단위 — FontMatrix로 1000 단위로 바꾼다
  const fm = subtype === 'Type3' ? arrayOf(ctx, font.get(PDFName.of('FontMatrix')))?.map((v) => numOf(ctx, v) ?? 0) : undefined;
  const scale = fm && fm[0] ? fm[0] * 1000 : 1;
  return {
    split: (b) => [...b].map((c) => Uint8Array.of(c)),
    width(code) {
      const c = code[0];
      const w = ws?.[c - first];
      if (w !== undefined && (w !== 0 || subtype === 'Type3')) return w * scale;
      return stdWidth(base, c) ?? missing ?? 500;
    },
    simple: true,
    ascent: (fd instanceof PDFDict && numOf(ctx, fd.get(PDFName.of('Ascent')))) || 750,
    descent: (fd instanceof PDFDict && numOf(ctx, fd.get(PDFName.of('Descent')))) || -250,
  };
}
