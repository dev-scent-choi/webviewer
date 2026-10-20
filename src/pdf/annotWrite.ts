/**
 * PDF 편집기 모델의 주석을 표준 PDF 주석으로 쓴다 — 다른 뷰어(Acrobat·브라우저·pdf.js)에서도 보이고 고칠 수 있게,
 * 사전과 함께 모양(AP/N) 스트림을 만든다. 모양은 쪽 좌표 그대로 그린다(BBox = Rect, Matrix 없음).
 * 글상자: 글이 표준 글꼴(WinAnsi) 안이면 Helvetica, 한글 등이 있으면 동봉 Pretendard를 통째로 넣는다(pdf-lib+fontkit의
 * subset은 글자가 빠지는 문제가 있다 — rendering.ts와 같은 이유).
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import fontkit from '@pdf-lib/fontkit';
import { PDFDocument, PDFHexString, PDFName, PDFRef, StandardFonts, type PDFContext, type PDFFont, type PDFPage } from 'pdf-lib';
import { origIndexOf, type PdfAnnot } from './annotTypes.js';

/** 프로젝트 루트 assets/fonts — 이 파일 위치 기준이라 실행 cwd와 무관하다(rendering.ts와 같은 방식) */
/** pdf-lib의 사전 리터럴 형식(패키지가 따로 내보내지 않아 flateStream의 매개변수에서 꺼낸다) */
type Dict = NonNullable<Parameters<PDFContext['flateStream']>[1]>;

const PRETENDARD = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..', 'assets', 'fonts', 'Pretendard-Regular.ttf');

const f = (n: number) => (Math.round(n * 1000) / 1000).toString();

function rgbOf(hex: string | null | undefined): [number, number, number] | null {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex ?? '');
  if (!m) return null;
  const v = parseInt(m[1], 16);
  return [((v >> 16) & 255) / 255, ((v >> 8) & 255) / 255, (v & 255) / 255];
}
const colorArray = (hex: string | null | undefined) => rgbOf(hex) ?? [0, 0, 0];
const rg = (hex: string, op: 'rg' | 'RG') => `${colorArray(hex).map(f).join(' ')} ${op}`;

/** ISO → PDF 날짜 "D:YYYYMMDDHHmmSSZ" */
function pdfDate(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  const p = (n: number) => String(n).padStart(2, '0');
  return `D:${d.getUTCFullYear()}${p(d.getUTCMonth() + 1)}${p(d.getUTCDate())}${p(d.getUTCHours())}${p(d.getUTCMinutes())}${p(d.getUTCSeconds())}Z`;
}

/** 모양을 그릴 범위 — 선 굵기·화살촉만큼 넓힌 사각형 */
function boundsOf(a: PdfAnnot): [number, number, number, number] {
  const pad = Math.max(1, a.width) + (a.type === 'arrow' ? Math.max(6, a.width * 4) : 0);
  const pts: number[] = [];
  if (a.quads) pts.push(...a.quads);
  if (a.paths) for (const p of a.paths) pts.push(...p);
  if (a.line) pts.push(...a.line);
  if (pts.length === 0) pts.push(a.rect[0], a.rect[1], a.rect[2], a.rect[3]);
  const xs = pts.filter((_, i) => i % 2 === 0);
  const ys = pts.filter((_, i) => i % 2 === 1);
  if (a.type === 'note' || a.type === 'freetext' || a.type === 'rect' || a.type === 'ellipse') {
    return [a.rect[0], a.rect[1], a.rect[2], a.rect[3]];
  }
  return [Math.min(...xs) - pad, Math.min(...ys) - pad, Math.max(...xs) + pad, Math.max(...ys) + pad];
}

/** 베지어 4개로 사각형 안 타원 */
function ellipsePath(x1: number, y1: number, x2: number, y2: number): string {
  const k = 0.5522847498;
  const cx = (x1 + x2) / 2;
  const cy = (y1 + y2) / 2;
  const rx = (x2 - x1) / 2;
  const ry = (y2 - y1) / 2;
  return [
    `${f(cx + rx)} ${f(cy)} m`,
    `${f(cx + rx)} ${f(cy + ry * k)} ${f(cx + rx * k)} ${f(cy + ry)} ${f(cx)} ${f(cy + ry)} c`,
    `${f(cx - rx * k)} ${f(cy + ry)} ${f(cx - rx)} ${f(cy + ry * k)} ${f(cx - rx)} ${f(cy)} c`,
    `${f(cx - rx)} ${f(cy - ry * k)} ${f(cx - rx * k)} ${f(cy - ry)} ${f(cx)} ${f(cy - ry)} c`,
    `${f(cx + rx * k)} ${f(cy - ry)} ${f(cx + rx)} ${f(cy - ry * k)} ${f(cx + rx)} ${f(cy)} c`,
  ].join('\n');
}

/** 끝점에 열린 화살촉(두 줄) */
function arrowHead(x1: number, y1: number, x2: number, y2: number, width: number): string {
  const len = Math.max(6, width * 4);
  const ang = Math.atan2(y2 - y1, x2 - x1);
  const a1 = ang + Math.PI - Math.PI / 7;
  const a2 = ang + Math.PI + Math.PI / 7;
  return `${f(x2 + len * Math.cos(a1))} ${f(y2 + len * Math.sin(a1))} m ${f(x2)} ${f(y2)} l ${f(x2 + len * Math.cos(a2))} ${f(y2 + len * Math.sin(a2))} l S`;
}

/** 글상자 글을 상자 폭에 맞춰 줄 나누기(글자 단위 — 한글은 띄어쓰기 없이도 나뉜다) */
function wrapText(text: string, font: PDFFont, size: number, maxWidth: number): string[] {
  const lines: string[] = [];
  for (const para of text.split(/\r?\n/)) {
    let line = '';
    for (const ch of para) {
      const next = line + ch;
      if (line && font.widthOfTextAtSize(next, size) > maxWidth) {
        // 영어 단어 중간이면 마지막 띄어쓰기에서 자른다
        const cut = /[A-Za-z0-9]$/.test(line) && /[A-Za-z0-9]/.test(ch) ? line.lastIndexOf(' ') : -1;
        if (cut > 0) {
          lines.push(line.slice(0, cut));
          line = line.slice(cut + 1) + ch;
        } else {
          lines.push(line);
          line = ch;
        }
      } else line = next;
    }
    lines.push(line);
  }
  return lines;
}

const WIN_ANSI = /^[\x20-\x7e\n\r\t]*$/;

interface Fonts {
  standard(): Promise<PDFFont>;
  korean(): Promise<PDFFont>;
}

function fontsFor(doc: PDFDocument): Fonts {
  let std: Promise<PDFFont> | null = null;
  let ko: Promise<PDFFont> | null = null;
  return {
    standard: () => (std ??= doc.embedFont(StandardFonts.Helvetica)),
    korean: () => {
      if (!ko) {
        doc.registerFontkit(fontkit);
        ko = doc.embedFont(fs.readFileSync(PRETENDARD), { subset: false });
      }
      return ko;
    },
  };
}

/** 모양 스트림 내용과 필요한 리소스 */
async function appearanceOf(a: PdfAnnot, fonts: Fonts): Promise<{ content: string; font?: PDFFont; multiply?: boolean }> {
  const w = Math.max(0.5, a.width);
  const lines: string[] = [];
  switch (a.type) {
    case 'highlight': {
      const q = a.quads ?? [];
      lines.push(rg(a.color, 'rg'));
      for (let i = 0; i + 7 < q.length; i += 8) {
        lines.push(`${f(q[i])} ${f(q[i + 1])} m ${f(q[i + 2])} ${f(q[i + 3])} l ${f(q[i + 6])} ${f(q[i + 7])} l ${f(q[i + 4])} ${f(q[i + 5])} l h f`);
      }
      return { content: lines.join('\n'), multiply: true };
    }
    case 'underline':
    case 'strike': {
      const q = a.quads ?? [];
      lines.push(rg(a.color, 'RG'));
      for (let i = 0; i + 7 < q.length; i += 8) {
        const top = (q[i + 1] + q[i + 3]) / 2;
        const bottom = (q[i + 5] + q[i + 7]) / 2;
        const thick = Math.max(0.8, Math.abs(top - bottom) * 0.07);
        const y = a.type === 'underline' ? bottom + thick : (top + bottom) / 2;
        lines.push(`${f(thick)} w ${f(q[i + 4])} ${f(y)} m ${f(q[i + 6])} ${f(y)} l S`);
      }
      return { content: lines.join('\n') };
    }
    case 'ink':
      lines.push(`${rg(a.color, 'RG')} ${f(w)} w 1 J 1 j`);
      for (const p of a.paths ?? []) {
        if (p.length < 2) continue;
        let seg = `${f(p[0])} ${f(p[1])} m`;
        for (let i = 2; i + 1 < p.length; i += 2) seg += ` ${f(p[i])} ${f(p[i + 1])} l`;
        if (p.length === 2) seg += ` ${f(p[0] + 0.01)} ${f(p[1])} l`; // 점 하나도 보이게
        lines.push(`${seg} S`);
      }
      return { content: lines.join('\n') };
    case 'rect': {
      const [x1, y1, x2, y2] = a.rect;
      const fill = a.fill ? `${rg(a.fill, 'rg')} ` : '';
      lines.push(`${fill}${rg(a.color, 'RG')} ${f(w)} w ${f(x1 + w / 2)} ${f(y1 + w / 2)} ${f(x2 - x1 - w)} ${f(y2 - y1 - w)} re ${a.fill ? 'B' : 'S'}`);
      return { content: lines.join('\n') };
    }
    case 'ellipse': {
      const [x1, y1, x2, y2] = a.rect;
      const fill = a.fill ? `${rg(a.fill, 'rg')} ` : '';
      lines.push(`${fill}${rg(a.color, 'RG')} ${f(w)} w`, ellipsePath(x1 + w / 2, y1 + w / 2, x2 - w / 2, y2 - w / 2), a.fill ? 'B' : 'S');
      return { content: lines.join('\n') };
    }
    case 'line':
    case 'arrow': {
      const [x1, y1, x2, y2] = a.line ?? [a.rect[0], a.rect[1], a.rect[2], a.rect[3]];
      lines.push(`${rg(a.color, 'RG')} ${f(w)} w 1 J 1 j ${f(x1)} ${f(y1)} m ${f(x2)} ${f(y2)} l S`);
      if (a.type === 'arrow') lines.push(arrowHead(x1, y1, x2, y2, w));
      return { content: lines.join('\n') };
    }
    case 'note': {
      // 24×24 말풍선 아이콘(색 채우기 + 진한 테두리 + 줄 두 개)
      const [x, y] = [a.rect[0], a.rect[1]];
      lines.push(
        `${rg(a.color, 'rg')} 0.25 0.2 0.1 RG 1 w`,
        `${f(x + 2)} ${f(y + 8)} m ${f(x + 2)} ${f(y + 22)} l ${f(x + 22)} ${f(y + 22)} l ${f(x + 22)} ${f(y + 8)} l ${f(x + 10)} ${f(y + 8)} l ${f(x + 5)} ${f(y + 2)} l ${f(x + 6)} ${f(y + 8)} l h B`,
        `${f(x + 6)} ${f(y + 17)} m ${f(x + 18)} ${f(y + 17)} l S ${f(x + 6)} ${f(y + 13)} m ${f(x + 15)} ${f(y + 13)} l S`,
      );
      return { content: lines.join('\n') };
    }
    case 'freetext': {
      const size = a.fontSize ?? 12;
      const text = a.text ?? '';
      const font = WIN_ANSI.test(text) ? await fonts.standard() : await fonts.korean();
      const [x1, , x2, y2] = a.rect;
      const pad = 2;
      const leading = size * 1.2;
      const wrapped = wrapText(text, font, size, Math.max(10, x2 - x1 - pad * 2));
      lines.push('BT', `/F0 ${f(size)} Tf`, rg(a.color, 'rg'));
      let y = y2 - pad - size;
      for (const line of wrapped) {
        lines.push(`1 0 0 1 ${f(x1 + pad)} ${f(y)} Tm`, `${font.encodeText(line).toString()} Tj`);
        y -= leading;
      }
      lines.push('ET');
      return { content: lines.join('\n'), font };
    }
    default:
      return { content: '' };
  }
}

const SUBTYPE: Record<PdfAnnot['type'], string> = {
  note: 'Text',
  highlight: 'Highlight',
  underline: 'Underline',
  strike: 'StrikeOut',
  ink: 'Ink',
  freetext: 'FreeText',
  rect: 'Square',
  ellipse: 'Circle',
  line: 'Line',
  arrow: 'Line',
};

async function writeOne(doc: PDFDocument, page: PDFPage, a: PdfAnnot, fonts: Fonts): Promise<void> {
  const ctx = doc.context;
  const bounds = boundsOf(a);
  const ap = await appearanceOf(a, fonts);
  const resources: Dict = {};
  const gs: Dict = {};
  if (a.opacity < 1) Object.assign(gs, { CA: a.opacity, ca: a.opacity });
  if (ap.multiply) gs.BM = 'Multiply';
  let content = ap.content;
  if (Object.keys(gs).length > 0) {
    resources.ExtGState = ctx.obj({ GS0: ctx.obj({ Type: 'ExtGState', ...gs }) });
    content = `/GS0 gs\n${content}`;
  }
  if (ap.font) resources.Font = ctx.obj({ F0: ap.font.ref });
  const stream = ctx.flateStream(content, {
    Type: 'XObject',
    Subtype: 'Form',
    BBox: bounds,
    Resources: ctx.obj(resources),
  });
  const apRef = ctx.register(stream);
  const dict: Dict = {
    Type: 'Annot',
    Subtype: SUBTYPE[a.type],
    Rect: a.type === 'line' || a.type === 'arrow' || a.type === 'ink' || a.type === 'highlight' || a.type === 'underline' || a.type === 'strike' ? bounds : a.rect,
    NM: PDFHexString.fromText(a.id),
    T: PDFHexString.fromText(a.author),
    C: colorArray(a.color),
    F: 4,
    P: page.ref,
    AP: ctx.obj({ N: apRef }),
    Border: [0, 0, a.width],
    BS: ctx.obj({ W: a.width, S: 'S' }),
  };
  if (a.opacity < 1) dict.CA = a.opacity;
  const m = pdfDate(a.modifiedAt);
  if (m) dict.M = PDFHexString.fromText(m);
  const c = pdfDate(a.createdAt);
  if (c) dict.CreationDate = PDFHexString.fromText(c);
  if (a.text !== undefined && a.text !== '') dict.Contents = PDFHexString.fromText(a.text);
  switch (a.type) {
    case 'note':
      dict.Name = 'Comment';
      dict.Open = false;
      if (dict.Contents === undefined) dict.Contents = PDFHexString.fromText('');
      break;
    case 'highlight':
    case 'underline':
    case 'strike':
      dict.QuadPoints = a.quads ?? [];
      break;
    case 'ink':
      dict.InkList = (a.paths ?? []).map((p) => [...p]);
      break;
    case 'freetext':
      dict.DA = PDFHexString.fromText(`/Helv ${f(a.fontSize ?? 12)} Tf ${colorArray(a.color).map(f).join(' ')} rg`);
      dict.C = [];
      dict.Contents = PDFHexString.fromText(a.text ?? '');
      dict.Border = [0, 0, 0];
      dict.BS = ctx.obj({ W: 0 });
      break;
    case 'rect':
    case 'ellipse':
      if (a.fill) dict.IC = colorArray(a.fill);
      break;
    case 'line':
    case 'arrow':
      dict.L = [...(a.line ?? [a.rect[0], a.rect[1], a.rect[2], a.rect[3]])];
      dict.LE = [PDFName.of('None'), PDFName.of(a.type === 'arrow' ? 'OpenArrow' : 'None')];
      break;
  }
  const ref: PDFRef = ctx.register(ctx.obj(dict));
  page.node.addAnnot(ref);
}

/** 모델의 주석을 쪽마다 쓴다. pageKey가 가리키는 쪽이 없으면 건너뛴다. */
export async function writeAnnots(doc: PDFDocument, annots: readonly PdfAnnot[], pageOfKey?: (pageKey: string) => PDFPage | undefined): Promise<void> {
  const fonts = fontsFor(doc);
  const pages = doc.getPages();
  for (const a of annots) {
    const page = pageOfKey ? pageOfKey(a.pageKey) : pages[origIndexOf(a.pageKey) ?? -1];
    if (!page) continue;
    await writeOne(doc, page, a, fonts);
  }
}
