/**
 * 파일에 있던 주석 중 PDF 편집기가 "다루는 종류"를 읽고(readManagedAnnots), 쪽에서 빼낸다(stripManagedAnnots).
 * 다루는 종류: Text(메모)·Highlight·Underline·StrikeOut·Ink·FreeText·Square·Circle·Line — 단 답글(IRT)은 원래 자리에 둔다
 * (부모 메모와 묶여 있어 따로 다시 쓰면 대화가 끊긴다). Popup은 그 부모가 다루는 주석일 때만 함께 뺀다.
 * 편집기는 뺀 PDF를 그리고(pdf-source) 다루는 주석은 자기 층에 그린다 — 저장할 때 모델의 주석을 다시 쓴다(annotWrite.ts).
 */
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNumber, PDFRef, PDFString, type PDFObject } from 'pdf-lib';
import { isStampImage, pageKeyOf, type PdfAnnot, type PdfAnnotType } from './annotTypes.js';

const SUBTYPES: Record<string, PdfAnnotType> = {
  Text: 'note',
  Highlight: 'highlight',
  Underline: 'underline',
  StrikeOut: 'strike',
  Ink: 'ink',
  FreeText: 'freetext',
  Square: 'rect',
  Circle: 'ellipse',
  Line: 'line',
  Stamp: 'stamp',
};
/** 편집기가 만든 서명·도장(Stamp)에만 붙는 그림 사본 — 다시 열었을 때 고칠 수 있게. 다른 Stamp는 손대지 않는다 */
export const STAMP_IMAGE_KEY = 'DocviewImage';
const ARROW_ENDS = new Set(['OpenArrow', 'ClosedArrow', 'ROpenArrow', 'RClosedArrow']);

function lookupDict(doc: PDFDocument, obj: PDFObject | undefined): PDFDict | undefined {
  const value = obj instanceof PDFRef ? doc.context.lookup(obj) : obj;
  return value instanceof PDFDict ? value : undefined;
}

function num(doc: PDFDocument, obj: PDFObject | undefined): number | undefined {
  const value = obj instanceof PDFRef ? doc.context.lookup(obj) : obj;
  return value instanceof PDFNumber && Number.isFinite(value.asNumber()) ? value.asNumber() : undefined;
}

function nums(doc: PDFDocument, obj: PDFObject | undefined): number[] | undefined {
  const value = obj instanceof PDFRef ? doc.context.lookup(obj) : obj;
  if (!(value instanceof PDFArray)) return undefined;
  const out: number[] = [];
  for (const item of value.asArray()) {
    const n = num(doc, item);
    if (n === undefined) return undefined;
    out.push(n);
  }
  return out;
}

function text(doc: PDFDocument, obj: PDFObject | undefined): string | undefined {
  const value = obj instanceof PDFRef ? doc.context.lookup(obj) : obj;
  if (value instanceof PDFString || value instanceof PDFHexString) {
    try {
      return value.decodeText();
    } catch {
      return undefined;
    }
  }
  return undefined;
}

function nameOf(doc: PDFDocument, obj: PDFObject | undefined): string | undefined {
  const value = obj instanceof PDFRef ? doc.context.lookup(obj) : obj;
  return value instanceof PDFName ? value.decodeText() : undefined;
}

const hex2 = (v: number) => Math.round(Math.min(1, Math.max(0, v)) * 255).toString(16).padStart(2, '0').toUpperCase();

/** 주석 색 배열(회색 1개·RGB 3개·CMYK 4개) → '#RRGGBB'. 빈 배열(투명)이면 null */
export function colorOf(components: number[] | undefined): string | null | undefined {
  if (!components) return undefined;
  if (components.length === 0) return null;
  if (components.length === 1) return `#${hex2(components[0]).repeat(3)}`;
  if (components.length === 3) return `#${components.map(hex2).join('')}`;
  if (components.length === 4) {
    const [c, m, y, k] = components;
    return `#${[c, m, y].map((v) => hex2((1 - v) * (1 - k))).join('')}`;
  }
  return undefined;
}

/** PDF 날짜 "D:YYYYMMDDHHmmSS+09'00'" → ISO 문자열(못 읽으면 '') */
export function pdfDateToIso(value: string | undefined): string {
  const m = /^D?:?(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?([Zz+-])?(\d{2})?'?(\d{2})?/.exec(value?.trim() ?? '');
  if (!m) return '';
  const [, y, mo = '01', d = '01', h = '00', mi = '00', s = '00', sign, oh = '00', om = '00'] = m;
  const offset = !sign || sign === 'Z' || sign === 'z' ? 'Z' : `${sign}${oh}:${om}`;
  const date = new Date(`${y}-${mo}-${d}T${h}:${mi}:${s}${offset}`);
  return Number.isNaN(date.getTime()) ? '' : date.toISOString();
}

function rectOf(r: number[] | undefined): [number, number, number, number] | undefined {
  if (!r || r.length !== 4) return undefined;
  return [Math.min(r[0], r[2]), Math.min(r[1], r[3]), Math.max(r[0], r[2]), Math.max(r[1], r[3])];
}

/** 이 사전이 편집기가 다루는 주석인지(답글·Popup·숨김 아님) */
function managedType(doc: PDFDocument, dict: PDFDict): PdfAnnotType | undefined {
  const sub = nameOf(doc, dict.get(PDFName.of('Subtype')));
  if (!sub || !(sub in SUBTYPES)) return undefined;
  if (dict.has(PDFName.of('IRT'))) return undefined;
  if (sub === 'Stamp' && !dict.has(PDFName.of(STAMP_IMAGE_KEY))) return undefined;
  return SUBTYPES[sub];
}

function widthOf(doc: PDFDocument, dict: PDFDict): number | undefined {
  const bs = lookupDict(doc, dict.get(PDFName.of('BS')));
  const w = bs ? num(doc, bs.get(PDFName.of('W'))) : undefined;
  if (w !== undefined) return w;
  const border = nums(doc, dict.get(PDFName.of('Border')));
  return border && border.length >= 3 ? border[2] : undefined;
}

/** FreeText의 DA "/Helv 14 Tf 0 0 1 rg" → 글자 크기·색 */
function parseDa(da: string | undefined): { fontSize?: number; color?: string } {
  if (!da) return {};
  const size = /(\d+(?:\.\d+)?)\s+Tf/.exec(da);
  const rgbColor = /(\d*\.?\d+)\s+(\d*\.?\d+)\s+(\d*\.?\d+)\s+rg/.exec(da);
  const gray = /(\d*\.?\d+)\s+g(?:\s|$)/.exec(da);
  return {
    fontSize: size ? Number(size[1]) : undefined,
    color: rgbColor ? colorOf([Number(rgbColor[1]), Number(rgbColor[2]), Number(rgbColor[3])]) ?? undefined : gray ? colorOf([Number(gray[1])]) ?? undefined : undefined,
  };
}

function toAnnot(doc: PDFDocument, dict: PDFDict, type: PdfAnnotType, pageIndex: number, order: number): PdfAnnot | undefined {
  const rect = rectOf(nums(doc, dict.get(PDFName.of('Rect'))));
  if (!rect) return undefined;
  // NM은 보통 문자열이지만 이름(/NM /abc)으로 쓰는 프로그램도 있다
  const id = text(doc, dict.get(PDFName.of('NM'))) || nameOf(doc, dict.get(PDFName.of('NM'))) || `f${pageIndex}-${order}`;
  const modified = pdfDateToIso(text(doc, dict.get(PDFName.of('M'))));
  const created = pdfDateToIso(text(doc, dict.get(PDFName.of('CreationDate')))) || modified;
  const base = {
    id,
    pageKey: pageKeyOf(pageIndex),
    rect,
    color: colorOf(nums(doc, dict.get(PDFName.of('C')))) ?? '#000000',
    opacity: Math.min(1, Math.max(0, num(doc, dict.get(PDFName.of('CA'))) ?? 1)),
    width: widthOf(doc, dict) ?? 1,
    author: text(doc, dict.get(PDFName.of('T'))) ?? '',
    createdAt: created,
    modifiedAt: modified,
  };
  const contents = text(doc, dict.get(PDFName.of('Contents')));
  switch (type) {
    case 'note':
      return { ...base, type, text: contents ?? '', color: base.color === '#000000' ? '#FFD84D' : base.color };
    case 'highlight':
    case 'underline':
    case 'strike': {
      const quads = nums(doc, dict.get(PDFName.of('QuadPoints')));
      const valid = quads && quads.length >= 8 && quads.length % 8 === 0 ? quads : [rect[0], rect[3], rect[2], rect[3], rect[0], rect[1], rect[2], rect[1]];
      return { ...base, type, quads: valid, ...(contents ? { text: contents } : {}) };
    }
    case 'ink': {
      const list = dict.get(PDFName.of('InkList'));
      const arr = list instanceof PDFRef ? doc.context.lookup(list) : list;
      if (!(arr instanceof PDFArray)) return undefined;
      const paths = arr
        .asArray()
        .map((p) => nums(doc, p))
        .filter((p): p is number[] => !!p && p.length >= 2 && p.length % 2 === 0);
      return paths.length ? { ...base, type, paths } : undefined;
    }
    case 'freetext': {
      const da = parseDa(text(doc, dict.get(PDFName.of('DA'))));
      // 글상자는 테두리 없이 쓰므로(annotWrite) 굵기는 의미가 없다 — 늘 1로 맞춰 다시 읽어도 같은 값
      return { ...base, type, width: 1, text: contents ?? '', fontSize: da.fontSize ?? 12, color: da.color ?? '#000000' };
    }
    case 'rect':
    case 'ellipse': {
      const fill = colorOf(nums(doc, dict.get(PDFName.of('IC'))));
      return { ...base, type, fill: fill ?? null, ...(contents ? { text: contents } : {}) };
    }
    case 'stamp': {
      const image = text(doc, dict.get(PDFName.of(STAMP_IMAGE_KEY)));
      return isStampImage(image) ? { ...base, type, width: 0, color: '#000000', image } : undefined;
    }
    case 'line': {
      const l = nums(doc, dict.get(PDFName.of('L')));
      if (!l || l.length !== 4) return undefined;
      const le = dict.get(PDFName.of('LE'));
      const leArr = le instanceof PDFRef ? doc.context.lookup(le) : le;
      const ends = leArr instanceof PDFArray ? leArr.asArray().map((e) => nameOf(doc, e) ?? 'None') : ['None', 'None'];
      const endArrow = ARROW_ENDS.has(ends[1] ?? '');
      const startArrow = ARROW_ENDS.has(ends[0] ?? '');
      // 화살촉은 끝(두 번째 점)에 둔다 — 시작에만 있으면 두 점을 바꾼다
      const line: [number, number, number, number] = startArrow && !endArrow ? [l[2], l[3], l[0], l[1]] : [l[0], l[1], l[2], l[3]];
      return { ...base, type: endArrow || startArrow ? 'arrow' : 'line', line, ...(contents ? { text: contents } : {}) };
    }
    default:
      return undefined;
  }
}

function annotRefsOf(doc: PDFDocument, pageIndex: number): PDFObject[] {
  const page = doc.getPage(pageIndex);
  const annots = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray);
  return annots ? annots.asArray() : [];
}

export function readManagedAnnots(doc: PDFDocument): PdfAnnot[] {
  const out: PdfAnnot[] = [];
  const seen = new Set<string>();
  doc.getPages().forEach((_page, pageIndex) => {
    annotRefsOf(doc, pageIndex).forEach((obj, order) => {
      const dict = lookupDict(doc, obj);
      if (!dict) return;
      const type = managedType(doc, dict);
      if (!type) return;
      try {
        const annot = toAnnot(doc, dict, type, pageIndex, order);
        if (!annot) return;
        // NM이 겹치는 파일(복사해 붙인 주석 등) — id가 겹치면 맵에서 하나가 사라지므로 뒤엣것에 쪽·순번을 붙인다
        const id = seen.has(annot.id) ? `${annot.id}~${pageIndex}-${order}` : annot.id;
        seen.add(id);
        out.push(id === annot.id ? annot : { ...annot, id });
      } catch {
        // 이상한 주석 하나 때문에 전체를 버리지 않는다
      }
    });
  });
  return out;
}

export function stripManagedAnnots(doc: PDFDocument): void {
  doc.getPages().forEach((page, pageIndex) => {
    const refs = annotRefsOf(doc, pageIndex);
    if (refs.length === 0) return;
    const removed = new Set<PDFObject>();
    const removedPopups = new Set<string>();
    for (const obj of refs) {
      const dict = lookupDict(doc, obj);
      if (!dict || !managedType(doc, dict)) continue;
      removed.add(obj);
      const popup = dict.get(PDFName.of('Popup'));
      if (popup instanceof PDFRef) removedPopups.add(popup.toString());
    }
    if (removed.size === 0) return;
    const removedRefs = new Set([...removed].filter((o): o is PDFRef => o instanceof PDFRef).map((r) => r.toString()));
    const kept = refs.filter((obj) => {
      if (removed.has(obj)) return false;
      if (obj instanceof PDFRef && removedPopups.has(obj.toString())) return false;
      // Parent가 뺀 주석인 Popup도 함께
      const dict = lookupDict(doc, obj);
      const parent = dict?.get(PDFName.of('Parent'));
      return !(nameOf(doc, dict?.get(PDFName.of('Subtype'))) === 'Popup' && parent instanceof PDFRef && removedRefs.has(parent.toString()));
    });
    if (kept.length === 0) page.node.delete(PDFName.of('Annots'));
    else page.node.set(PDFName.of('Annots'), doc.context.obj(kept));
  });
}
