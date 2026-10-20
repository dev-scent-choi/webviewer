/**
 * 쪽 내용 고쳐 쓰기 — 글자 직접 고치기와 가리기(설계 §3).
 *
 * 내용 스트림(contentStream.ts)을 그래픽·글자 상태(CTM·Tm·Tf·Tc·Tw·Tz·TL·Ts)와 함께 따라가며 글자마다 쪽 좌표의 상자를 계산한다
 * (글꼴 너비는 pdfFonts.ts). 영역에 드는 글자는 **실제로 지운다** — 그 글자 자리는 TJ 간격 값으로 바꿔 남은 글자의 위치는
 * 그대로 둔다. Form XObject 안의 글자도 따라 들어가 고친 사본을 만든다(다른 쪽과 같이 쓰는 원본은 건드리지 않는다).
 *  - 고치기(edit): 글자 가운데가 영역 안이면 지우고, 같은 자리(같은 방향·크기·색)에 새 글을 쓴다. 새 글이 지운 글의 글자들로만
 *    되어 있으면 원래 글꼴을 그대로 쓰고(모양이 같다), 아니면 Pretendard(한글)·Helvetica.
 *  - 가리기(redact): 영역에 조금이라도 걸치는 글자를 지우고, 그림은 그 부분을 검게 칠한 사본으로 바꾸며(못 다루는 형식이면 그림을
 *    뺀다), 인라인 그림은 뺀다. 마지막에 검은 상자를 그린다. 걸치는 주석·링크·양식 칸은 쪽에서 뺀다.
 * 원래 내용은 `q … Q`로 감싸 뒤에 덧붙이는 그림이 원래 좌표계의 영향을 받지 않게 한다.
 */
import {
  decodePDFRawStream,
  PDFArray,
  PDFDict,
  PDFName,
  PDFNumber,
  PDFRawStream,
  PDFRef,
  PDFStream,
  type PDFContext,
  type PDFDocument,
  type PDFObject,
  type PDFPage,
} from 'pdf-lib';
import { apply, hexOf, IDENTITY, invert, mul, nameOf, num, opToString, parseContent, type ContentOp, type Matrix, type Operand } from './contentStream.js';
import { fontInfoOf, type FontInfo } from './pdfFonts.js';
import { decodeImage } from './pdfImages.js';
import type { Fonts } from './annotWrite.js';
import { dropFieldsOnRemovedPages } from './pdfForm.js';

export type Rect = readonly [number, number, number, number];

export interface RewriteTarget {
  readonly id: string;
  readonly mode: 'edit' | 'redact';
  /** PDF 좌표(x1, y1, x2, y2) */
  readonly rect: Rect;
  /** 고치기: 원래 글(편집기가 고른 글 조각)·새 글·글자 크기(어림) */
  readonly oldText?: string;
  readonly newText?: string;
  readonly fontSize?: number;
}

/** 고치기 영역에서 지운 글자들 — 새 글을 쓸 자리·글꼴·색 */
interface RemovedRun {
  readonly targetId: string;
  /** 쪽 공간으로의 글자 행렬(Th·Ts 포함, 글자 크기 빼고) — 첫 글자의 기준점 */
  readonly matrix: Matrix;
  readonly size: number;
  /** 쪽 자원의 글꼴 이름(Form 안이면 null — 원래 글꼴을 다시 쓰지 않는다) */
  readonly fontRes: string | null;
  readonly codes: Uint8Array[];
  readonly fill: string;
}

interface ImageHit {
  readonly ref: PDFRef;
  readonly hits: Array<{ ctm: Matrix; rect: Rect }>;
}

const rectOf = (pts: Array<[number, number]>): Rect => {
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)];
};
const intersects = (a: Rect, b: Rect) => a[0] < b[2] && a[2] > b[0] && a[1] < b[3] && a[3] > b[1];
const contains = (r: Rect, x: number, y: number) => x >= r[0] && x <= r[2] && y >= r[1] && y <= r[3];
const isNum = (o: Operand | undefined): o is number => typeof o === 'number';
const isName = (o: Operand | undefined): o is { name: string } => typeof o === 'object' && o !== null && 'name' in o;
const isStr = (o: Operand | undefined): o is { bytes: Uint8Array; hex: boolean } => typeof o === 'object' && o !== null && 'bytes' in o;
const unitSquare = (ctm: Matrix) => rectOf([apply(ctm, 0, 0), apply(ctm, 1, 0), apply(ctm, 0, 1), apply(ctm, 1, 1)]);

function lookup(ctx: PDFContext, o: PDFObject | undefined): PDFObject | undefined {
  return o instanceof PDFRef ? ctx.lookup(o) : o;
}
function dictOf(ctx: PDFContext, o: PDFObject | undefined): PDFDict | undefined {
  const v = lookup(ctx, o);
  return v instanceof PDFDict ? v : undefined;
}

/** 스트림 내용 바이트(필터 풀기) */
export function streamBytes(stream: PDFObject | undefined): Uint8Array | null {
  if (stream instanceof PDFRawStream) {
    try {
      return decodePDFRawStream(stream).decode();
    } catch {
      return null;
    }
  }
  // pdf-lib가 이번에 만든 스트림(압축 전 내용을 가진 PDFFlateStream 등)
  if (stream instanceof PDFStream) {
    const s = stream as PDFStream & { getUnencodedContents?: () => Uint8Array };
    return typeof s.getUnencodedContents === 'function' ? s.getUnencodedContents() : stream.getContents();
  }
  return null;
}

/** 쪽 내용(여러 스트림이면 이어서) */
function pageContent(ctx: PDFContext, page: PDFPage): Uint8Array | null {
  const contents = lookup(ctx, page.node.get(PDFName.of('Contents')));
  const parts: Uint8Array[] = [];
  const list = contents instanceof PDFArray ? contents.asArray().map((o) => lookup(ctx, o)) : [contents];
  for (const s of list) {
    if (!s) continue;
    const b = streamBytes(s);
    if (!b) return null;
    parts.push(b, Uint8Array.of(0x0a));
  }
  return Buffer.concat(parts);
}

/** 자원 사전 안의 하위 사전(Font·XObject)을 이 자원만의 사본으로 — 같이 쓰는 원본을 고치지 않게 */
function ownSubdict(ctx: PDFContext, res: PDFDict, key: 'Font' | 'XObject', owned: WeakSet<PDFDict>): PDFDict {
  const cur = dictOf(ctx, res.get(PDFName.of(key)));
  if (cur && owned.has(cur)) return cur;
  const copy = cur ? cur.clone(ctx) : ctx.obj({});
  owned.add(copy);
  res.set(PDFName.of(key), copy);
  return copy;
}
function uniqueName(dict: PDFDict, prefix: string): string {
  for (let i = 1; ; i += 1) if (!dict.has(PDFName.of(`${prefix}${i}`))) return `${prefix}${i}`;
}

interface TextState {
  ctm: Matrix;
  fontRes: string | null;
  font: FontInfo | null;
  size: number;
  tc: number;
  tw: number;
  th: number;
  tl: number;
  rise: number;
  fill: string;
}

interface RunContext {
  /** 1차(걸친 그림 모으기)면 true — Form 사본을 만들지 않는다 */
  readonly analyze: boolean;
  readonly ctx: PDFContext;
  readonly targets: readonly RewriteTarget[];
  /** 그림 ref → 검게 칠한 사본 ref 또는 'drop'(2차) */
  readonly decisions: ReadonlyMap<string, PDFRef | 'drop'>;
  readonly imageHits: Map<string, ImageHit>;
  readonly runs: Map<string, RemovedRun>;
  readonly owned: WeakSet<PDFDict>;
  removedGlyphs: number;
  droppedImages: number;
}

/** 내용 스트림 한 개를 고쳐 쓴다. 자원(res)은 이 호출만의 사본이어야 한다(그림·Form 이름을 더할 수 있다). */
function rewriteStream(rc: RunContext, bytes: Uint8Array, res: PDFDict | undefined, ctm0: Matrix, inForm: boolean, depth: number): { bytes: Uint8Array; changed: boolean } {
  const { ctx, targets } = rc;
  const edits = targets.filter((t) => t.mode === 'edit');
  const redacts = targets.filter((t) => t.mode === 'redact');
  const ops = parseContent(bytes);
  const fontCache = new Map<string, FontInfo>();
  let st: TextState = { ctm: ctm0, fontRes: null, font: null, size: 0, tc: 0, tw: 0, th: 1, tl: 0, rise: 0, fill: '' };
  const stack: TextState[] = [];
  let tm: Matrix = IDENTITY;
  let tlm: Matrix = IDENTITY;
  const out: Uint8Array[] = [];
  let cursor = 0;
  let changed = false;
  const replace = (op: ContentOp, text: string | null) => {
    out.push(bytes.subarray(cursor, op.start));
    if (text !== null) out.push(Buffer.from(`${text}\n`, 'latin1'));
    cursor = op.end;
    changed = true;
  };
  const fontsDict = res ? dictOf(ctx, res.get(PDFName.of('Font'))) : undefined;
  const xobjects = () => (res ? dictOf(ctx, res.get(PDFName.of('XObject'))) : undefined);

  const nextLine = () => {
    tlm = mul([1, 0, 0, 1, 0, -st.tl], tlm);
    tm = tlm;
  };

  /** 글 보이기 — 바뀌었으면 새 TJ 배열(피연산자), 아니면 null */
  const show = (items: Operand[]): Operand[] | null => {
    const font = st.font;
    if (!font || st.size === 0) {
      // 글꼴을 모르면 위치만 대강 넘긴다(지우지 않는다)
      return null;
    }
    const outItems: Operand[] = [];
    let pendingBytes: number[] = [];
    let pendingNum = 0;
    let any = false;
    const flushBytes = () => {
      if (pendingBytes.length) {
        if (pendingNum) outItems.push(pendingNum);
        pendingNum = 0;
        outItems.push({ bytes: Uint8Array.from(pendingBytes), hex: true });
        pendingBytes = [];
      }
    };
    for (const item of items) {
      if (isNum(item)) {
        flushBytes();
        pendingNum += item;
        tm = mul([1, 0, 0, 1, (-item / 1000) * st.size * st.th, 0], tm);
        continue;
      }
      if (!isStr(item)) continue;
      for (const code of font.split(item.bytes)) {
        const w0 = font.width(code);
        const trm = mul([st.size * st.th, 0, 0, st.size, 0, st.rise], mul(tm, st.ctm));
        const box = rectOf([apply(trm, 0, font.descent / 1000), apply(trm, w0 / 1000, font.descent / 1000), apply(trm, 0, font.ascent / 1000), apply(trm, w0 / 1000, font.ascent / 1000)]);
        const [cx, cy] = apply(trm, w0 / 2000, (font.ascent + font.descent) / 2000);
        const edit = edits.find((t) => contains(t.rect, cx, cy));
        const hit = edit ?? redacts.find((t) => intersects(t.rect, box));
        const space = font.simple && code.length === 1 && code[0] === 32 ? st.tw : 0;
        const advance = ((w0 / 1000) * st.size + st.tc + space) * st.th;
        if (hit) {
          any = true;
          rc.removedGlyphs += 1;
          if (edit) {
            const prev = rc.runs.get(edit.id);
            if (prev) prev.codes.push(code);
            else {
              rc.runs.set(edit.id, {
                targetId: edit.id,
                matrix: mul([st.th, 0, 0, 1, 0, st.rise], mul(tm, st.ctm)),
                size: st.size,
                fontRes: inForm ? null : st.fontRes,
                codes: [code],
                fill: st.fill,
              });
            }
          }
          // 지운 글자 자리만큼 띄운다(Tc·Tw까지)
          flushBytes();
          pendingNum += -(w0 + ((st.tc + space) * 1000) / st.size);
        } else {
          if (pendingNum) {
            outItems.push(pendingNum);
            pendingNum = 0;
          }
          pendingBytes.push(...code);
        }
        tm = mul([1, 0, 0, 1, advance, 0], tm);
      }
    }
    flushBytes();
    if (pendingNum) outItems.push(pendingNum);
    return any ? outItems : null;
  };

  for (const op of ops) {
    const a = op.operands;
    switch (op.op) {
      case 'q':
        stack.push({ ...st });
        break;
      case 'Q':
        st = stack.pop() ?? st;
        break;
      case 'cm':
        if (a.length === 6 && a.every(isNum)) st.ctm = mul(a as number[], st.ctm);
        break;
      case 'g':
      case 'rg':
      case 'k':
        st.fill = opToString(op.op, a);
        break;
      case 'cs':
        st.fill = opToString('cs', a);
        break;
      case 'sc':
      case 'scn': {
        // 색 공간(cs)은 앞 줄에 남겨 둔다
        const cs = st.fill.split('\n')[0];
        st.fill = `${cs.endsWith(' cs') ? `${cs}\n` : ''}${opToString(op.op, a)}`;
        break;
      }
      case 'BT':
        tm = IDENTITY;
        tlm = IDENTITY;
        break;
      case 'Tf':
        if (isName(a[0]) && isNum(a[1])) {
          st.fontRes = a[0].name;
          st.size = a[1];
          let info = fontCache.get(a[0].name);
          if (!info) {
            info = fontInfoOf(ctx, fontsDict?.get(PDFName.of(a[0].name)));
            fontCache.set(a[0].name, info);
          }
          st.font = info;
        }
        break;
      case 'Tc':
        if (isNum(a[0])) st.tc = a[0];
        break;
      case 'Tw':
        if (isNum(a[0])) st.tw = a[0];
        break;
      case 'Tz':
        if (isNum(a[0])) st.th = a[0] / 100;
        break;
      case 'TL':
        if (isNum(a[0])) st.tl = a[0];
        break;
      case 'Ts':
        if (isNum(a[0])) st.rise = a[0];
        break;
      case 'Td':
      case 'TD':
        if (isNum(a[0]) && isNum(a[1])) {
          if (op.op === 'TD') st.tl = -a[1];
          tlm = mul([1, 0, 0, 1, a[0], a[1]], tlm);
          tm = tlm;
        }
        break;
      case 'Tm':
        if (a.length === 6 && a.every(isNum)) {
          tlm = [...(a as number[])] as Matrix;
          tm = tlm;
        }
        break;
      case 'T*':
        nextLine();
        break;
      case 'Tj':
      case 'TJ':
      case "'":
      case '"': {
        let pre = '';
        let items: Operand[] = [];
        if (op.op === '"') {
          if (isNum(a[0])) st.tw = a[0];
          if (isNum(a[1])) st.tc = a[1];
          pre = `${num(st.tw)} Tw ${num(st.tc)} Tc T* `;
          nextLine();
          items = a[2] ? [a[2]] : [];
        } else if (op.op === "'") {
          pre = 'T* ';
          nextLine();
          items = a[0] ? [a[0]] : [];
        } else if (op.op === 'Tj') items = a[0] ? [a[0]] : [];
        else items = a[0] && typeof a[0] === 'object' && a[0] !== null && 'array' in a[0] ? a[0].array : [];
        const next = show(items);
        if (next) replace(op, `${pre}${opToString('TJ', [{ array: next }])}`);
        break;
      }
      case 'BI': {
        if (redacts.some((t) => intersects(t.rect, unitSquare(st.ctm)))) {
          replace(op, null);
          rc.droppedImages += 1;
        }
        break;
      }
      case 'Do': {
        if (!isName(a[0])) break;
        const xdict = xobjects();
        const refObj = xdict?.get(PDFName.of(a[0].name));
        const xo = lookup(ctx, refObj);
        if (!(xo instanceof PDFRawStream || xo instanceof PDFStream)) break;
        const subtype = xo.dict.get(PDFName.of('Subtype'));
        if (subtype === PDFName.of('Image')) {
          const bbox = unitSquare(st.ctm);
          const hits = redacts.filter((t) => intersects(t.rect, bbox));
          if (hits.length === 0 || !(refObj instanceof PDFRef) || !res) break;
          const decision = rc.decisions.get(refObj.tag);
          if (decision === undefined) {
            const entry = rc.imageHits.get(refObj.tag) ?? { ref: refObj, hits: [] };
            for (const t of hits) entry.hits.push({ ctm: st.ctm, rect: t.rect });
            rc.imageHits.set(refObj.tag, entry);
          } else if (decision === 'drop') {
            replace(op, null);
            rc.droppedImages += 1;
          } else {
            const own = ownSubdict(ctx, res, 'XObject', rc.owned);
            const name = uniqueName(own, 'DvIm');
            own.set(PDFName.of(name), decision);
            replace(op, `${nameOf(name)} Do`);
          }
        } else if (subtype === PDFName.of('Form') && depth < 4 && res) {
          const matrixArr = lookup(ctx, xo.dict.get(PDFName.of('Matrix')));
          const fm = matrixArr instanceof PDFArray ? (matrixArr.asArray().map((v) => (lookup(ctx, v) instanceof PDFNumber ? (lookup(ctx, v) as PDFNumber).asNumber() : 0)) as Matrix) : IDENTITY;
          const formCtm = mul(fm, st.ctm);
          // Form의 영역이 어떤 대상과도 겹치지 않으면 들어가지 않는다
          const bboxArr = lookup(ctx, xo.dict.get(PDFName.of('BBox')));
          if (bboxArr instanceof PDFArray) {
            const bb = bboxArr.asArray().map((v) => (lookup(ctx, v) instanceof PDFNumber ? (lookup(ctx, v) as PDFNumber).asNumber() : 0));
            const area = rectOf([apply(formCtm, bb[0], bb[1]), apply(formCtm, bb[2], bb[1]), apply(formCtm, bb[0], bb[3]), apply(formCtm, bb[2], bb[3])]);
            if (!targets.some((t) => intersects(t.rect, area))) break;
          }
          const content = streamBytes(xo);
          if (!content) break;
          const formResOrig = dictOf(ctx, xo.dict.get(PDFName.of('Resources')));
          const formRes = (formResOrig ?? res).clone(ctx);
          const inner = rewriteStream(rc, content, formRes, formCtm, true, depth + 1);
          if (!inner.changed || rc.analyze) break;
          const entries: Record<string, PDFObject> = {};
          for (const [k, v] of xo.dict.entries()) {
            const key = k.decodeText();
            if (key !== 'Length' && key !== 'Filter' && key !== 'DecodeParms') entries[key] = v;
          }
          entries.Resources = formRes;
          const copyRef = ctx.register(ctx.flateStream(inner.bytes, entries));
          const own = ownSubdict(ctx, res, 'XObject', rc.owned);
          const name = uniqueName(own, 'DvFm');
          own.set(PDFName.of(name), copyRef);
          replace(op, `${nameOf(name)} Do`);
        }
        break;
      }
    }
  }
  out.push(bytes.subarray(cursor));
  return { bytes: changed ? Buffer.concat(out) : bytes, changed };
}

// ---------- 그림 검게 칠하기 ----------
/** 그림의 걸친 부분을 검게 칠한 사본 — 다룰 수 없는 형식이면 null(그림을 뺀다) */
async function blackoutImage(ctx: PDFContext, image: ImageHit): Promise<PDFRef | null> {
  const stream = ctx.lookup(image.ref);
  const decoded = await decodeImage(ctx, stream);
  if (!decoded || !(stream instanceof PDFRawStream)) return null;
  const d = stream.dict;
  const { raw, channels, colorSpace } = decoded;
  const W = decoded.width;
  const H = decoded.height;
  const black = channels === 4 ? [0, 0, 0, 255] : new Array(channels).fill(0);
  for (const { ctm, rect } of image.hits) {
    const inv = invert(ctm);
    if (!inv) return null;
    const uv = [apply(inv, rect[0], rect[1]), apply(inv, rect[2], rect[1]), apply(inv, rect[0], rect[3]), apply(inv, rect[2], rect[3])];
    const xs = uv.map((p) => p[0] * W);
    const ys = uv.map((p) => (1 - p[1]) * H);
    const x0 = Math.max(0, Math.floor(Math.min(...xs)));
    const x1 = Math.min(W, Math.ceil(Math.max(...xs)));
    const y0 = Math.max(0, Math.floor(Math.min(...ys)));
    const y1 = Math.min(H, Math.ceil(Math.max(...ys)));
    for (let y = y0; y < y1; y += 1) {
      for (let x = x0; x < x1; x += 1) {
        const o = (y * W + x) * channels;
        for (let c = 0; c < channels; c += 1) raw[o + c] = black[c];
      }
    }
  }
  const entries: Record<string, PDFObject> = {
    Type: PDFName.of('XObject'),
    Subtype: PDFName.of('Image'),
    Width: PDFNumber.of(W),
    Height: PDFNumber.of(H),
    BitsPerComponent: PDFNumber.of(8),
    ColorSpace: colorSpace,
  };
  const smask = d.get(PDFName.of('SMask'));
  if (smask) entries.SMask = smask;
  return ctx.register(ctx.flateStream(raw, entries));
}

export interface PageRewriteResult {
  readonly removedGlyphs: number;
  readonly droppedImages: number;
  readonly blackedImages: number;
  readonly removedAnnots: number;
}

/**
 * 한 쪽에 고치기·가리기를 적용한다. 가리기는 검은 상자까지, 고치기는 새 글까지 그린다.
 * 바뀐 것이 없으면 쪽을 건드리지 않는다.
 */
export async function rewritePage(doc: PDFDocument, page: PDFPage, targets: readonly RewriteTarget[], fonts: Fonts): Promise<PageRewriteResult> {
  const ctx = doc.context;
  const empty = { removedGlyphs: 0, droppedImages: 0, blackedImages: 0, removedAnnots: 0 };
  if (targets.length === 0) return empty;
  const content = pageContent(ctx, page) ?? new Uint8Array();
  const resOrig = page.node.Resources();
  const res = resOrig ? resOrig.clone(ctx) : ctx.obj({});
  const owned = new WeakSet<PDFDict>();
  // 1차: 걸친 그림을 모은다
  const first: RunContext = { analyze: true, ctx, targets, decisions: new Map(), imageHits: new Map(), runs: new Map(), owned, removedGlyphs: 0, droppedImages: 0 };
  rewriteStream(first, content, res.clone(ctx), IDENTITY, false, 0);
  const decisions = new Map<string, PDFRef | 'drop'>();
  let blackedImages = 0;
  for (const [tag, hit] of first.imageHits) {
    let ref: PDFRef | null = null;
    try {
      ref = await blackoutImage(ctx, hit);
    } catch (error) {
      console.error('가리기: 그림 칠하기 실패(그림을 뺀다):', error);
    }
    decisions.set(tag, ref ?? 'drop');
    if (ref) blackedImages += 1;
  }
  // 2차: 실제로 고쳐 쓴다
  const rc: RunContext = { analyze: false, ctx, targets, decisions, imageHits: new Map(), runs: new Map(), owned, removedGlyphs: 0, droppedImages: 0 };
  const rewritten = rewriteStream(rc, content, res, IDENTITY, false, 0);

  // 덧그리기: 고친 글 · 검은 상자
  const extra: string[] = [];
  const redactAreas = targets.filter((t) => t.mode === 'redact').map((t) => t.rect);
  for (const t of targets) {
    // 가린 곳과 겹치는 고친 글은 쓰지 않는다(검은 상자 밑에 글이 남지 않게)
    if (t.mode === 'edit' && t.newText && !redactAreas.some((r) => intersects(r, t.rect))) {
      const run = rc.runs.get(t.id);
      extra.push(await drawEditText(ctx, res, owned, t, run, fonts));
    }
  }
  for (const t of targets) {
    if (t.mode === 'redact') extra.push(`0 g ${num(t.rect[0])} ${num(t.rect[1])} ${num(t.rect[2] - t.rect[0])} ${num(t.rect[3] - t.rect[1])} re f`);
  }
  // 가린 곳에 걸치는 주석(링크·양식 칸 포함)은 뺀다
  let removedAnnots = 0;
  const redactRects = targets.filter((t) => t.mode === 'redact').map((t) => t.rect);
  if (redactRects.length) {
    const annots = page.node.Annots();
    const removed = new Set<string>();
    if (annots) {
      for (let i = annots.size() - 1; i >= 0; i -= 1) {
        const item = annots.get(i);
        const a = dictOf(ctx, item);
        const r = a ? lookup(ctx, a.get(PDFName.of('Rect'))) : undefined;
        if (!(r instanceof PDFArray)) continue;
        const v = r.asArray().map((x) => (lookup(ctx, x) instanceof PDFNumber ? (lookup(ctx, x) as PDFNumber).asNumber() : 0));
        const ar: Rect = [Math.min(v[0], v[2]), Math.min(v[1], v[3]), Math.max(v[0], v[2]), Math.max(v[1], v[3])];
        if (redactRects.some((rr) => intersects(rr, ar))) {
          annots.remove(i);
          if (item instanceof PDFRef) removed.add(item.tag);
          removedAnnots += 1;
        }
      }
    }
    dropFieldsOnRemovedPages(doc, removed);
  }

  if (!rewritten.changed && extra.length === 0) return { ...empty, removedAnnots };
  const body = Buffer.concat([Buffer.from('q\n', 'latin1'), Buffer.from(rewritten.bytes), Buffer.from('\nQ\n', 'latin1'), Buffer.from(extra.join('\n'), 'latin1')]);
  page.node.set(PDFName.of('Contents'), ctx.register(ctx.flateStream(body)));
  page.node.set(PDFName.of('Resources'), res);
  return { removedGlyphs: rc.removedGlyphs, droppedImages: rc.droppedImages, blackedImages, removedAnnots };
}

const WIN_ANSI = /^[\x20-\x7e]*$/;

/** 고친 글 — 지운 글자들의 자리·방향·크기·색으로. 원래 글꼴로 쓸 수 있으면(새 글이 지운 글자들로만) 그 글꼴 */
async function drawEditText(ctx: PDFContext, res: PDFDict, owned: WeakSet<PDFDict>, t: RewriteTarget, run: RemovedRun | undefined, fonts: Fonts): Promise<string> {
  const text = (t.newText ?? '').replace(/[\r\n]+/g, ' ');
  if (!text) return '';
  const oldChars = [...(t.oldText ?? '')];
  let matrix: Matrix;
  let size: number;
  let fill = '0 g';
  if (run) {
    matrix = run.matrix;
    size = run.size;
    if (run.fill) fill = run.fill;
  } else {
    // 지운 글자가 없으면(글꼴을 모르는 글 등) 영역 아래쪽을 기준선으로
    size = t.fontSize ?? Math.max(4, (t.rect[3] - t.rect[1]) * 0.8);
    matrix = [1, 0, 0, 1, t.rect[0], t.rect[1] + (t.rect[3] - t.rect[1]) * 0.2];
  }
  // 원래 글꼴: 지운 글자 수와 원래 글의 글자 수가 같으면 글자 ↔ 코드를 맞춰 쓸 수 있다
  let fontName: string | null = null;
  let encoded: string | null = null;
  if (run?.fontRes && run.codes.length === oldChars.length) {
    const codeOf = new Map<string, Uint8Array>();
    oldChars.forEach((ch, i) => codeOf.set(ch, run.codes[i]));
    if ([...text].every((ch) => codeOf.has(ch))) {
      fontName = run.fontRes;
      encoded = hexOf(Buffer.concat([...text].map((ch) => codeOf.get(ch)!)));
    }
  }
  if (!fontName) {
    const font = WIN_ANSI.test(text) ? await fonts.standard() : await fonts.korean();
    const own = ownSubdict(ctx, res, 'Font', owned);
    fontName = uniqueName(own, 'DvF');
    own.set(PDFName.of(fontName), font.ref);
    encoded = font.encodeText(text).toString();
  }
  return `q BT ${fill} ${nameOf(fontName)} ${num(size)} Tf ${matrix.map(num).join(' ')} Tm ${encoded} Tj ET Q`;
}
