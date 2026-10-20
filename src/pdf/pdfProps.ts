/**
 * 문서 속성(설계 §3) — 고칠 수 있는 정보(제목·작성자·주제·키워드)는 방의 pdfProps에 두고 저장 때 쓴다.
 * 읽기만 하는 정보(PDF 버전·쪽 크기·글꼴과 포함 여부·만든 프로그램·날짜)는 pdfInfo로 읽는다.
 */
import { PDFArray, PDFDict, PDFDocument, PDFName, PDFRef, PDFStream, type PDFContext, type PDFObject } from 'pdf-lib';

export const PROP_KEYS = ['title', 'author', 'subject', 'keywords'] as const;
export type PropKey = (typeof PROP_KEYS)[number];
export type PdfProps = Partial<Record<PropKey, string>>;

export function readProps(doc: PDFDocument): PdfProps {
  const safe = (fn: () => string | undefined) => {
    try {
      return fn() ?? '';
    } catch {
      return '';
    }
  };
  return {
    title: safe(() => doc.getTitle()),
    author: safe(() => doc.getAuthor()),
    subject: safe(() => doc.getSubject()),
    keywords: safe(() => doc.getKeywords()),
  };
}

/** 바뀐 속성만 쓴다(같으면 원래 정보 사전을 건드리지 않는다) */
export function writeProps(doc: PDFDocument, props: PdfProps, seeded: PdfProps): void {
  for (const key of PROP_KEYS) {
    const v = props[key];
    if (typeof v !== 'string' || v === (seeded[key] ?? '')) continue;
    const text = v.slice(0, 2000);
    if (key === 'title') doc.setTitle(text);
    else if (key === 'author') doc.setAuthor(text);
    else if (key === 'subject') doc.setSubject(text);
    else doc.setKeywords(text.split(/[,;]\s*/).filter(Boolean));
  }
}

export interface PdfFontInfo {
  readonly name: string;
  readonly type: string;
  readonly embedded: boolean;
}

export interface PdfInfo {
  readonly version: string;
  readonly pageCount: number;
  /** 서로 다른 쪽 크기(pt)와 그 수 */
  readonly pageSizes: ReadonlyArray<{ w: number; h: number; count: number }>;
  readonly fonts: readonly PdfFontInfo[];
  readonly creator: string;
  readonly producer: string;
  readonly created: string | null;
  readonly modified: string | null;
}

function lookup(ctx: PDFContext, o: PDFObject | undefined): PDFObject | undefined {
  return o instanceof PDFRef ? ctx.lookup(o) : o;
}

function fontEntry(ctx: PDFContext, font: PDFDict): PdfFontInfo {
  const name = (lookup(ctx, font.get(PDFName.of('BaseFont'))) as PDFName | undefined)?.decodeText?.() ?? '(이름 없음)';
  const type = (lookup(ctx, font.get(PDFName.of('Subtype'))) as PDFName | undefined)?.decodeText?.() ?? '';
  let desc = lookup(ctx, font.get(PDFName.of('FontDescriptor')));
  if (type === 'Type0') {
    const d = lookup(ctx, font.get(PDFName.of('DescendantFonts')));
    const cid = d instanceof PDFArray ? lookup(ctx, d.get(0)) : undefined;
    desc = cid instanceof PDFDict ? lookup(ctx, cid.get(PDFName.of('FontDescriptor'))) : undefined;
  }
  const embedded = type === 'Type3' || (desc instanceof PDFDict && ['FontFile', 'FontFile2', 'FontFile3'].some((k) => desc.has(PDFName.of(k))));
  return { name: name.replace(/^[A-Z]{6}\+/, ''), type, embedded };
}

export function pdfInfo(doc: PDFDocument, bytes: Uint8Array): PdfInfo {
  const ctx = doc.context;
  const header = Buffer.from(bytes.subarray(0, 16)).toString('latin1');
  const version = /%PDF-(\d\.\d)/.exec(header)?.[1] ?? '';
  const sizes = new Map<string, { w: number; h: number; count: number }>();
  const fonts = new Map<string, PdfFontInfo>();
  const seen = new Set<PDFObject>();
  const scanResources = (res: PDFObject | undefined, depth: number) => {
    const r = lookup(ctx, res);
    if (!(r instanceof PDFDict) || seen.has(r) || depth > 4) return;
    seen.add(r);
    const fdict = lookup(ctx, r.get(PDFName.of('Font')));
    if (fdict instanceof PDFDict) {
      for (const [, v] of fdict.entries()) {
        const f = lookup(ctx, v);
        if (f instanceof PDFDict) {
          const info = fontEntry(ctx, f);
          const key = `${info.name}|${info.type}`;
          if (!fonts.has(key) || (!fonts.get(key)!.embedded && info.embedded)) fonts.set(key, info);
        }
      }
    }
    const xdict = lookup(ctx, r.get(PDFName.of('XObject')));
    if (xdict instanceof PDFDict) {
      for (const [, v] of xdict.entries()) {
        const x = lookup(ctx, v);
        if (x instanceof PDFStream && x.dict.get(PDFName.of('Subtype')) === PDFName.of('Form')) scanResources(x.dict.get(PDFName.of('Resources')), depth + 1);
      }
    }
  };
  for (const page of doc.getPages()) {
    const { width, height } = page.getSize();
    const turned = page.getRotation().angle % 180 !== 0;
    const w = Math.round(turned ? height : width);
    const h = Math.round(turned ? width : height);
    const key = `${w}x${h}`;
    sizes.set(key, { w, h, count: (sizes.get(key)?.count ?? 0) + 1 });
    scanResources(page.node.Resources(), 0);
  }
  const date = (fn: () => Date | undefined) => {
    try {
      return fn()?.toISOString() ?? null;
    } catch {
      return null;
    }
  };
  const text = (fn: () => string | undefined) => {
    try {
      return fn() ?? '';
    } catch {
      return '';
    }
  };
  return {
    version,
    pageCount: doc.getPageCount(),
    pageSizes: [...sizes.values()].sort((a, b) => b.count - a.count),
    fonts: [...fonts.values()].sort((a, b) => a.name.localeCompare(b.name)),
    creator: text(() => doc.getCreator()),
    producer: text(() => doc.getProducer()),
    created: date(() => doc.getCreationDate()),
    modified: date(() => doc.getModificationDate()),
  };
}
