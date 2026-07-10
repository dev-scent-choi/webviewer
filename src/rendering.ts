import fs from 'node:fs/promises';
import { PDFDocument, StandardFonts, degrees, rgb } from 'pdf-lib';
import { pdf } from 'pdf-to-img';
import { config } from './config.js';

type PdfImageDocument = Awaited<ReturnType<typeof pdf>>;

interface PageSize {
  readonly width: number;
  readonly height: number;
}

const DOCUMENT_CACHE_LIMIT = 6;
/** 페이지당 렌더링 픽셀 상한 — 초대형 시트/도면에서 메모리 폭주를 막는다. */
const MAX_BITMAP_EDGE_PX = 6000;
/* 실질적인 안전장치는 페이지 픽셀 상한(MAX_BITMAP_EDGE_PX)이며,
   작은 페이지(크롭된 도면 등)는 높은 배율로 렌더링해야 확대 시 선명하다 */
export const MAX_RENDER_SCALE = 16;

const documentCache = new Map<string, Promise<PdfImageDocument>>();
const pageSizeCache = new Map<string, Promise<readonly PageSize[]>>();

export class PageOutOfRangeError extends Error {
  constructor(pageNumber: number, pageCount: number) {
    super(`페이지 범위를 벗어났습니다: ${pageNumber} (전체 ${pageCount}페이지)`);
    this.name = 'PageOutOfRangeError';
  }
}

function evictOldest<K, V>(cache: Map<K, V>, limit: number): void {
  while (cache.size > limit) {
    const oldestKey = cache.keys().next().value;
    if (oldestKey === undefined) return;
    cache.delete(oldestKey);
  }
}

function openDocument(pdfPath: string, scale: number): Promise<PdfImageDocument> {
  const key = `${pdfPath}#${scale.toFixed(2)}`;
  const cached = documentCache.get(key);
  if (cached) {
    return cached;
  }
  const loading = pdf(pdfPath, { scale });
  loading.catch(() => documentCache.delete(key));
  documentCache.set(key, loading);
  evictOldest(documentCache, DOCUMENT_CACHE_LIMIT);
  return loading;
}

function getPageSizes(pdfPath: string): Promise<readonly PageSize[]> {
  const cached = pageSizeCache.get(pdfPath);
  if (cached) {
    return cached;
  }
  const loading = (async () => {
    const bytes = await fs.readFile(pdfPath);
    const doc = await PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
    return doc.getPages().map((page) => {
      const { width, height } = page.getSize();
      return { width, height };
    });
  })();
  loading.catch(() => pageSizeCache.delete(pdfPath));
  pageSizeCache.set(pdfPath, loading);
  evictOldest(pageSizeCache, DOCUMENT_CACHE_LIMIT * 4);
  return loading;
}

export async function getPageCount(pdfPath: string): Promise<number> {
  const sizes = await getPageSizes(pdfPath);
  return sizes.length;
}

/**
 * 요청 배율을 페이지 크기 기준 픽셀 상한에 맞게 낮춘다.
 * 도면처럼 원본이 벡터인 PDF는 배율을 올릴수록 선명해진다.
 */
function clampScale(requestedScale: number, size: PageSize): number {
  const bounded = Math.min(Math.max(requestedScale, 1), MAX_RENDER_SCALE);
  const longestEdge = Math.max(size.width, size.height);
  const maxByPixels = MAX_BITMAP_EDGE_PX / longestEdge;
  const clamped = Math.min(bounded, maxByPixels);
  // 캐시 키가 무한히 늘지 않도록 0.5 단위로 정량화한다
  return Math.max(1, Math.floor(clamped * 2) / 2);
}

export async function renderPage(
  pdfPath: string,
  pageNumber: number,
  requestedScale: number = config.renderScale,
): Promise<Buffer> {
  const sizes = await getPageSizes(pdfPath);
  if (pageNumber < 1 || pageNumber > sizes.length) {
    throw new PageOutOfRangeError(pageNumber, sizes.length);
  }
  const scale = clampScale(requestedScale, sizes[pageNumber - 1]);
  const document = await openDocument(pdfPath, scale);
  return document.getPage(pageNumber);
}

/** Helvetica는 WinAnsi 범위만 인코딩 가능하므로 워터마크 문자열을 ASCII로 정리한다. */
function toWinAnsiSafe(text: string): string {
  const safe = text.replace(/[^\x20-\x7E]/g, '').trim();
  return safe.length > 0 ? safe : 'CONFIDENTIAL';
}

export async function buildPrintPdf(pdfPath: string, watermarkText: string): Promise<Uint8Array> {
  const bytes = await fs.readFile(pdfPath);
  const doc = await PDFDocument.load(bytes, { ignoreEncryption: true });
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const text = toWinAnsiSafe(watermarkText);
  for (const page of doc.getPages()) {
    const { width, height } = page.getSize();
    const fontSize = Math.max(18, Math.min(width, height) / 14);
    const textWidth = font.widthOfTextAtSize(text, fontSize);
    page.drawText(text, {
      x: width / 2 - textWidth / 2.5,
      y: height / 3,
      size: fontSize,
      font,
      color: rgb(0.55, 0.55, 0.55),
      opacity: 0.3,
      rotate: degrees(35),
    });
  }
  return doc.save();
}
