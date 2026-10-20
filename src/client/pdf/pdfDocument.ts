/**
 * pdf.js 열기와 문서 정보 읽기 — 화면과 분리된 얇은 래퍼(PDF 편집기 전용).
 * 보안: PDF 안 자바스크립트는 실행하지 않고(enableScripting 안 씀), eval도 막는다. worker·cMap·표준 글꼴은 자체 호스팅
 * (build-editor.mjs가 node_modules/pdfjs-dist에서 복사) — cMap이 있어야 글꼴이 안 들어 있는 한글 PDF(Adobe-Korea1)도 그린다.
 */
import * as pdfjs from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';

pdfjs.GlobalWorkerOptions.workerSrc = '/js/pdf.worker.mjs';

export { pdfjs };
export type { PDFDocumentProxy };

export async function openPdf(url: string): Promise<PDFDocumentProxy> {
  return pdfjs.getDocument({
    url,
    cMapUrl: '/pdfjs/cmaps/',
    cMapPacked: true,
    standardFontDataUrl: '/pdfjs/standard_fonts/',
    isEvalSupported: false,
    withCredentials: false,
  }).promise;
}

export interface OutlineNode {
  readonly title: string;
  /** 1부터 센 쪽 번호(목적지를 풀 수 없으면 null) */
  readonly page: number | null;
  readonly children: readonly OutlineNode[];
}

type RawOutline = Awaited<ReturnType<PDFDocumentProxy['getOutline']>>;

async function pageOfDest(doc: PDFDocumentProxy, dest: unknown): Promise<number | null> {
  try {
    const explicit = typeof dest === 'string' ? await doc.getDestination(dest) : dest;
    if (!Array.isArray(explicit) || explicit.length === 0) return null;
    const ref = explicit[0];
    if (typeof ref === 'number') return ref + 1;
    if (ref && typeof ref === 'object') return (await doc.getPageIndex(ref as Parameters<PDFDocumentProxy['getPageIndex']>[0])) + 1;
  } catch {
    // 깨진 목적지 — 쪽 없이 제목만 보인다
  }
  return null;
}

export async function readOutline(doc: PDFDocumentProxy): Promise<OutlineNode[]> {
  const raw: RawOutline = await doc.getOutline();
  if (!raw) return [];
  const walk = async (items: NonNullable<RawOutline>, depth: number): Promise<OutlineNode[]> =>
    Promise.all(
      items.map(async (item) => ({
        title: item.title || '(제목 없음)',
        page: await pageOfDest(doc, item.dest),
        children: depth < 32 && item.items?.length ? await walk(item.items, depth + 1) : [],
      })),
    );
  return walk(raw, 0);
}

export interface LayerInfo {
  readonly id: string;
  readonly name: string;
  readonly visible: boolean;
}

export type OptionalContentConfig = Awaited<ReturnType<PDFDocumentProxy['getOptionalContentConfig']>>;

/** 레이어(선택적 콘텐츠 그룹) 목록 — 문서가 정한 표시 순서(getOrder, 중첩 배열을 펼침)대로, 순서에 없는 그룹은 뒤에 */
export function readLayers(config: OptionalContentConfig): LayerInfo[] {
  const groups = (config.getGroups() ?? {}) as Record<string, { name?: string | null; visible?: boolean }>;
  const flatten = (items: unknown[]): string[] =>
    items.flatMap((item) => (Array.isArray(item) ? flatten(item) : typeof item === 'string' ? [item] : []));
  const ordered = flatten((config.getOrder() as unknown[] | null) ?? []).filter((id) => id in groups);
  const ids = [...new Set([...ordered, ...Object.keys(groups)])];
  return ids.map((id) => ({ id, name: groups[id].name || '(이름 없는 레이어)', visible: groups[id].visible !== false }));
}
