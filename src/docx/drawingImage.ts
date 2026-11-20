import { createHash } from 'node:crypto';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { strFromU8, strToU8, zipSync } from 'fflate';
import { SizedLruCache } from '../cache.js';
import { officeConverter } from '../converters/libreOffice.js';
import { renderPage } from '../rendering.js';
import { guardedUnzipSync } from '../zipGuard.js';
import { importDocx } from './import.js';
import { refHashOf } from './ooxml.js';

/**
 * 브라우저가 못 그리는 그림 조각(EMF/WMF 그림·차트·도형 묶음·도형)을 LibreOffice로 그림(PNG)으로 바꾼다. 문서 한 판(내용 해시)마다
 * 그런 조각을 **모두 한 번에** — 조각마다 한 쪽(구역, 종이 = 조각 크기)인 docx 하나를 원본 파트(그림·차트 파트·관계)를 그대로 지닌 채
 * 만들어 PDF로 한 번 바꾸고, 요청이 오면 그 쪽을 그린다. 원본 문서는 바꾸지 않는다(저장 땐 원본 조각을 그대로 쓴다).
 * 요청마다 문서를 다시 읽지 않는다: 그린 그림 → 문서 판의 PDF → 문서 판의 조각 목록 순으로 메모리에서 먼저 찾는다.
 */
const EMU_PER_TWIP = 635;

interface DrawingBatch {
  /** 조각 해시 → PDF 쪽 번호(1부터) */
  readonly pages: ReadonlyMap<string, number>;
  /** 쪽 크기(pt) — 만들 때 정한 종이 크기라 PDF를 다시 읽어 셀 필요가 없다(200쪽 PDF 읽기는 주 스레드를 1초 막는다) */
  readonly sizes: ReadonlyArray<{ width: number; height: number }>;
  readonly pdf: Buffer;
  readonly key: string;
}

/** 문서 판(docId + 내용 해시) → 조각 목록·저장소 — 같은 판은 한 번만 가져온다(400쪽 문서 가져오기 약 0.65초) */
const versions = new SizedLruCache<{ fragments: ReadonlyMap<string, string>; drawn: readonly string[] }>({ maxBytes: 64 * 1024 * 1024, maxEntries: 50, sizeOf: () => 1024 * 1024 });
/** 문서 판의 그림 묶음 PDF */
const batches = new SizedLruCache<DrawingBatch>({ maxBytes: 128 * 1024 * 1024, maxEntries: 50, sizeOf: (b) => b.pdf.length });
/** 그린 그림 — 문서 id + 조각 해시(같은 조각이라도 문서마다 그림 파트가 다르다) */
const pngs = new SizedLruCache<Buffer>({ maxBytes: 64 * 1024 * 1024, maxEntries: 2000, sizeOf: (b) => b.length });

/** 조각 크기의 종이(twip) — 크기를 모르면(VML) A6쯤 */
function pageTwipsOf(runXml: string): { w: number; h: number } {
  const cx = Number(/<wp:extent\b[^>]*\bcx="(\d+)"/.exec(runXml)?.[1] ?? 0);
  const cy = Number(/<wp:extent\b[^>]*\bcy="(\d+)"/.exec(runXml)?.[1] ?? 0);
  return { w: cx > 0 ? Math.max(240, Math.ceil(cx / EMU_PER_TWIP) + 40) : 5954, h: cy > 0 ? Math.max(240, Math.ceil(cy / EMU_PER_TWIP) + 40) : 8391 };
}

/** 떠 있는 개체를 종이 왼쪽 위로 옮긴 조각과, 그 조각 크기의 종이(구역 나눔) */
function drawingPage(runXml: string, last: boolean): string {
  const placed = runXml
    .replace(/<wp:positionH\b[^>]*>[\s\S]*?<\/wp:positionH>/, '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>')
    .replace(/<wp:positionV\b[^>]*>[\s\S]*?<\/wp:positionV>/, '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>');
  const { w, h } = pageTwipsOf(runXml);
  const sectPr = `<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>`;
  const spacing = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
  return last ? `<w:p><w:pPr>${spacing}</w:pPr>${placed}</w:p>${sectPr}` : `<w:p><w:pPr>${spacing}${sectPr}</w:pPr>${placed}</w:p>`;
}

async function renderBatch(docxBytes: Uint8Array, fragments: ReadonlyMap<string, string>, hashes: readonly string[], key: string): Promise<DrawingBatch> {
  const entries = guardedUnzipSync(docxBytes, 'renderDocxDrawings');
  const documentXml = entries['word/document.xml'] ? strFromU8(entries['word/document.xml']) : '';
  const openTag = /<w:document\b[^>]*>/.exec(documentXml)?.[0] ?? '<w:document>';
  const body = hashes.map((hash, i) => drawingPage(fragments.get(hash)!, i === hashes.length - 1)).join('');
  entries['word/document.xml'] = strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${openTag}<w:body>${body}</w:body></w:document>`);
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-drawing-'));
  try {
    const source = path.join(dir, 'drawing.docx');
    const target = path.join(dir, 'drawing.pdf');
    await fs.writeFile(source, zipSync(entries));
    await officeConverter.convert(source, target);
    const sizes = hashes.map((hash) => {
      const { w, h } = pageTwipsOf(fragments.get(hash)!);
      return { width: w / 20, height: h / 20 };
    });
    return { pages: new Map(hashes.map((hash, i) => [hash, i + 1])), sizes, pdf: await fs.readFile(target), key };
  } finally {
    await fs.rm(dir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 }).catch(() => undefined);
  }
}

/** 쪽 그리기(pdf.js, 주 스레드)는 한 번에 하나씩, 사이사이 다른 요청을 받는다 — 그림 200개를 한꺼번에 요청해도 서버가 멈추지 않게 */
let renderChain: Promise<unknown> = Promise.resolve();
function queued<T>(task: () => Promise<T>): Promise<T> {
  const run = renderChain.then(() => new Promise<void>((resolve) => setImmediate(resolve))).then(task);
  renderChain = run.catch(() => undefined);
  return run;
}

/** 이미 그려 둔 그림(문서 판과 무관하게 문서 id + 조각 해시) — 있으면 상한 없이 바로 돌려준다 */
export function cachedDocxDrawing(docId: string, ref: string): Promise<Buffer> | undefined {
  return pngs.get(`${docId}:${refHashOf(ref)}`);
}

/**
 * ref: 편집기 노드의 조각 참조 — 이 문서(지금 원본, 열려 있는 방이면 그 방의 저장소까지)에 있는 조각만 그린다(없으면 null).
 * roomFragments: 열린 방의 보존 조각 저장소(지웠다 되돌린 조각도 그리게).
 */
export async function renderDocxDrawingPng(
  loadOriginal: () => Promise<Uint8Array>,
  docId: string,
  ref: string,
  roomFragments: ReadonlyMap<string, string> | null,
): Promise<Buffer | null> {
  const hash = refHashOf(ref);
  const pngKey = `${docId}:${hash}`;
  const cached = pngs.get(pngKey);
  if (cached) return cached;
  const bytes = await loadOriginal();
  const version = `${docId}:${createHash('sha1').update(bytes).digest('base64url')}`;
  const info = await versions.getOrLoad(version, async () => {
    const imported = importDocx(bytes);
    return { fragments: imported.fragments, drawn: [...imported.serverDrawn] };
  });
  // 지금 원본엔 없지만 방 저장소엔 있는 조각(지웠다 저장 뒤 되돌린 것) — 그 하나만 따로 묶음을 만든다
  const inFile = info.drawn.includes(hash);
  const fragments = inFile ? info.fragments : roomFragments;
  if (!fragments?.get(hash)) return null;
  const batchKey = inFile ? `${version}:all` : `${version}:${hash}`;
  batchedDocs.add(docId);
  const batch = await batches.getOrLoad(batchKey, () => renderBatch(bytes, fragments, inFile ? info.drawn : [hash], batchKey));
  const page = batch.pages.get(hash);
  if (!page) return null;
  return pngs.getOrLoad(pngKey, () => queued(() => renderPage(`docx-drawing:${batch.key}`, page, 2, async () => batch.pdf, { load: async () => batch.sizes, save: async () => undefined })));
}

/** 그 문서 판의 묶음 PDF가 이미 있거나 만드는 중인지 — 있으면 그리기 상한을 세지 않는다(쪽 하나 그리기는 가볍다) */
export function hasDocxDrawingBatch(docId: string): boolean {
  return batchedDocs.has(docId);
}
const batchedDocs = new Set<string>();
