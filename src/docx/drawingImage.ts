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
 * 브라우저가 못 그리는 그림 조각(EMF/WMF 그림·차트·도형 묶음·도형)을 LibreOffice로 그림(PNG)으로 바꾼다 — 그 조각 하나만 담은 docx를
 * 원본 파트(그림·차트 파트·관계)를 그대로 지닌 채 만들어, 종이를 조각 크기로 하고 PDF로 바꾼 뒤 첫 쪽을 그린다. 원본 문서는 바꾸지 않는다
 * (편집기는 보여 주기만 하고 저장 땐 원본 조각을 그대로 쓴다). 결과는 조각 내용(해시)으로 메모리에 둔다.
 */
const cache = new SizedLruCache<Buffer>({ maxBytes: 64 * 1024 * 1024, maxEntries: 500, sizeOf: (b) => b.length });

const EMU_PER_TWIP = 635;

/** 떠 있는 개체를 종이 왼쪽 위로 옮기고, 종이를 개체 크기로 — 그 조각만 담은 document.xml */
function drawingDocumentXml(originalDocumentXml: string, runXml: string): string {
  const openTag = /<w:document\b[^>]*>/.exec(originalDocumentXml)?.[0] ?? '<w:document>';
  const cx = Number(/<wp:extent\b[^>]*\bcx="(\d+)"/.exec(runXml)?.[1] ?? 0);
  const cy = Number(/<wp:extent\b[^>]*\bcy="(\d+)"/.exec(runXml)?.[1] ?? 0);
  const placed = runXml
    .replace(/<wp:positionH\b[^>]*>[\s\S]*?<\/wp:positionH>/, '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>')
    .replace(/<wp:positionV\b[^>]*>[\s\S]*?<\/wp:positionV>/, '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>');
  // 크기를 모르면(VML) A6쯤 — 그린 뒤 여백은 보이지 않을 만큼 작다
  const w = Math.max(240, Math.ceil(cx / EMU_PER_TWIP) + 40) || 5954;
  const h = Math.max(240, Math.ceil(cy / EMU_PER_TWIP) + 40) || 8391;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${openTag}<w:body><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>${placed}</w:p>` +
    `<w:sectPr><w:pgSz w:w="${w}" w:h="${h}"/><w:pgMar w:top="0" w:right="0" w:bottom="0" w:left="0" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>`
  );
}

/** ref: 편집기 노드의 조각 참조 — 이 문서에 실제로 있는 조각만 그린다(없으면 null). docKey: 문서 id(같은 조각이라도 문서마다 그림 파트가 다르다) */
export async function renderDocxDrawingPng(docxBytes: Uint8Array, ref: string, docKey: string): Promise<Buffer | null> {
  const hash = refHashOf(ref);
  const runXml = importDocx(docxBytes).fragments.get(hash);
  if (!runXml) return null;
  const key = `docx-drawing:${docKey}:${hash}`;
  return cache.getOrLoad(key, async () => {
    const entries = guardedUnzipSync(docxBytes, 'renderDocxDrawingPng');
    const documentXml = entries['word/document.xml'] ? strFromU8(entries['word/document.xml']) : '';
    entries['word/document.xml'] = strToU8(drawingDocumentXml(documentXml, runXml));
    const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-drawing-'));
    try {
      const source = path.join(dir, 'drawing.docx');
      const target = path.join(dir, 'drawing.pdf');
      await fs.writeFile(source, zipSync(entries));
      await officeConverter.convert(source, target);
      const pdf = await fs.readFile(target);
      return await renderPage(key, 1, 2, async () => pdf, { load: async () => null, save: async () => undefined });
    } finally {
      await fs.rm(dir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 }).catch(() => undefined);
    }
  });
}
