import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import type { FastifyInstance } from 'fastify';
import { recordAudit } from '../audit.js';
import { config } from '../config.js';
import { findConverter, supportedExtensions } from '../converters/index.js';
import { ConversionError, type Converter } from '../converters/types.js';
import { getPageCountFromBytes } from '../rendering.js';
import { issueViewToken } from '../security.js';
import { parseCsvToGrid, parseXlsxToGrid, serializeGrid, type GridDocument } from '../spreadsheet.js';
import { newDocumentId, store } from '../storage.js';
import type { DocumentMeta } from '../types.js';

const tokenBodySchema = {
  type: 'object',
  required: ['viewer'],
  additionalProperties: false,
  properties: {
    viewer: { type: 'string', minLength: 1, maxLength: 100 },
    canPrint: { type: 'boolean', default: false },
    canDownload: { type: 'boolean', default: false },
    canCopy: { type: 'boolean', default: false },
  },
} as const;

interface TokenBody {
  readonly viewer: string;
  readonly canPrint: boolean;
  readonly canDownload: boolean;
  readonly canCopy: boolean;
}

/** LibreOffice가 중간 진행률을 제공하지 않으므로 단계 완료 시점마다 진행률을 기록한다. */
const PROGRESS_CONVERTING = 25;
const PROGRESS_PDF_DONE = 85;

/** 그리드 뷰로 여는 스프레드시트 확장자 (PDF는 출력 시점에 지연 생성) */
const GRID_VIEW_EXTS = new Set(['xlsx', 'csv']);

async function buildGrid(meta: DocumentMeta): Promise<GridDocument> {
  const original = await store.readOriginal(meta.id, meta.ext);
  if (meta.ext === 'csv') {
    return parseCsvToGrid(original.toString('utf8'));
  }
  return parseXlsxToGrid(original);
}

/**
 * 스프레드시트는 PDF 변환 없이 셀 데이터만 추출해 즉시 열람 가능하게 한다.
 * (대형 리스트의 수 분짜리 PDF 변환 대기 제거 — PDF는 출력 요청 시 생성)
 */
async function runGridExtraction(app: FastifyInstance, meta: DocumentMeta): Promise<void> {
  try {
    await store.writeMeta({ ...meta, progress: PROGRESS_CONVERTING });
    const grid = await buildGrid(meta);
    if (grid.sheets.length === 0) {
      throw new ConversionError('표시할 시트가 없습니다. 파일이 비어 있거나 지원되지 않는 내용입니다.');
    }
    await store.saveDerived(meta.id, 'grid', serializeGrid(grid));
    await store.writeMeta({
      ...meta,
      status: 'ready',
      viewKind: 'grid',
      pageCount: grid.sheets.length,
      progress: 100,
    });
    await recordAudit({ type: 'converted', docId: meta.id, detail: `${grid.sheets.length} sheets (grid)` });
  } catch (error) {
    app.log.error(error, 'grid extraction failed');
    const userMessage =
      error instanceof ConversionError
        ? error.message
        : '스프레드시트를 읽지 못했습니다. 파일이 손상되었거나 지원되지 않는 형식입니다.';
    await store.writeMeta({ ...meta, status: 'failed', error: userMessage });
    await recordAudit({ type: 'conversion_failed', docId: meta.id, detail: meta.name });
  }
}

/**
 * 변환기는 외부 도구(LibreOffice/LibreDWG) 특성상 파일 경로가 필요하다.
 * 스토리지 모드와 무관하게 임시 작업 폴더에 꺼냈다가 변환 후 즉시 삭제하므로
 * memory 모드에서는 변환 순간 외에 디스크에 남는 것이 없다.
 */
async function runConversion(
  app: FastifyInstance,
  converter: Converter,
  meta: DocumentMeta,
): Promise<void> {
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-convert-'));
  try {
    const sourcePath = path.join(workDir, `source.${meta.ext}`);
    await fs.writeFile(sourcePath, await store.readOriginal(meta.id, meta.ext));
    await store.writeMeta({ ...meta, progress: PROGRESS_CONVERTING });

    const targetPdfPath = path.join(workDir, 'converted.pdf');
    await converter.convert(sourcePath, targetPdfPath, meta.name);

    const pdfBytes = await fs.readFile(targetPdfPath);
    await store.saveDerived(meta.id, 'pdf', pdfBytes);
    await store.writeMeta({ ...meta, progress: PROGRESS_PDF_DONE });

    const pageCount = await getPageCountFromBytes(pdfBytes);
    await store.writeMeta({ ...meta, status: 'ready', viewKind: 'pages', pageCount, progress: 100 });
    await recordAudit({ type: 'converted', docId: meta.id, detail: `${pageCount} pages` });
  } catch (error) {
    app.log.error(error, 'document conversion failed');
    const userMessage =
      error instanceof ConversionError
        ? error.message
        : '문서 변환에 실패했습니다. 파일이 손상되었거나 지원되지 않는 내용입니다.';
    await store.writeMeta({ ...meta, status: 'failed', error: userMessage });
    await recordAudit({ type: 'conversion_failed', docId: meta.id, detail: meta.name });
  } finally {
    await fs.rm(workDir, { recursive: true, force: true });
  }
}

export async function documentRoutes(app: FastifyInstance): Promise<void> {
  app.post('/api/documents', async (request, reply) => {
    let file;
    let data: Buffer;
    try {
      file = await request.file();
      if (!file) {
        return await reply.code(400).send({ success: false, error: '업로드된 파일이 없습니다.' });
      }
      data = await file.toBuffer();
    } catch (error) {
      const code = (error as { code?: string }).code;
      if (code === 'FST_REQ_FILE_TOO_LARGE') {
        return reply.code(413).send({ success: false, error: '파일이 업로드 용량 제한을 초과했습니다.' });
      }
      request.log.error(error, 'upload parsing failed');
      return reply.code(400).send({ success: false, error: '업로드 요청을 처리할 수 없습니다.' });
    }

    const ext = path.extname(file.filename).slice(1).toLowerCase();
    const converter = findConverter(ext);
    if (!converter) {
      return reply.code(415).send({
        success: false,
        error: `지원하지 않는 형식입니다: .${ext || '(없음)'} — 지원 형식: ${supportedExtensions.join(', ')}`,
      });
    }

    const id = newDocumentId();
    try {
      await store.saveOriginal(id, ext, data);
      const meta: DocumentMeta = {
        id,
        name: file.filename,
        ext,
        status: 'processing',
        pageCount: 0,
        createdAt: new Date().toISOString(),
        progress: 10,
      };
      await store.writeMeta(meta);
      await recordAudit({ type: 'uploaded', docId: id, detail: file.filename });
      // 변환은 백그라운드로 진행하고 즉시 응답한다. 클라이언트는 상태를 폴링한다.
      if (GRID_VIEW_EXTS.has(ext)) {
        void runGridExtraction(app, meta);
      } else {
        void runConversion(app, converter, meta);
      }
      return await reply.code(201).send({ success: true, data: meta });
    } catch (error) {
      request.log.error(error, 'upload persisting failed');
      return reply.code(500).send({ success: false, error: '업로드 저장에 실패했습니다.' });
    }
  });

  app.get<{ Params: { id: string } }>('/api/documents/:id', async (request, reply) => {
    const meta = await store.readMeta(request.params.id);
    if (!meta) {
      return reply.code(404).send({ success: false, error: '문서를 찾을 수 없습니다.' });
    }
    return { success: true, data: meta };
  });

  app.get('/api/documents', async () => {
    const documents = await store.listMeta();
    return { success: true, data: documents };
  });

  app.get('/api/formats', async () => ({ success: true, data: supportedExtensions }));

  app.post<{ Params: { id: string }; Body: TokenBody }>(
    '/api/documents/:id/tokens',
    { schema: { body: tokenBodySchema } },
    async (request, reply) => {
      const meta = await store.readMeta(request.params.id);
      if (!meta || meta.status !== 'ready') {
        return reply.code(404).send({ success: false, error: '문서를 찾을 수 없습니다.' });
      }
      const { viewer, canPrint, canDownload, canCopy } = request.body;
      const token = issueViewToken(
        { docId: meta.id, viewer, perms: { canPrint, canDownload, canCopy } },
        config.tokenTtlSeconds,
      );
      await recordAudit({
        type: 'token_issued',
        docId: meta.id,
        viewer,
        detail: `print=${canPrint} download=${canDownload} copy=${canCopy}`,
      });
      return {
        success: true,
        data: {
          token,
          viewerUrl: `/viewer.html?token=${encodeURIComponent(token)}`,
          expiresInSeconds: config.tokenTtlSeconds,
        },
      };
    },
  );
}
