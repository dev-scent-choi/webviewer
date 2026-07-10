import fs from 'node:fs/promises';
import type { FastifyInstance, FastifyReply } from 'fastify';
import { recordAudit } from '../audit.js';
import { PageOutOfRangeError, buildPrintPdf, renderPage } from '../rendering.js';
import { InvalidTokenError, verifyViewToken } from '../security.js';
import { convertedPath, originalPath, readMeta } from '../storage.js';
import type { DocumentMeta, ViewClaims } from '../types.js';

interface ViewContext {
  readonly claims: ViewClaims;
  readonly meta: DocumentMeta;
}

/**
 * 토큰 검증 + 문서 조회. 실패 시 reply를 직접 전송하고 null을 반환한다.
 * 토큰은 경로가 아닌 쿼리 파라미터로 받는다(JWT의 `.`이 경로 라우팅을 깨뜨리고
 * 액세스 로그에 토큰이 남는 문제를 피하기 위함).
 */
async function resolveView(token: string | undefined, reply: FastifyReply): Promise<ViewContext | null> {
  if (!token) {
    await reply.code(401).send({ success: false, error: '열람 토큰이 없습니다.' });
    return null;
  }
  let claims: ViewClaims;
  try {
    claims = verifyViewToken(token);
  } catch (error) {
    if (error instanceof InvalidTokenError) {
      await reply.code(401).send({ success: false, error: error.message });
      return null;
    }
    throw error;
  }
  const meta = await readMeta(claims.docId);
  if (!meta || meta.status !== 'ready') {
    await reply.code(404).send({ success: false, error: '문서를 찾을 수 없습니다.' });
    return null;
  }
  return { claims, meta };
}

// token 누락/무효는 resolveView가 401로 일관 처리하도록 스키마에서 required로 강제하지 않는다.
const tokenQuerySchema = {
  type: 'object',
  properties: { token: { type: 'string' } },
} as const;

const pageRequestSchema = {
  params: {
    type: 'object',
    required: ['page'],
    properties: { page: { type: 'integer', minimum: 1 } },
  },
  querystring: {
    type: 'object',
    properties: {
      token: { type: 'string' },
      scale: { type: 'integer', minimum: 1, maximum: 16, default: 2 },
    },
  },
} as const;

interface TokenQuery {
  readonly token?: string;
}

interface PageQuery extends TokenQuery {
  readonly scale: number;
}

export async function viewerRoutes(app: FastifyInstance): Promise<void> {
  app.get<{ Querystring: TokenQuery }>(
    '/api/view/info',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      await recordAudit({ type: 'opened', docId: view.meta.id, viewer: view.claims.viewer });
      return {
        success: true,
        data: {
          name: view.meta.name,
          pageCount: view.meta.pageCount,
          permissions: view.claims.perms,
          viewer: view.claims.viewer,
        },
      };
    },
  );

  app.get<{ Params: { page: number }; Querystring: PageQuery }>(
    '/api/view/pages/:page',
    { schema: pageRequestSchema },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      try {
        const png = await renderPage(
          convertedPath(view.meta.id),
          request.params.page,
          request.query.scale,
        );
        return await reply.type('image/png').header('cache-control', 'private, no-store').send(png);
      } catch (error) {
        if (error instanceof PageOutOfRangeError) {
          return reply.code(404).send({ success: false, error: error.message });
        }
        request.log.error(error, 'page rendering failed');
        return reply.code(500).send({ success: false, error: '페이지 렌더링에 실패했습니다.' });
      }
    },
  );

  app.get<{ Querystring: TokenQuery }>(
    '/api/view/print',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canPrint) {
        await recordAudit({ type: 'print_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '출력 권한이 없습니다.' });
      }
      const watermark = `${view.claims.viewer}  ${new Date().toISOString().slice(0, 10)}`;
      const pdfBytes = await buildPrintPdf(convertedPath(view.meta.id), watermark);
      await recordAudit({ type: 'printed', docId: view.meta.id, viewer: view.claims.viewer });
      return reply
        .type('application/pdf')
        .header('content-disposition', 'inline; filename="print.pdf"')
        .send(Buffer.from(pdfBytes));
    },
  );

  app.get<{ Querystring: TokenQuery }>(
    '/api/view/download',
    { schema: { querystring: tokenQuerySchema } },
    async (request, reply) => {
      const view = await resolveView(request.query.token, reply);
      if (!view) return reply;
      if (!view.claims.perms.canDownload) {
        await recordAudit({ type: 'download_denied', docId: view.meta.id, viewer: view.claims.viewer });
        return reply.code(403).send({ success: false, error: '저장(다운로드) 권한이 없습니다.' });
      }
      const data = await fs.readFile(originalPath(view.meta.id, view.meta.ext));
      await recordAudit({ type: 'downloaded', docId: view.meta.id, viewer: view.claims.viewer });
      const encodedName = encodeURIComponent(view.meta.name);
      return reply
        .type('application/octet-stream')
        .header('content-disposition', `attachment; filename*=UTF-8''${encodedName}`)
        .send(data);
    },
  );
}
