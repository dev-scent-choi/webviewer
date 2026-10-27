import path from 'node:path';
import { fileURLToPath } from 'node:url';
import multipart from '@fastify/multipart';
import rateLimit from '@fastify/rate-limit';
import fastifyStatic from '@fastify/static';
import fastifyWebsocket from '@fastify/websocket';
import Fastify, { type FastifyInstance } from 'fastify';
import { versionAssetUrls } from './assetVersion.js';
import { config } from './config.js';
import { adminRoutes } from './routes/admin.js';
import { editRoutes } from './routes/edit.js';
import { documentRoutes } from './routes/documents.js';
import { viewerRoutes } from './routes/viewer.js';

/** 로그용 — 쿼리스트링(토큰)을 지운 주소 */
export function redactQuery(url: string): string {
  const q = url.indexOf('?');
  return q < 0 ? url : `${url.slice(0, q)}?…`;
}

const publicDir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'public');

export async function buildApp(): Promise<FastifyInstance> {
  const app = Fastify({
    logger: {
      level: process.env.LOG_LEVEL ?? 'info',
      // 뷰어·편집기 주소의 ?token= 은 열람 권한 그 자체다 — 로그에는 쿼리 없이 경로만 남긴다
      serializers: {
        req: (req: { method: string; url: string; ip?: string }) => ({ method: req.method, url: redactQuery(req.url), remoteAddress: req.ip }),
      },
    },
    // 기본 본문 상한은 작게 — 파일을 받는 라우트(업로드·hwp 저장·PDF 넣기·파일→PDF)만 라우트에서 크게 연다.
    // 토큰 확인은 핸들러 안에서 하므로, 전역 상한이 크면 인증 없이 큰 본문을 읽게 된다.
    bodyLimit: config.defaultBodyLimit,
    // 앞단 프록시(cloudflared 등 같은 컴퓨터)가 붙인 X-Forwarded-For로 실제 접속 IP를 쓴다 — 요청 수 제한·관리자 IP 검사가
    // 모든 사용자를 127.0.0.1 하나로 보지 않게. 믿는 것은 이 컴퓨터(루프백)에서 온 연결의 헤더뿐이다.
    trustProxy: config.trustProxy,
  });

  // 모든 응답에 기본으로 걸어도 무해한 보안 헤더 — 라우트를 가리지 않고 전역으로 건다.
  // Referrer-Policy: 뷰어/편집기 URL에는 열람 토큰이 쿼리스트링으로 실려 있다(iframe 임베드
  // 특성상 불가피 — README 참고) — 그 페이지가 외부 리소스를 참조할 때 Referer 헤더로 토큰 붙은
  // 전체 URL이 제3자 서버로 새어나가지 않게 한다.
  // X-Content-Type-Options: 브라우저가 Content-Type을 무시하고 응답 바이트를 스니핑해 다른
  // 타입(예: 업로드된 파일을 HTML로 오인)으로 실행하는 것을 막는다.
  app.addHook('onSend', async (_request, reply, payload) => {
    reply.header('Referrer-Policy', 'same-origin');
    reply.header('X-Content-Type-Options', 'nosniff');
    return payload;
  });

  // 이 화면(뷰어/편집기)을 iframe으로 띄울 수 있는 출처를 제한한다 — 설정 안 하면(개발 편의)
  // 아무 출처에서나 임베드할 수 있으니, 운영 환경에서는 DOCVIEW_ALLOWED_ORIGINS를 반드시 채워야
  // 한다(server.ts가 부팅 시 경고). JSON 응답 등 프레임과 무관한 응답에 붙어도 무해하므로
  // 라우트를 가리지 않고 전역으로 건다.
  if (config.allowedOrigins.length > 0) {
    const frameAncestors = `frame-ancestors ${config.allowedOrigins.join(' ')}`;
    app.addHook('onSend', async (_request, reply, payload) => {
      // 라우트가 먼저 단 정책(문서 SVG의 sandbox 등)은 지우지 않고 뒤에 붙인다
      const existing = reply.getHeader('content-security-policy');
      reply.header('Content-Security-Policy', existing ? `${String(existing)}; ${frameAncestors}` : frameAncestors);
      return payload;
    });
  }

  // global:false — 뷰어 화면은 그리드 스크롤·페이지 넘김마다 /api/view/* 를 정상적으로 자주
  // 부르므로, 여기에 전역으로 걸면 실사용 중인 브라우저가 429를 맞는다. 대신 남용 비용이 큰
  // 엔드포인트(업로드·토큰 발급·삭제·관리자 로그인)에만 라우트별로 opt-in한다(각 route 등록의
  // config.rateLimit). 이 앱의 다른 모든 응답이 { success, error } 봉투를 쓰므로 429도 맞춘다.
  await app.register(rateLimit, {
    global: false,
    // 플러그인은 이 함수가 돌려준 값을 그대로 throw한다(node_modules/@fastify/rate-limit/index.js) —
    // Fastify의 기본 에러 핸들러는 그 값의 statusCode 필드를 읽어 실제 HTTP 상태 코드로 쓰므로,
    // 이걸 빠뜨리면 429가 아니라 500(제네릭 에러)으로 응답한다(직접 겪은 버그 — 라이브 서버로
    // 429를 실제로 유발해보고서야 statusCode 누락을 발견했다).
    errorResponseBuilder: (_request, context) => ({
      statusCode: context.statusCode,
      success: false,
      error: `요청이 너무 많습니다. ${context.after} 후 다시 시도하세요.`,
    }),
  });

  // 500 오류의 내부 메시지(파일 경로 등)를 밖에 보이지 않는다 — 로그에만. 4xx는 원래 메시지 그대로.
  app.setErrorHandler((error: Error & { statusCode?: number; validation?: unknown }, request, reply) => {
    const status = error.statusCode && error.statusCode >= 400 ? error.statusCode : 500;
    if (status >= 500) {
      request.log.error(error);
      return reply.code(status).send({ success: false, error: '서버 오류가 발생했습니다.' });
    }
    return reply.code(status).send(error); // 4xx(검증 오류·429 등)는 기본 처리 그대로
  });

  await app.register(multipart, {
    limits: { fileSize: config.maxUploadBytes, files: 1 },
  });
  // POST /api/documents/stream(documents.ts)용 — multipart 경계 없이 파일 바이트를 요청 본문에
  // 그대로 실어 보내는 업로드 경로. Fastify는 등록된 파서가 없는 Content-Type의 본문을 기본
  // 거부하므로(415), 이 타입만 명시적으로 "본문 전체를 버퍼로 그대로 넘긴다"로 등록해 둔다 —
  // JSON/멀티파트 등 기존 타입 파싱에는 전혀 영향이 없다.
  app.addContentTypeParser('application/octet-stream', { parseAs: 'buffer' }, (_request, payload, done) => {
    done(null, payload);
  });
  // 편집 메시지(Yjs 갱신 등) 한 개 상한 — 기본값(100MB)은 인증 전에도 받아 쌓을 수 있어 너무 크다
  await app.register(fastifyWebsocket, { options: { maxPayload: config.wsMaxPayload } });
  // 화면 HTML의 JS·CSS 주소에 내용 해시를 붙여 내보낸다(assetVersion.ts 참고 — 앞단 캐시가 예전 파일을 내주지 않게)
  app.addHook('onSend', async (_request, reply, payload) => {
    if (reply.statusCode !== 200 || !String(reply.getHeader('content-type') ?? '').startsWith('text/html')) return payload;
    let html: string;
    if (typeof payload === 'string') html = payload;
    else if (Buffer.isBuffer(payload)) html = payload.toString('utf8');
    else if (payload && typeof (payload as NodeJS.ReadableStream).pipe === 'function') {
      const chunks: Buffer[] = [];
      for await (const chunk of payload as AsyncIterable<Buffer | string>) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
      html = Buffer.concat(chunks).toString('utf8');
    } else return payload;
    const next = versionAssetUrls(html, publicDir);
    reply.removeHeader('content-length');
    reply.removeHeader('etag');
    return next;
  });
  await app.register(fastifyStatic, {
    root: publicDir,
    // 뷰어/관리 화면 JS가 브라우저에 캐시되어 구버전이 돌지 않도록 항상 재검증시킨다.
    // 단 화면 HTML이 내용 해시(?v=, assetVersion.ts)를 붙여 부른 주소는 내용이 바뀌면 주소도 바뀌므로 오래 보관해도 된다 —
    // 화면을 열 때마다 번들마다 재검증 요청이 가던 것을 없앤다.
    setHeaders(res) {
      const url = (res as { req?: { url?: string } }).req?.url ?? '';
      res.setHeader('cache-control', /[?&]v=[0-9a-f]+/.test(url) ? 'public, max-age=31536000, immutable' : 'no-cache');
    },
  });
  await app.register(documentRoutes);
  await app.register(viewerRoutes);
  await app.register(editRoutes);
  await app.register(adminRoutes);

  return app;
}
