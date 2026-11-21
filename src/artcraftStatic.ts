/**
 * /artcraft/* — 아트크래프트 wasm·글꼴(이름에 내용 해시). 브라우저가 br을 받으면 미리 압축한 .br을(content-encoding: br),
 * 아니면 원본을 내준다. 이름이 내용과 함께 바뀌므로 1년 immutable(설계 §5). manifest.json은 서버만 읽는다.
 */
import { mkdirSync } from 'node:fs';
import fastifyStatic from '@fastify/static';
import type { FastifyInstance } from 'fastify';

export async function registerArtcraftStatic(app: FastifyInstance, root: string): Promise<void> {
  // 빌드 전이라 폴더가 없어도 서버는 뜬다(빈 폴더 = 아트크래프트 꺼짐)
  mkdirSync(root, { recursive: true });
  await app.register(fastifyStatic, {
    root,
    prefix: '/artcraft/',
    decorateReply: false,
    preCompressed: true,
    index: false,
    allowedPath: (pathName) => pathName !== '/manifest.json',
    // cache-control은 send가 setHeaders 뒤에 덮어쓰므로 send 옵션으로 준다
    maxAge: '365d',
    immutable: true,
    setHeaders(res) {
      res.setHeader('vary', 'accept-encoding');
    },
  });
}
