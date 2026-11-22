/**
 * /artcraft/* — 아트크래프트 wasm·글꼴(이름에 내용 해시). 브라우저가 br을 받으면 미리 압축한 .br을(content-encoding: br),
 * 아니면 원본을 내준다. 이름이 내용과 함께 바뀌므로 1년 immutable(설계 §5). manifest.json은 서버만 읽는다.
 */
import { mkdirSync } from 'node:fs';
import fastifyStatic from '@fastify/static';
import type { FastifyInstance } from 'fastify';

const ASSET_PATH = /^\/(fonts\/)?[A-Za-z0-9_]+-[0-9a-f]{10}\.(wasm|otf)$/;

/** /artcraft-app/ 허용 목록 — 해시 폴더 안 접착 JS·wasm, 접착 JS가 상대 경로로 부르는 snippets/<이름>/inline<n>.js만 */
export const APP_ASSET_PATH = /^\/(word|deck|grid)craft-[0-9a-f]{10}\/((word|deck|grid)craft-web(_bg)?\.(js|wasm)|snippets\/[A-Za-z0-9_-]+\/inline[0-9]+\.js)$/;

export async function registerArtcraftStatic(app: FastifyInstance, root: string, prefix = '/artcraft/', assetPath: RegExp = ASSET_PATH): Promise<void> {
  // 빌드 전이라 폴더가 없어도 서버는 뜬다(빈 폴더 = 아트크래프트 꺼짐). 읽기 전용 설치본이라 못 만들면 폴더가 없다는 뜻이니
  // 등록을 건너뛴다(새어 나갈 manifest도 없다)
  try {
    mkdirSync(root, { recursive: true });
  } catch (err) {
    app.log.warn({ err, root }, `artcraft 자산 폴더를 만들 수 없어 ${prefix} 서빙을 건너뜁니다`);
    return;
  }
  await app.register(fastifyStatic, {
    root,
    prefix,
    decorateReply: false,
    preCompressed: true,
    index: false,
    // 허용 목록: 해시 이름 wasm·otf만(.br은 preCompressed가 붙이므로 직접 요청은 막힌다). 거부 목록이면 //·대소문자로 manifest가 샌다
    allowedPath: (pathName) => assetPath.test(pathName),
    // cache-control은 send가 setHeaders 뒤에 덮어쓰므로 send 옵션으로 준다
    maxAge: '365d',
    immutable: true,
    setHeaders(res) {
      res.setHeader('vary', 'accept-encoding');
    },
  });
}
