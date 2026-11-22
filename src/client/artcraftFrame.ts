/**
 * 아트크래프트 앱 iframe 문서(public/artcraft-frame.html) — 우리가 빌드한 같은 출처 앱 JS만 불러 캔버스에 띄운다(spike app-*.html과 같은 꼴).
 * 앱은 이 문서 주소의 ?host=parent(·chrome=none)를 읽는다. 불러오기·시작이 실패하면 부모에 <앱>craft:failed — 부모가 엔진 보기로 넘긴다.
 * sandbox 없음(WebGPU 필요, 설계 §3 보안) — 그래서 주소의 js를 허용 목록으로만 받는다.
 */
import { isAppJsUrl, type AppKind } from './artcraftAppLogic.js';

const q = new URLSearchParams(location.search);
const kind = (q.get('app') ?? '') as AppKind;
const js = q.get('js') ?? '';
const failToParent = (error: string): void => {
  if (window.parent !== window) window.parent.postMessage({ type: `${kind}craft:failed`, error }, location.origin);
};

if (!isAppJsUrl(kind, js)) {
  document.body.textContent = '잘못된 주소입니다.';
  failToParent('허용되지 않은 앱 주소');
} else {
  const canvas = document.createElement('canvas');
  canvas.id = `${kind}craft_canvas`;
  const loading = document.createElement('div');
  loading.id = `${kind}craft_loading`;
  loading.className = 'app-loading';
  loading.textContent = '불러오는 중…';
  document.body.append(canvas, loading);
  import(/* 주소는 위 허용 목록으로 확인 */ js)
    .then((m: { default: () => Promise<unknown> }) => m.default())
    .catch((error: unknown) => failToParent(String(error)));
}
