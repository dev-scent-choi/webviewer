/**
 * 아트크래프트 엔진 Worker(문서 하나 = wasm 인스턴스 하나) — DOM·Canvas 없음, RGBA 버퍼를 넘겨준다(spike §7.2).
 * 받는 것은 같은 출처의 정적 자산(/artcraft/*)과 원본(/api/view/engine-source)뿐이다(설계 §6). 작업은 artcraftSource.ts가 하나씩 보낸다.
 */
import { MAX_ORIGINAL_BYTES, originalProblem, stubImports, type PageText, type WorkerReply, type WorkerRequest } from './artcraftLogic.js';

interface Engine {
  readonly memory: WebAssembly.Memory;
  sp_alloc(n: number): number;
  sp_free(ptr: number, n: number): void;
  sp_add_font(ptr: number, len: number, aliasPtr: number, aliasLen: number): number;
  sp_open(namePtr: number, nameLen: number, ptr: number, len: number): number;
  sp_count(h: number): number;
  sp_page_w(h: number, i: number): number;
  sp_page_h(h: number, i: number): number;
  sp_render(h: number, i: number, scaleMilli: number): number;
  sp_text(h: number, i: number): number;
  sp_out_ptr(): number;
  sp_out_len(): number;
  sp_out_w(): number;
  sp_out_h(): number;
  sp_out_clear(): void;
  sp_msg_ptr(): number;
  sp_msg_len(): number;
}

// tsconfig는 DOM 타입만 싣는다 — Worker 전역에서 쓰는 두 가지만 좁게 적는다
const scope = self as unknown as {
  onmessage: ((event: MessageEvent<WorkerRequest>) => void) | null;
  postMessage(message: WorkerReply, transfer?: Transferable[]): void;
};
const encoder = new TextEncoder();
const decoder = new TextDecoder();
let engine: Engine | null = null;
let docHandle = 0;

async function bytesOf(url: string): Promise<Uint8Array> {
  const res = await fetch(url, { credentials: 'same-origin' });
  if (!res.ok) throw new Error(`받지 못했습니다(${res.status}): ${url.split('?')[0]}`);
  // 큰 원본은 받기 전에 거절(압축 전송이면 길이가 작게 오므로 받은 뒤 originalProblem이 다시 본다)
  if (Number(res.headers.get('content-length')) > MAX_ORIGINAL_BYTES) throw new Error('원본이 브라우저 엔진에 너무 큽니다');
  return new Uint8Array(await res.arrayBuffer());
}

function put(e: Engine, bytes: Uint8Array): [number, number] {
  const ptr = e.sp_alloc(bytes.length);
  if (!ptr) throw new Error('브라우저 엔진 메모리가 모자랍니다');
  new Uint8Array(e.memory.buffer, ptr, bytes.length).set(bytes);
  return [ptr, bytes.length];
}

function lastMessage(e: Engine): string {
  return decoder.decode(new Uint8Array(e.memory.buffer, e.sp_msg_ptr(), e.sp_msg_len()));
}

function opened(): Engine {
  if (!engine || !docHandle) throw new Error('문서가 열려 있지 않습니다');
  return engine;
}

async function init(m: Extract<WorkerRequest, { cmd: 'init' }>): Promise<WorkerReply> {
  const t0 = performance.now();
  const wasm = fetch(m.wasm, { credentials: 'same-origin' }).then((res) => {
    if (!res.ok) throw new Error(`엔진을 받지 못했습니다(${res.status})`);
    return res;
  });
  // 컴파일은 받는 동안 함께(compileStreaming — 크롬은 결과를 캐시해 두 번째부터 수십 ms), 글꼴·원본도 동시에 받는다
  const [module, fonts, original] = await Promise.all([WebAssembly.compileStreaming(wasm), Promise.all(m.fonts.map(bytesOf)), bytesOf(m.original)]);
  const instance = await WebAssembly.instantiate(
    module,
    stubImports(WebAssembly.Module.imports(module), (name) => console.warn(`[artcraft] 채우지 않은 바깥 함수가 불렸습니다: ${name}`)),
  );
  const e = instance.exports as unknown as Engine;
  const problem = originalProblem(m.name.slice(m.name.lastIndexOf('.') + 1).toLowerCase(), original);
  if (problem) throw new Error(problem);
  for (const font of fonts) {
    const [ptr, n] = put(e, font);
    const faces = e.sp_add_font(ptr, n, 0, 0);
    e.sp_free(ptr, n);
    // 글꼴이 하나도 들어가지 않으면 한글이 빠진 채 그려진다 — 열기 실패(→ 서버 변환, 설계 §4.2)
    if (faces < 1) throw new Error('글꼴을 넣지 못했습니다');
  }
  const t1 = performance.now();
  const [namePtr, nameLen] = put(e, encoder.encode(m.name));
  const [ptr, n] = put(e, original);
  const h = e.sp_open(namePtr, nameLen, ptr, n);
  e.sp_free(ptr, n);
  e.sp_free(namePtr, nameLen);
  if (!h) throw new Error(lastMessage(e) || '문서를 열지 못했습니다');
  engine = e;
  docHandle = h;
  const count = e.sp_count(h);
  const sizes = Array.from({ length: count }, (_, i) => ({ w: e.sp_page_w(h, i), h: e.sp_page_h(h, i) }));
  return { id: m.id, count, sizes, initMs: Math.round(t1 - t0), openMs: Math.round(performance.now() - t1) };
}

function render(m: Extract<WorkerRequest, { cmd: 'render' }>): { id: number; w: number; h: number; rgba: ArrayBuffer; ms: number } {
  const t0 = performance.now();
  const e = opened();
  if (e.sp_render(docHandle, m.page - 1, Math.round(m.scale * 1000))) throw new Error(lastMessage(e) || '쪽을 그리지 못했습니다');
  const w = e.sp_out_w();
  const h = e.sp_out_h();
  const ptr = e.sp_out_ptr();
  const rgba = e.memory.buffer.slice(ptr, ptr + w * h * 4);
  e.sp_out_clear(); // 큰 RGBA를 엔진 쪽에서 바로 놓는다
  return { id: m.id, w, h, rgba, ms: Math.round(performance.now() - t0) };
}

function text(page: number): PageText {
  const e = opened();
  if (e.sp_text(docHandle, page - 1)) throw new Error(lastMessage(e) || '글자를 꺼내지 못했습니다');
  const json = decoder.decode(new Uint8Array(e.memory.buffer, e.sp_out_ptr(), e.sp_out_len()));
  e.sp_out_clear();
  return JSON.parse(json) as PageText;
}

async function answer(m: WorkerRequest): Promise<void> {
  try {
    if (m.cmd === 'init') scope.postMessage(await init(m));
    else if (m.cmd === 'render') {
      const reply = render(m);
      scope.postMessage(reply, [reply.rgba]);
    } else scope.postMessage({ id: m.id, text: text(m.page) });
  } catch (error) {
    const fatal = error instanceof WebAssembly.RuntimeError;
    if (fatal) engine = null; // trap 뒤 인스턴스는 다시 쓰지 않는다
    scope.postMessage({ id: m.id, error: error instanceof Error ? error.message : String(error), fatal });
  }
}

scope.onmessage = (event) => void answer(event.data);
