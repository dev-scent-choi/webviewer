/**
 * 아트크래프트 앱 화면(public/artcraft-app.html)의 순수 판정 — tests/artcraftAppLogic.test.ts. 앱과의 계약은 시험(SPIKE-WEBAPP §3)의
 * craft-web-host 그대로: ?host=parent → <app>craft:ready, 부모 → <app>craft:open {name, bytes}, 앱 → :save {name, saveAs, bytes}·:dirty·:failed.
 */
export type AppKind = 'word' | 'deck' | 'grid';

/** 앱이 준비(ready)를 알리기까지 — 넘으면 엔진 보기로(설계 §3 실패 전환) */
export const READY_TIMEOUT_MS = 20_000;
/** 문서를 보낸 뒤 앱 화면에 문서 이름이 보이기까지 — 읽기 전용은 확인 못 하면 보호를 못 걸므로 전환 */
export const OPEN_TIMEOUT_MS = 15_000;

export const NOTICES = {
  failed: '아트크래프트 앱 화면을 열지 못해 다른 화면으로 보여 줍니다',
  readOnlyDeck: '이 화면의 편집은 저장되지 않습니다',
  readOnlyIgnored: '읽기 전용이라 저장하지 않았습니다',
  saveAs: '다른 이름으로 저장·내보내기는 지원하지 않습니다 — Ctrl+S로 저장하세요',
  appError: '아트크래프트 앱에 오류가 났습니다. 저장한 뒤 다시 열어 주세요',
  stale: '문서를 연 뒤에 다른 사람이 먼저 저장했습니다. 내 변경은 저장하지 않았습니다 — 문서를 다시 열어 주세요',
  noBase: '문서를 연 시점 정보가 없어 저장하지 않았습니다. 문서를 다시 열어 주세요',
} as const;

export const SAVE_LABELS = { clean: '', dirty: '저장 안 됨', saving: '저장하는 중…', saved: '저장됨', failed: '저장하지 못했습니다' } as const;
export type SaveState = keyof typeof SAVE_LABELS;

const APP_JS = /^\/artcraft-app\/(word|deck|grid)craft-[0-9a-f]{10}\/(word|deck|grid)craft-web\.js$/;

/** iframe이 불러도 되는 앱 JS — 같은 출처, 해시 폴더, 그 앱 것만(주소를 꾸며 다른 스크립트를 이 출처에서 돌리지 못하게) */
export function isAppJsUrl(kind: AppKind, js: string): boolean {
  const m = APP_JS.exec(js);
  return m !== null && m[1] === kind && m[2] === kind;
}

export function frameUrl(kind: AppKind, js: string, readOnly: boolean): string {
  const q = new URLSearchParams({ app: kind, js, host: 'parent' });
  // 엑셀 읽기 전용은 제목 줄·File 탭을 숨긴다(spike §5)
  if (readOnly && kind === 'grid') q.set('chrome', 'none');
  return `/artcraft-frame.html?${q.toString()}`;
}

export interface AppMessage {
  readonly kind: 'ready' | 'save' | 'dirty' | 'failed';
  readonly data: Record<string, unknown>;
}

/** 앱 → 부모 메시지만 — 같은 출처(origin), 우리 iframe(source), <앱>craft: 이름. 그 밖은 null(무시) */
export function readAppMessage(e: { readonly origin: string; readonly source: unknown; readonly data: unknown }, origin: string, frame: unknown, kind: AppKind): AppMessage | null {
  if (e.origin !== origin || e.source !== frame || typeof e.data !== 'object' || e.data === null) return null;
  const type = (e.data as { type?: unknown }).type;
  const prefix = `${kind}craft:`;
  if (typeof type !== 'string' || !type.startsWith(prefix)) return null;
  const k = type.slice(prefix.length);
  return k === 'ready' || k === 'save' || k === 'dirty' || k === 'failed' ? { kind: k, data: e.data as Record<string, unknown> } : null;
}

/** 읽기 전용으로 열 때 앱에 보낼 명령(spike §5 확인) — PPT는 막는 명령이 없어 1차는 띠만(설계 §3) */
export function readOnlyCommands(kind: AppKind, password: string, sheetCount: number): Array<readonly [string, Record<string, unknown>]> {
  if (kind === 'word') return [['review.restrict', { mode: 'readOnly', password }]];
  if (kind === 'grid') return Array.from({ length: Math.max(1, sheetCount) }, (_, sheet) => ['review.protectSheet', { sheet, password }] as const);
  return [];
}

export type SaveDecision = 'upload' | 'readonly' | 'saveas' | 'empty';

export function saveDecision(data: Record<string, unknown>, canSave: boolean): SaveDecision {
  if (!canSave) return 'readonly';
  if (data.saveAs === true) return 'saveas';
  // 신뢰 경계 — 진짜 바이트(ArrayBuffer·TypedArray)만, byteLength만 흉내 낸 객체는 거절
  const bytes = data.bytes;
  return (bytes instanceof ArrayBuffer || ArrayBuffer.isView(bytes)) && bytes.byteLength > 0 ? 'upload' : 'empty';
}

/** :save의 바이트를 ArrayBuffer로(TypedArray면 그 구간만 복사) — saveDecision이 'upload'라고 본 값에만 */
export function toArrayBuffer(bytes: unknown): ArrayBuffer {
  if (bytes instanceof ArrayBuffer) return bytes;
  const view = bytes as ArrayBufferView;
  return new Uint8Array(view.buffer, view.byteOffset, view.byteLength).slice().buffer;
}

/** 앱에 줄 파일 이름 — 앱은 확장자로 형식을 고르므로 저장된 바이트 형식(meta.ext)으로 */
export function openName(name: string, ext: string): string {
  return `${name.replace(/\.[^./\\]+$/, '')}.${ext}`;
}

/** /artcraft/fonts/SCDream4-0123456789.otf → SCDream4.otf */
export function fontName(url: string): string {
  return (url.split('/').pop() ?? '').replace(/-[0-9a-f]{10}(\.otf)$/, '$1');
}

/** 앱 화면 상태(uiInspect 결과)에 문서 이름이 보이면 열린 것(시험 run.mjs와 같은 판정) */
export function uiShowsDocument(ui: unknown, fileName: string): boolean {
  const base = fileName.replace(/\.[^./\\]+$/, '');
  return base !== '' && JSON.stringify(ui ?? null).includes(JSON.stringify(base).slice(1, -1));
}

/** ui=content로 앱 화면에 오면 — 워드·PPT는 엔진 쪽 그림 보기, 엑셀은 격자 편집기의 콘텐츠만(설계 §4) */
export function contentRedirect(kind: AppKind, token: string): string {
  const q = `?token=${encodeURIComponent(token)}&ui=content`;
  return kind === 'grid' ? `/editor-grid.html${q}` : `/view.html${q}`;
}

/** font.add 명령은 base64를 받는다(spike host.html과 같다) */
export function toBase64(u8: Uint8Array): string {
  let s = '';
  for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode(...u8.subarray(i, i + 0x8000));
  return btoa(s);
}

/** 되돌리기·다시 실행 키(Ctrl/⌘+Z·Y) — 읽기 전용 엑셀에서 막는다. code(KeyZ)가 있으면 한글 입력 상태에서도 안다 */
export function isUndoKey(e: { readonly key: string; readonly code?: string; readonly ctrl: boolean }): boolean {
  if (!e.ctrl) return false;
  const letter = /^Key([A-Z])$/.exec(e.code ?? '')?.[1].toLowerCase() ?? e.key.toLowerCase();
  return letter === 'z' || letter === 'y';
}

/** 엑셀 docInspect() 결과에서 보호가 풀린 시트 번호 — 되돌리기(메뉴·단추 포함)로 풀렸으면 다시 건다 */
export function unprotectedSheets(inspect: unknown): number[] {
  const sheets = (inspect as { sheets?: unknown } | null)?.sheets;
  if (!Array.isArray(sheets)) return [];
  return sheets.flatMap((sheet, i) => ((sheet as { protected?: unknown } | null)?.protected === true ? [] : [i]));
}

/** 서버 engine-source ETag와 같은 꼴 — 저장(200) 뒤엔 원본 = 보낸 바이트라 그 해시가 다음 If-Match */
export async function sha256Etag(bytes: BufferSource): Promise<string> {
  const digest = new Uint8Array(await crypto.subtle.digest('SHA-256', bytes));
  return `"${[...digest].map((b) => b.toString(16).padStart(2, '0')).join('')}"`;
}

/** 저장 실패 안내 — 409 먼저 저장(stale)·428 기준 없음은 다시 열어야 풀린다. 편집 방 열림(409) 등은 서버 문구 그대로 */
export function saveFailure(status: number, error: string | undefined): { text: string; reload: boolean } {
  if (status === 409 && error?.includes('먼저 저장')) return { text: NOTICES.stale, reload: true };
  if (status === 428) return { text: NOTICES.noBase, reload: true };
  return { text: error ?? '서버에 연결할 수 없습니다', reload: false };
}

/**
 * 저장을 한 줄로 — :save가 겹쳐도(Ctrl+S 두 번) 앞 저장이 끝난 뒤 보낸다. If-Match는 보내는 순간의 기준이고,
 * 성공하면 보낸 바이트의 해시가 다음 기준(서버 원본 = 보낸 바이트). 그러지 않으면 둘째가 옛 기준으로 나가 409 stale이 된다
 */
export function createSaveQueue(send: (bytes: ArrayBuffer, ifMatch: string | null) => Promise<boolean>): {
  setBase(etag: string | null): void;
  save(bytes: ArrayBuffer): Promise<boolean>;
} {
  let base: string | null = null;
  let chain: Promise<unknown> = Promise.resolve();
  return {
    setBase(etag) {
      base = etag;
    },
    save(bytes) {
      const run = chain.then(async () => {
        const next = sha256Etag(bytes);
        const ok = await send(bytes, base);
        if (ok) base = await next;
        return ok;
      });
      chain = run.catch(() => undefined);
      return run;
    },
  };
}

/** 상한 안에 끝나지 않으면 message로 거부 — 준비 뒤 열기(글꼴·원본·font.add·확인 전체)가 멈춰 loading에 남지 않게 */
export function withTimeout<T>(p: Promise<T>, ms: number, message: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const limit = new Promise<never>((_, reject) => (timer = setTimeout(() => reject(new Error(message)), ms)));
  return Promise.race([p, limit]).finally(() => clearTimeout(timer));
}

/** 같은 출처 정적 파일 받기 — 200이 아니면 거부(404 HTML을 글꼴·문서로 넘기지 않게) */
export async function fetchBytes(url: string): Promise<Uint8Array> {
  const r = await fetch(url);
  if (!r.ok) throw new Error(`${url} ${r.status}`);
  return new Uint8Array(await r.arrayBuffer());
}

/** 처음 한 번만 — 앱이 ready를 두 번 보내도 열기는 한 번(둘째는 이미 넘긴 바이트로 DataCloneError가 나 화면을 넘겨 버린다) */
export function once(fn: () => void): () => void {
  let done = false;
  return () => {
    if (done) return;
    done = true;
    fn();
  };
}
