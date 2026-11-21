import { describe, expect, test, vi } from 'vitest';
import {
  ByteLru,
  JobQueue,
  MAX_ORIGINAL_BYTES,
  drawPlan,
  originalProblem,
  pageCountSuspicious,
  quickPixelWidth,
  runHits,
  runSpanGeom,
  stubImports,
  targetPixelWidth,
  type PageText,
} from '../src/client/view/artcraftLogic.js';

const LETTER = { w: 612, h: 792 };

describe('배율', () => {
  test('요청 화소 폭 = 쪽 CSS 폭 × dpr(캔버스 상한 안)', () => {
    expect(targetPixelWidth(800, 2, LETTER)).toBe(1600);
    expect(targetPixelWidth(800, 1, LETTER)).toBe(800);
  });
  test('아주 크게 확대해도 캔버스 한 장 1677만 화소를 넘지 않는다', () => {
    const px = targetPixelWidth(5000, 2, LETTER);
    expect(px).toBeGreaterThan(3500);
    expect(px * Math.ceil(px * (792 / 612))).toBeLessThanOrEqual(16_777_216 * 1.002);
  });
  test('1배 그림 화소 폭 = 쪽 폭 pt', () => {
    expect(quickPixelWidth(LETTER)).toBe(612);
    expect(quickPixelWidth({ w: 0.2, h: 1 })).toBe(1);
  });
});

describe('두 단계 그리기 판단', () => {
  test('같은 화소 폭이 이미 있으면 그대로', () => {
    expect(drawPlan([1600], 1600, 612)).toBe('cached');
  });
  test('선명한 그림이 아직 없고 1.5배를 넘으면 1배 먼저', () => {
    expect(drawPlan([], 1600, 612)).toBe('two-stage');
    expect(drawPlan([612], 1600, 612)).toBe('two-stage');
  });
  test('이미 선명한 그림이 있으면(확대 바꿈) 바로 요청 배율', () => {
    expect(drawPlan([1200], 1600, 612)).toBe('direct');
  });
  test('요청이 1.5배 이하면 바로', () => {
    expect(drawPlan([], 900, 612)).toBe('direct');
  });
});

describe('찾기 — run 상자 안에서 글자 수 비율로 자른 사각형', () => {
  const text: PageText = {
    w: 400,
    h: 800,
    runs: [
      { t: '가나다라', x0: 100, x1: 140, y: 200, s: 10 },
      { t: 'Hello World', x0: 0, x1: 110, y: 400, s: 20 },
    ],
  };
  test('한글 가운데 두 글자', () => {
    expect(runHits(3, text, '나다')).toEqual([{ page: 3, rects: [{ x: 0.275, y: 0.24, w: 0.05, h: 0.0125 }] }]);
  });
  test('대소문자 무시, 한 run에 여러 번', () => {
    const hits = runHits(1, text, 'O');
    expect(hits).toHaveLength(2);
    expect(hits[0].rects[0].x).toBeCloseTo(40 / 400);
    expect(runHits(1, text, 'world')[0].rects[0]).toMatchObject({ y: (400 - 16) / 800, h: 20 / 800 });
  });
  test('빈 검색어·없는 글자는 결과 없음', () => {
    expect(runHits(1, text, '  ')).toEqual([]);
    expect(runHits(1, text, '없음')).toEqual([]);
  });
});

describe('글자 층 좌표', () => {
  test('쪽 비율 % 위치, CSS 화소 글자 크기·폭', () => {
    expect(runSpanGeom({ t: '가나', x0: 100, x1: 140, y: 200, s: 10 }, { w: 400, h: 800 }, 800)).toEqual({ left: 25, top: 24, fontPx: 20, widthPx: 80 });
  });
});

describe('.doc(실험) 쪽 수 전환 규칙', () => {
  test('서버 쪽 수와 2쪽 넘게 다르면 전환, 2쪽까지는 그대로', () => {
    expect(pageCountSuspicious(true, 5, 8)).toBe(true);
    expect(pageCountSuspicious(true, 5, 7)).toBe(false);
    expect(pageCountSuspicious(true, 9, 6)).toBe(true);
  });
  test('실험이 아니거나 서버 쪽 수를 모르면(0) 보지 않는다', () => {
    expect(pageCountSuspicious(false, 1, 9)).toBe(false);
    expect(pageCountSuspicious(true, 5, 0)).toBe(false);
  });
});

describe('그림 LRU(바이트 상한)', () => {
  test('넘으면 오래 안 쓴 것부터 버리고, 방금 쓴 것은 남긴다', () => {
    const lru = new ByteLru<{ n: number }>(10, (v) => v.n);
    lru.set('1:100', { n: 4 });
    lru.set('2:100', { n: 4 });
    expect(lru.get('1:100')).toEqual({ n: 4 }); // 1을 최근으로
    lru.set('3:100', { n: 4 }); // 12 > 10 → 2를 버림
    expect(lru.keys()).toEqual(['1:100', '3:100']);
    expect(lru.bytes).toBe(8);
  });
  test('하나가 상한보다 커도 그것은 남긴다', () => {
    const lru = new ByteLru<{ n: number }>(10, (v) => v.n);
    lru.set('a', { n: 3 });
    lru.set('b', { n: 50 });
    expect(lru.keys()).toEqual(['b']);
  });
  test('같은 키를 다시 넣으면 크기를 바꿔 센다', () => {
    const lru = new ByteLru<{ n: number }>(100, (v) => v.n);
    lru.set('a', { n: 30 });
    lru.set('a', { n: 10 });
    expect(lru.bytes).toBe(10);
  });
});

describe('작업 줄 — 보이는 쪽(급함) 먼저, 호출 순서 유지, 취소는 빼기', () => {
  test('급한 것이 먼저, 같은 급끼리는 넣은 순서', () => {
    const q = new JobQueue<string>();
    q.push('선명하게 1', false);
    q.push('1쪽', true);
    q.push('2쪽', true);
    expect([q.shift(), q.shift(), q.shift(), q.shift()]).toEqual(['1쪽', '2쪽', '선명하게 1', undefined]);
  });
  test('remove는 줄에 있던 것만 true', () => {
    const q = new JobQueue<string>();
    q.push('a', true);
    q.push('b', false);
    expect(q.remove('b')).toBe(true);
    expect(q.remove('b')).toBe(false);
    expect(q.size).toBe(1);
  });
});

describe('바깥 함수 채우기(wasm-bindgen 없이)', () => {
  const descs: WebAssembly.ModuleImportDescriptor[] = [
    { module: 'wbg', name: '__wbg_now_2c95c9de', kind: 'function' },
    { module: 'wbg', name: '__wbg_getTimezoneOffset_6b5752021c499c47', kind: 'function' },
    { module: 'wbg', name: '__wbg_unknown_99', kind: 'function' },
    { module: 'wbg', name: '__wbindgen_throw', kind: 'function' },
    { module: 'env', name: 'memory', kind: 'memory' },
  ];
  test('시각·시간대만 진짜, 나머지는 0 + 이름마다 경고 한 번', () => {
    const warn = vi.fn();
    const imports = stubImports(descs, warn) as unknown as Record<string, Record<string, () => number>>;
    expect(Math.abs(imports.wbg.__wbg_now_2c95c9de() - Date.now())).toBeLessThan(1000);
    expect(imports.wbg.__wbg_getTimezoneOffset_6b5752021c499c47()).toBe(new Date().getTimezoneOffset());
    expect(imports.wbg.__wbg_unknown_99()).toBe(0); // "unknown" 안의 now에 걸리지 않는다
    imports.wbg.__wbg_unknown_99();
    expect(imports.wbg.__wbindgen_throw()).toBe(0);
    expect(warn.mock.calls).toEqual([['__wbg_unknown_99'], ['__wbindgen_throw']]);
    expect(imports.env).toBeUndefined();
  });
});

describe('엔진에 넘기기 전 원본 검사(빈 입력을 엔진이 빈 문서로 여는 것 막기)', () => {
  const zip = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 1, 2, 3]);
  const ole = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0]);
  test('docx·pptx는 zip, doc는 OLE 머리여야 통과', () => {
    expect(originalProblem('docx', zip)).toBeNull();
    expect(originalProblem('pptx', zip)).toBeNull();
    expect(originalProblem('doc', ole)).toBeNull();
    expect(originalProblem('doc', zip)).toMatch(/형식/);
    expect(originalProblem('docx', ole)).toMatch(/형식/);
  });
  test('빈 파일·쓰레기·너무 큰 파일은 거절', () => {
    expect(originalProblem('docx', new Uint8Array())).toMatch(/비어/);
    expect(originalProblem('docx', new TextEncoder().encode('<html>'))).toMatch(/형식/);
    expect(originalProblem('docx', { length: MAX_ORIGINAL_BYTES + 1, subarray: () => zip } as unknown as Uint8Array)).toMatch(/큽니다/);
  });
});

describe('Worker 시간 제한(WorkerClient)', () => {
  test('답이 없으면 제한 시간에 ArtcraftUnavailable로 끝내고 Worker를 끈다, 그 뒤 호출도 거절', async () => {
    vi.useFakeTimers();
    try {
      const { WorkerClient, ArtcraftUnavailable, RENDER_TIMEOUT_MS } = await import('../src/client/view/artcraftSource.js');
      const worker = { postMessage: vi.fn(), terminate: vi.fn(), onmessage: null, onerror: null } as unknown as Worker;
      const client = new WorkerClient(worker);
      const first = client.call({ cmd: 'text', page: 1 }, RENDER_TIMEOUT_MS);
      const queued = client.call({ cmd: 'text', page: 2 }, RENDER_TIMEOUT_MS);
      const settled = Promise.allSettled([first, queued]);
      vi.advanceTimersByTime(RENDER_TIMEOUT_MS);
      const [a, b] = await settled;
      expect(a.status === 'rejected' && a.reason instanceof ArtcraftUnavailable).toBe(true);
      expect(b.status === 'rejected' && b.reason instanceof ArtcraftUnavailable).toBe(true);
      expect(worker.terminate).toHaveBeenCalledOnce();
      expect(worker.postMessage).toHaveBeenCalledOnce(); // 한 번에 하나씩
      await expect(client.call({ cmd: 'text', page: 3 }, 1)).rejects.toBeInstanceOf(ArtcraftUnavailable);
    } finally {
      vi.useRealTimers();
    }
  });
});
