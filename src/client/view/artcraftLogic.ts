/**
 * 아트크래프트 원천의 순수 계산과 Worker 주고받기 모양 — DOM 없이 시험한다(tests/artcraftLogic.test.ts).
 * 엔진 글자 JSON·C ABI는 tools/artcraft-shell/abi.rs, 설계 docs/superpowers/specs/2026-11-21-artcraft-browser-engine-design.md §4.
 */
import { canvasPixelRatio } from './logic.js';
import type { NormRect, PageSize, SearchHit } from './types.js';

/** 엔진 sp_text의 글자 조각 — pt, y 아래로, y는 기준선, s는 글자 크기 */
export interface TextRun {
  readonly t: string;
  readonly x0: number;
  readonly x1: number;
  readonly y: number;
  readonly s: number;
}
export interface PageText {
  readonly w: number;
  readonly h: number;
  readonly runs: readonly TextRun[];
}

export type WorkerCall =
  | { readonly cmd: 'init'; readonly wasm: string; readonly fonts: readonly string[]; readonly original: string; readonly name: string }
  | { readonly cmd: 'render'; readonly page: number; readonly scale: number }
  | { readonly cmd: 'text'; readonly page: number };
export type WorkerRequest = WorkerCall & { readonly id: number };
/** initMs = 받기·컴파일·글꼴 넣기, openMs = sp_open·쪽 크기(Worker 안 performance.now — e2e CPU×4 환산에 쓴다) */
export interface InitReply { readonly id: number; readonly count: number; readonly sizes: PageSize[]; readonly initMs?: number; readonly openMs?: number }
export interface RenderReply { readonly id: number; readonly w: number; readonly h: number; readonly rgba: ArrayBuffer; readonly ms?: number }
export interface TextReply { readonly id: number; readonly text: PageText }
/** fatal = wasm trap(RuntimeError) — 엔진 상태가 오염됐을 수 있어 Worker를 끈다 */
export interface ErrorReply { readonly id: number; readonly error: string; readonly fatal?: boolean }
export type WorkerReply = InitReply | RenderReply | TextReply | ErrorReply;

/** 글자 상자 위 = 기준선 − 0.8 × 글자 크기, 높이 = 글자 크기(엔진은 기준선과 크기만 준다) */
const ASCENT = 0.8;
/** 요청 화소 폭이 쪽 폭 pt의 이만큼을 넘어야 두 단계로 그린다(그 아래면 1배와 차이가 작다) */
const TWO_STAGE_FACTOR = 1.5;

/** 그릴 화소 폭 — CSS 폭 × dpr, 캔버스 한 장 상한(MAX_CANVAS_PIXELS)을 넘지 않게 */
export function targetPixelWidth(cssWidth: number, dpr: number, size: PageSize): number {
  const cssHeight = cssWidth * (size.h / size.w);
  return Math.max(1, Math.ceil(cssWidth * canvasPixelRatio(cssWidth, cssHeight, dpr)));
}

/** 1배(scale = 1, 쪽 폭 pt = 화소) 그림의 화소 폭 */
export function quickPixelWidth(size: PageSize): number {
  return Math.max(1, Math.round(size.w));
}

/**
 * 이 쪽을 어떻게 그릴지 — cachedPx는 이 쪽에 대해 보관 중인 그림의 화소 폭들.
 * 같은 폭이 있으면 그대로, 선명한(1.5배 넘는) 그림이 아직 없으면 1배 먼저 보이고 이어서 요청 배율(설계 §4.1), 그 밖은 바로.
 */
export function drawPlan(cachedPx: readonly number[], wantPx: number, pageWidthPt: number): 'cached' | 'direct' | 'two-stage' {
  if (cachedPx.includes(wantPx)) return 'cached';
  const sharp = TWO_STAGE_FACTOR * pageWidthPt;
  if (wantPx > sharp && !cachedPx.some((px) => px > sharp)) return 'two-stage';
  return 'direct';
}

/**
 * 한 쪽에서 찾기(대소문자 무시) — run 상자 안을 글자 수 비율로 잘라 쪽 비율 사각형을 만든다.
 * shortcut: matches that cross two runs are not found; join neighbouring runs if users report misses.
 */
export function runHits(page: number, text: PageText, q: string): SearchHit[] {
  const needle = q.trim().toLowerCase();
  if (!needle || text.w <= 0 || text.h <= 0) return [];
  const hits: SearchHit[] = [];
  for (const run of text.runs) {
    const hay = run.t.toLowerCase();
    // 소문자로 바꾸며 길이가 바뀌는 글자(예: İ)가 있으면 run 전체를 칠한다
    const exact = hay.length === run.t.length;
    for (let at = hay.indexOf(needle); at >= 0; at = hay.indexOf(needle, at + needle.length)) {
      const a = exact ? at / run.t.length : 0;
      const b = exact ? (at + needle.length) / run.t.length : 1;
      const x0 = run.x0 + (run.x1 - run.x0) * a;
      const x1 = run.x0 + (run.x1 - run.x0) * b;
      const rect: NormRect = { x: x0 / text.w, y: (run.y - run.s * ASCENT) / text.h, w: (x1 - x0) / text.w, h: run.s / text.h };
      hits.push({ page, rects: [rect] });
    }
  }
  return hits;
}

/** 글자 층 span 하나의 자리 — left·top은 쪽 비율 %, 글자 크기·폭은 지금 쪽 CSS 폭 기준 화소 */
export function runSpanGeom(run: TextRun, size: PageSize, cssWidth: number): { left: number; top: number; fontPx: number; widthPx: number } {
  const scale = cssWidth / size.w;
  return {
    left: (run.x0 / size.w) * 100,
    top: ((run.y - run.s * ASCENT) / size.h) * 100,
    fontPx: run.s * scale,
    widthPx: (run.x1 - run.x0) * scale,
  };
}

/** .doc(실험): 서버 변환 쪽 수(알 때만, >0)와 2쪽 넘게 다르면 의심 → 리브레오피스로(설계 §4.2) */
export function pageCountSuspicious(experimental: boolean, engineCount: number, serverCount: number): boolean {
  return experimental && serverCount > 0 && Math.abs(engineCount - serverCount) > 2;
}

/** 바이트 상한 LRU — 쪽 그림(ImageData)을 보관한다. 방금 넣은 것은 상한보다 커도 남긴다 */
/** 쪽마다 연달아 실패한 횟수 — 한 번은 다음 그리기에 다시, 같은 쪽 두 번째면 엔진을 버린다(설계 §4.2, 최종 검토 I2) */
export class PageFailures {
  private readonly counts = new Map<number, number>();
  /** 실패를 적고, 같은 쪽 연속 두 번째 이상이면 true */
  fail(page: number): boolean {
    const n = (this.counts.get(page) ?? 0) + 1;
    this.counts.set(page, n);
    return n >= 2;
  }
  ok(page: number): void {
    this.counts.delete(page);
  }
}

export class ByteLru<V> {
  private readonly map = new Map<string, V>();
  private total = 0;
  constructor(
    private readonly maxBytes: number,
    private readonly sizeOf: (value: V) => number,
  ) {}
  get bytes(): number {
    return this.total;
  }
  get(key: string): V | undefined {
    const value = this.map.get(key);
    if (value !== undefined) {
      this.map.delete(key);
      this.map.set(key, value);
    }
    return value;
  }
  set(key: string, value: V): void {
    const old = this.map.get(key);
    if (old !== undefined) {
      this.total -= this.sizeOf(old);
      this.map.delete(key);
    }
    this.map.set(key, value);
    this.total += this.sizeOf(value);
    for (const [k, v] of this.map) {
      if (this.total <= this.maxBytes || k === key) break;
      this.map.delete(k);
      this.total -= this.sizeOf(v);
    }
  }
  keys(): string[] {
    return [...this.map.keys()];
  }
  delete(key: string): void {
    const value = this.map.get(key);
    if (value === undefined) return;
    this.map.delete(key);
    this.total -= this.sizeOf(value);
  }
}

/** Worker 작업 줄 — 급한 것(보이는 쪽) 먼저, 같은 급끼리는 넣은 순서(PagesBody가 보이는 쪽부터 부른다) */
export class JobQueue<T> {
  private readonly urgent: T[] = [];
  private readonly later: T[] = [];
  get size(): number {
    return this.urgent.length + this.later.length;
  }
  push(job: T, urgent: boolean): void {
    (urgent ? this.urgent : this.later).push(job);
  }
  shift(): T | undefined {
    return this.urgent.shift() ?? this.later.shift();
  }
  remove(job: T): boolean {
    for (const list of [this.urgent, this.later]) {
      const at = list.indexOf(job);
      if (at >= 0) {
        list.splice(at, 1);
        return true;
      }
    }
    return false;
  }
}

/**
 * wasm-bindgen 없이 돌릴 때 엔진이 요구하는 바깥 함수(spike §7.6) — 시각·시간대만 진짜, 나머지는 0을 돌려주고 이름마다 한 번 경고.
 * 시험 중 실제로 불린 것은 없었다(불리면 엔진 바인딩 문제).
 */
export function stubImports(descriptors: readonly WebAssembly.ModuleImportDescriptor[], warn: (name: string) => void): WebAssembly.Imports {
  const imports: Record<string, Record<string, () => number>> = {};
  const warned = new Set<string>();
  for (const d of descriptors) {
    if (d.kind !== 'function') continue;
    const fns = (imports[d.module] ??= {});
    if (/(^|_)now(_|$)/.test(d.name)) fns[d.name] = () => Date.now();
    else if (/getTimezoneOffset/.test(d.name)) fns[d.name] = () => new Date().getTimezoneOffset();
    else
      fns[d.name] = () => {
        if (!warned.has(d.name)) {
          warned.add(d.name);
          warn(d.name);
        }
        return 0;
      };
  }
  return imports;
}

/** 엔진에 넘길 원본 상한 — wasm 메모리 상한이 512 MB(--max-memory)이고 엔진이 입력을 두 번 복사하므로 약 150 MB */
/** 글꼴 하나를 넣은 결과(sp_add_font = 들어간 글꼴 수) — 0개면 한글이 빠진 채 그려지므로 열기 실패(→ 서버 변환, 설계 §4.2) */
export function fontProblem(faces: number): string | null {
  return faces < 1 ? '글꼴을 넣지 못했습니다' : null;
}

export const MAX_ORIGINAL_BYTES = 150 * 1024 * 1024;
const ZIP = [0x50, 0x4b, 0x03, 0x04];
const OLE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

/**
 * 엔진 open 전 원본 검사 — 문제면 이유, 괜찮으면 null. word 엔진은 빈·깨진 입력도 빈 문서로 "열어" 버리므로 여기서 거른다.
 * docx·pptx는 zip, doc는 OLE 머리(그 밖의 확장자는 서버가 보내지 않는다).
 */
export function originalProblem(ext: string, bytes: Uint8Array): string | null {
  if (bytes.length === 0) return '원본이 비어 있습니다';
  if (bytes.length > MAX_ORIGINAL_BYTES) return '원본이 브라우저 엔진에 너무 큽니다';
  const magic = ext === 'doc' ? OLE : ZIP;
  const head = bytes.subarray(0, magic.length);
  return magic.every((b, i) => head[i] === b) ? null : `원본이 ${ext} 형식이 아닙니다`;
}
