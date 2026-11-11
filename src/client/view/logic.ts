/**
 * 통합 보기 껍데기의 순수 계산 — DOM 없이 시험한다(tests/viewShellLogic.test.ts).
 * 값은 지금 보기 화면(public/js/viewer.js)과 같게 둔다(확대 단계·PNG 배율·쪽 맞춤·측정).
 */
import type { FindMatch } from '../pdf/findText.js';
import type { NormRect, PageSize, ViewPermissions } from './types.js';

/** 확대 단계(%) — viewer.js와 같다. 100% = 쪽 칸 폭 전체 */
export const ZOOM_STEPS = [50, 67, 80, 100, 125, 150, 200, 250, 300, 400, 500] as const;
export const DEFAULT_ZOOM_INDEX = 3;
/** 캔버스 한 장 최대 화소(pageView.ts와 같은 값) — 넘으면 해상도를 낮춰 그리고 CSS로 늘린다 */
export const MAX_CANVAS_PIXELS = 16_777_216;
/** 찾기 결과 상한(서버 /api/view/search limit 최대와 같다) */
export const SEARCH_LIMIT = 500;

/** 권한 → 안쪽: 내려받기 권한이 있을 때만 PDF를 브라우저로 받는다(설계 V5·§4) */
export function chooseBodyKind(perms: Pick<ViewPermissions, 'canDownload'>): 'pages-pdfjs' | 'pages-png' {
  return perms.canDownload ? 'pages-pdfjs' : 'pages-png';
}

export function clampZoomIndex(index: number): number {
  return Math.max(0, Math.min(ZOOM_STEPS.length - 1, index));
}

/** 가장 가까운 확대 단계(같으면 작은 쪽) */
export function snapZoomIndex(percent: number): number {
  let best = 0;
  ZOOM_STEPS.forEach((step, i) => {
    if (Math.abs(step - percent) < Math.abs(ZOOM_STEPS[best] - percent)) best = i;
  });
  return best;
}

/** 서버 PNG 배율(viewer.js desiredScale 그대로) */
export function pngScale(percent: number, dpr: number): number {
  const ratio = Math.min(dpr || 1, 2);
  return Math.min(16, Math.max(2, Math.ceil((percent / 100) * 2 * ratio)));
}

/** 다시 그릴지 — 처음이거나, 더 선명해야 하거나, 지금 그림이 필요의 두 배를 넘어 메모리를 낭비할 때 */
export function needsRedraw(drawnKey: number, wantKey: number): boolean {
  return drawnKey === 0 || wantKey > drawnKey || wantKey * 2 < drawnKey;
}

/** 글자 층을 다시 그릴지 — 그림은 두 배까지 그대로 쓰지만 글자 층은 쪽 폭에 맞춰야 선택이 그림과 겹친다(1px 넘게 다르면) */
export function textLayerStale(drawnWidth: number | undefined, cssWidth: number): boolean {
  return drawnWidth === undefined || Math.abs(drawnWidth - cssWidth) >= 1;
}

/** 늦게 끝난 그림을 쓸지 — 그리는 사이 창(보이는 쪽 + 지금 쪽 ±2) 밖이 됐으면 버린다 */
export function keepFinishedDraw(page: number, current: number, count: number, radius: number, visible: Iterable<number>): boolean {
  return pagesToKeep(current, count, radius, visible).has(page);
}

/** 캔버스를 남겨 둘 쪽 — 보이는 쪽 + 지금 쪽 앞뒤 radius(설계 §7 "보이는 쪽 ±2") */
export function pagesToKeep(current: number, count: number, radius: number, visible: Iterable<number>): Set<number> {
  const keep = new Set<number>(visible);
  for (let page = Math.max(1, current - radius); page <= Math.min(count, current + radius); page += 1) keep.add(page);
  return keep;
}

/** "쪽 맞춤" 확대(%) — viewer.js pageFitPercent 그대로(좌우·위아래 여백 32px). 100%에서 쪽이 남은 높이를 넘으면 줄인다 */
export function fitPagePercent(containerWidth: number, availableHeight: number, size: PageSize): number {
  const heightAt100 = (containerWidth - 32) * (size.h / size.w);
  const room = availableHeight - 32;
  return heightAt100 <= room ? 100 : (room / heightAt100) * 100;
}

/** 캔버스 화소 배율 — 화면 배율로 그리되 maxPixels를 넘지 않게 */
export function canvasPixelRatio(cssW: number, cssH: number, dpr: number, maxPixels = MAX_CANVAS_PIXELS): number {
  const ratio = dpr || 1;
  return cssW * cssH * ratio * ratio > maxPixels ? Math.sqrt(maxPixels / (cssW * cssH)) : ratio;
}

/** 편집기 ↔ 보기 화면을 오갈 때 들고 가는 임베드 옵션(editor*.html·viewer.html의 home·back) */
const CARRIED_PARAMS = ['home', 'back'] as const;

/** target(토큰 포함 상대 주소)에 지금 주소의 home·back과 extra를 붙인다. from 등 다른 값은 옮기지 않는다 */
export function carryViewParams(target: string, currentSearch: string, extra: Readonly<Record<string, string>> = {}): string {
  const url = new URL(target, 'http://local.invalid');
  const current = new URLSearchParams(currentSearch);
  for (const key of CARRIED_PARAMS) {
    const value = current.get(key);
    if (value !== null) url.searchParams.set(key, value);
  }
  for (const [key, value] of Object.entries(extra)) url.searchParams.set(key, value);
  return `${url.pathname}${url.search}`;
}

/** 편집 단추가 갈 주소 — 편집기 "보기로"가 보이게 from=view를 붙인다(설계 §6) */
export function editHref(editorUrl: string, currentSearch: string): string {
  return carryViewParams(editorUrl, currentSearch, { from: 'view' });
}

/** 편집기의 "보기로" 단추를 보일지 — 껍데기에서 편집으로 들어온 경우만(설계 V7) */
export function offersBackToView(currentSearch: string, viewerUrl: string | undefined): boolean {
  return Boolean(viewerUrl) && new URLSearchParams(currentSearch).get('from') === 'view';
}

export function pathOf(url: string): string {
  return new URL(url, 'http://local.invalid').pathname;
}

/** 순환 번호(찾기 이전·다음) — n이 0이면 -1 */
export function wrapIndex(index: number, n: number): number {
  return n === 0 ? -1 : ((index % n) + n) % n;
}

/** 찾기 칸 표시 — viewer.js gotoHit와 같은 "3 / 500+" */
export function hitCountLabel(current: number, total: number): string {
  if (total === 0) return '결과 없음';
  return `${current + 1} / ${total >= SEARCH_LIMIT ? `${SEARCH_LIMIT}+` : String(total)}`;
}

/** pdf.js 글 조각의 모양(TextItem에서 쓰는 것만) */
export interface TextItemGeom {
  readonly str: string;
  readonly transform: readonly number[];
  readonly width: number;
  readonly height: number;
}

/** 배율 1·회전 0 viewport의 변환·크기 */
export interface ViewportGeom {
  readonly transform: readonly number[];
  readonly width: number;
  readonly height: number;
}

const applyTransform = (t: readonly number[], x: number, y: number): [number, number] => [t[0] * x + t[2] * y + t[4], t[1] * x + t[3] * y + t[5]];

/**
 * 찾은 곳(findText.ts FindMatch — 조각 번호·조각 안 글자 위치) → 쪽 대비 사각형. 글자 층(DOM) 없이 계산하므로 글자 선택이
 * 꺼진(canCopy 없음) pdf.js 모드에서도 찾기 강조가 된다. 조각 안 글자 폭은 고르게 나눈다.
 * shortcut: 가로 글자만 정확하다(돌아간·세로 글 조각은 기준점에서 가로로 근사) — 도면 PDF에서 어긋나 보이면 transform의 a·b로 방향을 쓴다.
 */
export function matchRects(m: FindMatch, items: readonly TextItemGeom[], vp: ViewportGeom): NormRect[] {
  const out: NormRect[] = [];
  for (let i = m.itemStart; i <= m.itemEnd; i += 1) {
    const item = items[i];
    if (!item || item.str.length === 0) continue;
    const len = item.str.length;
    const start = i === m.itemStart ? Math.min(m.offsetStart, len) : 0;
    const end = i === m.itemEnd ? Math.min(m.offsetEnd, len) : len;
    if (end <= start) continue;
    const [, , c, d, tx, ty] = item.transform;
    const height = item.height || Math.hypot(c, d);
    const x0 = tx + (item.width * start) / len;
    const x1 = tx + (item.width * end) / len;
    const [ax, ay] = applyTransform(vp.transform, x0, ty);
    const [bx, by] = applyTransform(vp.transform, x1, ty + height);
    const left = Math.min(ax, bx);
    const top = Math.min(ay, by);
    out.push({ x: left / vp.width, y: top / vp.height, w: (Math.max(ax, bx) - left) / vp.width, h: (Math.max(ay, by) - top) / vp.height });
  }
  return out;
}

/** 측정: 두 점(쪽 비율) 사이 용지 거리(mm) — viewer.js addMeasurement 그대로(1pt = 1/72in) */
export function paperMm(a: { readonly fx: number; readonly fy: number }, b: { readonly fx: number; readonly fy: number }, size: PageSize): number {
  return (Math.hypot((b.fx - a.fx) * size.w, (b.fy - a.fy) * size.h) * 25.4) / 72;
}

/** 측정 이름표 — viewer.js measurementLabelText 그대로 */
export function measureLabel(mm: number, factor: number, unit: string): string {
  const value = mm * factor;
  return `${value >= 100 ? value.toFixed(0) : value.toFixed(1)} ${unit}`;
}

/** 640~1023px에서 "도구▾" 메뉴로 들어가는 것(설계 §2 도구▾). 1024px 이상은 막대에 그대로 둔다(verify:viewer-quicktext가 #text-tool-btn을 누른다) */
export const TOOL_MENU_IDS = ['markup-add-btn', 'text-tool-btn', 'markup-include-wrap', 'measure-btn', 'invert-btn', 'fullscreen-btn'] as const;
/** 640px 미만에서 더 들어가는 것 — 막대에는 확대·쪽·찾기·도구▾만 남는다(설계 §2) */
export const NARROW_MENU_IDS = ['thumbs-btn', 'rotate-btn', 'print-btn', 'download-btn', 'download-as-btn'] as const;

export function menuIdsFor(width: number): string[] {
  if (width >= 1024) return [];
  return width >= 640 ? [...TOOL_MENU_IDS] : [...TOOL_MENU_IDS, ...NARROW_MENU_IDS];
}

/** 썸네일 칸 — 640px 미만은 닫힌 채 시작(설계 §2) */
export function thumbsOpenAtStart(width: number): boolean {
  return width >= 640;
}

export type ControlName = 'print' | 'download' | 'markup' | 'measure';

/** 흐린 단추의 이유(툴팁) — null이면 쓸 수 있다(설계 §2 "흐리게 하고 이유를 툴팁으로") */
export function disabledReason(control: ControlName, perms: ViewPermissions, ext: string): string | null {
  switch (control) {
    case 'print':
      return perms.canPrint ? null : '인쇄 권한이 없습니다';
    case 'download':
      return perms.canDownload ? null : '내려받기 권한이 없습니다';
    case 'markup':
      return perms.canEdit ? null : '표시를 넣을 권한이 없습니다';
    case 'measure':
      return ext === 'pdf' ? null : 'PDF 문서에서만 쓸 수 있습니다';
  }
}

/** Esc 처리 순서 — priority가 작은 것부터 불러 true를 돌려준 곳에서 멈춘다 */
export function createEscapeChain(): { add(priority: number, handler: () => boolean): void; run(): boolean } {
  const list: Array<{ readonly priority: number; readonly handler: () => boolean }> = [];
  return {
    add(priority, handler) {
      list.push({ priority, handler });
      list.sort((a, b) => a.priority - b.priority);
    },
    run() {
      return list.some((entry) => entry.handler());
    },
  };
}
