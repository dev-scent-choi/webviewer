import { describe, expect, test } from 'vitest';
import { findInPage } from '../src/client/pdf/findText.js';
import {
  canvasPixelRatio,
  carryViewParams,
  chooseBodyKind,
  createEscapeChain,
  disabledReason,
  editHref,
  fitPagePercent,
  hitCountLabel,
  keepFinishedDraw,
  textLayerStale,
  matchRects,
  measureLabel,
  menuIdsFor,
  needsRedraw,
  offersBackToView,
  pagesToKeep,
  paperMm,
  pathOf,
  pngScale,
  snapZoomIndex,
  thumbsOpenAtStart,
  wrapIndex,
  ZOOM_STEPS,
} from '../src/client/view/logic.js';

const perms = (p: Partial<{ canPrint: boolean; canDownload: boolean; canCopy: boolean; canEdit: boolean }> = {}) => ({
  canPrint: false,
  canDownload: false,
  canCopy: false,
  canEdit: false,
  ...p,
});

describe('권한 → 안쪽', () => {
  test('내려받기 권한이 있을 때만 pdf.js, 아니면 서버 PNG', () => {
    expect(chooseBodyKind(perms({ canDownload: true }))).toBe('pages-pdfjs');
    expect(chooseBodyKind(perms({ canEdit: true, canCopy: true, canPrint: true }))).toBe('pages-png');
  });
});

describe('확대·그리기', () => {
  test('가장 가까운 확대 단계', () => {
    expect(ZOOM_STEPS[snapZoomIndex(100)]).toBe(100);
    expect(ZOOM_STEPS[snapZoomIndex(72)]).toBe(67);
    expect(ZOOM_STEPS[snapZoomIndex(10)]).toBe(50);
    expect(ZOOM_STEPS[snapZoomIndex(9999)]).toBe(500);
  });
  test('서버 PNG 배율은 viewer.js와 같다(2~16, dpr은 2까지)', () => {
    expect(pngScale(100, 1)).toBe(2);
    expect(pngScale(100, 1.5)).toBe(3);
    expect(pngScale(100, 3)).toBe(4);
    expect(pngScale(50, 1)).toBe(2);
    expect(pngScale(500, 2)).toBe(16);
  });
  test('다시 그리기: 처음·더 선명하게·두 배 넘게 큰 그림', () => {
    expect(needsRedraw(0, 4)).toBe(true);
    expect(needsRedraw(4, 4)).toBe(false);
    expect(needsRedraw(4, 6)).toBe(true);
    expect(needsRedraw(4, 3)).toBe(false);
    expect(needsRedraw(1000, 400)).toBe(true);
  });
  test('남길 쪽 = 보이는 쪽 + 지금 쪽 ±2(문서 범위 안)', () => {
    expect([...pagesToKeep(1, 10, 2, [])].sort((a, b) => a - b)).toEqual([1, 2, 3]);
    expect([...pagesToKeep(5, 10, 2, [9])].sort((a, b) => a - b)).toEqual([3, 4, 5, 6, 7, 9]);
    expect([...pagesToKeep(10, 10, 2, [])].sort((a, b) => a - b)).toEqual([8, 9, 10]);
  });
  test('쪽 맞춤: 높이가 넉넉하면 100, 아니면 줄인다', () => {
    expect(fitPagePercent(1032, 2000, { w: 595, h: 842 })).toBe(100);
    // 폭 1000(-32=968) → 100%에서 높이 968*842/595 ≈ 1369.9, 남은 높이 800-32=768 → 약 56%
    expect(fitPagePercent(1000, 800, { w: 595, h: 842 })).toBeCloseTo(56.06, 1);
  });
  test('캔버스 화소 상한', () => {
    expect(canvasPixelRatio(1000, 1000, 2)).toBe(2);
    const r = canvasPixelRatio(4000, 4000, 2);
    expect(4000 * 4000 * r * r).toBeLessThanOrEqual(16_777_216 + 1);
  });
});

describe('편집 ↔ 보기 주소', () => {
  test('home·back만 들고 가고 extra를 붙인다', () => {
    expect(carryViewParams('/editor.html?token=a.b', '?token=a.b&home=hidden&back=parent&from=x', { from: 'view' })).toBe(
      '/editor.html?token=a.b&home=hidden&back=parent&from=view',
    );
    expect(carryViewParams('/view.html?token=t', '?token=t&from=view')).toBe('/view.html?token=t');
  });
  test('편집 단추 주소에는 from=view', () => {
    expect(editHref('/editor-slides.html?token=t', '?token=t')).toBe('/editor-slides.html?token=t&from=view');
  });
  test('"보기로"는 from=view이고 viewerUrl이 있을 때만', () => {
    expect(offersBackToView('?token=t&from=view', '/view.html?token=t')).toBe(true);
    expect(offersBackToView('?token=t', '/view.html?token=t')).toBe(false);
    expect(offersBackToView('?token=t&from=view', undefined)).toBe(false);
  });
  test('경로만', () => {
    expect(pathOf('/view.html?token=t')).toBe('/view.html');
  });
});

describe('찾기', () => {
  test('순환 번호와 표시', () => {
    expect(wrapIndex(-1, 3)).toBe(2);
    expect(wrapIndex(3, 3)).toBe(0);
    expect(wrapIndex(0, 0)).toBe(-1);
    expect(hitCountLabel(0, 3)).toBe('1 / 3');
    expect(hitCountLabel(4, 500)).toBe('5 / 500+');
    expect(hitCountLabel(-1, 0)).toBe('결과 없음');
  });
  test('pdf.js 글 조각 → 쪽 대비 사각형(배율 1·회전 0, A4)', () => {
    const items = [{ str: 'hello world', transform: [10, 0, 0, 10, 100, 700], width: 55, height: 10 }];
    const [m] = findInPage(1, items, 'world');
    const vp = { transform: [1, 0, 0, -1, 0, 842], width: 595, height: 842 };
    const [r] = matchRects(m, items, vp);
    expect(r.x).toBeCloseTo((100 + (55 * 6) / 11) / 595, 6);
    expect(r.w).toBeCloseTo((55 * 5) / 11 / 595, 6);
    expect(r.y).toBeCloseTo((842 - 710) / 842, 6);
    expect(r.h).toBeCloseTo(10 / 842, 6);
  });
  test('조각 두 개에 걸친 일치는 사각형 두 개', () => {
    const items = [
      { str: 'abc', transform: [10, 0, 0, 10, 0, 100], width: 30, height: 10 },
      { str: 'def', transform: [10, 0, 0, 10, 30, 100], width: 30, height: 10 },
    ];
    const [m] = findInPage(1, items, 'cd');
    expect(matchRects(m, items, { transform: [1, 0, 0, -1, 0, 200], width: 100, height: 200 })).toHaveLength(2);
  });
});

describe('측정', () => {
  test('A4 폭 전체 = 210mm 남짓, 이름표 자리수', () => {
    expect(paperMm({ fx: 0, fy: 0 }, { fx: 1, fy: 0 }, { w: 595, h: 842 })).toBeCloseTo(209.9, 1);
    expect(measureLabel(209.9, 1, 'mm(용지)')).toBe('210 mm(용지)');
    expect(measureLabel(12.34, 1, 'mm')).toBe('12.3 mm');
  });
});

describe('막대 배치·단추 상태', () => {
  test('1024px 이상은 메뉴 없음, 640~1023은 도구만, 그 아래는 더', () => {
    expect(menuIdsFor(1400)).toEqual([]);
    expect(menuIdsFor(800)).toContain('text-tool-btn');
    expect(menuIdsFor(800)).not.toContain('print-btn');
    expect(menuIdsFor(390)).toContain('print-btn');
    expect(menuIdsFor(390)).not.toContain('zoom-in');
  });
  test('썸네일은 휴대폰 폭에서 닫힌 채', () => {
    expect(thumbsOpenAtStart(390)).toBe(false);
    expect(thumbsOpenAtStart(1280)).toBe(true);
  });
  test('흐린 단추의 이유', () => {
    expect(disabledReason('print', perms(), 'docx')).toBe('인쇄 권한이 없습니다');
    expect(disabledReason('print', perms({ canPrint: true }), 'docx')).toBeNull();
    expect(disabledReason('download', perms(), 'pdf')).toBe('내려받기 권한이 없습니다');
    expect(disabledReason('markup', perms(), 'pdf')).toBe('표시를 넣을 권한이 없습니다');
    expect(disabledReason('measure', perms(), 'docx')).toBe('PDF 문서에서만 쓸 수 있습니다');
    expect(disabledReason('measure', perms(), 'pdf')).toBeNull();
  });
  test('Esc는 우선순위 순서로 처음 받은 곳에서 멈춘다', () => {
    const calls: string[] = [];
    const chain = createEscapeChain();
    chain.add(90, () => (calls.push('search'), true));
    chain.add(10, () => (calls.push('measure'), false));
    chain.add(40, () => (calls.push('markup'), true));
    expect(chain.run()).toBe(true);
    expect(calls).toEqual(['measure', 'markup']);
  });
});

describe('늦게 끝난 그림·글자 층', () => {
  test('textLayerStale: 처음이거나 폭이 1px 이상 다르면', () => {
    expect(textLayerStale(undefined, 800)).toBe(true);
    expect(textLayerStale(800, 800.4)).toBe(false);
    expect(textLayerStale(800, 400)).toBe(true);
  });
  test('keepFinishedDraw: 창(보이는 쪽 + 지금 쪽 ±2) 밖이면 버린다', () => {
    expect(keepFinishedDraw(3, 1, 20, 2, [1])).toBe(true);
    expect(keepFinishedDraw(10, 1, 20, 2, [1])).toBe(false);
    expect(keepFinishedDraw(10, 1, 20, 2, [10])).toBe(true);
  });
});
