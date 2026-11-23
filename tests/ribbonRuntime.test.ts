import { describe, expect, it, vi } from 'vitest';
import { collapseOrder } from '../src/client/ribbonRuntime.js';

describe('collapseOrder — 오피스처럼 접는 순서', () => {
  it('collapse 작은 수 먼저, 같으면 오른쪽 묶음 먼저', () => {
    const groups = [
      { priority: 9, index: 0 },
      { priority: 2, index: 1 },
      { priority: 5, index: 2 },
      { priority: 2, index: 3 },
      { priority: 1, index: 4 },
    ];
    expect(collapseOrder(groups)).toEqual([4, 3, 1, 2, 0]);
  });
  it('빈 목록', () => expect(collapseOrder([])).toEqual([]));
});

// 브라우저 없이 fitPanel·묶음 판 열기만 돌려 보는 가짜 DOM(필요한 것만)
function fakeEl() {
  const classes = new Set<string>();
  const attrs = new Map<string, string>();
  return {
    classes,
    attrs,
    dataset: { collapse: '1' } as Record<string, string>,
    classList: {
      add: (c: string) => classes.add(c),
      remove: (...cs: string[]) => cs.forEach((c) => classes.delete(c)),
      contains: (c: string) => classes.has(c),
    },
    setAttribute: (k: string, v: string) => attrs.set(k, v),
  };
}

describe('fitPanel — 열린 묶음 판은 폭이 바뀔 때만 닫는다(M2)', () => {
  it('같은 폭으로 다시 맞추면 열린 판 유지, 폭이 바뀌면 닫음', async () => {
    const listeners: Record<string, (e: unknown) => void> = {};
    let focused: unknown = null;
    const inside = {};
    (globalThis as Record<string, unknown>).document = {
      addEventListener: (t: string, f: (e: unknown) => void) => (listeners[t] = f),
      get activeElement() {
        return focused;
      },
    };
    (globalThis as Record<string, unknown>).window = { innerWidth: 1000 };
    const { fitPanel, installGroupToggles } = await import('../src/client/ribbonRuntime.js');
    const toggle = Object.assign(fakeEl(), { focus: vi.fn(), getBoundingClientRect: () => ({ left: 10, bottom: 20 }) });
    const body = { id: '', offsetWidth: 100, style: {} as Record<string, string> };
    const group = Object.assign(fakeEl(), {
      querySelector: (sel: string) => (sel.includes('rb-group-toggle') ? toggle : sel.includes('ribbon-group-buttons') ? body : null),
      contains: (n: unknown) => n === inside,
    });
    Object.assign(toggle, { closest: () => group });
    Object.defineProperty(group, 'offsetWidth', { get: () => (group.classes.has('rb-collapsed') ? 60 : 400) });
    // 묶음을 접어야만 판이 넘치지 않는다
    const panel = {
      querySelectorAll: () => [group],
      get scrollWidth() {
        return group.classes.has('rb-collapsed') ? 100 : 500;
      },
      clientWidth: 200,
    } as unknown as HTMLElement;
    installGroupToggles({ addEventListener: (t: string, f: (e: unknown) => void) => (listeners[t] = f) } as unknown as HTMLElement);

    fitPanel(panel, 1000);
    expect(group.classes.has('rb-collapsed')).toBe(true);
    listeners.click({ target: { closest: () => toggle } });
    expect(group.classes.has('rb-open')).toBe(true);

    fitPanel(panel, 1000); // 상황별 탭·글꼴 받음으로 render 다시
    expect(group.classes.has('rb-open')).toBe(true);
    expect(toggle.attrs.get('aria-expanded')).toBe('true');

    // M1/M3: 속성·포커스 복귀·Esc
    expect(toggle.attrs.get('aria-haspopup')).toBe('true');
    expect(toggle.attrs.get('aria-controls')).toBe(body.id);
    listeners.keydown({ key: 'Escape', defaultPrevented: true, target: { closest: () => null } });
    expect(group.classes.has('rb-open')).toBe(true); // 이미 처리된 Esc는 무시
    focused = inside;
    listeners.keydown({ key: 'Escape', defaultPrevented: false, target: { closest: () => null } });
    expect(group.classes.has('rb-open')).toBe(false);
    expect(toggle.focus).toHaveBeenCalled();
    listeners.click({ target: { closest: () => toggle } });
    listeners.click({ target: { closest: (s: string) => (s === 'button' || s === '.ribbon-group-buttons' ? {} : null) } }); // 판 안 명령
    expect(group.classes.has('rb-open')).toBe(false);
    listeners.click({ target: { closest: () => toggle } });

    fitPanel(panel, 900); // 창 폭이 바뀜
    expect(group.classes.has('rb-open')).toBe(false);
    expect(toggle.attrs.get('aria-expanded')).toBe('false');
  });
});

describe('fitPanel — 오피스 순서: 크게 → 작은 단추 3줄 → 접기, 남으면 다시 펴기', () => {
  // 묶음 폭은 상태(크게·작게·접힘)에 따라 정해 두고, 판 scrollWidth는 그 합
  function setup(clientWidth: number) {
    (globalThis as Record<string, unknown>).document = { addEventListener: () => {} };
    const mk = (priority: number, full: number, compact: number, folded: number) => {
      const g = fakeEl();
      g.dataset.collapse = String(priority);
      // Object.assign은 getter를 값으로 굳히므로 defineProperty로
      return Object.defineProperty(g, 'offsetWidth', { get: () => (g.classes.has('rb-collapsed') ? folded : g.classes.has('rb-compact') ? compact : full) }) as typeof g & { offsetWidth: number };
    };
    const groups = [mk(9, 100, 100, 50), mk(2, 120, 80, 40), mk(1, 60, 60, 30)];
    const panel = {
      querySelectorAll: () => groups,
      get scrollWidth() {
        return groups.reduce((n, g) => n + g.offsetWidth, 0);
      },
      clientWidth,
    } as unknown as HTMLElement;
    const state = () => groups.map((g) => (g.classes.has('rb-collapsed') ? 'folded' : g.classes.has('rb-compact') ? 'compact' : 'full'));
    return { panel, state };
  }
  it('작게 줄여 들어가면 접지 않는다(줄어들지 않는 묶음은 크게 둔다)', async () => {
    const { fitPanel } = await import('../src/client/ribbonRuntime.js');
    const { panel, state } = setup(250);
    fitPanel(panel, 1400);
    expect(state()).toEqual(['full', 'compact', 'full']);
  });
  it('접은 뒤 먼저 접은 묶음이 다시 들어가면 편다', async () => {
    const { fitPanel } = await import('../src/client/ribbonRuntime.js');
    const { panel, state } = setup(200);
    fitPanel(panel, 1300);
    expect(state()).toEqual(['full', 'folded', 'full']);
  });
  it('넉넉하면 모두 크게', async () => {
    const { fitPanel } = await import('../src/client/ribbonRuntime.js');
    const { panel, state } = setup(400);
    fitPanel(panel, 1920);
    expect(state()).toEqual(['full', 'full', 'full']);
  });
});
