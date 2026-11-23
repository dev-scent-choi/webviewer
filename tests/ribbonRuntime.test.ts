import { describe, expect, it } from 'vitest';
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
      remove: (c: string) => classes.delete(c),
      contains: (c: string) => classes.has(c),
    },
    setAttribute: (k: string, v: string) => attrs.set(k, v),
  };
}

describe('fitPanel — 열린 묶음 판은 폭이 바뀔 때만 닫는다(M2)', () => {
  it('같은 폭으로 다시 맞추면 열린 판 유지, 폭이 바뀌면 닫음', async () => {
    const listeners: Record<string, (e: unknown) => void> = {};
    (globalThis as Record<string, unknown>).document = { addEventListener: () => {} };
    (globalThis as Record<string, unknown>).window = { innerWidth: 1000 };
    const { fitPanel, installGroupToggles } = await import('../src/client/ribbonRuntime.js');
    const toggle = fakeEl();
    const group = Object.assign(fakeEl(), {
      querySelector: (sel: string) => (sel.includes('rb-group-toggle') ? toggle : null),
    });
    Object.assign(toggle, { closest: () => group });
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

    fitPanel(panel, 900); // 창 폭이 바뀜
    expect(group.classes.has('rb-open')).toBe(false);
    expect(toggle.attrs.get('aria-expanded')).toBe('false');
  });
});
