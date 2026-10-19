import { describe, expect, test } from 'vitest';
import { isOpenable, trackInitialConversion, waitForInitialConversion } from '../src/conversionGate.js';
import type { DocumentMeta } from '../src/types.js';

const meta = (over: Partial<DocumentMeta>): DocumentMeta => ({ id: 'd', name: 'x', ext: 'docx', status: 'processing', pageCount: 0, createdAt: '', ...over });

describe('변환 전 편집기 열기(editReady)', () => {
  test('변환이 끝났거나, 변환 중이어도 editReady면 연다 — 실패·다른 형식의 변환 중은 아님', () => {
    expect(isOpenable(meta({ status: 'ready' }))).toBe(true);
    expect(isOpenable(meta({ status: 'processing', editReady: true }))).toBe(true);
    expect(isOpenable(meta({ status: 'processing' }))).toBe(false);
    expect(isOpenable(meta({ status: 'failed', editReady: true }))).toBe(false);
    expect(isOpenable(null)).toBe(false);
  });

  test('자동저장은 첫 변환이 끝날 때까지(실패해도) 기다린다', async () => {
    const order: string[] = [];
    let finish!: () => void;
    trackInitialConversion('doc-1', new Promise<void>((resolve) => (finish = resolve)).then(() => order.push('첫 변환 끝')));
    const waiting = waitForInitialConversion('doc-1').then(() => order.push('자동저장 변환 시작'));
    await new Promise((r) => setTimeout(r, 20));
    expect(order).toEqual([]);
    finish();
    await waiting;
    expect(order).toEqual(['첫 변환 끝', '자동저장 변환 시작']);
    trackInitialConversion('doc-2', Promise.reject(new Error('변환 실패')));
    await expect(waitForInitialConversion('doc-2')).resolves.toBeUndefined();
    await expect(waitForInitialConversion('없는 문서')).resolves.toBeUndefined();
  });
});
