import { describe, expect, test } from 'vitest';
import * as Y from 'yjs';
import { Awareness, applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import { sanitizeAwarenessUpdate } from '../src/collab/awarenessGuard.js';

function updateFrom(clientID: number, state: Record<string, unknown> | null): Uint8Array {
  const doc = new Y.Doc();
  doc.clientID = clientID;
  const a = new Awareness(doc);
  a.setLocalState(state);
  return encodeAwarenessUpdate(a, [clientID]);
}

function applied(update: Uint8Array): Map<number, Record<string, unknown>> {
  const a = new Awareness(new Y.Doc());
  applyAwarenessUpdate(a, update, null);
  return a.getStates() as Map<number, Record<string, unknown>>;
}

describe('sanitizeAwarenessUpdate', () => {
  test('이름은 토큰의 서명된 이름으로, 색은 #rrggbb 만', () => {
    const evil = updateFrom(7, { user: { name: '대표이사', color: 'red;position:fixed;inset:0' }, cursor: { anchor: 1 } });
    const state = applied(sanitizeAwarenessUpdate(evil, new Set(), '홍길동')).get(7)!;
    expect(state.user).toEqual({ name: '홍길동' });
    expect(state.cursor).toEqual({ anchor: 1 });
    const ok = applied(sanitizeAwarenessUpdate(updateFrom(7, { user: { name: 'x', color: '#1971c2' } }), new Set(), '홍길동')).get(7)!;
    expect(ok.user).toEqual({ name: '홍길동', color: '#1971c2' });
  });

  test('다른 연결의 clientID 항목(지우기·덮어쓰기)은 버린다', () => {
    const forged = updateFrom(42, { user: { name: '가짜' } });
    expect(applied(sanitizeAwarenessUpdate(forged, new Set([42]), 'a')).has(42)).toBe(false);
    expect(applied(sanitizeAwarenessUpdate(forged, new Set([7]), 'a')).has(42)).toBe(true);
  });
});
