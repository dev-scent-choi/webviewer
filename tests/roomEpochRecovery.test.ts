import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { recoveryText } from '../src/client/roomEpoch.js';

// R3-I1: 다시 열린 방 안내에서 건네는 '내 입력 내용' — 고친 문단·셀의 현재 글자
describe('recoveryText', () => {
  const collect = (doc: Y.Doc) => {
    const edits = new Map<Y.AbstractType<any>, Set<string | null>>(); // eslint-disable-line @typescript-eslint/no-explicit-any
    doc.on('afterTransaction', (tr: Y.Transaction) => tr.changed.forEach((keys, type) => {
      const s = edits.get(type) ?? new Set();
      keys.forEach((k) => s.add(k));
      edits.set(type, s);
    }));
    return edits;
  };

  test('문단 글자와 맵 칸 값을 한 줄씩 모은다', () => {
    const doc = new Y.Doc();
    const frag = doc.getXmlFragment('prosemirror');
    const p1 = new Y.XmlElement('paragraph');
    const p2 = new Y.XmlElement('paragraph');
    p1.insert(0, [new Y.XmlText('그대로')]);
    p2.insert(0, [new Y.XmlText('')]);
    frag.insert(0, [p1, p2]);
    const edits = collect(doc);
    (p2.get(0) as Y.XmlText).insert(0, '새로 쓴 글');
    doc.getMap<string>('cells').set('Sheet1!B3', '42');
    expect(recoveryText(edits).split('\n')).toEqual(['새로 쓴 글', 'Sheet1!B3: 42']);
  });

  test('고친 것이 없으면 빈 글자', () => {
    expect(recoveryText(new Map())).toBe('');
  });
});
