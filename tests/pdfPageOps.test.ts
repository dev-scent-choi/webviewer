import * as Y from 'yjs';
import { describe, expect, test } from 'vitest';
import { PageOps } from '../src/client/pdf/pageOps.js';
import { LOCAL_ORIGIN } from '../src/client/pdf/annotLayer.js';
import { identityPages, type PdfAnnot, type PdfPageEntry } from '../src/pdf/annotTypes.js';

function setup(n = 5) {
  const ydoc = new Y.Doc();
  const arr = ydoc.getArray<PdfPageEntry>('pdfPages');
  arr.push(identityPages(n));
  const ops = new PageOps(ydoc, arr, ydoc.getMap<PdfAnnot>('pdfAnnots'));
  const keys = () => ops.list().map((e) => e.key.slice(2));
  return { ydoc, arr, ops, keys };
}

describe('쪽 목록 바꾸기(정리 화면)', () => {
  test('옮기기 — 앞으로·뒤로·여러 쪽·끝으로', () => {
    let s = setup();
    s.ops.move(['o:3'], 1);
    expect(s.keys()).toEqual(['0', '3', '1', '2', '4']);
    s = setup();
    s.ops.move(['o:0', 'o:2'], 5);
    expect(s.keys()).toEqual(['1', '3', '4', '0', '2']);
    s = setup();
    s.ops.move(['o:1', 'o:3'], 0);
    expect(s.keys()).toEqual(['1', '3', '0', '2', '4']);
    s = setup();
    s.ops.move(['o:1'], 3);
    expect(s.keys()).toEqual(['0', '2', '1', '3', '4']);
    s = setup();
    s.ops.move(['o:1'], 2); // 제자리
    expect(s.keys()).toEqual(['0', '1', '2', '3', '4']);
  });

  test('회전·지우기(전부는 안 됨)·빈 쪽·넣기', () => {
    const s = setup(3);
    s.ops.rotate(['o:1'], -90);
    expect(s.ops.list()[1]).toEqual({ key: 'o:1', rotate: 270 });
    s.ops.rotate(['o:1'], 90);
    expect(s.ops.list()[1].rotate).toBe(0);
    expect(s.ops.remove(['o:0', 'o:1', 'o:2'])).toBe(false);
    expect(s.ops.remove(['o:0'])).toBe(true);
    expect(s.keys()).toEqual(['1', '2']);
    const blank = s.ops.insertBlank(1, { w: 300, h: 400 });
    expect(s.ops.list()[1]).toEqual({ key: blank, rotate: 0, blank: { w: 300, h: 400 } });
    s.ops.insert(3, [{ key: 'p:ab:0', rotate: 0 }]);
    expect(s.keys()).toEqual(['1', blank.slice(2), '2', 'ab:0']);
  });

  test('실행 취소는 한 동작씩(내 편집만)', () => {
    const s = setup(3);
    const um = new Y.UndoManager(s.arr, { trackedOrigins: new Set([LOCAL_ORIGIN]), captureTimeout: 0 });
    s.ops.remove(['o:0']);
    s.ops.move(['o:2'], 0);
    expect(s.keys()).toEqual(['2', '1']);
    um.undo();
    expect(s.keys()).toEqual(['1', '2']);
    um.undo();
    expect(s.keys()).toEqual(['0', '1', '2']);
  });
});
