/**
 * 쪽 목록(방의 pdfPages) 바꾸기 — 정리 화면이 쓴다. 모두 LOCAL_ORIGIN 트랜잭션(실행 취소 대상, 서버는 그 연결의 편집으로 기록).
 * 목록을 통째로 다시 쓰지 않고 바뀐 항목만 지우고 넣는다 — 두 사람이 동시에 다른 쪽을 고쳐도 둘 다 남게.
 * 같은 열쇠가 두 번 들어가는 경우(동시에 같은 쪽을 옮김)는 읽는 쪽(cleanPageList)에서 첫 번째만 쓴다.
 */
import * as Y from 'yjs';
import { cleanPageList, type PdfAnnot, type PdfPageEntry } from '../../pdf/annotTypes.js';
import { LOCAL_ORIGIN } from './annotLayer.js';

type Rotation = PdfPageEntry['rotate'];

export class PageOps {
  constructor(
    private readonly ydoc: Y.Doc,
    private readonly arr: Y.Array<PdfPageEntry>,
    private readonly annots: Y.Map<PdfAnnot>,
  ) {}

  /** 지금 목록(보이는 순서) */
  list(): PdfPageEntry[] {
    return cleanPageList(this.arr.toArray());
  }

  /** 열쇠의 실제 배열 자리(중복이면 첫 번째) */
  private rawIndexOf(key: string): number {
    return this.arr.toArray().findIndex((e) => e?.key === key);
  }

  rotate(keys: readonly string[], delta: 90 | -90): void {
    this.ydoc.transact(() => {
      for (const key of keys) {
        const i = this.rawIndexOf(key);
        if (i < 0) continue;
        const e = this.arr.get(i);
        this.arr.delete(i, 1);
        this.arr.insert(i, [{ ...e, rotate: (((e.rotate + delta) % 360) + 360) % 360 as Rotation }]);
      }
    }, LOCAL_ORIGIN);
  }

  /** 지운다 — 전부 지우려 하면 false(쪽이 하나는 있어야 한다) */
  remove(keys: readonly string[]): boolean {
    const drop = new Set(keys);
    if (this.list().every((e) => drop.has(e.key))) return false;
    this.ydoc.transact(() => {
      const raw = this.arr.toArray();
      for (let i = raw.length - 1; i >= 0; i -= 1) if (raw[i] && drop.has(raw[i].key)) this.arr.delete(i, 1);
    }, LOCAL_ORIGIN);
    return true;
  }

  /** at: 보이는 목록에서의 자리(0이면 맨 앞, 길이면 맨 끝) */
  private rawInsertIndex(at: number): number {
    const visible = this.list();
    if (at <= 0) return 0;
    if (at >= visible.length) return this.arr.length;
    return this.rawIndexOf(visible[at].key);
  }

  insert(at: number, entries: readonly PdfPageEntry[], annots: readonly PdfAnnot[] = []): void {
    this.ydoc.transact(() => {
      this.arr.insert(this.rawInsertIndex(at), [...entries]);
      for (const a of annots) this.annots.set(a.id, a);
    }, LOCAL_ORIGIN);
  }

  insertBlank(at: number, size: { w: number; h: number }): string {
    const key = `b:${Math.random().toString(36).slice(2, 12)}`;
    this.insert(at, [{ key, rotate: 0, blank: { w: Math.round(size.w), h: Math.round(size.h) } }]);
    return key;
  }

  /** 고른 쪽들을(지금 순서대로) 보이는 목록의 to 자리 앞으로 옮긴다(to는 옮기기 전 목록 기준) */
  move(keys: readonly string[], to: number): void {
    const visible = this.list();
    const chosen = new Set(keys);
    const moving = visible.filter((e) => chosen.has(e.key));
    if (moving.length === 0) return;
    const before = visible.slice(0, Math.max(0, Math.min(to, visible.length))).filter((e) => !chosen.has(e.key));
    const anchorKey = visible.slice(Math.max(0, to)).find((e) => !chosen.has(e.key))?.key ?? null;
    const unchanged = visible.filter((e) => !chosen.has(e.key));
    const nextOrder = [...before, ...moving, ...unchanged.slice(before.length)].map((e) => e.key);
    if (nextOrder.every((k, i) => k === visible[i].key)) return;
    this.ydoc.transact(() => {
      const raw = this.arr.toArray();
      for (let i = raw.length - 1; i >= 0; i -= 1) if (raw[i] && chosen.has(raw[i].key)) this.arr.delete(i, 1);
      const at = anchorKey === null ? this.arr.length : this.rawIndexOf(anchorKey);
      this.arr.insert(at < 0 ? this.arr.length : at, moving);
    }, LOCAL_ORIGIN);
  }
}
