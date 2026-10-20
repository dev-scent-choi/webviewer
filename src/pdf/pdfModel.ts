/**
 * PDF 편집 방의 Y.Doc 모델 — 원본 PDF는 그대로 두고 바뀐 것만 담는다(설계 docs/superpowers/specs/2026-10-20-pdf-editor-design.md §2.2).
 *
 * 3단계(쪽 정리): 쪽 목록(`pdfPages`)은 방을 열 때 원본 순서로 채운다. 저장은 늘 **방을 열 때의 원본**에서 다시 만든다 —
 * 저장본에 목록을 또 적용하면 순서·삭제가 두 번 먹으므로(room.ts의 originalPdfBytes). 넣은 PDF(부속)의 쪽은 copyPages로 가져온다.
 *
 * 2단계(주석): 편집기가 다루는 종류의 주석(annotRead.ts)은 방을 열 때 파일에서 읽어 `pdfAnnots`에 넣고(PDF_SEED_ORIGIN — 편집이
 * 아니므로 자동저장·변경 기록에서 뺀다), 편집기에는 그 주석을 뺀 PDF를 준다(basePdfForEditor — 편집기가 직접 그리므로 두 번
 * 그려지지 않게). 저장은 원본에서 그 주석을 빼고 맵의 주석을 다시 쓴다(annotWrite.ts). 바꾼 것이 없으면 원본 그대로다.
 */
import { createHash } from 'node:crypto';
import { degrees, PDFArray, PDFDict, PDFDocument, PDFName, PDFPageTree, PDFRef, PDFStream, type PDFObject, type PDFPage } from 'pdf-lib';
import * as Y from 'yjs';
import { readManagedAnnots, stripManagedAnnots } from './annotRead.js';
import { fontsFor, writeAnnots } from './annotWrite.js';
import { applyFormValues, dropFieldsOnRemovedPages } from './pdfForm.js';
import {
  ANNOT_LABELS,
  cleanPageList,
  identityPages,
  origIndexOf,
  parsePageKey,
  PDF_ANNOTS_MAP,
  PDF_FORM_MAP,
  PDF_META_MAP,
  PDF_PAGES_ARRAY,
  type PdfAnnot,
  type PdfPageEntry,
} from './annotTypes.js';

export { PDF_ANNOTS_MAP, PDF_FORM_MAP, PDF_META_MAP, PDF_PAGES_ARRAY };

/** 원본 지문 — 편집기가 받은 원본이 방의 원본과 같은지 확인한다 */
export function pdfBaseFingerprint(bytes: Uint8Array): string {
  return createHash('sha1').update(bytes).digest('hex').slice(0, 16);
}

/** 넣은 PDF(부속) 바이트를 돌려준다 — 없으면 undefined(그 쪽은 빠진다) */
export type PdfPartSource = (partId: string) => Uint8Array | undefined;
/** 방을 열 때 파일 주석을 맵에 넣는 트랜잭션 표시 — 편집이 아니다 */
export const PDF_SEED_ORIGIN = Symbol('pdf-seed');

export function pdfModelIsEmpty(ydoc: Y.Doc): boolean {
  return ydoc.getArray(PDF_PAGES_ARRAY).length === 0 && ydoc.getMap(PDF_ANNOTS_MAP).size === 0 && ydoc.getMap(PDF_FORM_MAP).size === 0;
}

async function loadPdf(bytes: Uint8Array): Promise<PDFDocument> {
  return PDFDocument.load(bytes, { updateMetadata: false });
}

/** 방을 열 때 — 원본 지문, 쪽 목록(원본 순서), 파일의 (다루는) 주석을 넣는다. pdf-lib로 못 여는 파일(암호 등)이면 지문만. */
export async function seedPdfModel(original: Uint8Array, ydoc: Y.Doc): Promise<void> {
  const meta = ydoc.getMap<unknown>(PDF_META_MAP);
  let annots: PdfAnnot[] = [];
  let pageCount = 0;
  try {
    const doc = await loadPdf(original);
    annots = readManagedAnnots(doc);
    pageCount = doc.getPageCount();
  } catch (error) {
    console.error('PDF 읽기 실패(주석·쪽 정리 없이 연다):', error);
  }
  const map = ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP);
  const pages = ydoc.getArray<PdfPageEntry>(PDF_PAGES_ARRAY);
  ydoc.transact(() => {
    meta.set('base', pdfBaseFingerprint(original));
    if (pages.length === 0 && pageCount > 0) pages.push(identityPages(pageCount));
    for (const a of annots) map.set(a.id, a);
  }, PDF_SEED_ORIGIN);
}

/** 편집기가 그릴 PDF — 다루는 주석을 뺀 것(편집기가 자기 층에 그린다). 못 열면 그대로. */
export async function basePdfForEditor(bytes: Uint8Array): Promise<Uint8Array> {
  try {
    const doc = await loadPdf(bytes);
    stripManagedAnnots(doc);
    return await doc.save();
  } catch {
    return bytes;
  }
}

/**
 * 저장할 바이트. edited가 false면(방을 연 뒤 사용자가 바꾼 것이 없으면) 원본 그대로 — 열기만 해도 파일이 다시 쓰이지 않게.
 */
export async function exportPdfFromRoom(original: Uint8Array, ydoc: Y.Doc, edited = true, parts?: PdfPartSource): Promise<Uint8Array> {
  if (!edited) return original;
  return (await buildPdfFromRoom(original, ydoc, parts)).bytes;
}

/** 저장 결과와 그 쪽들의 열쇠(최종 순서) — 추출이 쪽 열쇠를 최종 순번으로 바꿀 때 쓴다 */
export async function buildPdfFromRoom(original: Uint8Array, ydoc: Y.Doc, parts?: PdfPartSource): Promise<{ bytes: Uint8Array; keys: string[] }> {
  const doc = await loadPdf(original);
  stripManagedAnnots(doc);
  const raw = ydoc.getArray<PdfPageEntry>(PDF_PAGES_ARRAY).toArray();
  const pageOfKey = await arrangePages(doc, raw.length > 0 ? cleanPageList(raw) : identityPages(doc.getPageCount()), parts);
  const order = new Map([...pageOfKey.keys()].map((key, i) => [key, i]));
  const annots = [...ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP).values()]
    .filter((a) => order.has(a.pageKey))
    .sort((a, b) => order.get(a.pageKey)! - order.get(b.pageKey)! || a.createdAt.localeCompare(b.createdAt));
  const fonts = fontsFor(doc);
  await writeAnnots(doc, annots, (key) => pageOfKey.get(key), fonts);
  // 양식 값(바꾼 칸만) — 한글은 주석과 같은 Pretendard 한 벌을 쓴다
  await applyFormValues(doc, ydoc.getMap<unknown>(PDF_FORM_MAP).entries(), () => fonts.korean());
  return { bytes: await doc.save(), keys: [...pageOfKey.keys()] };
}

/** 상속되는 쪽 속성 — 쪽 트리를 새로 짜기 전에 각 쪽에 직접 적어 둔다(중간 노드에서 물려받던 값이 사라지지 않게) */
const INHERITABLE = ['Resources', 'MediaBox', 'CropBox', 'Rotate'].map((n) => PDFName.of(n));

/**
 * 쪽 목록대로 문서의 쪽을 다시 짠다. 목록이 원본 그대로(순서·회전 같음)면 아무것도 바꾸지 않는다.
 * 원래 쪽은 같은 객체를 다시 끼우고(양식·책갈피가 그 쪽을 그대로 가리키게), 부속 쪽은 copyPages, 빈 쪽은 새로 만든다.
 * 열쇠 → 최종 쪽(주석 쓰기용). 원래 쪽을 하나라도 지웠으면 어디서도 가리키지 않는 객체를 지운다(지운 쪽 내용이 파일에 남지 않게).
 */
async function arrangePages(doc: PDFDocument, entries: readonly PdfPageEntry[], parts?: PdfPartSource): Promise<Map<string, PDFPage>> {
  const origPages = doc.getPages();
  const identity = entries.length === origPages.length && entries.every((e, i) => e.key === `o:${i}` && e.rotate === 0);
  const result = new Map<string, PDFPage>();
  if (identity) {
    origPages.forEach((p, i) => result.set(`o:${i}`, p));
    return result;
  }
  const partDocs = new Map<string, PDFDocument | null>();
  const partDoc = async (partId: string): Promise<PDFDocument | null> => {
    if (!partDocs.has(partId)) {
      const bytes = parts?.(partId);
      partDocs.set(partId, bytes ? await loadPdf(bytes).catch(() => null) : null);
    }
    return partDocs.get(partId) ?? null;
  };
  const ordered: Array<{ entry: PdfPageEntry; page: PDFPage }> = [];
  for (const entry of entries) {
    const parsed = parsePageKey(entry.key);
    let page: PDFPage | undefined;
    if (parsed?.kind === 'orig') page = origPages[parsed.index];
    else if (parsed?.kind === 'part') {
      const src = await partDoc(parsed.partId);
      if (src && parsed.index < src.getPageCount()) [page] = await doc.copyPages(src, [parsed.index]);
    } else if (parsed?.kind === 'blank' && entry.blank) {
      // addPage는 쪽 트리 끝에 붙인다 — 아래에서 트리를 새로 짜므로 자리는 상관없다
      page = doc.addPage([entry.blank.w, entry.blank.h]);
    }
    if (page) ordered.push({ entry, page });
  }
  if (ordered.length === 0) {
    origPages.forEach((p, i) => result.set(`o:${i}`, p));
    return result;
  }
  const ctx = doc.context;
  for (const { page } of ordered) {
    for (const name of INHERITABLE) {
      if (page.node.get(name) !== undefined) continue;
      const inherited = page.node.getInheritableAttribute(name);
      if (inherited !== undefined) page.node.set(name, inherited);
    }
  }
  const tree = PDFPageTree.withContext(ctx);
  const pagesRef = ctx.register(tree);
  for (const { page } of ordered) tree.pushLeafNode(page.ref);
  for (const { entry, page } of ordered) {
    page.node.set(PDFName.of('Parent'), pagesRef);
    if (entry.rotate) page.setRotation(degrees((page.getRotation().angle + entry.rotate) % 360));
    result.set(entry.key, page);
  }
  doc.catalog.set(PDFName.of('Pages'), pagesRef);
  const kept = new Set(ordered.map(({ page }) => page.ref.tag));
  const removedPages = origPages.filter((p) => !kept.has(p.ref.tag));
  if (removedPages.length > 0) {
    const removedAnnots = new Set<string>();
    for (const p of removedPages) {
      const annots = p.node.Annots();
      for (const item of annots?.asArray() ?? []) if (item instanceof PDFRef) removedAnnots.add(item.tag);
    }
    dropFieldsOnRemovedPages(doc, removedAnnots);
    dropUnreachable(doc);
  }
  return result;
}

/** 문서 끝(trailer)에서 닿지 않는 간접 객체를 지운다 */
function dropUnreachable(doc: PDFDocument): void {
  const ctx = doc.context;
  const seen = new Set<string>();
  const stack: PDFObject[] = [];
  const t = ctx.trailerInfo;
  for (const o of [t.Root, t.Info, t.Encrypt, t.ID]) if (o) stack.push(o);
  while (stack.length > 0) {
    const o = stack.pop()!;
    if (o instanceof PDFRef) {
      if (seen.has(o.tag)) continue;
      seen.add(o.tag);
      const v = ctx.lookup(o);
      if (v) stack.push(v);
    } else if (o instanceof PDFDict) {
      for (const [, v] of o.entries()) stack.push(v);
    } else if (o instanceof PDFArray) {
      for (const v of o.asArray()) stack.push(v);
    } else if (o instanceof PDFStream) {
      stack.push(o.dict);
    }
  }
  for (const [ref] of ctx.enumerateIndirectObjects()) if (!seen.has(ref.tag)) ctx.delete(ref);
}

/** 이미 만든 PDF에서 쪽 묶음마다 새 PDF(쪽 순번은 0부터) — 추출·나누기 */
export async function extractPdf(full: Uint8Array, groups: readonly (readonly number[])[]): Promise<Uint8Array[]> {
  const src = await loadPdf(full);
  const count = src.getPageCount();
  const out: Uint8Array[] = [];
  for (const group of groups) {
    const indices = group.filter((i) => Number.isInteger(i) && i >= 0 && i < count);
    if (indices.length === 0) continue;
    const doc = await PDFDocument.create();
    for (const page of await doc.copyPages(src, indices)) doc.addPage(page);
    out.push(await doc.save());
  }
  return out;
}

/** 넣을 PDF(부속) 준비 — 다루는 주석은 빼서 그 부속의 쪽 열쇠·새 id로 돌려주고(편집기가 맵에 넣는다), 바이트는 주석을 뺀 것.
 * 암호가 걸렸거나 PDF가 아니면 예외. */
export async function preparePdfPart(
  bytes: Uint8Array,
  partId: string,
): Promise<{ bytes: Uint8Array; pageCount: number; sizes: Array<{ w: number; h: number; rotate: number }>; annots: PdfAnnot[] }> {
  const doc = await loadPdf(bytes);
  const pageCount = doc.getPageCount();
  if (pageCount === 0) throw new Error('쪽이 없는 PDF입니다.');
  const annots = readManagedAnnots(doc).map((a, n) => ({ ...a, id: `${partId}-${n}`, pageKey: `p:${partId}:${origIndexOf(a.pageKey) ?? 0}` }));
  stripManagedAnnots(doc);
  const sizes = doc.getPages().map((p) => ({ w: p.getWidth(), h: p.getHeight(), rotate: p.getRotation().angle }));
  return { bytes: await doc.save(), pageCount, sizes, annots };
}

/** 변경 기록 한 줄 — 같은 주석은 마지막 상태로 하나만, 쪽 정리는 요약 한 줄 */
export interface PdfChangeEntry {
  readonly kind: 'annot_added' | 'annot_changed' | 'annot_removed' | 'pages_changed' | 'form_changed';
  readonly annotType?: string;
  readonly label: string;
  readonly page: number | null;
  readonly text?: string;
  readonly viewer: string | null;
  readonly at: string;
}

/** 처음 쪽 목록과 지금 목록의 차이를 한 줄로 — 같으면 null */
export function summarizePageChanges(initial: readonly PdfPageEntry[], current: readonly PdfPageEntry[]): string | null {
  const before = new Map(initial.map((e) => [e.key, e]));
  const now = new Map(current.map((e) => [e.key, e]));
  const removed = initial.filter((e) => !now.has(e.key)).length;
  const added = current.filter((e) => !before.has(e.key)).length;
  const rotated = current.filter((e) => before.has(e.key) && before.get(e.key)!.rotate !== e.rotate).length;
  const commonBefore = initial.filter((e) => now.has(e.key)).map((e) => e.key);
  const commonNow = current.filter((e) => before.has(e.key)).map((e) => e.key);
  const reordered = commonBefore.some((k, i) => k !== commonNow[i]);
  const parts = [removed ? `${removed}쪽 지움` : '', added ? `${added}쪽 넣음` : '', rotated ? `${rotated}쪽 돌림` : '', reordered ? '순서 바꿈' : ''].filter(Boolean);
  return parts.length ? parts.join(' · ') : null;
}

/** 주석 추가·변경·삭제와 쪽 정리를 열람자·시각과 함께 모은다(시딩은 빼고). 방을 연 뒤 생겼다가 지운 주석은 남기지 않는다. */
export class PdfChangeTracker {
  private readonly entries = new Map<string, PdfChangeEntry>();
  private readonly seeded = new Set<string>();
  private initialPages: PdfPageEntry[];
  private pagesEdit: { viewer: string | null; at: string } | null = null;

  constructor(
    private readonly ydoc: Y.Doc,
    private readonly resolveViewer: (origin: unknown) => string | null,
  ) {
    const pages = ydoc.getArray<PdfPageEntry>(PDF_PAGES_ARRAY);
    this.initialPages = cleanPageList(pages.toArray());
    pages.observe((_event, transaction) => {
      if (transaction.origin === PDF_SEED_ORIGIN) {
        this.initialPages = cleanPageList(pages.toArray());
        return;
      }
      this.pagesEdit = { viewer: this.resolveViewer(transaction.origin), at: new Date().toISOString() };
    });
    // 양식 값 — 칸마다 마지막 값 한 줄
    const form = ydoc.getMap<unknown>(PDF_FORM_MAP);
    form.observe((event, transaction) => {
      if (transaction.origin === PDF_SEED_ORIGIN) return;
      const viewer = this.resolveViewer(transaction.origin);
      const at = new Date().toISOString();
      for (const name of event.keysChanged) {
        const value = form.get(name);
        const shown = value === undefined ? '(되돌림)' : typeof value === 'boolean' ? (value ? '체크' : '체크 해제') : Array.isArray(value) ? value.join(', ') : String(value);
        this.entries.set(`form:${name}`, { kind: 'form_changed', label: '양식', page: null, text: `${name}: ${shown.slice(0, 200)}`, viewer, at });
      }
    });
    const map = ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP);
    for (const key of map.keys()) this.seeded.add(key);
    map.observe((event, transaction) => {
      if (transaction.origin === PDF_SEED_ORIGIN) {
        for (const key of event.keysChanged) this.seeded.add(key);
        return;
      }
      const viewer = this.resolveViewer(transaction.origin);
      const at = new Date().toISOString();
      event.changes.keys.forEach((change, key) => {
        const now = map.get(key);
        const before = change.oldValue as PdfAnnot | undefined;
        const annot = now ?? before;
        if (!annot) return;
        const common = { annotType: annot.type, label: ANNOT_LABELS[annot.type], page: this.pageNumberOf(annot.pageKey), viewer, at, ...(annot.text ? { text: annot.text } : {}) };
        if (change.action === 'delete') {
          if (this.seeded.has(key)) this.entries.set(key, { kind: 'annot_removed', ...common });
          else this.entries.delete(key); // 이번에 만들었다가 지움 — 남길 것 없음
        } else if (this.seeded.has(key)) this.entries.set(key, { kind: 'annot_changed', ...common });
        else this.entries.set(key, { kind: 'annot_added', ...common });
      });
    });
  }

  /** 지금 쪽 목록에서 몇 쪽인지(1부터). 목록이 없으면 원본 순번, 지운 쪽이면 null */
  private pageNumberOf(pageKey: string): number | null {
    const pages = this.ydoc.getArray<PdfPageEntry>(PDF_PAGES_ARRAY).toArray();
    if (pages.length === 0) {
      const i = origIndexOf(pageKey);
      return i === null ? null : i + 1;
    }
    const i = pages.findIndex((e) => e.key === pageKey);
    return i < 0 ? null : i + 1;
  }

  list(): readonly PdfChangeEntry[] {
    const out = [...this.entries.values()];
    if (this.pagesEdit) {
      const summary = summarizePageChanges(this.initialPages, cleanPageList(this.ydoc.getArray<PdfPageEntry>(PDF_PAGES_ARRAY).toArray()));
      if (summary) out.push({ kind: 'pages_changed', label: '쪽 정리', page: null, text: summary, ...this.pagesEdit });
    }
    return out;
  }
}

export function pdfChangeLogOf(tracker: PdfChangeTracker | undefined): readonly object[] {
  return tracker?.list() ?? [];
}
