/**
 * PDF 편집 방의 Y.Doc 모델 — 원본 PDF는 그대로 두고 바뀐 것만 담는다(설계 docs/superpowers/specs/2026-10-20-pdf-editor-design.md §2.2).
 *
 * 2단계(주석): 편집기가 다루는 종류의 주석(annotRead.ts)은 방을 열 때 파일에서 읽어 `pdfAnnots`에 넣고(PDF_SEED_ORIGIN — 편집이
 * 아니므로 자동저장·변경 기록에서 뺀다), 편집기에는 그 주석을 뺀 PDF를 준다(basePdfForEditor — 편집기가 직접 그리므로 두 번
 * 그려지지 않게). 저장은 원본에서 그 주석을 빼고 맵의 주석을 다시 쓴다(annotWrite.ts). 바꾼 것이 없으면 원본 그대로다.
 */
import { PDFDocument } from 'pdf-lib';
import * as Y from 'yjs';
import { readManagedAnnots, stripManagedAnnots } from './annotRead.js';
import { writeAnnots } from './annotWrite.js';
import { ANNOT_LABELS, origIndexOf, type PdfAnnot } from './annotTypes.js';

/** 최종 쪽 순서(3단계 쪽 정리) */
export const PDF_PAGES_ARRAY = 'pdfPages';
/** 주석 id → 주석 */
export const PDF_ANNOTS_MAP = 'pdfAnnots';
/** 양식 칸 이름 → 값(4단계) */
export const PDF_FORM_MAP = 'pdfForm';
/** 방을 열 때 파일 주석을 맵에 넣는 트랜잭션 표시 — 편집이 아니다 */
export const PDF_SEED_ORIGIN = Symbol('pdf-seed');

export function pdfModelIsEmpty(ydoc: Y.Doc): boolean {
  return ydoc.getArray(PDF_PAGES_ARRAY).length === 0 && ydoc.getMap(PDF_ANNOTS_MAP).size === 0 && ydoc.getMap(PDF_FORM_MAP).size === 0;
}

async function loadPdf(bytes: Uint8Array): Promise<PDFDocument> {
  return PDFDocument.load(bytes, { updateMetadata: false });
}

/** 방을 열 때 — 파일의 (다루는) 주석을 맵에 넣는다. pdf-lib로 못 여는 파일(암호 등)이면 비워 둔다. */
export async function seedPdfModel(original: Uint8Array, ydoc: Y.Doc): Promise<void> {
  let annots: PdfAnnot[] = [];
  try {
    annots = readManagedAnnots(await loadPdf(original));
  } catch (error) {
    console.error('PDF 주석 읽기 실패(주석 없이 연다):', error);
    return;
  }
  const map = ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP);
  ydoc.transact(() => {
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
export async function exportPdfFromRoom(original: Uint8Array, ydoc: Y.Doc, edited = true): Promise<Uint8Array> {
  if (!edited) return original;
  const doc = await loadPdf(original);
  stripManagedAnnots(doc);
  const annots = [...ydoc.getMap<PdfAnnot>(PDF_ANNOTS_MAP).values()].sort(
    (a, b) => (origIndexOf(a.pageKey) ?? 0) - (origIndexOf(b.pageKey) ?? 0) || a.createdAt.localeCompare(b.createdAt),
  );
  await writeAnnots(doc, annots);
  return doc.save();
}

/** 변경 기록 한 줄 — 같은 주석은 마지막 상태로 하나만 */
export interface PdfChangeEntry {
  readonly kind: 'annot_added' | 'annot_changed' | 'annot_removed';
  readonly annotType: string;
  readonly label: string;
  readonly page: number | null;
  readonly text?: string;
  readonly viewer: string | null;
  readonly at: string;
}

/** 주석 추가·변경·삭제를 열람자·시각과 함께 모은다(시딩은 빼고). 방을 연 뒤 생겼다가 지운 주석은 남기지 않는다. */
export class PdfChangeTracker {
  private readonly entries = new Map<string, PdfChangeEntry>();
  private readonly seeded = new Set<string>();

  constructor(
    private readonly ydoc: Y.Doc,
    private readonly resolveViewer: (origin: unknown) => string | null,
  ) {
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
        const page = origIndexOf(annot.pageKey);
        const common = { annotType: annot.type, label: ANNOT_LABELS[annot.type], page: page === null ? null : page + 1, viewer, at, ...(annot.text ? { text: annot.text } : {}) };
        if (change.action === 'delete') {
          if (this.seeded.has(key)) this.entries.set(key, { kind: 'annot_removed', ...common });
          else this.entries.delete(key); // 이번에 만들었다가 지움 — 남길 것 없음
        } else if (this.seeded.has(key)) this.entries.set(key, { kind: 'annot_changed', ...common });
        else this.entries.set(key, { kind: 'annot_added', ...common });
      });
    });
  }

  list(): readonly PdfChangeEntry[] {
    return [...this.entries.values()];
  }
}

export function pdfChangeLogOf(tracker: PdfChangeTracker | undefined): readonly object[] {
  return tracker?.list() ?? [];
}
