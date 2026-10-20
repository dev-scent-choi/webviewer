/**
 * PDF 편집 방의 Y.Doc 모델 — 원본 PDF는 그대로 두고 바뀐 것만 담는다(설계 docs/superpowers/specs/2026-10-20-pdf-editor-design.md §2.2).
 * 1단계에는 편집 도구가 없어 맵이 비어 있고, 저장은 원본 바이트를 그대로 돌려준다. 2단계 이후가 exportPdfFromRoom을 채운다
 * (주석·쪽 정리·양식).
 */
import * as Y from 'yjs';

/** 최종 쪽 순서(쪽 정리) */
export const PDF_PAGES_ARRAY = 'pdfPages';
/** 주석 id → 주석 */
export const PDF_ANNOTS_MAP = 'pdfAnnots';
/** 양식 칸 이름 → 값 */
export const PDF_FORM_MAP = 'pdfForm';

export function pdfModelIsEmpty(ydoc: Y.Doc): boolean {
  return ydoc.getArray(PDF_PAGES_ARRAY).length === 0 && ydoc.getMap(PDF_ANNOTS_MAP).size === 0 && ydoc.getMap(PDF_FORM_MAP).size === 0;
}

export async function exportPdfFromRoom(original: Uint8Array, ydoc: Y.Doc): Promise<Uint8Array> {
  if (pdfModelIsEmpty(ydoc)) return original;
  // 2단계부터 채운다 — 그 전까지는 편집기에 모델을 바꾸는 도구가 없다
  return original;
}

/** 변경 기록(GET /api/view/changes) — 2단계부터 주석·쪽 정리·양식 값을 남긴다 */
export function pdfChangeLogOf(_ydoc: Y.Doc): readonly object[] {
  return [];
}
