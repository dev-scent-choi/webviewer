/**
 * PDF 편집기 주석 형식 — 서버(읽기·쓰기)와 브라우저(그리기·도구)가 함께 쓴다. 좌표는 원래 쪽의 PDF 좌표(회전 전, 아래가 y=0).
 * 주석은 쪽 열쇠(pageKey)에 붙는다 — 원본 쪽 'o:<0부터>', 넣은 PDF의 쪽 'p:<부속 id>:<0부터>', 빈 쪽 'b:<무작위>'.
 * 쪽 목록(pdfPages)도 같은 열쇠를 써서, 순서가 바뀌거나 쪽이 지워져도 주석은 그 쪽을 따라간다.
 */
/** PDF 편집 방의 Y.Doc 이름들 — 브라우저(편집기)와 서버(pdfModel.ts)가 같이 쓴다 */
/** 최종 쪽 순서(3단계 쪽 정리) */
export const PDF_PAGES_ARRAY = 'pdfPages';
/** 주석 id → 주석 */
export const PDF_ANNOTS_MAP = 'pdfAnnots';
/** 양식 칸 이름 → 값(4단계) */
export const PDF_FORM_MAP = 'pdfForm';
/** OCR(8단계) — 쪽 열쇠 → 인식한 낱말들. 저장할 때 보이지 않는 글자 층(렌더 모드 3)으로 쓴다 */
export const PDF_OCR_MAP = 'pdfOcr';

/** 인식한 낱말 하나 — 기준선 시작(x, y)·끝(ex, ey)과 글자 높이 h(PDF 좌표). 방향이 있어 회전한 쪽에도 맞는다 */
export interface OcrWord {
  readonly t: string;
  readonly x: number;
  readonly y: number;
  readonly ex: number;
  readonly ey: number;
  readonly h: number;
}

export interface PdfOcrPage {
  readonly words: readonly OcrWord[];
  readonly author: string;
  readonly at: string;
}

export function isOcrPage(v: unknown): v is PdfOcrPage {
  const p = v as PdfOcrPage;
  return (
    !!p &&
    Array.isArray(p.words) &&
    p.words.length <= 20000 &&
    p.words.every((w) => w && typeof w.t === 'string' && w.t.length <= 200 && [w.x, w.y, w.ex, w.ey, w.h].every((n) => typeof n === 'number' && Number.isFinite(n)))
  );
}

/** 문서 속성(6단계) — title·author·subject·keywords → 글 */
export const PDF_PROPS_MAP = 'pdfProps';
/** 기존 글자 고치기(5단계) — id → PdfTextEdit */
export const PDF_TEXT_EDITS_MAP = 'pdfTextEdits';
/** 가리기(5단계) — id → PdfRedaction. applied가 true인 것만 저장 때 실제로 지운다 */
export const PDF_REDACT_MAP = 'pdfRedactions';

export interface PdfTextEdit {
  readonly id: string;
  readonly pageKey: string;
  /** 고친 글 조각이 있던 자리(PDF 좌표) */
  readonly rect: readonly [number, number, number, number];
  readonly oldText: string;
  readonly newText: string;
  readonly fontSize: number;
  readonly author: string;
  readonly createdAt: string;
}

export interface PdfRedaction {
  readonly id: string;
  readonly pageKey: string;
  readonly rect: readonly [number, number, number, number];
  readonly applied: boolean;
  readonly author: string;
  readonly createdAt: string;
}

const finiteRect = (r: unknown): r is [number, number, number, number] =>
  Array.isArray(r) && r.length === 4 && r.every((v) => typeof v === 'number' && Number.isFinite(v)) && r[2] > r[0] && r[3] > r[1];

export function isTextEdit(v: unknown): v is PdfTextEdit {
  const e = v as PdfTextEdit;
  return !!e && typeof e.id === 'string' && typeof e.pageKey === 'string' && finiteRect(e.rect) && typeof e.oldText === 'string' && typeof e.newText === 'string' && e.newText.length <= 2000 && typeof e.fontSize === 'number';
}
export function isRedaction(v: unknown): v is PdfRedaction {
  const r = v as PdfRedaction;
  return !!r && typeof r.id === 'string' && typeof r.pageKey === 'string' && finiteRect(r.rect) && typeof r.applied === 'boolean';
}

/** 방 정보 — base: 방을 열 때의 원본 지문(편집기가 받은 원본과 같은지 확인) */
export const PDF_META_MAP = 'pdfMeta';

/** 쪽 목록 한 항목. rotate는 쪽 자체의 /Rotate에 더하는 각도. blank면 그 크기(pt)의 빈 쪽 */
export interface PdfPageEntry {
  readonly key: string;
  readonly rotate: 0 | 90 | 180 | 270;
  readonly blank?: { readonly w: number; readonly h: number };
}

export type ParsedPageKey = { kind: 'orig'; index: number } | { kind: 'part'; partId: string; index: number } | { kind: 'blank' };

export function parsePageKey(key: string): ParsedPageKey | null {
  let m = /^o:(\d+)$/.exec(key);
  if (m) return { kind: 'orig', index: Number(m[1]) };
  m = /^p:([A-Za-z0-9_-]{1,40}):(\d+)$/.exec(key);
  if (m) return { kind: 'part', partId: m[1], index: Number(m[2]) };
  if (/^b:[A-Za-z0-9_-]{1,40}$/.test(key)) return { kind: 'blank' };
  return null;
}

/** 같은 열쇠가 두 번 있으면(동시에 옮긴 경우 등) 첫 번째만, 형식이 틀린 항목은 뺀다 */
export function cleanPageList(entries: readonly PdfPageEntry[]): PdfPageEntry[] {
  const seen = new Set<string>();
  const out: PdfPageEntry[] = [];
  for (const e of entries) {
    if (!e || typeof e.key !== 'string' || seen.has(e.key)) continue;
    const parsed = parsePageKey(e.key);
    if (!parsed) continue;
    if (parsed.kind === 'blank' && !(e.blank && e.blank.w > 0 && e.blank.h > 0)) continue;
    seen.add(e.key);
    out.push(e);
  }
  return out;
}

export function identityPages(count: number): PdfPageEntry[] {
  return Array.from({ length: count }, (_, i) => ({ key: `o:${i}`, rotate: 0 as const }));
}

export type PdfAnnotType = 'note' | 'highlight' | 'underline' | 'strike' | 'ink' | 'freetext' | 'rect' | 'ellipse' | 'line' | 'arrow' | 'stamp';

export const PDF_ANNOT_TYPES: readonly PdfAnnotType[] = ['note', 'highlight', 'underline', 'strike', 'ink', 'freetext', 'rect', 'ellipse', 'line', 'arrow', 'stamp'];

export interface PdfAnnot {
  readonly id: string;
  readonly pageKey: string;
  readonly type: PdfAnnotType;
  /** x1, y1, x2, y2 (x1<x2, y1<y2) */
  readonly rect: readonly [number, number, number, number];
  /** 형광펜·밑줄·취소선: 줄마다 8개(왼위 x,y · 오른위 · 왼아래 · 오른아래 — Acrobat 순서) */
  readonly quads?: readonly number[];
  /** 펜: 획마다 [x, y, x, y, …] */
  readonly paths?: readonly (readonly number[])[];
  /** 선·화살표: 시작 x, y, 끝 x, y */
  readonly line?: readonly [number, number, number, number];
  /** '#RRGGBB' */
  readonly color: string;
  /** 사각형·타원 안쪽 색(없으면 비움) */
  readonly fill?: string | null;
  /** 0~1 */
  readonly opacity: number;
  /** 선 굵기(pt) */
  readonly width: number;
  /** 메모 내용·글상자 글 */
  readonly text?: string;
  readonly fontSize?: number;
  /** 서명·도장(stamp): 그림 data URL(image/png 또는 image/jpeg, base64) */
  readonly image?: string;
  readonly author: string;
  readonly createdAt: string;
  readonly modifiedAt: string;
}

export function pageKeyOf(origIndex: number): string {
  return `o:${origIndex}`;
}

/** 'o:3' → 3, 형식이 다르면 null */
export function origIndexOf(pageKey: string): number | null {
  const m = /^o:(\d+)$/.exec(pageKey);
  return m ? Number(m[1]) : null;
}

/** 종류별 기본 색·굵기·투명도(Acrobat 기본값과 비슷하게) */
export const ANNOT_DEFAULTS: Record<PdfAnnotType, { color: string; width: number; opacity: number }> = {
  note: { color: '#FFD84D', width: 1, opacity: 1 },
  highlight: { color: '#FFE600', width: 1, opacity: 0.5 },
  underline: { color: '#1A73E8', width: 1, opacity: 1 },
  strike: { color: '#D93025', width: 1, opacity: 1 },
  ink: { color: '#D93025', width: 2, opacity: 1 },
  freetext: { color: '#1F1F23', width: 1, opacity: 1 },
  rect: { color: '#D93025', width: 2, opacity: 1 },
  ellipse: { color: '#D93025', width: 2, opacity: 1 },
  line: { color: '#D93025', width: 2, opacity: 1 },
  arrow: { color: '#D93025', width: 2, opacity: 1 },
  stamp: { color: '#000000', width: 0, opacity: 1 },
};

export const ANNOT_LABELS: Record<PdfAnnotType, string> = {
  note: '메모',
  highlight: '형광펜',
  underline: '밑줄',
  strike: '취소선',
  ink: '펜',
  freetext: '글상자',
  rect: '사각형',
  ellipse: '타원',
  line: '선',
  arrow: '화살표',
  stamp: '서명·도장',
};

/** 서명·도장 그림으로 받는 data URL(크기 상한 — Y.Doc·저장 파일이 커지지 않게) */
export const STAMP_IMAGE_MAX_CHARS = 1_500_000;
export function isStampImage(value: unknown): value is string {
  return typeof value === 'string' && value.length <= STAMP_IMAGE_MAX_CHARS && /^data:image\/(png|jpeg);base64,[A-Za-z0-9+/=]+$/.test(value);
}
