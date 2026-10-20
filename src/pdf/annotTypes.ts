/**
 * PDF 편집기 주석 형식 — 서버(읽기·쓰기)와 브라우저(그리기·도구)가 함께 쓴다. 좌표는 원래 쪽의 PDF 좌표(회전 전, 아래가 y=0).
 * 주석은 "원본 몇 쪽"(pageKey 'o:<0부터>')에 붙는다 — 3단계 쪽 정리에서 순서가 바뀌어도 그 쪽을 따라간다.
 */
export type PdfAnnotType = 'note' | 'highlight' | 'underline' | 'strike' | 'ink' | 'freetext' | 'rect' | 'ellipse' | 'line' | 'arrow';

export const PDF_ANNOT_TYPES: readonly PdfAnnotType[] = ['note', 'highlight', 'underline', 'strike', 'ink', 'freetext', 'rect', 'ellipse', 'line', 'arrow'];

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
};
