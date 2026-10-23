/**
 * PDF 쪽의 글 조각(pdf.js items)을 줄 → 문단(블록)으로 묶고 제목 블록을 가린다 — 찾기 패널이 "결과가 속한 제목",
 * "그 문단", "결과 안에서 다시 찾기(블록 단위)"에 쓴다. DOM 없이 도는 순수 함수(tests/pdfTextStructure.test.ts).
 *
 * PDF에는 문단·제목 정보가 없으므로 위치로 추정한다: 같은 높이의 조각은 한 줄, 줄 사이가 크게 벌어지거나 글자 크기가
 * 바뀌면 새 블록. 제목은 본문보다 글자가 확실히 크거나, 짧은 한 줄이 "제1장·1.·Ⅰ.·가." 같은 번호로 시작하거나,
 * 문서 책갈피 제목과 같은 줄이다.
 */
import { normalizeForFind } from './findText.js';

export interface StructItem {
  readonly str: string;
  /** [a, b, c, d, e, f] — e, f가 기준점(PDF 좌표, y는 위로) */
  readonly transform: readonly number[];
  readonly width: number;
  readonly height: number;
}

export interface TextBlock {
  /** 이 블록에 든 조각 번호(쪽의 items 순번) */
  readonly items: readonly number[];
  readonly text: string;
  readonly isHeading: boolean;
  /** 이 블록이 속한 제목(제목 블록 자신이면 자기 글) — 쪽 앞에 제목이 없으면 앞 쪽에서 이어진 제목, 그것도 없으면 '' */
  readonly heading: string;
  /** PDF 좌표 [x0, y0, x1, y1] (y0 < y1) */
  readonly bbox: readonly [number, number, number, number];
}

export interface PageStructure {
  readonly page: number;
  readonly blocks: readonly TextBlock[];
  /** 조각 번호 → 블록 번호(-1 = 빈 조각) */
  readonly blockOfItem: readonly number[];
  /** 쪽 전체 글(블록 사이 빈 줄) */
  readonly text: string;
  /** 이 쪽 끝에서 이어지는 제목 — 다음 쪽 첫 블록들의 heading */
  readonly lastHeading: string;
}

interface Line {
  items: number[];
  text: string;
  y: number;
  size: number;
  x0: number;
  x1: number;
}

const HEADING_NUMBER = /^(제\s*\d+\s*[장절편관조항]|[0-9]{1,2}(\.[0-9]{1,2})*\.?\s|[0-9]{1,2}\)\s|[IVX]{1,4}\.\s|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\.?\s?|[가-하]\.\s|\[[^\]]{1,20}\]$|【[^】]{1,20}】)/;

const sizeOf = (it: StructItem) => (it.height > 0 ? it.height : Math.hypot(it.transform[2], it.transform[3])) || 0;

function median(values: number[]): number {
  if (values.length === 0) return 0;
  const s = [...values].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}

/** 쪽 하나를 블록으로 — prevHeading 은 앞 쪽에서 이어지는 제목, outlineTitles 는 책갈피 제목(정규화해서 비교) */
export function analyzePage(page: number, items: readonly StructItem[], prevHeading = '', outlineTitles: ReadonlySet<string> = new Set()): PageStructure {
  const idx = items.map((_, i) => i).filter((i) => items[i].str.trim() !== '');
  // 위에서 아래(y 큰 것부터), 같은 높이면 왼쪽부터
  idx.sort((a, b) => items[b].transform[5] - items[a].transform[5] || items[a].transform[4] - items[b].transform[4]);
  const lines: Line[] = [];
  for (const i of idx) {
    const it = items[i];
    const size = sizeOf(it);
    const y = it.transform[5];
    const x = it.transform[4];
    const line = lines[lines.length - 1];
    if (line && Math.abs(line.y - y) <= Math.max(line.size, size) * 0.5) {
      line.items.push(i);
      line.size = Math.max(line.size, size);
      line.x0 = Math.min(line.x0, x);
      line.x1 = Math.max(line.x1, x + it.width);
    } else {
      lines.push({ items: [i], text: '', y, size, x0: x, x1: x + it.width });
    }
  }
  // 줄 글은 왼쪽부터 — 조각 사이가 떨어져 있으면 한 칸 띄운다(살짝 높이가 다른 조각이 먼저 들어와도 순서가 맞게)
  for (const line of lines) {
    line.items.sort((a, b) => items[a].transform[4] - items[b].transform[4]);
    let text = '';
    let right = -Infinity;
    for (const i of line.items) {
      const it = items[i];
      const x = it.transform[4];
      if (text && x - right > Math.max(sizeOf(it), line.size) * 0.25 && !text.endsWith(' ') && !it.str.startsWith(' ')) text += ' ';
      text += it.str;
      right = Math.max(right, x + it.width);
    }
    line.text = text.replace(/\s+/g, ' ').trim();
  }
  // 본문 크기 — 글자 수로 가중한 중앙값(제목 몇 줄에 끌려가지 않게)
  const weighted: number[] = [];
  for (const line of lines) for (let k = 0; k < Math.min(line.text.length, 200); k += 1) weighted.push(line.size);
  const body = median(weighted) || 10;
  const isHeadingLine = (line: Line) => {
    const len = line.text.length;
    if (len === 0 || len > 80) return false;
    if (outlineTitles.size > 0 && outlineTitles.has(normalizeForFind(line.text).trim())) return true;
    if (line.size >= body * 1.18) return true;
    return len <= 50 && line.size >= body * 0.95 && HEADING_NUMBER.test(line.text) && !/[.。]$/.test(line.text);
  };

  const blocks: { lines: Line[]; isHeading: boolean }[] = [];
  let prev: Line | null = null;
  for (const line of lines) {
    const heading = isHeadingLine(line);
    const current = blocks[blocks.length - 1];
    const gap = prev ? prev.y - line.y : 0;
    const sameKind = current && current.isHeading === heading;
    const sizeJump = prev && Math.abs(prev.size - line.size) > Math.max(prev.size, line.size) * 0.15;
    // 제목은 한 줄씩 따로(번호 붙은 제목 줄이 이어져도 각각 다른 절이다), 본문은 줄 간격이 벌어지면 새 문단
    if (current && sameKind && !heading && !sizeJump && gap <= Math.max(prev!.size, line.size) * 1.75) current.lines.push(line);
    else blocks.push({ lines: [line], isHeading: heading });
    prev = line;
  }

  let heading = prevHeading;
  const out: TextBlock[] = [];
  const blockOfItem = items.map(() => -1);
  blocks.forEach((b, bi) => {
    const text = b.lines.map((l) => l.text).join('\n');
    if (b.isHeading) heading = text.replace(/\n/g, ' ');
    const blockItems = b.lines.flatMap((l) => l.items);
    for (const i of blockItems) blockOfItem[i] = bi;
    const x0 = Math.min(...b.lines.map((l) => l.x0));
    const x1 = Math.max(...b.lines.map((l) => l.x1));
    const top = b.lines[0].y + b.lines[0].size;
    const bottom = b.lines[b.lines.length - 1].y - b.lines[b.lines.length - 1].size * 0.25;
    out.push({ items: blockItems, text, isHeading: b.isHeading, heading, bbox: [x0, bottom, x1, top] });
  });
  return { page, blocks: out, blockOfItem, text: out.map((b) => b.text).join('\n\n'), lastHeading: heading };
}
