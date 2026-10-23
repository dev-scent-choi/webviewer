/**
 * PDF 쪽의 글 조각(pdf.js items)을 줄 → 문단(블록)으로 묶고 제목 블록을 가린다 — 찾기 패널이 "결과가 속한 제목",
 * "그 문단", "결과 안에서 다시 찾기(블록 단위)"에 쓴다. DOM 없이 도는 순수 함수(tests/pdfSearchEngine.test.ts).
 *
 * PDF에는 대개 문단·제목 정보가 없어 줄마다 여러 단서를 점수로 더해 제목을 가린다(headingScore, 3점 이상이면 제목):
 *   글자 크기(본문 대비) · 굵기·다른 글꼴 · 쪽 위쪽 · 번호 패턴(1. / 제1장 / Ⅰ. / 가.) · 제목 단어(개요·목적…) ·
 *   위아래 여백 · 짧은 줄 / 긴 문장·마침표(−) · 숫자·기호뿐인 도면 치수(−) · 표처럼 칸이 벌어진 줄(−).
 * 책갈피 제목과 같은 줄, 태그 PDF의 H1~H6 은 확실한 제목(+10), 태그 P 는 본문(−3).
 * 제목을 찾은 뒤 그 아래 글은 그 제목의 본문 블록으로 묶는다.
 */
import { normalizeForFind } from './findText.js';

export interface StructItem {
  readonly str: string;
  /** [a, b, c, d, e, f] — e, f가 기준점(PDF 좌표, y는 위로) */
  readonly transform: readonly number[];
  readonly width: number;
  readonly height: number;
  /** pdf.js 글꼴 열쇠(같은 글꼴이면 같은 값) */
  readonly fontName?: string;
  /** 글꼴이 굵은지(글꼴 이름·속성으로 판단) */
  readonly bold?: boolean;
  /** 태그 PDF의 구조 역할(H1~H6, P 등) — 없으면 null */
  readonly role?: string | null;
}

export interface AnalyzeOptions {
  /** 쪽 높이(PDF 단위) — 쪽 위쪽 판단에 쓴다 */
  readonly pageHeight?: number;
}

export interface TextBlock {
  /** 이 블록에 든 조각 번호(쪽의 items 순번) */
  readonly items: readonly number[];
  readonly text: string;
  readonly isHeading: boolean;
  /** 제목 단계 — 1 큰 제목, 2 소제목, 본문은 0 */
  readonly level: 0 | 1 | 2;
  /** 이 블록이 속한 제목(제목 블록 자신이면 자기 글) — 쪽 앞에 제목이 없으면 앞 쪽에서 이어진 제목, 그것도 없으면 '' */
  readonly heading: string;
  /** PDF 좌표 [x0, y0, x1, y1] (y0 < y1) */
  readonly bbox: readonly [number, number, number, number];
  /** 제목 점수(블록 첫 줄) — 확인·조정용 */
  readonly score: number;
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
  /** 줄 글자 중 굵은 글자 비율 */
  boldShare: number;
  /** 가장 많이 쓴 글꼴 */
  font: string;
  /** 칸 사이가 크게 벌어진 곳 수(표 줄 단서) */
  bigGaps: number;
  role: string | null;
}

/** 줄 하나의 제목 단서 — 점수와 이유(시험·조정용) */
export interface HeadingSignals {
  readonly sizeRatio: number;
  readonly bold: boolean;
  readonly fontDiffers: boolean;
  readonly nearTop: boolean;
  readonly numbered: boolean;
  readonly keyword: boolean;
  readonly spaced: boolean;
  readonly text: string;
  readonly bigGaps: number;
  readonly role: string | null;
  readonly inOutline: boolean;
}

export const HEADING_THRESHOLD = 3;

const HEADING_NUMBER = /^(제\s*\d+\s*[장절편관조항부]|[0-9]{1,2}(\.[0-9]{1,2})*\.?\s|[0-9]{1,2}\)\s|\([0-9]{1,2}\)\s?|[IVX]{1,4}\.\s|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\.?\s?|[가-하]\.\s|[①-⑳]\s?|\[[^\]]{1,20}\]$|【[^】]{1,20}】|<[^>]{1,20}>$|《[^》]{1,20}》$)/;
const HEADING_WORDS = /(개요|목적|적용\s?범위|범위|시스템\s?구성|구성|배경|서론|결론|요약|일반\s?사항|정의|용어|참고|부록|목차|추진\s?(배경|계획|방향)|현황|문제점|개선\s?(방안|사항)|기대\s?효과|세부\s?(내용|사항)|주요\s?(내용|사항)|유의\s?사항|주의\s?사항|절차|방법|기준|설계|사양|운영|관리|점검)$/;
const BOLD_FONT_NAME = /(bold|black|heavy|extrabold|semibold|demibold|extra|headline|견고딕|굵은|-b$|_b$|,b$)/i;

/** 글꼴 이름으로 굵기 짐작 — pdf.js 글꼴 속성(bold/black)이 없을 때 */
export function isBoldFontName(name: string | undefined): boolean {
  return !!name && BOLD_FONT_NAME.test(name.replace(/^[A-Z]{6}\+/, ''));
}

/** 줄 하나의 제목 점수 — 붙여 받은 기준표를 따른 가중치(실제 문서로 조정할 값) */
export function headingScore(s: HeadingSignals): number {
  const len = s.text.length;
  let score = 0;
  if (s.inOutline) score += 10;
  if (s.role && /^H[1-6]?$/i.test(s.role)) score += 10;
  if (s.role && /^(P|LBody|Span|Caption|Note|TD|TH|Figure)$/i.test(s.role)) score -= 3;
  if (s.sizeRatio >= 1.5) score += 4;
  else if (s.sizeRatio >= 1.15) score += 3;
  else if (s.sizeRatio >= 1.05) score += 1;
  else if (s.sizeRatio < 0.95) score -= 2;
  if (s.bold) score += 2;
  else if (s.fontDiffers) score += 1;
  if (s.nearTop) score += 1;
  if (s.numbered) score += 3;
  if (s.keyword && len <= 30) score += 1;
  if (s.spaced) score += 1;
  if (len <= 40) score += 1;
  if (len > 60 || /(다|요|음|함|임)\.$|[.。]$/.test(s.text)) score -= 2;
  // 숫자·기호뿐인 줄(도면 치수 300, 1:50, Ø25, 표의 숫자 칸) — 글자가 커도 제목이 될 수 없게 크게 뺀다
  if (/^[\d\s.,:;%Øø±×xX*/()+\-=~·°'"mMcCkKgGlL]+$/.test(s.text)) score -= 6;
  if (s.bigGaps >= 2) score -= 1;
  return score;
}

const sizeOf = (it: StructItem) => (it.height > 0 ? it.height : Math.hypot(it.transform[2], it.transform[3])) || 0;

function median(values: number[]): number {
  if (values.length === 0) return 0;
  const s = [...values].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)];
}

/** 쪽 하나를 블록으로 — prevHeading 은 앞 쪽에서 이어지는 제목, outlineTitles 는 책갈피 제목(정규화해서 비교) */
export function analyzePage(
  page: number,
  items: readonly StructItem[],
  prevHeading = '',
  outlineTitles: ReadonlySet<string> = new Set(),
  opts: AnalyzeOptions = {},
): PageStructure {
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
      lines.push({ items: [i], text: '', y, size, x0: x, x1: x + it.width, boldShare: 0, font: '', bigGaps: 0, role: null });
    }
  }
  // 줄 글은 왼쪽부터 — 조각 사이가 떨어져 있으면 한 칸 띄운다(살짝 높이가 다른 조각이 먼저 들어와도 순서가 맞게)
  for (const line of lines) {
    line.items.sort((a, b) => items[a].transform[4] - items[b].transform[4]);
    let text = '';
    let right = -Infinity;
    let boldChars = 0;
    let chars = 0;
    const fontChars = new Map<string, number>();
    for (const i of line.items) {
      const it = items[i];
      const x = it.transform[4];
      const gap = x - right;
      if (text && gap > Math.max(sizeOf(it), line.size) * 0.25 && !text.endsWith(' ') && !it.str.startsWith(' ')) text += ' ';
      if (text && gap > line.size * 2.5) line.bigGaps += 1;
      text += it.str;
      right = Math.max(right, x + it.width);
      const n = it.str.trim().length;
      chars += n;
      if (it.bold) boldChars += n;
      if (it.fontName) fontChars.set(it.fontName, (fontChars.get(it.fontName) ?? 0) + n);
      if (it.role && !line.role) line.role = it.role;
    }
    line.text = text.replace(/\s+/g, ' ').trim();
    line.boldShare = chars ? boldChars / chars : 0;
    line.font = [...fontChars.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? '';
  }
  // 본문 크기·글꼴 — 글자 수로 가중한 중앙값·최빈값(제목 몇 줄에 끌려가지 않게)
  const weighted: number[] = [];
  const bodyFonts = new Map<string, number>();
  let allChars = 0;
  let boldChars = 0;
  for (const line of lines) {
    for (let k = 0; k < Math.min(line.text.length, 200); k += 1) weighted.push(line.size);
    allChars += line.text.length;
    boldChars += line.text.length * line.boldShare;
  }
  const body = median(weighted) || 10;
  for (const line of lines) if (Math.abs(line.size - body) <= body * 0.05 && line.font) bodyFonts.set(line.font, (bodyFonts.get(line.font) ?? 0) + line.text.length);
  const bodyFont = [...bodyFonts.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? '';
  // 쪽 대부분이 굵으면 굵기는 단서가 못 된다
  const boldMeaningful = allChars === 0 || boldChars / allChars < 0.6;
  const pageTop = opts.pageHeight ?? (lines.length ? Math.max(...lines.map((l) => l.y + l.size)) : 0);

  const scores = lines.map((line, k) => {
    const above = k > 0 ? lines[k - 1].y - line.y : Infinity;
    const below = k < lines.length - 1 ? line.y - lines[k + 1].y : Infinity;
    const gapUnit = Math.max(line.size, body);
    return headingScore({
      text: line.text,
      sizeRatio: line.size / body,
      bold: boldMeaningful && line.boldShare >= 0.6,
      fontDiffers: !!bodyFont && !!line.font && line.font !== bodyFont,
      nearTop: opts.pageHeight !== undefined && line.y >= pageTop * 0.85,
      numbered: HEADING_NUMBER.test(line.text),
      keyword: HEADING_WORDS.test(line.text.replace(HEADING_NUMBER, '').trim()),
      spaced: above >= gapUnit * 1.8 || below >= gapUnit * 1.8,
      bigGaps: line.bigGaps,
      role: line.role,
      inOutline: outlineTitles.size > 0 && outlineTitles.has(normalizeForFind(line.text).trim()),
    });
  });
  const isHeadingLine = (k: number) => lines[k].text.length > 0 && lines[k].text.length <= 120 && scores[k] >= HEADING_THRESHOLD;

  const blocks: { lines: Line[]; isHeading: boolean; score: number }[] = [];
  let prev: Line | null = null;
  lines.forEach((line, k) => {
    const heading = isHeadingLine(k);
    const current = blocks[blocks.length - 1];
    const gap = prev ? prev.y - line.y : 0;
    const sameKind = current && current.isHeading === heading;
    const sizeJump = prev && Math.abs(prev.size - line.size) > Math.max(prev.size, line.size) * 0.15;
    // 제목은 한 줄씩 따로(번호 붙은 제목 줄이 이어져도 각각 다른 절이다), 본문은 줄 간격이 벌어지면 새 문단
    if (current && sameKind && !heading && !sizeJump && gap <= Math.max(prev!.size, line.size) * 1.75) current.lines.push(line);
    else blocks.push({ lines: [line], isHeading: heading, score: scores[k] });
    prev = line;
  });

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
    const first = b.lines[0];
    const level: 0 | 1 | 2 = !b.isHeading ? 0 : first.size >= body * 1.4 || /^H1$/i.test(first.role ?? '') ? 1 : 2;
    out.push({ items: blockItems, text, isHeading: b.isHeading, level, heading, bbox: [x0, bottom, x1, top], score: b.score });
  });
  return { page, blocks: out, blockOfItem, text: out.map((b) => b.text).join('\n\n'), lastHeading: heading };
}
