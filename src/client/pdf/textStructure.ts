/**
 * PDF 쪽의 글 조각(pdf.js items)을 줄 → 덩어리(블록)로 묶고, 덩어리마다 제목인지 본문인지 가린다 — 찾기 패널이
 * "결과가 속한 제목(절)", "그 덩어리", "제목만/본문만", "결과 안에서(절 단위)"에 쓴다. DOM 없는 순수 함수.
 *
 * 순서가 중요하다(실제 문서로 검증하며 정한 것 — scripts/analyze-pdf-structure.ts):
 *  1. 줄 만들기: 같은 높이의 조각을 왼쪽부터 잇는다. 줄마다 서식 열쇠(글꼴·크기·굵기)를 단다.
 *  2. 문서 기준(buildProfile): 여러 쪽에서 가장 많이 쓴 서식을 본문으로, 보통 줄 간격과 여러 쪽에 반복되는
 *     머리말·쪽 번호를 구한다. 쪽 하나만 보면 조문 상자·해설이 섞인 쪽에서 본문 기준이 흔들렸다.
 *  3. 덩어리 묶기: 서식이 같고 줄 간격이 보통이면 한 덩어리 — 두 줄로 꺾인 질문은 한 제목, 두 줄로 이어진 조문은
 *     한 본문이 된다. 번호·글머리로 시작하는 줄은 새 덩어리.
 *  4. 덩어리 점수(blockScore, 3점 이상이면 제목): 본문보다 큰 글자·굵기·다른 글꼴, 서식이 다른 번호 제목,
 *     여백으로 떨어진 짧은 한 줄, 굵은 질문(?)은 +. 여러 줄·긴 글·'~다.'로 끝나는 문장·글머리 목록·숫자뿐인
 *     줄(도면 치수)·표 줄은 −. 책갈피 제목·태그 PDF의 H1~H6 은 확정(+10), 태그 P 는 −3.
 * 제목 아래의 덩어리는 다음 제목까지 그 제목의 절(본문)로 묶인다(쪽을 넘어서도 이어진다).
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
  /** 글꼴이 굵은지(글꼴 속성·이름으로 판단) */
  readonly bold?: boolean;
  /** 태그 PDF의 구조 역할(H1~H6, P 등) — 없으면 null */
  readonly role?: string | null;
}

/** 문서 전체의 본문 기준 — 여러 쪽에서 구한다(buildProfile) */
export interface DocProfile {
  /** 본문 서식 열쇠(글꼴|크기|굵기) */
  readonly bodyStyle: string;
  readonly bodyFont: string;
  readonly bodySize: number;
  readonly bodyBold: boolean;
  /** 보통 줄 간격 ÷ 글자 크기 */
  readonly lineGap: number;
  /** 여러 쪽의 위·아래 끝에 반복되는 머리말·꼬리말(숫자는 #) */
  readonly repeated: ReadonlySet<string>;
}

export interface AnalyzeOptions {
  /** 쪽 높이(PDF 단위) — 머리말·꼬리말 자리 판단에 쓴다 */
  readonly pageHeight?: number;
  /** 문서 기준 — 없으면 이 쪽만으로 구한다 */
  readonly profile?: DocProfile;
}

export interface TextBlock {
  /** 이 블록에 든 조각 번호(쪽의 items 순번) */
  readonly items: readonly number[];
  readonly text: string;
  readonly isHeading: boolean;
  /** 제목 단계 — 1 큰 제목, 2 소제목, 본문 0 */
  readonly level: 0 | 1 | 2;
  /** 이 블록이 속한 제목(제목 블록 자신이면 자기 글) — 쪽 앞에 제목이 없으면 앞 쪽에서 이어진 제목 */
  readonly heading: string;
  /** 이 블록이 속한 절 열쇠(그 제목 블록의 "쪽:블록", 첫 제목 앞이면 '') — '결과 안에서'가 절 단위로 좁힌다 */
  readonly section: string;
  /** 머리말·꼬리말·쪽 번호 — 찾기에서 뺀다 */
  readonly furniture: boolean;
  /** 목차(점선 뒤 쪽 번호 줄, 목차 쪽 전체) — '전체'에서만 찾고 제목·본문 어느 쪽으로도 치지 않는다 */
  readonly toc: boolean;
  /** 본문 첫머리에 붙은 조 제목('제3조(목적)')의 글자 수 — 이 앞의 일치는 제목으로 친다(없으면 0) */
  readonly inlineTitle: number;
  /** PDF 좌표 [x0, y0, x1, y1] (y0 < y1) */
  readonly bbox: readonly [number, number, number, number];
  /** 제목 점수 — 확인·조정용 */
  readonly score: number;
}

export interface PageStructure {
  readonly page: number;
  readonly blocks: readonly TextBlock[];
  /** 조각 번호 → 블록 번호(-1 = 빈 조각) */
  readonly blockOfItem: readonly number[];
  /** 쪽 전체 글(블록 사이 빈 줄) */
  readonly text: string;
  /** 이 쪽 끝에서 이어지는 제목·절 — 다음 쪽 첫 블록들이 이어받는다 */
  readonly lastHeading: string;
  readonly lastSection: string;
}

interface Line {
  items: number[];
  text: string;
  y: number;
  size: number;
  x0: number;
  x1: number;
  bold: boolean;
  font: string;
  style: string;
  /** 칸 사이가 크게 벌어진 곳 수(표 줄 단서) — 맨 앞 번호 뒤 칸은 세지 않는다 */
  bigGaps: number;
  role: string | null;
}

export const HEADING_THRESHOLD = 3;

const HEADING_NUMBER = /^(제\s*\d+\s*[장절편관조항부]|[0-9]{1,2}(\.[0-9]{1,2})*\.?\s|[0-9]{1,2}\)\s|\([0-9]{1,2}\)\s?|[IVX]{1,4}\.\s|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]\.?\s?|[가-하]\.\s+(?=\S)|[①-⑳]\s?|\[[^\]]{1,20}\]$|【[^】]{1,20}】|<[^>]{1,20}>$|《[^》]{1,20}》$)/;
const LIST_START = /^([-–—·∙•‣▪▫◦○●◎※*]|\d{1,2}\)\s|[가-하]\)\s)/;
const HEADING_WORDS = /(개요|목적|적용\s?범위|범위|시스템\s?구성|구성|배경|서론|결론|요약|일반\s?(사항|정보)|정의|용어|참고(\s?자료)?|부록|목차|추진\s?(배경|계획|방향)|현황|문제점|개선\s?(방안|사항)|기대\s?효과|세부\s?(내용|사항)|주요\s?(내용|사항)|유의\s?사항|주의\s?사항|절차|방법|기준|설계|사양|운영|관리|점검|해설|예시|사례|질의\s?응답|도입\s?배경)$/;
const SENTENCE_END = /((다|요|음|함|임|됨|것|니다)\.|[.。])$/;
// 문단 첫머리에 붙은 조 제목 — '제3조(목적)', '제12조의2 (정의)'
const ARTICLE_TITLE = /^(제\s*\d+\s*조(의\s*\d+)?\s*[(（][^)）]{1,40}[)）])/;
// 목차 줄 — 점선(·, ., …)이 이어지다 쪽 번호로 끝난다(번호가 잘린 줄도 있다)
const TOC_LEADER = /[·.…ㆍ･‥]{5,}\s*[\divxlc]*$/i;
// 알림 상자 표지 — '참고!', '주의 사항!', '참조' 같은 것은 제목이 아니다
const ADMONITION = /^(참고|참조|주의|경고|알림|팁|노트|중요|위험|비고|note|tip|hint|warning|caution|important|danger|notice)(\s?(사항|항목|하십시오))?\s*[!:：]?$/i;
const NUMERIC_ONLY = /^[\d\s.,:;%Øø±×xX*/()+\-=~·°'"mMcCkKgGlL#]+$/;
const BOLD_FONT_NAME = /(bold|black|heavy|extrabold|semibold|demibold|extra|headline|견고딕|굵은|-b$|_b$|,b$)/i;

/** 글꼴 이름으로 굵기 짐작 — pdf.js 글꼴 속성(bold/black)이 없을 때 */
export function isBoldFontName(name: string | undefined): boolean {
  return !!name && BOLD_FONT_NAME.test(name.replace(/^[A-Z]{6}\+/, ''));
}

const sizeOf = (it: StructItem) => (it.height > 0 ? it.height : Math.hypot(it.transform[2], it.transform[3])) || 0;
const roundSize = (n: number) => Math.round(n * 2) / 2;
/** 머리말·꼬리말 비교용 — 숫자는 같은 것으로 */
const furnitureKey = (text: string) => normalizeForFind(text).replace(/\d+/g, '#').trim();

/** 덩어리로 묶어도 되는 같은 모양 — 크기·굵기가 같으면. 글꼴 열쇠는 보지 않는다(오래된 한글 PDF는 같은 본문을
 * Type3 글꼴 여러 개로 나눠 담아 줄마다 열쇠가 달랐다) */
const sameLook = (a: Line, b: Line) => Math.abs(a.size - b.size) <= 0.6 && a.bold === b.bold;

/** 쪽 조각을 줄로 — 위에서 아래, 줄 안은 왼쪽부터 */
export function buildLines(items: readonly StructItem[]): Line[] {
  const idx = items.map((_, i) => i).filter((i) => items[i].str.trim() !== '');
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
      lines.push({ items: [i], text: '', y, size, x0: x, x1: x + it.width, bold: false, font: '', style: '', bigGaps: 0, role: null });
    }
  }
  for (const line of lines) {
    line.items.sort((a, b) => items[a].transform[4] - items[b].transform[4]);
    let text = '';
    let right = -Infinity;
    let boldChars = 0;
    let chars = 0;
    const fontChars = new Map<string, number>();
    const sizeChars = new Map<number, number>();
    line.items.forEach((i, k) => {
      const it = items[i];
      const x = it.transform[4];
      const gap = x - right;
      // 낱말 사이(한글 띄어쓰기는 글자 폭의 0.2~0.3 남짓) — 0.25로는 경계에 걸린 줄이 붙어 버려 0.15로 본다
      if (text && gap > Math.max(sizeOf(it), line.size) * 0.15 && !text.endsWith(' ') && !it.str.startsWith(' ')) text += ' ';
      // 표 칸처럼 크게 벌어진 곳 — "1.1   대상"처럼 맨 앞 번호 뒤의 들여쓰기는 빼고 센다
      if (text && gap > line.size * 2.5 && !(k === 1 && HEADING_NUMBER.test(`${text.trim()} `))) line.bigGaps += 1;
      text += it.str;
      right = Math.max(right, x + it.width);
      const n = it.str.trim().length;
      chars += n;
      if (it.bold) boldChars += n;
      if (it.fontName) fontChars.set(it.fontName, (fontChars.get(it.fontName) ?? 0) + n);
      const s = roundSize(sizeOf(it));
      sizeChars.set(s, (sizeChars.get(s) ?? 0) + n);
      if (it.role && !line.role) line.role = it.role;
    });
    line.text = text.replace(/\s+/g, ' ').trim();
    line.bold = chars > 0 && boldChars / chars >= 0.6;
    line.font = [...fontChars.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] ?? '';
    // 줄 크기는 가장 많이 쓴 글자 크기(번호·기호 하나가 크다고 줄 전체가 커지지 않게)
    const mainSize = [...sizeChars.entries()].sort((a, b) => b[1] - a[1])[0]?.[0];
    if (mainSize) line.size = mainSize;
    line.style = `${line.font}|${roundSize(line.size)}|${line.bold ? 1 : 0}`;
  }
  return lines;
}

/** 여러 쪽에서 문서 기준을 구한다 — 쪽마다 {items, pageHeight} */
export function buildProfile(pages: readonly { items: readonly StructItem[]; pageHeight?: number }[]): DocProfile {
  const styleChars = new Map<string, { chars: number; font: string; size: number; bold: boolean }>();
  const gaps: number[] = [];
  const edgeCount = new Map<string, number>();
  for (const { items, pageHeight } of pages) {
    const lines = buildLines(items);
    const seen = new Set<string>();
    lines.forEach((line, k) => {
      const e = styleChars.get(line.style) ?? { chars: 0, font: line.font, size: roundSize(line.size), bold: line.bold };
      e.chars += line.text.length;
      styleChars.set(line.style, e);
      const next = lines[k + 1];
      if (next && next.style === line.style) {
        const g = (line.y - next.y) / line.size;
        if (g > 0.8 && g < 3) gaps.push(g);
      }
      // 쪽 위·아래 끝 10% 안의 줄 — 여러 쪽에 반복되면 머리말·꼬리말
      if (pageHeight && (line.y >= pageHeight * 0.9 || line.y <= pageHeight * 0.1)) {
        const key = furnitureKey(line.text);
        if (key && !seen.has(key)) {
          seen.add(key);
          edgeCount.set(key, (edgeCount.get(key) ?? 0) + 1);
        }
      }
    });
  }
  const body = [...styleChars.entries()].sort((a, b) => b[1].chars - a[1].chars)[0];
  const sortedGaps = gaps.sort((a, b) => a - b);
  const minRepeat = Math.max(3, Math.ceil(pages.length * 0.3));
  return {
    bodyStyle: body?.[0] ?? '',
    bodyFont: body?.[1].font ?? '',
    bodySize: body?.[1].size || 10,
    bodyBold: body?.[1].bold ?? false,
    lineGap: sortedGaps.length ? sortedGaps[Math.floor(sortedGaps.length / 2)] : 1.4,
    repeated: new Set([...edgeCount.entries()].filter(([key, n]) => n >= minRepeat && (key === '#' || key.length <= 60)).map(([key]) => key)),
  };
}

/** 덩어리 하나의 제목 단서 */
export interface BlockSignals {
  readonly text: string;
  readonly lines: number;
  readonly sizeRatio: number;
  readonly bold: boolean;
  readonly fontDiffers: boolean;
  /** 위로 보통 줄 간격보다 크게 떨어져 있다(쪽 첫 덩어리 포함) */
  readonly isolated: boolean;
  /** 바로 아래 덩어리가 글머리 목록이거나 조금 들여 쓴 글 — 목록을 이끄는 줄 */
  readonly introducesList: boolean;
  /** 바로 위가 표 줄(칸이 크게 벌어진 줄)이고 붙어 있다 — 표 칸 안에서 꺾인 글 */
  readonly inTable: boolean;
  /** 앞이나 뒤 덩어리도 같은 서식의 번호 줄 — 번호 목록의 항목(조문 안 1. 2. …) */
  readonly inNumberedList: boolean;
  /** 쪽 폭을 거의 다 채운 한 줄이 바로 아래 같은 모양의 줄로 이어진다 — 줄 간격이 넓은 큰 글씨 산문(발간사 등) */
  readonly wrapsIntoNext: boolean;
  readonly bigGaps: number;
  readonly role: string | null;
  readonly inOutline: boolean;
}

/** 덩어리 하나의 제목 점수 — 3점 이상이면 제목(가중치는 실제 문서로 맞춘 값) */
export function blockScore(s: BlockSignals): number {
  const text = s.text;
  const len = text.length;
  const numbered = HEADING_NUMBER.test(`${text} `);
  // 굵거나 본문보다 큰 서식만 강한 단서 — 글꼴만 다른 것(조문 상자 전체가 다른 글꼴인 경우 등)은 약하다
  const strongStyle = s.bold || s.sizeRatio >= 1.08;
  const question = /\?$/.test(text);
  let score = 0;
  if (s.inOutline) score += 10;
  if (s.role && /^H[1-6]?$/i.test(s.role)) score += 10;
  if (s.role && /^(P|LBody|Span|Caption|Note|TD|TH|Figure|Lbl|LI)$/i.test(s.role)) score -= 3;
  if (s.sizeRatio >= 1.5) score += 4;
  else if (s.sizeRatio >= 1.2) score += 3;
  else if (s.sizeRatio >= 1.08) score += 2;
  else if (s.sizeRatio < 0.97) score -= 2;
  const sentence = SENTENCE_END.test(text);
  if (s.bold) score += 2;
  else if (s.fontDiffers) score += 1;
  // 굵게 따로 선 한두 줄(문장이 아님) — 상자 제목·참고 제목
  if (s.bold && s.lines <= 2 && !sentence) score += 1;
  // 번호 제목: 서식이 본문과 다르거나, 여백으로 떨어진 짧은 한 줄일 때만 — 조문의 '가.'·'1)' 항목은 본문
  if (numbered && strongStyle) score += 3;
  else if (numbered && s.lines === 1 && len <= 30 && s.isolated && !s.inNumberedList) score += 2;
  if (HEADING_WORDS.test(text.replace(HEADING_NUMBER, '').trim()) && len <= 30) score += 1;
  if (question && strongStyle) score += 2;
  // 여백으로 떨어졌거나 아래 목록을 이끄는 짧은 한 줄 — '□ 절 제목'처럼 서식이 본문과 같은 제목.
  // 번호 줄은 위에서 따로 셌고(겹쳐 주지 않는다), 본문보다 작은 글씨(조문·각주)는 빼다
  const listItem = LIST_START.test(text);
  if (s.lines === 1 && len <= 40 && (s.isolated || s.introducesList) && !numbered && !listItem && !sentence && s.sizeRatio >= 0.97) score += 2;
  if (s.lines === 1 && len <= 40) score += 1;
  if (!question) {
    if (s.lines >= 3) score -= 3;
    if (len > 80) score -= 3;
    else if (len > 60) score -= 1;
  }
  if (sentence) score -= 3;
  // 문장 끝이 중간에 들어 있는 글('…됩니다. 자세한 내용은 …') — 여러 문장의 산문(표 칸에서 꺾인 조각 등)
  if (/(다|요|음|함)\.(?=\s|[가-힣A-Za-z0-9「(“])/.test(text)) score -= 3;
  // ':'로 끝나는 안내 문구('메시지:', '권장 조치:', '…하는 방법:')와 '이름: 설명' 꼴 — 제목이 아니라 본문의 표지
  if (/[:：]$/.test(text) || /^[^\s:：]{1,12}[:：]\s/.test(text)) score -= 3;
  // 파일 이름(표 칸의 xxx.wav 등)
  if (/\.[A-Za-z0-9]{2,4}$/.test(text) && !/\s/.test(text)) score -= 4;
  if (listItem) score -= 2;
  // 숫자·기호뿐인 줄(도면 치수 300, 1:50, Ø25, 표의 숫자 칸) — 글자가 커도 제목이 될 수 없게 크게 뺀다
  if (NUMERIC_ONLY.test(text)) score -= 6;
  if (ADMONITION.test(text)) score -= 6;
  // 글자(한글·영문)가 두 자도 안 되는 조각
  if ((text.match(/[가-힣A-Za-z]/g) ?? []).length < 2) score -= 6;
  if (s.inTable) score -= 4;
  if (s.wrapsIntoNext && !question) score -= 3;
  if (s.bigGaps >= 1) score -= 2;
  return score;
}

/** 쪽 하나를 블록으로 — prev* 는 앞 쪽에서 이어지는 제목·절, outlineTitles 는 책갈피 제목(정규화해서 비교) */
export function analyzePage(
  page: number,
  items: readonly StructItem[],
  prevHeading = '',
  outlineTitles: ReadonlySet<string> = new Set(),
  opts: AnalyzeOptions = {},
  prevSection = '',
): PageStructure {
  const lines = buildLines(items);
  const profile = opts.profile ?? buildProfile([{ items, pageHeight: opts.pageHeight }]);
  const body = profile.bodySize;
  const isFurniture = (line: Line) =>
    opts.pageHeight !== undefined && (line.y >= opts.pageHeight * 0.9 || line.y <= opts.pageHeight * 0.1) && (profile.repeated.has(furnitureKey(line.text)) || NUMERIC_ONLY.test(line.text));

  // 덩어리 묶기 — 서식이 같고 줄 간격이 보통이면 같은 덩어리, 번호·글머리로 시작하면 새 덩어리
  const normalGap = profile.lineGap + 0.55;
  const groups: { lines: Line[]; furniture: boolean; gapAbove: number }[] = [];
  lines.forEach((line, k) => {
    const prev = lines[k - 1];
    const current = groups[groups.length - 1];
    const furniture = isFurniture(line);
    const gap = prev ? (prev.y - line.y) / Math.max(line.size, prev.size) : Infinity;
    const startsItem = HEADING_NUMBER.test(`${line.text} `) || LIST_START.test(line.text);
    // 위로 떨어진 짧은 첫 줄 다음에 훨씬 긴 줄이 오면 첫 줄은 따로 — 서식이 본문과 같은 제목('제어 입력 구성' + 설명)
    const head = current?.lines.length === 1 ? current.lines[0] : null;
    const titleLine =
      !!head && head.text.length <= 40 && !SENTENCE_END.test(head.text) && !LIST_START.test(head.text) && current!.gapAbove > profile.lineGap * 1.35 && line.x1 - line.x0 > (head.x1 - head.x0) * 1.8;
    if (current && !furniture && !current.furniture && sameLook(prev, line) && gap <= normalGap && !startsItem && !titleLine) current.lines.push(line);
    else groups.push({ lines: [line], furniture, gapAbove: gap });
  });

  // 글 영역의 왼쪽·오른쪽 끝(머리말 빼고) — 쪽 폭을 채운 줄 판단에 쓴다
  const bodyLines = groups.filter((g) => !g.furniture).flatMap((g) => g.lines);
  const textLeft = bodyLines.length ? Math.min(...bodyLines.map((l) => l.x0)) : 0;
  const textRight = bodyLines.length ? Math.max(...bodyLines.map((l) => l.x1)) : 0;
  // 목차 쪽 — 점선 줄이 여럿이고 글 줄의 4분의 1을 넘으면 그 쪽의 '제1장 총칙' 같은 줄까지 모두 목차로 본다
  const leaderLines = bodyLines.filter((l) => TOC_LEADER.test(l.text)).length;
  const tocPage = leaderLines >= 3 && leaderLines >= bodyLines.length * 0.25;
  let heading = prevHeading;
  let section = prevSection;
  const out: TextBlock[] = [];
  const blockOfItem = items.map(() => -1);
  groups.forEach((g, bi) => {
    const text = g.lines.map((l) => l.text).join('\n');
    const flat = text.replace(/\n/g, ' ');
    const first = g.lines[0];
    const chars = g.lines.reduce((n, l) => n + l.text.length, 0) || 1;
    const boldShare = g.lines.reduce((n, l) => n + (l.bold ? l.text.length : 0), 0) / chars;
    const size = g.lines.reduce((n, l) => n + l.size * l.text.length, 0) / chars;
    const font = first.font;
    const next = groups[bi + 1];
    const nx = next?.lines[0];
    const introducesList = !!nx && !next.furniture && (LIST_START.test(nx.text) || (nx.x0 > first.x0 + first.size * 0.8 && nx.x0 < first.x0 + first.size * 6));
    const prevLine = groups[bi - 1]?.lines[groups[bi - 1].lines.length - 1];
    const inTable = !!prevLine && prevLine.bigGaps >= 1 && g.gapAbove <= normalGap;
    // 번호 목록: 바로 앞이나 뒤 덩어리도 같은 서식의 번호 줄이면(사이에 다른 글 없이) 목록 항목이다 — '1. 개요 … 본문 … 2. 목적'은 아님
    const numberedSameStyle = (o: (typeof groups)[number] | undefined) => !!o && !o.furniture && sameLook(o.lines[0], first) && HEADING_NUMBER.test(`${o.lines[0].text} `);
    const inNumberedList = numberedSameStyle(groups[bi - 1]) || numberedSameStyle(groups[bi + 1]);
    const lastLine = g.lines[g.lines.length - 1];
    const wrapsIntoNext =
      !!nx && !next.furniture && sameLook(lastLine, nx) && lastLine.x1 >= textRight - lastLine.size * 2 && lastLine.x1 - lastLine.x0 >= (textRight - textLeft) * 0.6 && (lastLine.y - nx.y) / lastLine.size <= 2.8;
    const toc = !g.furniture && (tocPage || g.lines.every((l) => TOC_LEADER.test(l.text)));
    const score = g.furniture || toc
      ? -10
      : blockScore({
          text: flat,
          lines: g.lines.length,
          sizeRatio: size / body,
          bold: boldShare >= 0.6 && !profile.bodyBold,
          fontDiffers: !!font && !!profile.bodyFont && font !== profile.bodyFont,
          isolated: g.gapAbove > profile.lineGap * 1.35,
          introducesList,
          inTable,
          inNumberedList,
          wrapsIntoNext,
          bigGaps: Math.max(...g.lines.map((l) => l.bigGaps)),
          role: g.lines.find((l) => l.role)?.role ?? null,
          inOutline: outlineTitles.size > 0 && outlineTitles.has(normalizeForFind(flat).trim()),
        });
    const isHeading = !g.furniture && !toc && flat.length <= 160 && score >= HEADING_THRESHOLD;
    let inlineTitle = 0;
    if (isHeading) {
      heading = flat;
      section = `${page}:${bi}`;
    } else if (!g.furniture && !toc) {
      // '제3조(목적) 이 규정은 …'처럼 조 제목이 문단 첫머리에 붙은 본문 — 그 조 제목으로 새 절을 연다
      const article = ARTICLE_TITLE.exec(flat)?.[1];
      if (article) {
        heading = article.replace(/\s+/g, ' ').trim();
        section = `${page}:${bi}`;
        inlineTitle = article.length;
      }
    }
    const blockItems = g.lines.flatMap((l) => l.items);
    for (const i of blockItems) blockOfItem[i] = bi;
    const x0 = Math.min(...g.lines.map((l) => l.x0));
    const x1 = Math.max(...g.lines.map((l) => l.x1));
    const top = first.y + first.size;
    const last = g.lines[g.lines.length - 1];
    const bottom = last.y - last.size * 0.25;
    const level: 0 | 1 | 2 = !isHeading ? 0 : size >= body * 1.25 || /^H1$/i.test(first.role ?? '') ? 1 : 2;
    out.push({ items: blockItems, text, isHeading, level, heading: g.furniture ? '' : toc ? '목차' : heading, section: g.furniture || toc ? '' : section, furniture: g.furniture, toc, inlineTitle, bbox: [x0, bottom, x1, top], score });
  });
  return {
    page,
    blocks: out,
    blockOfItem,
    text: out.filter((b) => !b.furniture).map((b) => b.text).join('\n\n'),
    lastHeading: heading,
    lastSection: section,
  };
}
