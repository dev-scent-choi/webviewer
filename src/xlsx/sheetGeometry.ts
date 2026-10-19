/**
 * 엑셀 시트의 열 너비·행 높이를 화면 픽셀로 바꾸는 규칙과, 행마다 높이가 다른 시트를 가상 스크롤하기
 * 위한 RowLayout. 서버(spreadsheet.ts·lazyGridStore.ts)와 그리드 편집기 번들이 함께 쓰므로 node
 * 전용 모듈을 import하지 않는다(viewer.js에는 같은 규칙을 JS로 옮겨 두었다).
 */

/** 기본 글꼴(맑은 고딕/Calibri 11pt)의 숫자 한 자 폭(px) — 엑셀 열 너비 단위의 기준 */
export const EXCEL_MAX_DIGIT_WIDTH_PX = 7;
/** 시트에 기본 행 높이가 없을 때 엑셀이 쓰는 값(pt) */
export const EXCEL_DEFAULT_ROW_HEIGHT_PT = 15;

/** 파일에 저장된 열 너비(문자 단위, 여백 포함) → 화면 px. OOXML 명세의 변환식 그대로다.
 * mdw는 통합 문서 기본 글꼴의 숫자 한 자 폭(maxDigitWidthPx). */
export function excelColWidthToPx(width: number, mdw = EXCEL_MAX_DIGIT_WIDTH_PX): number {
  return Math.trunc(((256 * width + Math.trunc(128 / mdw)) / 256) * mdw);
}

/** 너비를 따로 정하지 않은 열의 폭 — sheetFormatPr의 defaultColWidth가 있으면 그것, 없으면
 * baseColWidth(기본 8자)에 여백 5px을 더해 8px 단위로 올린 값(Calibri 기본 64px). */
export function defaultColWidthPx(defaultColWidth?: number, baseColWidth = 8, mdw = EXCEL_MAX_DIGIT_WIDTH_PX): number {
  if (defaultColWidth && defaultColWidth > 0) return excelColWidthToPx(defaultColWidth, mdw);
  return Math.ceil((baseColWidth * mdw + 5) / 8) * 8;
}

/** 글꼴별 숫자 폭(글자 크기 대비) — 엑셀이 화면(96dpi)에서 재는 값에 맞춘 근사치. 한글 글꼴은 숫자가
 * 넓다(예: 돋움 11pt는 한 자 9px — 같은 "너비 3.66"이 Calibri 26px, 돋움 33px로 그려진다). */
const DIGIT_WIDTH_RATIO: Readonly<Record<string, number>> = {
  calibri: 0.507,
  arial: 0.556,
  'times new roman': 0.5,
  tahoma: 0.546,
  verdana: 0.636,
  '맑은 고딕': 0.55,
  'malgun gothic': 0.55,
  '나눔고딕': 0.55,
  'nanumgothic': 0.55,
  '돋움': 0.6,
  dotum: 0.6,
  '돋움체': 0.6,
  dotumche: 0.6,
  '굴림': 0.6,
  gulim: 0.6,
  '굴림체': 0.6,
  gulimche: 0.6,
  '바탕': 0.6,
  batang: 0.6,
  '바탕체': 0.6,
  batangche: 0.6,
  '궁서': 0.6,
  gungsuh: 0.6,
};

/** 통합 문서 기본 글꼴(styles.xml 첫 글꼴 = 표준 스타일)의 숫자 한 자 폭(px) — 열 너비 단위의 기준 */
export function maxDigitWidthPx(fontName: string | undefined, sizePt: number | undefined): number {
  const ratio = DIGIT_WIDTH_RATIO[(fontName ?? 'Calibri').trim().toLowerCase()] ?? 0.55;
  const px = ((sizePt && sizePt > 0 ? sizePt : 11) * 96) / 72;
  return Math.max(1, Math.round(px * ratio));
}

/** styles.xml에서 기본 글꼴(첫 <font>)의 이름·크기 */
/** <tag>…</tag> 또는 <tag/> 자식 요소 목록(원문 그대로) */
function childElements(containerXml: string, tag: string): string[] {
  const out: string[] = [];
  const re = new RegExp(`<(?:\\w+:)?${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</(?:\\w+:)?${tag}>)`, 'g');
  for (const m of containerXml.matchAll(re)) out.push(m[0]);
  return out;
}

/**
 * 통합 문서의 "표준" 스타일 글꼴 — 엑셀이 열 너비 단위와 서식 없는 칸의 글꼴로 쓰는 것. 표준 스타일
 * (<cellStyle builtinId="0">)이 가리키는 cellStyleXfs의 fontId 글꼴이다(대개 첫 글꼴이지만 아닐 때도 있다).
 * index는 <fonts> 안 위치, xml은 그 <font> 원문.
 */
export function normalFontOfStylesXml(stylesXml: string | undefined): { name?: string; size?: number; index?: number; xml?: string } {
  if (!stylesXml) return {};
  const fontsXml = /<(?:\w+:)?fonts\b[^>]*>([\s\S]*?)<\/(?:\w+:)?fonts>/.exec(stylesXml)?.[1];
  if (!fontsXml) return {};
  const fonts = childElements(fontsXml, 'font');
  const normalStyle = /<(?:\w+:)?cellStyle\b[^>]*\bbuiltinId="0"[^>]*>/.exec(stylesXml)?.[0];
  const xfId = Number(normalStyle ? /\bxfId="(\d+)"/.exec(normalStyle)?.[1] ?? 0 : 0);
  const styleXfsXml = /<(?:\w+:)?cellStyleXfs\b[^>]*>([\s\S]*?)<\/(?:\w+:)?cellStyleXfs>/.exec(stylesXml)?.[1] ?? '';
  const styleXf = childElements(styleXfsXml, 'xf')[xfId] ?? '';
  const index = Number(/\bfontId="(\d+)"/.exec(styleXf)?.[1] ?? 0);
  const xml = fonts[index] ?? fonts[0];
  if (!xml) return {};
  const size = Number(/<(?:\w+:)?sz\b[^>]*\bval="([^"]+)"/.exec(xml)?.[1]);
  const name = /<(?:\w+:)?name\b[^>]*\bval="([^"]+)"/.exec(xml)?.[1];
  return { name, size: Number.isFinite(size) && size > 0 ? size : undefined, index: fonts[index] ? index : 0, xml };
}

/** styles.xml에서 기본(표준 스타일) 글꼴의 이름·크기 */
export function defaultFontOfStylesXml(stylesXml: string | undefined): { name?: string; size?: number } {
  const { name, size } = normalFontOfStylesXml(stylesXml);
  return { name, size };
}

export function rowHeightPtToPx(pt: number): number {
  return Math.round((pt * 96) / 72);
}

/** 기본과 다른 행 높이를 [시작 행(0-based), 행 수, px] 묶음으로 — 같은 높이가 이어지는 큰 시트도
 * 몇 개 항목으로 끝난다. 숨긴 행은 0px. 시작 행 오름차순, 겹치지 않는다. */
export type RowHeightRuns = ReadonlyArray<readonly [number, number, number]>;

/** (행, px) 목록(행 오름차순)을 묶음으로 압축한다 — 기본 높이와 같은 행은 넣지 않는다. */
export function encodeRowHeights(entries: Iterable<readonly [number, number]>, defaultPx: number): RowHeightRuns {
  const runs: Array<[number, number, number]> = [];
  for (const [row, px] of entries) {
    if (px === defaultPx) continue;
    const last = runs[runs.length - 1];
    if (last && last[0] + last[1] === row && last[2] === px) last[1] += 1;
    else runs.push([row, 1, px]);
  }
  return runs;
}

/** 행 삽입/삭제에 맞춰 높이 묶음을 옮긴다 — 끼워 넣은 행은 기본 높이, 지운 행의 높이는 사라진다. */
export function shiftRowHeightRuns(runs: RowHeightRuns, at: number, count: number, insert: boolean): RowHeightRuns {
  const out: Array<[number, number, number]> = [];
  const push = (start: number, len: number, px: number) => {
    if (len <= 0) return;
    const last = out[out.length - 1];
    if (last && last[0] + last[1] === start && last[2] === px) last[1] += len;
    else out.push([start, len, px]);
  };
  for (const [start, len, px] of runs) {
    const end = start + len;
    if (insert) {
      if (end <= at) push(start, len, px);
      else if (start >= at) push(start + count, len, px);
      else {
        push(start, at - start, px);
        push(at + count, end - at, px);
      }
    } else {
      const delEnd = at + count;
      if (end <= at) push(start, len, px);
      else if (start >= delEnd) push(start - count, len, px);
      else {
        push(start, Math.max(0, at - start), px);
        if (end > delEnd) push(at, end - delEnd, px);
      }
    }
  }
  return out;
}

/**
 * 행마다 높이가 다른(숨긴 행은 0) 시트의 세로 위치 계산 — top(행)과 rowAt(y)를 이분 탐색으로 답한다.
 * 기본 높이가 아닌 구간만 segment로 들고 있어, 높이 지정이 없는 시트는 예전 고정 높이 계산과 같다.
 */
export class RowLayout {
  private readonly starts: number[] = [];
  private readonly ends: number[] = [];
  private readonly heights: number[] = [];
  private readonly tops: number[] = [];

  constructor(
    readonly defaultPx: number,
    runs: RowHeightRuns = [],
    hiddenSorted: readonly number[] = [],
  ) {
    const segs: Array<[number, number, number]> = [];
    const add = (s: number, e: number, h: number) => {
      if (e <= s) return;
      const last = segs[segs.length - 1];
      if (last && last[1] === s && last[2] === h) last[1] = e;
      else segs.push([s, e, h]);
    };
    let hi = 0;
    for (const [start, len, px] of runs) {
      const end = start + len;
      // 이 묶음 앞의 숨긴 행
      while (hi < hiddenSorted.length && hiddenSorted[hi] < start) {
        add(hiddenSorted[hi], hiddenSorted[hi] + 1, 0);
        hi += 1;
      }
      let cursor = start;
      while (hi < hiddenSorted.length && hiddenSorted[hi] < end) {
        const h = hiddenSorted[hi];
        add(cursor, h, px);
        add(h, h + 1, 0);
        cursor = h + 1;
        hi += 1;
      }
      add(cursor, end, px);
    }
    while (hi < hiddenSorted.length) {
      add(hiddenSorted[hi], hiddenSorted[hi] + 1, 0);
      hi += 1;
    }
    let delta = 0;
    for (const [s, e, h] of segs) {
      this.starts.push(s);
      this.ends.push(e);
      this.heights.push(h);
      this.tops.push(s * defaultPx + delta);
      delta += (e - s) * (h - defaultPx);
    }
  }

  /** start <= row인 마지막 segment 번호(없으면 -1) */
  private segBefore(row: number): number {
    let lo = 0;
    let hi = this.starts.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.starts[mid] <= row) lo = mid + 1;
      else hi = mid;
    }
    return lo - 1;
  }

  height(row: number): number {
    const i = this.segBefore(row);
    return i >= 0 && row < this.ends[i] ? this.heights[i] : this.defaultPx;
  }

  /** 0행 맨 위를 0으로 둔 row행의 위쪽 y(px) */
  top(row: number): number {
    const i = this.segBefore(row);
    if (i < 0) return row * this.defaultPx;
    const s = this.starts[i];
    const e = this.ends[i];
    const h = this.heights[i];
    if (row < e) return this.tops[i] + (row - s) * h;
    return this.tops[i] + (e - s) * h + (row - e) * this.defaultPx;
  }

  /** y(px)에 걸친 행 — 숨긴 행은 건너뛴다 */
  rowAt(y: number): number {
    if (y <= 0) return this.firstVisibleFrom(0);
    let lo = 0;
    let hi = this.tops.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.tops[mid] <= y) lo = mid + 1;
      else hi = mid;
    }
    const i = lo - 1;
    if (i < 0) return Math.floor(y / this.defaultPx);
    const s = this.starts[i];
    const e = this.ends[i];
    const h = this.heights[i];
    const segEndY = this.tops[i] + (e - s) * h;
    if (h > 0 && y < segEndY) return s + Math.floor((y - this.tops[i]) / h);
    return this.firstVisibleFrom(e + Math.floor((y - segEndY) / this.defaultPx));
  }

  private firstVisibleFrom(row: number): number {
    let r = row;
    while (this.height(r) === 0) r += 1;
    return r;
  }
}

/** 줄 바꿈 칸의 글자를 그 폭에 넣었을 때 몇 줄이 되는지 어림한다(한글·한자는 글자 크기만큼, 영문·숫자는
 * 그 절반 남짓). 행 높이가 파일에 없을 때 엑셀이 "자동 맞춤"으로 늘려 보여 주는 높이를 흉내 내는 용도. */
export function estimateWrappedLines(text: string, widthPx: number, fontPx: number): number {
  const avail = Math.max(widthPx - 6, fontPx);
  let lines = 0;
  for (const paragraph of text.split('\n')) {
    let w = 0;
    for (const ch of paragraph) {
      const code = ch.codePointAt(0) ?? 0;
      if (code >= 0x1100 && (code <= 0x11ff || (code >= 0x2e80 && code <= 0xa4cf) || (code >= 0xac00 && code <= 0xd7a3) || (code >= 0xf900 && code <= 0xfaff) || (code >= 0xff00 && code <= 0xffef))) w += fontPx;
      else if (ch === ' ') w += fontPx * 0.3;
      else if (/[A-Z0-9%@#&MW]/.test(ch)) w += fontPx * 0.62;
      else w += fontPx * 0.52;
    }
    lines += Math.max(1, Math.ceil(w / avail));
  }
  return lines;
}

/** 기본 글꼴 크기(11pt) px */
export const EXCEL_DEFAULT_FONT_PX = (11 * 96) / 72;

/** 자동 맞춤 행 높이(px) — 한 줄 높이는 기본 행 높이를 글자 크기 비율로 늘린 값(엑셀이 11pt 한 줄에
 * 기본 행 높이를 쓰는 것과 같은 비율). */
export function autoFitRowHeightPx(lines: number, fontPx: number, defaultRowPx: number): number {
  return Math.ceil(lines * defaultRowPx * Math.max(1, fontPx / EXCEL_DEFAULT_FONT_PX));
}
