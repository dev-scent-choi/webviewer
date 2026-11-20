/**
 * 엑셀 표시 형식(numFmt) 서식기 — 보기 화면(서버가 그리드 파일을 만들 때)과 편집 화면(브라우저가 칸을 그릴 때)이
 * 반드시 같은 결과를 내도록 둘 다 이 함수만 쓴다. 실제 해석은 검증된 SheetJS SSF(fast-formula-parser에 들어 있는
 * 사본 — TEXT 함수와 같은 서식기)가 하고, 여기서는 SSF가 모르는 한국어 요일(aaa/aaaa)·[$-412] 로캘과 글자 색
 * ([Red] 등)만 보탠다.
 */
import SSF from 'fast-formula-parser/ssf/ssf.js';
import { dateTextToSerial } from './cellValueShared.js';

export interface FormattedValue {
  readonly text: string;
  /** 서식 코드의 색 지정([Red] 등) — '#RRGGBB' */
  readonly color?: string;
}

/** 서식을 고를 때 보는 칸 서식 — gridShared.ts CellStyleAttrs·spreadsheet.ts GridStyle의 공통 부분 */
export interface NumFmtStyle {
  /** 원본 서식 코드(엑셀 그대로) — 있으면 이것이 이긴다 */
  readonly fmt?: string;
  readonly nf?: 'number' | 'currency' | 'percent';
  readonly nd?: number;
  readonly df?: 1;
}

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const KO_DAYS = ['일', '월', '화', '수', '목', '금', '토'];
const NAMED_COLORS: Record<string, string> = {
  black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00',
  magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00',
};
/** 엑셀 [ColorN] 처음 8개(나머지는 드물어 검정으로 둔다) */
const INDEXED_COLORS = ['#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF'];

/** 툴바 범주(nf/nd)·날짜 표시(df)를 엑셀 서식 코드로 — 원본 코드(fmt)가 없을 때 */
export function numFmtCodeOf(style: NumFmtStyle | undefined): string | undefined {
  if (!style) return undefined;
  if (style.fmt) return style.fmt;
  const decimals = style.nd ? `.${'0'.repeat(style.nd)}` : '';
  if (style.nf === 'number') return `#,##0${decimals}`;
  if (style.nf === 'currency') return `"₩"#,##0${decimals}`;
  if (style.nf === 'percent') return `0${decimals}%`;
  return undefined;
}

/** 텍스트 칸 서식(@)인지 — 숫자처럼 보여도 글자로 다룬다 */
export function isTextNumFmt(code: string | undefined): boolean {
  return code === '@';
}

/** 따옴표·이스케이프·대괄호 밖에서만 바꾼다 */
function replaceOutsideLiterals(code: string, pattern: RegExp, fn: (m: string) => string): string {
  return code.replace(new RegExp(`"[^"]*"|\\\\.|\\[[^\\]]*\\]|${pattern.source}`, 'gi'), (m) =>
    m.startsWith('"') || m.startsWith('\\') || m.startsWith('[') ? m : fn(m),
  );
}

/** SSF가 모르는 한국어 표기를 글자 그대로("…")로 바꾼다 — aaa/aaaa(요일), [$-412]의 ddd/dddd·AM/PM */
function koreanize(code: string, serial: number): string {
  const korean = /\[\$-(?:412|ko)/i.test(code);
  if (!korean && !/a{3}/i.test(code)) return code;
  const day = KO_DAYS[new Date(EXCEL_EPOCH_MS + Math.floor(serial) * 86_400_000).getUTCDay()];
  let out = replaceOutsideLiterals(code, /a{3,4}/, (m) => (m.length === 4 ? `"${day}요일"` : `"${day}"`));
  if (korean) {
    out = replaceOutsideLiterals(out, /d{3,4}/, (m) => (m.length === 4 ? `"${day}요일"` : `"${day}"`));
    const pm = (serial - Math.floor(serial)) * 24 >= 12;
    out = replaceOutsideLiterals(out, /AM\/PM/, () => (pm ? '"오후"' : '"오전"'));
  }
  return out;
}

/** 서식 코드별 준비 결과 — 칸마다 다시 하지 않는다(큰 시트는 같은 코드가 수십만 번 온다) */
interface PreparedCode {
  readonly korean: boolean;
  /** 색 지정([Red] 등)이 있는 코드만 구역 목록을 둔다 */
  readonly sections?: readonly string[];
}
const prepared = new Map<string, PreparedCode>();
function prepare(code: string): PreparedCode {
  let p = prepared.get(code);
  if (!p) {
    const colored = /\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)\]/i.test(code);
    p = {
      korean: /a{3}/i.test(code) || /\[\$-(?:412|ko)/i.test(code),
      ...(colored ? { sections: SSF._split(code).filter((_s, i, all) => !(i === all.length - 1 && all.length === 4)) } : {}),
    };
    if (prepared.size > 500) prepared.clear(); // 서식 코드 종류는 파일마다 수십 개 — 넘치면 비운다
    prepared.set(code, p);
  }
  return p;
}

/** 값이 고를 서식 구역(양수;음수;0;글자, [조건]) — 색을 알아내는 데 쓴다(글자 자체는 SSF가 같은 규칙으로 고른다) */
function sectionFor(sections: readonly string[], value: number): string {
  const cond = (s: string) => /\[(=|<>|>=|<=|>|<)\s*(-?[\d.]+)\]/.exec(s);
  if (sections.some((s) => cond(s))) {
    for (const s of sections) {
      const m = cond(s);
      if (!m) return s;
      const n = Number(m[2]);
      const ok = m[1] === '=' ? value === n : m[1] === '<>' ? value !== n : m[1] === '>=' ? value >= n : m[1] === '<=' ? value <= n : m[1] === '>' ? value > n : value < n;
      if (ok) return s;
    }
    return sections[sections.length - 1];
  }
  if (sections.length >= 3 && value === 0) return sections[2];
  if (sections.length >= 2 && value < 0) return sections[1];
  return sections[0];
}

function colorOf(section: string): string | undefined {
  const m = /\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*(\d+))\]/i.exec(section);
  if (!m) return undefined;
  if (m[2]) return INDEXED_COLORS[Number(m[2]) - 1] ?? '#000000';
  return NAMED_COLORS[m[1].toLowerCase()];
}

/** 숫자(또는 날짜 일련번호) 하나를 엑셀 서식 코드대로 */
export function formatNumberCode(value: number, code: string | undefined): FormattedValue {
  if (!Number.isFinite(value)) return { text: String(value) };
  const fmt = code && code !== 'General' ? code : 'General';
  // 가장 흔한 경우(일반 서식의 11자리 이하 정수)는 SSF를 거치지 않는다 — 결과는 같다
  if (fmt === 'General' && Number.isInteger(value) && Math.abs(value) < 1e11) return { text: String(value) };
  try {
    const p = prepare(fmt);
    const text = SSF.format(p.korean ? koreanize(fmt, value) : fmt, value);
    const color = p.sections ? colorOf(sectionFor(p.sections, value)) : undefined;
    return color ? { text, color } : { text };
  } catch {
    // 서식 코드가 깨졌으면 엑셀 일반 서식으로
    return { text: SSF.format('General', value) };
  }
}

const NUMERIC_TEXT = /^-?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$/;

/**
 * 칸의 값 글자(편집기 cells 맵·그리드의 값 — 숫자는 "1234.5", 날짜는 "2026-11-20 14:05:09")를 칸 서식대로 보여 줄 글자로.
 * 서식이 없거나 글자 서식(@)이면, 또는 값이 숫자·날짜가 아니면 그대로 돌려준다.
 */
export function formatValueText(raw: string, style: NumFmtStyle | undefined): FormattedValue {
  const code = numFmtCodeOf(style);
  const t = raw.trim();
  if (t === '' || (!code && !style?.df) || isTextNumFmt(code)) return { text: raw };
  if (NUMERIC_TEXT.test(t)) return formatNumberCode(Number(t), code ?? 'yyyy-mm-dd');
  const serial = dateTextToSerial(t);
  if (serial !== null) {
    // 날짜 값인데 서식 코드가 숫자 서식(툴바로 바꾼 쉼표 등)이면 엑셀처럼 일련번호를 그 서식으로
    const dateCode = code ?? (serial % 1 === 0 ? 'yyyy-mm-dd' : 'yyyy-mm-dd hh:mm');
    return formatNumberCode(serial, dateCode);
  }
  return { text: raw };
}

/** 서식이 걸린 숫자·날짜 글자가 칸 폭에 안 들어가면 엑셀처럼 '#'으로 채운다(일반 서식·글자는 엑셀도 그대로 둔다).
 * 너비는 보수적인 어림(숫자 0.45em, 쉼표·마침표·하이픈 등 0.27em, 한글 등 1em, 그 밖 0.55em, 여백 2px) — 실제보다 좁게 잡아
 * 들어가는 값을 #으로 가리는 일이 없게 한다. 보기·편집 화면이 같은 어림을 쓴다. */
export function hashIfTooWide(text: string, widthPx: number, fontPx: number): string {
  let width = 2;
  for (const ch of text) {
    const code = ch.codePointAt(0) ?? 0;
    width += fontPx * (code >= 0x2e80 ? 1 : /[0-9]/.test(ch) ? 0.45 : /[-,.:/ ()]/.test(ch) ? 0.27 : 0.55);
  }
  if (width <= widthPx || widthPx <= 0) return text;
  return '#'.repeat(Math.max(1, Math.floor((widthPx - 2) / (fontPx * 0.45))));
}

/** 숫자·날짜 값인지(서식이 걸리면 '#' 채움·오른쪽 정렬 대상) */
export function isNumberOrDateValue(raw: string): boolean {
  const t = raw.trim();
  return NUMERIC_TEXT.test(t) || dateTextToSerial(t) !== null;
}
