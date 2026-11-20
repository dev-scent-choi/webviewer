/**
 * 엑셀 표시 형식(numFmt) 서식기 — 보기 화면(서버가 그리드 파일을 만들 때)과 편집 화면(브라우저가 칸을 그릴 때)이
 * 반드시 같은 결과를 내도록 둘 다 이 함수만 쓴다. 실제 해석은 검증된 SheetJS SSF(fast-formula-parser에 들어 있는
 * 사본 — TEXT 함수와 같은 서식기)가 하고, 여기서는 SSF가 모르는 것(한국어 요일 aaa/aaaa·[$-412] 로캘·오전/오후,
 * 따옴표 없는 한글, 날짜의 점, 채움 *, 글자 색 [Red]·[ColorN])을 미리 고치거나 보탠다.
 */
import SSF from 'fast-formula-parser/ssf/ssf.js';
import { dateTextToSerial, serialToDateText } from './cellValueShared.js';

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

/** 엑셀 색 번호 표(indexed 0~63) — 칸 색의 indexed와 서식 코드 [ColorN](= 8+N-1번)이 함께 쓴다 */
export const EXCEL_INDEXED_COLORS: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00',
  '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF',
  '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300',
  '993300', '993366', '333399', '333333',
];

/** 채움 표시(* 다음 글자로 칸을 채우는 자리) — 이 글자 앞은 칸 왼쪽 끝, 뒤는 오른쪽 끝에 붙는다(회계 서식의 ₩ … 25,000) */
export const FILL_MARK = '\u0001';

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

/** 서식 코드를 글자 그대로 부분("…"·\x)·대괄호([…])·나머지로 나눠 나머지 부분에만 fn을 적용한다 */
function mapOutsideLiterals(code: string, fn: (plain: string) => string): string {
  let out = '';
  let plain = '';
  for (let i = 0; i < code.length; i += 1) {
    const ch = code[i];
    const end = ch === '"' ? code.indexOf('"', i + 1) : ch === '[' ? code.indexOf(']', i + 1) : ch === '\\' ? i + 1 : -2;
    if (end === -2) {
      plain += ch;
      continue;
    }
    out += fn(plain);
    plain = '';
    const stop = end < 0 ? code.length - 1 : end;
    out += code.slice(i, stop + 1);
    i = stop;
  }
  return out + fn(plain);
}

/** 날짜 구역인지(y·m·d·h·s가 있고 숫자 자리 0·#·?가 없다) */
const isDateSection = (section: string) => {
  const plain = mapOutsideLiterals(section, (p) => p).replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  return /[ymdhs]/i.test(plain) && !/[0#?]/.test(plain.replace(/s+\.0+/gi, ''));
};

/**
 * SSF가 못 읽는 것을 미리 고친다(코드마다 한 번):
 * - 통화 기호 없는 로캘·숫자 표기 지정([$-412], [$-ko-KR], [DBNum1] …) — 글자로 새어 나오지 않게 지운다.
 * - 따옴표 없는 한글 등 비ASCII 글자(년·월·원·개 …)는 "…"로 감싼다.
 * - 날짜 구역의 점(yyyy.mm.dd, yyyy\.mm\.dd)은 "."로 — SSF가 소수점으로 읽다 실패한다(초의 소수 ss.0은 둔다).
 * - 채움(* 다음 글자)은 채움 표시 글자로 — SSF는 채움을 버리므로 우리가 칸 양 끝 배치로 그린다.
 */
function normalizeForSsf(code: string): string {
  const withoutLocale = code.replace(/\[\$-[^\]]*\]|\[(DBNum|NatNum)\d+\]/gi, '');
  return withoutLocale
    .split(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/)
    .map((section) => {
      const date = isDateSection(section);
      let s = date ? section.replace(/\\\.(?!0)/g, '"."') : section;
      s = mapOutsideLiterals(s, (plain) => {
        let p = plain.replace(/[^\x00-\x7f]+/g, (m) => `"${m}"`);
        if (date) p = p.replace(/(?<!s)\.(?!0)/gi, '"."');
        return p.replace(/\*(.)/g, `"${FILL_MARK}"`);
      });
      return s;
    })
    .join(';');
}

/** 서식 코드별 준비 결과 — 칸마다 다시 하지 않는다(큰 시트는 같은 코드가 수십만 번 온다) */
interface PreparedCode {
  readonly ssf: string;
  /** 값마다 요일 글자를 넣어야 하는 코드(aaa, [$-412]의 ddd) */
  readonly days: boolean;
  /** 한국어 로캘 — ddd/dddd도 한국어 요일로 */
  readonly koDdd: boolean;
  /** 한국어 로캘의 오전/오후 — SSF가 그린 AM/PM을 바꾼다(12시간제는 SSF가 그대로 지킨다) */
  readonly koAmPm: boolean;
  readonly date: boolean;
  /** 숫자 서식이 SSF로 안 될 때 쓰는 소수 자릿수 */
  readonly decimals: number;
  /** 색 지정([Red] 등)이 있는 코드만 구역 목록을 둔다 */
  readonly sections?: readonly string[];
}
const prepared = new Map<string, PreparedCode>();
function prepare(code: string): PreparedCode {
  let p = prepared.get(code);
  if (!p) {
    const korean = /\[\$-(?:412|ko)/i.test(code);
    const plain = mapOutsideLiterals(code, (x) => x).replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
    const colored = /\[(black|blue|cyan|green|magenta|red|white|yellow|color\s*\d+)\]/i.test(code);
    let sections: string[] | undefined;
    try {
      sections = colored ? SSF._split(code).filter((_s, i, all) => !(i === all.length - 1 && all.length === 4)) : undefined;
    } catch {
      sections = undefined;
    }
    p = {
      ssf: normalizeForSsf(code),
      days: /a{3}/i.test(plain) || (korean && /d{3}/i.test(plain)),
      koDdd: korean,
      koAmPm: korean && /AM\/PM/i.test(plain),
      date: code.split(';').some(isDateSection),
      decimals: /\.(0+)/.exec(plain)?.[1].length ?? 0,
      ...(sections ? { sections } : {}),
    };
    if (prepared.size > 500) prepared.clear(); // 서식 코드 종류는 파일마다 수십 개 — 넘치면 비운다
    prepared.set(code, p);
  }
  return p;
}

/** 값마다 다른 요일 글자를 넣는다(aaa/aaaa, 한국어 로캘의 ddd/dddd) */
function withDays(ssf: string, serial: number, koreanDdd: boolean): string {
  const day = KO_DAYS[new Date(EXCEL_EPOCH_MS + Math.floor(serial) * 86_400_000).getUTCDay()];
  return mapOutsideLiterals(ssf, (plain) =>
    plain.replace(koreanDdd ? /a{3,4}|d{3,4}/gi : /a{3,4}/gi, (m) => (m.length === 4 ? `"${day}요일"` : `"${day}"`)),
  );
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
  if (m[2]) {
    const hex = EXCEL_INDEXED_COLORS[7 + Number(m[2])];
    return hex ? `#${hex}` : undefined;
  }
  return NAMED_COLORS[m[1].toLowerCase()];
}

/** 숫자(또는 날짜 일련번호) 하나를 엑셀 서식 코드대로. 채움(*)이 있으면 text에 FILL_MARK가 든다 */
export function formatNumberCode(value: number, code: string | undefined): FormattedValue {
  if (!Number.isFinite(value)) return { text: String(value) };
  const fmt = code && code !== 'General' ? code : 'General';
  // 가장 흔한 경우(일반 서식의 11자리 이하 정수)는 SSF를 거치지 않는다 — 결과는 같다
  if (fmt === 'General' && Number.isInteger(value) && Math.abs(value) < 1e11) return { text: String(value) };
  const p = prepare(fmt);
  // 날짜·시각은 초 단위로 반올림해서 그린다(23:59:59.9999가 23:59:59로 깎이지 않게 — 엑셀은 반올림)
  const v = p.date ? Math.round(value * 86_400) / 86_400 : value;
  try {
    let text = SSF.format(p.days ? withDays(p.ssf, v, p.koDdd) : p.ssf, v);
    if (p.koAmPm) text = text.replace(/AM|PM/, (m) => (m === 'AM' ? '오전' : '오후'));
    const color = p.sections ? colorOf(sectionFor(p.sections, value)) : undefined;
    return color ? { text, color } : { text };
  } catch {
    // SSF가 끝내 못 읽는 코드 — 일련번호를 그대로 보이지 않고 종류대로(날짜는 날짜 글자, 숫자는 천 단위 쉼표)
    if (p.date) return { text: serialToDateText(value) };
    if (!/[0#]/.test(fmt)) return { text: SSF.format('General', value) };
    return { text: SSF.format(`#,##0${p.decimals ? `.${'0'.repeat(p.decimals)}` : ''}`, value) };
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
