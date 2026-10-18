/**
 * 엑셀의 "값 해석" 규칙 중 이 서비스가 흉내 내야 하는 부분 — 서버(수식 계산·xlsx 내보내기)와
 * 브라우저(셀 입력·표시)가 반드시 같은 규칙을 써야 해서 DOM/Node 의존성 없이 한 곳에 둔다
 * (gridShared.ts / pptxShared.ts와 같은 공유 모듈 관례).
 *
 * 배경: 이 앱은 날짜 칸을 엑셀의 날짜 일련번호가 아니라 "YYYY-MM-DD"(시각이 있으면
 * "YYYY-MM-DD HH:MM") 텍스트로 들고 다닌다(spreadsheet.ts의 formatDate가 유일한 날짜 표시 경로).
 * 엑셀은 이런 날짜 모양 텍스트를 산술에 쓰면 자동으로 날짜로 바꿔 계산하고(="2026-07-15"-"2026-07-05"
 * → 10), 사용자가 셀에 "1,000"/"50%"/"₩1,000"을 치면 텍스트가 아니라 숫자(+해당 표시 형식)로
 * 저장한다. 이 모듈이 그 두 규칙을 제공한다.
 */

/** formatDate(spreadsheet.ts)가 만드는 모양 그대로 — 이 앱 안의 날짜 텍스트는 항상 이 형태다. */
const DATE_TEXT = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{2}):(\d{2}))?$/;

/** 엑셀 날짜 일련번호의 기준(1900 날짜 체계, 1900-02-29 버그 보정이 이미 반영된 1899-12-30). */
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

/** "2026-07-05"(또는 "2026-07-05 13:30") → UTC Date. 날짜 모양이 아니거나 존재하지 않는 날짜면 null. */
export function dateTextToDate(text: string): Date | null {
  const m = DATE_TEXT.exec(text.trim());
  if (!m) return null;
  const [y, mo, d, hh = '0', mm = '0'] = m.slice(1).map((v) => v ?? '0');
  const date = new Date(Date.UTC(Number(y), Number(mo) - 1, Number(d), Number(hh), Number(mm)));
  // 2026-02-31 같은 존재하지 않는 날짜는 Date가 조용히 다음 달로 넘겨버린다 — 날짜로 취급하지 않는다.
  if (date.getUTCMonth() !== Number(mo) - 1 || date.getUTCDate() !== Number(d)) return null;
  return date;
}

/** 날짜 모양 텍스트 → 엑셀 날짜 일련번호(시각은 소수부). 날짜가 아니면 null. */
export function dateTextToSerial(text: string): number | null {
  const date = dateTextToDate(text);
  return date ? (date.getTime() - EXCEL_EPOCH_MS) / MS_PER_DAY : null;
}

/** 엑셀 날짜 일련번호 → 이 앱의 날짜 표시 텍스트(formatDate와 같은 모양). */
export function serialToDateText(serial: number): string {
  // 분 단위로 반올림 — 부동소수 잡음(…:59.9999초)이 날짜/분을 하나 깎아 먹지 않게
  const ms = Math.round((serial * MS_PER_DAY) / 60_000) * 60_000;
  const date = new Date(EXCEL_EPOCH_MS + ms);
  const y = date.getUTCFullYear();
  const mo = String(date.getUTCMonth() + 1).padStart(2, '0');
  const d = String(date.getUTCDate()).padStart(2, '0');
  const hh = date.getUTCHours();
  const mm = date.getUTCMinutes();
  return hh === 0 && mm === 0
    ? `${y}-${mo}-${d}`
    : `${y}-${mo}-${d} ${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
}

/** 엑셀 표시 형식 코드가 날짜/시각 형식인지 — 따옴표 안 문자열·[색]/[$통화] 대괄호·이스케이프를
 * 걷어낸 뒤 연/월/일/시 토큰(y, d, h, 그리고 시간 문맥이 아닌 m)이 남아 있으면 날짜로 본다.
 * 경과 시간 형식([h]:mm)은 대괄호를 걷어내 판단하지 않는다(날짜가 아니라 기간이다). */
export function isDateNumFmt(numFmt: string | undefined): boolean {
  if (!numFmt || numFmt === 'General' || numFmt === '@') return false;
  const stripped = numFmt
    .replace(/"[^"]*"/g, '')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/\\./g, '')
    .toLowerCase();
  return /[yd]/.test(stripped) || /m/.test(stripped.replace(/h+:?m+|m+:?s+/g, ''));
}

export type TypedNumberKind = 'number' | 'currency' | 'percent';

/**
 * 사용자가 셀에 직접 친 문자열을 엑셀처럼 숫자로 해석한다 — "1,000"(천 단위 쉼표), "50%",
 * "₩1,000"만 대상이다(일반 숫자 "123"은 원래부터 숫자로 저장되므로 여기서 다루지 않는다).
 * 쉼표 위치가 천 단위 규칙에 안 맞는 "1,00" 같은 입력은 엑셀도 텍스트로 두므로 null.
 */
export function parseTypedNumber(text: string): { value: number; kind: TypedNumberKind; decimals: number } | null {
  const t = text.trim();
  const grouped = '-?\\d{1,3}(?:,\\d{3})*(?:\\.(\\d+))?';
  const plainOrGrouped = `(?:${grouped}|-?\\d+(?:\\.(\\d+))?)`;
  let m = new RegExp(`^(${plainOrGrouped})%$`).exec(t);
  if (m) {
    const n = Number(m[1].replace(/,/g, ''));
    if (!Number.isFinite(n)) return null;
    return { value: Math.round((n / 100) * 1e12) / 1e12, kind: 'percent', decimals: (m[2] ?? m[3] ?? '').length };
  }
  m = new RegExp(`^₩\\s?(${plainOrGrouped})$`).exec(t);
  if (m) {
    const n = Number(m[1].replace(/,/g, ''));
    return Number.isFinite(n) ? { value: n, kind: 'currency', decimals: (m[2] ?? m[3] ?? '').length } : null;
  }
  m = new RegExp(`^${grouped}$`).exec(t);
  if (m && t.includes(',')) {
    const n = Number(t.replace(/,/g, ''));
    return Number.isFinite(n) ? { value: n, kind: 'number', decimals: (m[1] ?? '').length } : null;
  }
  return null;
}
