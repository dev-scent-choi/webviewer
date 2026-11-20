import { describe, expect, test } from 'vitest';
import { FILL_MARK, formatNumberCode } from '../src/collab/numberFormat.js';
import excelBuiltinNumFmts from 'exceljs/lib/xlsx/defaultnumformats.js';
import '../src/spreadsheet.js'; // 한국어 엑셀 기본 서식 번호 표를 채운다

const D = 46346.6; // 2026-11-20(금) 14:24:00
const N = 1234.5;
const shown = (value: number, code: string) => formatNumberCode(value, code).text.replace(FILL_MARK, '').trim();

/** 한국어 엑셀(ko-KR)이 보여 주는 글자 — 실제 문서에서 흔한 서식 코드 표 */
const CASES: ReadonlyArray<readonly [string, number, string]> = [
  ['General', N, '1234.5'],
  ['0', N, '1235'],
  ['0.00', N, '1234.50'],
  ['#,##0', N, '1,235'],
  ['#,##0.00', -N, '-1,234.50'],
  ['0%', 0.256, '26%'],
  ['0.00%', 0.256, '25.60%'],
  ['0.00E+00', N, '1.23E+03'],
  ['##0.0E+0', N, '1.2E+3'],
  ['yyyy-mm-dd', D, '2026-11-20'],
  ['d-mmm-yy', D, '20-Nov-26'],
  ['mmmm d, yyyy', D, 'November 20, 2026'],
  ['m/d', D, '11/20'],
  ['yyyy/mm/dd', D, '2026/11/20'],
  ['h:mm AM/PM', D, '2:24 PM'],
  ['h:mm:ss', D, '14:24:00'],
  ['mm:ss', D, '24:00'],
  ['[h]:mm:ss', 1.5, '36:00:00'],
  ['hh:mm:ss', 0.9999999, '00:00:00'], // 자정 직전은 엑셀처럼 올림
  ['yyyy-mm-dd h:mm', D, '2026-11-20 14:24'],
  // 한국 문서에서 흔한 사용자 서식(SSF가 그대로는 못 읽던 것)
  ['yyyy.mm.dd', D, '2026.11.20'],
  ['yyyy\\.mm\\.dd', D, '2026.11.20'],
  ['yy.mm.dd', D, '26.11.20'],
  ['yyyy. m. d', D, '2026. 11. 20'],
  ['mm.dd', D, '11.20'],
  ['yyyy.mm.dd hh:mm', D, '2026.11.20 14:24'],
  ['yyyy년 mm월 dd일', D, '2026년 11월 20일'],
  ['yyyy"년" m"월" d"일"', D, '2026년 11월 20일'],
  ['h시 mm분', D, '14시 24분'],
  ['[$-412]yyyy년 m월 d일 dddd', D, '2026년 11월 20일 금요일'],
  ['yyyy-mm-dd (aaa)', D, '2026-11-20 (금)'],
  ['aaaa', D, '금요일'],
  ['[$-ko-KR]dddd', D, '금요일'],
  ['[$-412]AM/PM h:mm:ss', 0.6, '오후 2:24:00'],
  ['[$-412]AM/PM hh:mm', 0.75, '오후 06:00'],
  ['[$-412]AM/PM h:mm', 0.25, '오전 6:00'],
  ['[$-412]General', 5, '5'],
  ['[$-412]0', 5, '5'],
  ['[DBNum1]#,##0', N, '1,235'],
  ['#,##0원', 1500, '1,500원'],
  ['#,##0 원', 1500, '1,500 원'],
  ['0개', 1500, '1500개'],
  ['"일금 "#,##0"원정"', 9108000, '일금 9,108,000원정'],
  ['#,##0,"천원"', 12345678, '12,346천원'],
  ['#,##0.0,,"백만"', 123456789, '123.5백만'],
  ['[>=1000]"많음 "#,##0;"적음 "#,##0', 999, '적음 999'],
  ['"₩"#,##0;\\-"₩"#,##0', -N, '-₩1,235'],
  ['[$₩-412]#,##0', N, '₩1,235'],
  ['[$$-409]#,##0.00', N, '$1,234.50'],
  ['_-"₩"* #,##0_-;\\-"₩"* #,##0_-;_-"₩"* "-"_-;_-@_-', -25000, '-₩25,000'],
  ['_-"₩"* #,##0_-;\\-"₩"* #,##0_-;_-"₩"* "-"_-;_-@_-', 0, '₩-'],
  ['000-0000', 1234567, '123-4567'],
];

describe('엑셀 서식 코드 표(보기·편집 공용 서식기)', () => {
  test.each(CASES)('%s · %s → %s', (code, value, expected) => {
    expect(shown(value, code)).toBe(expected);
  });

  test('회계 채움(*): ₩와 숫자 사이에 채움 표시가 들어간다(칸 양 끝 배치용)', () => {
    expect(formatNumberCode(-25000, '_-"₩"* #,##0_-;\\-"₩"* #,##0_-;_-"₩"* "-"_-;_-@_-').text).toContain(`₩${FILL_MARK}25,000`);
  });

  test('[ColorN]은 엑셀 56색 표', () => {
    expect(formatNumberCode(5, '[Color10]0').color).toBe('#008000');
    expect(formatNumberCode(5, '[Color3]0').color).toBe('#FF0000');
  });

  test('못 읽는 서식도 일련번호를 그대로 보이지 않는다(날짜는 날짜 글자, 숫자는 쉼표)', () => {
    expect(shown(D, 'yyyy-mm-dd [[[')).not.toBe(String(D));
    expect(shown(1500, '#,##0 [[[')).toBe('1,500');
  });

  test('한국어 엑셀 기본 서식 번호(ExcelJS 표를 한국어 코드로 채움)', () => {
    const f = (id: number) => (excelBuiltinNumFmts as Record<string, { f?: string }>)[id]?.f ?? '';
    expect(shown(D, f(14))).toBe('2026-11-20');
    expect(shown(D, f(31))).toBe('2026년 11월 20일');
    expect(shown(D, f(32))).toBe('14시 24분');
    expect(shown(-N, f(5))).toBe('-₩1,235');
    expect(shown(-N, f(38))).toBe('-1,235');
    expect(shown(-25000, f(42))).toBe('-₩25,000');
    expect(shown(D, f(22))).toBe('2026-11-20 14:24');
    for (let id = 0; id <= 58; id += 1) {
      const code = f(id);
      if (!code) continue;
      const out = shown(id === 49 ? 0 : D, code);
      expect(out, `기본 서식 ${id} ${code}`).not.toMatch(/\$|\[|undefined/);
    }
  });
});
