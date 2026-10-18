import FormulaParser from 'fast-formula-parser';
import { dateTextToSerial } from '../collab/cellValueShared.js';

/**
 * fast-formula-parser의 "값 → 숫자" 변환(FormulaHelpers.acceptNumber)에 엑셀의 두 규칙을 더한다 —
 * 이 모듈을 import하는 것만으로 한 번 적용된다(부수효과 모듈, formulaParserPool.ts가 import).
 *
 * 1) **빈 칸은 산술에서 0**: 이 앱은 빈 칸을 ''로 넘긴다(그래야 =A1="", LEN, IF(A1="",…) 같은
 *    문자열·비교 동작이 엑셀과 같다 — null로 넘기면 A1=""가 FALSE, LEN이 15가 되는 등 더 틀린다,
 *    실측). 그런데 라이브러리는 ''를 산술에 쓰면 #VALUE!를 내서, 월별 칸 중 하나만 비어 있어도
 *    =B3+C3+D3 같은 흔한 수식이 통째로 #VALUE!가 됐다. 엑셀은 빈 칸을 산술에서 0으로 본다.
 *    (빈 문자열을 "만드는" 수식 결과 ""까지 0으로 보는 점은 엑셀(#VALUE!)보다 관대하지만, 빈 칸과
 *    구분할 표현이 이 앱에 없고 오류보다 계산되는 쪽이 낫다.)
 * 2) **날짜 모양 텍스트는 산술에서 날짜 일련번호**: 날짜 칸이 "2026-07-05" 텍스트라 =B4-A4, =A4+7이
 *    #VALUE!였다. 엑셀은 날짜 모양 텍스트를 산술에서 자동으로 날짜로 바꾼다(="2026-07-15"-"2026-07-05"
 *    → 10). 날짜 "함수"(DAYS/DATEDIF/YEAR)는 라이브러리가 이미 처리하므로 막힌 건 이 변환뿐이었다.
 *
 * 연산자 모듈(formulas/operators.js)은 호출 시점마다 `FormulaHelpers.acceptNumber`를 찾아 쓰고, 그
 * 객체는 index.js가 FormulaParser에 그대로 붙여 내보내는 싱글톤이라 여기서 바꾸면 모든 파서
 * 인스턴스(풀 포함)에 적용된다. 원래 숫자로 읽히던 입력은 원래 함수가 그대로 처리하고, 원래 #VALUE!가
 * 나던 두 경우만 새로 숫자가 된다 — 기존에 맞던 결과는 바뀌지 않는다.
 */
type AcceptNumber = (obj: unknown, isArray?: boolean, allowBoolean?: boolean) => unknown;

const helpers = (FormulaParser as unknown as { FormulaHelpers: { acceptNumber: AcceptNumber; __excelCoercion?: true } })
  .FormulaHelpers;

if (!helpers.__excelCoercion) {
  const original = helpers.acceptNumber;
  helpers.acceptNumber = function acceptNumberExcelLike(this: unknown, obj, isArray, allowBoolean) {
    if (obj === '') return 0;
    if (typeof obj === 'string' && Number.isNaN(Number(obj))) {
      const serial = dateTextToSerial(obj);
      if (serial !== null) return serial;
    }
    return original.call(this, obj, isArray, allowBoolean);
  };
  helpers.__excelCoercion = true;
}
