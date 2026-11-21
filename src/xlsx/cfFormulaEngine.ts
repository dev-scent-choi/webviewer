/**
 * 조건부 서식 수식의 대체 계산 — 작은 계산기(condFormatShared.ts evalCfValue)가 못 하는 범위·함수(COUNTIF($A:$A,$A2)>1,
 * SUM(…), VLOOKUP …)를 이 앱의 수식 엔진(fast-formula-parser + formulaOverrides, 서버 수식 계산과 같은 함수 표)으로 계산한다.
 * 서버(보기 화면 조건부 서식 굽기)는 바로 쓰고, 편집 화면은 그런 규칙이 있을 때만 이 모듈을 따로 불러온다(번들 분리).
 * DOM을 쓰지 않는다.
 */
import FormulaParser, { type CellRef, type RangeRef } from 'fast-formula-parser';
import './formulaExcelCoercion.js';
import { formulaFunctions, formulaFunctionsNeedContext } from '../client/formulaOverrides.js';
import { CF_CANNOT, cfNumber, type CfFormulaFallback, type CfVal } from '../collab/condFormatShared.js';

const SHEET = 'S';
const COLUMN_RANGE = /(?<![A-Za-z0-9_$!.])(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(])/g;

/** 칸 값 글자 → 수식 엔진 값(빈 칸 '', 수·날짜는 수 — 날짜 칸을 TODAY()와 비교할 수 있게, 논리값, 나머지 글자) */
function engineValue(text: string): unknown {
  if (text.trim() === '') return '';
  if (text === 'TRUE' || text === 'FALSE') return text === 'TRUE';
  return cfNumber(text) ?? text;
}

/**
 * valueAt: 이 시트의 칸 값 글자(0부터 센 행·열), extent: 시트 크기(A:A 같은 시트 끝까지 범위를 여기까지만 읽는다).
 * 다른 시트 참조는 빈 칸으로 본다(조건부 서식은 엑셀도 같은 시트 위주다).
 */
export function createCfFormulaFallback(
  valueAt: (row: number, col: number) => string,
  extent: () => { rows: number; cols: number },
): CfFormulaFallback {
  const parser = new FormulaParser({
    functions: formulaFunctions,
    functionsNeedContext: formulaFunctionsNeedContext,
    onCell: (ref: CellRef) => (ref.sheet === SHEET ? engineValue(valueAt(ref.row - 1, ref.col - 1)) : ''),
    onRange: (ref: RangeRef) => {
      const { rows, cols } = extent();
      const r1 = Math.min(ref.to.row, rows);
      const c1 = Math.min(ref.to.col, cols);
      const out: unknown[][] = [];
      for (let r = ref.from.row; r <= r1; r += 1) {
        const line: unknown[] = [];
        for (let c = ref.from.col; c <= c1; c += 1) line.push(ref.sheet === SHEET ? engineValue(valueAt(r - 1, c - 1)) : '');
        out.push(line);
      }
      return out.length > 0 ? out : [['']];
    },
  });
  return (formula, row, col) => {
    // 이 엔진은 SUMPRODUCT 안의 배열 비교((범위=값)·(범위>값))를 거짓으로 잘못 계산한다 — 계산하지 못함으로 알린다
    // (자주 쓰는 SUMPRODUCT((범위=값)*1)은 condFormatShared.ts rewriteSumproduct가 미리 COUNTIF로 바꾼다)
    if (sumproductCompares(formula)) return CF_CANNOT;
    try {
      // 열 전체 범위($A:$A)는 엔진이 잘못 읽는다(열 번호가 음수가 됨) — 시트 끝 행까지의 보통 범위로 바꿔 넘긴다
      const rows = Math.max(1, extent().rows);
      const text = formula.replace(COLUMN_RANGE, (_m, a1: string, c1: string, a2: string, c2: string) => `${a1}${c1}$1:${a2}${c2}$${rows}`);
      const v = parser.parse(text, { sheet: SHEET, row: row + 1, col: col + 1 }) as unknown;
      const first = Array.isArray(v) ? (Array.isArray(v[0]) ? v[0][0] : v[0]) : v;
      // 모르는 함수(#NAME?)는 계산하지 못함 — 알린다. 그 밖의 오류값(#N/A 등)은 안 맞음
      if (first instanceof Error || (first !== null && typeof first === 'object')) return /#NAME\?/.test(String((first as { _error?: string })._error ?? first)) ? CF_CANNOT : undefined;
      if (first === undefined || first === '') return null;
      return first as CfVal;
    } catch {
      return CF_CANNOT; // 엔진이 읽지 못한 수식
    }
  };
}

/** SUMPRODUCT( … ) 괄호 안에 비교(=·<·>)가 있는지(따옴표 안 글자는 빼고) */
export function sumproductCompares(formula: string): boolean {
  const re = /SUMPRODUCT\s*\(/gi;
  for (let m = re.exec(formula); m; m = re.exec(formula)) {
    let depth = 1;
    for (let i = m.index + m[0].length; i < formula.length && depth > 0; i += 1) {
      const ch = formula[i];
      if (ch === '"') {
        const end = formula.indexOf('"', i + 1);
        i = end < 0 ? formula.length : end;
      } else if (ch === '(') depth += 1;
      else if (ch === ')') depth -= 1;
      else if (ch === '=' || ch === '<' || ch === '>') return true;
    }
  }
  return false;
}
