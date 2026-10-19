/**
 * 조건부 서식을 저장 파일(ExcelJS 워크시트)에 쓴다.
 *
 * 1) applyEditedCondFormats — 편집기에서 규칙을 바꿨거나 행/열 이동으로 범위가 바뀐 시트(condFormatsEdited)만 규칙 목록을
 *    sheetMeta.condFormats 순서·범위대로 다시 쓴다. 파일에 원래 있던 규칙은 ExcelJS가 읽어 둔 원래 객체를 priority로 찾아
 *    그대로 쓰므로 화면에 못 그리는 종류(수식·아이콘 집합 등)도 잃지 않는다.
 * 2) preserveUnsupportedCfRules — ExcelJS는 "중복 값·고유 값·다음으로 시작/끝남·텍스트 포함 안 함" 규칙을 읽기만 하고
 *    쓰지 못해(cf-rule-xform의 render에 그 종류가 없다) 저장하면 조용히 사라졌다. 엑셀에서 똑같이 동작하는 수식 규칙
 *    (type: expression)으로 바꿔 보존한다. 모든 시트에 적용한다 — 그런 규칙이 없는 시트에는 아무 일도 하지 않는다.
 */
import type ExcelJS from 'exceljs';
import { parseSqref, rectA1, type CfRect, type CfRule, type CfStyle, type CondFormatEntry } from '../collab/condFormatShared.js';
import type { LazyGridSheetMeta } from './lazyGridStore.js';

type ExcelRule = Record<string, unknown> & { type?: string; priority?: number; formulae?: string[]; text?: string; operator?: string };
interface ExcelCf {
  ref: string;
  rules: ExcelRule[];
}
type CfSheet = { conditionalFormattings: ExcelCf[] };

const quote = (text: string) => `"${text.replace(/"/g, '""')}"`;
const argb = (hex: string) => `FF${hex.replace('#', '').toUpperCase()}`;

function excelStyle(s: CfStyle): Record<string, unknown> {
  const style: Record<string, unknown> = {};
  if (s.col || s.b) style.font = { ...(s.col ? { color: { argb: argb(s.col) } } : {}), ...(s.b ? { bold: true } : {}) };
  if (s.bg) style.fill = { type: 'pattern', pattern: 'solid', bgColor: { argb: argb(s.bg) } };
  return style;
}

/** 규칙 값 자리 — 숫자·칸 주소·이미 따옴표 친 글자는 그대로, 그 밖의 글자는 따옴표로 */
function operand(value: string): string {
  const v = value.trim().replace(/^=/, '').trim();
  if (/^-?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$/.test(v) || /^\$?[A-Z]{1,3}\$?\d+$/i.test(v) || /^".*"$/.test(v)) return v;
  return quote(v);
}

function countIfDup(rects: readonly CfRect[], tl: string): string {
  const sum = rects.map((r) => `COUNTIF(${rectA1(r, true)},${tl})`).join('+');
  return rects.length > 1 ? `(${sum})` : sum;
}

/** 텍스트 규칙(포함·포함 안 함·시작·끝)과 중복/고유 값을 엑셀 수식 규칙으로 */
function textFormula(op: string, text: string, tl: string): string {
  const t = quote(text);
  if (op === 'contains') return `NOT(ISERROR(SEARCH(${t},${tl})))`;
  if (op === 'notContains') return `ISERROR(SEARCH(${t},${tl}))`;
  if (op === 'begins') return `LEFT(${tl},LEN(${t}))=${t}`;
  return `RIGHT(${tl},LEN(${t}))=${t}`;
}

const CELL_OPS: Record<string, string> = {
  gt: 'greaterThan',
  lt: 'lessThan',
  ge: 'greaterThanOrEqual',
  le: 'lessThanOrEqual',
  eq: 'equal',
  ne: 'notEqual',
  between: 'between',
  notBetween: 'notBetween',
};

/** 편집기 형식의 규칙 → ExcelJS 규칙(쓸 수 없는 종류면 null) */
export function toExcelRule(rule: CfRule, rects: readonly CfRect[]): ExcelRule | null {
  if (rects.length === 0) return null;
  const tl = rectA1({ r0: rects[0].r0, r1: rects[0].r0, c0: rects[0].c0, c1: rects[0].c0 });
  switch (rule.t) {
    case 'cell':
      return {
        type: 'cellIs',
        operator: CELL_OPS[rule.op],
        formulae: rule.v2 !== undefined ? [operand(rule.v1), operand(rule.v2)] : [operand(rule.v1)],
        style: excelStyle(rule.s),
      };
    case 'text':
      return { type: 'expression', formulae: [textFormula(rule.op, rule.text, tl)], style: excelStyle(rule.s) };
    case 'dup':
      return {
        type: 'expression',
        formulae: [`AND(${tl}<>"",${countIfDup(rects, tl)}${rule.unique ? '=1' : '>1'})`],
        style: excelStyle(rule.s),
      };
    case 'top':
      return { type: 'top10', rank: rule.n, percent: Boolean(rule.percent), bottom: Boolean(rule.bottom), style: excelStyle(rule.s) };
    case 'avg':
      return { type: 'aboveAverage', aboveAverage: !rule.below, style: excelStyle(rule.s) };
    case 'scale':
      return {
        type: 'colorScale',
        cfvo: rule.colors.length === 3 ? [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }] : [{ type: 'min' }, { type: 'max' }],
        color: rule.colors.map((c) => ({ argb: argb(c) })),
      };
    case 'bar':
      return { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: argb(rule.color) } };
    default:
      return null;
  }
}

/** ExcelJS가 쓰지 못하는 원래 규칙을 같은 뜻의 수식 규칙으로(바꿀 필요 없으면 그대로) */
function preservable(rule: ExcelRule, ref: string): ExcelRule {
  const rects = parseSqref(ref);
  if (rects.length === 0) return rule;
  const tl = rectA1({ r0: rects[0].r0, r1: rects[0].r0, c0: rects[0].c0, c1: rects[0].c0 });
  const text = typeof rule.text === 'string' ? rule.text : '';
  const keep = { style: rule.style, priority: rule.priority, ...(rule.stopIfTrue ? { stopIfTrue: rule.stopIfTrue } : {}) };
  switch (rule.type) {
    case 'duplicateValues':
    case 'uniqueValues':
      return { type: 'expression', formulae: [`AND(${tl}<>"",${countIfDup(rects, tl)}${rule.type === 'uniqueValues' ? '=1' : '>1'})`], ...keep };
    case 'beginsWith':
    case 'endsWith':
    case 'notContainsText': {
      const op = rule.type === 'beginsWith' ? 'begins' : rule.type === 'endsWith' ? 'ends' : 'notContains';
      const formula = rule.formulae?.[0] ?? (text ? textFormula(op, text, tl) : undefined);
      return formula ? { type: 'expression', formulae: [formula], ...keep } : rule;
    }
    default:
      return rule;
  }
}

export function preserveUnsupportedCfRules(sheet: ExcelJS.Worksheet): void {
  const ws = sheet as unknown as CfSheet;
  if (!ws.conditionalFormattings?.length) return;
  ws.conditionalFormattings = ws.conditionalFormattings.map((cf) => ({ ...cf, rules: cf.rules.map((rule) => preservable(rule, cf.ref)) }));
}

export function applyEditedCondFormats(
  sheetMeta: readonly LazyGridSheetMeta[],
  sheets: ReadonlyArray<ExcelJS.Worksheet | undefined>,
  removedSheets: ReadonlySet<number>,
): void {
  sheetMeta.forEach((meta, sheetIdx) => {
    const sheet = sheets[sheetIdx];
    if (!meta.condFormatsEdited || !sheet || removedSheets.has(sheetIdx)) return;
    const ws = sheet as unknown as CfSheet;
    const originals = (ws.conditionalFormattings ?? []).flatMap((cf) => cf.rules.map((rule) => rule));
    const byPriority = new Map<number, ExcelRule>();
    for (const rule of originals) if (typeof rule.priority === 'number' && !byPriority.has(rule.priority)) byPriority.set(rule.priority, rule);
    const next: ExcelCf[] = [];
    (meta.condFormats ?? []).forEach((entry: CondFormatEntry) => {
      const ref = entry.ranges.map((r) => rectA1(r)).join(' ');
      if (!ref) return;
      const fileOrder = /^f(\d+)$/.exec(entry.id);
      const original = entry.p !== undefined ? byPriority.get(entry.p) : fileOrder ? originals[Number(fileOrder[1]) - 1] : undefined;
      const rule = original ? { ...original } : toExcelRule(entry.rule, entry.ranges);
      if (!rule) return;
      next.push({ ref, rules: [{ ...rule, priority: next.length + 1 }] });
    });
    // 편집기가 읽지 않은 원래 규칙(extLst 안의 x14 규칙 등)은 지우지도 옮기지도 않은 채 뒤에 그대로 둔다
    const known = new Set(meta.condFormatFilePriorities ?? []);
    for (const cf of ws.conditionalFormattings ?? []) {
      const untouched = cf.rules.filter((rule) => typeof rule.priority === 'number' && !known.has(rule.priority));
      for (const rule of untouched) next.push({ ref: cf.ref, rules: [{ ...rule, priority: next.length + 1 }] });
    }
    ws.conditionalFormattings = next;
  });
}
