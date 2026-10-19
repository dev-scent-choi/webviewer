/**
 * 파일의 조건부 서식(<conditionalFormatting>)과 그 강조 서식(styles.xml의 <dxfs>)을 읽어 편집기가 쓰는 형식으로 바꾼다.
 * 화면에 그리지 못하는 종류(수식·아이콘 집합·날짜 등)도 `t: 'other'`로 목록에 남긴다 — 순서(우선순위)와 행/열 이동을
 * 함께 관리해야 저장 때 잃지 않는다(condFormatShared.ts 머리 주석 참고).
 */
import { parseSqref, type CfCellOp, type CfRule, type CfStyle, type CondFormatEntry } from '../collab/condFormatShared.js';
import { MAX_COLS, MAX_ROWS, resolveColor } from '../spreadsheet.js';

function decodeXml(value: string): string {
  return value
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(Number(dec)))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function attr(attrs: string, name: string): string | undefined {
  const m = new RegExp(`\\b${name}="([^"]*)"`).exec(attrs);
  return m ? decodeXml(m[1]) : undefined;
}

/** <color rgb=… | theme=… tint=… | indexed=…> → '#RRGGBB' */
function colorOf(tag: string | undefined, themeColors: readonly string[]): string | undefined {
  if (!tag) return undefined;
  const rgb = attr(tag, 'rgb');
  const theme = attr(tag, 'theme');
  const tint = attr(tag, 'tint');
  const indexed = attr(tag, 'indexed');
  const css = resolveColor(
    {
      argb: rgb,
      theme: theme === undefined ? undefined : Number(theme),
      tint: tint === undefined ? undefined : Number(tint),
      indexed: indexed === undefined ? undefined : Number(indexed),
    },
    themeColors,
  );
  return css?.toUpperCase();
}

/** styles.xml의 <dxfs> — 조건부 서식이 dxfId로 가리키는 강조 서식(글자 색·굵게·채우기) */
export function parseDxfs(stylesXml: string | undefined, themeColors: readonly string[]): CfStyle[] {
  const block = stylesXml?.match(/<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/)?.[1];
  if (!block) return [];
  return [...block.matchAll(/<dxf>([\s\S]*?)<\/dxf>|<dxf\/>/g)].map((m) => {
    const xml = m[1] ?? '';
    const font = xml.match(/<font>([\s\S]*?)<\/font>/)?.[1] ?? '';
    const fill = xml.match(/<fill>([\s\S]*?)<\/fill>/)?.[1] ?? '';
    const col = colorOf(font.match(/<color\b[^>]*\/?>/)?.[0], themeColors);
    const bg = colorOf(fill.match(/<bgColor\b[^>]*\/?>/)?.[0], themeColors) ?? colorOf(fill.match(/<fgColor\b[^>]*\/?>/)?.[0], themeColors);
    const bold = /<b\s*\/>|<b\s+val="(1|true)"\s*\/>/.test(font);
    const style: { bg?: string; col?: string; b?: 1 } = {};
    if (bg) style.bg = bg;
    if (col) style.col = col;
    if (bold) style.b = 1;
    return style;
  });
}

const CELL_OPS: Record<string, CfCellOp> = {
  greaterThan: 'gt',
  lessThan: 'lt',
  greaterThanOrEqual: 'ge',
  lessThanOrEqual: 'le',
  equal: 'eq',
  notEqual: 'ne',
  between: 'between',
  notBetween: 'notBetween',
};

function ruleOf(attrs: string, body: string, dxfs: readonly CfStyle[], themeColors: readonly string[]): CfRule {
  const type = attr(attrs, 'type');
  const dxfId = attr(attrs, 'dxfId');
  const s: CfStyle = dxfId === undefined ? {} : (dxfs[Number(dxfId)] ?? {});
  const formulas = [...body.matchAll(/<formula>([\s\S]*?)<\/formula>/g)].map((m) => decodeXml(m[1]));
  switch (type) {
    case 'cellIs': {
      const op = CELL_OPS[attr(attrs, 'operator') ?? ''];
      if (!op || formulas.length === 0) return { t: 'other' };
      return { t: 'cell', op, v1: formulas[0], ...(formulas[1] !== undefined ? { v2: formulas[1] } : {}), s };
    }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith': {
      const text = attr(attrs, 'text');
      if (text === undefined) return { t: 'other' };
      const op = type === 'containsText' ? 'contains' : type === 'notContainsText' ? 'notContains' : type === 'beginsWith' ? 'begins' : 'ends';
      return { t: 'text', op, text, s };
    }
    case 'duplicateValues':
      return { t: 'dup', s };
    case 'uniqueValues':
      return { t: 'dup', unique: 1, s };
    case 'top10': {
      const rank = Number(attr(attrs, 'rank') ?? '10');
      return {
        t: 'top',
        n: Number.isFinite(rank) && rank > 0 ? rank : 10,
        ...(/^(1|true)$/.test(attr(attrs, 'bottom') ?? '') ? { bottom: 1 as const } : {}),
        ...(/^(1|true)$/.test(attr(attrs, 'percent') ?? '') ? { percent: 1 as const } : {}),
        s,
      };
    }
    case 'aboveAverage': {
      if (attr(attrs, 'stdDev') !== undefined || /^(1|true)$/.test(attr(attrs, 'equalAverage') ?? '')) return { t: 'other' };
      return { t: 'avg', ...(/^(0|false)$/.test(attr(attrs, 'aboveAverage') ?? '') ? { below: 1 as const } : {}), s };
    }
    case 'colorScale': {
      const colors = [...body.matchAll(/<color\b[^>]*\/?>/g)].map((m) => colorOf(m[0], themeColors)).filter((c): c is string => Boolean(c));
      return colors.length === 2 || colors.length === 3 ? { t: 'scale', colors } : { t: 'other' };
    }
    case 'dataBar': {
      const color = colorOf(body.match(/<color\b[^>]*\/?>/)?.[0], themeColors);
      return color ? { t: 'bar', color } : { t: 'other' };
    }
    default:
      return { t: 'other' };
  }
}

/** 시트 XML 하나의 조건부 서식 — 우선순위 순(priority 1이 맨 앞). extLst 안의 x14 규칙은 읽지 않는다. */
export function parseCondFormats(sheetXml: string, dxfs: readonly CfStyle[], themeColors: readonly string[]): CondFormatEntry[] {
  const found: Array<CondFormatEntry & { order: number }> = [];
  let order = 0;
  for (const block of sheetXml.matchAll(/<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g)) {
    const sqref = attr(block[1], 'sqref');
    if (!sqref) continue;
    const ranges = parseSqref(sqref, MAX_ROWS - 1, MAX_COLS - 1);
    if (ranges.length === 0) continue;
    for (const rule of block[2].matchAll(/<cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cfRule>)/g)) {
      const priority = Number(attr(rule[1], 'priority') ?? NaN);
      const p = Number.isFinite(priority) ? priority : undefined;
      order += 1;
      found.push({ id: `f${order}`, ranges, rule: ruleOf(rule[1], rule[2] ?? '', dxfs, themeColors), ...(p !== undefined ? { p } : {}), order });
    }
  }
  found.sort((a, b) => (a.p ?? Number.MAX_SAFE_INTEGER) - (b.p ?? Number.MAX_SAFE_INTEGER) || a.order - b.order);
  return found.map(({ order: _order, ...entry }) => entry);
}
