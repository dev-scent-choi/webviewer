/**
 * 파일의 조건부 서식(<conditionalFormatting>)과 그 강조 서식(styles.xml의 <dxfs>)을 읽어 편집기가 쓰는 형식으로 바꾼다.
 * 화면에 그리지 못하는 종류(수식·아이콘 집합·날짜 등)도 `t: 'other'`로 목록에 남긴다 — 순서(우선순위)와 행/열 이동을
 * 함께 관리해야 저장 때 잃지 않는다(condFormatShared.ts 머리 주석 참고).
 */
import { parseSqref, type CfCellOp, type CfRule, type CfStyle, type CfThreshold, type CondFormatEntry } from '../collab/condFormatShared.js';
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

/** cfvo 목록(문턱값) — <cfvo type val/>, x14는 <x14:cfvo type><xm:f>값</xm:f></x14:cfvo>. 모르는 종류가 있으면 undefined */
function cfvosOf(body: string): CfThreshold[] | undefined {
  const out: CfThreshold[] = [];
  for (const m of body.matchAll(/<(?:x14:)?cfvo\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:x14:)?cfvo>)/g)) {
    const kind = attr(m[1], 'type');
    const raw = attr(m[1], 'val') ?? m[2]?.match(/<xm:f>([\s\S]*?)<\/xm:f>/)?.[1];
    const val = raw === undefined ? undefined : decodeXml(raw);
    const gt = /^(0|false)$/.test(attr(m[1], 'gte') ?? '') ? { gt: 1 as const } : {};
    if (kind === 'min' || kind === 'max' || kind === 'autoMin' || kind === 'autoMax') out.push({ type: kind, ...gt });
    else if (kind === 'num' || kind === 'percent' || kind === 'percentile' || kind === 'formula') {
      const n = val === undefined ? NaN : Number(val);
      // 값 자리에 수식(=$B$1, 다른 칸)이 있으면 수식 문턱값으로
      if (Number.isFinite(n) && kind !== 'formula') out.push({ type: kind, v: n, ...gt });
      else if (val) out.push({ type: kind === 'formula' ? 'formula' : kind, f: val.replace(/^=/, ''), ...gt });
      else return undefined;
    } else return undefined;
  }
  return out.length > 0 ? out : undefined;
}

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
      const th = cfvosOf(body);
      if (colors.length !== 2 && colors.length !== 3) return { t: 'other' };
      return th && th.length === colors.length ? { t: 'scale', colors, th } : { t: 'scale', colors };
    }
    case 'dataBar': {
      const color = colorOf(body.match(/<color\b[^>]*\/?>/)?.[0], themeColors);
      if (!color) return { t: 'other' };
      const th = cfvosOf(body);
      const tag = body.match(/<dataBar\b[^>]*>/)?.[0] ?? '';
      const minL = Number(attr(tag, 'minLength') ?? 10);
      const maxL = Number(attr(tag, 'maxLength') ?? 90);
      return {
        t: 'bar',
        color,
        ...(th && th.length === 2 ? { th: [th[0], th[1]] as const } : {}),
        len: [Number.isFinite(minL) ? minL : 10, Number.isFinite(maxL) ? maxL : 90] as const,
      };
    }
    case 'iconSet': {
      const tag = body.match(/<(?:x14:)?iconSet\b[^>]*>/)?.[0] ?? '';
      const th = cfvosOf(body);
      if (!th) return { t: 'other' };
      return {
        t: 'icon',
        set: attr(tag, 'iconSet') ?? '3TrafficLights1',
        th: th.slice(1),
        ...(/^(1|true)$/.test(attr(tag, 'reverse') ?? '') ? { reverse: 1 as const } : {}),
        ...(/^(0|false)$/.test(attr(tag, 'showValue') ?? '') ? { hide: 1 as const } : {}),
      };
    }
    case 'expression':
      return formulas[0] ? { t: 'expr', f: formulas[0], s } : { t: 'other' };
    default:
      return { t: 'other' };
  }
}

/** 시트 XML 하나의 조건부 서식 — 우선순위 순(priority 1이 맨 앞). "true이면 중지"도 읽는다.
 * extLst의 x14 확장(엑셀 2010+): 아이콘 집합(3Stars·3Triangles·5Boxes 등 확장 집합)은 따로 읽어 화면 판정에만 쓰고(저장은 파일의
 * 확장 부분을 그대로 둔다 — p 대신 xp), 데이터 막대 확장 설정(단색·막대 길이·자동 최소/최대·음수 색)은 본 규칙에 덧붙인다. */
export function parseCondFormats(sheetXml: string, dxfs: readonly CfStyle[], themeColors: readonly string[]): CondFormatEntry[] {
  const found: Array<CondFormatEntry & { order: number }> = [];
  let order = 0;
  // x14 확장 규칙: id → 본문(데이터 막대 확장 설정을 본 규칙에 붙일 때 쓴다)
  const extById = new Map<string, string>();
  for (const block of sheetXml.matchAll(/<x14:conditionalFormatting\b[^>]*>([\s\S]*?)<\/x14:conditionalFormatting>/g)) {
    const sqref = block[1].match(/<xm:sqref>([\s\S]*?)<\/xm:sqref>/)?.[1];
    for (const rule of block[1].matchAll(/<x14:cfRule\b([^>]*?)>([\s\S]*?)<\/x14:cfRule>/g)) {
      const id = attr(rule[1], 'id');
      if (id) extById.set(id, rule[2]);
      if (attr(rule[1], 'type') !== 'iconSet' || !sqref) continue;
      const ranges = parseSqref(sqref, MAX_ROWS - 1, MAX_COLS - 1);
      const parsed = ruleOf(rule[1], rule[2], dxfs, themeColors);
      const priority = Number(attr(rule[1], 'priority') ?? NaN);
      order += 1;
      if (ranges.length > 0 && parsed.t === 'icon') found.push({ id: `x${order}`, ranges, rule: parsed, ...(Number.isFinite(priority) ? { xp: priority } : {}), order });
    }
  }
  for (const block of sheetXml.matchAll(/<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g)) {
    const sqref = attr(block[1], 'sqref');
    if (!sqref) continue;
    const ranges = parseSqref(sqref, MAX_ROWS - 1, MAX_COLS - 1);
    if (ranges.length === 0) continue;
    for (const rule of block[2].matchAll(/<cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cfRule>)/g)) {
      const priority = Number(attr(rule[1], 'priority') ?? NaN);
      const p = Number.isFinite(priority) ? priority : undefined;
      let parsed = ruleOf(rule[1], rule[2] ?? '', dxfs, themeColors);
      const extId = rule[2]?.match(/<x14:id>([\s\S]*?)<\/x14:id>/)?.[1];
      const ext = extId ? extById.get(extId) : undefined;
      if (parsed.t === 'bar' && ext) parsed = withBarExt(parsed, ext, themeColors);
      order += 1;
      found.push({
        id: `f${order}`,
        ranges,
        rule: parsed,
        ...(p !== undefined ? { p } : {}),
        ...(/^(1|true)$/.test(attr(rule[1], 'stopIfTrue') ?? '') ? { stop: 1 as const } : {}),
        order,
      });
    }
  }
  const prio = (e: CondFormatEntry) => e.p ?? e.xp ?? Number.MAX_SAFE_INTEGER;
  found.sort((a, b) => prio(a) - prio(b) || a.order - b.order);
  return found.map(({ order: _order, ...entry }) => entry);
}

/** x14 데이터 막대 설정 — 막대 길이(minLength·maxLength, 보통 0·100), 자동 최소·최대(0을 포함), 음수 막대 색 */
function withBarExt(rule: Extract<CfRule, { t: 'bar' }>, ext: string, themeColors: readonly string[]): CfRule {
  const tag = ext.match(/<x14:dataBar\b[^>]*>/)?.[0] ?? '';
  const minL = Number(attr(tag, 'minLength') ?? 0);
  const maxL = Number(attr(tag, 'maxLength') ?? 100);
  const th = cfvosOf(ext);
  const neg = colorOf(ext.match(/<x14:negativeFillColor\b[^>]*\/?>/)?.[0], themeColors);
  return {
    ...rule,
    len: [Number.isFinite(minL) ? minL : 0, Number.isFinite(maxL) ? maxL : 100],
    ...(th && th.length === 2 ? { th: [th[0], th[1]] as const } : {}),
    ...(neg ? { neg } : {}),
  };
}
