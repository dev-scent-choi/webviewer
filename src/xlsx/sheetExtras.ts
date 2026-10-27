/**
 * 시트의 하이퍼링크(바로가기)·목록 상자(데이터 유효성 "목록")·표(Excel 표 서식)의 자동 필터를 원본 xlsx에서 읽는다.
 * 그리드 시딩(ExcelJS)은 이 셋을 버리거나(하이퍼링크는 글자만 남김) 아예 모른다(표 안 필터) — 그래서 편집 화면에서
 * 바로가기를 눌러도 아무 일이 없었고, 표 서식 범위의 필터 단추가 보이지 않았다.
 *
 * 편집 화면은 이 정보를 읽기 전용으로 받아 쓴다(링크 이동·목록에서 고르기). 저장할 때 하이퍼링크·유효성 검사는
 * ExcelJS가 원본에서 읽은 것을 그대로 다시 쓴다.
 */
import { strFromU8, unzipSync } from 'fflate';
import { parseRangeFormula } from './chartModel.js';
import { decodeXmlAttr, relationshipsOf, relsPathOf, resolvePartPath, sheetPathsByName } from './drawingPassthrough.js';
import type { SheetFilter } from './filterModel.js';

export interface SheetLink {
  readonly sheetName: string;
  /** 0부터, 링크가 걸린 칸 범위 */
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
  /** 바깥 주소(http·mailto 등) — 있으면 새 창으로 */
  readonly url?: string;
  /** 문서 안 위치('자료'!A1, 요약!C5 등) — 있으면 그 시트·칸으로 이동 */
  readonly location?: string;
  /** 마우스를 올렸을 때 보일 설명 */
  readonly tooltip?: string;
  /** 주소를 칸 값으로 만드는 수식 링크 — =HYPERLINK("#'"&$A2&"'!A1", …)의 첫 인자. 편집 화면이 누른 칸 기준으로 푼다
   * (anchor는 수식이 적힌 칸 — 아래로 복사한 공유 수식이면 상대 참조를 그만큼 옮긴다) */
  readonly expr?: string;
  readonly anchor?: { readonly row: number; readonly col: number };
}

export interface SheetListValidation {
  readonly sheetName: string;
  /** 이 목록이 걸린 칸 범위들(0부터) */
  readonly ranges: ReadonlyArray<{ readonly r1: number; readonly c1: number; readonly r2: number; readonly c2: number }>;
  /** "김철수,이영희"처럼 직접 적은 항목 */
  readonly values?: readonly string[];
  /** =$H$1:$H$5, =목록!$A$1:$A$3처럼 칸 범위를 가리키는 목록 — 편집 화면이 지금 값으로 푼다 */
  readonly source?: { readonly sheetName?: string; readonly r1: number; readonly c1: number; readonly r2: number; readonly c2: number };
}

export interface SheetExtras {
  readonly links: SheetLink[];
  readonly lists: SheetListValidation[];
}

/** 바깥 주소로 열어도 되는 것만(스크립트 주소 등은 버린다) */
function safeUrl(url: string): string | undefined {
  return /^(https?:|mailto:|tel:)/i.test(url.trim()) ? url.trim() : undefined;
}

function rangesOfSqref(sqref: string) {
  return sqref
    .split(/\s+/)
    .map((part) => parseRangeFormula(part))
    .filter((r): r is NonNullable<typeof r> => r !== null)
    .map(({ r1, c1, r2, c2 }) => ({ r1, c1, r2, c2 }));
}

/** 수식 안 HYPERLINK( 의 첫 인자 글 — IFERROR(HYPERLINK(...), "")처럼 감싸여 있어도 찾는다. 없으면 undefined */
export function hyperlinkFirstArg(formula: string): string | undefined {
  const start = /\bHYPERLINK\s*\(/i.exec(formula);
  if (!start) return undefined;
  const from = start.index + start[0].length;
  let depth = 0;
  let quoted = false;
  for (let i = from; i < formula.length; i += 1) {
    const ch = formula[i];
    if (quoted) {
      if (ch === '"' && formula[i + 1] === '"') i += 1;
      else if (ch === '"') quoted = false;
    } else if (ch === '"') quoted = true;
    else if (ch === '(') depth += 1;
    else if (ch === ')' && depth > 0) depth -= 1;
    else if ((ch === ',' || ch === ')') && depth === 0) return formula.slice(from, i).trim();
  }
  return undefined;
}

function linkOf(sheetName: string, ref: { r1: number; c1: number; r2: number; c2: number }, target: string, tooltip?: string): SheetLink | null {
  if (target.startsWith('#')) return { sheetName, ...ref, location: target.slice(1), ...(tooltip ? { tooltip } : {}) };
  const url = safeUrl(target);
  return url ? { sheetName, ...ref, url, ...(tooltip ? { tooltip } : {}) } : null;
}

export function parseSheetExtras(sheetName: string, sheetXml: string, sheetRelsXml: string | undefined): SheetExtras {
  const rels = new Map(relationshipsOf(sheetRelsXml ?? '').map((r) => [r.id, r]));
  const links: SheetLink[] = [];
  const linked = new Set<string>();
  for (const m of sheetXml.matchAll(/<hyperlink\b[^>]*\/?>/g)) {
    const attr = (name: string) => {
      const v = new RegExp(`\\b${name}="([^"]*)"`).exec(m[0])?.[1];
      return v === undefined ? undefined : decodeXmlAttr(v);
    };
    const ref = rangesOfSqref(attr('ref') ?? '')[0];
    if (!ref) continue;
    const tooltip = attr('tooltip');
    const rid = attr('r:id');
    const location = attr('location');
    const rel = rid ? rels.get(rid) : undefined;
    // 문서 안 위치가 우선 — ExcelJS는 내부 링크에도 관계(r:id)를 같이 쓰므로 location을 먼저 본다
    const link = location
      ? { sheetName, ...ref, location: location.replace(/^#/, ''), ...(tooltip ? { tooltip } : {}) }
      : rel
        ? linkOf(sheetName, ref, rel.external ? decodeXmlAttr(rel.target) : `#${decodeXmlAttr(rel.target)}`, tooltip)
        : null;
    if (link) {
      links.push(link);
      linked.add(`${ref.r1}:${ref.c1}`);
    }
  }
  // =HYPERLINK("#시트!A1","글자") 수식 — 링크 목록에 없는 칸만. IFERROR 등으로 감싸도, 주소를 칸 값으로 만들어도
  // ("#'"&$A2&"'!a1") 링크로 본다. 아래로 복사한 공유 수식(t="shared" ref="B2:B50")은 첫 칸에만 글이 있어 그 범위 전체를 링크로.
  // 빈 칸(<c r="A1"/>)에서 시작해 다음 칸의 수식을 잘못 잡지 않게, 닫는 칸·다음 칸을 넘지 않는다
  for (const m of sheetXml.matchAll(/<c\b[^>]*\br="([A-Z]{1,3}\d+)"[^>]*(?<!\/)>(?:(?!<\/c>|<c\b)[\s\S])*?<f\b([^>]*)>([^<]*)<\/f>/g)) {
    const arg = hyperlinkFirstArg(decodeXmlAttr(m[3]));
    const cell = parseRangeFormula(m[1]);
    if (arg === undefined || !cell || linked.has(`${cell.r1}:${cell.c1}`)) continue;
    const literal = /^"((?:[^"]|"")*)"$/.exec(arg);
    if (literal) {
      const link = linkOf(sheetName, { r1: cell.r1, c1: cell.c1, r2: cell.r2, c2: cell.c2 }, literal[1].replace(/""/g, '"'));
      if (link) links.push(link);
      continue;
    }
    const shared = /\bt="shared"/.test(m[2]) ? parseRangeFormula(/\bref="([^"]+)"/.exec(m[2])?.[1] ?? '') : null;
    const range = shared ?? cell;
    links.push({ sheetName, r1: range.r1, c1: range.c1, r2: range.r2, c2: range.c2, expr: arg, anchor: { row: cell.r1, col: cell.c1 } });
  }

  const lists: SheetListValidation[] = [];
  for (const m of sheetXml.matchAll(/<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/g)) {
    if (!/\btype="list"/.test(m[1])) continue;
    if (/\bshowDropDown="(1|true)"/.test(m[1])) continue; // 엑셀에서 "드롭다운 표시" 끔(이름과 반대 뜻)
    const ranges = rangesOfSqref(decodeXmlAttr(/\bsqref="([^"]*)"/.exec(m[1])?.[1] ?? ''));
    const formula = decodeXmlAttr(/<formula1>([\s\S]*?)<\/formula1>/.exec(m[2])?.[1] ?? '').trim();
    if (ranges.length === 0 || !formula) continue;
    if (formula.startsWith('"')) {
      const values = formula.replace(/^"|"$/g, '').replace(/""/g, '"').split(',').map((v) => v.trim()).filter(Boolean);
      if (values.length > 0) lists.push({ sheetName, ranges, values });
      continue;
    }
    const source = parseRangeFormula(formula.replace(/^=/, ''));
    if (source) lists.push({ sheetName, ranges, source });
  }
  return { links, lists };
}

/** 원본 xlsx의 모든 시트 */
export function readSheetExtras(xlsx: Uint8Array): SheetExtras {
  return readSheetExtrasFromEntries(unzipSync(xlsx));
}

/** 이미 푼 파트 모음에서 — 큰 파일은 데이터를 뺀 뼈대(zipSkeleton.ts readXlsxSkeleton)를 넘긴다 */
export function readSheetExtrasFromEntries(entries: Readonly<Record<string, Uint8Array>>): SheetExtras {
  const out: SheetExtras = { links: [], lists: [] };
  for (const [sheetName, sheetPath] of sheetPathsByName(entries)) {
    const sheet = entries[sheetPath];
    if (!sheet) continue;
    const rels = entries[relsPathOf(sheetPath)];
    const extras = parseSheetExtras(sheetName, strFromU8(sheet), rels ? strFromU8(rels) : undefined);
    out.links.push(...extras.links);
    out.lists.push(...extras.lists);
  }
  return out;
}

/**
 * 시트에 연결된 표(<tableParts>) 가운데 자동 필터가 켜진 첫 표의 필터 범위 — 시트 자체 자동 필터가 없을 때
 * 편집 화면의 필터 단추로 쓴다. table 필드에 표 파트 경로를 남겨, 저장할 때 시트 자동 필터로 다시 쓰지 않게 한다
 * (표와 겹치는 시트 자동 필터는 엑셀이 파일을 고쳐야 한다며 연다).
 */
export function tableFilterOf(entries: Record<string, Uint8Array>, sheetPath: string): SheetFilter | null {
  const relsBytes = entries[relsPathOf(sheetPath)];
  const sheet = entries[sheetPath];
  if (!relsBytes || !sheet) return null;
  const rels = new Map(relationshipsOf(strFromU8(relsBytes)).map((r) => [r.id, r]));
  for (const m of strFromU8(sheet).matchAll(/<tablePart\b[^>]*\br:id="([^"]+)"/g)) {
    const rel = rels.get(m[1]);
    if (!rel) continue;
    const tablePath = resolvePartPath(sheetPath, rel.target);
    const tableBytes = entries[tablePath];
    if (!tableBytes) continue;
    const xml = strFromU8(tableBytes);
    const ref = /<autoFilter\b[^>]*\bref="([^"]+)"/.exec(xml)?.[1];
    const range = ref ? parseRangeFormula(ref) : null;
    if (range) return { r1: range.r1, c1: range.c1, r2: range.r2, c2: range.c2, criteria: {}, table: tablePath };
  }
  return null;
}
