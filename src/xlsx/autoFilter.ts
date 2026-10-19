/**
 * 시트 자동 필터(<autoFilter>)를 파일에서 읽고, 저장 파일에 조건(<filterColumn>)까지 다시 쓴다.
 * ExcelJS는 자동 필터 범위만 다루고 조건은 버리므로, 조건은 쓰기가 끝난 zip에서 직접 붙인다.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { parseRangeFormula } from './chartModel.js';
import { sheetPathsByName } from './drawingPassthrough.js';
import { tableFilterOf } from './sheetExtras.js';
import type { SheetFilter } from './filterModel.js';

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

function escapeXml(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** 시트 XML 하나의 자동 필터(범위 + 값 목록 조건) */
export function parseAutoFilter(sheetXml: string): SheetFilter | null {
  const m = /<autoFilter\b[^>]*\bref="([^"]+)"[^>]*?(?:\/>|>([\s\S]*?)<\/autoFilter>)/.exec(sheetXml);
  if (!m) return null;
  const range = parseRangeFormula(m[1]);
  if (!range) return null;
  const criteria: Record<string, string[]> = {};
  for (const col of (m[2] ?? '').matchAll(/<filterColumn\b[^>]*colId="(\d+)"[^>]*>([\s\S]*?)<\/filterColumn>/g)) {
    const values = [...col[2].matchAll(/<filter\b[^>]*val="([^"]*)"/g)].map((v) => decodeXml(v[1]));
    if (/<filters\b[^>]*blank="1"/.test(col[2])) values.push('');
    if (values.length > 0) criteria[String(range.c1 + Number(col[1]))] = values;
  }
  return { r1: range.r1, c1: range.c1, r2: Math.max(range.r2, contiguousDataEnd(sheetXml, range.r1, range.c1, range.c2)), c2: range.c2, criteria };
}

function columnIndexOf(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/**
 * 필터 범위가 머리글 한 줄뿐일 때(예: ref="A1:G1" — 엑셀이 아닌 도구가 만든 파일에 흔하다) 엑셀은 그 아래로 빈 줄 없이
 * 이어진 데이터를 필터 대상으로 삼는다. 같은 규칙으로 머리글 아래 연속된 데이터의 마지막 행(0부터)을 찾는다 — 예전엔
 * 범위를 적힌 그대로 써서 머리글 아래 행이 0개로 계산돼 필터 값 목록이 비었다(실사용 파일 "인터페이스 정의서").
 */
export function contiguousDataEnd(sheetXml: string, headerRow: number, c1: number, c2: number): number {
  const rowsWithData = new Set<number>();
  for (const cell of sheetXml.matchAll(/<c\b[^>]*\br="([A-Z]{1,3})(\d+)"[^>]*?(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const inner = cell[3];
    if (!inner || !/<(?:v|is|f)\b[^>]*>(?!<\/)/.test(inner)) continue; // 서식만 있는 빈 칸은 데이터가 아니다
    const col = columnIndexOf(cell[1]);
    if (col < c1 || col > c2) continue;
    rowsWithData.add(Number(cell[2]) - 1);
  }
  let last = headerRow;
  while (rowsWithData.has(last + 1)) last += 1;
  return last;
}

/** 시트 이름 순서(편집기 시트 자리 번호)대로 파일의 자동 필터 */
export function readAutoFilters(xlsx: Uint8Array, sheetNames: readonly string[]): Map<number, SheetFilter> {
  const entries: Record<string, Uint8Array> = unzipSync(xlsx);
  const paths = sheetPathsByName(entries);
  const out = new Map<number, SheetFilter>();
  sheetNames.forEach((name, index) => {
    const path = paths.get(name);
    const bytes = path ? entries[path] : undefined;
    // 시트 자동 필터가 없으면 표(표 서식) 안의 필터 — 예전엔 이걸 몰라 표 머리글의 필터 단추가 안 보였다
    const filter = bytes ? (parseAutoFilter(strFromU8(bytes)) ?? (path ? tableFilterOf(entries, path) : null)) : null;
    if (filter) out.set(index, filter);
  });
  return out;
}

/** 저장 파일의 <autoFilter ref=…/>에 조건을 붙인다(ExcelJS가 범위만 쓰므로) */
export function writeFilterCriteria(xlsx: Uint8Array, filtersBySheetName: ReadonlyMap<string, SheetFilter>): Uint8Array {
  const withCriteria = [...filtersBySheetName].filter(([, f]) => !f.table && Object.keys(f.criteria).length > 0);
  if (withCriteria.length === 0) return xlsx;
  const entries: Record<string, Uint8Array> = unzipSync(xlsx);
  const paths = sheetPathsByName(entries);
  for (const [name, filter] of withCriteria) {
    const path = paths.get(name);
    if (!path || !entries[path]) continue;
    const columns = Object.entries(filter.criteria)
      .map(([col, values]) => {
        const blank = values.includes('') ? ' blank="1"' : '';
        const items = values.filter((v) => v !== '').map((v) => `<filter val="${escapeXml(v)}"/>`).join('');
        return `<filterColumn colId="${Number(col) - filter.c1}"><filters${blank}>${items}</filters></filterColumn>`;
      })
      .join('');
    const xml = strFromU8(entries[path]).replace(/<autoFilter\b([^>]*?)\/>/, `<autoFilter$1>${columns}</autoFilter>`);
    entries[path] = strToU8(xml);
  }
  return zipSync(entries);
}
