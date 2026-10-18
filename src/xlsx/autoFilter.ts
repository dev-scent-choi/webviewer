/**
 * 시트 자동 필터(<autoFilter>)를 파일에서 읽고, 저장 파일에 조건(<filterColumn>)까지 다시 쓴다.
 * ExcelJS는 자동 필터 범위만 다루고 조건은 버리므로, 조건은 쓰기가 끝난 zip에서 직접 붙인다.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { parseRangeFormula } from './chartModel.js';
import { sheetPathsByName } from './drawingPassthrough.js';
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
  return { r1: range.r1, c1: range.c1, r2: range.r2, c2: range.c2, criteria };
}

/** 시트 이름 순서(편집기 시트 자리 번호)대로 파일의 자동 필터 */
export function readAutoFilters(xlsx: Uint8Array, sheetNames: readonly string[]): Map<number, SheetFilter> {
  const entries: Record<string, Uint8Array> = unzipSync(xlsx);
  const paths = sheetPathsByName(entries);
  const out = new Map<number, SheetFilter>();
  sheetNames.forEach((name, index) => {
    const path = paths.get(name);
    const bytes = path ? entries[path] : undefined;
    const filter = bytes ? parseAutoFilter(strFromU8(bytes)) : null;
    if (filter) out.set(index, filter);
  });
  return out;
}

/** 저장 파일의 <autoFilter ref=…/>에 조건을 붙인다(ExcelJS가 범위만 쓰므로) */
export function writeFilterCriteria(xlsx: Uint8Array, filtersBySheetName: ReadonlyMap<string, SheetFilter>): Uint8Array {
  const withCriteria = [...filtersBySheetName].filter(([, f]) => Object.keys(f.criteria).length > 0);
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
