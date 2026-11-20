import { strFromU8, strToU8, zipSync } from 'fflate';
import { guardedUnzipSync } from '../zipGuard.js';

/**
 * LibreOffice로 PDF를 만들기 전에 pptx 사본에만 적용하는 전처리(원본·저장 파일은 그대로) — 보기 화면과 편집 화면의
 * 표 조각이 이 PDF를 쓴다.
 *
 * 표 머리 줄 굵게: PowerPoint 기본 표 스타일은 "머리글 행"(tblPr firstRow="1")의 글자를 굵게 그리는데(스타일의
 * <a:firstRow><a:tcTxStyle b="on">), LibreOffice(26.2)는 이 굵기를 적용하지 않아 머리 줄이 가는 글씨로 나왔다
 * (품질 점검 P-v2, p04·a2 2장). 그 표의 첫 줄 글자에 굵게(b="1")를 직접 적어 같은 모양을 만든다. 글자에 이미
 * b가 적혀 있으면(b="0" 포함) 그대로 둔다. 스타일 정의가 tableStyles.xml에 있으면 그 정의를 따르고, 없으면
 * PowerPoint 내장 스타일 — 내장 스타일은 서식 없는 두 개를 빼고 모두 머리 줄이 굵다.
 */
const PLAIN_BUILTIN_STYLES = new Set(['{2D5ABB26-0587-4C30-8999-92F81FD0307C}', '{5940675A-B579-460E-94D1-54222C63F5DA}']);

function headerIsBold(styleId: string, tableStylesXml: string): boolean {
  const def = new RegExp(`<a:tblStyle\\b[^>]*styleId="${styleId.replace(/[{}-]/g, (c) => `\\${c}`)}"[\\s\\S]*?</a:tblStyle>`).exec(tableStylesXml)?.[0];
  if (def) return /<a:firstRow>[\s\S]*?<a:tcTxStyle\b[^>]*\bb="on"/.test(def.match(/<a:firstRow>[\s\S]*?<\/a:firstRow>/)?.[0] ?? '');
  return !PLAIN_BUILTIN_STYLES.has(styleId);
}

/** 표 하나(<a:tbl>…</a:tbl>)의 첫 줄 글자를 굵게 — 머리글 행이 켜져 있고 스타일이 굵게일 때만 */
export function boldTableHeaderXml(tblXml: string, tableStylesXml: string): string {
  const tblPr = /<a:tblPr\b[^>]*>/.exec(tblXml)?.[0] ?? '';
  if (!/\bfirstRow="(1|true)"/.test(tblPr)) return tblXml;
  const styleId = /<a:tableStyleId>([^<]+)<\/a:tableStyleId>/.exec(tblXml)?.[1];
  if (!styleId || !headerIsBold(styleId.trim(), tableStylesXml)) return tblXml;
  return tblXml.replace(/<a:tr\b[\s\S]*?<\/a:tr>/, (row) =>
    row
      .replace(/<a:rPr\b(?![^>]*\bb=")/g, '<a:rPr b="1"')
      .replace(/<a:r>(?!\s*<a:rPr\b)/g, '<a:r><a:rPr b="1"/>'),
  );
}

export function boldPptxTableHeaders(pptxBytes: Uint8Array): Uint8Array {
  const entries = guardedUnzipSync(pptxBytes, 'boldPptxTableHeaders');
  const tableStyles = entries['ppt/tableStyles.xml'] ? strFromU8(entries['ppt/tableStyles.xml']) : '';
  let changed = false;
  const updated: Record<string, Uint8Array> = { ...entries };
  for (const [name, bytes] of Object.entries(entries)) {
    if (!/^ppt\/slides\/slide\d+\.xml$/.test(name)) continue;
    const xml = strFromU8(bytes);
    if (!xml.includes('<a:tbl>')) continue;
    const next = xml.replace(/<a:tbl>[\s\S]*?<\/a:tbl>/g, (tbl) => boldTableHeaderXml(tbl, tableStyles));
    if (next !== xml) {
      updated[name] = strToU8(next);
      changed = true;
    }
  }
  return changed ? zipSync(updated) : pptxBytes;
}
