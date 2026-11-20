import JSZip from 'jszip';

/**
 * LibreOffice로 PDF를 만들기 전에 pptx 사본에만 적용하는 전처리(원본·저장 파일은 그대로) — 보기 화면과 편집 화면의
 * 표 조각이 이 PDF를 쓴다.
 *
 * 표 머리 줄 굵게: PowerPoint 기본 표 스타일은 "머리글 행"(tblPr firstRow="1")의 글자를 굵게 그리는데(스타일의
 * <a:firstRow><a:tcTxStyle b="on">), LibreOffice(26.2)는 이 굵기를 적용하지 않아 머리 줄이 가는 글씨로 나왔다
 * (품질 점검 P-v2, p04·a2 2장). 그 표의 첫 줄 글자에 굵게(b="1")를 직접 적어 같은 모양을 만든다. 글자에 이미
 * b가 적혀 있으면(b="0" 포함) 그대로 둔다. 스타일 정의가 tableStyles.xml에 있으면 그 정의를 따르고, 없으면
 * PowerPoint 내장 스타일 — 내장 스타일은 서식 없는 두 개를 빼고 모두 머리 줄이 굵다.
 *
 * 안전: 스타일 id는 파일 값이라 정규식에 넣지 않는다(예전엔 넣어서 표 하나로 서버가 수십 초 멈췄다 — 검토 C1).
 * GUID 모양({8-4-4-4-12 16진수})만 받고 평범한 문자열 검색으로 찾는다.
 */
const PLAIN_BUILTIN_STYLES = new Set(['{2D5ABB26-0587-4C30-8999-92F81FD0307C}', '{5940675A-B579-460E-94D1-54222C63F5DA}']);
const GUID = /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/;

function headerIsBold(styleId: string, tableStylesXml: string): boolean {
  if (!GUID.test(styleId)) return false;
  const at = tableStylesXml.indexOf(`styleId="${styleId}"`);
  if (at < 0) return !PLAIN_BUILTIN_STYLES.has(styleId.toUpperCase());
  const end = tableStylesXml.indexOf('</a:tblStyle>', at);
  const def = tableStylesXml.slice(at, end < 0 ? undefined : end);
  const first = def.indexOf('<a:firstRow>');
  if (first < 0) return false;
  const firstEnd = def.indexOf('</a:firstRow>', first);
  return /<a:tcTxStyle\b[^>]*\bb="on"/.test(def.slice(first, firstEnd < 0 ? undefined : firstEnd));
}

/** 표 하나(<a:tbl>…</a:tbl>)의 첫 줄 글자를 굵게 — 머리글 행이 켜져 있고 스타일이 굵게일 때만 */
export function boldTableHeaderXml(tblXml: string, tableStylesXml: string): string {
  const tblPr = /<a:tblPr\b[^>]*>/.exec(tblXml)?.[0] ?? '';
  if (!/\bfirstRow="(1|true)"/.test(tblPr)) return tblXml;
  const styleId = /<a:tableStyleId>([^<]{1,64})<\/a:tableStyleId>/.exec(tblXml)?.[1];
  if (!styleId || !headerIsBold(styleId.trim(), tableStylesXml)) return tblXml;
  return tblXml.replace(/<a:tr\b[\s\S]*?<\/a:tr>/, (row) =>
    row.replace(/<a:rPr\b(?![^>]*\bb=")/g, '<a:rPr b="1"').replace(/<a:r>(?!\s*<a:rPr\b)/g, '<a:r><a:rPr b="1"/>'),
  );
}

/**
 * 표가 있는 슬라이드만 고쳐 다시 묶는다 — 비동기(JSZip, 조각마다 이벤트 루프에 양보), 고치지 않은 파트는 원래 압축
 * 바이트를 그대로 옮긴다(이미 압축된 큰 미디어를 다시 압축하느라 서버가 수 초 멈추던 문제 — 검토 I1). 고칠 게
 * 없으면 원본 바이트를 그대로 돌려준다.
 */
export async function boldPptxTableHeaders(pptxBytes: Uint8Array): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(pptxBytes);
  const stylesFile = zip.file('ppt/tableStyles.xml');
  const tableStyles = stylesFile ? await stylesFile.async('string') : '';
  let changed = false;
  for (const file of zip.file(/^ppt\/slides\/slide\d+\.xml$/)) {
    const xml = await file.async('string');
    if (!xml.includes('<a:tbl>')) continue;
    const next = xml.replace(/<a:tbl>[\s\S]*?<\/a:tbl>/g, (tbl) => boldTableHeaderXml(tbl, tableStyles));
    if (next !== xml) {
      zip.file(file.name, next);
      changed = true;
    }
  }
  if (!changed) return pptxBytes;
  // JSZip은 원래 압축 방식과 내보낼 압축 방식이 같은 파트를 다시 압축하지 않고 원래 바이트를 그대로 쓴다 — 무압축(STORE)으로
  // 들어 있던 파트(대개 큰 미디어)도 STORE로 내보내야 다시 압축되지 않는다.
  // shortcut: JSZip 내부 필드(_data.compression.magic, options.compression)를 읽고 쓴다 — JSZip을 올릴 때 이 시험(loPreprocess 시간)이 깨지면 다시 볼 것
  zip.forEach((_path, file) => {
    const inner = file as unknown as { _data?: { compression?: { magic?: string } }; options: { compression?: string } };
    if (inner._data?.compression?.magic === '\x00\x00') inner.options.compression = 'STORE';
  });
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}
