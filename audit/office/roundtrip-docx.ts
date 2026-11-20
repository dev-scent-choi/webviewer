// docx 열기 → (고치지 않고) 저장 → 다시 열기 왕복 점검: 글자·구조(노드 종류별 개수)·보존 조각(그림·필드 등)이 그대로인지,
// 저장한 document.xml이 올바른 XML인지 본다. 저장본은 out 폴더에 남겨 LibreOffice·Word로 열어 볼 수 있다.
// 사용: npx tsx audit/office/roundtrip-docx.ts [corpus 폴더=audit/office/corpus/docx] [출력 폴더=docs/quality-audit/out/roundtrip]
import fs from 'node:fs';
import path from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { XMLValidator } from 'fast-xml-parser';
import type { Node as PMNode } from 'prosemirror-model';
import { exportDocx } from '../../src/docx/export.js';
import { importDocx } from '../../src/docx/import.js';

const [dir = 'audit/office/corpus/docx', outDir = 'docs/quality-audit/out/roundtrip'] = process.argv.slice(2);
fs.mkdirSync(outDir, { recursive: true });

function shapeOf(doc: PMNode) {
  const counts: Record<string, number> = {};
  const raw: string[] = [];
  doc.descendants((node) => {
    const key = node.type.name === 'ooxml' ? `ooxml:${node.attrs.kind as string}` : node.type.name;
    counts[key] = (counts[key] ?? 0) + 1;
    if (node.type.name === 'ooxml' || node.type.name === 'ooxml_block') raw.push(String(node.attrs.xml));
    for (const k of ['pPrExtra', 'tcPrExtra', 'trPrExtra', 'tblPrExtra']) if (Array.isArray(node.attrs[k])) raw.push(...(node.attrs[k] as string[]));
    for (const mark of node.marks) {
      counts[`mark:${mark.type.name}`] = (counts[`mark:${mark.type.name}`] ?? 0) + 1;
      if (mark.type.name === 'runExtra' && Array.isArray(mark.attrs.xml)) raw.push(...(mark.attrs.xml as string[]));
    }
  });
  return { counts, raw, text: doc.textBetween(0, doc.content.size, '\n', (leaf) => leaf.type.spec.leafText?.(leaf) ?? '') };
}

let failed = 0;
for (const name of fs.readdirSync(dir).filter((f) => f.endsWith('.docx') && !f.startsWith('~$')).sort()) {
  const bytes = new Uint8Array(fs.readFileSync(path.join(dir, name)));
  const first = importDocx(bytes);
  const saved = exportDocx(bytes, first.doc, null, first.comments);
  fs.writeFileSync(path.join(outDir, name), saved);
  const second = importDocx(saved);
  const a = shapeOf(first.doc);
  const b = shapeOf(second.doc);
  const problems: string[] = [];
  if (a.text !== b.text) problems.push('글자 다름');
  for (const key of new Set([...Object.keys(a.counts), ...Object.keys(b.counts)])) {
    if (a.counts[key] !== b.counts[key]) problems.push(`${key} ${a.counts[key] ?? 0}→${b.counts[key] ?? 0}`);
  }
  const lostRaw = a.raw.filter((x) => !b.raw.includes(x));
  if (lostRaw.length) problems.push(`보존 조각 ${lostRaw.length}개 사라짐`);
  const xml = strFromU8(unzipSync(saved)['word/document.xml']);
  // 원본과 저장본의 주요 요소 개수 — 번호·구역·그림·글상자·각주·표 스타일·탭·필드·링크·책갈피가 그대로 남았는지
  const originalXml = strFromU8(unzipSync(bytes)['word/document.xml']);
  for (const tag of ['w:numPr', 'w:sectPr', 'w:drawing', 'w:txbxContent', 'w:footnoteReference', 'w:tblStyle', 'w:tabs', 'w:tab', 'w:fldChar', 'w:fldSimple', 'w:hyperlink', 'w:bookmarkStart', 'w:pBdr', 'w:tblHeader', 'w:rStyle']) {
    const count = (s: string) => (s.match(new RegExp(`<${tag}[\\s/>]`, 'g')) ?? []).length;
    if (count(originalXml) !== count(xml)) problems.push(`<${tag}> ${count(originalXml)}→${count(xml)}`);
  }
  const valid = XMLValidator.validate(xml);
  if (valid !== true) problems.push(`XML 오류 ${JSON.stringify(valid).slice(0, 120)}`);
  // 같은 파일을 한 번 더 저장해도 document.xml이 같아야 한다(저장할 때마다 조금씩 바뀌지 않는다)
  const again = strFromU8(unzipSync(exportDocx(saved, second.doc, null, second.comments))['word/document.xml']);
  if (again !== xml) problems.push('두 번째 저장 결과가 다름');
  const summary = Object.entries(a.counts).filter(([k]) => /ooxml|table$|columns|tab$|mark:link|heading/.test(k)).map(([k, v]) => `${k}=${v}`).join(' ');
  console.log(`${problems.length ? 'FAIL' : 'ok  '} ${name}  ${summary}${problems.length ? `  ← ${problems.join(', ')}` : ''}`);
  if (problems.length) failed += 1;
}
process.exitCode = failed ? 1 : 0;
