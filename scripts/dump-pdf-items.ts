// 실제 PDF 쪽의 글 조각(위치·크기·글꼴·굵기·태그 역할)을 그대로 뽑는다 — 제목 판별 조정용
// 사용: npx tsx scripts/dump-pdf-items.ts <pdf> <쪽>
import fs from 'node:fs';
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs';
import { isBoldFontName } from '../src/client/pdf/textStructure.js';

const [file, pageNo] = process.argv.slice(2);
const doc = await pdfjs.getDocument({ data: new Uint8Array(fs.readFileSync(file)), cMapUrl: 'node_modules/pdfjs-dist/cmaps/', cMapPacked: true, standardFontDataUrl: 'node_modules/pdfjs-dist/standard_fonts/', isEvalSupported: false }).promise;
const p = await doc.getPage(Number(pageNo));
await p.getOperatorList();
const content = await p.getTextContent({ includeMarkedContent: true });
const roleOf = new Map<string, string>();
const tree = await p.getStructTree().catch(() => null);
const walk = (node: unknown, role: string | null) => {
  const n = node as { role?: string; type?: string; id?: string; children?: unknown[] } | null;
  if (!n) return;
  if (n.type === 'content' && n.id && role) roleOf.set(n.id, role);
  const next = n.role && !['Root', 'Document', 'Part', 'Sect', 'Div'].includes(n.role) ? n.role : role;
  for (const c of n.children ?? []) walk(c, next);
};
walk(tree, null);
const stack: string[] = [];
const fonts = new Map<string, string>();
for (const item of content.items) {
  const it = item as { type?: string; id?: string; str?: string; fontName?: string; transform?: number[]; width?: number; height?: number };
  if (it.type === 'beginMarkedContentProps' || it.type === 'beginMarkedContent') stack.push(it.id ?? '');
  else if (it.type === 'endMarkedContent') stack.pop();
  else if (typeof it.str === 'string' && it.str.trim()) {
    let role: string | null = null;
    for (let k = stack.length - 1; k >= 0 && !role; k -= 1) role = roleOf.get(stack[k]) ?? null;
    let name = fonts.get(it.fontName!);
    let bold = false;
    try {
      const f = p.commonObjs.get(it.fontName!) as { name?: string; bold?: boolean; black?: boolean };
      name = f?.name ?? '?';
      bold = !!(f?.bold || f?.black) || isBoldFontName(f?.name);
    } catch {
      name = '?';
    }
    const t = it.transform!;
    const cps = [...it.str].slice(0, 3).map((c) => c.codePointAt(0)!.toString(16)).join(' ');
    console.log(`y${t[5].toFixed(1).padStart(6)} x${t[4].toFixed(1).padStart(6)} w${it.width!.toFixed(0).padStart(4)} h${it.height!.toFixed(1)} ${bold ? 'B' : '-'} ${(role ?? '-').padEnd(5)} ${String(name).slice(0, 28).padEnd(28)} [${cps}] ${it.str.slice(0, 50)}`);
  }
}
console.log('page size', p.getViewport({ scale: 1 }).width, p.getViewport({ scale: 1 }).height);
