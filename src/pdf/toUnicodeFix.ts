/**
 * pdf-lib가 통째로 넣은 글꼴(subset:false)의 글자 대응표(ToUnicode)를 바로잡는다.
 *
 * pdf-lib는 글리프마다 fontkit의 codePoints를 그대로 쓰는데, Pretendard처럼 한 글리프를 여러 유니코드가 같이 쓰는 글꼴에서는
 * 호환 문자가 대신 들어가 복사·찾기·변환에서 다른 글자가 나왔다(실측: "-" → "㏄", "—" → "㏆"). 글리프마다 그 글리프로 가는
 * 유니코드 중 가장 작은 값(기본 문자 쪽)을 골라 대응표를 새로 쓴다. 저장 직전(글꼴이 문서에 들어간 뒤)에 부른다.
 */
import { PDFDict, PDFName, type PDFDocument, type PDFFont } from 'pdf-lib';

interface FontkitFont {
  readonly characterSet: number[];
  glyphForCodePoint(cp: number): { id: number };
}

const cache = new WeakMap<object, string>();

function utf16Hex(cp: number): string {
  if (cp <= 0xffff) return cp.toString(16).padStart(4, '0');
  const v = cp - 0x10000;
  return ((0xd800 | (v >> 10)).toString(16) + (0xdc00 | (v & 0x3ff)).toString(16)).padStart(8, '0');
}

export function buildToUnicode(font: FontkitFont): string {
  const cached = cache.get(font);
  if (cached) return cached;
  const best = new Map<number, number>();
  for (const cp of font.characterSet) {
    const gid = font.glyphForCodePoint(cp).id;
    if (gid === 0) continue;
    const prev = best.get(gid);
    if (prev === undefined || cp < prev) best.set(gid, cp);
  }
  const entries = [...best.entries()].sort((a, b) => a[0] - b[0]);
  const chunks: string[] = [];
  for (let i = 0; i < entries.length; i += 100) {
    const part = entries.slice(i, i + 100);
    chunks.push(`${part.length} beginbfchar\n${part.map(([gid, cp]) => `<${gid.toString(16).padStart(4, '0')}> <${utf16Hex(cp)}>`).join('\n')}\nendbfchar`);
  }
  const text = `/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <ffff>\nendcodespacerange\n${chunks.join('\n')}\nendcmap\nCMapName currentdict /CMapResource defineresource pop\nend\nend\n`;
  cache.set(font, text);
  return text;
}

/** 문서에 들어간 그 글꼴의 ToUnicode를 바꿔 끼운다 */
export async function fixToUnicode(doc: PDFDocument, font: PDFFont): Promise<void> {
  const fk = (font as unknown as { embedder?: { font?: FontkitFont } }).embedder?.font;
  if (!fk?.characterSet) return;
  await doc.flush();
  const dict = doc.context.lookup(font.ref);
  if (!(dict instanceof PDFDict)) return;
  dict.set(PDFName.of('ToUnicode'), doc.context.register(doc.context.flateStream(buildToUnicode(fk))));
}
