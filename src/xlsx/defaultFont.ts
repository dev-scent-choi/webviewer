/**
 * ExcelJS는 저장할 때 styles.xml을 새로 쓰면서 통합 문서의 "표준" 스타일 글꼴을 자기 기본값(Calibri 11)으로
 * 바꿔 버린다 — 한국어 문서(표준 글꼴 돋움·맑은 고딕)를 편집기로 한 번 저장하면 엑셀에서 열 너비 단위가
 * 바뀌어 모든 열이 좁아지고, 서식 없는 칸의 글꼴도 달라졌다. 저장 결과의 표준 글꼴을 원본 것으로 되돌린다.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { normalFontOfStylesXml } from './sheetGeometry.js';

export function restoreDefaultFont(outputXlsx: Uint8Array, originalXlsx: Uint8Array): Uint8Array {
  const pick = (bytes: Uint8Array) => unzipSync(bytes, { filter: (file) => file.name === 'xl/styles.xml' })['xl/styles.xml'];
  const originalStyles = pick(originalXlsx);
  if (!originalStyles) return outputXlsx;
  const original = normalFontOfStylesXml(strFromU8(originalStyles));
  if (!original.xml || !original.name) return outputXlsx;
  const entries = unzipSync(outputXlsx);
  const stylesBytes = entries['xl/styles.xml'];
  if (!stylesBytes) return outputXlsx;
  const styles = strFromU8(stylesBytes);
  const output = normalFontOfStylesXml(styles);
  if (!output.xml || output.index === undefined) return outputXlsx;
  if (output.name === original.name && output.size === original.size) return outputXlsx;
  // 원본 글꼴 요소의 접두사(있다면)를 출력 파일 쪽 표기에 맞춘다 — ExcelJS 출력은 접두사가 없다
  const replacement = original.xml.replace(/<(\/?)\w+:/g, '<$1');
  const fontsMatch = /(<(?:\w+:)?fonts\b[^>]*>)([\s\S]*?)(<\/(?:\w+:)?fonts>)/.exec(styles);
  if (!fontsMatch) return outputXlsx;
  let seen = -1;
  const fontsBody = fontsMatch[2].replace(/<(?:\w+:)?font\b[^>]*?(?:\/>|>[\s\S]*?<\/(?:\w+:)?font>)/g, (font) => {
    seen += 1;
    return seen === output.index ? replacement : font;
  });
  entries['xl/styles.xml'] = strToU8(styles.replace(fontsMatch[0], fontsMatch[1] + fontsBody + fontsMatch[3]));
  return zipSync(entries);
}
