import { createHash } from 'node:crypto';
import { XMLBuilder, XMLParser } from 'fast-xml-parser';

/**
 * docx XML을 읽고(가져오기) 원본 조각을 그대로 되돌려 쓰는(내보내기) 데 함께 쓰는 도구.
 * fast-xml-parser를 preserveOrder 모드로 쓴다 — OOXML은 같은 태그(w:p, w:tbl 등)가 반복되며 형제 순서가 의미를 가진다.
 */
const XML_OPTIONS = { ignoreAttributes: false, attributeNamePrefix: '@_', preserveOrder: true, textNodeName: '#text', trimValues: false } as const;
export const parser = new XMLParser(XML_OPTIONS);
const builder = new XMLBuilder({ ...XML_OPTIONS, suppressEmptyNode: true });

export type XmlNode = Record<string, unknown>;

export function tagOf(node: XmlNode): string | undefined {
  return Object.keys(node).find((key) => key !== ':@');
}

export function attrsOf(node: XmlNode): Record<string, string> {
  return (node[':@'] as Record<string, string> | undefined) ?? {};
}

export function childrenOf(node: XmlNode): XmlNode[] {
  const tag = tagOf(node);
  if (!tag) return [];
  const value = node[tag];
  return Array.isArray(value) ? (value as XmlNode[]) : [];
}

export function findAll(list: readonly XmlNode[], tag: string): XmlNode[] {
  return list.filter((node) => tagOf(node) === tag);
}

export function findFirst(list: readonly XmlNode[], tag: string): XmlNode | undefined {
  return list.find((node) => tagOf(node) === tag);
}

/** 깊이 우선으로 첫 번째 tag 요소 */
export function findDeep(list: readonly XmlNode[], tag: string): XmlNode | undefined {
  for (const node of list) {
    if (tagOf(node) === tag) return node;
    const inner = findDeep(childrenOf(node), tag);
    if (inner) return inner;
  }
  return undefined;
}

export function textOf(list: readonly XmlNode[]): string {
  return list
    .filter((node) => '#text' in node)
    .map((node) => String((node as { '#text': unknown })['#text']))
    .join('');
}

/** 요소 안의 모든 글자(w:t·m:t·w:delText, 탭은 \t) — 그대로 보존하는 조각의 화면 글자에 쓴다. */
export function deepTextOf(list: readonly XmlNode[]): string {
  let out = '';
  for (const node of list) {
    const tag = tagOf(node);
    if (tag === 'w:t' || tag === 'm:t' || tag === 'w:delText' || tag === 'w:instrText') {
      if (tag !== 'w:instrText') out += textOf(childrenOf(node));
    } else if (tag === 'w:tab') out += '\t';
    else if (tag) out += deepTextOf(childrenOf(node));
  }
  return out;
}

/** 파싱된 요소 하나를 다시 XML 문자열로 — 모델링하지 않는 원본 조각을 그대로 되돌려 쓰기 위함. 같은 입력이면 늘 같은 문자열이다. */
export function xmlOf(node: XmlNode): string {
  return builder.build([node]) as string;
}

/** w:b·w:i 같은 켜기/끄기 속성 — 요소가 있고 val이 0/false/off가 아니면 켜짐. 요소가 없으면 undefined. */
export function onOff(list: readonly XmlNode[], tag: string): boolean | undefined {
  const el = findFirst(list, tag);
  if (!el) return undefined;
  const val = attrsOf(el)['@_w:val'];
  return !(val === '0' || val === 'false' || val === 'off');
}

/** CT_ 스키마가 정한 자식 순서 — 모델링한 값과 원본 그대로의 나머지를 합쳐 쓸 때 이 순서로 정렬한다(Word는 순서가 틀리면 파일을 못 연다). */
export const PR_ORDER: Readonly<Record<string, readonly string[]>> = {
  'w:pPr': ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs',
    'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid',
    'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl',
    'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'],
  'w:rPr': ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof',
    'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
    'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath', 'rPrChange'],
  'w:tblPr': ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize', 'tblW', 'jc', 'tblCellSpacing', 'tblInd',
    'tblBorders', 'shd', 'tblLayout', 'tblCellMar', 'tblLook', 'tblCaption', 'tblDescription', 'tblPrChange'],
  'w:trPr': ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight', 'tblHeader', 'tblCellSpacing', 'jc', 'hidden',
    'ins', 'del', 'trPrChange'],
  'w:tcPr': ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar', 'textDirection', 'tcFitText', 'vAlign', 'hideMark',
    'headers', 'cellIns', 'cellDel', 'cellMerge', 'tcPrChange'],
};

/**
 * 보존 조각 참조 — 원본 XML은 서버의 조각 저장소(해시 → XML, 방이 열려 있는 동안 유지)에만 두고 문서 노드에는 "태그@해시"만 싣는다.
 * 그래서 다른 참여자가 Yjs로 임의 XML을 보낼 길이 없고(저장소에 없는 해시는 저장 때 버려진다), 지웠다가 되돌린 조각도 저장소에 남아 있어
 * 다시 저장된다. unique: 내용이 같아도 따로 세야 하는 조각(구역 나눔)에 붙이는 순번("#3").
 */
export type FragmentStore = Map<string, string>;

function escapeAttrValue(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** 감싸개(w:hyperlink·w:ins 등)의 속성 XML — 이름이 XML 이름 꼴인 것만(다른 참여자가 보낸 값으로 XML을 깨지 못하게) */
export function wrapAttrsXml(attrs: Readonly<Record<string, unknown>> | null | undefined): string {
  return Object.entries(attrs ?? {})
    .filter(([name]) => /^[A-Za-z][\w.-]*(:[A-Za-z][\w.-]*)?$/.test(name))
    .map(([name, value]) => ` ${name}="${escapeAttrValue(String(value))}"`)
    .join('');
}

/**
 * 변경 추적 삽입(w:ins·w:moveTo) 감싸개 표지 — 가져오기 때 원본에 있던 감싸개(작성자·날짜·id 그대로)를 이 모양(빈 요소)으로 조각
 * 저장소에 넣어 두고, 저장 때 이것과 똑같은 감싸개만 쓴다. 다른 참여자가 지어낸 작성자 감싸개는 저장소에 없으니 빼고 글자만 쓴다.
 */
export const TRACKED_WRAP_TAGS: ReadonlySet<string> = new Set(['w:ins', 'w:moveTo']);
export const trackedWrapKey = (tag: string, attrs: Readonly<Record<string, unknown>> | null | undefined): string => `<${tag}${wrapAttrsXml(attrs)}/>`;

/** 조각 해시(참조의 @ 뒤) — 내용에서 바로 나오므로 저장해 둔 조각을 다시 읽을 때 맞는지 확인할 수 있다 */
export const fragmentHash = (xml: string): string => createHash('sha256').update(xml).digest('base64url').slice(0, 20);

export function fragmentRef(store: FragmentStore, tag: string, xml: string, unique?: number): string {
  const hash = fragmentHash(xml);
  store.set(hash, xml);
  return `${tag}@${hash}${unique === undefined ? '' : `#${unique}`}`;
}

/** 참조의 태그("w:numPr@…" → "w:numPr") — 브라우저 쪽(wordCommands)도 같은 모양을 쓴다 */
export const refTagOf = (ref: string): string => ref.slice(0, Math.max(0, ref.indexOf('@')));
export const refHashOf = (ref: string): string => ref.slice(ref.indexOf('@') + 1).replace(/#.*$/, '');

/** 속성 묶음(w:pPr 등)의 자식 중 우리가 모델링하지 않는 것들 — 원본 그대로 저장소에 두고 참조 목록을 노드 속성에 담는다. */
export function extrasOf(children: readonly XmlNode[], modeled: ReadonlySet<string>, store: FragmentStore): string[] | null {
  const extras = children.filter((node) => {
    const tag = tagOf(node);
    return tag !== undefined && tag !== '#text' && !modeled.has(tag);
  });
  return extras.length > 0 ? extras.map((node) => fragmentRef(store, tagOf(node)!, xmlOf(node))) : null;
}

const tagNameOf = (xml: string): string => /^<([\w:.-]+)/.exec(xml)?.[1] ?? '';

/**
 * 모델링한 값(우리가 새로 만든 XML 조각)과 보존한 원본 조각을 합쳐 스키마 순서로 정렬한 속성 묶음 하나를 만든다.
 * resolve: 참조 → 원본 XML(저장소에 없는 참조는 undefined — 버린다).
 */
export function prXml(wrapper: keyof typeof PR_ORDER, parts: readonly string[], extras: unknown, resolve: (ref: string) => string | undefined): string {
  const order = PR_ORDER[wrapper];
  const rank = (xml: string) => {
    const i = order.indexOf(tagNameOf(xml).replace(/^w:/, ''));
    return i < 0 ? order.length : i;
  };
  const modeled = parts.filter(Boolean);
  const modeledTags = new Set(modeled.map(tagNameOf));
  // 편집으로 새로 생긴 값(예: 끄기였던 굵게를 켬)과 같은 요소의 원본 조각은 뺀다 — 같은 요소가 두 번 나오면 Word가 파일을 못 연다
  const raw = Array.isArray(extras)
    ? extras
        .map((ref) => (typeof ref === 'string' ? resolve(ref) : undefined))
        .filter((x): x is string => typeof x === 'string' && !modeledTags.has(tagNameOf(x)))
    : [];
  const all = [...modeled, ...raw].map((xml, index) => ({ xml, index, rank: rank(xml) }));
  all.sort((a, b) => a.rank - b.rank || a.index - b.index);
  return all.length > 0 ? `<${wrapper}>${all.map((p) => p.xml).join('')}</${wrapper}>` : '';
}
