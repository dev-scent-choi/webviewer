import { XMLBuilder, XMLParser } from 'fast-xml-parser';
import { partialLevelFrom } from './inheritance.js';
import { loadThemeColors, relTargetOf, type ThemeColors } from './theme.js';
import type { Node as PMNode } from 'prosemirror-model';
import { pptxTextSchema as schema } from './textSchema.js';
import { SLIDE_WIDTH_EMU, SLIDE_HEIGHT_EMU, type SlideVisual } from '../collab/pptxShared.js';

export interface ParsedShape {
  readonly shapeId: string;
  readonly x: number | null;
  readonly y: number | null;
  readonly w: number | null;
  readonly h: number | null;
  readonly doc: PMNode;
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  // 기본값(true)이면 <a:t> 텍스트의 앞뒤 공백을 잘라낸다 — pptx 런(run) 경계의 공백(예: "Bold"
  // 다음 런의 " plain")이 서식 구분을 위한 의미 있는 문자라서, 잘리면 텍스트가 눈에 띄게
  // 달라진다(docx/import.ts가 같은 이유로 trimValues: false를 쓰는 것과 동일한 근거).
  trimValues: false,
  // 기본값(true)이면 태그 텍스트가 숫자처럼 보일 때(예: "007", "1.0", "0.10", "+7", "1e3",
  // "0x1A") fast-xml-parser가 조용히 JS 숫자로 바꾸고, 다시 문자열로 꺼낼 때 그 숫자의 "canonical"
  // 형태로 되돌린다("007" -> "7", "1.0" -> "1" 등) — 이 파서가 읽는 태그 텍스트 중 의미 있는 건
  // <a:t>(런 텍스트)뿐이다(위치/크기 같은 속성값은 attributeNamePrefix로 따로 읽혀
  // parseAttributeValue 기본값 false의 영향만 받으므로 이 옵션과 무관하다). 이 문제는 그 도형을
  // 한 번도 편집 안 하면 무해하다(baseline 비교 양쪽이 똑같이 망가지므로 "편집 안 된 도형은
  // 건드리지 않는다" 로직이 여전히 원본을 그대로 둔다) — 하지만 그 도형 안 어느 런이든 한 번
  // 편집되는 순간 <p:txBody> 전체가 이미 망가진 파싱 결과로부터 재조립되어, 숫자처럼 생긴 그
  // 텍스트가 저장 파일에 영구히 다른 값으로 남는다(FINAL REVIEW C3). parseTagValue만 꺼서 태그
  // 텍스트는 항상 원문 문자열 그대로 읽는다.
  parseTagValue: false,
  // <a:r>가 하나뿐이면 배열이 아니라 객체로 파싱되는 fast-xml-parser의 기본 동작을 막는다 —
  // 문단 하나에 런이 하나뿐인 경우가 흔해서, 이 옵션 없이는 매번 "배열인지 객체인지" 분기해야 한다.
  isArray: (name) => ['a:p', 'a:r', 'p:sp'].includes(name),
});

type XmlNode = Record<string, unknown>;

/** 원본 색 노드(<a:solidFill>의 내용)를 XML 문자열로 되돌린다 — 테마 색을 저장에서 그대로 유지하기 위함 */
const colorBuilder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', suppressEmptyNode: true });

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

/** <a:rPr>에서 마크(굵게/기울임/밑줄/색/크기)를 뽑아 그 런의 텍스트에 적용한다. */
function marksFromRPr(rPr: XmlNode | undefined, theme?: ThemeColors): ReturnType<typeof schema.marks.strong.create>[] {
  if (!rPr) return [];
  const marks = [];
  if (rPr['@_b'] === '1') marks.push(schema.marks.strong.create());
  if (rPr['@_i'] === '1') marks.push(schema.marks.em.create());
  if (rPr['@_u'] && rPr['@_u'] !== 'none') marks.push(schema.marks.underline.create());
  const solidFill = rPr['a:solidFill'] as XmlNode | undefined;
  const srgbClr = solidFill?.['a:srgbClr'] as XmlNode | undefined;
  const colorVal = srgbClr?.['@_val'] as string | undefined;
  if (colorVal && Object.keys(srgbClr ?? {}).every((k) => k === '@_val')) {
    marks.push(schema.marks.textColor.create({ color: `#${colorVal}` }));
  } else if (solidFill && typeof solidFill === 'object' && Object.keys(solidFill).length > 0) {
    // 테마 색(schemeClr)·밝기 조정이 붙은 색 등 — 원본 XML을 그대로 보관하고(저장 때 되돌림) 화면엔
    // 테마로 풀어 낸 색을 쓴다. 테마를 모르면(저장 기준선 비교 등) 표시 색 없이 원본만 보관.
    const raw = colorBuilder.build(solidFill) as string;
    marks.push(schema.marks.textColor.create({ color: theme?.resolve(solidFill) ?? '', raw }));
  }
  const sz = rPr['@_sz'] as string | undefined;
  if (sz) marks.push(schema.marks.fontSize.create({ size: Number(sz) / 100 }));
  // 한글 문서는 <a:ea>(동아시아 글꼴)가 실제 보이는 글꼴이라 그쪽을 우선한다. 테마 참조("+mn-lt"
  // 등 '+'로 시작)는 실제 글꼴 이름이 아니므로 마크로 만들지 않는다 — 원본 테마 글꼴이 그대로 쓰인다.
  const typeface =
    ((rPr['a:ea'] as XmlNode | undefined)?.['@_typeface'] as string | undefined) ??
    ((rPr['a:latin'] as XmlNode | undefined)?.['@_typeface'] as string | undefined);
  if (typeface && !typeface.startsWith('+')) marks.push(schema.marks.fontFamily.create({ family: typeface }));
  return marks;
}

const ALIGN_VALUES = new Set(['l', 'ctr', 'r', 'just']);

function paragraphFromXml(pNode: XmlNode, theme?: ThemeColors): PMNode {
  const runs = asArray<XmlNode>(pNode['a:r'] as XmlNode | XmlNode[] | undefined);
  const inline = runs
    .map((run) => {
      const text = (run['a:t'] as string | number | undefined) ?? '';
      const textStr = String(text);
      if (textStr === '') return null;
      return schema.text(textStr, marksFromRPr(run['a:rPr'] as XmlNode | undefined, theme));
    })
    .filter((n): n is NonNullable<typeof n> => n !== null);
  const pPrNode = pNode['a:pPr'] as XmlNode | string | undefined;
  const pPr = typeof pPrNode === 'object' ? pPrNode : undefined; // <a:pPr/>(속성 없음)는 파서가 ''를 준다
  const algn = pPr?.['@_algn'] as string | undefined;
  const lvl = Number(pPr?.['@_lvl'] ?? 0);
  const bullet = partialLevelFrom(pPr).bullet;
  return schema.node(
    'paragraph',
    {
      align: algn && ALIGN_VALUES.has(algn) ? algn : null,
      pPr: pPr ? pPrXmlWithoutAlgn(pPr) : null,
      lvl: Number.isInteger(lvl) && lvl >= 0 && lvl <= 8 ? lvl : 0,
      bu: bullet === null ? 'none' : (bullet ?? null),
    },
    inline,
  );
}

const pPrBuilder = new XMLBuilder({ ignoreAttributes: false, attributeNamePrefix: '@_', suppressEmptyNode: true });

/** 파싱된 <a:pPr>를 algn만 뺀 원래 XML로 되돌린다(정렬은 align 속성이 따로 맡는다). 자식 요소
 * 순서는 파서가 문서 순서대로 담아 두므로 스키마 순서(lnSpc, spcBef, …, buChar, defRPr)가 유지된다. */
function pPrXmlWithoutAlgn(pPr: XmlNode): string {
  const { '@_algn': _algn, ...rest } = pPr;
  return pPrBuilder.build({ 'a:pPr': rest }) as string;
}

/** 저장할 <a:pPr>: 보관해 둔 원본(algn 제외)에 현재 정렬만 다시 얹는다. */
function pPrXmlFor(paragraph: PMNode): string {
  const align = paragraph.attrs.align as string | null;
  const raw = paragraph.attrs.pPr as string | null;
  const algnAttr = align ? ` algn="${align}"` : '';
  if (raw && raw.startsWith('<a:pPr')) return raw.replace(/^<a:pPr/, `<a:pPr${algnAttr}`);
  return align ? `<a:pPr${algnAttr}/>` : '';
}

function emuAttr(xfrmChild: XmlNode | undefined, attr: '@_x' | '@_y' | '@_cx' | '@_cy'): number | null {
  const value = xfrmChild?.[attr];
  return typeof value === 'number' ? value : value !== undefined ? Number(value) : null;
}

/** <p:grpSp>(그룹 도형) 하나의 좌표 변환 — 그룹 안 자식 도형의 <a:xfrm>은 그룹의 "자식 좌표계"
 * (<a:chOff>/<a:chExt>)로 표현되는데, 이는 그룹 자신의 겉보기 위치/크기(<a:off>/<a:ext>, 부모
 * 좌표계 기준)와 다를 수 있다(PowerPoint가 그룹을 늘이거나 줄이면 어긋난다) — 그래서 자식 좌표를
 * 그대로 쓰면 그룹을 옮기거나 늘였을 때 안의 도형들이 엉뚱한 자리에 보인다. 회전(rot=)은 이
 * 파일이 최상위 도형에서도 다루지 않는 것과 같은 이유로 무시한다(완전히 틀린 자리보다 약간 어긋난
 * 자리가 낫다는 이 파일의 기존 단순화 방침과 일관됨). */
interface GroupTransform {
  readonly scaleX: number;
  readonly scaleY: number;
  readonly chOffX: number;
  readonly chOffY: number;
  readonly offX: number;
  readonly offY: number;
}

function groupTransformFromGrpSpPr(grpSpPr: XmlNode | undefined): GroupTransform | null {
  const xfrm = grpSpPr?.['a:xfrm'] as XmlNode | undefined;
  if (!xfrm) return null; // 구조적으로 있어야 하지만(PowerPoint가 항상 씀), 없으면 변환 없이 그대로 둔다.
  const off = xfrm['a:off'] as XmlNode | undefined;
  const ext = xfrm['a:ext'] as XmlNode | undefined;
  const chOff = xfrm['a:chOff'] as XmlNode | undefined;
  const chExt = xfrm['a:chExt'] as XmlNode | undefined;
  const extCx = emuAttr(ext, '@_cx') ?? 0;
  const extCy = emuAttr(ext, '@_cy') ?? 0;
  const chExtCx = emuAttr(chExt, '@_cx') ?? extCx;
  const chExtCy = emuAttr(chExt, '@_cy') ?? extCy;
  return {
    scaleX: chExtCx !== 0 ? extCx / chExtCx : 1,
    scaleY: chExtCy !== 0 ? extCy / chExtCy : 1,
    chOffX: emuAttr(chOff, '@_x') ?? 0,
    chOffY: emuAttr(chOff, '@_y') ?? 0,
    offX: emuAttr(off, '@_x') ?? 0,
    offY: emuAttr(off, '@_y') ?? 0,
  };
}

/** 중첩된 그룹을 거슬러 올라가며(가장 안쪽 그룹부터 바깥쪽 순서로) 로컬 좌표를 슬라이드 절대
 * 좌표로 누적 변환한다. transforms는 바깥쪽부터 안쪽 순서로 쌓이므로(재귀가 바깥 그룹부터 먼저
 * 방문하기 때문) 뒤에서부터(안쪽부터) 적용해야 한다. 네 값 중 하나라도 원본에 없으면(null)
 * 변환하지 않고 그대로 null을 돌려준다 — 이 파일 전체의 "모르면 null, 억지로 계산하지 않는다"
 * 원칙과 일관된다. */
function applyGroupTransforms(
  transforms: readonly GroupTransform[],
  x: number | null,
  y: number | null,
  w: number | null,
  h: number | null,
): { x: number | null; y: number | null; w: number | null; h: number | null } {
  if (x === null || y === null || w === null || h === null) return { x: null, y: null, w: null, h: null };
  let px = x;
  let py = y;
  let pw = w;
  let ph = h;
  for (let i = transforms.length - 1; i >= 0; i--) {
    const t = transforms[i];
    px = (px - t.chOffX) * t.scaleX + t.offX;
    py = (py - t.chOffY) * t.scaleY + t.offY;
    pw = pw * t.scaleX;
    ph = ph * t.scaleY;
  }
  return { x: Math.round(px), y: Math.round(py), w: Math.round(pw), h: Math.round(ph) };
}

function collectShapesRecursive(
  node: XmlNode,
  transforms: readonly GroupTransform[],
  shapes: ParsedShape[],
  theme?: ThemeColors,
): void {
  for (const sp of asArray<XmlNode>(node['p:sp'] as XmlNode | XmlNode[] | undefined)) {
    const txBody = sp['p:txBody'] as XmlNode | undefined;
    if (!txBody) continue; // <p:txBody> 없는 도형(그림 등)은 편집 대상이 아니다 — 건너뛴다.

    const cNvPr = (sp['p:nvSpPr'] as XmlNode)['p:cNvPr'] as XmlNode;
    const shapeId = String(cNvPr['@_id']);

    const xfrm = (sp['p:spPr'] as XmlNode | undefined)?.['a:xfrm'] as XmlNode | undefined;
    const off = xfrm?.['a:off'] as XmlNode | undefined;
    const ext = xfrm?.['a:ext'] as XmlNode | undefined;
    const pos = applyGroupTransforms(transforms, emuAttr(off, '@_x'), emuAttr(off, '@_y'), emuAttr(ext, '@_cx'), emuAttr(ext, '@_cy'));

    const paragraphs = asArray<XmlNode>(txBody['a:p'] as XmlNode | XmlNode[] | undefined).map((p) => paragraphFromXml(p, theme));
    const doc = schema.node(
      'doc',
      null,
      paragraphs.length > 0 ? paragraphs : [schema.node('paragraph', null, [])],
    );

    shapes.push({ shapeId, ...pos, doc });
  }

  for (const grp of asArray<XmlNode>(node['p:grpSp'] as XmlNode | XmlNode[] | undefined)) {
    const groupTransform = groupTransformFromGrpSpPr(grp['p:grpSpPr'] as XmlNode | undefined);
    const nextTransforms = groupTransform ? [...transforms, groupTransform] : transforms;
    collectShapesRecursive(grp, nextTransforms, shapes, theme);
  }
}

/** pptx 슬라이드 XML(p:sld) 안의 각 텍스트 도형(p:sp)을 위치 정보 + ProseMirror 문서로 파싱한다.
 * <p:txBody>가 없는 도형(그림 등)은 편집 대상이 아니므로 결과에서 제외한다. <p:grpSp>(그룹 도형)
 * 안에 중첩된 도형도 재귀로 찾아내 그룹의 좌표 변환을 적용한 절대 좌표로 반환한다 — 그룹 도형은
 * 재질서(저장 시 position은 쓰지 않으므로, hasShape/spliceShapeTxBody/removeShape의 원본 XML
 * 텍스트 기반 매칭은 이 중첩 여부와 무관하게 이미 그대로 동작한다, 이 함수만 읽기 전용으로 몰랐을
 * 뿐이다). */
export function parseSlideShapes(slideXml: string, theme?: ThemeColors): ParsedShape[] {
  const parsed = parser.parse(slideXml) as XmlNode;
  const sld = parsed['p:sld'] as XmlNode;
  const spTree = (sld['p:cSld'] as XmlNode)['p:spTree'] as XmlNode;
  const shapes: ParsedShape[] = [];
  collectShapesRecursive(spTree, [], shapes, theme);
  return shapes;
}

function escapeXmlText(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/** 텍스트 노드의 마크를 <a:rPr>로 되돌린다. schema.marks 순서와 무관하게 항상 같은 속성
 * 순서로 내보내 스냅샷/디버깅이 안정적이게 한다. */
function rPrXmlFor(node: PMNode): string {
  const b = node.marks.some((m) => m.type.name === 'strong') ? ' b="1"' : '';
  const i = node.marks.some((m) => m.type.name === 'em') ? ' i="1"' : '';
  const u = node.marks.some((m) => m.type.name === 'underline') ? ' u="sng"' : '';
  const sizeMark = node.marks.find((m) => m.type.name === 'fontSize');
  const sz = sizeMark ? ` sz="${Math.round((sizeMark.attrs.size as number) * 100)}"` : '';
  const colorMark = node.marks.find((m) => m.type.name === 'textColor');
  const familyMark = node.marks.find((m) => m.type.name === 'fontFamily');
  const attrs = `${b}${i}${u}${sz}`;
  // CT_TextCharacterProperties의 자식 순서: ... solidFill ... latin, ea, cs ... (스키마상 고정)
  // 원본 테마 색(raw)이 있으면 그대로 되돌리고, 리본으로 새로 고른 색이면 RGB로 쓴다.
  const fill = !colorMark
    ? ''
    : colorMark.attrs.raw
      ? `<a:solidFill>${colorMark.attrs.raw as string}</a:solidFill>`
      : /^#[0-9A-Fa-f]{6}$/.test(colorMark.attrs.color as string)
        ? `<a:solidFill><a:srgbClr val="${(colorMark.attrs.color as string).replace('#', '').toUpperCase()}"/></a:solidFill>`
        : '';
  const face = familyMark
    ? (() => {
        const name = escapeXmlAttr(familyMark.attrs.family as string);
        return `<a:latin typeface="${name}"/><a:ea typeface="${name}"/>`;
      })()
    : '';
  const children = fill + face;
  if (!children) return attrs ? `<a:rPr${attrs}/>` : '';
  return `<a:rPr${attrs}>${children}</a:rPr>`;
}

function escapeXmlAttr(text: string): string {
  return escapeXmlText(text).replace(/"/g, '&quot;');
}

export function serializeTxBodyXml(doc: PMNode): string {
  const paragraphsXml: string[] = [];
  doc.forEach((paragraph) => {
    const runsXml: string[] = [];
    paragraph.forEach((textNode) => {
      if (!textNode.isText || !textNode.text) return;
      const rPr = rPrXmlFor(textNode);
      runsXml.push(`<a:r>${rPr}<a:t>${escapeXmlText(textNode.text)}</a:t></a:r>`);
    });
    paragraphsXml.push(`<a:p>${pPrXmlFor(paragraph)}${runsXml.join('')}</a:p>`);
  });
  // CT_TextBody(OOXML 스키마)는 <a:p>가 최소 1개 있어야 한다(minOccurs="1") — 문단이 하나도 없는
  // <p:txBody>는 PowerPoint가 열 수 없는 손상된 파일이다(FINAL REVIEW C2). pptxTextSchema는
  // doc: 'paragraph+'라 정상적으로 만들어진 문서는 항상 문단이 1개 이상이지만, "텍스트박스 추가"
  // 직후 아직 아무도 타이핑하지 않은 빈 Y.XmlFragment를 y-prosemirror가 ProseMirror 문서로
  // 되돌릴 때는 문단이 0개인 doc이 나올 수 있다(마운트된 에디터 자신의 초기화가 최소 문단을 채워
  // 넣기 전에 자동저장이 먼저 끼어드는 경우 등) — 그 경우에도 항상 빈 문단 하나는 반드시 내보낸다.
  // 빈 문단의 출력 형태는 위 루프가 런이 0개인 문단에 이미 쓰는 것과 똑같이 `<a:p></a:p>`로
  // 맞춘다(자기 닫힘 `<a:p/>` 대신 — 이 함수의 기존 출력 스타일과 일관되게).
  if (paragraphsXml.length === 0) {
    paragraphsXml.push('<a:p></a:p>');
  }
  return `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphsXml.join('')}</p:txBody>`;
}

/** 이미 있는 도형(shapeId로 찾음)의 <p:txBody>만 통째로 교체한다 — 그 도형의 <p:spPr>(위치·
 * 스타일)이나 다른 도형은 전혀 건드리지 않는다. 정규식으로 그 <p:sp>...</p:sp> 블록을 찾고,
 * 그 안의 <p:txBody>...</p:txBody>만 문자열 치환한다(전체를 다시 파싱해 직렬화하면 원본에
 * 있던, 우리가 모델링하지 않는 속성(원본에만 있는 <p:spPr> 하위 요소 등)이 유실될 위험이 있어
 * 문자열 서저리를 유지한다 — src/pptx/collab.ts의 기존 방식과 같은 원칙).
 *
 * id 매칭은 숫자로 제한하지 않는다(`id="([^"]+)"` — 따옴표로 감싸인 값이면 무엇이든 매치) —
 * 원본 슬라이드의 도형은 항상 숫자 id를 쓰지만, 클라이언트가 새로 추가한 텍스트박스는
 * `new-<uuid>` 형태의 문자열 id를 쓴다(appendNewShape 참고). 이 함수가 숫자만 매치했다면
 * 두 번째 저장부터 그 도형을 "찾지 못했다"고 판단해 매번 새로 append하게 되어 저장할 때마다
 * 같은 텍스트박스가 중복 추가되는 버그가 있었다 — 이 완화는 순수하게 넓히는 방향이라(숫자
 * id는 전부 `[^"]+`에도 그대로 매치되므로) 기존에 매치되던 것은 전부 그대로 매치된다.
 *
 * 블록 탐지 정규식(`<p:sp\b...`)도 마찬가지로 순수 확장이다 — 원래 리터럴 `<p:sp>`만 매치했는데,
 * PowerPoint가 실제로 자주 내보내는 `<p:sp useBgFill="1">`처럼 여는 태그에 속성이 붙은 형태는
 * 매치하지 못했다(이 함수는 "못 찾음"으로 throw하지는 않지만, 아래 hasShape가 같은 이유로 "없다"고
 * 잘못 답하면 이 함수가 아니라 appendNewShape가 불려 같은 도형이 중복 추가됐다 — Task 13 fix로
 * `\b`(단어 경계)로 넓혀 속성 유무와 무관하게 매치하면서도 `<p:spPr>`/`<p:spTree>`는 계속
 * 거른다: `sp` 다음이 `Pr`/`Tree`처럼 단어 문자로 이어지면 그 사이엔 경계가 없기 때문).
 *
 * `<p:txBody>...</p:txBody>` 탐지 정규식도 (리뷰에서 지적된) 같은 계열의 문제를 안고 있었다 —
 * 리터럴 `<p:txBody>`만 매치해, 여는 태그에 속성이 붙은 변형(OOXML 스키마 자체는 CT_TextBody에
 * 속성을 정의하지 않으므로 `<p:sp>`보다 훨씬 드문 트리거지만, 어떤 재직렬화 도구가 네임스페이스
 * 재선언 등을 붙이면 여전히 가능하다)을 만나면 이 함수가 "<p:txBody> 없음"으로 잘못 throw했다.
 * 아래에서 `<p:txBody\b`로 넓혀 같은 방식으로 고쳤다. */
/** 원본 txBody의 상자 수준 속성(<a:bodyPr>: 세로 맞춤·여백·자동 맞춤(글자 줄이기)·줄바꿈,
 * <a:lstStyle>: 단계별 기본 글꼴 크기/글머리)을 새로 직렬화한 txBody에 그대로 옮겨 심는다.
 * serializeTxBodyXml은 항상 빈 `<a:bodyPr/><a:lstStyle/>`를 내보내므로, 이걸 안 하면 글자를
 * 한 자만 고쳐도 아래 맞춤 제목이 위로 붙거나 "넘치면 글자 줄이기" 상자가 넘쳐 보이는 등
 * PowerPoint에서 보이는 모양이 바뀌었다(편집 안 한 상자는 애초에 splice되지 않아 무관). */
function carryOverBodyProps(originalTxBody: string, newTxBody: string): string {
  const bodyPr = /<a:bodyPr\b(?:[^>]*\/>|[\s\S]*?<\/a:bodyPr>)/.exec(originalTxBody)?.[0];
  const lstStyle = /<a:lstStyle\b(?:[^>]*\/>|[\s\S]*?<\/a:lstStyle>)/.exec(originalTxBody)?.[0];
  let out = newTxBody;
  if (bodyPr) out = out.replace('<a:bodyPr/>', () => bodyPr);
  if (lstStyle) out = out.replace('<a:lstStyle/>', () => lstStyle);
  return out;
}

export function spliceShapeTxBody(slideXml: string, shapeId: string, txBodyXml: string): string {
  const spBlockRe = /<p:sp\b(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
  let replaced = false;
  const result = slideXml.replace(spBlockRe, (spBlock) => {
    const idMatch = /<p:cNvPr\b[^>]*\bid="([^"]+)"/.exec(spBlock);
    if (!idMatch || idMatch[1] !== shapeId) return spBlock;
    // id는 맞아도 이 <p:sp> 블록에 <p:txBody>가 구조적으로 없을 수 있다(그림 등) — 그 경우
    // 아래 정규식 치환은 아무것도 바꾸지 못한 채 spBlock을 그대로 돌려주는데, 그걸 놓치고
    // replaced만 true로 찍으면 "교체 성공"으로 잘못 보고되며 수정이 조용히 사라진다.
    // 원본 유실 금지가 최우선 기준이므로, 실제로 바꿀 대상이 없으면 여기서 즉시 throw한다.
    //
    // 이 <p:txBody> 탐지 정규식도 위 <p:sp\b 블록 정규식과 정확히 같은 이유로 `\b`(단어 경계)를
    // 쓴다 — 리터럴 <p:txBody>만 매치하면, 어떤 재직렬화 도구가 <p:txBody>에 네임스페이스
    // 재선언 같은 속성을 붙인(OOXML 스키마 자체는 CT_TextBody에 속성을 정의하지 않지만, 실제로
    // 그런 속성 붙은 변형을 만들어내는 도구가 있을 수 있다) 슬라이드를 만나면, parseSlideShapes
    // (실제 XML 파서)는 그 도형의 텍스트 본문을 정상적으로 인식하는데 이 함수만 "<p:txBody> 없음"
    // 으로 잘못 throw한다 — 이 throw는 저장/다운로드 경로 안에서 일어나므로 그 문서의 이후 모든
    // 저장을 영구히 막는다(<p:sp\b 정규식을 리터럴 <p:sp>로 뒀을 때 일어났던 것과 완전히 같은
    // 종류의 결과). `\b`는 <p:txBody>(바로 뒤가 `>`)도, <p:txBody xmlns:a="...">(바로 뒤가 공백)도
    // 매치하면서, <p:txBody> 접두어를 공유하되 실제로는 다른 태그(가정상의 <p:txBodyFoo> 등, 바로
    // 뒤가 단어 문자로 이어짐)는 여전히 거른다.
    if (!/<p:txBody\b[\s\S]*?<\/p:txBody>/.test(spBlock)) {
      throw new Error(`spliceShapeTxBody: shapeId ${shapeId}의 <p:sp> 블록에 <p:txBody>가 없습니다.`);
    }
    replaced = true;
    const originalTxBody = /<p:txBody\b[\s\S]*?<\/p:txBody>/.exec(spBlock)![0];
    return spBlock.replace(/<p:txBody\b[\s\S]*?<\/p:txBody>/, () => carryOverBodyProps(originalTxBody, txBodyXml));
  });
  if (!replaced) {
    throw new Error(`spliceShapeTxBody: shapeId ${shapeId}를 슬라이드 XML에서 찾지 못했습니다.`);
  }
  return result;
}

/** shapeId를 가진 <p:sp> 블록이 슬라이드 XML 안에 이미 존재하는지만 확인한다(내용은 건드리지
 * 않는다) — spliceShapeTxBody와 정확히 같은 블록 탐색/속성 매칭 정규식을 쓴다(둘이 어긋나면
 * "있다"고 했는데 spliceShapeTxBody는 못 찾는 모순이 생길 수 있으므로 반드시 같아야 한다).
 *
 * src/pptx/collab.ts의 export 병합 루프가 도형별로 이 함수를 먼저 호출해 "이미 있는 도형(교체 —
 * spliceShapeTxBody)"과 "처음 쓰는 도형(추가 — appendNewShape)"을 구분한다. 예전에는 shapeId가
 * `new-`로 시작하는지만 보고 무조건 appendNewShape를 불렀는데, 그러면 이미 한 번 저장되어 슬라이드
 * XML에 실제로 존재하는 `new-` 도형을 두 번째 저장에서 또 추가해 버려 저장할 때마다 같은
 * 텍스트박스가 중복되는 버그가 있었다 — 이 함수로 먼저 존재 여부를 확인하면 그 문제가 사라진다.
 *
 * **Task 13 fix로 고친, 그와 같은 계열의 별개 버그**: 이 함수가 쓰는 블록 탐지 정규식이 최근까지
 * (Task 3/6이 만든 그대로) 리터럴 `<p:sp>`만 매치했다 — 여는 태그에 속성이 붙은
 * `<p:sp useBgFill="1">` 같은, PowerPoint가 실제로 흔히 만들어내는 형태는 매치하지 못했다. 그
 * 결과 이 함수가 실제로 존재하는 속성 있는 도형을 "없다"고 잘못 답해, 위 병합 루프가 교체
 * (spliceShapeTxBody) 대신 추가(appendNewShape) 분기를 타 버려 그 도형이 같은 shapeId로 중복
 * 저장됐다(구 텍스트와 새 텍스트가 둘 다 파일에 남는 조용한 데이터 손상). `\b`(단어 경계)로
 * 정규식을 넓혀 지금은 속성 유무와 무관하게 매치한다. */
export function hasShape(slideXml: string, shapeId: string): boolean {
  const spBlockRe = /<p:sp\b(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
  for (const spBlock of slideXml.match(spBlockRe) ?? []) {
    const idMatch = /<p:cNvPr\b[^>]*\bid="([^"]+)"/.exec(spBlock);
    if (idMatch && idMatch[1] === shapeId) return true;
  }
  return false;
}

/** 원본에 없던 도형(새로 추가한 텍스트박스)을 <p:spTree> 끝에 끼워 넣는다. shapeId는
 * new-<uuid> 형태를 그대로 <p:cNvPr name="..">에는 안 쓰고 id 자리(문자열)에 그대로 쓴다 —
 * PowerPoint의 <p:cNvPr id>는 원래 숫자지만, 우리가 다시 읽을 때(parseSlideShapes)도
 * 문자열로만 비교하므로 실제로는 아무 문제가 없다(PowerPoint 자체가 다시 열 때 재검증하지
 * 않는 필드). */
export function appendNewShape(
  slideXml: string,
  meta: { shapeId: string; x: number; y: number; w: number; h: number; geom?: string; fill?: string },
  txBodyXml: string,
): string {
  // geom/fill은 호출자(exportPptxFromSlides)가 화이트리스트·형식 검증을 끝낸 값만 넘긴다.
  const isShape = meta.geom !== undefined;
  const fillXml = isShape && meta.fill ? `<a:solidFill><a:srgbClr val="${meta.fill.replace('#', '').toUpperCase()}"/></a:solidFill>` : '';
  const spXml =
    `<p:sp><p:nvSpPr><p:cNvPr id="${meta.shapeId}" name="${isShape ? 'Shape' : 'TextBox'}"/>` +
    `${isShape ? '<p:cNvSpPr/>' : '<p:cNvSpPr txBox="1"/>'}<p:nvPr/></p:nvSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${meta.x}" y="${meta.y}"/><a:ext cx="${meta.w}" cy="${meta.h}"/></a:xfrm>` +
    `<a:prstGeom prst="${meta.geom ?? 'rect'}"><a:avLst/></a:prstGeom>${fillXml}</p:spPr>` +
    // 새 도형의 글자는 실제 프레젠테이션 프로그램처럼 도형 한가운데(세로 가운데) — 이후 재저장에서도
    // carryOverBodyProps가 이 bodyPr를 그대로 유지한다.
    (isShape ? txBodyXml.replace('<a:bodyPr/>', '<a:bodyPr anchor="ctr"/>') : txBodyXml) +
    '</p:sp>';
  const closeIdx = slideXml.lastIndexOf('</p:spTree>');
  if (closeIdx === -1) {
    throw new Error('appendNewShape: 슬라이드 XML에서 </p:spTree>를 찾지 못했습니다.');
  }
  return slideXml.slice(0, closeIdx) + spXml + slideXml.slice(closeIdx);
}

/** shapeId에 해당하는 <p:sp> 블록 전체를 슬라이드 XML에서 제거한다(사용자가 텍스트박스를
 * 삭제했을 때 씀 — src/pptx/collab.ts의 exportPptxFromSlides가 "baseline에는 있었는데 지금
 * perSlide에는 없는" 도형에 대해 호출한다). hasShape/spliceShapeTxBody와 정확히 같은 블록 탐색/
 * 속성 매칭 정규식을 쓴다(id 매칭을 숫자로 제한하지 않으므로 `new-<uuid>` 형태의 id도 그대로
 * 매치된다) — 셋이 어긋나면 "있다"고 판단한 도형을 못 찾거나, 지워야 할 도형을 못 찾는 모순이
 * 생길 수 있으므로 반드시 같아야 한다. 다른 도형의 <p:sp> 블록은 전혀 건드리지 않는다.
 *
 * 못 찾으면 throw한다(이 파일의 다른 함수들과 동일한 "절대 조용히 실패하지 않는다" 원칙) — 이
 * 호출부(exportPptxFromSlides)는 항상 baseline에 실제로 존재하는 shapeId로만 이 함수를 부르므로,
 * 정상적인 사용에서는 이 throw가 절대 발생하지 않아야 한다(발생한다면 그 자체가 병합 로직의
 * 불변식이 깨졌다는 신호이므로, 조용히 넘어가지 않고 크게 실패하는 편이 "원본 유실 금지"
 * 원칙에 맞다).
 *
 * **주의(Task 13 fix 이전에는 이 불변식이 실제로 깨져 있었다)**: baseline은 parseSlideShapes(실제
 * XML 파서)로 만들어지므로 여는 태그에 속성이 붙은 `<p:sp useBgFill="1">` 같은 도형도 정확히
 * 인식해 baseline에 포함시킨다. 반면 이 함수가 쓰는 블록 탐지 정규식은 최근까지 리터럴 `<p:sp>`만
 * 매치했다 — 그래서 baseline에는 있고(파서가 봤으니) perSlide에는 없는(사용자가 지웠으니) 속성
 * 있는 도형을 지우려 하면, 그 블록이 XML 안에 실제로 있는데도 정규식이 찾지 못해 이 throw가
 * 정상적인 삭제 조작만으로 발생했다 — 그리고 이 throw는 exportPptxFromSlides(저장/다운로드 경로)
 * 안에서 일어나므로, 그 문서의 이후 모든 저장·다운로드를 영구히 막았다(orphan된 baseline id는
 * 절대 사라지지 않으므로 사용자 쪽에서 되돌릴 방법도 없었다). 지금은 세 함수의 블록 정규식을
 * 전부 `<p:sp\b`로 넓혀(속성 유무와 무관하게 매치하되 `<p:spPr>`/`<p:spTree>`는 여전히 제외)
 * baseline과 이 함수가 항상 같은 도형 집합을 인식한다.
 *
 * **위 "절대 발생하지 않아야 한다"는 주장의 정확한 범위**: 이건 `<p:sp\b` 블록 탐지 정규식에
 * 한정된 얘기다 — 세 함수가 함께 쓰는 또 다른 정규식, `<p:cNvPr>`에서 id 속성값을 뽑는
 * `/<p:cNvPr\b[^>]*\bid="([^"]+)"/`는 여전히 실제 XML 파서와 완전히 동일하지 않다. PowerPoint와
 * LibreOffice가 실제로 내보내는, 큰따옴표를 쓰고 `id`가 다른 속성보다 먼저 오는 직렬화 형태에서는
 * 항상 일치하지만, 합법적이지만 어느 실제 생산자도 만들어내지 않는 변형 — 작은따옴표
 * (`id='7'`), `=` 앞뒤 공백(`id = "7"`), 또는 `id`보다 앞선 속성의 값 안에 이스케이프 안 된 `>`가
 * 들어있는 경우(예: `name="a > b" id="7"`) — 에서는 이 id 정규식이 이론적으로 여전히 실제 파서와
 * 다르게 판단할 수 있다. 그러므로 "이 throw가 절대 발생하지 않는다"는 건 실제 생산자가 만들어내는
 * 직렬화 형태에 한해 참이지, 위와 같은 이론상 가능한 모든 합법적 XML 변형까지 포괄하는 절대적
 * 보장은 아니다. */
export function removeShape(slideXml: string, shapeId: string): string {
  const spBlockRe = /<p:sp\b(?:(?!<\/p:sp>)[\s\S])*?<\/p:sp>/g;
  let removed = false;
  const result = slideXml.replace(spBlockRe, (spBlock) => {
    const idMatch = /<p:cNvPr\b[^>]*\bid="([^"]+)"/.exec(spBlock);
    if (!idMatch || idMatch[1] !== shapeId) return spBlock;
    removed = true;
    return '';
  });
  if (!removed) {
    throw new Error(`removeShape: shapeId ${shapeId}를 슬라이드 XML에서 찾지 못했습니다.`);
  }
  return result;
}

/* ==================== 표시 전용 요소(이미지·단순 도형·표/차트/스마트아트) ====================
 * 아래 전부는 parseSlideShapes(편집 대상 텍스트박스)와 완전히 별개 경로다 — 결과가
 * SlideVisual[]로 slideVisuals(협업 편집 상태와 무관한 표시 전용 맵)에만 쓰이고, 절대
 * export되지 않는다(그 요소들이 있던 자리는 원본 XML 그대로 복사되므로 여기서 뭘 하든
 * 저장되는 파일에는 영향이 없다). */

/** .rels 파일에서 "rId → Target" 매핑만 뽑는다(src/pptx/collab.ts의 slideRelsMap과 같은 패턴,
 * 이 파일은 그쪽을 import할 이유가 없을 만큼 단순해 그대로 복제했다). */
function relsMap(relsXml: string): Map<string, string> {
  const map = new Map<string, string>();
  const relRe = /<Relationship\b[^>]*\/>/g;
  for (const relTag of relsXml.match(relRe) ?? []) {
    const idMatch = /Id="(rId\d+)"/.exec(relTag);
    const targetMatch = /Target="([^"]+)"/.exec(relTag);
    if (idMatch && targetMatch) map.set(idMatch[1], targetMatch[1]);
  }
  return map;
}

/** .rels의 Target(예: "../media/image1.png")을 슬라이드 파트 기준 zip 내부 절대 경로로 푼다. */
function resolveRelativePath(basePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const stack = basePart.split('/').slice(0, -1);
  for (const part of target.split('/')) {
    if (part === '..') stack.pop();
    else if (part !== '.') stack.push(part);
  }
  return stack.join('/');
}

/** src/routes/viewer.ts의 /api/view/slide-media도 Content-Type 헤더를 정하는 데 그대로
 * 재사용한다 — 여기서 "지원 포맷"으로 판단한 것과 실제로 내려주는 헤더가 어긋나지 않게. */
export const IMAGE_MIME_BY_EXT: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  webp: 'image/webp',
  // emf/wmf/tiff 등은 의도적으로 뺐다 — 브라우저가 <img>로 직접 그릴 수 없는 포맷이라
  // data URI를 만들어봐야 화면에 아무것도 안 나온다(깨진 이미지 아이콘보다는 그냥 생략이 낫다).
};

/** presentation.xml의 실제 슬라이드 크기(<p:sldSz cx cy>, EMU). 없거나 읽을 수 없으면 표준 16:9.
 * 4:3(9144000×6858000) 같은 다른 크기의 문서를 16:9 좌표계로 그리면 모든 도형이 왼쪽으로
 * 찌그러지고 비율이 틀어진다 — 공공기관 문서에 흔한 4:3 파일로 실제 재현. */
export function readSlideSizeEmu(zipEntries: Record<string, Uint8Array>): { cx: number; cy: number } {
  const bytes = zipEntries['ppt/presentation.xml'];
  const tag = bytes ? /<p:sldSz\b[^>]*>/.exec(Buffer.from(bytes).toString('utf-8'))?.[0] : undefined;
  const cx = Number(tag ? /\bcx="(\d+)"/.exec(tag)?.[1] : NaN);
  const cy = Number(tag ? /\bcy="(\d+)"/.exec(tag)?.[1] : NaN);
  return cx > 0 && cy > 0 ? { cx, cy } : { cx: SLIDE_WIDTH_EMU, cy: SLIDE_HEIGHT_EMU };
}

function presetGeomFromSpPr(spPr: XmlNode | undefined): string | undefined {
  const prst = (spPr?.['a:prstGeom'] as XmlNode | undefined)?.['@_prst'];
  return typeof prst === 'string' && prst !== 'rect' && /^[A-Za-z0-9]+$/.test(prst) ? prst : undefined;
}

/**
 * 도형의 보이는 채우기 색. 우선순위는 PowerPoint와 같다: 명시적 <a:noFill/> → 없음, <a:solidFill> →
 * 그 색, <a:gradFill> → 첫 색 지점(그라데이션은 근사), 그 밖엔 도형 스타일 <p:style><a:fillRef idx>의
 * 테마 색(기본 도형은 대부분 이 경로로 accent1 같은 색을 갖는다 — 예전엔 이걸 몰라 기본 도형이
 * 투명하게 그려졌다). 테마를 모르면 직접 지정 RGB만 푼다.
 */
function fillColorFromSpPr(spPr: XmlNode | undefined, theme?: ThemeColors, style?: XmlNode): string | undefined {
  if (spPr?.['a:noFill'] !== undefined) return undefined;
  const solid = spPr?.['a:solidFill'] as XmlNode | undefined;
  if (solid !== undefined) {
    if (theme) return theme.resolve(solid);
    const val = (solid['a:srgbClr'] as XmlNode | undefined)?.['@_val'] as string | undefined;
    return val ? `#${val}` : undefined;
  }
  const grad = spPr?.['a:gradFill'] as XmlNode | undefined;
  if (grad && theme) return gradientCss(grad, theme);
  if (spPr?.['a:blipFill'] !== undefined || spPr?.['a:pattFill'] !== undefined) return undefined;
  const fillRef = style?.['a:fillRef'] as XmlNode | undefined;
  if (fillRef && theme && Number(fillRef['@_idx'] ?? 0) > 0) return theme.resolve(fillRef);
  return undefined;
}

/** <a:gradFill> → CSS 그라데이션. 정지점은 파일에 적힌 순서가 아니라 위치(pos) 순으로 정렬한다
 * (PowerPoint는 끝 정지점을 먼저 적기도 한다). 선형 각도: OOXML 0°=왼→오른쪽, CSS 90°=to right. */
function gradientCss(grad: XmlNode, theme: ThemeColors): string | undefined {
  const stops = asArray<XmlNode>((grad['a:gsLst'] as XmlNode | undefined)?.['a:gs'] as XmlNode | XmlNode[] | undefined)
    .map((gs) => ({ pos: Number(gs['@_pos'] ?? 0) / 1000, color: theme.resolve(gs) }))
    .filter((stop): stop is { pos: number; color: string } => Boolean(stop.color))
    .sort((a, b) => a.pos - b.pos);
  if (stops.length === 0) return undefined;
  if (stops.length === 1) return stops[0].color;
  const round = (n: number) => Number(n.toFixed(2));
  const list = stops.map((stop) => `${stop.color} ${round(stop.pos)}%`).join(', ');
  if (grad['a:path'] !== undefined) return `radial-gradient(${list})`;
  const deg = Number((grad['a:lin'] as XmlNode | undefined)?.['@_ang'] ?? 0) / 60000;
  return `linear-gradient(${round((deg + 90) % 360)}deg, ${list})`;
}

/** 색이 완전히 투명(<a:alpha val="0"/>)한지 — 그런 선은 PowerPoint에서도 안 보인다 */
function isFullyTransparent(container: XmlNode | undefined): boolean {
  if (!container || typeof container !== 'object') return false;
  for (const key of ['a:srgbClr', 'a:schemeClr', 'a:sysClr', 'a:prstClr']) {
    const node = container[key] as XmlNode | undefined;
    const alpha = node && typeof node === 'object' ? (node['a:alpha'] as XmlNode | undefined)?.['@_val'] : undefined;
    if (alpha !== undefined) return Number(alpha) <= 0;
  }
  return false;
}

/**
 * 도형의 테두리(선) — <a:ln>의 <a:noFill/>이면 없음, <a:solidFill>이면 그 색, 둘 다 없으면 도형 스타일
 * <p:style><a:lnRef idx>의 테마 색(기본 도형의 테두리). 두께는 EMU(기본 0.75pt), 점선 종류는 CSS
 * border-style 셋으로 줄인다. 예전엔 테두리를 전혀 읽지 않아 "채우기 없음 + 빨간 점선" 강조 상자가
 * 편집 화면에서 통째로 사라졌다.
 */
function strokeFromSpPr(spPr: XmlNode | undefined, theme?: ThemeColors, style?: XmlNode): SlideVisual['stroke'] {
  const ln = spPr?.['a:ln'] as XmlNode | undefined;
  if (ln?.['a:noFill'] !== undefined) return undefined;
  if (ln?.['a:gradFill'] !== undefined || ln?.['a:pattFill'] !== undefined) return undefined;
  const solid = ln?.['a:solidFill'] as XmlNode | undefined;
  let color: string | undefined;
  if (solid !== undefined) {
    if (isFullyTransparent(solid)) return undefined;
    color = theme ? theme.resolve(solid) : undefined;
    if (!color) {
      const rgb = (solid['a:srgbClr'] as XmlNode | undefined)?.['@_val'];
      color = typeof rgb === 'string' && /^[0-9A-Fa-f]{6}$/.test(rgb) ? `#${rgb.toUpperCase()}` : undefined;
    }
  } else {
    const lnRef = style?.['a:lnRef'] as XmlNode | undefined;
    if (lnRef && theme && Number(lnRef['@_idx'] ?? 0) > 0 && !isFullyTransparent(lnRef)) color = theme.resolve(lnRef);
  }
  if (!color) return undefined;
  const width = Number(ln?.['@_w'] ?? 9525) || 9525;
  const prstDash = String((ln?.['a:prstDash'] as XmlNode | undefined)?.['@_val'] ?? 'solid');
  const dash = prstDash === 'solid' ? 'solid' : /dot/i.test(prstDash) && !/dash/i.test(prstDash) ? 'dotted' : 'dashed';
  return { color, width, dash };
}

/** 회전(<a:xfrm rot>, 1/60000도)이 걸린 개체가 실제로 차지하는 영역 — 원본 렌더링에서 잘라 올 때 쓴다.
 * 회전 전 상자를 그대로 쓰면(예: 90° 돌린 가로 연결선 = 실제로는 세로선) 엉뚱한 곳을 잘라 와 선이 사라졌다. */
function rotatedBounds(
  pos: { x: number | null; y: number | null; w: number | null; h: number | null },
  rot: number,
): { x: number | null; y: number | null; w: number | null; h: number | null } {
  if (!rot || pos.x === null || pos.y === null || pos.w === null || pos.h === null) return pos;
  const rad = ((rot / 60000) * Math.PI) / 180;
  const cos = Math.abs(Math.cos(rad));
  const sin = Math.abs(Math.sin(rad));
  const w = pos.w * cos + pos.h * sin;
  const h = pos.w * sin + pos.h * cos;
  const cx = pos.x + pos.w / 2;
  const cy = pos.y + pos.h / 2;
  return { x: cx - w / 2, y: cy - h / 2, w, h };
}

/** 선·연결선을 원본 렌더링에서 잘라 올 때 사방으로 넓히는 여유(약 6px) — 화살촉과 두께가 들어오게 */
const LINE_SNAPSHOT_PAD_EMU = 6 * 9525;

/** 편집 화면이 CSS 테두리로 그대로 그릴 수 있는 모양(사각형·둥근 사각형·타원 계열) */
const STROKABLE_GEOMS = new Set([undefined, 'roundRect', 'ellipse', 'flowChartConnector', 'flowChartAlternateProcess', 'round2SameRect', 'flowChartTerminator', 'flowChartProcess']);

/** 글자 도형에 실제로 보이는 글자(공백 아닌 a:t, 또는 쪽 번호 같은 필드)가 있는지. */
function hasVisibleText(txBody: XmlNode | undefined): boolean {
  for (const para of asArray<XmlNode>(txBody?.['a:p'] as XmlNode | XmlNode[] | undefined)) {
    if (para['a:fld'] !== undefined) return true;
    for (const run of asArray<XmlNode>(para['a:r'] as XmlNode | XmlNode[] | undefined)) {
      if (String(run['a:t'] ?? '').trim() !== '') return true;
    }
  }
  return false;
}

function graphicFrameXfrm(gf: XmlNode) {
  const xfrm = gf['p:xfrm'] as XmlNode | undefined; // 다른 도형과 달리 p:spPr 없이 바로 p:xfrm
  const off = xfrm?.['a:off'] as XmlNode | undefined;
  const ext = xfrm?.['a:ext'] as XmlNode | undefined;
  return { x: emuAttr(off, '@_x'), y: emuAttr(off, '@_y'), w: emuAttr(ext, '@_cx'), h: emuAttr(ext, '@_cy') };
}

/** <a:tc>(표 셀) 하나의 텍스트만 뽑는다 — 서식은 버리고 문단은 줄바꿈으로 잇는다(표는 편집
 * 대상이 아니라 화면 표시용이라 이 정도 단순화로 충분하다). */
function tableCellText(tc: XmlNode): string {
  const txBody = tc['a:txBody'] as XmlNode | undefined;
  if (!txBody) return '';
  return asArray<XmlNode>(txBody['a:p'] as XmlNode | XmlNode[] | undefined)
    .map((p) =>
      asArray<XmlNode>(p['a:r'] as XmlNode | XmlNode[] | undefined)
        .map((r) => String((r['a:t'] as string | number | undefined) ?? ''))
        .join(''),
    )
    .join('\n');
}

const TABLE_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/table';
const CHART_GRAPHIC_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

/** 파트(슬라이드/레이아웃/마스터) 하나를 표시용으로 읽어 둔 것 — 미디어 경로는 그 파트의 rels 기준. */
interface VisualPart {
  readonly part: string;
  readonly cSld: XmlNode;
  readonly root: XmlNode;
  readonly rels: Map<string, string>;
  /** 도형 id → 원문에 나오는 순서(= 겹침 순서, 앞이 아래) */
  readonly order: Map<string, number>;
}

function loadVisualPart(zipEntries: Record<string, Uint8Array>, part: string | null, rootTag: string): VisualPart | null {
  if (!part || !zipEntries[part]) return null;
  try {
    const xml = Buffer.from(zipEntries[part]).toString('utf-8');
    const parsed = parser.parse(xml) as XmlNode;
    const root = parsed[rootTag] as XmlNode | undefined;
    const cSld = root?.['p:cSld'] as XmlNode | undefined;
    if (!root || !cSld) return null;
    const slash = part.lastIndexOf('/');
    const relsBytes = zipEntries[`${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`];
    const rels = relsBytes ? relsMap(Buffer.from(relsBytes).toString('utf-8')) : new Map<string, string>();
    const order = new Map<string, number>();
    for (const m of xml.matchAll(/<p:cNvPr\b[^>]*?\bid="(\d+)"/g)) if (!order.has(m[1])) order.set(m[1], order.size);
    return { part, cSld, root, rels, order };
  } catch {
    return null;
  }
}

/** 슬라이드 XML 안의 이미지(p:pic)·글자 없는 단순 도형(p:sp)·표/차트/스마트아트
 * (p:graphicFrame)를 화면 표시용으로 뽑는다. 표는 셀 텍스트까지 그리고, 차트·스마트아트는
 * (렌더링 엔진 없이는 실제로 그릴 수 없어) 위치만 맞는 자리표시자로 남긴다.
 *
 * 그리는 순서(아래부터): 배경(슬라이드 → 레이아웃 → 마스터 중 먼저 정의된 것) → 마스터의 그림·도형 →
 * 레이아웃의 그림·도형 → 슬라이드의 그림·도형. 레이아웃/마스터의 개체 틀(p:ph)은 "틀"일 뿐 실제로
 * 보이지 않으므로 건너뛰고, PowerPoint의 "배경 그래픽 숨기기"(showMasterSp="0")를 따른다 — 예전엔
 * 슬라이드만 읽어서 회사 로고·색 띠처럼 마스터에 있는 요소와 마스터 배경이 편집 화면에서 빠졌다. */
export function parseSlideVisuals(zipEntries: Record<string, Uint8Array>, sourceSlidePart: string): SlideVisual[] {
  const slidePartPath = `ppt/slides/${sourceSlidePart}`;
  const slide = loadVisualPart(zipEntries, slidePartPath, 'p:sld');
  if (!slide) return [];
  let theme: ThemeColors | undefined;
  try {
    theme = loadThemeColors(zipEntries, sourceSlidePart);
  } catch {
    theme = undefined;
  }
  const layoutPath = relTargetOf(zipEntries, slidePartPath, 'slideLayout');
  const layout = loadVisualPart(zipEntries, layoutPath, 'p:sldLayout');
  const master = loadVisualPart(zipEntries, layoutPath ? relTargetOf(zipEntries, layoutPath, 'slideMaster') : null, 'p:sldMaster');
  const visuals: SlideVisual[] = [];

  function resolveImageMediaPath(owner: VisualPart, rId: string | undefined): string | undefined {
    if (!rId) return undefined;
    const target = owner.rels.get(rId);
    if (!target) return undefined;
    const normalized = resolveRelativePath(owner.part, target);
    if (!zipEntries[normalized]) return undefined; // 관계는 있는데 실제 미디어 파트가 없음(손상)
    const ext = normalized.split('.').pop()?.toLowerCase() ?? '';
    if (!IMAGE_MIME_BY_EXT[ext]) return undefined; // 브라우저가 못 그리는 포맷(emf/wmf 등)
    return normalized;
  }

  // 배경 — <p:bg>는 <p:spTree> 형제라 도형 루프가 보지 못한다. 슬라이드에 없으면 레이아웃,
  // 그다음 마스터의 배경이 보인다(PowerPoint와 같음). bgPr(직접 지정)과 bgRef(테마 배경 스타일의
  // 색) 둘 다 처리한다. 다른 모든 요소보다 먼저 push해 맨 아래에 깐다.
  const bgOwner = [slide, layout, master].find((p) => p && p.cSld['p:bg'] !== undefined) ?? null;
  if (bgOwner) {
    const { cx, cy } = readSlideSizeEmu(zipEntries);
    const bg = bgOwner.cSld['p:bg'] as XmlNode;
    const bgPr = bg['p:bgPr'] as XmlNode | undefined;
    const bgRef = bg['p:bgRef'] as XmlNode | undefined;
    const blip = (bgPr?.['a:blipFill'] as XmlNode | undefined)?.['a:blip'] as XmlNode | undefined;
    const mediaPath = resolveImageMediaPath(bgOwner, blip?.['@_r:embed'] as string | undefined);
    if (mediaPath) {
      visuals.push({ kind: 'image', x: 0, y: 0, w: cx, h: cy, mediaPath, background: true });
    } else {
      const fill = bgPr ? fillColorFromSpPr(bgPr, theme) : bgRef && theme ? theme.resolve(bgRef) : undefined;
      if (fill) visuals.push({ kind: 'shape', x: 0, y: 0, w: cx, h: cy, fill, background: true });
    }
  }

  // <p:grpSp>(그룹 도형) 안에 중첩된 이미지·단순 도형·표/차트도 재귀로 찾아 그룹의 좌표 변환을
  // 적용한 절대 좌표로 push한다(parseSlideShapes의 collectShapesRecursive와 같은 이유 —
  // GroupTransform 설명 참고).
  // 그림·도형·표·선을 종류별로 모으므로, 원본의 겹침 순서(파트 원문 순서)를 따로 기록해 마지막에 정렬한다
  const sortKey = new Map<SlideVisual, number>();
  let partRank = 0;
  function collectVisualsRecursive(owner: VisualPart, node: XmlNode, transforms: readonly GroupTransform[], template: boolean): void {
    const isPlaceholder = (el: XmlNode, nvTag: string) =>
      ((el[nvTag] as XmlNode | undefined)?.['p:nvPr'] as XmlNode | undefined)?.['p:ph'] !== undefined;
    const push = (el: XmlNode, nvTag: string, visual: SlideVisual) => {
      const id = String(((el[nvTag] as XmlNode | undefined)?.['p:cNvPr'] as XmlNode | undefined)?.['@_id'] ?? '');
      visuals.push(visual);
      sortKey.set(visual, partRank * 1_000_000 + (owner.order.get(id) ?? 999_999));
    };

    for (const pic of asArray<XmlNode>(node['p:pic'] as XmlNode | XmlNode[] | undefined)) {
      if (template && isPlaceholder(pic, 'p:nvPicPr')) continue;
      const spPr = pic['p:spPr'] as XmlNode | undefined;
      const xfrm = spPr?.['a:xfrm'] as XmlNode | undefined;
      const off = xfrm?.['a:off'] as XmlNode | undefined;
      const ext = xfrm?.['a:ext'] as XmlNode | undefined;
      const blip = (pic['p:blipFill'] as XmlNode | undefined)?.['a:blip'] as XmlNode | undefined;
      const mediaPath = resolveImageMediaPath(owner, blip?.['@_r:embed'] as string | undefined);
      const pos = applyGroupTransforms(transforms, emuAttr(off, '@_x'), emuAttr(off, '@_y'), emuAttr(ext, '@_cx'), emuAttr(ext, '@_cy'));
      // 브라우저가 못 그리는 포맷(EMF/WMF — 한글에서 붙여넣은 도식에 흔함)이면 원본 렌더링에서 잘라 쓴다
      if (!mediaPath) {
        push(pic, 'p:nvPicPr', { kind: 'snapshot', ...pos, snapshot: true });
        continue;
      }
      // 그림에 직접 준 테두리(스크린샷 강조용 빨간 테두리 등)
      const picStroke = strokeFromSpPr(spPr, theme, pic['p:style'] as XmlNode | undefined);
      push(pic, 'p:nvPicPr', { kind: 'image', ...pos, mediaPath, ...(picStroke ? { stroke: picStroke } : {}) });
    }

    for (const sp of asArray<XmlNode>(node['p:sp'] as XmlNode | XmlNode[] | undefined)) {
      if (template && isPlaceholder(sp, 'p:nvSpPr')) continue; // 레이아웃·마스터의 개체 틀은 보이지 않는 "틀"
      const spPr = sp['p:spPr'] as XmlNode | undefined;
      const fill = fillColorFromSpPr(spPr, theme, sp['p:style'] as XmlNode | undefined);
      const xfrm = spPr?.['a:xfrm'] as XmlNode | undefined;
      const off = xfrm?.['a:off'] as XmlNode | undefined;
      const ext = xfrm?.['a:ext'] as XmlNode | undefined;
      const pos = applyGroupTransforms(transforms, emuAttr(off, '@_x'), emuAttr(off, '@_y'), emuAttr(ext, '@_cx'), emuAttr(ext, '@_cy'));
      // 레이아웃·마스터의 글자 도형(꼬리말 문구, 쪽 번호 등)은 편집 대상이 아니다 — 글자(쪽 번호는
      // 슬라이드마다 다르다)까지 원본 렌더링에서 잘라 그대로 보여준다.
      if (template && hasVisibleText(sp['p:txBody'] as XmlNode | undefined) && pos.x !== null && pos.y !== null) {
        push(sp, 'p:nvSpPr', { kind: 'snapshot', ...pos, snapshot: true });
        continue;
      }
      // 글자가 든 도형의 "글자"는 parseSlideShapes가 편집 대상으로 다룬다 — 하지만 그 도형에 채우기
      // 색이 있으면(색 칠한 사각형 위의 제목 등, 실사용 문서에 매우 흔함) 예전엔 배경색이 통째로
      // 빠져 글자만 허공에 떠 보였다. 채우기가 있을 때만 그 색 판을 텍스트 상자 아래에 깔아 준다.
      const stroke = strokeFromSpPr(spPr, theme, sp['p:style'] as XmlNode | undefined);
      if (sp['p:txBody'] && !fill && !stroke) continue;
      const geom = presetGeomFromSpPr(spPr);
      const custom = spPr?.['a:custGeom'] !== undefined;
      // 채우기 없이 선만 있는 화살표·곡선·자유형처럼 CSS 테두리로 흉내 낼 수 없는 모양은 원본 렌더링에서
      // 잘라 쓴다(자유형은 채우기만 있어도 사각형으로 그리면 모양이 틀리므로 마찬가지). 채우기가 있는
      // 프리셋 모양(화살표 등)은 지금처럼 모양대로 칠하기만 한다. 글자 도형은 글자가 편집 대상이라 잘라
      // 쓰지 않고(글자가 겹쳐 보인다) 채우기만 깐다.
      const needsSnapshot = custom ? Boolean(fill || stroke) : Boolean(stroke && !fill && !STROKABLE_GEOMS.has(geom));
      if (!sp['p:txBody'] && pos.x !== null && pos.y !== null && needsSnapshot) {
        // 선만 있는 모양(직선·곡선 화살표 등)은 선 두께·화살촉이 상자 밖으로 나간다 — 연결선처럼 넓혀 자른다
        const pad = stroke && !fill ? Math.max(LINE_SNAPSHOT_PAD_EMU, stroke.width * 3) : 0;
        const area = rotatedBounds(pos, Number(xfrm?.['@_rot'] ?? 0));
        push(sp, 'p:nvSpPr', {
          kind: 'snapshot',
          x: (area.x ?? 0) - pad,
          y: (area.y ?? 0) - pad,
          w: (area.w ?? 0) + pad * 2,
          h: (area.h ?? 0) + pad * 2,
          snapshot: true,
        });
        continue;
      }
      if (!fill && !(stroke && STROKABLE_GEOMS.has(geom) && !custom)) continue;
      push(sp, 'p:nvSpPr', {
        kind: 'shape',
        ...pos,
        fill,
        ...(geom ? { geom } : {}),
        ...(stroke && STROKABLE_GEOMS.has(geom) && !custom ? { stroke } : {}),
      });
    }

    for (const gf of asArray<XmlNode>(node['p:graphicFrame'] as XmlNode | XmlNode[] | undefined)) {
      if (template && isPlaceholder(gf, 'p:nvGraphicFramePr')) continue;
      const raw = graphicFrameXfrm(gf);
      const { x, y, w, h } = applyGroupTransforms(transforms, raw.x, raw.y, raw.w, raw.h);
      const graphicData = (gf['a:graphic'] as XmlNode | undefined)?.['a:graphicData'] as XmlNode | undefined;
      const uri = graphicData?.['@_uri'] as string | undefined;
      if (uri === TABLE_GRAPHIC_URI) {
        const tbl = graphicData?.['a:tbl'] as XmlNode | undefined;
        const rows = asArray<XmlNode>(tbl?.['a:tr'] as XmlNode | XmlNode[] | undefined).map((tr) =>
          asArray<XmlNode>(tr['a:tc'] as XmlNode | XmlNode[] | undefined).map(tableCellText),
        );
        push(gf, 'p:nvGraphicFramePr', { kind: 'table', x, y, w, h, rows, snapshot: true });
      } else {
        const label = uri === CHART_GRAPHIC_URI ? '차트' : uri?.includes('/diagram') ? '스마트아트' : '콘텐츠';
        push(gf, 'p:nvGraphicFramePr', { kind: 'placeholder', x, y, w, h, label, snapshot: true });
      }
    }

    // 선·연결선(<p:cxnSp>) — 예전엔 아예 그리지 않았다(도형 사이 화살표·구분선이 사라짐)
    for (const cxn of asArray<XmlNode>(node['p:cxnSp'] as XmlNode | XmlNode[] | undefined)) {
      if (template && isPlaceholder(cxn, 'p:nvCxnSpPr')) continue;
      const xfrm = (cxn['p:spPr'] as XmlNode | undefined)?.['a:xfrm'] as XmlNode | undefined;
      const off = xfrm?.['a:off'] as XmlNode | undefined;
      const ext = xfrm?.['a:ext'] as XmlNode | undefined;
      const pos = rotatedBounds(
        applyGroupTransforms(transforms, emuAttr(off, '@_x'), emuAttr(off, '@_y'), emuAttr(ext, '@_cx'), emuAttr(ext, '@_cy')),
        Number(xfrm?.['@_rot'] ?? 0),
      );
      if (pos.x === null || pos.y === null) continue;
      // 선은 상자 테두리 위에 그려지고 화살촉·두께가 상자 밖으로 나간다 — 똑바른 세로/가로선은 상자 폭(높이)이
      // 0이라 잘라 낸 조각이 비어 버렸다. 사방으로 조금 넓혀 잘라 온다.
      const lineWidth = Number(((cxn['p:spPr'] as XmlNode | undefined)?.['a:ln'] as XmlNode | undefined)?.['@_w'] ?? 0) || 0;
      const pad = Math.max(LINE_SNAPSHOT_PAD_EMU, lineWidth * 3);
      push(cxn, 'p:nvCxnSpPr', {
        kind: 'snapshot',
        x: pos.x - pad,
        y: pos.y - pad,
        w: (pos.w ?? 0) + pad * 2,
        h: (pos.h ?? 0) + pad * 2,
        snapshot: true,
      });
    }

    for (const grp of asArray<XmlNode>(node['p:grpSp'] as XmlNode | XmlNode[] | undefined)) {
      const groupTransform = groupTransformFromGrpSpPr(grp['p:grpSpPr'] as XmlNode | undefined);
      const nextTransforms = groupTransform ? [...transforms, groupTransform] : transforms;
      collectVisualsRecursive(owner, grp, nextTransforms, template);
    }
  }

  const hides = (p: VisualPart | null) => p?.root['@_showMasterSp'] === '0' || p?.root['@_showMasterSp'] === 'false';
  const spTreeOf = (p: VisualPart) => p.cSld['p:spTree'] as XmlNode | undefined;
  if (!hides(slide)) {
    if (master && !hides(layout)) {
      const tree = spTreeOf(master);
      partRank = 1;
      if (tree) collectVisualsRecursive(master, tree, [], true);
    }
    if (layout) {
      const tree = spTreeOf(layout);
      partRank = 2;
      if (tree) collectVisualsRecursive(layout, tree, [], true);
    }
  }
  const slideTree = spTreeOf(slide);
  partRank = 3;
  if (slideTree) collectVisualsRecursive(slide, slideTree, [], false);

  // 배경(정렬 키 없음)은 맨 아래, 나머지는 마스터 → 레이아웃 → 슬라이드, 각 파트 안에서는 원본 순서
  return visuals
    .map((visual, index) => ({ visual, index, key: sortKey.get(visual) ?? -1 }))
    .sort((a, b) => a.key - b.key || a.index - b.index)
    .map((entry) => entry.visual);
}
