import { XMLParser } from 'fast-xml-parser';

/**
 * 슬라이드 도형의 "보이는 모양" 중 슬라이드 XML이 아니라 레이아웃·마스터에서 물려받는 값을 푼다 —
 * 편집 캔버스 표시 전용(저장 경로는 이 값을 쓰지 않는다: 원본 도형은 txBody만 교체하므로).
 *
 * 실제 문서의 대부분 슬라이드는 제목/본문 "개체 틀(placeholder)"을 쓰고, 그 경우 슬라이드 XML에는
 * <p:ph type="title"/>만 있고 위치·크기·글자 크기·세로 맞춤·정렬은 레이아웃(없으면 마스터)에
 * 있다. 예전엔 이를 읽지 않아 개체 틀이 임의 위치(가로 20%·세로 40%)에 18pt로 그려졌다.
 *
 * 우선순위(PowerPoint와 동일):
 * - 위치/크기:   슬라이드 → 레이아웃 개체 틀 → 마스터 개체 틀
 * - 세로 맞춤:   슬라이드 bodyPr → 레이아웃 → 마스터 → 위쪽
 * - 1수준 글자 크기·정렬: 슬라이드 lstStyle → 레이아웃 → 마스터 개체 틀 lstStyle →
 *   마스터 txStyles(제목/본문/기타) → presentation.xml defaultTextStyle
 */

export type Anchor = 't' | 'ctr' | 'b';
export type Align = 'l' | 'ctr' | 'r' | 'just';

/** 목록 수준(0~8) 하나의 보이는 모양 — 글머리 기호·들여쓰기·글자 크기. */
export interface LevelLook {
  /** 왼쪽 여백(pt). */
  readonly marLPt: number;
  /** 첫 줄 들여쓰기(pt, 음수면 내어쓰기 — 글머리 기호가 그 자리에 온다). */
  readonly indentPt: number;
  /** 글머리 문자. null = 글머리 없음. */
  readonly bullet: string | null;
  readonly sizePt: number;
}

export interface InheritedShapeLook {
  readonly x: number | null;
  readonly y: number | null;
  readonly w: number | null;
  readonly h: number | null;
  readonly anchor: Anchor;
  /** 1수준 문단의 기본 글자 크기(pt) — 글자에 크기 서식이 없을 때 보이는 크기. */
  readonly defaultSizePt: number;
  /** 1수준 문단의 기본 정렬 — 문단에 정렬이 없을 때 보이는 정렬. */
  readonly defaultAlign: Align;
  /** 수준 0~8(lvl1pPr~lvl9pPr)의 글머리·들여쓰기·크기. */
  readonly levels: readonly LevelLook[];
}

type XmlNode = Record<string, unknown>;

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  isArray: (name) => ['p:sp', 'p:grpSp', 'Relationship'].includes(name),
});

const decoder = new TextDecoder();

function asArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined) return [];
  return Array.isArray(value) ? value : [value];
}

function readXml(zip: Record<string, Uint8Array>, part: string): XmlNode | null {
  const bytes = zip[part];
  if (!bytes) return null;
  try {
    return parser.parse(decoder.decode(bytes)) as XmlNode;
  } catch {
    return null;
  }
}

/** rels에서 Type이 `…/${typeSuffix}`인 첫 관계의 대상 파트 경로(zip 기준 절대 경로). */
function relTarget(zip: Record<string, Uint8Array>, part: string, typeSuffix: string): string | null {
  const slash = part.lastIndexOf('/');
  const dir = part.slice(0, slash);
  const rels = readXml(zip, `${dir}/_rels/${part.slice(slash + 1)}.rels`);
  const list = asArray((rels?.Relationships as XmlNode | undefined)?.Relationship as XmlNode | XmlNode[] | undefined);
  const rel = list.find((r) => String(r['@_Type'] ?? '').endsWith(`/${typeSuffix}`));
  if (!rel) return null;
  const segments = `${dir}/${String(rel['@_Target'])}`.split('/');
  const out: string[] = [];
  for (const s of segments) {
    if (s === '..') out.pop();
    else if (s !== '.' && s !== '') out.push(s);
  }
  return out.join('/');
}

interface PlaceholderInfo {
  readonly type: string; // 원본 type, 없으면 'obj'(스키마 기본값)
  readonly idx: string | undefined;
}

interface ShapeLook {
  readonly ph: PlaceholderInfo | null;
  readonly id: string;
  readonly x: number | null;
  readonly y: number | null;
  readonly w: number | null;
  readonly h: number | null;
  readonly anchor: Anchor | undefined;
  readonly sizePt: number | undefined;
  readonly align: Align | undefined;
  readonly levels: readonly PartialLevel[];
}

const ANCHORS = new Set(['t', 'ctr', 'b']);
const ALIGNS = new Set(['l', 'ctr', 'r', 'just']);

function num(value: unknown): number | null {
  const n = typeof value === 'number' ? value : value === undefined ? NaN : Number(value);
  return Number.isFinite(n) ? n : null;
}

/** 한 수준(a:lvlNpPr)에서 실제로 "적혀 있는" 값만 — 없는 값은 undefined로 두어 상위로 넘긴다. */
interface PartialLevel {
  readonly sizePt?: number;
  readonly align?: Align;
  readonly marLPt?: number;
  readonly indentPt?: number;
  /** 글머리: 문자열 = 그 문자, null = <a:buNone/>, undefined = 언급 없음 */
  readonly bullet?: string | null;
}

const EMU_PER_PT = 12700;

/** 문단 속성 노드(<a:lvlNpPr> 또는 <a:pPr>)에서 이 모듈이 다루는 값만 뽑는다. */
export function partialLevelFrom(node: XmlNode | undefined): PartialLevel {
  if (!node || typeof node !== 'object') return {};
  const sz = num((node['a:defRPr'] as XmlNode | undefined)?.['@_sz']);
  const algn = node['@_algn'] as string | undefined;
  const marL = num(node['@_marL']);
  const indent = num(node['@_indent']);
  const buChar = (node['a:buChar'] as XmlNode | undefined)?.['@_char'];
  const bullet =
    node['a:buNone'] !== undefined ? null : buChar !== undefined ? String(buChar) : node['a:buAutoNum'] !== undefined ? '•' : undefined;
  return {
    ...(sz !== null ? { sizePt: sz / 100 } : {}),
    ...(algn && ALIGNS.has(algn) ? { align: algn as Align } : {}),
    ...(marL !== null ? { marLPt: marL / EMU_PER_PT } : {}),
    ...(indent !== null ? { indentPt: indent / EMU_PER_PT } : {}),
    ...(bullet !== undefined ? { bullet } : {}),
  };
}

/** <a:lstStyle>/<p:titleStyle> 등에서 수준 1~9 */
function levelsOf(styleNode: XmlNode | undefined): PartialLevel[] {
  return Array.from({ length: 9 }, (_, i) => partialLevelFrom(styleNode?.[`a:lvl${i + 1}pPr`] as XmlNode | undefined));
}


function shapeLooks(root: XmlNode | null, rootTag: string): ShapeLook[] {
  const spTree = (((root?.[rootTag] as XmlNode | undefined)?.['p:cSld'] as XmlNode | undefined)?.['p:spTree']) as
    | XmlNode
    | undefined;
  if (!spTree) return [];
  return asArray(spTree['p:sp'] as XmlNode | XmlNode[] | undefined).map((sp) => {
    const nvSpPr = sp['p:nvSpPr'] as XmlNode | undefined;
    const phNode = (nvSpPr?.['p:nvPr'] as XmlNode | undefined)?.['p:ph'] as XmlNode | string | undefined;
    // <p:ph/>처럼 속성이 하나도 없으면 파서가 빈 문자열을 준다 — 그것도 "개체 틀 있음"이다.
    const ph: PlaceholderInfo | null =
      phNode === undefined
        ? null
        : {
            type: typeof phNode === 'object' && phNode['@_type'] ? String(phNode['@_type']) : 'obj',
            idx: typeof phNode === 'object' && phNode['@_idx'] !== undefined ? String(phNode['@_idx']) : undefined,
          };
    const xfrm = (sp['p:spPr'] as XmlNode | undefined)?.['a:xfrm'] as XmlNode | undefined;
    const off = xfrm?.['a:off'] as XmlNode | undefined;
    const ext = xfrm?.['a:ext'] as XmlNode | undefined;
    const txBody = sp['p:txBody'] as XmlNode | undefined;
    const bodyPr = txBody?.['a:bodyPr'] as XmlNode | undefined;
    const anchorRaw = typeof bodyPr === 'object' ? (bodyPr['@_anchor'] as string | undefined) : undefined;
    const levels = levelsOf(txBody?.['a:lstStyle'] as XmlNode | undefined);
    const { sizePt, align } = levels[0];
    return {
      levels,
      ph,
      id: String((nvSpPr?.['p:cNvPr'] as XmlNode | undefined)?.['@_id'] ?? ''),
      x: num(off?.['@_x']),
      y: num(off?.['@_y']),
      w: num(ext?.['@_cx']),
      h: num(ext?.['@_cy']),
      anchor: anchorRaw && ANCHORS.has(anchorRaw) ? (anchorRaw as Anchor) : undefined,
      sizePt,
      align,
    };
  });
}

/** 제목 계열/본문 계열/기타 — 마스터 개체 틀과 txStyles를 고를 때 쓰는 분류 */
function phClass(type: string): 'title' | 'body' | 'other' {
  if (type === 'title' || type === 'ctrTitle') return 'title';
  if (type === 'body' || type === 'subTitle' || type === 'obj') return 'body';
  return 'other';
}

/** 레이아웃에서 짝 찾기: idx가 있으면 idx로, 없으면 type으로(제목/가운데 제목은 서로 대체 가능). */
function matchLayout(ph: PlaceholderInfo, candidates: readonly ShapeLook[]): ShapeLook | undefined {
  if (ph.idx !== undefined) {
    const byIdx = candidates.find((c) => c.ph && c.ph.idx === ph.idx);
    if (byIdx) return byIdx;
  }
  return (
    candidates.find((c) => c.ph && c.ph.type === ph.type) ??
    candidates.find((c) => c.ph && phClass(c.ph.type) === 'title' && phClass(ph.type) === 'title')
  );
}

function matchMaster(ph: PlaceholderInfo, candidates: readonly ShapeLook[]): ShapeLook | undefined {
  const cls = phClass(ph.type);
  return (
    candidates.find((c) => c.ph && c.ph.type === ph.type) ??
    candidates.find((c) => c.ph && phClass(c.ph.type) === cls && cls !== 'other')
  );
}

/**
 * 슬라이드 파트(예: "slide3.xml") 안 각 텍스트 도형(cNvPr id 기준)의 상속까지 반영한 보이는 모양.
 * 레이아웃/마스터가 없거나 손상돼도 던지지 않는다 — 찾을 수 있는 만큼만 채우고 나머지는 기본값.
 */
export function resolveSlideShapeLooks(
  zip: Record<string, Uint8Array>,
  sourceSlidePart: string,
): Map<string, InheritedShapeLook> {
  const slidePart = `ppt/slides/${sourceSlidePart}`;
  const slideShapes = shapeLooks(readXml(zip, slidePart), 'p:sld');
  const layoutPart = relTarget(zip, slidePart, 'slideLayout');
  const layoutShapes = layoutPart ? shapeLooks(readXml(zip, layoutPart), 'p:sldLayout') : [];
  const masterPart = layoutPart ? relTarget(zip, layoutPart, 'slideMaster') : null;
  const masterRoot = masterPart ? readXml(zip, masterPart) : null;
  const masterShapes = shapeLooks(masterRoot, 'p:sldMaster');
  const txStyles = (masterRoot?.['p:sldMaster'] as XmlNode | undefined)?.['p:txStyles'] as XmlNode | undefined;
  const classLevels = {
    title: levelsOf(txStyles?.['p:titleStyle'] as XmlNode | undefined),
    body: levelsOf(txStyles?.['p:bodyStyle'] as XmlNode | undefined),
    other: levelsOf(txStyles?.['p:otherStyle'] as XmlNode | undefined),
  };
  const classStyle = { title: classLevels.title[0], body: classLevels.body[0], other: classLevels.other[0] };
  const presentation = readXml(zip, 'ppt/presentation.xml');
  const defaultLevels = levelsOf(
    (presentation?.['p:presentation'] as XmlNode | undefined)?.['p:defaultTextStyle'] as XmlNode | undefined,
  );
  const defaultText = defaultLevels[0];

  const result = new Map<string, InheritedShapeLook>();
  for (const shape of slideShapes) {
    const layout = shape.ph ? matchLayout(shape.ph, layoutShapes) : undefined;
    const master = shape.ph ? matchMaster(shape.ph, masterShapes) : undefined;
    const cls = shape.ph ? phClass(shape.ph.type) : 'other';
    const chain = [shape, layout, master];
    const pick = <K extends 'anchor' | 'sizePt' | 'align'>(key: K) => chain.find((s) => s?.[key] !== undefined)?.[key];
    // 위치는 네 값을 한 묶음으로 물려받는다(슬라이드가 위치만 바꾸고 크기는 안 쓰는 경우는 없다).
    const positioned = chain.find((s) => s && s.x !== null && s.y !== null && s.w !== null && s.h !== null);
    result.set(shape.id, {
      x: positioned?.x ?? null,
      y: positioned?.y ?? null,
      w: positioned?.w ?? null,
      h: positioned?.h ?? null,
      anchor: (pick('anchor') as Anchor | undefined) ?? 't',
      defaultSizePt:
        (pick('sizePt') as number | undefined) ?? classStyle[cls].sizePt ?? defaultText.sizePt ?? 18,
      defaultAlign: (pick('align') as Align | undefined) ?? classStyle[cls].align ?? defaultText.align ?? 'l',
      levels: Array.from({ length: 9 }, (_, i): LevelLook => {
        const sources = [shape.levels[i], layout?.levels[i], master?.levels[i], classLevels[cls][i], defaultLevels[i]];
        const first = <K extends keyof PartialLevel>(k: K) => sources.find((s) => s && s[k] !== undefined)?.[k];
        return {
          marLPt: (first('marLPt') as number | undefined) ?? 0,
          indentPt: (first('indentPt') as number | undefined) ?? 0,
          bullet: (first('bullet') as string | null | undefined) ?? null,
          sizePt: (first('sizePt') as number | undefined) ?? 18,
        };
      }),
    });
  }
  return result;
}
