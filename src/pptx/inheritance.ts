import { XMLParser } from 'fast-xml-parser';
import { loadThemeColors } from './theme.js';
import { DEFAULT_INSETS_EMU as DEFAULT_INSETS, type ParagraphSpacing, type SlideLevelLook } from '../collab/pptxShared.js';

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

/** 목록 수준(0~8) 하나의 보이는 모양 — 편집 화면과 같은 타입(pptxShared). */
export type LevelLook = SlideLevelLook;

export interface InheritedShapeLook {
  readonly x: number | null;
  readonly y: number | null;
  readonly w: number | null;
  readonly h: number | null;
  readonly anchor: Anchor;
  /** bodyPr wrap="none" — 상자 너비에서 줄을 바꾸지 않는다(PowerPoint의 "도형의 텍스트 줄 바꿈" 끔). */
  readonly wrapNone: boolean;
  /** 글상자 안쪽 여백(EMU) [왼, 위, 오른, 아래] — bodyPr lIns/tIns/rIns/bIns, 없으면 기본 0.1"/0.05". */
  readonly insets: readonly [number, number, number, number];
  /** 1수준 문단의 기본 글자 크기(pt) — 글자에 크기 서식이 없을 때 보이는 크기. */
  readonly defaultSizePt: number;
  /** 1수준 문단의 기본 정렬 — 문단에 정렬이 없을 때 보이는 정렬. */
  readonly defaultAlign: Align;
  /** 수준 0~8(lvl1pPr~lvl9pPr)의 글머리·들여쓰기·크기. */
  readonly levels: readonly LevelLook[];
  /** 자동 맞춤(bodyPr normAutofit) — 글자 배율·줄 간격 줄임(0~1). 없으면 맞춤 없음. */
  readonly autofit?: { readonly fontScale: number; readonly lnSpcReduction: number };
  /** 세로 쓰기(bodyPr vert: vert·vert270·eaVert·wordArtVert·mongolianVert·wordArtVertRtl). 가로면 없음. */
  readonly vert?: string;
  /** 하이퍼링크 글자 색(테마 hlink) */
  readonly hlinkColor?: string;
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
  readonly wrap: 'none' | 'square' | undefined;
  readonly lIns: number | undefined;
  readonly tIns: number | undefined;
  readonly rIns: number | undefined;
  readonly bIns: number | undefined;
  readonly sizePt: number | undefined;
  readonly align: Align | undefined;
  readonly levels: readonly PartialLevel[];
  /** 도형 스타일 <p:style><a:fontRef>의 색 노드 — 기본 도형의 글자 색(대개 흰색 lt1)이 여기서 온다 */
  readonly fontRefColor?: XmlNode;
  /** 도형 스타일 <p:style><a:fontRef idx="major|minor"> — 글꼴 지정이 없을 때 쓰는 테마 글꼴 */
  readonly fontRefIdx?: string;
  /** bodyPr 자동 맞춤: undefined = 언급 없음(상속), null = 맞춤 없음(noAutofit/spAutoFit) */
  readonly autofit: { readonly fontScale: number; readonly lnSpcReduction: number } | null | undefined;
  readonly vert: string | undefined;
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
  /** <a:defRPr><a:solidFill>…</a:solidFill> 원본 노드 — 테마를 알아야 풀리므로 노드째 들고 간다 */
  readonly colorFill?: Record<string, unknown>;
  /** 자동 번호 — bullet과 같은 "글머리 종류" 칸이라 bullet이 적힌 수준에선 함께 정해진다 */
  readonly autoNum?: { readonly scheme: string; readonly startAt: number } | null;
  readonly line?: { readonly pct: number } | { readonly pts: number };
  readonly before?: { readonly pct: number } | { readonly pts: number };
  readonly after?: { readonly pct: number } | { readonly pts: number };
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly caps?: boolean;
  readonly eaLnBrk?: boolean;
  /** 원본 글꼴 이름(테마 참조 +mj-lt 그대로) — 테마를 알아야 풀린다 */
  readonly latin?: string;
  readonly ea?: string;
  readonly buFont?: string;
  readonly buClr?: Record<string, unknown>;
  readonly buSize?: { readonly pct: number } | { readonly pts: number };
}

/** <a:lnSpc>/<a:spcBef>/<a:spcAft> 하나 — spcPct(1/1000%) → pct 비율, spcPts(1/100pt) → pts */
function spacingOf(node: unknown): { pct: number } | { pts: number } | undefined {
  if (!node || typeof node !== 'object') return undefined;
  const pct = num(((node as XmlNode)['a:spcPct'] as XmlNode | undefined)?.['@_val']);
  if (pct !== null) return { pct: pct / 100000 };
  const pts = num(((node as XmlNode)['a:spcPts'] as XmlNode | undefined)?.['@_val']);
  return pts !== null ? { pts: pts / 100 } : undefined;
}

const onOff = (v: unknown): boolean | undefined => (v === undefined ? undefined : v === '1' || v === 'true');

const EMU_PER_PT = 12700;

/** 문단 속성 노드(<a:lvlNpPr> 또는 <a:pPr>)에서 이 모듈이 다루는 값만 뽑는다. */
export function partialLevelFrom(node: XmlNode | undefined): PartialLevel {
  if (!node || typeof node !== 'object') return {};
  const sz = num((node['a:defRPr'] as XmlNode | undefined)?.['@_sz']);
  const defRPrFill = (node['a:defRPr'] as XmlNode | undefined)?.['a:solidFill'];
  const algn = node['@_algn'] as string | undefined;
  const marL = num(node['@_marL']);
  const indent = num(node['@_indent']);
  const buChar = (node['a:buChar'] as XmlNode | undefined)?.['@_char'];
  const autoNumNode = node['a:buAutoNum'] as XmlNode | undefined;
  // 그림 글머리(buBlip)는 그림 대신 •로 보인다(그림 파일을 따라가지 않는다 — FIX-PPTX.md 참고)
  const bullet =
    node['a:buNone'] !== undefined ? null : buChar !== undefined ? String(buChar) : autoNumNode !== undefined || node['a:buBlip'] !== undefined ? '•' : undefined;
  const autoNum =
    autoNumNode !== undefined
      ? { scheme: String((autoNumNode as XmlNode)['@_type'] ?? 'arabicPeriod'), startAt: num((autoNumNode as XmlNode)['@_startAt']) ?? 1 }
      : bullet !== undefined
        ? null
        : undefined;
  const defRPr = typeof node['a:defRPr'] === 'object' ? (node['a:defRPr'] as XmlNode) : undefined;
  const face = (tag: string) => {
    const v = (defRPr?.[tag] as XmlNode | undefined)?.['@_typeface'];
    return typeof v === 'string' && v ? v : undefined;
  };
  const buSzPct = num((node['a:buSzPct'] as XmlNode | undefined)?.['@_val']);
  const buSzPts = num((node['a:buSzPts'] as XmlNode | undefined)?.['@_val']);
  const buClr = node['a:buClr'];
  const buFont = (node['a:buFont'] as XmlNode | undefined)?.['@_typeface'];
  const line = spacingOf(node['a:lnSpc']);
  const before = spacingOf(node['a:spcBef']);
  const after = spacingOf(node['a:spcAft']);
  const bold = onOff(defRPr?.['@_b']);
  const italic = onOff(defRPr?.['@_i']);
  const cap = defRPr?.['@_cap'];
  return {
    ...(autoNum !== undefined ? { autoNum } : {}),
    ...(line ? { line } : {}),
    ...(before ? { before } : {}),
    ...(after ? { after } : {}),
    ...(bold !== undefined ? { bold } : {}),
    ...(italic !== undefined ? { italic } : {}),
    ...(cap !== undefined ? { caps: cap === 'all' } : {}),
    ...(node['@_eaLnBrk'] !== undefined ? { eaLnBrk: onOff(node['@_eaLnBrk']) } : {}),
    ...(face('a:latin') ? { latin: face('a:latin') } : {}),
    ...(face('a:ea') ? { ea: face('a:ea') } : {}),
    ...(typeof buFont === 'string' && buFont ? { buFont } : {}),
    ...(typeof buClr === 'object' && buClr ? { buClr: buClr as XmlNode } : {}),
    ...(buSzPct !== null ? { buSize: { pct: buSzPct / 100000 } } : buSzPts !== null ? { buSize: { pts: buSzPts / 100 } } : {}),
    ...(sz !== null ? { sizePt: sz / 100 } : {}),
    ...(algn && ALIGNS.has(algn) ? { align: algn as Align } : {}),
    ...(marL !== null ? { marLPt: marL / EMU_PER_PT } : {}),
    ...(indent !== null ? { indentPt: indent / EMU_PER_PT } : {}),
    ...(bullet !== undefined ? { bullet } : {}),
    ...(typeof defRPrFill === 'object' && defRPrFill ? { colorFill: defRPrFill as XmlNode } : {}),
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
  // 그룹(<p:grpSp>, 중첩 포함) 안 도형도 모은다 — 실제 문서의 제목 상자가 그룹 안에 든 경우가 흔하다.
  const collect = (node: XmlNode): XmlNode[] => [
    ...asArray(node['p:sp'] as XmlNode | XmlNode[] | undefined),
    ...asArray(node['p:grpSp'] as XmlNode | XmlNode[] | undefined).flatMap(collect),
  ];
  return collect(spTree).map((sp) => {
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
    const wrapRaw = typeof bodyPr === 'object' ? (bodyPr['@_wrap'] as string | undefined) : undefined;
    const levels = levelsOf(txBody?.['a:lstStyle'] as XmlNode | undefined);
    const { sizePt, align } = levels[0];
    const inset = (key: string) => (typeof bodyPr === 'object' ? num(bodyPr[key]) ?? undefined : undefined);
    const fontRef = (sp['p:style'] as XmlNode | undefined)?.['a:fontRef'] as XmlNode | undefined;
    const norm = typeof bodyPr === 'object' ? bodyPr['a:normAutofit'] : undefined;
    const pct = (v: unknown, fallback: number) => (num(v) ?? fallback) / 100000;
    const autofit =
      norm !== undefined
        ? {
            fontScale: pct(typeof norm === 'object' ? (norm as XmlNode)['@_fontScale'] : undefined, 100000),
            lnSpcReduction: pct(typeof norm === 'object' ? (norm as XmlNode)['@_lnSpcReduction'] : undefined, 0),
          }
        : typeof bodyPr === 'object' && (bodyPr['a:noAutofit'] !== undefined || bodyPr['a:spAutoFit'] !== undefined)
          ? null
          : undefined;
    const vertRaw = typeof bodyPr === 'object' ? (bodyPr['@_vert'] as string | undefined) : undefined;
    return {
      levels,
      fontRefColor: fontRef,
      fontRefIdx: typeof fontRef?.['@_idx'] === 'string' ? (fontRef['@_idx'] as string) : undefined,
      autofit,
      vert: vertRaw,
      ph,
      id: String((nvSpPr?.['p:cNvPr'] as XmlNode | undefined)?.['@_id'] ?? ''),
      x: num(off?.['@_x']),
      y: num(off?.['@_y']),
      w: num(ext?.['@_cx']),
      h: num(ext?.['@_cy']),
      anchor: anchorRaw && ANCHORS.has(anchorRaw) ? (anchorRaw as Anchor) : undefined,
      wrap: wrapRaw === 'none' || wrapRaw === 'square' ? wrapRaw : undefined,
      lIns: inset('@_lIns'),
      tIns: inset('@_tIns'),
      rIns: inset('@_rIns'),
      bIns: inset('@_bIns'),
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

  const theme = loadThemeColors(zip, sourceSlidePart);
  const result = new Map<string, InheritedShapeLook>();
  for (const shape of slideShapes) {
    const layout = shape.ph ? matchLayout(shape.ph, layoutShapes) : undefined;
    const master = shape.ph ? matchMaster(shape.ph, masterShapes) : undefined;
    const cls = shape.ph ? phClass(shape.ph.type) : 'other';
    const chain = [shape, layout, master];
    const pick = <K extends 'anchor' | 'wrap' | 'sizePt' | 'align' | 'lIns' | 'tIns' | 'rIns' | 'bIns' | 'autofit' | 'vert'>(key: K) =>
      chain.find((s) => s?.[key] !== undefined)?.[key];
    const autofit = pick('autofit') as ShapeLook['autofit'];
    const vert = pick('vert') as string | undefined;
    const hlinkColor = theme.resolve({ 'a:schemeClr': { '@_val': 'hlink' } });
    // 위치는 네 값을 한 묶음으로 물려받는다(슬라이드가 위치만 바꾸고 크기는 안 쓰는 경우는 없다).
    const positioned = chain.find((s) => s && s.x !== null && s.y !== null && s.w !== null && s.h !== null);
    result.set(shape.id, {
      x: positioned?.x ?? null,
      y: positioned?.y ?? null,
      w: positioned?.w ?? null,
      h: positioned?.h ?? null,
      anchor: (pick('anchor') as Anchor | undefined) ?? 't',
      wrapNone: pick('wrap') === 'none',
      // 안쪽 여백은 네 값을 따로 물려받는다(lIns="0"만 쓴 도형은 위·아래·오른쪽은 기본값)
      insets: [
        (pick('lIns') as number | undefined) ?? DEFAULT_INSETS[0],
        (pick('tIns') as number | undefined) ?? DEFAULT_INSETS[1],
        (pick('rIns') as number | undefined) ?? DEFAULT_INSETS[2],
        (pick('bIns') as number | undefined) ?? DEFAULT_INSETS[3],
      ],
      defaultSizePt:
        (pick('sizePt') as number | undefined) ?? classStyle[cls].sizePt ?? defaultText.sizePt ?? 18,
      // 정렬이 도형·개체 틀 어디에도 없으면 마스터 txStyles(그린 도형·글상자는 otherStyle, 대개 왼쪽) —
      // PowerPoint가 도형을 새로 그릴 때는 algn="ctr"을 직접 적어 넣으므로, 적힌 것이 없으면 왼쪽이 맞다
      // (MS Office 정답 PDF p07: 정렬 없는 도형 글자 전부 왼쪽). 별·타원처럼 글자 칸이 도형보다 좁은 모양은
      // 편집기가 프리셋의 글자 칸(text rect)으로 여백을 줘서 글자가 도형 안에 놓인다.
      defaultAlign: (pick('align') as Align | undefined) ?? classStyle[cls].align ?? defaultText.align ?? 'l',
      ...(autofit && (autofit.fontScale < 1 || autofit.lnSpcReduction > 0) ? { autofit } : {}),
      ...(vert && vert !== 'horz' ? { vert } : {}),
      ...(hlinkColor ? { hlinkColor } : {}),
      levels: Array.from({ length: 9 }, (_, i): LevelLook => {
        const sources = [shape.levels[i], layout?.levels[i], master?.levels[i], classLevels[cls][i], defaultLevels[i]];
        const first = <K extends keyof PartialLevel>(k: K) => sources.find((s) => s && s[k] !== undefined)?.[k];
        // 글머리 "종류"(문자·없음·자동 번호)는 처음 언급한 수준에서 한꺼번에 정해진다
        const bulletSource = sources.find((s) => s && (s.bullet !== undefined || s.autoNum !== undefined));
        const autoNum = bulletSource?.autoNum ?? undefined;
        // 테마 글꼴: 수준 defRPr의 글꼴 → 도형 스타일 fontRef(major/minor) — 둘 다 없으면 편집기 기본(맑은 고딕)
        const refSlot = shape.fontRefIdx === 'major' ? 'mj' : shape.fontRefIdx === 'minor' ? 'mn' : undefined;
        const fontLatin = theme.font((first('latin') as string | undefined) ?? (refSlot ? `+${refSlot}-lt` : undefined));
        const fontEa = theme.font((first('ea') as string | undefined) ?? (refSlot ? `+${refSlot}-ea` : undefined));
        const line = first('line') as ParagraphSpacing['line'];
        const before = first('before') as ParagraphSpacing['before'];
        const after = first('after') as ParagraphSpacing['after'];
        const bold = first('bold') as boolean | undefined;
        const italic = first('italic') as boolean | undefined;
        const caps = first('caps') as boolean | undefined;
        const keepWords = first('eaLnBrk') === false;
        const bulletFont = first('buFont') as string | undefined;
        const bulletColor = theme.resolve(first('buClr') as XmlNode | undefined);
        const bulletSize = first('buSize') as LevelLook['bulletSize'];
        return {
          marLPt: (first('marLPt') as number | undefined) ?? 0,
          indentPt: (first('indentPt') as number | undefined) ?? 0,
          bullet: bulletSource?.bullet ?? null,
          ...(autoNum ? { autoNum } : {}),
          ...(line || before || after ? { spacing: { ...(line ? { line } : {}), ...(before ? { before } : {}), ...(after ? { after } : {}) } } : {}),
          ...(bold ? { bold } : {}),
          ...(italic ? { italic } : {}),
          ...(caps ? { caps } : {}),
          ...(keepWords ? { keepWords } : {}),
          ...(fontLatin ? { fontLatin } : {}),
          ...(fontEa ? { fontEa } : {}),
          ...(bulletFont ? { bulletFont: theme.font(bulletFont) ?? bulletFont } : {}),
          ...(bulletColor ? { bulletColor } : {}),
          ...(bulletSize ? { bulletSize } : {}),
          sizePt: (first('sizePt') as number | undefined) ?? 18,
          // 색 우선순위(PowerPoint와 같음): 도형·레이아웃·마스터 개체 틀의 목록 스타일 → 도형 스타일의
          // 글꼴 색(fontRef, 기본 도형의 흰 글자가 여기서 온다) → 마스터 txStyles/기본 텍스트 스타일.
          // fontRef를 txStyles보다 늦게 보면 마스터 otherStyle의 tx1(검정)이 이겨 파란 도형 위 글자가
          // 검정으로 보였다(실제 Office 템플릿 파일로 확인).
          color:
            theme.resolve(
              [shape.levels[i], layout?.levels[i], master?.levels[i]].find((s) => s?.colorFill)?.colorFill,
            ) ??
            theme.resolve(shape.fontRefColor) ??
            theme.resolve([classLevels[cls][i], defaultLevels[i]].find((s) => s?.colorFill)?.colorFill) ??
            null,
        };
      }),
    });
  }
  return result;
}
