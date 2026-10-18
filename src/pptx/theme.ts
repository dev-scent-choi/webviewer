import { XMLParser } from 'fast-xml-parser';

/**
 * pptx 테마 색 해석 — 편집 캔버스 표시 전용.
 *
 * 실제 슬라이드의 색은 대부분 직접 지정(srgbClr)이 아니라 테마 참조다: 글자 <a:schemeClr val="tx1">
 * + 밝기 조정(<a:lumMod>/<a:lumOff>/<a:tint>/<a:shade>), 도형 스타일 <a:fillRef><a:schemeClr
 * val="accent1"/>, 배경 <p:bgRef>. 예전엔 이걸 못 풀어서 테마 색 글자는 검정, 테마 색 도형·배경은
 * 아예 빈 채로 그려졌다(부제목이 원본의 회색 대신 검정으로 보이는 것을 LibreOffice 렌더링과 비교해 확인).
 *
 * 해석 경로: 슬라이드 → 레이아웃 → 마스터(<p:clrMap>: tx1→dk1 등 별칭) → 마스터의 테마
 * (<a:clrScheme>: dk1/lt1/accent1… 실제 RGB).
 */

type XmlNode = Record<string, unknown>;

const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', parseTagValue: false, isArray: (n) => n === 'Relationship' });
const decoder = new TextDecoder();

/** 지정 색 이름(prstClr) 중 실제 문서에 자주 나오는 것들 */
const PRESET: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00',
  gray: '808080', darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500', purple: '800080',
};

export interface ThemeColors {
  /** 색 노드(<a:solidFill>나 <a:fillRef>의 내용 — srgbClr/schemeClr/sysClr/prstClr 중 하나를 자식으로
   * 가진 노드)를 '#RRGGBB'로 푼다. 풀 수 없으면 undefined. */
  resolve(container: XmlNode | undefined): string | undefined;
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

export function relTargetOf(zip: Record<string, Uint8Array>, part: string, typeSuffix: string): string | null {
  const slash = part.lastIndexOf('/');
  const dir = part.slice(0, slash);
  const rels = readXml(zip, `${dir}/_rels/${part.slice(slash + 1)}.rels`);
  const list = ((rels?.Relationships as XmlNode | undefined)?.Relationship as XmlNode[] | undefined) ?? [];
  const rel = list.find((r) => String(r['@_Type'] ?? '').endsWith(`/${typeSuffix}`));
  if (!rel) return null;
  const out: string[] = [];
  for (const s of `${dir}/${String(rel['@_Target'])}`.split('/')) {
    if (s === '..') out.pop();
    else if (s !== '.' && s !== '') out.push(s);
  }
  return out.join('/');
}

/* ---------- 색 변환(DrawingML) ---------- */

function hexToRgb(hex: string): [number, number, number] {
  const n = parseInt(hex, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
function rgbToHex([r, g, b]: [number, number, number]): string {
  const c = (v: number) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0');
  return `${c(r)}${c(g)}${c(b)}`.toUpperCase();
}
function rgbToHsl([r, g, b]: [number, number, number]): [number, number, number] {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}
function hslToRgb([h, s, l]: [number, number, number]): [number, number, number] {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t: number) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}

/** 색 노드의 자식 변환(lumMod/lumOff/tint/shade)을 순서대로 적용한다. 값은 1/1000 퍼센트(100000=100%). */
function applyTransforms(hex: string, node: XmlNode): string {
  let rgb = hexToRgb(hex);
  const pct = (name: string) => {
    const v = (node[name] as XmlNode | undefined)?.['@_val'];
    return v === undefined ? undefined : Number(v) / 100000;
  };
  const tint = pct('a:tint');
  if (tint !== undefined) rgb = rgb.map((c) => c * tint + 255 * (1 - tint)) as [number, number, number];
  const shade = pct('a:shade');
  if (shade !== undefined) rgb = rgb.map((c) => c * shade) as [number, number, number];
  const lumMod = pct('a:lumMod');
  const lumOff = pct('a:lumOff');
  if (lumMod !== undefined || lumOff !== undefined) {
    const [h, s, l] = rgbToHsl(rgb);
    rgb = hslToRgb([h, s, Math.max(0, Math.min(1, l * (lumMod ?? 1) + (lumOff ?? 0)))]);
  }
  return rgbToHex(rgb);
}

/** 슬라이드에 적용되는 테마 색표(clrMap 별칭 포함). 테마를 못 찾으면 Office 기본 테마로 대체한다. */
export function loadThemeColors(zip: Record<string, Uint8Array>, sourceSlidePart: string): ThemeColors {
  const slidePart = `ppt/slides/${sourceSlidePart}`;
  const layoutPart = relTargetOf(zip, slidePart, 'slideLayout');
  const masterPart = layoutPart ? relTargetOf(zip, layoutPart, 'slideMaster') : null;
  const master = masterPart ? readXml(zip, masterPart) : null;
  const themePart = masterPart ? relTargetOf(zip, masterPart, 'theme') : null;
  const theme = themePart ? readXml(zip, themePart) : null;

  // 실제 테마 파트에 적힌 값만 쓴다 — 테마를 못 찾았을 때 Office 기본 색표로 "추측"하면 원본과 다른
  // 색을 그럴듯하게 그려 버린다(테마 참조 색은 그냥 풀지 않는 편이 낫다 — 기존 정책과 같음).
  const scheme: Record<string, string> = {};
  const SCHEME_NAMES = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const clrScheme = (((theme?.['a:theme'] as XmlNode | undefined)?.['a:themeElements'] as XmlNode | undefined)?.[
    'a:clrScheme'
  ]) as XmlNode | undefined;
  for (const name of SCHEME_NAMES) {
    const entry = clrScheme?.[`a:${name}`] as XmlNode | undefined;
    const srgb = (entry?.['a:srgbClr'] as XmlNode | undefined)?.['@_val'];
    const sys = (entry?.['a:sysClr'] as XmlNode | undefined)?.['@_lastClr'];
    const value = srgb ?? sys;
    if (typeof value === 'string' && /^[0-9A-Fa-f]{6}$/.test(value)) scheme[name] = value.toUpperCase();
  }
  // 마스터 clrMap: tx1/bg1 등 별칭 → 실제 스킴 이름. 없으면 표준 매핑.
  const clrMapNode = (master?.['p:sldMaster'] as XmlNode | undefined)?.['p:clrMap'] as XmlNode | undefined;
  const alias = (name: string): string => {
    const mapped = clrMapNode?.[`@_${name}`];
    if (typeof mapped === 'string') return mapped;
    return ({ tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' } as Record<string, string>)[name] ?? name;
  };

  return {
    resolve(container) {
      if (!container || typeof container !== 'object') return undefined;
      const srgb = container['a:srgbClr'] as XmlNode | undefined;
      const schemeClr = container['a:schemeClr'] as XmlNode | undefined;
      const sysClr = container['a:sysClr'] as XmlNode | undefined;
      const prstClr = container['a:prstClr'] as XmlNode | undefined;
      let base: string | undefined;
      let node: XmlNode | undefined;
      if (srgb && typeof srgb === 'object') {
        base = String(srgb['@_val'] ?? '');
        node = srgb;
      } else if (schemeClr && typeof schemeClr === 'object') {
        // phClr = "자리표시 색"(스타일 행렬 안에서만 쓰임) — 이 앱 표시엔 의미 없어 풀지 않는다
        base = scheme[alias(String(schemeClr['@_val'] ?? ''))];
        node = schemeClr;
      } else if (sysClr && typeof sysClr === 'object') {
        base = String(sysClr['@_lastClr'] ?? '');
        node = sysClr;
      } else if (prstClr && typeof prstClr === 'object') {
        base = PRESET[String(prstClr['@_val'] ?? '')];
        node = prstClr;
      }
      if (!base || !/^[0-9A-Fa-f]{6}$/.test(base) || !node) return undefined;
      return `#${applyTransforms(base.toUpperCase(), node)}`;
    },
  };
}
