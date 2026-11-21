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

// htmlEntities: 테마 글꼴 이름의 숫자 문자 참조(&#xB9D1;)를 푼다 — slideContent.ts와 같은 이유
const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_', parseTagValue: false, htmlEntities: true, isArray: (n) => n === 'Relationship' });
const decoder = new TextDecoder();

/** 지정 색 이름(prstClr) 중 실제 문서에 자주 나오는 것들 */
const PRESET: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00',
  gray: '808080', darkGray: 'A9A9A9', lightGray: 'D3D3D3', orange: 'FFA500', purple: '800080',
};

export interface ThemeColors {
  /** 색 노드(<a:solidFill>나 <a:fillRef>의 내용 — srgbClr/schemeClr/sysClr/prstClr 중 하나를 자식으로
   * 가진 노드)를 '#RRGGBB'로 푼다. 풀 수 없으면 undefined. ph는 스타일 행렬(fmtScheme) 안의 phClr
   * 자리에 들어갈 색('#RRGGBB' — 그 스타일을 부른 fillRef/lnRef/effectRef의 색). */
  resolve(container: XmlNode | undefined, ph?: string): string | undefined;
  /** 테마 서식 행렬(fmtScheme) — 도형 스타일 <a:fillRef idx>·<a:lnRef idx>·<a:effectRef idx>가 가리키는 항목.
   * idx는 1부터, fillRef 1001 이상은 배경 채우기 목록(bgFillStyleLst). 없으면 undefined.
   * fillStyle은 채우기 요소를 자식으로 가진 노드(spPr처럼 {'a:gradFill': …}), lineStyle은 <a:ln>, effectStyle은 <a:effectStyle>. */
  fillStyle(idx: number): XmlNode | undefined;
  lineStyle(idx: number): XmlNode | undefined;
  effectStyle(idx: number): XmlNode | undefined;
  /** 테마 글꼴 참조(+mj-lt, +mn-ea …)를 실제 글꼴 이름으로. 테마 참조가 아니면 그대로, 풀 수 없으면 undefined.
   * 동아시아(ea) 칸이 비어 있으면 한글 글꼴(<a:font script="Hang">)을 쓴다 — PowerPoint가 한국어에 쓰는 것과 같다. */
  font(typeface: string | undefined): string | undefined;
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

/** .rels 파일에서 "rId → Target" 매핑만 뽑는다(src/pptx/collab.ts의 slideRelsMap과 같은 패턴). */
export function relsMap(relsXml: string): Map<string, string> {
  const map = new Map<string, string>();
  const relRe = /<Relationship\b[^>]*\/>/g;
  for (const relTag of relsXml.match(relRe) ?? []) {
    const idMatch = /\bId="([^"]+)"/.exec(relTag); // 관계 id는 rIdN이 흔하지만 아무 이름이나 된다
    const targetMatch = /Target="([^"]+)"/.exec(relTag);
    if (idMatch && targetMatch) map.set(idMatch[1], targetMatch[1]);
  }
  return map;
}

/** .rels의 Target(예: "../media/image1.png")을 슬라이드 파트 기준 zip 내부 절대 경로로 푼다. */
export function resolveRelativePath(basePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const stack = basePart.split('/').slice(0, -1);
  for (const part of target.split('/')) {
    if (part === '..') stack.pop();
    else if (part !== '.') stack.push(part);
  }
  return stack.join('/');
}

/** .rels의 관계 id → 관계 종류(Type URI의 마지막 이름: hyperlink·image·slide·audio …) */
export function relTypesOf(relsXml: string): Map<string, string> {
  const map = new Map<string, string>();
  for (const relTag of relsXml.match(/<Relationship\b[^>]*\/>/g) ?? []) {
    const id = /\bId="([^"]+)"/.exec(relTag)?.[1];
    const type = /\bType="([^"]+)"/.exec(relTag)?.[1];
    if (id && type) map.set(id, type.slice(type.lastIndexOf('/') + 1));
  }
  return map;
}

/** 파트(슬라이드·레이아웃·마스터) 안 관계 id(r:embed)를 그 파트의 rels로 따라가 브라우저가 그릴 수 있는 그림의 zip 경로로.
 * 관계가 없거나, 미디어 파트가 없거나(손상), 브라우저가 못 그리는 형식(EMF 등)이면 undefined. rels를 이미 읽었으면 넘긴다. */
export function mediaResolverFor(zip: Record<string, Uint8Array>, part: string, rels?: Map<string, string>): (rId: string) => string | undefined {
  let map = rels;
  return (rId) => {
    if (!map) {
      const slash = part.lastIndexOf('/');
      const bytes = zip[`${part.slice(0, slash)}/_rels/${part.slice(slash + 1)}.rels`];
      map = bytes ? relsMap(decoder.decode(bytes)) : new Map();
    }
    const target = map.get(rId);
    if (!target) return undefined;
    const normalized = resolveRelativePath(part, target);
    if (!zip[normalized]) return undefined; // 관계는 있는데 실제 미디어 파트가 없음(손상)
    const ext = normalized.split('.').pop()?.toLowerCase() ?? '';
    return IMAGE_MIME_BY_EXT[ext] ? normalized : undefined; // 브라우저가 못 그리는 포맷(emf/wmf 등)은 빼고
  };
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

const toLinear = (c: number) => {
  const v = c / 255;
  return v <= 0.04045 ? v / 12.92 : ((v + 0.055) / 1.055) ** 2.4;
};
const fromLinear = (v: number) => 255 * (v <= 0.0031308 ? v * 12.92 : 1.055 * v ** (1 / 2.4) - 0.055);

/** 색 노드의 자식 변환(tint/shade/satMod/lumMod/lumOff)을 적용한다. 값은 1/1000 퍼센트(100000=100%).
 * tint·shade는 PowerPoint처럼 선형 밝기(scRGB)에서 흰색·검정과 섞는다 — sRGB로 섞으면 "tx1 tint 75%"가
 * #404040(너무 진함)이 되는데 PowerPoint·LibreOffice는 #898989로 그린다(부제목 회색이 검게 보이던 원인). */
function applyTransforms(hex: string, node: XmlNode): string {
  type Rgb = [number, number, number];
  let rgb = hexToRgb(hex);
  const clamp = (v: number) => Math.max(0, Math.min(1, v));
  const hsl = (f: (h: number, s: number, l: number) => [number, number, number]) => {
    const [h, s, l] = rgbToHsl(rgb);
    const [h2, s2, l2] = f(h, s, l);
    rgb = hslToRgb([((h2 % 1) + 1) % 1, clamp(s2), clamp(l2)]);
  };
  const lin = (f: (c: number, i: number) => number) => {
    rgb = rgb.map((c, i) => fromLinear(clamp(f(toLinear(c), i)))) as Rgb;
  };
  // 변환은 파일에 적힌 순서대로 적용한다(OOXML) — 파서가 담은 키 순서가 문서 순서다(같은 이름이 두 번 나오는 일은 없다).
  // 값은 1/1000 퍼센트(100000=100%), 색상(hue)은 1/60000도.
  for (const key of Object.keys(node)) {
    const raw = (node[key] as XmlNode | undefined)?.['@_val'];
    const v = raw === undefined ? NaN : Number(raw) / 100000;
    switch (key) {
      case 'a:tint': lin((c) => c * v + (1 - v)); break;
      case 'a:shade': lin((c) => c * v); break;
      case 'a:satMod': hsl((h, s, l) => [h, s * v, l]); break;
      case 'a:satOff': hsl((h, s, l) => [h, s + v, l]); break;
      case 'a:sat': hsl((h, _s, l) => [h, v, l]); break;
      case 'a:lumMod': hsl((h, s, l) => [h, s, l * v]); break;
      case 'a:lumOff': hsl((h, s, l) => [h, s, l + v]); break;
      case 'a:lum': hsl((h, s) => [h, s, v]); break;
      case 'a:hueMod': hsl((h, s, l) => [h * v, s, l]); break;
      case 'a:hueOff': hsl((h, s, l) => [h + Number(raw) / 60000 / 360, s, l]); break;
      case 'a:hue': hsl((_h, s, l) => [Number(raw) / 60000 / 360, s, l]); break;
      case 'a:comp': hsl((h, s, l) => [h + 0.5, s, l]); break;
      case 'a:inv': rgb = rgb.map((c) => 255 - c) as Rgb; break;
      case 'a:gray': { const g = 0.3 * rgb[0] + 0.59 * rgb[1] + 0.11 * rgb[2]; rgb = [g, g, g]; break; }
      case 'a:gamma': rgb = rgb.map((c) => 255 * toLinear(c)) as Rgb; break;
      case 'a:invGamma': rgb = rgb.map((c) => fromLinear(c / 255)) as Rgb; break;
      case 'a:redMod': lin((c, i) => (i === 0 ? c * v : c)); break;
      case 'a:redOff': lin((c, i) => (i === 0 ? c + v : c)); break;
      case 'a:greenMod': lin((c, i) => (i === 1 ? c * v : c)); break;
      case 'a:greenOff': lin((c, i) => (i === 1 ? c + v : c)); break;
      case 'a:blueMod': lin((c, i) => (i === 2 ? c * v : c)); break;
      case 'a:blueOff': lin((c, i) => (i === 2 ? c + v : c)); break;
      default: break; // alpha 등 색을 바꾸지 않는 값
    }
  }
  return rgbToHex(rgb);
}

/** fmtScheme 목록(fillStyleLst 등)의 자식 요소를 문서 순서대로 — 파서는 이름이 다른 형제(solidFill·gradFill)를
 * 이름별로 묶어 순서를 잃으므로 원문에서 순서대로 잘라 하나씩 읽는다. */
function styleList(themeXml: string, listTag: string, childRe: string): XmlNode[] {
  const list = new RegExp(`<a:${listTag}\\b[^>]*>([\\s\\S]*?)</a:${listTag}>`).exec(themeXml)?.[1] ?? '';
  const out: XmlNode[] = [];
  for (const m of list.matchAll(new RegExp(`<a:(${childRe})\\b(?:[^>]*/>|[\\s\\S]*?</a:\\1>)`, 'g'))) {
    try {
      out.push(parser.parse(m[0]) as XmlNode);
    } catch {
      out.push({});
    }
  }
  return out;
}

/** 슬라이드에 적용되는 테마 색표(clrMap 별칭 포함). 테마를 못 찾으면 Office 기본 테마로 대체한다. */
export function loadThemeColors(zip: Record<string, Uint8Array>, sourceSlidePart: string): ThemeColors {
  const slidePart = `ppt/slides/${sourceSlidePart}`;
  const layoutPart = relTargetOf(zip, slidePart, 'slideLayout');
  const masterPart = layoutPart ? relTargetOf(zip, layoutPart, 'slideMaster') : null;
  const master = masterPart ? readXml(zip, masterPart) : null;
  const themePart = masterPart ? relTargetOf(zip, masterPart, 'theme') : null;
  const theme = themePart ? readXml(zip, themePart) : null;
  const themeXml = themePart && zip[themePart] ? decoder.decode(zip[themePart]) : '';
  const FILLS = 'noFill|solidFill|gradFill|blipFill|pattFill|grpFill';
  const fills = styleList(themeXml, 'fillStyleLst', FILLS);
  const bgFills = styleList(themeXml, 'bgFillStyleLst', FILLS);
  const lines = styleList(themeXml, 'lnStyleLst', 'ln');
  const effects = styleList(themeXml, 'effectStyleLst', 'effectStyle');
  const fontScheme = ((theme?.['a:theme'] as XmlNode | undefined)?.['a:themeElements'] as XmlNode | undefined)?.['a:fontScheme'] as XmlNode | undefined;
  const slotFont = (slot: 'majorFont' | 'minorFont', script: 'latin' | 'ea' | 'cs'): string | undefined => {
    const group = fontScheme?.[`a:${slot}`] as XmlNode | undefined;
    const direct = (group?.[`a:${script}`] as XmlNode | undefined)?.['@_typeface'];
    if (typeof direct === 'string' && direct) return direct;
    if (script !== 'ea') return undefined;
    const scripts = group?.['a:font'];
    const list = (Array.isArray(scripts) ? scripts : scripts ? [scripts] : []) as XmlNode[];
    const hang = list.find((f) => f['@_script'] === 'Hang')?.['@_typeface'];
    return typeof hang === 'string' && hang ? hang : undefined;
  };

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
    fillStyle: (idx) => (idx >= 1001 ? bgFills[idx - 1001] : fills[idx - 1]),
    lineStyle: (idx) => lines[idx - 1]?.['a:ln'] as XmlNode | undefined,
    effectStyle: (idx) => effects[idx - 1]?.['a:effectStyle'] as XmlNode | undefined,
    font(typeface) {
      if (!typeface) return undefined;
      const m = /^\+(mj|mn)-(lt|ea|cs)$/.exec(typeface);
      if (!m) return typeface.startsWith('+') ? undefined : typeface;
      return slotFont(m[1] === 'mj' ? 'majorFont' : 'minorFont', m[2] === 'lt' ? 'latin' : m[2] === 'ea' ? 'ea' : 'cs');
    },
    resolve(container, ph) {
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
        // phClr = "자리표시 색"(스타일 행렬 안에서만 쓰임) — 그 스타일을 부른 참조의 색(ph)이 들어간다
        const val = String(schemeClr['@_val'] ?? '');
        base = val === 'phClr' ? ph?.replace('#', '') : scheme[alias(val)];
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
