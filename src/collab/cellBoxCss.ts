/**
 * 칸 상자 모양(테두리 선 종류·색, 무늬 채우기, 글자 회전·세로쓰기, 들여쓰기)을 CSS로 — 편집 화면(grid-editor.ts)이 칸마다,
 * 보기 화면은 서버가 서식 목록을 내려 줄 때(routes/viewer.ts) 같은 함수로 만들어 viewer.js가 그대로 붙인다.
 * DOM을 쓰지 않는다(서버·브라우저 공용).
 */

/** 엑셀 테두리 한 변: [선 종류(thin·dashed·double …), 색 '#RRGGBB'] */
export type BorderSpec = readonly [string, string?];

export interface CellBoxStyle {
  readonly bt?: 1 | 2;
  readonly br?: 1 | 2;
  readonly bb?: 1 | 2;
  readonly bl?: 1 | 2;
  /** 굵기(bt…)만으로 못 나타내는 변(색이 있거나 점선·이중선 등) */
  readonly bd?: { readonly t?: BorderSpec; readonly r?: BorderSpec; readonly b?: BorderSpec; readonly l?: BorderSpec };
  readonly bg?: string;
  /** 무늬 채우기: [무늬 이름, 무늬 색, 바탕색] — solid는 bg로만 둔다 */
  readonly pat?: readonly [string, string?, string?];
  /** 글자 회전(엑셀 textRotation: 1~90 반시계, 91~180 = 시계 방향 1~90, 255 = 세로쓰기) */
  readonly rot?: number;
  /** 들여쓰기 단계 */
  readonly ind?: number;
  readonly a?: 'l' | 'c' | 'r';
}

const BORDER_CSS: Readonly<Record<string, string>> = {
  hair: '1px dotted',
  thin: '1px solid',
  medium: '2px solid',
  thick: '3px solid',
  dotted: '1px dotted',
  dashed: '1px dashed',
  dashDot: '1px dashed',
  dashDotDot: '1px dashed',
  mediumDashed: '2px dashed',
  mediumDashDot: '2px dashed',
  mediumDashDotDot: '2px dashed',
  slantDashDot: '2px dashed',
  double: '3px double',
};

/** 한 변의 CSS border 값('' = 칸 서식 테두리 없음). 기본 색은 화면 글자 색(var(--v-text)) */
export function borderSideCss(style: CellBoxStyle | undefined, side: 't' | 'r' | 'b' | 'l'): string {
  const weight = style?.[`b${side}` as 'bt' | 'br' | 'bb' | 'bl'];
  if (!weight) return '';
  const spec = style?.bd?.[side];
  if (!spec) return `${weight}px solid var(--v-text)`;
  return `${BORDER_CSS[spec[0]] ?? `${weight}px solid`} ${spec[1] ?? 'var(--v-text)'}`;
}

function mix(a: string, b: string, t: number): string {
  const pa = [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16));
  const pb = [1, 3, 5].map((i) => parseInt(b.slice(i, i + 2), 16));
  return `#${pa.map((x, i) => Math.round(x + (pb[i] - x) * t).toString(16).padStart(2, '0')).join('')}`;
}

const GRAY_SHARE: Readonly<Record<string, number>> = { darkGray: 0.75, mediumGray: 0.5, lightGray: 0.25, gray125: 0.125, gray0625: 0.0625 };
const LINE_ANGLE: Readonly<Record<string, string>> = { Horizontal: '0deg', Vertical: '90deg', Down: '45deg', Up: '-45deg' };

/** 무늬 채우기 → background 계열 CSS(엑셀 무늬를 줄무늬·격자 그라데이션으로 근사) */
function patternCss(pat: readonly [string, string?, string?]): Record<string, string> {
  const [name, fgRaw, bgRaw] = pat;
  const fg = fgRaw ?? '#000000';
  const bg = bgRaw ?? '#ffffff';
  const gray = GRAY_SHARE[name];
  if (gray !== undefined) return { 'background-color': mix(bg, fg, gray), 'background-image': 'none' };
  const m = /^(dark|light)(Horizontal|Vertical|Down|Up|Grid|Trellis)$/.exec(name);
  if (!m) return {};
  const line = m[1] === 'dark' ? 2 : 1;
  const stripes = (angle: string) => `repeating-linear-gradient(${angle}, ${fg} 0 ${line}px, transparent ${line}px 4px)`;
  const image =
    m[2] === 'Grid' ? `${stripes('0deg')}, ${stripes('90deg')}` : m[2] === 'Trellis' ? `${stripes('45deg')}, ${stripes('-45deg')}` : stripes(LINE_ANGLE[m[2]]);
  return { 'background-color': bg, 'background-image': image };
}

/**
 * 칸(td)과 칸 안 글자(.g-cell-spill)에 붙일 CSS. 테두리는 편집 화면이 위·왼쪽을 따로 그리므로(grid-editor.ts applyStyleToTd)
 * 여기 넣지 않는다 — 보기 화면용 글자(cellBoxCssText)에서만 네 변을 함께 낸다.
 */
export function cellBoxCss(style: CellBoxStyle | undefined): { td: Record<string, string>; inner: Record<string, string> } {
  const td: Record<string, string> = {};
  const inner: Record<string, string> = {};
  if (!style) return { td, inner };
  if (style.pat) Object.assign(td, patternCss(style.pat));
  if (style.ind) td[style.a === 'r' ? 'padding-right' : 'padding-left'] = `${3 + style.ind * 9}px`;
  if (style.rot === 255) {
    // 세로쓰기 — 글자를 위에서 아래로 한 자씩 쌓는다
    inner['writing-mode'] = 'vertical-rl';
    inner['text-orientation'] = 'upright';
    inner['letter-spacing'] = '-1px';
  } else if (style.rot) {
    const deg = style.rot <= 90 ? -style.rot : style.rot - 90;
    inner.display = 'inline-block';
    inner.transform = `rotate(${deg}deg)`;
    inner['transform-origin'] = 'center';
  }
  return { td, inner };
}

const toText = (props: Record<string, string>) => Object.entries(props).map(([k, v]) => `${k}:${v}`).join(';');

/** 보기 화면용 — 칸 CSS(네 변 테두리 포함)와 안쪽 글자 CSS를 글자로. 둘 다 비면 undefined */
export function cellBoxCssText(style: CellBoxStyle): { x?: string; xi?: string } {
  const { td, inner } = cellBoxCss(style);
  if (style.bd) {
    for (const [side, prop] of [['t', 'border-top'], ['r', 'border-right'], ['b', 'border-bottom'], ['l', 'border-left']] as const) {
      const css = borderSideCss(style, side);
      if (css) td[prop] = css;
    }
  }
  const x = toText(td);
  const xi = toText(inner);
  return { ...(x ? { x } : {}), ...(xi ? { xi } : {}) };
}
