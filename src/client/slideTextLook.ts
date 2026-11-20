import { Plugin, PluginKey } from 'prosemirror-state';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { Node as PMNode } from 'prosemirror-model';
import { DEFAULT_INSETS_EMU, type SlideLevelLook, type SlideShapeMeta } from '../collab/pptxShared.js';
import { autoNumberPlan, gapCss, lineHeightCss } from '../pptx/paragraphLook.js';
import { fontSizeCss, slideMediaUrl } from '../pptx/textSchema.js';

/**
 * 편집 캔버스 글상자의 "물려받은 모양"(src/pptx/inheritance.ts가 푼 meta)을 화면에 입히는 부분 —
 * 수준별 글머리·간격·글꼴·굵기, 자동 맞춤 글자 배율, 세로 쓰기, 회전, 프리셋 모양의 글자 칸.
 * 저장과는 무관하다(표시 전용).
 */

const BASE_FONT = "'맑은 고딕', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif";
const cssFont = (name: string) => `'${name.replace(/["';<>\\]/g, '')}'`;

/** 수준(data-lvl)마다 --lN-* 변수를 문단 공용 변수(--marl, --lv-lh …)로 옮기는 규칙 — 한 번만 문서에 넣는다. */
export function ensureLevelRules(): void {
  if (document.getElementById('slide-level-rules')) return;
  const rules = Array.from({ length: 9 }, (_, i) => {
    const v = (name: string, fallback: string) => `var(--l${i}-${name}, ${fallback})`;
    return (
      `.slide-textbox .ProseMirror p[data-lvl="${i}"] {` +
      ` --marl: ${v('marl', '0')}; --ind: ${v('ind', '0')}; --bu: ${v('bu', 'none')};` +
      // 1수준 크기는 글상자 기본 크기(view.dom font-size)가 맡는다
      (i > 0 ? ` font-size: ${v('sz', 'inherit')};` : '') +
      ` color: ${v('color', 'inherit')}; font-weight: ${v('fw', 'inherit')}; font-style: ${v('fi', 'inherit')};` +
      ` text-transform: ${v('tt', 'none')}; font-family: ${v('ff', 'inherit')}; word-break: ${v('wb', 'normal')};` +
      ` --lv-lh: ${v('lh', 'calc(1.2 * (1 - var(--ln-red, 0)))')}; --lv-sb: ${v('sb', '0px')}; --lv-sa: ${v('sa', '0px')};` +
      ` --bu-ff: var(--p-bu-ff, ${v('buff', 'inherit')}); --bu-color: var(--p-bu-color, ${v('buclr', 'inherit')}); --bu-size: var(--p-bu-size, ${v('busz', 'inherit')});` +
      ` --bu-img: var(--p-bu-img, ${v('buimg', 'none')}); --bu-h: var(--p-bu-h, ${v('buh', 'auto')}); }`
    );
  });
  const style = document.createElement('style');
  style.id = 'slide-level-rules';
  style.textContent = rules.join('\n');
  document.head.append(style);
}

/** 글꼴 목록(라틴 → 동아시아 → 편집기 기본): 영문은 라틴 글꼴, 한글은 그 글꼴에 없으면 동아시아 글꼴로 그려진다 */
function fontStack(latin?: string, ea?: string): string | undefined {
  if (!latin && !ea) return undefined;
  return [latin, ea].filter((f): f is string => Boolean(f)).map(cssFont).concat(BASE_FONT).join(', ');
}

/** 수준별 값을 --lN-* 변수로(ensureLevelRules가 문단에 옮긴다). 글머리 글꼴(Wingdings 등)이 브라우저에 없으면 기본 글꼴로 대체된다. */
function setLevelVars(dom: HTMLElement, levels: ReadonlyArray<SlideLevelLook>): void {
  levels.forEach((level, i) => {
    const set = (name: string, value: string | undefined) => {
      if (value !== undefined) dom.style.setProperty(`--l${i}-${name}`, value);
    };
    set('marl', String(level.marLPt));
    set('ind', String(level.indentPt));
    set('bu', level.bulletImage ? '""' : level.bullet === null ? 'none' : JSON.stringify(level.bullet));
    if (level.bulletImage) {
      set('buimg', `url("${slideMediaUrl(level.bulletImage)}")`);
      set('buh', '1em');
    }
    set('sz', fontSizeCss(level.sizePt));
    if (level.color && /^#[0-9A-Fa-f]{6}$/.test(level.color)) set('color', level.color);
    if (level.bold) set('fw', '700');
    if (level.italic) set('fi', 'italic');
    if (level.caps) set('tt', 'uppercase');
    if (level.keepWords) set('wb', 'keep-all');
    set('ff', fontStack(level.fontLatin, level.fontEa));
    if (level.spacing?.line) set('lh', lineHeightCss(level.spacing.line));
    if (level.spacing?.before) set('sb', gapCss(level.spacing.before));
    if (level.spacing?.after) set('sa', gapCss(level.spacing.after));
    if (level.bulletFont) set('buff', `${cssFont(level.bulletFont)}, ${BASE_FONT}`);
    if (level.bulletColor && /^#[0-9A-Fa-f]{6}$/.test(level.bulletColor)) set('buclr', level.bulletColor);
    if (level.bulletSize) set('busz', 'pct' in level.bulletSize ? `${level.bulletSize.pct}em` : fontSizeCss(level.bulletSize.pts));
  });
}

/** 자동 번호 글머리 — 문단 순서를 알아야 셀 수 있어 문서 전체를 보는 플러그인이 data-bu로 번호를 붙인다. */
export function autoNumberPlugin(levels: ReadonlyArray<SlideLevelLook> | undefined): Plugin {
  const build = (doc: PMNode) => {
    const plan = autoNumberPlan(doc, levels);
    if (plan.every((t) => t === null)) return DecorationSet.empty;
    const decos: Decoration[] = [];
    let i = 0;
    doc.forEach((node, offset) => {
      const text = plan[i++];
      if (text !== null) decos.push(Decoration.node(offset, offset + node.nodeSize, { 'data-bu': text }));
    });
    return DecorationSet.create(doc, decos);
  };
  return new Plugin({
    key: autoNumberKey,
    state: { init: (_config, state) => build(state.doc), apply: (tr, old) => (tr.docChanged ? build(tr.doc) : old) },
    props: { decorations: (state) => autoNumberKey.getState(state) },
  });
}

const autoNumberKey = new PluginKey<DecorationSet>('slideAutoNumber');

/**
 * 프리셋 모양 안의 글자 칸(OOXML presetShapeDefinitions의 text rect)을 상자 가장자리에서 들어간 거리로(EMU,
 * [왼, 위, 오른, 아래]). 타원·별·삼각형처럼 글자 칸이 도형보다 좁은 모양에서 글자가 모양 밖으로 나가지 않게 한다.
 * 조정값(avLst)은 기본값 기준. 모르는 모양은 0(상자 전체).
 */
export function textRectInsets(geom: string | undefined, w: number, h: number): [number, number, number, number] {
  const ss = Math.min(w, h);
  switch (geom) {
    case 'ellipse':
    case 'flowChartConnector': {
      const k = 1 - Math.SQRT1_2; // 내접 사각형
      return [(w * k) / 2, (h * k) / 2, (w * k) / 2, (h * k) / 2];
    }
    case 'triangle':
      return [w / 4, h / 2, w / 4, 0];
    case 'diamond':
    case 'flowChartDecision':
      return [w / 4, h / 4, w / 4, h / 4];
    case 'star5':
      return [w * 0.31, h * 0.38, w * 0.31, h * 0.27];
    case 'homePlate':
      return [0, 0, ss / 4, 0];
    case 'chevron':
      return [ss / 2, 0, ss / 2, 0];
    case 'rightArrow':
      return [0, h / 4, ss / 4, h / 4];
    case 'leftArrow':
      return [ss / 4, h / 4, 0, h / 4];
    case 'can':
      return [0, ss / 4, 0, ss / 8];
    case 'cloud':
      return [w * 0.138, h * 0.151, w * 0.209, h * 0.197];
    case 'roundRect': {
      const r = (ss / 6) * 0.29289; // 기본 반지름(짧은 변 1/6)의 45° 지점
      return [r, r, r, r];
    }
    default:
      return [0, 0, 0, 0];
  }
}

/** 회전·뒤집기 CSS transform. 글자 상자는 뒤집어도 글자를 거울로 만들지 않는다(PowerPoint: 세로 뒤집기만 글자가 180° 돈다). */
export function turnCss(turn: { rot?: number; flipH?: boolean; flipV?: boolean }, mirror: boolean): string {
  const parts: string[] = [];
  const rot = (turn.rot ?? 0) + (!mirror && turn.flipV ? 180 : 0);
  if (rot) parts.push(`rotate(${Math.round(rot * 100) / 100}deg)`);
  if (mirror && (turn.flipH || turn.flipV)) parts.push(`scale(${turn.flipH ? -1 : 1}, ${turn.flipV ? -1 : 1})`);
  return parts.join(' ');
}

/**
 * 뒤집힌 도형(그룹 뒤집기 포함)의 선형 그라데이션 방향 — 도형 상자는 CSS로 거울처럼 뒤집지만(모양이 뒤집혀야 하므로),
 * 그라데이션 방향은 뒤집지 않는다. 근거: ECMA-376(20.1.8.33 rotWithShape)은 회전만 정하고 뒤집기는 말이 없고, 보기 화면
 * (LibreOffice)이 뒤집힌 도형의 그라데이션을 뒤집지 않고 그린다 — 보기·편집이 같게 그쪽을 따른다(MS Office 정답이 생기면 확인할 것).
 * 미리 반대로 돌려 두면 상자의 거울 변환과 상쇄된다: 가로 뒤집기 a → −a, 세로 뒤집기 a → 180°−a.
 */
export function unmirroredFill(fill: string, turn: { flipH?: boolean; flipV?: boolean }): string {
  const m = /^linear-gradient\((-?[\d.]+)deg,/.exec(fill);
  if (!m || (!turn.flipH && !turn.flipV)) return fill;
  let a = Number(m[1]);
  if (turn.flipH) a = -a;
  if (turn.flipV) a = 180 - a;
  return fill.replace(m[0], `linear-gradient(${((a % 360) + 360) % 360}deg,`);
}

const VERT: Record<string, { mode: string; orient: string; flip?: boolean }> = {
  vert: { mode: 'vertical-rl', orient: 'sideways' },
  vert270: { mode: 'vertical-rl', orient: 'sideways', flip: true },
  eaVert: { mode: 'vertical-rl', orient: 'mixed' },
  wordArtVert: { mode: 'vertical-rl', orient: 'upright' },
  wordArtVertRtl: { mode: 'vertical-rl', orient: 'upright' },
  mongolianVert: { mode: 'vertical-lr', orient: 'mixed' },
};

/**
 * 원본 글상자(시딩된 meta)의 물려받은 모양을 상자(box)와 편집 영역(dom = view.dom)에 입힌다. 글자 크기는 자동 맞춤 배율을
 * 곱하고 0.5pt 단위로 내린 크기(textSchema fontSizeCss)로, 여백·들여쓰기는 슬라이드 1pt(--slide-pt)로 잰다.
 */
export function applyTextboxLook(box: HTMLElement, dom: HTMLElement, meta: SlideShapeMeta, emuToPx: (emu: number) => number): void {
  if (meta.autofit) applyAutofit(dom, meta.autofit);
  if (meta.defaultSizePt) dom.style.fontSize = fontSizeCss(meta.defaultSizePt);
  if (meta.levels) setLevelVars(dom, meta.levels);
  const base = meta.levels?.[0];
  const family = fontStack(base?.fontLatin, base?.fontEa);
  if (family) dom.style.fontFamily = family;
  if (meta.hlinkColor) dom.style.setProperty('--hlink', meta.hlinkColor);
  const turn = turnCss(meta, false);
  if (turn) box.style.transform = turn;
  const vert = meta.vert ? VERT[meta.vert] : undefined;
  if (vert) {
    box.classList.add('slide-vert');
    dom.style.writingMode = vert.mode;
    dom.style.textOrientation = vert.orient;
    if (vert.flip) dom.style.transform = 'rotate(180deg)';
  }
  // 프리셋 모양의 글자 칸 — 안쪽 여백(원본 또는 기본)에 더한다
  if (meta.presetGeom && meta.w && meta.h && !meta.geom) {
    const extra = textRectInsets(meta.presetGeom, meta.w, meta.h);
    if (extra.some((v) => v > 0)) {
      const [l, t, r, b] = (meta.insets ?? DEFAULT_INSETS_EMU).map((v, i) => v + extra[i]);
      box.style.padding = [t, r, b, l].map((emu) => `${emuToPx(emu)}px`).join(' ');
    }
  }
}

type Autofit = { readonly fontScale: number; readonly lnSpcReduction: number };

/** 자동 맞춤(normAutofit): 글자 크기만 배율만큼(들여쓰기·여백은 그대로), 줄 간격은 줄임 비율만큼 뺀다 */
export function applyAutofit(dom: HTMLElement, autofit: Autofit): void {
  dom.style.setProperty('--fs-scale', String(autofit.fontScale));
  dom.style.setProperty('--fs-step', autofit.fontScale < 1 ? '0.5' : '0.01');
  dom.style.setProperty('--ln-red', String(autofit.lnSpcReduction));
}

/** PowerPoint "넘치면 글자 줄이기"가 고르는 단계: 먼저 줄 간격을 10%·20% 줄이고, 그래도 넘치면 글자를 7.5%씩 줄인다
 * (PowerPoint가 저장하는 값 92500·85000·77500…·25000과 같은 단계). */
const AUTOFIT_STEPS: readonly Autofit[] = [
  { fontScale: 1, lnSpcReduction: 0 },
  { fontScale: 1, lnSpcReduction: 0.1 },
  { fontScale: 1, lnSpcReduction: 0.2 },
  ...Array.from({ length: 10 }, (_, i) => ({ fontScale: Math.round((0.925 - i * 0.075) * 1000) / 1000, lnSpcReduction: 0.2 })),
];

/** 글을 고친 뒤 상자에 맞는 가장 큰 단계를 다시 고른다(글상자 높이 안에 글이 다 들어가는 첫 단계, 다 넘치면 마지막). */
export function refitAutofit(box: HTMLElement, dom: HTMLElement): Autofit {
  const cs = getComputedStyle(box);
  const room = box.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom) + 1;
  for (const step of AUTOFIT_STEPS) {
    applyAutofit(dom, step);
    if (dom.scrollHeight <= room) return step;
  }
  return AUTOFIT_STEPS[AUTOFIT_STEPS.length - 1];
}

/** 캔·구름 모양 — 다각형(clip-path polygon)으로 못 그리는 곡선 모양을 px 경로로(상자 크기를 알아야 한다). */
export function curvedShapePath(geom: string | undefined, w: number, h: number): string | undefined {
  const f = (n: number) => Math.round(n * 10) / 10;
  if (geom === 'can') {
    const ry = Math.min(w, h) / 8; // 위·아래 타원 반높이(조정값 기본 25%의 절반)
    return `M0,${f(ry)} A${f(w / 2)},${f(ry)} 0 0 1 ${f(w)},${f(ry)} L${f(w)},${f(h - ry)} A${f(w / 2)},${f(ry)} 0 0 1 0,${f(h - ry)} Z`;
  }
  if (geom === 'cloud') {
    // 원 여러 개의 합(같은 방향으로 그려 nonzero 채우기로 합쳐진다) — PowerPoint 구름 윤곽 근사
    const blobs: ReadonlyArray<[number, number, number]> = [
      [0.2, 0.6, 0.17], [0.3, 0.32, 0.19], [0.52, 0.22, 0.2], [0.74, 0.3, 0.18], [0.86, 0.52, 0.14],
      [0.78, 0.76, 0.16], [0.56, 0.8, 0.18], [0.33, 0.8, 0.16], [0.5, 0.52, 0.3],
    ];
    return blobs
      .map(([cx, cy, r]) => {
        const rx = r * w;
        const ry = r * h;
        return `M${f(cx * w - rx)},${f(cy * h)} a${f(rx)},${f(ry)} 0 1 0 ${f(2 * rx)},0 a${f(rx)},${f(ry)} 0 1 0 ${f(-2 * rx)},0 Z`;
      })
      .join(' ');
  }
  return undefined;
}
