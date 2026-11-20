import type { Node as PMNode } from 'prosemirror-model';
import type { SlideLevelLook, SpacingValue } from '../collab/pptxShared.js';

/**
 * 문단 하나의 "보이는 모양" 중 문단에 직접 적힌 값(<a:pPr>) — 줄·단락 간격, 글머리 재정의(없음·문자·자동 번호).
 * 편집 화면 표시 전용 순수 함수 모음(서버·브라우저 공용, DOM 없음). 원본 <a:pPr>는 문단 속성 pPr에 문자열로
 * 보관돼 저장 때 그대로 되돌아가므로(textSchema.ts) 여기서는 읽기만 한다.
 */

export type BulletOverride =
  | { readonly kind: 'none' }
  | { readonly kind: 'char'; readonly char: string }
  | { readonly kind: 'auto'; readonly scheme: string; readonly startAt: number };

export interface PPrLook {
  readonly line?: SpacingValue;
  readonly before?: SpacingValue;
  readonly after?: SpacingValue;
  readonly bullet?: BulletOverride;
  /** 문단에 직접 적힌 왼쪽 여백·첫 줄 들여쓰기(pt) — 수준 상속값보다 우선 */
  readonly marLPt?: number;
  readonly indentPt?: number;
  /** eaLnBrk — true: 한글 단어를 끊지 않음(eaLnBrk="0"), false: 글자 단위(eaLnBrk="1"), 없으면 상속 */
  readonly keepWords?: boolean;
  /** 글머리 꾸밈 재정의 — 크기 비율(buSzPct), 글꼴(buFont). 색(buClr)은 테마 색까지 풀어야 해서 서버가 읽을 때 문단
   * 속성 buColor로 넣어 둔다(slideContent.ts). */
  readonly bulletSizePct?: number;
  readonly bulletFont?: string;
}

const decodeXml = (v: string) =>
  v.replace(/&(lt|gt|quot|apos|amp|#(\d+)|#x([0-9a-fA-F]+));/g, (_m, name: string, dec?: string, hex?: string) =>
    dec ? String.fromCodePoint(Number(dec)) : hex ? String.fromCodePoint(parseInt(hex, 16)) : ({ lt: '<', gt: '>', quot: '"', apos: "'", amp: '&' } as Record<string, string>)[name],
  );

function spacingIn(raw: string, tag: string): SpacingValue | undefined {
  const m = new RegExp(`<a:${tag}>\\s*<a:spc(Pct|Pts)\\s+val="(-?\\d+)"`).exec(raw);
  if (!m) return undefined;
  return m[1] === 'Pct' ? { pct: Number(m[2]) / 100000 } : { pts: Number(m[2]) / 100 };
}

/** 보관된 원본 <a:pPr …>…</a:pPr> 문자열에서 표시용 값만 읽는다(없으면 빈 객체). */
export function pPrLook(raw: string | null | undefined): PPrLook {
  if (!raw) return {};
  const line = spacingIn(raw, 'lnSpc');
  const before = spacingIn(raw, 'spcBef');
  const after = spacingIn(raw, 'spcAft');
  const auto = /<a:buAutoNum\b([^>]*)\/?>/.exec(raw);
  const char = /<a:buChar\b[^>]*\bchar="([^"]*)"/.exec(raw);
  const bullet: BulletOverride | undefined = /<a:buNone\b/.test(raw)
    ? { kind: 'none' }
    : auto
      ? {
          kind: 'auto',
          scheme: /\btype="([^"]+)"/.exec(auto[1])?.[1] ?? 'arabicPeriod',
          startAt: Number(/\bstartAt="(\d+)"/.exec(auto[1])?.[1] ?? 1),
        }
      : char
        ? { kind: 'char', char: decodeXml(char[1]) }
        : /<a:buBlip\b/.test(raw)
          ? { kind: 'char', char: '•' } // 그림 글머리는 •로(그림 파일은 따라가지 않는다)
          : undefined;
  const eaLnBrk = /^<a:pPr\b[^>]*\beaLnBrk="([^"]*)"/.exec(raw)?.[1];
  const marL = /^<a:pPr\b[^>]*\bmarL="(-?\d+)"/.exec(raw)?.[1];
  const indent = /^<a:pPr\b[^>]*\bindent="(-?\d+)"/.exec(raw)?.[1];
  const sizePct = /<a:buSzPct\s+val="(\d+)"/.exec(raw)?.[1];
  const font = /<a:buFont\b[^>]*\btypeface="([^"]+)"/.exec(raw)?.[1];
  return {
    ...(line ? { line } : {}),
    ...(before ? { before } : {}),
    ...(after ? { after } : {}),
    ...(bullet ? { bullet } : {}),
    ...(marL !== undefined ? { marLPt: Number(marL) / 12700 } : {}),
    ...(indent !== undefined ? { indentPt: Number(indent) / 12700 } : {}),
    ...(eaLnBrk !== undefined ? { keepWords: eaLnBrk === '0' || eaLnBrk === 'false' } : {}),
    ...(sizePct ? { bulletSizePct: Number(sizePct) / 100000 } : {}),
    ...(font && !font.startsWith('+') ? { bulletFont: decodeXml(font) } : {}),
  };
}

/** PowerPoint "한 줄" 간격 ≈ 글자 크기의 1.2배(맑은 고딕·Calibri 등 기본 글꼴의 줄 높이). */
export const SINGLE_LINE = 1.2;

/**
 * 줄 간격 CSS(line-height 값). 배수(pct)는 단위 없는 값이라 문단 안 글자마다 그 글자 크기에 곱해진다.
 * 자동 맞춤의 줄 간격 줄임(--ln-red, 0~1)은 PowerPoint처럼 배수에서 뺀다(130% − 20% = 110%).
 */
export function lineHeightCss(v: SpacingValue): string {
  return 'pct' in v ? `calc(${SINGLE_LINE} * (${v.pct} - var(--ln-red, 0)))` : `calc(${v.pts} * var(--slide-pt, 1pt))`;
}

/** 단락 앞·뒤 간격 CSS(margin 값) — 비율(pct)은 한 줄 높이에 대한 비율, pts는 슬라이드 pt. 자동 맞춤의 줄임 비율은
 * 비율 간격에서도 뺀다(0 아래로는 안 감) — MS Office 정답 p03: 20% 줄인 상자에서 단락 사이가 거의 붙는다. */
export function gapCss(v: SpacingValue): string {
  return 'pct' in v ? `max(0px, calc((${v.pct} - var(--ln-red, 0)) * ${SINGLE_LINE}em))` : `calc(${v.pts} * var(--slide-pt, 1pt))`;
}

const ROMAN: ReadonlyArray<[number, string]> = [
  [1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i'],
];
const roman = (n: number) => ROMAN.reduce((out, [v, s]) => { let acc = out; while (n >= v) { acc += s; n -= v; } return acc; }, '');
const alpha = (n: number) => {
  let s = '';
  for (let k = n; k > 0; k = Math.floor((k - 1) / 26)) s = String.fromCharCode(97 + ((k - 1) % 26)) + s;
  return s;
};
const CJK_DIGITS = '〇一二三四五六七八九';
/** 한자 수(1~9999): 十一·二十·百五·千 — 10 이상을 자릿수 나열(一一)로 쓰던 것을 바로잡음 */
const cjk = (n: number): string => {
  if (n <= 0 || n >= 10000) return String(n);
  let out = '';
  for (const [unit, name] of [[1000, '千'], [100, '百'], [10, '十']] as const) {
    const d = Math.floor(n / unit) % 10;
    if (d) out += (d === 1 ? '' : CJK_DIGITS[d]) + name;
  }
  return out + (n % 10 ? CJK_DIGITS[n % 10] : '');
};
const circled = (n: number, base: number) => (n >= 1 && n <= 20 ? String.fromCodePoint(base + n - 1) : String(n));
const fullWidth = (n: number) => String(n).replace(/\d/g, (d) => String.fromCodePoint(0xff10 + Number(d)));
const digits = (n: number, zero: number) => String(n).replace(/\d/g, (d) => String.fromCodePoint(zero + Number(d)));
/** 문자 목록으로 세기(목록을 다 쓰면 같은 글자를 겹쳐 쓴다 — alpha 방식) */
const letters = (n: number, list: readonly string[]) => list[(n - 1) % list.length].repeat(Math.floor((n - 1) / list.length) + 1);
const THAI_ALPHA = [...'กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ'];
const HINDI_VOWELS = [...'अआइईउऊऋएऐओऔ'];
const HINDI_CONSONANTS = [...'कखगघङचछजझञटठडढणतथदधनपफबभमयरलवशषसह'];
const HEBREW = [...'אבגדהוזחטיכלמנסעפצקרשת'];

/** OOXML 자동 번호 체계(ST_TextAutonumberScheme, ECMA-376 20.1.10.61의 41가지 전부) n번째 글머리 문자열. 목록 밖 값은 "1." 모양.
 * 한국어 가나다 체계는 OOXML에 없다(PowerPoint 한국어판도 이 목록 안에서 고른다). */
export function autoNumberText(scheme: string, n: number): string {
  const m = /^(alphaLc|alphaUc|arabic|romanLc|romanUc)(ParenBoth|ParenR|Period|Plain|Minus)$/.exec(scheme);
  if (m) {
    const body =
      m[1] === 'alphaLc' ? alpha(n) : m[1] === 'alphaUc' ? alpha(n).toUpperCase() : m[1] === 'romanLc' ? roman(n) : m[1] === 'romanUc' ? roman(n).toUpperCase() : String(n);
    return m[2] === 'ParenBoth' ? `(${body})` : m[2] === 'ParenR' ? `${body})` : m[2] === 'Period' ? `${body}.` : m[2] === 'Minus' ? `- ${body} -` : body;
  }
  switch (scheme) {
    case 'circleNumDbPlain':
    case 'circleNumWdWhitePlain':
      return circled(n, 0x2460);
    case 'circleNumWdBlackPlain':
      return n <= 10 ? String.fromCodePoint(0x2776 + n - 1) : String(n);
    case 'arabicDbPlain':
      return fullWidth(n);
    case 'arabicDbPeriod':
      return `${fullWidth(n)}．`;
    case 'ea1JpnKorPlain':
    case 'ea1ChsPlain':
    case 'ea1ChtPlain':
      return cjk(n);
    case 'ea1JpnKorPeriod':
    case 'ea1ChsPeriod':
    case 'ea1ChtPeriod':
      return `${cjk(n)}.`;
    case 'ea1JpnChsDbPeriod':
      return `${cjk(n)}．`;
    case 'arabic1Minus':
      return `${n}-`;
    case 'arabic2Minus':
      return `-${n}-`;
    case 'hebrew2Minus':
      return `-${letters(n, HEBREW)}-`;
    case 'thaiAlphaPeriod':
      return `${letters(n, THAI_ALPHA)}.`;
    case 'thaiAlphaParenR':
      return `${letters(n, THAI_ALPHA)})`;
    case 'thaiAlphaParenBoth':
      return `(${letters(n, THAI_ALPHA)})`;
    case 'thaiNumPeriod':
      return `${digits(n, 0x0e50)}.`;
    case 'thaiNumParenR':
      return `${digits(n, 0x0e50)})`;
    case 'thaiNumParenBoth':
      return `(${digits(n, 0x0e50)})`;
    case 'hindiAlphaPeriod':
      return `${letters(n, HINDI_VOWELS)}.`;
    case 'hindiAlpha1Period':
      return `${letters(n, HINDI_CONSONANTS)}.`;
    case 'hindiNumPeriod':
      return `${digits(n, 0x0966)}.`;
    case 'hindiNumParenR':
      return `${digits(n, 0x0966)})`;
    default:
      return `${n}.`;
  }
}

/**
 * 글상자 문단들의 자동 번호 글머리(문단 순서대로, 번호가 아닌 문단은 null). PowerPoint와 같은 규칙:
 * 같은 수준의 번호 문단이 이어지면 1씩 늘고, 그 수준에서 번호 아닌 문단·다른 번호 체계가 끼면 다시 시작,
 * 더 얕은 수준 문단이 나오면 깊은 수준 번호는 다시 시작한다. 빈 문단은 번호를 쓰지도 세지도 않는다.
 */
export function autoNumberPlan(doc: PMNode, levels: ReadonlyArray<SlideLevelLook> | undefined): Array<string | null> {
  const counters: Array<{ scheme: string; n: number } | null> = Array(9).fill(null);
  const out: Array<string | null> = [];
  doc.forEach((p) => {
    const lvl = Math.max(0, Math.min(8, Number(p.attrs.lvl) || 0));
    for (let k = lvl + 1; k < 9; k++) counters[k] = null;
    if (p.content.size === 0) {
      out.push(null);
      return;
    }
    const own = pPrLook(p.attrs.pPr as string | null).bullet;
    const levelAuto = levels?.[lvl]?.autoNum;
    const auto = own ? (own.kind === 'auto' ? own : null) : levelAuto ? { scheme: levelAuto.scheme, startAt: levelAuto.startAt } : null;
    if (!auto) {
      counters[lvl] = null;
      out.push(null);
      return;
    }
    const prev = counters[lvl];
    const n = prev && prev.scheme === auto.scheme ? prev.n + 1 : auto.startAt;
    counters[lvl] = { scheme: auto.scheme, n };
    out.push(autoNumberText(auto.scheme, n));
  });
  return out;
}
