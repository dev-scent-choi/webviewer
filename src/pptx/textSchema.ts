import { Schema, type MarkSpec, type NodeSpec } from 'prosemirror-model';
import { safeCssColor } from '../docx/schema.js';
import { gapCss, lineHeightCss, pPrLook } from './paragraphLook.js';

/**
 * pptx 텍스트박스 전용 최소 ProseMirror 스키마. docx의 docxSchema(src/docx/schema.ts)와
 * 마크 정의(강조 종류·속성)는 동일하게 맞춘다 — import/export 시 같은 pt/색 표현을 그대로
 * 재사용하기 위해서다. 다만 텍스트박스는 표·목록·제목이 필요 없어 노드는 doc/paragraph/text만
 * 남긴다(docxSchema 그대로 재사용하면 안 쓸 커맨드·키맵까지 끌고 들어온다).
 */
/**
 * 글자 크기 pt → CSS 길이. 자동 맞춤 배율(--fs-scale, 글상자가 지정)을 곱하고, 배율이 걸렸으면 PowerPoint처럼 0.5pt 단위로
 * 내린다(--fs-step: 배율이 없으면 0.01 = 사실상 그대로). 근거: MS Office 정답 p03 1장 — 32pt × 47.5% = 15.2pt인데 글자 상자
 * 높이가 15.0pt이고 줄도 15pt 폭으로 나뉜다(15.2pt로 그리면 첫 단락이 한 줄 더 생겨 마지막 단락이 슬라이드 밖으로 밀렸다).
 */
export const fontSizeCss = (pt: number): string =>
  `calc(round(down, ${pt} * var(--fs-scale, 1), var(--fs-step, 0.01)) * var(--slide-pt, 1pt))`;

/** 슬라이드 그림(zip 안 경로)의 화면 주소 — 편집기가 문서 열람 토큰을 넣어 정한다(서버에선 DOM을 그리지 않아 쓰이지 않음).
 * 경로는 협업자가 쓸 수 있는 값이라 반드시 encodeURIComponent로 넣는다(CSS url("…") 안에 들어간다). */
let mediaUrlBuilder: (path: string) => string = () => '';
export function setSlideMediaUrl(build: (path: string) => string): void {
  mediaUrlBuilder = build;
}
export const slideMediaUrl = (path: string): string => mediaUrlBuilder(path);

const nodes: Record<string, NodeSpec> = {
  doc: { content: 'paragraph+' },
  paragraph: {
    /** align: pptx <a:pPr algn="..">의 값 그대로('l'|'ctr'|'r'|'just'), 없으면 null —
     * 원본에 algn이 아예 없던 문단을 저장할 때도 없는 그대로 내보내기 위해 기본값을 'l'로
     * 정규화하지 않는다(import/export 왕복 보존 원칙). */
    /**
     * pPr: 원본 <a:pPr …>…</a:pPr>를 algn 속성만 뺀 채 문자열 그대로 보관 — 저장 때 그대로 되돌린다
     * (수준 lvl, 글머리 없음/문자, 들여쓰기, 줄·문단 간격 등 이 편집기가 모델링하지 않는 문단 속성이
     * 예전엔 글자를 한 자만 고쳐도 저장에서 전부 사라졌다). lvl·bu는 화면 표시용으로 pPr에서 뽑아 둔
     * 값(목록 수준 0~8, 글머리 재정의: 문자 / 'none' / null=상속).
     */
    /** endRPr: 원본 <a:endParaRPr …/>(문단 끝 글자 속성 — 빈 문단의 높이를 정한다) 문자열 그대로, 저장 때 되돌린다. */
    /** buColor·buImage: 화면 표시 전용 — 문단에 직접 적힌 글머리 색(테마 색을 푼 '#RRGGBB')과 그림 글머리 그림의 zip 경로
     * (서버가 읽을 때 풀어 둔다. 저장은 pPr 원문이 맡는다). */
    attrs: {
      align: { default: null },
      pPr: { default: null },
      lvl: { default: 0 },
      bu: { default: null },
      endRPr: { default: null },
      buColor: { default: null },
      buImage: { default: null },
    },
    content: 'inline*',
    group: 'block',
    parseDOM: [
      {
        tag: 'p',
        getAttrs: (node) => {
          const textAlign = (node as HTMLElement).style?.textAlign;
          const align =
            textAlign === 'center' ? 'ctr' : textAlign === 'right' ? 'r' : textAlign === 'justify' ? 'just' : null;
          return { align };
        },
      },
    ],
    toDOM: (node) => {
      const align = node.attrs.align as string | null;
      const css =
        align === 'ctr' ? 'center' : align === 'r' ? 'right' : align === 'just' ? 'justify' : null;
      const attrs: Record<string, string> = { 'data-lvl': String(node.attrs.lvl ?? 0) };
      const styles: string[] = [];
      if (css) styles.push(`text-align: ${css}`);
      // 줄 높이는 문단 글자 크기를 바탕으로 잡힌다 — 모든 글자에 크기가 정해진 문단은 그중 가장 큰 크기를
      // 문단 크기로 둬서, 상자 기본 크기(18pt)가 11pt 글자의 줄 간격을 벌려 다음 줄이 아래 개체에 깔리거나
      // 작은 번호 원의 숫자가 아래로 밀리던 문제를 막는다(PowerPoint도 줄 높이를 그 줄 글자 크기로 정한다).
      let maxSize = 0;
      let allSized = node.childCount > 0;
      // 빈 문단(빈 줄)의 높이는 문단 끝 글자 크기(endParaRPr sz)로 — 큰 글자 사이의 빈 줄이 기본 크기로 줄어 보였다
      const endSz = node.childCount === 0 ? Number(/\bsz="(\d+)"/.exec(String(node.attrs.endRPr ?? ''))?.[1]) / 100 : NaN;
      if (endSz > 0) styles.push(`font-size: ${fontSizeCss(endSz)}`);
      node.forEach((child) => {
        if (!child.isText) return;
        const size = child.marks.find((mark) => mark.type.name === 'fontSize')?.attrs.size as number | undefined;
        if (typeof size === 'number' && size > 0) maxSize = Math.max(maxSize, size);
        else allSized = false;
      });
      if (allSized && maxSize > 0) styles.push(`font-size: ${fontSizeCss(maxSize)}`);
      // 모든 글자가 같은 글꼴이면 문단 자체도 그 글꼴로 — 문단(줄 기준선)과 글자의 글꼴이 다르면 두 글꼴의 위아래 여유가
      // 어긋나 줄마다 1~2pt씩 벌어졌다(정답 p08 실측: 8줄에 10pt 밀림). PowerPoint 줄 높이 = 줄 간격 × 그 줄 글자 크기.
      const families = new Set<string>();
      node.forEach((child) => {
        if (child.isText) families.add(String(child.marks.find((mark) => mark.type.name === 'fontFamily')?.attrs.family ?? ''));
      });
      const [family] = [...families];
      if (families.size === 1 && family) styles.push(`font-family: '${family.replace(/["';<>\\]/g, '')}'`);
      // 문단에 직접 적힌 줄·단락 간격과 글머리 꾸밈(<a:pPr>) — 상속값(수준별 --lN-*)보다 우선(CSS 변수로 넘긴다)
      const look = pPrLook(node.attrs.pPr as string | null);
      if (look.marLPt !== undefined) styles.push(`--marl: ${look.marLPt}`);
      if (look.indentPt !== undefined) styles.push(`--ind: ${look.indentPt}`);
      if (look.keepWords !== undefined) styles.push(`word-break: ${look.keepWords ? 'keep-all' : 'normal'}`);
      if (look.line) styles.push(`--p-lh: ${lineHeightCss(look.line)}`);
      if (look.before) styles.push(`--p-sb: ${gapCss(look.before)}`);
      if (look.after) styles.push(`--p-sa: ${gapCss(look.after)}`);
      const buColor = node.attrs.buColor as string | null;
      if (buColor && /^#[0-9A-Fa-f]{6}$/.test(buColor)) styles.push(`--p-bu-color: ${buColor}`);
      const buImage = node.attrs.buImage as string | null;
      if (buImage) styles.push(`--p-bu-img: url("${slideMediaUrl(String(buImage))}")`, '--p-bu-h: 1em');
      if (look.bulletSizePct) styles.push(`--p-bu-size: ${look.bulletSizePct}em`);
      if (look.bulletFont) styles.push(`--p-bu-ff: '${look.bulletFont.replace(/["';<>\\]/g, '')}'`);
      if (styles.length > 0) attrs.style = styles.join('; ');
      if (buImage) attrs['data-bu'] = ''; // 그림 글머리: 글자 대신 그림(--p-bu-img)
      else if (node.attrs.bu !== null && node.attrs.bu !== undefined) attrs['data-bu'] = String(node.attrs.bu);
      return ['p', attrs, 0];
    },
  },
  text: { group: 'inline' },
  /** 문단 안 줄 바꿈(<a:br>, PowerPoint Shift+Enter) — rPr: 원본 <a:rPr>(있으면) 문자열, 저장 때 되돌린다. */
  br: {
    inline: true,
    group: 'inline',
    selectable: false,
    attrs: { rPr: { default: null } },
    parseDOM: [{ tag: 'br' }],
    toDOM: () => ['br'],
  },
  /** 보존 조각 — 필드(<a:fld> 쪽 번호·날짜)와 이 편집기가 모르는 문단 자식(수식 mc:AlternateContent 등): 원본 XML(raw)을
   * 통째로 보관해 저장 때 그대로 되돌리고(서버가 원본에서 읽은 조각일 때만 — slideContent.ts TrustedFragments), 화면엔
   * 저장된 값·대체 글자(text)를 보인다. 한 덩어리로 지우거나 옮길 수는 있어도 안을 고칠 수는 없다. */
  field: {
    inline: true,
    group: 'inline',
    atom: true,
    attrs: { raw: {}, text: { default: '' } },
    toDOM: (node) => ['span', { class: 'slide-field', contenteditable: 'false' }, String(node.attrs.text ?? '')],
  },
};

const marks: Record<string, MarkSpec> = {
  strong: {
    parseDOM: [{ tag: 'strong' }, { tag: 'b' }],
    toDOM: () => ['strong', 0],
  },
  em: {
    parseDOM: [{ tag: 'em' }, { tag: 'i' }],
    toDOM: () => ['em', 0],
  },
  underline: {
    parseDOM: [{ tag: 'u' }],
    toDOM: () => ['u', 0],
  },
  textColor: {
    /** raw: 원본 <a:solidFill>의 내용 XML(테마 색 <a:schemeClr …><a:lumMod/></a:schemeClr> 등) —
     * 있으면 저장 때 그대로 되돌려 테마 연결을 유지한다(색을 고정 RGB로 바꿔 쓰지 않는다). color는
     * 화면 표시용으로 테마를 풀어 낸 '#RRGGBB'(풀 수 없으면 ''). 리본으로 색을 새로 고르면 raw 없이
     * 만들어져 일반 RGB 색으로 저장된다. */
    attrs: { color: {}, raw: { default: null } },
    parseDOM: [{ style: 'color', getAttrs: (value) => ({ color: value as string }) }],
    toDOM: (mark) => (safeCssColor(mark.attrs.color) ? ['span', { style: `color: ${mark.attrs.color as string}` }, 0] : ['span', 0]),
  },
  fontSize: {
    /** 포인트(pt) 단위 — pptx의 <a:rPr sz="..">(100분의 1 포인트)와는 import/export 시 변환한다. */
    attrs: { size: {} },
    parseDOM: [
      {
        style: 'font-size',
        // 우리 toDOM 형태("calc(40 * var(--slide-pt, 1pt))")와 일반 "40pt" 둘 다 첫 숫자가 pt 값이다.
        getAttrs: (value) => {
          const size = parseFloat(/[\d.]+/.exec(String(value))?.[0] ?? '');
          return Number.isFinite(size) ? { size } : false;
        },
      },
    ],
    // 캔버스는 실제 슬라이드를 축소/확대해 그리므로 글자도 같은 비율이어야 한다 — CSS의 pt(=1.333px)를
    // 그대로 쓰면 캔버스 축척과 어긋나 글자만 33% 크게 보여 PowerPoint에선 한 줄인 제목이 줄바꿈됐다.
    // --slide-pt(캔버스 1pt의 실제 px)는 편집기가 슬라이드 크기에 맞춰 .slide-canvas에 지정한다.
    toDOM: (mark) => ['span', { style: `font-size: ${fontSizeCss(Number(mark.attrs.size) || 18)}` }, 0],
  },
  /** 원본 <a:rPr> 전체(문자열) — 이 편집기가 모델링하지 않는 글자 속성(lang, 자간 spc, 하이퍼링크 hlinkClick, 강조색,
   * 테마 글꼴 참조 등)을 저장 때 그대로 되돌리기 위해 보관한다(slideContent.ts rPrXmlFor가 위 마크들과 합친다).
   * 화면엔 하이퍼링크만 보인다(테마 하이퍼링크 색 + 밑줄). */
  runProps: {
    attrs: { raw: {} },
    // 하이퍼링크 — PowerPoint 2019+ "링크도 글자 색 그대로"(<ahyp:hlinkClr val="tx"/>)면 테마 하이퍼링크 색 대신 글자 색
    // 자간(spc, 1/100pt)·b="0" — 안 그리면 spc="-80" 같은 한글 제목이 원본보다 넓어 PowerPoint·아트크래프트와 다른 곳에서 줄이 끊겼다
    toDOM: (mark) => {
      const raw = String(mark.attrs.raw);
      const spc = Number(/^<a:rPr\b[^>]*?\sspc="(-?\d+)"/.exec(raw)?.[1] ?? 0);
      const attrs: Record<string, string> = {};
      const classes: string[] = [];
      if (/<a:hlinkClick\b/.test(raw)) classes.push(/hlinkClr\b[^>]*\bval="tx"/.test(raw) ? 'slide-hlink slide-hlink-tx' : 'slide-hlink');
      // 글자에 직접 b="0" — 수준(lstStyle defRPr b="1")의 굵게를 끈다(규칙은 slideTextLook ensureLevelRules, 굵게 마크 안이면 굵게)
      if (/^<a:rPr\b[^>]*?\sb="(0|false)"/.test(raw)) classes.push('slide-b0');
      if (classes.length) attrs.class = classes.join(' ');
      if (spc) attrs.style = `letter-spacing: calc(${spc / 100} * var(--slide-pt, 1pt))`;
      return Object.keys(attrs).length ? ['span', attrs, 0] : ['span', 0];
    },
  },
  fontFamily: {
    /** 글꼴 이름 — pptx의 <a:rPr><a:latin typeface=".."/>(한글 글리프용 <a:ea>도 같이 내보낸다,
     * slideContent.ts 참고)와 import/export 시 서로 변환한다. */
    attrs: { family: {} },
    parseDOM: [{ style: 'font-family', getAttrs: (value) => ({ family: String(value).replace(/['"]/g, '') }) }],
    toDOM: (mark) => ['span', { style: `font-family: '${String(mark.attrs.family ?? '').replace(/["';<>\\]/g, '')}'` }, 0],
  },
};

export const pptxTextSchema = new Schema({ nodes, marks });

