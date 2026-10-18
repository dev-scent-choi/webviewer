import { Schema, type MarkSpec, type NodeSpec } from 'prosemirror-model';

/**
 * pptx 텍스트박스 전용 최소 ProseMirror 스키마. docx의 docxSchema(src/docx/schema.ts)와
 * 마크 정의(강조 종류·속성)는 동일하게 맞춘다 — import/export 시 같은 pt/색 표현을 그대로
 * 재사용하기 위해서다. 다만 텍스트박스는 표·목록·제목이 필요 없어 노드는 doc/paragraph/text만
 * 남긴다(docxSchema 그대로 재사용하면 안 쓸 커맨드·키맵까지 끌고 들어온다).
 */
const nodes: Record<string, NodeSpec> = {
  doc: { content: 'paragraph+' },
  paragraph: {
    /** align: pptx <a:pPr algn="..">의 값 그대로('l'|'ctr'|'r'|'just'), 없으면 null —
     * 원본에 algn이 아예 없던 문단을 저장할 때도 없는 그대로 내보내기 위해 기본값을 'l'로
     * 정규화하지 않는다(import/export 왕복 보존 원칙). */
    attrs: { align: { default: null } },
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
      return css ? ['p', { style: `text-align: ${css}` }, 0] : ['p', 0];
    },
  },
  text: { group: 'inline' },
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
    attrs: { color: {} },
    parseDOM: [{ style: 'color', getAttrs: (value) => ({ color: value as string }) }],
    toDOM: (mark) => ['span', { style: `color: ${mark.attrs.color as string}` }, 0],
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
    toDOM: (mark) => ['span', { style: `font-size: calc(${mark.attrs.size as number} * var(--slide-pt, 1pt))` }, 0],
  },
  fontFamily: {
    /** 글꼴 이름 — pptx의 <a:rPr><a:latin typeface=".."/>(한글 글리프용 <a:ea>도 같이 내보낸다,
     * slideContent.ts 참고)와 import/export 시 서로 변환한다. */
    attrs: { family: {} },
    parseDOM: [{ style: 'font-family', getAttrs: (value) => ({ family: String(value).replace(/['"]/g, '') }) }],
    toDOM: (mark) => ['span', { style: `font-family: '${mark.attrs.family as string}'` }, 0],
  },
};

export const pptxTextSchema = new Schema({ nodes, marks });
