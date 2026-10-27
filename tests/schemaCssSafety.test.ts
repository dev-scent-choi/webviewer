import { describe, expect, test } from 'vitest';
import type { DOMOutputSpec, Mark } from 'prosemirror-model';
import { docxSchema, safeCssColor } from '../src/docx/schema.js';
import { pptxTextSchema } from '../src/pptx/textSchema.js';

const EVIL = 'red;position:fixed;inset:0;background:url(https://evil.example/x)';
const styleOf = (spec: DOMOutputSpec) => JSON.stringify(spec);
const render = (mark: Mark) => mark.type.spec.toDOM!(mark, true);

describe('문서·참여자가 보낸 값이 style 속성에 CSS를 덧붙이지 못한다', () => {
  test('색: 정상 값은 그대로, CSS 덧붙인 값은 버린다', () => {
    for (const ok of ['#1971c2', '#abc', 'yellow', 'rgb(1, 2, 3)']) expect(safeCssColor(ok)).toBe(true);
    for (const bad of [EVIL, 'red;x:y', 'url(x)', '', null]) expect(safeCssColor(bad)).toBe(false);
    expect(styleOf(render(docxSchema.marks.textColor.create({ color: EVIL })))).not.toContain('evil');
    expect(styleOf(render(docxSchema.marks.highlight.create({ color: EVIL })))).not.toContain('evil');
    expect(styleOf(render(pptxTextSchema.marks.textColor.create({ color: EVIL })))).not.toContain('evil');
    expect(styleOf(render(docxSchema.marks.textColor.create({ color: '#ff0000' })))).toContain('color: #ff0000');
  });

  test('pptx 글꼴 이름의 따옴표·세미콜론, 글자 크기 자리의 문자열은 빠진다', () => {
    const font = styleOf(render(pptxTextSchema.marks.fontFamily.create({ family: "x';position:fixed;'" })));
    expect(font).not.toContain("x';");
    expect(font).not.toMatch(/position:fixed;/);
    expect(styleOf(render(pptxTextSchema.marks.fontSize.create({ size: '1);position:fixed;(' })))).not.toContain('position');
  });
});
