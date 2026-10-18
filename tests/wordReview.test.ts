import { EditorState, TextSelection } from 'prosemirror-state';
import { describe, expect, test } from 'vitest';
import { docxSchema } from '../src/docx/schema.js';
import { addCommentMark, commentAt, commentRanges, commentTarget, removeCommentMark, wordStats } from '../src/client/wordReview.js';

const { nodes } = docxSchema;
const docOf = (...texts: string[]) => nodes.doc.create(null, texts.map((t) => nodes.paragraph.create(null, t ? docxSchema.text(t) : [])));

describe('Word 검토 탭', () => {
  test('메모 달기·찾기·지우기: 겹친 다른 메모는 남는다', () => {
    let state = EditorState.create({ schema: docxSchema, doc: docOf('hello world', 'second line') });
    const apply = (cmd: (s: EditorState, d: (tr: import('prosemirror-state').Transaction) => void) => boolean) => {
      cmd(state, (tr) => {
        state = state.apply(tr);
      });
    };
    apply(addCommentMark('a', 1, 6));
    apply(addCommentMark('b', 3, 18));
    expect(commentRanges(state.doc)).toEqual([
      { id: 'a', from: 1, to: 6 },
      { id: 'b', from: 3, to: 18 },
    ]);
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2)));
    expect(commentAt(state)).toBe('a');
    apply(removeCommentMark('a'));
    expect(commentRanges(state.doc).map((r) => r.id)).toEqual(['b']);
  });

  test('선택이 없으면 커서가 있는 낱말이 메모 범위', () => {
    let state = EditorState.create({ schema: docxSchema, doc: docOf('hello world') });
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 9)));
    expect(commentTarget(state)).toEqual({ from: 7, to: 12 });
  });

  test('단어 개수: 단어·문자(공백 제외/포함)·단락, 범위 지정', () => {
    const doc = docOf('한글 단어 셋', '', 'two words');
    expect(wordStats(doc)).toEqual({ words: 5, charsNoSpaces: 13, charsWithSpaces: 16, paragraphs: 2 });
    expect(wordStats(doc, 1, 3).words).toBe(1);
  });
});

describe('글자가 없는 선택(페이지 나누기만 선택 등)', () => {
  test('메모 범위는 그 문단의 글자 전체로 넓힌다', async () => {
    const { NodeSelection } = await import('prosemirror-state');
    const doc = nodes.doc.create(null, [nodes.paragraph.create(null, [docxSchema.text('앞'), nodes.page_break.create(), docxSchema.text('뒤')])]);
    let state = EditorState.create({ schema: docxSchema, doc });
    state = state.apply(state.tr.setSelection(NodeSelection.create(state.doc, 2)));
    expect(commentTarget(state)).toEqual({ from: 1, to: 4 });
  });
});
