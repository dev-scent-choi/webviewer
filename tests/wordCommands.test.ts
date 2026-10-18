import { EditorState, TextSelection, type Command } from 'prosemirror-state';
import { describe, expect, test } from 'vitest';
import { docxSchema } from '../src/docx/schema.js';
import {
  applyMark,
  changeCase,
  changeIndent,
  clearFormatting,
  copyFormat,
  currentFontName,
  currentParagraphStyle,
  pasteFormat,
  replaceAll,
  setHeading,
  setLineSpacing,
  setParagraphStyle,
  setShading,
  stepFontSize,
} from '../src/client/wordCommands.js';

const { nodes, marks } = docxSchema;

function stateOf(paragraphs: string[], select?: [number, number]): EditorState {
  const doc = nodes.doc.create(null, paragraphs.map((text) => nodes.paragraph.create(null, text ? docxSchema.text(text) : [])));
  const state = EditorState.create({ schema: docxSchema, doc });
  if (!select) return state;
  return state.apply(state.tr.setSelection(TextSelection.create(state.doc, select[0], select[1])));
}

function run(state: EditorState, command: Command): EditorState {
  let next = state;
  command(state, (tr) => {
    next = state.apply(tr);
  });
  return next;
}

describe('Word 홈 리본 명령', () => {
  test('글꼴 상자: 한글·영문 글꼴을 같은 이름으로 칠하고, 현재 글꼴로 읽힌다', () => {
    const next = run(stateOf(['hello'], [1, 6]), applyMark(marks.fontFamily, { ascii: '굴림', eastAsia: '굴림' }));
    expect(next.doc.child(0).child(0).marks[0].attrs).toEqual({ ascii: '굴림', eastAsia: '굴림' });
    expect(currentFontName(next)).toBe('굴림');
  });

  test('글꼴 크기 크게/작게는 Word 단계(…10, 10.5, 11, 12, 14…)를 밟는다', () => {
    let state = stateOf(['abc'], [1, 4]);
    state = run(state, stepFontSize(1, 10));
    expect(state.doc.child(0).child(0).marks[0].attrs.size).toBe(10.5);
    state = run(state, stepFontSize(1, 10));
    expect(state.doc.child(0).child(0).marks[0].attrs.size).toBe(11);
    state = run(state, stepFontSize(-1, 10));
    state = run(state, stepFontSize(-1, 10));
    expect(state.doc.child(0).child(0).marks[0].attrs.size).toBe(10);
  });

  test('대/소문자 바꾸기: 서식은 그대로, 글자만 바뀐다', () => {
    let state = stateOf(['hello world. again'], [1, 19]);
    state = run(state, applyMark(marks.strong));
    expect(run(state, changeCase('upper')).doc.textContent).toBe('HELLO WORLD. AGAIN');
    expect(run(state, changeCase('title')).doc.textContent).toBe('Hello World. Again');
    expect(run(state, changeCase('sentence')).doc.textContent).toBe('Hello world. Again');
    expect(run(state, changeCase('upper')).doc.child(0).child(0).marks.map((m) => m.type.name)).toEqual(['strong']);
  });

  test('서식 지우기: 글자 마크와 문단 스타일을 지우되 강조색은 남긴다', () => {
    let state = stateOf(['abc'], [1, 4]);
    state = run(state, applyMark(marks.strong));
    state = run(state, applyMark(marks.highlight, { color: '#ffff00' }));
    state = run(state, setParagraphStyle('Quote'));
    state = run(state, clearFormatting);
    expect(state.doc.child(0).child(0).marks.map((m) => m.type.name)).toEqual(['highlight']);
    expect(state.doc.child(0).attrs.styleId).toBeNull();
  });

  test('들여쓰기: 목록 밖에서는 왼쪽 들여쓰기를 800dxa씩, 0 밑으로는 안 내려간다', () => {
    let state = stateOf(['abc'], [2, 2]);
    state = run(state, changeIndent(1));
    expect(state.doc.child(0).attrs.indent).toEqual({ left: 800, firstLine: 0, hanging: 0 });
    state = run(state, changeIndent(-1));
    state = run(state, changeIndent(-1));
    expect(state.doc.child(0).attrs.indent).toBeNull();
  });

  test('줄 간격·음영은 선택한 모든 문단에, 앞뒤 간격은 유지', () => {
    let state = stateOf(['a', 'b'], [1, 5]);
    state = run(state, setLineSpacing(1.5));
    state = run(state, setShading('#d9d9d9'));
    for (let i = 0; i < 2; i += 1) {
      expect(state.doc.child(i).attrs.spacing).toEqual({ line: 360, lineRule: 'auto', before: null, after: null });
      expect(state.doc.child(i).attrs.shading).toBe('#d9d9d9');
    }
  });

  test('스타일 갤러리: 제목 ↔ 문단 스타일 전환에 정렬은 남고, 현재 스타일로 읽힌다', () => {
    let state = stateOf(['abc'], [2, 2]);
    state = run(state, setHeading(1));
    expect(currentParagraphStyle(state)).toBe('Heading1');
    state = run(state, setParagraphStyle('Title'));
    expect(state.doc.child(0).type.name).toBe('paragraph');
    expect(currentParagraphStyle(state)).toBe('Title');
    state = run(state, setParagraphStyle(null));
    expect(currentParagraphStyle(state)).toBeNull();
  });

  test('서식 복사 → 다른 글자에 칠하기', () => {
    let state = stateOf(['source target'], [1, 7]);
    state = run(state, applyMark(marks.em));
    state = run(state, applyMark(marks.textColor, { color: '#c00000' }));
    const format = copyFormat(state.apply(state.tr.setSelection(TextSelection.create(state.doc, 2, 2))));
    state = state.apply(state.tr.setSelection(TextSelection.create(state.doc, 8, 14)));
    state = run(state, pasteFormat(format));
    const target = state.doc.child(0).child(2);
    expect(target.text).toBe('target');
    expect(target.marks.map((m) => m.type.name).sort()).toEqual(['em', 'textColor']);
  });

  test('모두 바꾸기: 대소문자 무시, 원래 서식 유지, 바꾼 개수', () => {
    let state = stateOf(['Cat cat CAT'], [1, 1]);
    state = state.apply(state.tr.addMark(1, 4, marks.strong.create()));
    let count = 0;
    count = replaceAll(state, (tr) => {
      state = state.apply(tr);
    }, 'cat', 'dog');
    expect(count).toBe(3);
    expect(state.doc.textContent).toBe('dog dog dog');
    expect(state.doc.child(0).child(0).marks.map((m) => m.type.name)).toEqual(['strong']);
  });
});

describe('Word 레이아웃 탭 명령', () => {
  test('페이지 나누기·들여쓰기 왼쪽/오른쪽·단락 앞/뒤 간격', async () => {
    const { insertPageBreak, setIndentSide, setSpacingSide, currentParagraphBox } = await import('../src/client/wordCommands.js');
    let state = stateOf(['abc'], [2, 2]);
    state = run(state, insertPageBreak);
    expect(state.doc.child(0).child(1).type.name).toBe('page_break');
    state = run(state, setIndentSide('right', 567));
    state = run(state, setIndentSide('left', 1134));
    state = run(state, setSpacingSide('after', 240));
    expect(currentParagraphBox(state)).toEqual({ left: 1134, right: 567, before: null, after: 240 });
  });
});
