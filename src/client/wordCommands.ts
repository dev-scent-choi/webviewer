/**
 * Word "홈" 리본 명령 — 전부 브라우저 없이 동작하는 ProseMirror 명령(state, dispatch)이라 단위
 * 테스트로 검증한다(tests/wordCommands.test.ts). 버튼 배선은 editor.ts가 한다.
 */

import type { Mark, MarkType, Node as PMNode } from 'prosemirror-model';
import { liftListItem, sinkListItem, wrapInList } from 'prosemirror-schema-list';
import { TextSelection, type Command, type EditorState, type Transaction } from 'prosemirror-state';
import { docxSchema, type ParagraphIndent, type ParagraphSpacing } from '../docx/schema.js';

const { marks, nodes } = docxSchema;

/** Word의 "글꼴 크기 크게/작게"가 밟는 크기 단계 */
export const FONT_SIZE_STEPS = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72] as const;
/** "들여쓰기 늘리기/줄이기" 한 번의 폭(dxa) — 한국어 Word 기본 탭 간격(40pt) */
export const INDENT_STEP_DXA = 800;

const FORMATTABLE = new Set(['paragraph', 'heading']);

/** 선택이 비어 있으면 "다음에 입력할 글자"에 적용한다(Word와 같다). */
export function applyMark(type: MarkType, attrs: Record<string, unknown> | null = null): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (dispatch) {
      const mark = type.create(attrs);
      dispatch(empty ? state.tr.addStoredMark(mark) : state.tr.addMark(from, to, mark));
    }
    return true;
  };
}

export function clearMark(type: MarkType): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (dispatch) dispatch(empty ? state.tr.removeStoredMark(type) : state.tr.removeMark(from, to, type));
    return true;
  };
}

/** 선택 영역(또는 커서)의 문단·제목에 같은 변경을 적용한다. 선택은 그대로 둔다. */
function updateParagraphs(state: EditorState, dispatch: ((tr: Transaction) => void) | undefined, patch: (node: PMNode) => Record<string, unknown> | null): boolean {
  const { from, to } = state.selection;
  const targets: Array<{ node: PMNode; pos: number }> = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (FORMATTABLE.has(node.type.name)) targets.push({ node, pos });
  });
  if (targets.length === 0) return false;
  if (dispatch) {
    const tr = state.tr;
    for (const { node, pos } of targets) {
      const attrs = patch(node);
      if (attrs) tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...attrs });
    }
    tr.setSelection(TextSelection.between(tr.doc.resolve(from), tr.doc.resolve(Math.min(to, tr.doc.content.size))));
    dispatch(tr);
  }
  return true;
}

export function setAlign(align: 'left' | 'center' | 'right' | 'justify'): Command {
  return (state, dispatch) => updateParagraphs(state, dispatch, () => ({ align: align === 'left' ? null : align }));
}

export function setShading(color: string | null): Command {
  return (state, dispatch) => updateParagraphs(state, dispatch, () => ({ shading: color }));
}

/** 줄 간격 배수(1, 1.15, 1.5 …) — 단락 앞/뒤 간격은 그대로 둔다. */
export function setLineSpacing(multiple: number): Command {
  return (state, dispatch) =>
    updateParagraphs(state, dispatch, (node) => {
      const prev = node.attrs.spacing as ParagraphSpacing | null;
      return { spacing: { line: Math.round(multiple * 240), lineRule: 'auto', before: prev?.before ?? null, after: prev?.after ?? null } };
    });
}

/** 커서가 목록 안이면 목록 수준을, 아니면 왼쪽 들여쓰기를 한 단계 바꾼다(Word와 같다). */
export function changeIndent(direction: 1 | -1): Command {
  return (state, dispatch, view) => {
    const listCommand = direction > 0 ? sinkListItem(nodes.list_item) : liftListItem(nodes.list_item);
    if (listCommand(state, dispatch, view)) return true;
    return updateParagraphs(state, dispatch, (node) => {
      const prev = (node.attrs.indent as ParagraphIndent | null) ?? { left: 0, firstLine: 0, hanging: 0 };
      const left = Math.max(0, prev.left + direction * INDENT_STEP_DXA);
      const next = { ...prev, left };
      return { indent: next.left || next.firstLine || next.hanging ? next : null };
    });
  };
}

/** 다단계 목록: 목록 안이면 한 수준 내리고, 아니면 번호 목록을 시작한다. */
export const multilevelList: Command = (state, dispatch, view) =>
  sinkListItem(nodes.list_item)(state, dispatch, view) || wrapInList(nodes.ordered_list)(state, dispatch, view);

/** 스타일 갤러리의 문단 스타일 — null이면 "표준". 제목 1/2는 heading 노드로 따로 처리한다. */
export function setParagraphStyle(styleId: string | null): Command {
  return (state, dispatch) =>
    updateParagraphsTo(state, dispatch, nodes.paragraph, (node) => ({
      align: node.attrs.align,
      indent: node.attrs.indent,
      spacing: node.attrs.spacing,
      shading: node.attrs.shading,
      styleId,
    }));
}

/** 제목 1/2 등 — 문단 서식(정렬·들여쓰기 등)은 그대로 둔다. */
export function setHeading(level: number): Command {
  return (state, dispatch) =>
    updateParagraphsTo(state, dispatch, nodes.heading, (node) => ({
      level,
      align: node.attrs.align,
      indent: node.attrs.indent,
      spacing: node.attrs.spacing,
      shading: node.attrs.shading,
    }));
}

function updateParagraphsTo(
  state: EditorState,
  dispatch: ((tr: Transaction) => void) | undefined,
  type: typeof nodes.paragraph,
  attrsOf: (node: PMNode) => Record<string, unknown>,
): boolean {
  const { from, to } = state.selection;
  const targets: Array<{ node: PMNode; pos: number }> = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (FORMATTABLE.has(node.type.name)) targets.push({ node, pos });
  });
  if (targets.length === 0) return false;
  if (dispatch) {
    const tr = state.tr;
    for (const { node, pos } of targets) tr.setNodeMarkup(pos, type, attrsOf(node));
    dispatch(tr);
  }
  return true;
}

/** 현재 커서 문단의 스타일 id(제목이면 'Heading1' 같은 값) — 갤러리 눌림 표시용. */
export function currentParagraphStyle(state: EditorState): string | null {
  const parent = state.selection.$from.parent;
  if (parent.type === nodes.heading) return `Heading${parent.attrs.level as number}`;
  if (parent.type === nodes.paragraph) return (parent.attrs.styleId as string | null) ?? null;
  return null;
}

/** 선택 영역의 글자 마크를 모두 지운다(강조색은 Word처럼 남긴다) + 문단을 표준으로 되돌린다. */
export const clearFormatting: Command = (state, dispatch) => {
  const { from, to, empty } = state.selection;
  if (dispatch) {
    const tr = state.tr;
    for (const type of Object.values(marks)) {
      if (type === marks.highlight) continue;
      if (empty) tr.removeStoredMark(type);
      else tr.removeMark(from, to, type);
    }
    state.doc.nodesBetween(from, to, (node, pos) => {
      if (FORMATTABLE.has(node.type.name)) {
        tr.setNodeMarkup(pos, nodes.paragraph, { align: null, indent: null, spacing: null, shading: null, styleId: null });
      }
    });
    dispatch(tr);
  }
  return true;
};

function currentMarkAttr(state: EditorState, type: MarkType): Mark | undefined {
  const { $from, empty, from, to } = state.selection;
  if (empty) return type.isInSet(state.storedMarks ?? $from.marks()) ?? undefined;
  let found: Mark | undefined;
  state.doc.nodesBetween(from, to, (node) => {
    if (!found && node.isText) found = type.isInSet(node.marks) ?? undefined;
  });
  return found;
}

/** 선택 시작 위치의 글자 크기(pt) — 크기 서식이 없으면 문서 기본 크기. */
export function currentFontSize(state: EditorState, defaultSizePt: number): number {
  const mark = currentMarkAttr(state, marks.fontSize);
  return mark ? Number(mark.attrs.size) : defaultSizePt;
}

/** 현재 글꼴 이름(한글 글꼴 우선) — 글꼴 서식이 없으면 null(문서 기본 글꼴). */
export function currentFontName(state: EditorState): string | null {
  const mark = currentMarkAttr(state, marks.fontFamily);
  return mark ? ((mark.attrs.eastAsia as string | null) ?? (mark.attrs.ascii as string | null)) : null;
}

export function stepFontSize(direction: 1 | -1, defaultSizePt: number): Command {
  return (state, dispatch, view) => {
    const size = currentFontSize(state, defaultSizePt);
    const next =
      direction > 0
        ? (FONT_SIZE_STEPS.find((step) => step > size) ?? Math.min(1638, size + 10))
        : ([...FONT_SIZE_STEPS].reverse().find((step) => step < size) ?? Math.max(1, size - 1));
    return applyMark(marks.fontSize, { size: next })(state, dispatch, view);
  };
}

export type CaseMode = 'lower' | 'upper' | 'title' | 'sentence' | 'toggle';

function transformCase(text: string, mode: CaseMode, atSentenceStart: boolean): string {
  switch (mode) {
    case 'lower':
      return text.toLowerCase();
    case 'upper':
      return text.toUpperCase();
    case 'toggle':
      return [...text].map((ch) => (ch === ch.toUpperCase() ? ch.toLowerCase() : ch.toUpperCase())).join('');
    case 'title':
      return text.toLowerCase().replace(/(^|[\s\-(["'])(\p{L})/gu, (_, sep: string, ch: string) => sep + ch.toUpperCase());
    case 'sentence': {
      let start = atSentenceStart;
      return [...text.toLowerCase()]
        .map((ch) => {
          if (start && /\p{L}/u.test(ch)) {
            start = false;
            return ch.toUpperCase();
          }
          if (/[.!?]/.test(ch)) start = true;
          return ch;
        })
        .join('');
    }
  }
}

/** 대/소문자 바꾸기 — 글자 서식은 그대로 두고 글자만 바꾼다. */
export function changeCase(mode: CaseMode): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) return false;
    if (dispatch) {
      const tr = state.tr;
      const pieces: Array<{ start: number; end: number; text: string; marks: readonly Mark[]; sentenceStart: boolean }> = [];
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (!node.isText || !node.text) return;
        const start = Math.max(from, pos);
        const end = Math.min(to, pos + node.text.length);
        const before = state.doc.textBetween(state.doc.resolve(start).start(), start);
        pieces.push({
          start,
          end,
          text: node.text.slice(start - pos, end - pos),
          marks: node.marks,
          sentenceStart: before.trim() === '' || /[.!?]\s*$/.test(before),
        });
      });
      for (const piece of pieces.reverse()) {
        const replaced = transformCase(piece.text, mode, piece.sentenceStart);
        if (replaced !== piece.text) tr.replaceWith(piece.start, piece.end, docxSchema.text(replaced, piece.marks));
      }
      tr.setSelection(TextSelection.create(tr.doc, from, to));
      dispatch(tr);
    }
    return true;
  };
}

/** 서식 복사 — 커서 위치의 글자 마크와 문단 서식을 담아 둔다. */
export interface CopiedFormat {
  readonly marks: readonly Mark[];
  readonly paragraph: Record<string, unknown> | null;
}

export function copyFormat(state: EditorState): CopiedFormat {
  const { $from, empty, from } = state.selection;
  const at = empty ? $from : state.doc.resolve(from + 1 <= state.doc.content.size ? from + 1 : from);
  const parent = $from.parent;
  return {
    marks: empty ? (state.storedMarks ?? $from.marks()) : at.marks(),
    paragraph: FORMATTABLE.has(parent.type.name)
      ? { align: parent.attrs.align, indent: parent.attrs.indent, spacing: parent.attrs.spacing, shading: parent.attrs.shading }
      : null,
  };
}

/** 담아 둔 서식을 선택 영역에 칠한다 — 글자 마크는 통째로 바꾸고, 문단 서식도 함께 옮긴다. */
export function pasteFormat(format: CopiedFormat): Command {
  return (state, dispatch) => {
    const { from, to, empty } = state.selection;
    if (empty) return false;
    if (dispatch) {
      const tr = state.tr;
      for (const type of Object.values(marks)) tr.removeMark(from, to, type);
      for (const mark of format.marks) tr.addMark(from, to, mark);
      if (format.paragraph) {
        state.doc.nodesBetween(from, to, (node, pos) => {
          if (FORMATTABLE.has(node.type.name)) tr.setNodeMarkup(pos, undefined, { ...node.attrs, ...format.paragraph });
        });
      }
      dispatch(tr);
    }
    return true;
  };
}

/** 찾은 위치 하나를 바꾼다 — 바뀐 글자는 원래 글자의 서식을 따른다. */
export function replaceRange(from: number, to: number, replacement: string): Command {
  return (state, dispatch) => {
    if (dispatch) {
      const marksAt = state.doc.resolve(from + (to > from ? 1 : 0)).marks();
      const tr = replacement
        ? state.tr.replaceWith(from, to, docxSchema.text(replacement, marksAt))
        : state.tr.delete(from, to);
      dispatch(tr);
    }
    return true;
  };
}

/** 같은 글자를 모두 바꾼다(대소문자 무시) — 바꾼 개수를 돌려준다. */
export function replaceAll(state: EditorState, dispatch: (tr: Transaction) => void, query: string, replacement: string): number {
  if (!query) return 0;
  const lower = query.toLowerCase();
  const hits: Array<{ from: number; to: number; marks: readonly Mark[] }> = [];
  state.doc.descendants((node, pos) => {
    if (!node.isText || !node.text) return;
    const text = node.text.toLowerCase();
    let idx = text.indexOf(lower);
    while (idx !== -1) {
      hits.push({ from: pos + idx, to: pos + idx + query.length, marks: node.marks });
      idx = text.indexOf(lower, idx + query.length);
    }
  });
  if (hits.length === 0) return 0;
  const tr = state.tr;
  for (const hit of hits.reverse()) {
    if (replacement) tr.replaceWith(hit.from, hit.to, docxSchema.text(replacement, hit.marks));
    else tr.delete(hit.from, hit.to);
  }
  dispatch(tr);
  return hits.length;
}
