/**
 * Word "검토" 탭 — 맞춤법 검사 표시, 단어 개수, 메모(새로 만들기·삭제·이전/다음·메모 창), 변경 내용 창.
 * 메모 범위는 문서의 comment 마크, 메모 내용은 공유 Y.Map(DOCX_COMMENTS_MAP)에 있다(저장 시 export.ts가
 * word/comments.xml로 쓴다).
 */
import type { Node as PMNode } from 'prosemirror-model';
import { TextSelection, type Command, type EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type * as Y from 'yjs';
import { docxSchema, type DocxComment } from '../docx/schema.js';

export interface CommentRange {
  readonly id: string;
  readonly from: number;
  readonly to: number;
}

/** 문서 순서대로 메모 범위(같은 메모가 여러 조각이면 처음~끝) */
export function commentRanges(doc: PMNode): CommentRange[] {
  const ranges = new Map<string, { from: number; to: number }>();
  doc.descendants((node, pos) => {
    // 글자뿐 아니라 페이지 나누기 같은 인라인 노드에도 메모가 달릴 수 있다
    if (!node.isInline) return;
    for (const mark of node.marks) {
      if (mark.type !== docxSchema.marks.comment) continue;
      const id = String(mark.attrs.id);
      const prev = ranges.get(id);
      const end = pos + node.nodeSize;
      ranges.set(id, prev ? { from: Math.min(prev.from, pos), to: Math.max(prev.to, end) } : { from: pos, to: end });
    }
  });
  return [...ranges.entries()].map(([id, r]) => ({ id, ...r })).sort((a, b) => a.from - b.from || a.to - b.to);
}

/** 커서·선택이 걸쳐 있는 메모 id(가장 안쪽 것) */
export function commentAt(state: EditorState): string | null {
  const { $from } = state.selection;
  const marks = state.selection.empty ? $from.marks() : state.doc.resolve(Math.min($from.pos + 1, state.doc.content.size)).marks();
  const found = marks.filter((mark) => mark.type === docxSchema.marks.comment);
  return found.length > 0 ? String(found[found.length - 1].attrs.id) : null;
}

/** 선택이 없으면 커서가 있는 낱말(공백으로 나뉜)을 메모 범위로 쓴다 — Word와 같다. */
export function commentTarget(state: EditorState): { from: number; to: number } | null {
  const { from, to, empty, $from } = state.selection;
  if (!empty) {
    if (state.doc.textBetween(from, to, '').trim()) return { from, to };
    // 글자 없이 페이지 나누기 같은 노드만 골랐으면 그 문단 전체에 단다(글자 아닌 노드의 마크는 공동 편집에 실리지 않는다)
    const block = $from.parent.isTextblock ? $from.parent : state.doc.resolve(Math.min(from + 1, state.doc.content.size)).parent;
    if (!block.isTextblock || !block.textContent.trim()) return null;
    const start = $from.parent === block ? $from.start() : state.doc.resolve(Math.min(from + 1, state.doc.content.size)).start();
    return { from: start, to: start + block.content.size };
  }
  const parent = $from.parent;
  if (!parent.isTextblock) return null;
  const text = parent.textBetween(0, parent.content.size, '\n', '\u0000');
  const offset = $from.parentOffset;
  let start = offset;
  let end = offset;
  while (start > 0 && !/\s/.test(text[start - 1])) start -= 1;
  while (end < text.length && !/\s/.test(text[end])) end += 1;
  if (start === end) return null;
  return { from: $from.start() + start, to: $from.start() + end };
}

export function addCommentMark(id: string, from: number, to: number): Command {
  return (state, dispatch) => {
    if (from >= to) return false;
    if (dispatch) dispatch(state.tr.addMark(from, to, docxSchema.marks.comment.create({ id })));
    return true;
  };
}

/** 그 id의 메모 마크만 문서 전체에서 지운다(겹친 다른 메모는 그대로). */
export function removeCommentMark(id: string): Command {
  return (state, dispatch) => {
    if (dispatch) {
      const tr = state.tr;
      state.doc.descendants((node, pos) => {
        if (!node.isInline) return;
        for (const mark of node.marks) {
          if (mark.type === docxSchema.marks.comment && String(mark.attrs.id) === id) tr.removeMark(pos, pos + node.nodeSize, mark);
        }
      });
      dispatch(tr);
    }
    return true;
  };
}

export interface WordStats {
  readonly words: number;
  readonly charsNoSpaces: number;
  readonly charsWithSpaces: number;
  readonly paragraphs: number;
}

/** Word "단어 개수" 대화상자와 같은 항목 — from/to를 주면 그 범위만 센다. */
export function wordStats(doc: PMNode, from = 0, to = doc.content.size): WordStats {
  let words = 0;
  let charsNoSpaces = 0;
  let charsWithSpaces = 0;
  let paragraphs = 0;
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isTextblock) return true;
    const start = Math.max(from, pos + 1);
    const end = Math.min(to, pos + 1 + node.content.size);
    const text = start < end ? doc.textBetween(start, end, '', ' ') : '';
    if (text.trim()) {
      paragraphs += 1;
      words += text.trim().split(/\s+/).length;
    }
    charsWithSpaces += text.length;
    charsNoSpaces += text.replace(/\s/g, '').length;
    return false;
  });
  return { words, charsNoSpaces, charsWithSpaces, paragraphs };
}

interface ReviewOptions {
  readonly view: EditorView;
  readonly comments: Y.Map<DocxComment>;
  readonly token: string;
  readonly viewer: string;
  readonly canEdit: boolean;
  readonly run: (command: Command) => void;
}

interface ParagraphChange {
  readonly paragraph: number | null;
  readonly before: string;
  readonly after: string;
  readonly viewer: string | null;
  readonly at: string;
}

function formatWhen(iso: string): string {
  const date = new Date(iso);
  if (Number.isNaN(date.getTime())) return '';
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function newCommentId(): string {
  return typeof crypto !== 'undefined' && 'randomUUID' in crypto ? crypto.randomUUID() : `c${Date.now()}${Math.random().toString(36).slice(2)}`;
}

/** 검토 탭 버튼·메모 창·변경 내용 창·단어 개수 대화상자를 연결한다. 반환값은 문서가 바뀔 때 부를 갱신 함수. */
export function setupWordReview(options: ReviewOptions): (state: EditorState) => void {
  const { view, comments, token, viewer, canEdit, run } = options;
  const pane = document.getElementById('review-pane') as HTMLElement;
  const paneTitle = document.getElementById('review-pane-title') as HTMLElement;
  const paneBody = document.getElementById('review-pane-body') as HTMLElement;
  const paneRefresh = document.getElementById('review-pane-refresh') as HTMLElement;
  const showCommentsBtn = document.getElementById('btn-show-comments');
  const showChangesBtn = document.getElementById('btn-show-changes');
  let mode: 'comments' | 'changes' | null = null;
  let activeId: string | null = null;
  const activeStyle = document.createElement('style');
  document.head.append(activeStyle);

  function setMode(next: typeof mode): void {
    mode = next;
    pane.hidden = mode === null;
    document.body.classList.toggle('review-pane-open', mode !== null);
    showCommentsBtn?.setAttribute('aria-pressed', String(mode === 'comments'));
    showChangesBtn?.setAttribute('aria-pressed', String(mode === 'changes'));
    paneTitle.textContent = mode === 'changes' ? '변경 내용' : '메모';
    paneRefresh.hidden = mode !== 'changes';
    if (mode === 'comments') renderComments();
    if (mode === 'changes') void renderChanges();
  }

  function selectRange(range: CommentRange): void {
    const tr = view.state.tr.setSelection(TextSelection.create(view.state.doc, range.from, range.to)).scrollIntoView();
    view.dispatch(tr);
    view.focus();
  }

  function renderComments(): void {
    if (mode !== 'comments') return;
    const ranges = commentRanges(view.state.doc);
    paneBody.replaceChildren();
    if (ranges.length === 0) {
      const empty = document.createElement('p');
      empty.className = 'review-empty';
      empty.textContent = canEdit ? '메모가 없습니다. 글자를 선택하고 "새 메모"를 누르세요.' : '메모가 없습니다.';
      paneBody.append(empty);
      return;
    }
    for (const range of ranges) {
      const comment = comments.get(range.id);
      const card = document.createElement('article');
      card.className = 'comment-card';
      card.dataset.commentId = range.id;
      card.classList.toggle('is-active', range.id === activeId);
      const head = document.createElement('header');
      const author = document.createElement('strong');
      author.textContent = comment?.author || '(이름 없음)';
      const when = document.createElement('time');
      when.textContent = comment?.date ? formatWhen(comment.date) : '';
      head.append(author, when);
      if (canEdit) {
        const del = document.createElement('button');
        del.className = 'icon-btn comment-delete';
        del.title = '메모 삭제';
        del.setAttribute('aria-label', '메모 삭제');
        del.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M18 6 6 18"/><path d="m6 6 12 12"/></svg>';
        del.addEventListener('click', (event) => {
          event.stopPropagation();
          deleteComment(range.id);
        });
        head.append(del);
      }
      const quote = document.createElement('blockquote');
      quote.textContent = view.state.doc.textBetween(range.from, range.to, ' ').slice(0, 80);
      const text = document.createElement('textarea');
      text.className = 'comment-text';
      text.value = comment?.text ?? '';
      text.readOnly = !canEdit;
      text.rows = Math.min(6, Math.max(2, text.value.split('\n').length));
      text.placeholder = '메모를 입력하세요';
      text.addEventListener('input', () => {
        const current = comments.get(range.id);
        if (current) comments.set(range.id, { ...current, text: text.value });
      });
      card.append(head, quote, text);
      card.addEventListener('click', (event) => {
        if (event.target === text) return;
        selectRange(range);
      });
      paneBody.append(card);
    }
  }

  async function renderChanges(): Promise<void> {
    paneBody.replaceChildren();
    const loading = document.createElement('p');
    loading.className = 'review-empty';
    loading.textContent = '불러오는 중…';
    paneBody.append(loading);
    let changes: ParagraphChange[] = [];
    try {
      const res = await fetch(`/api/view/changes?token=${encodeURIComponent(token)}`);
      const json = (await res.json()) as { data?: { changes?: ParagraphChange[] } };
      changes = json.data?.changes ?? [];
    } catch {
      loading.textContent = '변경 내용을 불러오지 못했습니다.';
      return;
    }
    if (mode !== 'changes') return;
    paneBody.replaceChildren();
    if (changes.length === 0) {
      loading.textContent = '이 편집 세션이 열린 뒤 바뀐 문단이 없습니다.';
      paneBody.append(loading);
      return;
    }
    for (const change of changes) {
      const card = document.createElement('article');
      card.className = 'change-card';
      const head = document.createElement('header');
      const who = document.createElement('strong');
      who.textContent = change.viewer ?? '(알 수 없음)';
      const where = document.createElement('span');
      where.textContent = change.paragraph === null ? '삭제된 문단' : `${change.paragraph}번째 문단`;
      const when = document.createElement('time');
      when.textContent = formatWhen(change.at);
      head.append(who, where, when);
      card.append(head);
      if (change.before) {
        const del = document.createElement('del');
        del.textContent = change.before;
        card.append(del);
      }
      if (change.after) {
        const ins = document.createElement('ins');
        ins.textContent = change.after;
        card.append(ins);
      }
      paneBody.append(card);
    }
  }

  function createComment(): void {
    if (!canEdit) return;
    const target = commentTarget(view.state);
    if (!target) {
      window.alert('메모를 달 글자를 선택하세요.');
      return;
    }
    const id = newCommentId();
    comments.set(id, { id, author: viewer, initials: viewer.slice(0, 1), date: new Date().toISOString(), text: '' });
    run(addCommentMark(id, target.from, target.to));
    activeId = id;
    setMode('comments');
    const card = paneBody.querySelector<HTMLElement>(`[data-comment-id="${CSS.escape(id)}"] textarea`);
    card?.focus();
  }

  function deleteComment(id: string | null): void {
    if (!canEdit || !id) return;
    run(removeCommentMark(id));
    comments.delete(id);
    if (activeId === id) activeId = null;
    renderComments();
  }

  function step(direction: 1 | -1): void {
    const ranges = commentRanges(view.state.doc);
    if (ranges.length === 0) return;
    const pos = view.state.selection.from;
    const target =
      direction > 0
        ? (ranges.find((r) => r.from > pos) ?? ranges[0])
        : ([...ranges].reverse().find((r) => r.to < pos) ?? ranges[ranges.length - 1]);
    activeId = target.id;
    selectRange(target);
    if (mode === 'comments') renderComments();
  }

  const on = (id: string, handler: () => void) => document.getElementById(id)?.addEventListener('click', handler);
  on('btn-new-comment', createComment);
  on('btn-delete-comment', () => deleteComment(commentAt(view.state) ?? activeId));
  on('btn-prev-comment', () => step(-1));
  on('btn-next-comment', () => step(1));
  on('btn-show-comments', () => setMode(mode === 'comments' ? null : 'comments'));
  on('btn-show-changes', () => setMode(mode === 'changes' ? null : 'changes'));
  on('review-pane-close', () => setMode(null));
  on('review-pane-refresh', () => void renderChanges());
  if (!canEdit) {
    for (const id of ['btn-new-comment', 'btn-delete-comment']) (document.getElementById(id) as HTMLButtonElement | null)?.setAttribute('disabled', '');
  }
  window.addEventListener('keydown', (event) => {
    if ((event.ctrlKey || event.metaKey) && event.altKey && event.key.toLowerCase() === 'm') {
      event.preventDefault();
      createComment();
    }
  });

  // ---------- 맞춤법 검사 표시 ----------
  const spellBtn = document.getElementById('btn-spellcheck');
  view.dom.spellcheck = true;
  spellBtn?.addEventListener('click', () => {
    view.dom.spellcheck = !view.dom.spellcheck;
    spellBtn.setAttribute('aria-pressed', String(view.dom.spellcheck));
    // 브라우저는 spellcheck를 바꿔도 이미 그려진 밑줄을 바로 지우지 않는다 — 한 번 포커스를 오가게 한다
    view.dom.blur();
    view.focus();
  });

  // ---------- 단어 개수 ----------
  const dialog = document.getElementById('word-count-dialog') as HTMLElement;
  const statsEl = document.getElementById('word-count-stats') as HTMLElement;
  const scopeEl = document.getElementById('word-count-scope') as HTMLElement;
  function openWordCount(): void {
    const { from, to, empty } = view.state.selection;
    const stats = empty ? wordStats(view.state.doc) : wordStats(view.state.doc, from, to);
    scopeEl.textContent = empty ? '문서 전체' : '선택한 부분';
    const rows: Array<[string, number]> = [
      ['단어', stats.words],
      ['문자(공백 제외)', stats.charsNoSpaces],
      ['문자(공백 포함)', stats.charsWithSpaces],
      ['단락', stats.paragraphs],
    ];
    statsEl.replaceChildren(
      ...rows.flatMap(([label, value]) => {
        const dt = document.createElement('dt');
        dt.textContent = label;
        const dd = document.createElement('dd');
        dd.textContent = value.toLocaleString();
        return [dt, dd];
      }),
    );
    dialog.hidden = false;
    (document.getElementById('word-count-close') as HTMLElement).focus();
  }
  on('btn-word-count', openWordCount);
  on('word-count-close', () => {
    dialog.hidden = true;
    view.focus();
  });
  dialog.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      dialog.hidden = true;
      view.focus();
    }
  });
  dialog.addEventListener('click', (event) => {
    if (event.target === dialog) dialog.hidden = true;
  });

  comments.observe(() => {
    // 다른 사람이 메모 내용을 고치는 중 내가 쓰고 있는 칸까지 다시 그리면 커서가 튄다 — 입력 중이면 건너뛴다
    if (!(document.activeElement instanceof HTMLTextAreaElement && pane.contains(document.activeElement))) renderComments();
  });

  let renderTimer: ReturnType<typeof setTimeout> | undefined;
  return (state: EditorState) => {
    const current = commentAt(state);
    if (current !== activeId && current !== null) activeId = current;
    activeStyle.textContent = activeId
      ? `.editor-surface .ProseMirror .comment-mark[data-comment-id="${CSS.escape(activeId)}"] { background: #ffd966; }`
      : '';
    for (const card of paneBody.querySelectorAll<HTMLElement>('.comment-card')) {
      card.classList.toggle('is-active', card.dataset.commentId === activeId);
    }
    clearTimeout(renderTimer);
    renderTimer = setTimeout(() => {
      if (!(document.activeElement instanceof HTMLTextAreaElement && pane.contains(document.activeElement))) renderComments();
    }, 300);
  };
}
