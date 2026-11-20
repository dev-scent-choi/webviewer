import { Plugin, PluginKey, type Transaction } from 'prosemirror-state';
import { DOMSerializer, Fragment, Slice, type Mark, type Node as PMNode, type Schema } from 'prosemirror-model';
import { ReplaceAroundStep, ReplaceStep } from 'prosemirror-transform';
import { ySyncPluginKey } from 'y-prosemirror';
import { docxSchema } from '../docx/schema.js';

/**
 * 붙여넣기 지킴이 — 그림·글상자·번호 같은 보존 조각은 원본 XML이 서버 저장소(이 문서를 연 방)에만 있어, 같은 문서 안에서 옮기거나
 * 복사한 것만 저장된다. 다른 문서에서 붙인 조각은 화면엔 보여도 저장 때 빠지므로, 붙여넣는 순간 빼고 알린다(보이는 것 = 저장되는 것).
 * 다른 문서의 조각은 해시가 같아도(내용이 바이트까지 같은 그림 조각 — 그런데 r:embed가 가리키는 그림 파트는 문서마다 다르다) 받지
 * 않는다: 복사할 때 클립보드에 문서 id를 적어 두고, 다른 문서에서 온 것이면 조각을 모두 뺀다(뜻이 몰래 바뀌지 않게, m5).
 * 새로 친 글자·붙인 글자는 변경 추적 삽입(w:ins·w:moveTo — 원래 작성자 이름)을 달지 않는다(m4): 그 감싸개 안에서 쳐도 추적 없는 글자.
 */
const SOURCE_ATTR = 'data-docx-source';
const REF_ATTRS = ['pPrExtra', 'tcPrExtra', 'trPrExtra', 'tblPrExtra'] as const;
const pasteKey = new PluginKey<Set<string>>('docxPaste');

function refsIn(node: PMNode, into: Set<string>): void {
  const add = (ref: unknown) => {
    if (typeof ref === 'string' && ref.includes('@')) into.add(ref);
  };
  const visit = (n: PMNode) => {
    if (n.type.name === 'ooxml' || n.type.name === 'ooxml_block') add(n.attrs.ref);
    for (const key of REF_ATTRS) if (Array.isArray(n.attrs[key])) (n.attrs[key] as unknown[]).forEach(add);
    for (const mark of n.marks) {
      if (mark.type.name === 'runExtra' && Array.isArray(mark.attrs.refs)) (mark.attrs.refs as unknown[]).forEach(add);
      if (mark.type.name === 'wrap') add(mark.attrs.pr);
    }
    n.forEach(visit);
  };
  visit(node);
}

/** 이 거래가 사용자가 넣은 글자(입력·붙여넣기·끌어 놓기)면, 넣은 범위의 변경 추적 삽입 감싸개를 뗀다 — 다른 참여자·되돌리기(Yjs)는 그대로 */
function stripTrackedWraps(transactions: readonly Transaction[], tr: Transaction): Transaction | null {
  const ranges: Array<[number, number]> = [];
  for (const t of transactions) {
    if (!t.docChanged || (t.getMeta(ySyncPluginKey) as { isChangeOrigin?: boolean } | undefined)?.isChangeOrigin) continue;
    t.steps.forEach((step, i) => {
      if (!(step instanceof ReplaceStep || step instanceof ReplaceAroundStep) || step.slice.size === 0) return;
      step.getMap().forEach((_oldStart, _oldEnd, newStart, newEnd) => {
        // 이 거래 뒤의 단계·같은 묶음의 다음 거래로 옮긴 범위
        let from = newStart;
        let to = newEnd;
        for (let k = i + 1; k < t.steps.length; k += 1) {
          from = t.steps[k].getMap().map(from, 1);
          to = t.steps[k].getMap().map(to, -1);
        }
        const after = transactions.slice(transactions.indexOf(t) + 1);
        for (const later of after) {
          from = later.mapping.map(from, 1);
          to = later.mapping.map(to, -1);
        }
        if (to > from) ranges.push([from, to]);
      });
    });
  }
  if (ranges.length === 0) return null;
  const doc = transactions[transactions.length - 1].doc;
  const removals: Array<[number, number, Mark]> = [];
  for (const [from, to] of ranges) {
    doc.nodesBetween(from, Math.min(to, doc.content.size), (node, pos) => {
      if (!node.isInline) return true;
      for (const mark of node.marks)
        if (mark.type.name === 'wrap' && (mark.attrs.tag === 'w:ins' || mark.attrs.tag === 'w:moveTo')) removals.push([Math.max(from, pos), Math.min(to, pos + node.nodeSize), mark]);
      return false;
    });
  }
  if (removals.length === 0) return null;
  for (const [from, to, mark] of removals) tr.removeMark(from, to, mark);
  return tr.setMeta('addToHistory', true);
}

/** 복사할 때 클립보드 HTML 끝에 이 문서 id 표지를 붙인다(붙일 때 transformPastedHTML이 본다) */
function sourceMarkingSerializer(schema: Schema, docId: string): DOMSerializer {
  const base = DOMSerializer.fromSchema(schema);
  const serializer = Object.create(base) as DOMSerializer;
  serializer.serializeFragment = (fragment, options, target) => {
    const out = base.serializeFragment(fragment, options, target);
    // 첫 요소에 속성으로 — 빈 표지 요소를 따로 넣으면 붙일 때 빈 문단이 생길 수 있다
    const first = out.firstChild;
    if (first instanceof Element) first.setAttribute(SOURCE_ATTR, docId);
    else {
      const marker = ((options?.document ?? document) as Document).createElement('span');
      marker.setAttribute(SOURCE_ATTR, docId);
      marker.textContent = '';
      out.insertBefore(marker, first);
    }
    return out;
  };
  return serializer;
}

export function docxPastePlugin(docId: string, onDropped: () => void): Plugin<Set<string>> {
  /** 지금 붙이는 내용이 다른 문서(또는 출처 모를 편집기 HTML)에서 왔나 — transformPastedHTML이 정하고 transformPasted가 쓴다 */
  let foreign = false;
  return new Plugin<Set<string>>({
    key: pasteKey,
    state: {
      init: (_, state) => {
        const known = new Set<string>();
        refsIn(state.doc, known);
        return known;
      },
      // 문서에 들어온 조각(처음 동기화·다른 참여자 편집·같은 문서 붙여넣기)은 이 문서의 것으로 기억한다 — 잘라낸 뒤 붙여도 알아본다
      apply: (tr, known) => {
        if (tr.docChanged) for (const step of tr.steps) {
          const slice = (step as unknown as { slice?: Slice }).slice;
          slice?.content.forEach((n) => refsIn(n, known));
        }
        return known;
      },
    },
    appendTransaction: (transactions, _old, state) => stripTrackedWraps(transactions, state.tr),
    props: {
      clipboardSerializer: sourceMarkingSerializer(docxSchema, docId),
      transformPastedHTML(html) {
        const source = new RegExp(`${SOURCE_ATTR}="([^"]*)"`).exec(html)?.[1];
        foreign = source !== docId;
        return html;
      },
      transformPasted(slice, view) {
        const known = pasteKey.getState(view.state) ?? new Set<string>();
        const fromOther = foreign;
        foreign = false;
        let dropped = false;
        const keep = (ref: unknown) => typeof ref !== 'string' || !ref.includes('@') || (!fromOther && known.has(ref)) || ((dropped = true), false);
        const clean = (fragment: Fragment): Fragment => {
          const out: PMNode[] = [];
          fragment.forEach((n) => {
            if ((n.type.name === 'ooxml' || n.type.name === 'ooxml_block') && !keep(n.attrs.ref)) return;
            const attrs: Record<string, unknown> = { ...n.attrs };
            for (const key of REF_ATTRS) {
              if (Array.isArray(attrs[key])) {
                const refs = (attrs[key] as unknown[]).filter(keep);
                attrs[key] = refs.length > 0 ? refs : null;
              }
            }
            const marks = n.marks.map((m) =>
              m.type.name === 'runExtra' && Array.isArray(m.attrs.refs)
                ? m.type.create({ ...m.attrs, refs: (m.attrs.refs as unknown[]).filter(keep) })
                : m.type.name === 'wrap' && !keep(m.attrs.pr)
                  ? m.type.create({ ...m.attrs, pr: null })
                  : m,
            );
            out.push(n.isText ? n.mark(marks) : (n.type.createAndFill(attrs, clean(n.content), marks) ?? n));
          });
          return Fragment.from(out);
        };
        const result = new Slice(clean(slice.content), slice.openStart, slice.openEnd);
        if (dropped) onDropped();
        return result;
      },
    },
  });
}
