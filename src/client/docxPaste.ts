import { Plugin, PluginKey } from 'prosemirror-state';
import { Fragment, Slice, type Node as PMNode } from 'prosemirror-model';

/**
 * 붙여넣기 지킴이 — 그림·글상자·번호 같은 보존 조각은 원본 XML이 서버 저장소(이 문서를 연 방)에만 있어, 같은 문서 안에서 옮기거나
 * 복사한 것만 저장된다. 다른 문서에서 붙인 조각은 화면엔 보여도 저장 때 빠지므로, 붙여넣는 순간 빼고 알린다(보이는 것 = 저장되는 것).
 */
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

export function docxPastePlugin(onDropped: () => void): Plugin<Set<string>> {
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
    props: {
      transformPasted(slice, view) {
        const known = pasteKey.getState(view.state) ?? new Set<string>();
        let dropped = false;
        const keep = (ref: unknown) => typeof ref !== 'string' || !ref.includes('@') || known.has(ref) || ((dropped = true), false);
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
