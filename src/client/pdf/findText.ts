/**
 * PDF 찾기의 순수 부분 — pdf.js 글자 조각(items) 위에서 찾는다. 조각 경계를 넘는 일치도 찾고, 결과는 강조에 쓰도록
 * "조각 번호 + 조각 안 글자 위치"로 돌려준다(offsetEnd는 끝 글자 다음 위치).
 */
export interface TextItemLite {
  readonly str: string;
}

export interface FindMatch {
  readonly page: number;
  readonly itemStart: number;
  readonly offsetStart: number;
  readonly itemEnd: number;
  readonly offsetEnd: number;
}

/** 전각·호환 문자를 펴고(NFKC) 소문자로, 공백 묶음은 한 칸으로 */
export function normalizeForFind(s: string): string {
  return s.normalize('NFKC').toLowerCase().replace(/\s+/g, ' ');
}

export function findInPage(page: number, items: readonly TextItemLite[], query: string): FindMatch[] {
  const q = normalizeForFind(query).trim();
  if (!q) return [];
  // 쪽 글자를 한 줄로 이으면서, 이은 글자의 각 위치가 어느 조각의 몇 번째 글자였는지 적어 둔다.
  // NFKC는 글자 수를 바꿀 수 있어(ﬁ → fi) 조각마다 글자 단위로 정규화해 원래 위치에 매단다.
  let text = '';
  const posItem: number[] = [];
  const posOffset: number[] = [];
  items.forEach((item, i) => {
    let k = 0;
    for (const raw of item.str) {
      const normalized = raw.normalize('NFKC').toLowerCase();
      for (const ch of normalized) {
        const c = /\s/.test(ch) ? ' ' : ch;
        if (c === ' ' && (text === '' || text.endsWith(' '))) continue;
        text += c;
        posItem.push(i);
        posOffset.push(k);
      }
      k += raw.length;
    }
  });
  const out: FindMatch[] = [];
  for (let at = text.indexOf(q); at !== -1; at = text.indexOf(q, at + q.length)) {
    const end = at + q.length - 1;
    const endItem = posItem[end];
    const endRaw = items[endItem].str;
    // 끝 글자가 서로게이트 쌍이면 두 칸을 덮는다
    const endWidth = endRaw.codePointAt(posOffset[end])! > 0xffff ? 2 : 1;
    out.push({ page, itemStart: posItem[at], offsetStart: posOffset[at], itemEnd: endItem, offsetEnd: posOffset[end] + endWidth });
  }
  return out;
}
