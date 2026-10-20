/**
 * PDF 찾기 — 입력하는 대로 1쪽부터 차례로 찾아 개수를 늘려 가고(긴 문서도 첫 결과가 바로 보인다), Enter·Shift+Enter로
 * 이동, 본문 글자 위에 노란 상자. 새 입력이 오면 앞선 찾기를 멈춘다.
 */
import { findInPage, type FindMatch } from './findText.js';
import type { PageView } from './pageView.js';

export interface FindUi {
  readonly bar: HTMLElement;
  readonly input: HTMLInputElement;
  readonly count: HTMLElement;
  readonly prev: HTMLButtonElement;
  readonly next: HTMLButtonElement;
  readonly close: HTMLButtonElement;
}

export class FindController {
  private matches: FindMatch[] = [];
  private index = -1;
  private generation = 0;
  private timer = 0;
  private searching = false;

  constructor(
    private readonly view: PageView,
    private readonly ui: FindUi,
  ) {
    ui.input.addEventListener('input', () => {
      clearTimeout(this.timer);
      this.timer = window.setTimeout(() => void this.run(ui.input.value), 200);
    });
    ui.input.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') {
        event.preventDefault();
        this.step(event.shiftKey ? -1 : 1);
      } else if (event.key === 'Escape') {
        event.preventDefault();
        this.close();
      }
    });
    ui.next.addEventListener('click', () => this.step(1));
    ui.prev.addEventListener('click', () => this.step(-1));
    ui.close.addEventListener('click', () => this.close());
    // 쪽 정리로 쪽 번호가 바뀌면 찾던 것을 다시 찾는다
    view.onPagesChange(() => {
      if (!ui.bar.hidden && ui.input.value.trim()) void this.run(ui.input.value);
    });
  }

  open(initial?: string): void {
    this.ui.bar.hidden = false;
    if (initial !== undefined && initial !== this.ui.input.value) {
      this.ui.input.value = initial;
      void this.run(initial);
    }
    this.ui.input.focus();
    this.ui.input.select();
  }

  close(): void {
    this.ui.bar.hidden = true;
    this.generation += 1;
    this.matches = [];
    this.index = -1;
    this.view.highlight([], null);
    this.ui.count.textContent = '';
  }

  private render(): void {
    const total = this.matches.length;
    const shown = this.index >= 0 ? `${this.index + 1}/${total}` : total > 0 ? `0/${total}` : this.ui.input.value.trim() ? (this.searching ? '찾는 중' : '없음') : '';
    this.ui.count.textContent = this.searching && total > 0 ? `${shown}…` : shown;
    this.view.highlight(this.matches, this.matches[this.index] ?? null);
  }

  private async run(query: string): Promise<void> {
    const gen = ++this.generation;
    this.matches = [];
    this.index = -1;
    this.searching = query.trim() !== '';
    this.render();
    if (!this.searching) return;
    // 지금 쪽부터 찾기 시작해 첫 결과를 가까운 곳에서 보여 준다(엑셀·워드의 찾기와 같다)
    const start = this.view.currentPage;
    const total = this.view.pageCount;
    // 처음 찾은 결과(지금 쪽에서 가장 가까운 것)를 "현재 결과"로 잡고, 뒤 쪽 결과가 더해져 정렬이 바뀌어도 같은 것을 가리킨다
    let currentMatch: FindMatch | null = null;
    const found: FindMatch[] = [];
    for (let k = 0; k < total; k += 1) {
      const page = ((start - 1 + k) % total) + 1;
      const items = await this.view.textItems(page);
      if (gen !== this.generation) return;
      const pageMatches = findInPage(page, items.map((str) => ({ str })), query);
      if (pageMatches.length === 0) continue;
      found.push(...pageMatches);
      this.matches = [...found].sort((a, b) => a.page - b.page || a.itemStart - b.itemStart || a.offsetStart - b.offsetStart);
      if (!currentMatch) {
        currentMatch = pageMatches[0];
        void this.view.scrollToMatch(currentMatch);
      }
      this.index = this.matches.indexOf(currentMatch);
      this.render();
    }
    if (gen !== this.generation) return;
    this.searching = false;
    this.render();
  }

  private step(delta: 1 | -1): void {
    if (this.matches.length === 0) {
      if (this.ui.input.value.trim()) void this.run(this.ui.input.value);
      return;
    }
    this.index = (this.index + delta + this.matches.length) % this.matches.length;
    this.render();
    void this.view.scrollToMatch(this.matches[this.index]);
  }
}
