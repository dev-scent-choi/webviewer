import type { PagesBody } from './pagesBody.js';
import type { ViewInfo } from './types.js';

/** 껍데기 모듈(막대·썸네일·찾기·표시·측정·편집 단추)이 함께 쓰는 것 — view.ts가 만든다 */
export interface ShellCtx {
  readonly token: string;
  /** `token=…`(인코딩됨) */
  readonly tokenQuery: string;
  readonly info: ViewInfo;
  /** 안쪽 — pdf.js가 실패해도 같은 객체가 원천만 바꾼다(replaceSource). 모듈은 body를 꺼내 두지 말고 ctx.body로 쓴다 */
  readonly body: PagesBody;
  /** Esc — priority가 작은 것부터(측정 10 · 쓰던 글 20 · 글 도구 30 · 강조 40 · 찾기 90), true를 돌려준 곳에서 멈춘다 */
  onEscape(priority: number, handler: () => boolean): void;
  /** 쪽 위 한 줄 안내(빈 글이면 숨김, 8초 뒤 저절로 숨김) */
  notice(text: string): void;
}
