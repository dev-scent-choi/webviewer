import fs from 'node:fs';
import { describe, expect, it } from 'vitest';
import { loadInventory, requiredIcons } from '../scripts/ribbonGenLib.mjs';

const SPRITE = 'public/ribbon/icons.svg';
const BUDGET_BYTES = 96 * 1024;

describe('리본 준비 중 아이콘 스프라이트(public/ribbon/icons.svg)', () => {
  const svg = fs.existsSync(SPRITE) ? fs.readFileSync(SPRITE, 'utf8') : '';
  const ids = [...svg.matchAll(/<symbol id="rbi-([a-z0-9-]+)"/g)].map((m) => m[1]);
  it(`크기 상한 ${BUDGET_BYTES} 바이트`, () => {
    expect(svg.length).toBeGreaterThan(0);
    expect(Buffer.byteLength(svg)).toBeLessThanOrEqual(BUDGET_BYTES);
  });
  it('목록이 쓰는 아이콘이 모두 있고 남는 것이 없다', () => {
    const want = requiredIcons(loadInventory());
    expect(want.filter((id) => !ids.includes(id))).toEqual([]);
    expect(ids.filter((id) => !want.includes(id))).toEqual([]);
  });
  it('id가 겹치지 않는다', () => expect(new Set(ids).size).toBe(ids.length));
  it('단색 선 — 24칸, 색·그림·글자·스크립트·바깥 주소 없음', () => {
    for (const m of svg.matchAll(/<symbol [^>]*>/g)) expect(m[0]).toMatch(/viewBox="0 0 24 24"/);
    expect(svg).not.toMatch(/#[0-9a-fA-F]{3,8}\b|rgb\(|<image|<text|<script|<style|<foreignObject|href="(?!#rbi-)/);
  });
  it('처리기 속성(on*=)·xlink:href 없음', () => {
    expect(svg).not.toMatch(/\son[a-z]+\s*=/i);
    expect(svg).not.toMatch(/xlink:href/i);
  });
  it('줄 번호(line-numbers)는 번호 매기기(numbering)와 다른 그림', () => {
    const body = (id: string) => svg.match(new RegExp(`<symbol id="rbi-${id}"[^>]*>(.*?)</symbol>`))?.[1];
    expect(body('line-numbers')).toBeTruthy();
    expect(body('line-numbers')).not.toBe(body('numbering'));
  });
  it('대리 기호(<use>)는 있는 기호만 가리킨다', () => {
    for (const m of svg.matchAll(/<use href="#rbi-([a-z0-9-]+)"/g)) expect(ids).toContain(m[1]);
  });
});
