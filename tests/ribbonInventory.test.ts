import fs from 'node:fs';
import { describe, expect, it } from 'vitest';
import { INVENTORY_PATH, loadInventory, requiredIcons, serializeInventory, summarize, validateInventory } from '../scripts/ribbonGenLib.mjs';

describe('오피스 리본 목록(src/client/ribbon/inventory.json)', () => {
  const inv = loadInventory();
  it('규칙 위반 없음(겹침·한글 이름·종류·출처 하나·자리 결정 note·줄 순서·요약 수)', () => {
    expect(validateInventory(inv)).toEqual([]);
  });
  it('파일 모양이 serializeInventory 그대로(단추 한 줄에 하나)', () => {
    // core.autocrlf=true 인 새 체크아웃은 CRLF로 풀린다 — 줄 끝은 보지 않는다
    expect(fs.readFileSync(INVENTORY_PATH, 'utf8').replaceAll('\r\n', '\n')).toBe(serializeInventory(inv));
  });
  it('기본 탭 순서는 설계 §2 그대로', () => {
    expect(inv.apps.word.tabs.map((t) => t.name)).toEqual(['홈', '삽입', '그리기', '디자인', '레이아웃', '참조', '편지', '검토', '보기']);
    expect(inv.apps.excel.tabs.map((t) => t.name)).toEqual(['홈', '삽입', '그리기', '페이지 레이아웃', '수식', '데이터', '검토', '보기']);
    expect(inv.apps.powerpoint.tabs.map((t) => t.name)).toEqual(['홈', '삽입', '그리기', '디자인', '전환', '애니메이션', '슬라이드 쇼', '녹음/녹화', '검토', '보기']);
  });
  it('기존 탭 id(편집기·시험이 쓰는 것)는 그대로', () => {
    const ids = (a: 'word' | 'excel' | 'powerpoint') => [...inv.apps[a].tabs, ...inv.apps[a].contextual].map((t) => t.id);
    expect(ids('word')).toEqual(expect.arrayContaining(['home', 'insert', 'layout', 'review', 'view', 'table']));
    expect(ids('excel')).toEqual(expect.arrayContaining(['home', 'insert', 'formulas', 'data', 'view', 'chart']));
    expect(ids('powerpoint')).toEqual(expect.arrayContaining(['home', 'insert', 'design', 'transitions', 'animations']));
  });
  it('기존 기능이 줄지 않는다(계획 작성 때 워드 64·엑셀 63·PPT 33, 대리 1·0·3)', () => {
    expect(summarize(inv.apps.word).existing).toBeGreaterThanOrEqual(64);
    expect(summarize(inv.apps.excel).existing).toBeGreaterThanOrEqual(63);
    expect(summarize(inv.apps.powerpoint).existing).toBeGreaterThanOrEqual(33);
    expect(summarize(inv.apps.powerpoint).proxy).toBeGreaterThanOrEqual(3);
  });
  it('unverified(공개 문서로 확인 못 한 이름 — Task 7 사람 확인)는 true만', () => {
    const bad = structuredClone(inv);
    (bad.apps.word.tabs[0].groups[0].controls[0] as { unverified?: unknown }).unverified = 'yes';
    expect(validateInventory(bad).some((x) => x.includes('unverified는 true만'))).toBe(true);
  });
  it('아이콘 키는 영문 소문자·숫자·하이픈', () => {
    for (const k of requiredIcons(inv)) expect(k).toMatch(/^[a-z0-9-]+$/);
  });
});
