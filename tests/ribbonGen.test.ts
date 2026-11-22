import { describe, expect, it } from 'vitest';
import { extractUnits, generatePage, pendingHtml, type RibbonApp } from '../scripts/ribbonGenLib.mjs';

const app: RibbonApp = {
  page: 'x.html',
  backstage: { office: [], existing: [] },
  tabs: [
    {
      id: 'home', name: '홈', editOnly: true, groups: [
        { id: 'clipboard', name: '클립보드', collapse: 9, layout: 'columns', launcher: { pending: '클립보드 창' }, controls: [
          { key: 'paste', name: '붙여넣기', size: 'big', kind: 'button', ref: '#btn-paste', map: 'same' },
          { key: 'cut', name: '잘라내기', size: 'small', kind: 'button', pending: true },
          { key: 'copy', name: '복사', size: 'small', kind: 'split', pending: true },
        ] },
        { id: 'font', name: '글꼴', collapse: 8, layout: 'rows', launcher: null, controls: [
          { key: 'bold', name: '굵게', size: 'small', kind: 'button', ref: '#btn-bold', map: 'same', row: 0, iconOnly: true },
          { key: 'shadow', name: '텍스트 그림자', size: 'small', kind: 'button', pending: true, row: 1, iconOnly: true },
        ] },
      ],
    },
  ],
  contextual: [
    { id: 'table', name: '레이아웃', editOnly: true, trigger: '표 안', wired: null, groups: [
      { id: 'cell-size', name: '셀 크기', collapse: 3, layout: 'fields', launcher: null, controls: [{ key: 'width', name: '너비', size: 'small', kind: 'input', pending: true, stackTitle: '크기' }] },
    ] },
  ],
};
const page = (panels: string) =>
  `<header>\n    <div class="ribbon-tabs">\n      <!-- ribbon:tabs:begin -->\n      <button>옛 탭</button>\n      <!-- ribbon:tabs:end -->\n      <span class="ribbon-tabs-spacer"></span>\n    </div>\n    <!-- ribbon:panels:begin -->\n${panels}\n    <!-- ribbon:panels:end -->\n  </header>\n`;
const OLD = page(`    <div class="ribbon-panel" data-panel="home">
      <button data-rb="home.clipboard.paste" id="btn-paste" class="icon-btn icon-btn-big" title="붙여넣기 (Ctrl+V)">
        <svg viewBox="0 0 24 24"><path d="M1 1"/></svg>
        <span>붙여넣기</span>
      </button>
      <span><button data-rb="home.font.bold" id="btn-bold" class="icon-btn" title="굵게 (Ctrl+Shift+>)"><strong>가</strong></button></span>
      <!-- 옛 주석 -->
      <button type="button" class="ribbon-launcher" data-proxy="btn-x">옛 자세히</button>
    </div>`);

describe('extractUnits', () => {
  it('data-rb 요소를 닫는 태그까지(속성 안의 > 포함) 꺼낸다', () => {
    const units = extractUnits(OLD);
    expect([...units.keys()]).toEqual(['home.clipboard.paste', 'home.font.bold']);
    expect(units.get('home.font.bold')).toBe('<button data-rb="home.font.bold" id="btn-bold" class="icon-btn" title="굵게 (Ctrl+Shift+>)"><strong>가</strong></button>');
    expect(units.get('home.clipboard.paste')).toMatch(/^<button data-rb="home\.clipboard\.paste"[^]*<\/button>$/);
  });
  it('안쪽에 같은 이름 태그가 겹쳐도 짝을 맞춘다', () => {
    const units = extractUnits('<span data-rb="a.b.c" class="dropdown-btn"><button>x</button><span>y<span>z</span></span></span><span>밖</span>');
    expect(units.get('a.b.c')).toBe('<span data-rb="a.b.c" class="dropdown-btn"><button>x</button><span>y<span>z</span></span></span>');
  });
  it('data-rb가 첫 속성이 아니면 멈춘다', () => {
    expect(() => extractUnits('<button id="x" data-rb="a.b.c">x</button>')).toThrow(/첫 속성/);
  });
  it('같은 data-rb가 두 번이면 멈춘다', () => {
    expect(() => extractUnits('<i data-rb="a.b.c"></i><i data-rb="a.b.c"></i>')).toThrow(/두 번/);
  });
});

describe('generatePage', () => {
  it('목록 순서대로 탭·판·묶음·단추를 만들고 기존 단추는 마크업 그대로 옮긴다', () => {
    const out = generatePage(OLD, app);
    expect(out).toContain('<button class="ribbon-tab ribbon-tab-file" role="tab" data-tab="file">파일</button>');
    expect(out).toContain('<button class="ribbon-tab" role="tab" data-tab="home" data-edit-only>홈</button>');
    expect(out).toContain('<button class="ribbon-tab ribbon-tab-contextual" role="tab" data-tab="table" data-edit-only hidden>레이아웃</button>');
    expect(out).not.toMatch(/옛 탭|옛 주석|옛 자세히/);
    expect(out.indexOf('data-rb="home.clipboard.paste"')).toBeLessThan(out.indexOf('data-rb="home.clipboard.cut"'));
    expect(out).toContain('title="굵게 (Ctrl+Shift+>)"');
    expect(out).toContain('<button data-rb="home.clipboard.cut" data-pending data-kind="button" title="준비 중 — 잘라내기" type="button" class="icon-btn ribbon-btn-labeled rb-pending" data-icon="cut" aria-label="잘라내기" disabled><span class="ribbon-btn-label">잘라내기</span></button>');
    expect(out).toContain('<button type="button" class="ribbon-launcher" data-pending disabled title="준비 중 — 클립보드 창" aria-label="클립보드 자세히">');
    expect(out).toContain('<div class="ribbon-group" data-group="font" data-collapse="8" role="group" aria-label="글꼴">');
    expect(out).toContain('class="icon-btn icon-btn-big rb-group-toggle" data-icon="paste" aria-expanded="false" title="클립보드"');
    expect(out).toContain('<div class="ribbon-group-buttons-wrap rb-rows">');
    expect(out).toContain('<span class="layout-field-title">크기</span>');
    expect(out).toMatch(/data-panel="home" role="tabpanel">/);
    expect(out).toMatch(/data-panel="table" role="tabpanel" hidden>/);
  });
  it('두 번 돌려도 같다(결정적)', () => {
    const once = generatePage(OLD, app);
    expect(generatePage(once, app)).toBe(once);
  });
  it('CRLF 파일은 CRLF로 둔다', () => {
    const once = generatePage(OLD.replace(/\n/g, '\r\n'), app);
    expect(once.includes('\r\n')).toBe(true);
    expect(/[^\r]\n/.test(once)).toBe(false);
  });
  it('기존 단추의 data-rb가 없으면·목록에 없는 data-rb가 있으면 멈춘다', () => {
    expect(() => generatePage(OLD.replace('data-rb="home.font.bold" ', ''), app)).toThrow(/home\.font\.bold/);
    expect(() => generatePage(OLD.replace('home.font.bold', 'home.font.zzz'), app)).toThrow(/목록에 없는 data-rb/);
  });
  it('기존 기능 자리에 준비 중 마크업이 있으면 멈춘다', () => {
    const fake = OLD.replace(/<span><button data-rb="home\.font\.bold"[^]*?<\/button><\/span>/, pendingHtml('home.font.bold', { key: 'bold', name: '굵게', size: 'small', kind: 'button' }));
    expect(() => generatePage(fake, app)).toThrow(/준비 중/);
  });
  it('표시가 없으면 멈춘다', () => {
    expect(() => generatePage('<header></header>', app)).toThrow(/표시 없음/);
  });
});
