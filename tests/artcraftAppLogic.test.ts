import { createHash } from 'node:crypto';
import { afterEach, describe, expect, test, vi } from 'vitest';
import {
  NOTICES, contentRedirect, fontName, frameUrl, isAppJsUrl, isUndoKey, openName, readAppMessage, readOnlyCommands, saveDecision, saveFailure, sha256Etag, toBase64,
  uiShowsDocument, unprotectedSheets, createSaveQueue, fetchBytes, once, withTimeout,
} from '../src/client/artcraftAppLogic.js';

const JS = '/artcraft-app/wordcraft-0123456789/wordcraft-web.js';

describe('isAppJsUrl — iframe이 불러도 되는 앱 JS(같은 출처 해시 폴더의 그 앱 것만)', () => {
  test.each([
    ['word', JS, true],
    ['grid', '/artcraft-app/gridcraft-abcdef0123/gridcraft-web.js', true],
    ['deck', JS, false],
    ['word', 'https://evil.example/artcraft-app/wordcraft-0123456789/wordcraft-web.js', false],
    ['word', '//evil.example/artcraft-app/wordcraft-0123456789/wordcraft-web.js', false],
    ['word', '/artcraft-app/wordcraft-0123456789/../x/wordcraft-web.js', false],
    ['word', `${JS}?x=1`, false],
    ['word', 'data:text/javascript,alert(1)', false],
    ['nope', JS, false],
  ] as const)('%s · %s → %s', (kind, js, want) => {
    expect(isAppJsUrl(kind as never, js)).toBe(want);
  });
});

describe('frameUrl', () => {
  test('host=parent, 엑셀 읽기 전용만 chrome=none', () => {
    const word = new URL(frameUrl('word', JS, true), 'http://x');
    expect(word.pathname).toBe('/artcraft-frame.html');
    expect(Object.fromEntries(word.searchParams)).toEqual({ app: 'word', js: JS, host: 'parent' });
    const grid = new URL(frameUrl('grid', '/artcraft-app/gridcraft-0123456789/gridcraft-web.js', true), 'http://x');
    expect(grid.searchParams.get('chrome')).toBe('none');
    expect(new URL(frameUrl('grid', '/artcraft-app/gridcraft-0123456789/gridcraft-web.js', false), 'http://x').searchParams.has('chrome')).toBe(false);
  });
});

describe('readAppMessage — 같은 출처·우리 iframe·<앱>craft: 이름만', () => {
  const frame = {};
  const ok = { origin: 'http://h', source: frame, data: { type: 'wordcraft:ready', version: '0.18.1' } };
  test('맞으면 종류와 자료', () => {
    expect(readAppMessage(ok, 'http://h', frame, 'word')).toEqual({ kind: 'ready', data: ok.data });
    expect(readAppMessage({ ...ok, data: { type: 'wordcraft:save', bytes: new ArrayBuffer(1) } }, 'http://h', frame, 'word')?.kind).toBe('save');
  });
  test('다른 출처·다른 창·다른 앱·모르는 종류·모양 틀림은 null', () => {
    expect(readAppMessage({ ...ok, origin: 'http://evil' }, 'http://h', frame, 'word')).toBeNull();
    expect(readAppMessage({ ...ok, source: {} }, 'http://h', frame, 'word')).toBeNull();
    expect(readAppMessage(ok, 'http://h', frame, 'deck')).toBeNull();
    expect(readAppMessage({ ...ok, data: { type: 'wordcraft:open' } }, 'http://h', frame, 'word')).toBeNull();
    expect(readAppMessage({ ...ok, data: 'wordcraft:ready' }, 'http://h', frame, 'word')).toBeNull();
    expect(readAppMessage({ ...ok, data: null }, 'http://h', frame, 'word')).toBeNull();
  });
});

describe('readOnlyCommands', () => {
  test('워드 review.restrict, 엑셀 시트마다 protectSheet, PPT 없음(띠만)', () => {
    expect(readOnlyCommands('word', 'pw', 0)).toEqual([['review.restrict', { mode: 'readOnly', password: 'pw' }]]);
    expect(readOnlyCommands('grid', 'pw', 3)).toEqual([
      ['review.protectSheet', { sheet: 0, password: 'pw' }],
      ['review.protectSheet', { sheet: 1, password: 'pw' }],
      ['review.protectSheet', { sheet: 2, password: 'pw' }],
    ]);
    expect(readOnlyCommands('grid', 'pw', 0)).toHaveLength(1);
    expect(readOnlyCommands('deck', 'pw', 0)).toEqual([]);
  });
});

describe('saveDecision', () => {
  const bytes = new ArrayBuffer(4);
  test.each([
    [{ bytes, saveAs: false }, true, 'upload'],
    [{ bytes: new Uint8Array(4), saveAs: false }, true, 'upload'],
    [{ bytes, saveAs: false }, false, 'readonly'],
    [{ bytes, saveAs: true }, true, 'saveas'],
    [{ bytes: new ArrayBuffer(0) }, true, 'empty'],
    [{}, true, 'empty'],
    [{ bytes: { byteLength: 4 } }, true, 'empty'],
    [{ bytes: 'abcd' }, true, 'empty'],
  ] as const)('%o · 저장 가능 %s → %s', (data, canSave, want) => {
    expect(saveDecision(data as Record<string, unknown>, canSave)).toBe(want);
  });
});

describe('이름·주소 도우미', () => {
  test('openName: 확장자는 저장된 형식으로', () => {
    expect(openName('보고서.doc', 'docx')).toBe('보고서.docx');
    expect(openName('a.b.xlsx', 'xlsx')).toBe('a.b.xlsx');
    expect(openName('이름없음', 'pptx')).toBe('이름없음.pptx');
  });
  test('fontName: 해시를 뺀 원래 글꼴 파일 이름', () => {
    expect(fontName('/artcraft/fonts/SCDream4-0123456789.otf')).toBe('SCDream4.otf');
  });
  test('uiShowsDocument: 앱 화면 상태에 문서 이름(확장자 뺀)이 보이면 열린 것', () => {
    expect(uiShowsDocument({ title: '한글 보고서 - WordCraft' }, '한글 보고서.docx')).toBe(true);
    expect(uiShowsDocument({ title: 'Untitled' }, '한글 보고서.docx')).toBe(false);
    expect(uiShowsDocument({ t: 'a "q" b' }, 'a "q" b.docx')).toBe(true);
    expect(uiShowsDocument(null, '.docx')).toBe(false);
  });
  test('contentRedirect: 워드·PPT는 엔진 보기, 엑셀은 격자 편집기', () => {
    expect(contentRedirect('word', 'a+b')).toBe('/view.html?token=a%2Bb&ui=content');
    expect(contentRedirect('deck', 't')).toBe('/view.html?token=t&ui=content');
    expect(contentRedirect('grid', 't')).toBe('/editor-grid.html?token=t&ui=content');
  });
  test('toBase64: Buffer와 같은 값(0x8000 넘는 길이도)', () => {
    const u8 = Uint8Array.from({ length: 70_000 }, (_, i) => i % 256);
    expect(toBase64(u8)).toBe(Buffer.from(u8).toString('base64'));
  });
  test('안내 문구는 한국어', () => {
    for (const text of Object.values(NOTICES)) expect(text).toMatch(/[가-힣]/);
  });
});

describe('읽기 전용 엑셀 — 되돌리기로 시트 보호가 풀리지 않게(Task 1: edit.undo가 Protect Sheet를 되돌린다)', () => {
  test.each([
    [{ key: 'z', ctrl: true }, true],
    [{ key: 'Z', ctrl: true, shift: true }, true],
    [{ key: 'y', ctrl: true }, true],
    [{ key: 'ㅋ', code: 'KeyZ', ctrl: true }, true],
    [{ key: 'z', ctrl: false }, false],
    [{ key: 'c', ctrl: true }, false],
    [{ key: 'Backspace', ctrl: true }, false],
  ])('isUndoKey %o → %s', (e, want) => {
    expect(isUndoKey(e)).toBe(want);
  });
  test('unprotectedSheets: 보호가 풀린 시트 번호만', () => {
    expect(unprotectedSheets({ sheets: [{ protected: true }, { protected: false }, {}] })).toEqual([1, 2]);
    expect(unprotectedSheets({ sheets: [{ protected: true }] })).toEqual([]);
    expect(unprotectedSheets(null)).toEqual([]);
    expect(unprotectedSheets({})).toEqual([]);
  });
});

describe('저장 기준(If-Match) 잇기', () => {
  test('sha256Etag: 서버 engine-source ETag와 같은 꼴(따옴표 + 소문자 hex 64)', async () => {
    const u8 = Uint8Array.from([1, 2, 3, 250]);
    expect(await sha256Etag(u8)).toBe(`"${createHash('sha256').update(u8).digest('hex')}"`);
    expect(await sha256Etag(u8.buffer)).toBe(`"${createHash('sha256').update(u8).digest('hex')}"`);
  });
  test('saveFailure: 409 먼저 저장(stale)만 다시 열기, 428은 기준 없음 안내, 그 밖은 서버 문구', () => {
    expect(saveFailure(409, '문서를 연 뒤에 다른 사람이 먼저 저장했습니다. …')).toEqual({ text: NOTICES.stale, reload: true });
    expect(saveFailure(409, '편집기 화면에서 이 문서를 열고 있어 저장하지 못했습니다.')).toEqual({ text: '편집기 화면에서 이 문서를 열고 있어 저장하지 못했습니다.', reload: false });
    expect(saveFailure(428, 'x')).toEqual({ text: NOTICES.noBase, reload: true });
    expect(saveFailure(0, undefined)).toEqual({ text: '서버에 연결할 수 없습니다', reload: false });
  });
});

describe('리뷰 수정 — 겹친 저장·열기 상한·글꼴 검사·ready 한 번', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    vi.useRealTimers();
  });
  const hex = (u8: Uint8Array) => `"${createHash('sha256').update(u8).digest('hex')}"`;

  test('createSaveQueue: 겹친 저장은 순서대로 하나씩, 둘째는 첫째 바이트의 해시를 If-Match로', async () => {
    const sent: Array<{ n: number; ifMatch: string | null }> = [];
    let release!: () => void;
    const firstGate = new Promise<void>((r) => (release = r));
    const queue = createSaveQueue(async (bytes, ifMatch) => {
      const n = new Uint8Array(bytes)[0];
      sent.push({ n, ifMatch });
      if (n === 1) await firstGate;
      return true;
    });
    queue.setBase('"base0"');
    const a = new Uint8Array([1, 1]);
    const b = new Uint8Array([2, 2]);
    const p1 = queue.save(a.buffer);
    const p2 = queue.save(b.buffer);
    await new Promise((r) => setTimeout(r, 10));
    expect(sent).toEqual([{ n: 1, ifMatch: '"base0"' }]); // 첫 저장이 끝나기 전엔 둘째를 보내지 않는다
    release();
    expect(await p1).toBe(true);
    expect(await p2).toBe(true);
    expect(sent).toEqual([{ n: 1, ifMatch: '"base0"' }, { n: 2, ifMatch: hex(a) }]);
  });

  test('createSaveQueue: 실패하면 기준을 바꾸지 않고, 앞 저장이 throw해도 뒤 저장은 나간다', async () => {
    const seen: Array<string | null> = [];
    let call = 0;
    const queue = createSaveQueue(async (_bytes, ifMatch) => {
      seen.push(ifMatch);
      call += 1;
      if (call === 1) return false;
      if (call === 2) throw new Error('끊김');
      return true;
    });
    queue.setBase('"b"');
    await queue.save(new Uint8Array([1]).buffer);
    await queue.save(new Uint8Array([2]).buffer).catch(() => undefined);
    await queue.save(new Uint8Array([3]).buffer);
    expect(seen).toEqual(['"b"', '"b"', '"b"']);
  });

  test('withTimeout: 상한 안이면 값, 넘으면 그 문구로 거부', async () => {
    await expect(withTimeout(Promise.resolve(7), 50, '느림')).resolves.toBe(7);
    vi.useFakeTimers();
    const slow = withTimeout(new Promise(() => undefined), 15_000, '열기 시간 초과');
    vi.advanceTimersByTime(15_000);
    await expect(slow).rejects.toThrow('열기 시간 초과');
  });

  test('fetchBytes: 200이면 바이트, 404면 거부(오류 화면 HTML을 글꼴로 넘기지 않게)', async () => {
    vi.stubGlobal('fetch', async (url: string) => (url === '/ok' ? new Response(new Uint8Array([9, 8])) : new Response('<html>', { status: 404 })));
    expect([...(await fetchBytes('/ok'))]).toEqual([9, 8]);
    await expect(fetchBytes('/missing')).rejects.toThrow('404');
  });

  test('once: ready가 두 번 와도 한 번만', () => {
    const fn = vi.fn();
    const go = once(fn);
    go();
    go();
    expect(fn).toHaveBeenCalledTimes(1);
  });
});
