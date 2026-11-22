import { describe, expect, test } from 'vitest';
import { ENGINE_REASONS, artcraftInfo, engineFor, type Engine } from '../src/artcraftEngine.js';

const M = { word: '/artcraft/artcraft_word-0123456789.wasm', deck: '/artcraft/artcraft_deck-0123456789.wasm', fonts: ['/artcraft/fonts/SCDream4-0123456789.otf', '/artcraft/fonts/SCDream6-0123456789.otf'] };

describe('engineFor — 형식 × 권한 × 요청 × 설치', () => {
  const R = ENGINE_REASONS;
  const table: Array<[string, boolean, Engine | undefined, boolean, Engine, string | undefined]> = [
    ['docx', true, 'artcraft', true, 'artcraft', undefined],
    ['pptx', true, 'artcraft', true, 'artcraft', undefined],
    ['doc', true, 'artcraft', true, 'artcraft', undefined],
    ['docx', false, 'artcraft', true, 'libreoffice', R.noDownload],
    ['ppt', true, 'artcraft', true, 'libreoffice', R.unsupported],
    ['xlsx', true, 'artcraft', true, 'libreoffice', R.unsupported],
    ['xls', true, 'artcraft', true, 'libreoffice', R.unsupported],
    ['hwp', true, 'artcraft', true, 'libreoffice', R.unsupported],
    ['pdf', false, 'artcraft', true, 'libreoffice', R.unsupported],
    ['docx', true, 'artcraft', false, 'libreoffice', R.notInstalled],
    ['docx', true, 'libreoffice', true, 'libreoffice', undefined],
    ['docx', true, undefined, true, 'libreoffice', undefined],
    ['xlsx', false, undefined, false, 'libreoffice', undefined],
    // 설치 안 됨이 먼저 — manifest가 없으면 형식과 무관하게 '설치되지 않았습니다'
    ['xlsx', true, 'artcraft', false, 'libreoffice', R.notInstalled],
  ];
  test.each(table)('%s · 내려받기 %s · 요청 %s · 설치 %s → %s', (ext, canDownload, requested, installed, engine, reason) => {
    expect(engineFor(ext, { canDownload }, requested, installed)).toEqual(reason ? { engine, reason } : { engine });
  });

  test('이유 문구는 설계 그대로', () => {
    expect(ENGINE_REASONS).toEqual({
      unsupported: '이 형식은 브라우저 엔진이 지원하지 않습니다',
      notInstalled: '브라우저 엔진이 설치되지 않았습니다',
      noDownload: '내려받기 권한이 없어 서버 변환으로 엽니다',
      noEditor: '이 형식은 읽기 전용 편집기 화면이 없습니다',
    });
  });

  test('constructor 같은 이름도 형식으로 보지 않는다', () => {
    expect(engineFor('constructor', { canDownload: true }, 'artcraft', true).engine).toBe('libreoffice');
  });
  test('앱 화면(appOk)이 되면 엔진이 안 그리는 형식도 artcraft — 내려받기 권한은 그대로 필요', () => {
    expect(engineFor('xlsx', { canDownload: true }, 'artcraft', true, true)).toEqual({ engine: 'artcraft' });
    expect(engineFor('xlsx', { canDownload: true }, 'artcraft', false, true)).toEqual({ engine: 'artcraft' });
    expect(engineFor('xlsx', { canDownload: false }, 'artcraft', true, true)).toEqual({ engine: 'libreoffice', reason: ENGINE_REASONS.noDownload });
    expect(engineFor('xlsx', { canDownload: true }, 'artcraft', true, false)).toEqual({ engine: 'libreoffice', reason: ENGINE_REASONS.unsupported });
    expect(engineFor('docx', { canDownload: true }, 'libreoffice', true, true)).toEqual({ engine: 'libreoffice' });
  });
});

describe('artcraftInfo — info의 artcraft 블록', () => {
  test('docx → word, doc → word 실험, pptx → deck, 그 밖 undefined', () => {
    expect(artcraftInfo('docx', 'a.b+c', M)).toEqual({
      kind: 'word',
      ext: 'docx',
      experimental: false,
      originalUrl: '/api/view/engine-source?token=a.b%2Bc',
      assets: { wasm: M.word, fonts: M.fonts },
    });
    expect(artcraftInfo('doc', 't', M)).toMatchObject({ kind: 'word', experimental: true });
    expect(artcraftInfo('pptx', 't', M)).toMatchObject({ kind: 'deck', experimental: false, assets: { wasm: M.deck } });
    expect(artcraftInfo('xlsx', 't', M)).toBeUndefined();
  });
});

describe('engineFor — 읽기 전용 편집기(editor, 설계 readonly-editor-look §3.2)', () => {
  const table: Array<[string, boolean, Engine, string | undefined]> = [
    ['docx', true, 'editor', undefined],
    ['xlsx', true, 'editor', undefined],
    ['pptx', true, 'editor', undefined],
    ['docx', false, 'libreoffice', ENGINE_REASONS.noDownload],
    ['xlsx', false, 'libreoffice', ENGINE_REASONS.noDownload],
    ['pptx', false, 'libreoffice', ENGINE_REASONS.noDownload],
    // 편집기가 없는 형식은 서버 변환 + 이유(내려받기 권한과 무관)
    ['pdf', true, 'libreoffice', ENGINE_REASONS.noEditor],
    ['pdf', false, 'libreoffice', ENGINE_REASONS.noEditor],
    ['hwp', true, 'libreoffice', ENGINE_REASONS.noEditor],
    ['doc', true, 'libreoffice', ENGINE_REASONS.noEditor],
    ['constructor', true, 'libreoffice', ENGINE_REASONS.noEditor],
  ];
  test.each(table)('%s · 내려받기 %s → %s', (ext, canDownload, engine, reason) => {
    // 아트크래프트 설치 여부(manifest)와 무관하다
    for (const installed of [true, false]) {
      expect(engineFor(ext, { canDownload }, 'editor', installed)).toEqual(reason ? { engine, reason } : { engine });
    }
  });
});
