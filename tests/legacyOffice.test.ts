// 시험 자료: Apache POI test-data(Apache License 2.0), 커밋 a4e8640da7c3bb906efc6fa11914d1b80f3cc9dc
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { describe, expect, test, vi } from 'vitest';
import { convertOfficeFile, sofficePath } from '../src/converters/libreOffice.js';
import { hasVbaMacros, isLegacyExt, LEGACY_EDIT_TARGET, sha256 } from '../src/legacyOffice.js';

const fx = (name: string) => readFileSync(path.resolve('tests', 'fixtures', 'legacy', name));

describe('옛 형식 표', () => {
  test('doc·xls·ppt만 옛 형식, 바꿀 곳은 OOXML', () => {
    expect(['doc', 'xls', 'ppt'].every(isLegacyExt)).toBe(true);
    expect(['docx', 'odt', 'hwp', ''].some(isLegacyExt)).toBe(false);
    expect(LEGACY_EDIT_TARGET).toEqual({ doc: 'docx', xls: 'xlsx', ppt: 'pptx' });
    expect(sha256(Buffer.from('a'))).toBe('ca978112ca1bbdcafac231b39a23dc4da786eff8147c4e72b9807785afee48bb');
  });
});

describe('매크로(VBA) 찾기', () => {
  test('매크로가 든 파일', () => {
    expect(hasVbaMacros(fx('SimpleMacro.doc'), 'doc')).toBe(true);
    expect(hasVbaMacros(fx('SimpleMacro.xls'), 'xls')).toBe(true);
    expect(hasVbaMacros(fx('SimpleMacro.ppt'), 'ppt')).toBe(true);
  });
  test('매크로 없는 파일·복합 파일이 아닌 것', () => {
    expect(hasVbaMacros(fx('table-merges.doc'), 'doc')).toBe(false);
    expect(hasVbaMacros(fx('chinese-provinces.xls'), 'xls')).toBe(false);
    expect(hasVbaMacros(fx('SampleShow.ppt'), 'ppt')).toBe(false);
    expect(hasVbaMacros(Buffer.from('PK\u0003\u0004 not ole'), 'doc')).toBe(false);
  });
});

describe.skipIf(!sofficePath)('LibreOffice로 바꾸기', () => {
  test('doc → docx → doc: OOXML(zip)과 OLE 머리, 글자 유지', async () => {
    const docx = await convertOfficeFile(fx('table-merges.doc'), 'doc', 'docx');
    expect(docx.subarray(0, 2).toString()).toBe('PK');
    const back = await convertOfficeFile(docx, 'docx', 'doc');
    expect(back.readUInt32LE(0)).toBe(0xe011cfd0);
    const again = await convertOfficeFile(back, 'doc', 'docx');
    const text = (b: Buffer) => strFromU8(unzipSync(b, { filter: (f) => f.name === 'word/document.xml' })['word/document.xml']!).replace(/<[^>]+>/g, '');
    expect(text(again)).toBe(text(docx));
  }, 120_000);
  test('xls·ppt도 OOXML로', async () => {
    expect((await convertOfficeFile(fx('chinese-provinces.xls'), 'xls', 'xlsx')).subarray(0, 2).toString()).toBe('PK');
    expect((await convertOfficeFile(fx('SampleShow.ppt'), 'ppt', 'pptx')).subarray(0, 2).toString()).toBe('PK');
  }, 120_000);
  test('모르는 형식은 거절', async () => {
    await expect(convertOfficeFile(Buffer.from('x'), 'doc', 'hwp')).rejects.toThrow('바꿀 수 없는 형식');
  });
});

describe('legacyDownload: 다시 바꾸기 실패면 편집본을 OOXML 이름으로', () => {
  test('fallback', async () => {
    vi.stubEnv('DOCVIEW_STORAGE', 'memory');
    vi.resetModules();
    const { store, newDocumentId } = await import('../src/storage.js');
    const { legacyDownload } = await import('../src/legacyOffice.js');
    const id = newDocumentId();
    await store.saveOriginal(id, 'docx', Buffer.from('편집본'));
    await store.saveDerived(id, 'legacy', Buffer.from('올린 원본'));
    const meta = { id, name: '보고서.doc', ext: 'docx', status: 'ready', pageCount: 1, createdAt: '', legacy: { ext: 'doc', sha256: '다름', hasMacros: false } } as const;
    const out = await legacyDownload(meta, async () => {
      throw new Error('soffice 없음');
    });
    expect(out).toEqual({ bytes: Buffer.from('편집본'), name: '보고서.docx', how: 'fallback' });
    const same = await legacyDownload({ ...meta, legacy: { ...meta.legacy, sha256: (await import('../src/legacyOffice.js')).sha256(Buffer.from('편집본')) } });
    expect(same).toEqual({ bytes: Buffer.from('올린 원본'), name: '보고서.doc', how: 'original' });
    vi.unstubAllEnvs();
  });
});
