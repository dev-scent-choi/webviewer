import { describe, expect, test } from 'vitest';
import { editorPathFor, editorUrlFor } from '../src/routes/documents.js';
import { issueViewToken, verifyViewToken } from '../src/security.js';

describe('PDF 편집기 연결', () => {
  const t = 'tok';
  test('pdf는 pdfEditor+canDownload일 때만 편집기 주소(8000은 pdfEditor를 안 보낸다)', () => {
    expect(editorUrlFor('pdf', { effectiveCanEdit: true, canDownload: true, pdfEditor: false }, t)).toBeUndefined();
    expect(editorUrlFor('pdf', { effectiveCanEdit: true, canDownload: false, pdfEditor: true }, t)).toBeUndefined();
    expect(editorUrlFor('pdf', { effectiveCanEdit: false, canDownload: true, pdfEditor: true }, t)).toBe('/editor-pdf.html?token=tok');
    expect(editorUrlFor('pdf', { effectiveCanEdit: true, canDownload: true, pdfEditor: true }, t)).toBe('/editor-pdf.html?token=tok');
  });
  test('기존 형식은 예전과 같다', () => {
    expect(editorUrlFor('docx', { effectiveCanEdit: false, canDownload: false, pdfEditor: false }, t)).toBe('/editor.html?token=tok');
    expect(editorUrlFor('xlsx', { effectiveCanEdit: false, canDownload: false, pdfEditor: false }, t)).toBe('/editor-grid.html?token=tok');
    expect(editorUrlFor('hwp', { effectiveCanEdit: false, canDownload: true, pdfEditor: true }, t)).toBeUndefined();
    expect(editorUrlFor('hwp', { effectiveCanEdit: true, canDownload: false, pdfEditor: false }, t)).toBe('/editor-hwp.html?token=tok');
    expect(editorUrlFor('png', { effectiveCanEdit: true, canDownload: true, pdfEditor: true }, t)).toBeUndefined();
  });
  test('경로', () => {
    expect(editorPathFor('pdf')).toBe('/editor-pdf.html');
  });
});

describe('PDF 편집기 토큰 표시(pdfEditor)', () => {
  const base = { docId: '11111111-2222-3333-4444-555555555555', viewer: '향기랩 방문자' };
  test('붙인 토큰만 그대로 돌아오고, 안 붙인 토큰은 예전 모양 그대로(키 없음)', () => {
    const withFlag = verifyViewToken(issueViewToken({ ...base, perms: { canPrint: true, canDownload: true, canCopy: true, canEdit: true, pdfEditor: true } }, 60));
    expect(withFlag.perms.pdfEditor).toBe(true);
    const without = verifyViewToken(issueViewToken({ ...base, perms: { canPrint: true, canDownload: true, canCopy: true, canEdit: true } }, 60));
    expect(without.perms).toEqual({ canPrint: true, canDownload: true, canCopy: true, canEdit: true });
  });
  test('true가 아닌 값은 표시로 인정하지 않는다', () => {
    const odd = verifyViewToken(issueViewToken({ ...base, perms: { canPrint: false, canDownload: false, canCopy: false, canEdit: false, pdfEditor: 'yes' as unknown as boolean } }, 60));
    expect(odd.perms.pdfEditor).toBeUndefined();
  });
});
