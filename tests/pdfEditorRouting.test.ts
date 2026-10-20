import { describe, expect, test } from 'vitest';
import { editorPathFor, editorUrlFor } from '../src/routes/documents.js';

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
