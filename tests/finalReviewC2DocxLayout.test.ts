import fs from 'node:fs';
import crypto from 'node:crypto';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { docxViewLayoutFixes } from '../src/converters/libreOffice.js';

/**
 * FINAL-REVIEW C2 — docxViewLayoutFixes의 sectPr·anchor 정규식이 안 닫힌 태그에서 제곱으로 느려지던 문제(업로드·저장마다 동기).
 * 1MB 꼬리가 100ms 안이어야 하고, 실제 문서(말뭉치 10·검토 문서 10)에서는 고치기 전과 결과가 바이트 그대로여야 한다.
 */

const sect = (w: number) => `<w:sectPr><w:pgSz w:w="${w}" w:h="16838"/><w:pgMar w:top="1440" w:right="1000" w:bottom="1440" w:left="1000"/></w:sectPr>`;

describe('C2: 안 닫힌 태그에서도 선형', () => {
  test('닫힌 sectPr 400개 + 안 닫힌 <w:sectPr 로 시작하는 1MB 꼬리가 100ms 안', () => {
    const xml = `<w:body>${Array.from({ length: 400 }, () => `<w:p><w:pPr>${sect(11906)}</w:pPr></w:p>`).join('')}<w:sectPr ${'x'.repeat(1_000_000)}`;
    const t0 = performance.now();
    expect(docxViewLayoutFixes(xml)).toBe(xml);
    expect(performance.now() - t0).toBeLessThan(100);
  });

  test('안 닫힌 <wp:anchor 가 1MB 동안 이어져도 100ms 안', () => {
    const xml = `<w:body>${'<wp:anchor distT="0">'.repeat(50_000)}${sect(11906)}</w:body>`;
    const t0 = performance.now();
    expect(docxViewLayoutFixes(xml)).toBe(xml);
    expect(performance.now() - t0).toBeLessThan(100);
  });
});

describe('C2: 실제 문서 결과는 그대로(고치기 전 구현의 결과 sha256 앞 16자)', () => {
  test.each(GOLDEN)('%s', (file, hash) => {
    const xml = strFromU8(unzipSync(new Uint8Array(fs.readFileSync(file)))['word/document.xml']);
    expect(crypto.createHash('sha256').update(docxViewLayoutFixes(xml)).digest('hex').slice(0, 16)).toBe(hash);
  });
});

const GOLDEN: Array<[string, string]> = [
  ['audit/office/corpus/docx/a1-korean-report.docx', '1755605c4f4ec358'],
  ['audit/office/corpus/docx/a2-korean-layout.docx', '1604f11dacf69050'],
  ['audit/office/corpus/docx/d01-styles-fonts.docx', '4fc366938955cc0a'],
  ['audit/office/corpus/docx/d02-numbering.docx', '3721e9f660dd17e0'],
  ['audit/office/corpus/docx/d03-tables.docx', '495dbf9d9ae51840'],
  ['audit/office/corpus/docx/d04-images.docx', 'e9dc6668f34e5bba'],
  ['audit/office/corpus/docx/d05-header-footer-sections.docx', '9665a432d17a18f8'],
  ['audit/office/corpus/docx/d06-columns-footnotes.docx', '3e265a6787e2d6ea'],
  ['audit/office/corpus/docx/d07-tabs-textbox-breaks.docx', 'f276dc7f60d9b65e'],
  ['audit/office/corpus/docx/d08-korean-official-letter.docx', 'da1fe78b4246aa9a'],
  ['tests/fixtures/docx-review/A-same-sections.docx', '12be267550dd256b'],
  ['tests/fixtures/docx-review/B-tracked.docx', 'f04b9f9b13cc0c7e'],
  ['tests/fixtures/docx-review/C-num-shared-abstract.docx', '39f86b85a033568e'],
  ['tests/fixtures/docx-review/D-merged-nested.docx', '475621620b26a8b1'],
  ['tests/fixtures/docx-review/E-css.docx', '123291291403557d'],
  ['tests/fixtures/docx-review/F-big.docx', '7cf7ce3428c4bd44'],
  ['tests/fixtures/docx-review/G-big400.docx', 'e6466276a6c73caf'],
  ['tests/fixtures/docx-review/H-emf-chart.docx', '1b82520e9a7e7ae0'],
  ['tests/fixtures/docx-review/I-num-cleanup.docx', '3721e9f660dd17e0'],
  ['tests/fixtures/docx-review/J-200-drawings.docx', 'df2daee906c039b1']
];
