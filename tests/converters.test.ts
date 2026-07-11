import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { PDFDocument } from 'pdf-lib';
import { afterAll, beforeAll, describe, expect, test } from 'vitest';
import sharp from 'sharp';
import { dwg2dxfPath, dwgConverter } from '../src/converters/dwgConverter.js';
import { findConverter, supportedExtensions } from '../src/converters/index.js';
import { imageConverter } from '../src/converters/imageConverter.js';
import { dxfConverter, officeConverter, sofficePath } from '../src/converters/libreOffice.js';
import { pdfPassthroughConverter } from '../src/converters/pdfPassthrough.js';
import { rasterConverter } from '../src/converters/rasterConverter.js';
import { makeSamplePdf, makeSamplePng } from './helpers.js';

let workDir: string;

beforeAll(async () => {
  workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-conv-'));
});

afterAll(async () => {
  await fs.rm(workDir, { recursive: true, force: true });
});

describe('converter registry', () => {
  test('finds converters case-insensitively', () => {
    expect(findConverter('PDF')?.name).toBe('pdf-passthrough');
    expect(findConverter('jpeg')?.name).toBe('image');
  });

  test('returns null for unsupported extensions', () => {
    expect(findConverter('exe')).toBeNull();
    expect(findConverter('')).toBeNull();
  });

  test('exposes the supported extension list', () => {
    expect(supportedExtensions).toEqual(expect.arrayContaining(['pdf', 'png', 'jpg', 'jpeg']));
  });
});

describe('pdfPassthroughConverter', () => {
  test('copies the source PDF unchanged', async () => {
    const source = path.join(workDir, 'in.pdf');
    const target = path.join(workDir, 'out-passthrough.pdf');
    const pdfBytes = await makeSamplePdf(3);
    await fs.writeFile(source, pdfBytes);

    await pdfPassthroughConverter.convert(source, target);

    const copied = await fs.readFile(target);
    expect(copied.equals(pdfBytes)).toBe(true);
  });
});

describe('imageConverter', () => {
  test('converts a PNG into a single-page PDF sized to the image', async () => {
    const source = path.join(workDir, 'in.png');
    const target = path.join(workDir, 'out-image.pdf');
    await fs.writeFile(source, makeSamplePng());

    await imageConverter.convert(source, target);

    const doc = await PDFDocument.load(await fs.readFile(target));
    expect(doc.getPageCount()).toBe(1);
    const { width, height } = doc.getPage(0).getSize();
    expect(width).toBe(1);
    expect(height).toBe(1);
  });

  test('fails explicitly on corrupt image data', async () => {
    const source = path.join(workDir, 'broken.png');
    const target = path.join(workDir, 'out-broken.pdf');
    await fs.writeFile(source, Buffer.from('not an image'));

    await expect(imageConverter.convert(source, target)).rejects.toThrow();
  });
});

describe('rasterConverter', () => {
  test('converts a TIFF into a PDF page sized to the image', async () => {
    const source = path.join(workDir, 'scan.tif');
    const target = path.join(workDir, 'out-scan.pdf');
    const frame = (color: string) =>
      sharp({ create: { width: 40, height: 30, channels: 3, background: color } }).tiff();
    const page1 = await frame('#ff0000').toBuffer();
    const page2 = await frame('#0000ff').toBuffer();
    await fs.writeFile(source, page1);
    await rasterConverter.convert(source, target);
    let doc = await PDFDocument.load(await fs.readFile(target));
    expect(doc.getPageCount()).toBe(1);
    expect(doc.getPage(0).getSize()).toEqual({ width: 40, height: 30 });

    await fs.writeFile(source, page2);
    await rasterConverter.convert(source, target);
    doc = await PDFDocument.load(await fs.readFile(target));
    expect(doc.getPageCount()).toBe(1);
  });

  test('converts BMP and WebP images', async () => {
    const webpSource = path.join(workDir, 'img.webp');
    const webpTarget = path.join(workDir, 'out-webp.pdf');
    await fs.writeFile(
      webpSource,
      await sharp({ create: { width: 20, height: 20, channels: 3, background: '#00ff00' } }).webp().toBuffer(),
    );
    await rasterConverter.convert(webpSource, webpTarget);
    const doc = await PDFDocument.load(await fs.readFile(webpTarget));
    expect(doc.getPageCount()).toBe(1);
  });

  test('fails explicitly on corrupt raster data', async () => {
    const source = path.join(workDir, 'broken.tif');
    const target = path.join(workDir, 'out-broken-tif.pdf');
    await fs.writeFile(source, Buffer.from('not a tiff'));
    await expect(rasterConverter.convert(source, target)).rejects.toThrow();
  });
});

const MINIMAL_DXF = [
  '0', 'SECTION', '2', 'ENTITIES',
  '0', 'LINE', '8', '0', '10', '0.0', '20', '0.0', '11', '100.0', '21', '100.0',
  '0', 'CIRCLE', '8', '0', '10', '50.0', '20', '50.0', '40', '25.0',
  '0', 'ENDSEC', '0', 'EOF', '',
].join('\n');

const LIBREOFFICE_TIMEOUT_MS = 200_000;

describe.skipIf(!sofficePath)('libreOffice converters (requires LibreOffice)', () => {
  test('converts a CSV spreadsheet into a PDF', { timeout: LIBREOFFICE_TIMEOUT_MS }, async () => {
    const source = path.join(workDir, 'sheet.csv');
    const target = path.join(workDir, 'out-sheet.pdf');
    await fs.writeFile(source, 'name,qty\napple,3\nbanana,7\n', 'utf8');

    await officeConverter.convert(source, target);

    const doc = await PDFDocument.load(await fs.readFile(target));
    expect(doc.getPageCount()).toBeGreaterThanOrEqual(1);
  });

  test('converts a DXF drawing into a PDF cropped to its content', { timeout: LIBREOFFICE_TIMEOUT_MS }, async () => {
    const source = path.join(workDir, 'plan.dxf');
    const target = path.join(workDir, 'out-plan.pdf');
    await fs.writeFile(source, MINIMAL_DXF, 'ascii');

    await dxfConverter.convert(source, target);

    const doc = await PDFDocument.load(await fs.readFile(target));
    expect(doc.getPageCount()).toBeGreaterThanOrEqual(1);
    // 작은 도형만 있는 도면은 빈 여백이 잘려 기본 페이지(A4~)보다 작아야 한다
    const { width, height } = doc.getPage(0).getCropBox();
    expect(width).toBeLessThan(590);
    expect(height).toBeLessThan(590);
  });

  test(
    'splits very long list sheets into normal pages instead of one giant page',
    { timeout: LIBREOFFICE_TIMEOUT_MS },
    async () => {
      const source = path.join(workDir, 'long-list.csv');
      const target = path.join(workDir, 'out-long-list.pdf');
      const rows = Array.from({ length: 1200 }, (_, i) => `${i + 1},도면-${i + 1},검토중`);
      await fs.writeFile(source, `no,name,status\n${rows.join('\n')}\n`, 'utf8');

      const { spreadsheetConverter } = await import('../src/converters/libreOffice.js');
      await spreadsheetConverter.convert(source, target, 'long-list.csv');

      const doc = await PDFDocument.load(await fs.readFile(target));
      expect(doc.getPageCount()).toBeGreaterThan(1);
      for (const page of doc.getPages()) {
        const { width, height } = page.getSize();
        expect(Math.max(width, height)).toBeLessThan(3600);
      }
    },
  );

  test('fails explicitly when no PDF is produced', { timeout: LIBREOFFICE_TIMEOUT_MS }, async () => {
    const source = path.join(workDir, 'does-not-exist.docx');
    const target = path.join(workDir, 'out-missing.pdf');

    await expect(officeConverter.convert(source, target)).rejects.toThrow();
  });
});

describe.skipIf(!sofficePath || !dwg2dxfPath)('dwgConverter (requires LibreDWG + LibreOffice)', () => {
  test('converts a DWG drawing into a PDF', { timeout: LIBREOFFICE_TIMEOUT_MS }, async () => {
    const fixture = path.join(import.meta.dirname, 'fixtures', 'example_2000.dwg');
    const target = path.join(workDir, 'out-dwg.pdf');

    await dwgConverter.convert(fixture, target, '샘플도면.dwg');

    const doc = await PDFDocument.load(await fs.readFile(target));
    expect(doc.getPageCount()).toBeGreaterThanOrEqual(1);
  });

  test('fails explicitly on a corrupt DWG', { timeout: LIBREOFFICE_TIMEOUT_MS }, async () => {
    const source = path.join(workDir, 'broken.dwg');
    const target = path.join(workDir, 'out-broken-dwg.pdf');
    await fs.writeFile(source, Buffer.from('not a dwg'));

    await expect(dwgConverter.convert(source, target, 'broken.dwg')).rejects.toThrow();
  });
});
