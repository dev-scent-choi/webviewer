import fs from 'node:fs/promises';
import { PDFDocument } from 'pdf-lib';
import sharp from 'sharp';
import type { Converter } from './types.js';

/**
 * pdf-lib이 직접 임베드하지 못하는 래스터 포맷(TIFF/BMP/WebP/GIF)을
 * sharp로 PNG 페이지로 풀어 PDF로 조립한다. 다페이지 TIFF(스캔 도면/팩스)도
 * 페이지별로 나눠 담는다.
 */
const RASTER_FORMATS = new Set(['tiff', 'webp', 'gif', 'bmp', 'png', 'jpeg']);
/** 여러 쪽 TIFF·움직이는 GIF의 쪽(프레임) 상한 */
const MAX_PAGES = 300;

export const rasterConverter: Converter = {
  name: 'raster-sharp',
  extensions: ['tif', 'tiff', 'bmp', 'webp', 'gif'],
  async convert(sourcePath, targetPdfPath) {
    const bytes = await fs.readFile(sourcePath);
    const metadata = await sharp(bytes).metadata();
    // 확장자만 그림이고 속은 SVG·PDF 등인 파일은 받지 않는다(sharp는 내용으로 형식을 고른다)
    if (!RASTER_FORMATS.has(String(metadata.format))) throw new Error(`지원하지 않는 그림 형식입니다: ${metadata.format}`);
    const pageCount = Math.min(MAX_PAGES, metadata.pages && metadata.pages > 0 ? metadata.pages : 1);

    const doc = await PDFDocument.create();
    for (let index = 0; index < pageCount; index += 1) {
      const png = await sharp(bytes, { page: index }).png().toBuffer();
      const image = await doc.embedPng(png);
      const page = doc.addPage([image.width, image.height]);
      page.drawImage(image, { x: 0, y: 0, width: image.width, height: image.height });
    }
    await fs.writeFile(targetPdfPath, await doc.save());
  },
};
