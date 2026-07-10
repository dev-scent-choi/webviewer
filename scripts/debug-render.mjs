import fs from 'node:fs/promises';
import { renderPage } from '../src/rendering.js';

const [pdfPath, pageArg, outPath] = process.argv.slice(2);
try {
  const png = await renderPage(pdfPath, Number(pageArg ?? 1));
  await fs.writeFile(outPath, png);
  console.log('render ok:', outPath, png.length, 'bytes');
} catch (error) {
  console.error('render failed:', error);
  process.exitCode = 1;
}
