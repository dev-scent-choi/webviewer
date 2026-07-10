import fs from 'node:fs/promises';
import { findConverter } from '../src/converters/index.js';
import { getPageCountFromBytes } from '../src/rendering.js';

const [sourcePath, ext, targetPdfPath, sourceName] = process.argv.slice(2);
const converter = findConverter(ext);
if (!converter) {
  console.error('no converter for', ext);
  process.exit(1);
}
await converter.convert(sourcePath, targetPdfPath, sourceName);
const pages = await getPageCountFromBytes(await fs.readFile(targetPdfPath));
console.log('converted with', converter.name, '→ pages:', pages);
