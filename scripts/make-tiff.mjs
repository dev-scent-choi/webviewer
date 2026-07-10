import fs from 'node:fs/promises';
import sharp from 'sharp';

const [outPath] = process.argv.slice(2);
const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="800" height="600">
  <rect width="800" height="600" fill="#f4f6fa"/>
  <rect x="60" y="60" width="680" height="480" fill="none" stroke="#2f5d9e" stroke-width="3"/>
  <line x1="60" y1="300" x2="740" y2="300" stroke="#2f5d9e" stroke-width="1.5"/>
  <circle cx="400" cy="300" r="120" fill="none" stroke="#b3402a" stroke-width="2"/>
  <text x="400" y="580" font-size="28" text-anchor="middle" fill="#1c1b1a">DOCVIEW TIFF TEST DRAWING</text>
</svg>`;
await fs.writeFile(outPath, await sharp(Buffer.from(svg)).tiff().toBuffer());
console.log('tiff written:', outPath);
