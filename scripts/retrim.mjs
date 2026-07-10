import { trimPdfMargins } from '../src/converters/pdfTrim.js';

for (const pdfPath of process.argv.slice(2)) {
  await trimPdfMargins(pdfPath);
  console.log('trimmed:', pdfPath);
}
