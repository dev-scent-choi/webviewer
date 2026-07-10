import fs from 'node:fs/promises';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';

const doc = await PDFDocument.create();
const font = await doc.embedFont(StandardFonts.HelveticaBold);
for (let i = 1; i <= 3; i += 1) {
  const page = doc.addPage([595, 842]);
  page.drawText(`DocView Sample - Page ${i}`, {
    x: 60, y: 760, size: 28, font, color: rgb(0.18, 0.36, 0.62),
  });
  page.drawRectangle({ x: 60, y: 80, width: 475, height: 640, borderColor: rgb(0.8, 0.78, 0.72), borderWidth: 1 });
  page.drawText(`Content block ${i}`, { x: 80, y: 680, size: 14, font, color: rgb(0.3, 0.3, 0.3) });
}
await fs.writeFile(process.argv[2], await doc.save());
console.log('written:', process.argv[2]);
