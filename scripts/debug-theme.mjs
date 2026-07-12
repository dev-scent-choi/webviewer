import fs from 'node:fs/promises';
import ExcelJS from 'exceljs';

const wb = new ExcelJS.Workbook();
await wb.xlsx.load(await fs.readFile(process.argv[2]));
const themes = wb.model.themes;
console.log('type:', typeof themes, Array.isArray(themes) ? 'array' : '');
if (themes && typeof themes === 'object') {
  const keys = Object.keys(themes);
  console.log('keys:', keys);
  const first = themes[keys[0]];
  console.log('first value type:', typeof first, String(first).slice(0, 120));
}
