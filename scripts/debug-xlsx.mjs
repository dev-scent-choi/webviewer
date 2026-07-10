import fs from 'node:fs/promises';
import { strFromU8, unzipSync } from 'fflate';

const [xlsxPath, sheetName, ...rowArgs] = process.argv.slice(2);
const entries = unzipSync(await fs.readFile(xlsxPath));
const read = (name) => strFromU8(entries[name]);

const shared = [...read('xl/sharedStrings.xml').matchAll(/<si>([\s\S]*?)<\/si>/g)].map((m) =>
  [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join(''),
);

const sheetXml = read(`xl/worksheets/${sheetName}.xml`);
for (const rowNum of rowArgs) {
  const row = sheetXml.match(new RegExp(`<row r="${rowNum}"[\\s\\S]*?</row>`));
  if (!row) {
    console.log(`row ${rowNum}: (없음)`);
    continue;
  }
  console.log(`--- row ${rowNum}:`);
  for (const cell of row[0].matchAll(/<c r="([A-Z]+)\d+"([^>]*)>(?:<v>([^<]*)<\/v>)?<\/c>|<c r="([A-Z]+)\d+"[^>]*\/>/g)) {
    const col = cell[1] ?? cell[4];
    const attrs = cell[2] ?? '';
    const value = cell[3];
    if (value === undefined) continue;
    const text = attrs.includes('t="s"') ? shared[Number(value)] : value;
    console.log(`  ${col} = ${text}`);
  }
}
