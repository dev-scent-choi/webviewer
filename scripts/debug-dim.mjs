import fs from 'node:fs/promises';
import { strFromU8, unzipSync } from 'fflate';

const entries = unzipSync(await fs.readFile(process.argv[2]));
for (const [name, bytes] of Object.entries(entries)) {
  if (!/^xl\/worksheets\/[^/]+\.xml$/.test(name)) continue;
  const match = strFromU8(bytes).match(/<dimension ref="([^"]+)"/);
  console.log(name, match ? match[1] : '(no dimension)', `${Math.round(bytes.length / 1024)}KB`);
}
