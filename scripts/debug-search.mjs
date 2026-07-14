import fs from 'node:fs/promises';
import { extractPdfText, searchPdfText } from '../src/search.js';

const pages = await extractPdfText(await fs.readFile(process.argv[2]));
console.log('pages:', pages.length, 'items page1:', pages[0]?.items.length);
console.log(JSON.stringify(searchPdfText(pages, process.argv[3] ?? 'test', 10), null, 2).slice(0, 1500));
