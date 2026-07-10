/**
 * 매핑정의서 시트(sheet3)의 열 밀림 수정:
 * 데이터 행(4행~)의 P/Q/R열 내용이 헤더 기준 한 칸 왼쪽에 저장되어 있어
 * P(NOT NULL)→Q(변환 규칙), Q→R(기본값), R→S(비고)로 옮긴다.
 * P열(NOT NULL) 값은 원본에 존재하지 않으므로 빈 셀로 남긴다.
 */
import fs from 'node:fs/promises';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

const [inputPath, outputPath] = process.argv.slice(2);
if (!inputPath || !outputPath) {
  console.error('사용법: node fix-mapping-shift.mjs <원본.xlsx> <수정본.xlsx>');
  process.exit(1);
}

const SHEET = 'xl/worksheets/sheet3.xml';
const FIRST_DATA_ROW = 4;

const entries = unzipSync(await fs.readFile(inputPath));
const xml = strFromU8(entries[SHEET]);

const sharedStrings = [...strFromU8(entries['xl/sharedStrings.xml']).matchAll(/<si>([\s\S]*?)<\/si>/g)].map(
  (m) => [...m[1].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map((t) => t[1]).join(''),
);

function cellText(cell) {
  if (!cell) return '';
  const value = cell.match(/<v>([^<]*)<\/v>/)?.[1];
  if (value === undefined) return '';
  return (cell.includes('t="s"') ? (sharedStrings[Number(value)] ?? '') : value).trim();
}

let shifted = 0;
let skipped = 0;

const fixed = xml.replace(/<row [^>]*?r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g, (rowXml, rowNumStr, inner) => {
  const rowNum = Number(rowNumStr);
  if (rowNum < FIRST_DATA_ROW) return rowXml;

  const cells = new Map();
  for (const match of inner.matchAll(/<c r="([A-Z]+)\d+"(?:[^>]*\/>|[^>]*>[\s\S]*?<\/c>)/g)) {
    cells.set(match[1], match[0]);
  }
  const hasContent = ['P', 'Q', 'R'].some((col) => cellText(cells.get(col)) !== '');
  if (!hasContent) return rowXml;
  if (cellText(cells.get('S')) !== '') {
    console.warn(`row ${rowNum}: S열에 이미 값이 있어 건너뜀`);
    skipped += 1;
    return rowXml;
  }

  const rename = (cell, from, to) => cell?.replace(`r="${from}${rowNum}"`, `r="${to}${rowNum}"`);
  const emptyCell = (col, styleSource) => {
    const style = styleSource?.match(/ s="(\d+)"/)?.[1];
    return `<c r="${col}${rowNum}"${style ? ` s="${style}"` : ''}/>`;
  };

  const newCells = new Map(cells);
  newCells.set('S', rename(cells.get('R'), 'R', 'S') ?? emptyCell('S', cells.get('S')));
  newCells.set('R', rename(cells.get('Q'), 'Q', 'R') ?? emptyCell('R', cells.get('R')));
  newCells.set('Q', rename(cells.get('P'), 'P', 'Q') ?? emptyCell('Q', cells.get('Q')));
  newCells.set('P', emptyCell('P', cells.get('P')));

  const colNum = (col) => [...col].reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
  const rebuilt = [...newCells.entries()]
    .sort(([a], [b]) => colNum(a) - colNum(b))
    .map(([, cell]) => cell)
    .join('');

  shifted += 1;
  const openTag = rowXml.slice(0, rowXml.indexOf('>') + 1);
  return `${openTag}${rebuilt}</row>`;
});

entries[SHEET] = strToU8(fixed);
await fs.writeFile(outputPath, zipSync(entries));
console.log(`완료: ${shifted}개 행 이동, ${skipped}개 행 건너뜀 → ${outputPath}`);
