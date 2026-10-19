/**
 * ExcelJS(4.4)는 차트를 모른다 — 차트가 든 그리기 파트가 있으면 불러오다 예외를 던지거나(실제
 * openpyxl·Excel 파일에서 확인: `drawing.anchors` undefined), 불러와도 저장할 때 차트를 버린다.
 * 그래서 ExcelJS에 넘기기 전에 시트의 그리기(<drawing>)와 그 아래 파트 전체(차트·그림·스타일 등)를
 * 떼어 두었다가(stripDrawings), ExcelJS가 쓴 결과에 시트 이름으로 찾아 그대로 되붙인다
 * (restoreDrawings). 차트·그림은 원본 바이트 그대로 보존된다. 새로 만든 차트도 같은 경로로 붙인다.
 *
 * 알려진 한계: 행/열 삽입·삭제로 칸이 밀려도 그리기 위치(앵커)와 차트 참조 범위는 옮기지 않는다.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import {
  buildChartAnchorXml,
  buildChartXml,
  chartAnchorsOf,
  parseChartXml,
  pictureAnchorsOf,
  type ChartAnchorPoint,
  type GridChart,
} from './chartModel.js';

export interface StashedSheetDrawing {
  /** 시트 이름 — ExcelJS가 저장하며 시트 파일 이름을 다시 매기므로 이름으로 다시 찾는다. */
  readonly sheetName: string;
  /** 그리기 파트 경로(zip 안 절대 경로, 예: xl/drawings/drawing1.xml) */
  readonly drawingPath: string;
}

export interface DrawingStash {
  readonly parts: Readonly<Record<string, Uint8Array>>;
  readonly sheets: readonly StashedSheetDrawing[];
  /** 떼어 낸 파트의 콘텐츠 형식(Override) — 되붙일 때 다시 등록한다. */
  readonly overrides: Readonly<Record<string, string>>;
}

const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const DRAWING_REL_TYPE = `${REL_NS}/drawing`;
const DRAWING_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml';
const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

const MEDIA_TYPES: Readonly<Record<string, string>> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  tif: 'image/tiff',
  tiff: 'image/tiff',
  svg: 'image/svg+xml',
};

/** "../drawings/drawing1.xml"을 기준 파트 위치로 풀어 zip 안 절대 경로로 */
export function resolvePartPath(basePart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = basePart.split('/').slice(0, -1);
  for (const segment of target.split('/')) {
    if (segment === '..') parts.pop();
    else if (segment !== '.' && segment !== '') parts.push(segment);
  }
  return parts.join('/');
}

function relsPathOf(part: string): string {
  const dir = part.split('/').slice(0, -1).join('/');
  const name = part.split('/').pop() ?? '';
  return `${dir ? `${dir}/` : ''}_rels/${name}.rels`;
}

interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly external: boolean;
  readonly raw: string;
}

function relationshipsOf(xml: string): Relationship[] {
  return [...xml.matchAll(/<Relationship\b[^>]*\/>/g)].map((match) => {
    const raw = match[0];
    const attr = (name: string) => new RegExp(`\\b${name}="([^"]*)"`).exec(raw)?.[1] ?? '';
    return { id: attr('Id'), type: attr('Type'), target: attr('Target'), external: attr('TargetMode') === 'External', raw };
  });
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function decodeXmlAttr(value: string): string {
  // ExcelJS는 한글 시트 이름을 &#49892;처럼 숫자 문자 참조로 쓴다
  return value
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(Number(dec)))
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

/** 시트 이름 → 시트 파트 경로(xl/workbook.xml + 관계) */
export function sheetPathsByName(entries: Record<string, Uint8Array>): Map<string, string> {
  const result = new Map<string, string>();
  const workbook = entries['xl/workbook.xml'];
  const rels = entries['xl/_rels/workbook.xml.rels'];
  if (!workbook || !rels) return result;
  const relById = new Map(relationshipsOf(strFromU8(rels)).map((rel) => [rel.id, rel]));
  for (const match of strFromU8(workbook).matchAll(/<sheet\b[^>]*\/>/g)) {
    const name = /\bname="([^"]*)"/.exec(match[0])?.[1];
    const rid = /\br:id="([^"]*)"/.exec(match[0])?.[1];
    const rel = rid ? relById.get(rid) : undefined;
    if (name !== undefined && rel) result.set(decodeXmlAttr(name), resolvePartPath('xl/workbook.xml', rel.target));
  }
  return result;
}

/** 그리기 파트와 그 관계를 따라 닿는 모든 파트(차트, 차트 스타일·색, 그림, 내장 통합 문서 …) */
function closureOf(entries: Record<string, Uint8Array>, start: string): string[] {
  const seen = new Set<string>();
  const queue = [start];
  while (queue.length > 0) {
    const part = queue.shift() as string;
    if (seen.has(part) || !entries[part]) continue;
    seen.add(part);
    const relsPath = relsPathOf(part);
    if (!entries[relsPath]) continue;
    seen.add(relsPath);
    for (const rel of relationshipsOf(strFromU8(entries[relsPath]))) {
      if (!rel.external) queue.push(resolvePartPath(part, rel.target));
    }
  }
  return [...seen];
}

export function stripDrawings(xlsx: Uint8Array): { bytes: Uint8Array; stash: DrawingStash } {
  const entries: Record<string, Uint8Array> = unzipSync(xlsx);
  const parts: Record<string, Uint8Array> = {};
  const sheets: StashedSheetDrawing[] = [];
  const overrides: Record<string, string> = {};
  const stripped = new Set<string>();
  for (const [sheetName, sheetPath] of sheetPathsByName(entries)) {
    const sheetBytes = entries[sheetPath];
    const relsPath = relsPathOf(sheetPath);
    if (!sheetBytes || !entries[relsPath]) continue;
    let sheetXml = strFromU8(sheetBytes);
    const drawingEl = /<drawing\b[^>]*\/>/.exec(sheetXml)?.[0];
    if (!drawingEl) continue;
    const rid = /\br:id="([^"]*)"/.exec(drawingEl)?.[1];
    let relsXml = strFromU8(entries[relsPath]);
    const rel = relationshipsOf(relsXml).find((r) => r.id === rid);
    if (!rel) continue;
    const drawingPath = resolvePartPath(sheetPath, rel.target);
    sheetXml = sheetXml.replace(drawingEl, '');
    relsXml = relsXml.replace(rel.raw, '');
    entries[sheetPath] = strToU8(sheetXml);
    entries[relsPath] = strToU8(relsXml);
    sheets.push({ sheetName, drawingPath });
    for (const part of closureOf(entries, drawingPath)) stripped.add(part);
  }
  // 어느 시트도 참조하지 않는 그리기 파트도 ExcelJS가 읽다가 죽을 수 있다 — 함께 떼어 둔다(되붙이지는 않는다).
  // 단 메모·머리글 그림용 VML(vmlDrawingN.vml과 그 관계)은 남긴다 — 시트 관계가 그대로 가리키고 있어서,
  // 떼면 ExcelJS가 메모를 읽다 "reading 'comments'"로 죽는다(편집기로 한 번 저장한 파일에서 재현).
  for (const name of Object.keys(entries)) {
    if (/^xl\/(drawings|charts)\//.test(name) && !/vmlDrawing[^/]*$/.test(name) && !stripped.has(name)) stripped.add(name);
  }
  for (const part of stripped) {
    parts[part] = entries[part];
    delete entries[part];
  }
  const types = entries['[Content_Types].xml'];
  if (types) {
    let xml = strFromU8(types);
    for (const part of stripped) {
      const re = new RegExp(`<Override\\b[^>]*PartName="/${escapeRegExp(part)}"[^>]*/>`);
      const found = re.exec(xml)?.[0];
      if (found) {
        overrides[part] = /\bContentType="([^"]*)"/.exec(found)?.[1] ?? '';
        xml = xml.replace(found, '');
      }
    }
    entries['[Content_Types].xml'] = strToU8(xml);
  }
  return { bytes: sheets.length > 0 || stripped.size > 0 ? zipSync(entries) : xlsx, stash: { parts, sheets, overrides } };
}

/** CT_Worksheet 순서상 <drawing>은 이 요소들보다 앞에 와야 한다 */
const AFTER_DRAWING = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

export function insertDrawingElement(sheetXml: string, rid: string): string {
  const element = `<drawing r:id="${rid}"/>`;
  let xml = sheetXml;
  if (!/xmlns:r="/.test(xml.slice(0, xml.indexOf('>', xml.indexOf('<worksheet')) + 1))) {
    xml = xml.replace(/<worksheet\b/, `<worksheet xmlns:r="${REL_NS}"`);
  }
  let at = -1;
  for (const tag of AFTER_DRAWING) {
    const idx = xml.search(new RegExp(`<${tag}\\b`));
    if (idx >= 0 && (at < 0 || idx < at)) at = idx;
  }
  if (at < 0) at = xml.lastIndexOf('</worksheet>');
  return xml.slice(0, at) + element + xml.slice(at);
}

function addRelationship(entries: Record<string, Uint8Array>, relsPath: string, type: string, target: string): string {
  const existing = entries[relsPath] ? strFromU8(entries[relsPath]) : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  const used = new Set(relationshipsOf(existing).map((rel) => rel.id));
  let n = 1;
  while (used.has(`rIdDv${n}`)) n += 1;
  const id = `rIdDv${n}`;
  entries[relsPath] = strToU8(existing.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`));
  return id;
}

/** zip 안 경로 a에서 b를 가리키는 상대 경로 */
function relativeTarget(fromPart: string, toPart: string): string {
  const from = fromPart.split('/').slice(0, -1);
  const to = toPart.split('/');
  let common = 0;
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common += 1;
  return [...from.slice(common).map(() => '..'), ...to.slice(common)].join('/');
}

export function restoreDrawings(xlsx: Uint8Array, stash: DrawingStash): Uint8Array {
  if (stash.sheets.length === 0) return xlsx;
  const entries: Record<string, Uint8Array> = unzipSync(xlsx);
  const byName = sheetPathsByName(entries);
  let typesXml = entries['[Content_Types].xml'] ? strFromU8(entries['[Content_Types].xml']) : '';
  const register = (part: string) => {
    const contentType =
      stash.overrides[part] ||
      (part.startsWith('xl/drawings/') && part.endsWith('.xml') ? DRAWING_CONTENT_TYPE : part.startsWith('xl/charts/chart') && part.endsWith('.xml') ? CHART_CONTENT_TYPE : '');
    if (contentType && !typesXml.includes(`PartName="/${part}"`)) {
      typesXml = typesXml.replace('</Types>', `<Override PartName="/${part}" ContentType="${contentType}"/></Types>`);
    }
    const ext = part.split('.').pop()?.toLowerCase() ?? '';
    if (!contentType && MEDIA_TYPES[ext] && !new RegExp(`<Default\\b[^>]*Extension="${ext}"`, 'i').test(typesXml)) {
      typesXml = typesXml.replace('</Types>', `<Default Extension="${ext}" ContentType="${MEDIA_TYPES[ext]}"/></Types>`);
    }
  };
  // ExcelJS가 같은 이름의 파트를 새로 썼다면(그림을 직접 넣은 경우 등) 덮어쓰지 않도록 원본 쪽을 다른 이름으로
  const rename = new Map<string, string>();
  for (const part of Object.keys(stash.parts)) {
    if (!entries[part] || part.includes('/_rels/')) continue;
    const dot = part.lastIndexOf('.');
    rename.set(part, `${part.slice(0, dot)}Dv${part.slice(dot)}`);
  }
  for (const [part, bytes] of Object.entries(stash.parts)) {
    let target = part;
    let data: Uint8Array = bytes;
    if (part.includes('/_rels/')) {
      const owner = part.replace('/_rels/', '/').replace(/\.rels$/, '');
      const renamedOwner = rename.get(owner);
      if (renamedOwner) target = relsPathOf(renamedOwner);
      // 관계 대상이 이름을 바꿨으면 그 상대 경로도 바꾼다
      if (rename.size > 0) {
        let xml = strFromU8(bytes);
        for (const rel of relationshipsOf(xml)) {
          if (rel.external) continue;
          const resolved = resolvePartPath(owner, rel.target);
          const renamed = rename.get(resolved);
          if (renamed) xml = xml.replace(rel.raw, rel.raw.replace(`Target="${rel.target}"`, `Target="${relativeTarget(renamedOwner ?? owner, renamed)}"`));
        }
        data = strToU8(xml);
      }
    } else {
      target = rename.get(part) ?? part;
    }
    entries[target] = data;
    if (!target.includes('/_rels/')) register(target);
  }
  for (const { sheetName, drawingPath } of stash.sheets) {
    const sheetPath = byName.get(sheetName);
    if (!sheetPath || !entries[sheetPath]) continue;
    const drawing = rename.get(drawingPath) ?? drawingPath;
    let sheetXml = strFromU8(entries[sheetPath]);
    if (/<drawing\b/.test(sheetXml)) continue; // ExcelJS가 이미 그리기를 달았다 — 겹쳐 달면 파일이 깨진다
    const rid = addRelationship(entries, relsPathOf(sheetPath), DRAWING_REL_TYPE, relativeTarget(sheetPath, drawing));
    sheetXml = insertDrawingElement(sheetXml, rid);
    entries[sheetPath] = strToU8(sheetXml);
  }
  if (typesXml) entries['[Content_Types].xml'] = strToU8(typesXml);
  return zipSync(entries);
}

// ---------------------------------------------------------------------------
// 차트 읽기(편집기 표시용)·새 차트 붙이기(저장)
// ---------------------------------------------------------------------------

/** 파일의 시트별 차트 — 편집기가 같은 자리에 SVG로 그린다(파일 쪽 차트는 저장 때 원본 그대로 보존). */
export function readSheetCharts(xlsx: Uint8Array): GridChart[] {
  const entries: Record<string, Uint8Array> = unzipSync(xlsx);
  const charts: GridChart[] = [];
  for (const [sheetName, sheetPath] of sheetPathsByName(entries)) {
    const sheetBytes = entries[sheetPath];
    const sheetRels = entries[relsPathOf(sheetPath)];
    if (!sheetBytes || !sheetRels) continue;
    const rid = /<drawing\b[^>]*r:id="([^"]+)"/.exec(strFromU8(sheetBytes))?.[1];
    const rel = relationshipsOf(strFromU8(sheetRels)).find((r) => r.id === rid);
    if (!rel) continue;
    const drawingPath = resolvePartPath(sheetPath, rel.target);
    const drawing = entries[drawingPath];
    const drawingRels = entries[relsPathOf(drawingPath)];
    if (!drawing || !drawingRels) continue;
    const relById = new Map(relationshipsOf(strFromU8(drawingRels)).map((r) => [r.id, r]));
    for (const anchor of chartAnchorsOf(strFromU8(drawing))) {
      const chartRel = relById.get(anchor.rid);
      const chartBytes = chartRel ? entries[resolvePartPath(drawingPath, chartRel.target)] : undefined;
      const parsed = chartBytes ? parseChartXml(strFromU8(chartBytes)) : null;
      if (!parsed) continue;
      charts.push({ id: `file:${drawingPath}:${anchor.rid}`, sheetName, from: anchor.from, to: anchor.to, size: anchor.size, source: 'file', ...parsed });
    }
  }
  return charts;
}

/** 시트 위에 놓인 그림(로고·사진) — 보기·편집 화면이 칸 위치에 맞춰 그대로 얹는다(읽기 전용). */
export interface GridImage {
  readonly id: string;
  readonly sheetName: string;
  readonly name: string;
  readonly from: ChartAnchorPoint;
  readonly to: ChartAnchorPoint | null;
  readonly size: { readonly width: number; readonly height: number } | null;
  /** data: URL */
  readonly src: string;
}

const IMAGE_MIME: Readonly<Record<string, string>> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
};
/** 그림을 응답에 data URL로 싣는 총량 상한 — 사진이 많은 큰 파일이 응답을 키우지 않게 */
const MAX_IMAGE_BYTES_TOTAL = 12 * 1024 * 1024;

/** 파일에 있는 그림들(시트 이름·위치·내용). 브라우저가 못 그리는 형식(EMF/WMF 등)은 뺀다. */
export function readSheetImages(xlsx: Uint8Array): GridImage[] {
  const entries: Record<string, Uint8Array> = unzipSync(xlsx);
  const images: GridImage[] = [];
  let total = 0;
  for (const [sheetName, sheetPath] of sheetPathsByName(entries)) {
    const sheetBytes = entries[sheetPath];
    const sheetRels = entries[relsPathOf(sheetPath)];
    if (!sheetBytes || !sheetRels) continue;
    const rid = /<drawing\b[^>]*r:id="([^"]+)"/.exec(strFromU8(sheetBytes))?.[1];
    const rel = relationshipsOf(strFromU8(sheetRels)).find((r) => r.id === rid);
    if (!rel) continue;
    const drawingPath = resolvePartPath(sheetPath, rel.target);
    const drawing = entries[drawingPath];
    const drawingRels = entries[relsPathOf(drawingPath)];
    if (!drawing || !drawingRels) continue;
    const relById = new Map(relationshipsOf(strFromU8(drawingRels)).map((r) => [r.id, r]));
    for (const anchor of pictureAnchorsOf(strFromU8(drawing))) {
      const imageRel = relById.get(anchor.rid);
      if (!imageRel) continue;
      const imagePath = resolvePartPath(drawingPath, imageRel.target);
      const bytes = entries[imagePath];
      const mime = IMAGE_MIME[imagePath.split('.').pop()?.toLowerCase() ?? ''];
      if (!bytes || !mime || total + bytes.length > MAX_IMAGE_BYTES_TOTAL) continue;
      total += bytes.length;
      images.push({
        id: `img:${drawingPath}:${anchor.rid}:${images.length}`,
        sheetName,
        name: anchor.name,
        from: anchor.from,
        to: anchor.to,
        size: anchor.size,
        src: `data:${mime};base64,${Buffer.from(bytes).toString('base64')}`,
      });
    }
  }
  return images;
}

/** 새 앵커는 xdr:·a:·r: 접두사로 쓰므로 그리기 루트에 세 네임스페이스가 모두 선언돼 있게 한다. */
function withRootNamespaces(drawingXml: string): string {
  const rootMatch = /<((?:[\w.-]+:)?wsDr)\b[^>]*>/.exec(drawingXml);
  if (!rootMatch) return drawingXml;
  let root = rootMatch[0];
  const required: ReadonlyArray<[string, string]> = [
    ['xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'],
    ['a', 'http://schemas.openxmlformats.org/drawingml/2006/main'],
    ['r', REL_NS],
  ];
  for (const [prefix, uri] of required) {
    if (!new RegExp(`\\sxmlns:${prefix}=`).test(root)) root = root.replace(/\/?>$/, (end) => ` xmlns:${prefix}="${uri}"${end}`);
  }
  return drawingXml.replace(rootMatch[0], root);
}

export interface ChartCache {
  readonly categories: readonly string[];
  readonly series: ReadonlyArray<{ name: string; values: readonly number[] }>;
}

/**
 * 편집기에서 넣은 차트를 저장 파일에 붙인다 — 시트에 이미 그리기가 있으면 그 그리기에 앵커를 더하고,
 * 없으면 그리기 파트를 새로 만든다. 차트 파트·관계·콘텐츠 형식을 모두 등록한다.
 */
export function addNewCharts(xlsx: Uint8Array, charts: ReadonlyArray<{ chart: GridChart; cache: ChartCache }>): Uint8Array {
  if (charts.length === 0) return xlsx;
  const entries: Record<string, Uint8Array> = unzipSync(xlsx);
  const byName = sheetPathsByName(entries);
  let typesXml = entries['[Content_Types].xml'] ? strFromU8(entries['[Content_Types].xml']) : '';
  const addOverride = (part: string, type: string) => {
    if (!typesXml.includes(`PartName="/${part}"`)) typesXml = typesXml.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
  };
  const freeName = (dir: string, base: string) => {
    let n = 1;
    while (entries[`${dir}/${base}${n}.xml`]) n += 1;
    return `${dir}/${base}${n}.xml`;
  };
  let shapeId = 5000;
  for (const { chart, cache } of charts) {
    const sheetPath = byName.get(chart.sheetName);
    if (!sheetPath || !entries[sheetPath]) continue;
    let sheetXml = strFromU8(entries[sheetPath]);
    const sheetRelsPath = relsPathOf(sheetPath);
    const sheetRid = /<drawing\b[^>]*r:id="([^"]+)"/.exec(sheetXml)?.[1];
    const existing = sheetRid && entries[sheetRelsPath] ? relationshipsOf(strFromU8(entries[sheetRelsPath])).find((r) => r.id === sheetRid) : undefined;
    let drawingPath: string;
    if (existing) {
      drawingPath = resolvePartPath(sheetPath, existing.target);
    } else {
      drawingPath = freeName('xl/drawings', 'drawing');
      entries[drawingPath] = strToU8(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></xdr:wsDr>',
      );
      addOverride(drawingPath, DRAWING_CONTENT_TYPE);
      const rid = addRelationship(entries, sheetRelsPath, DRAWING_REL_TYPE, relativeTarget(sheetPath, drawingPath));
      sheetXml = insertDrawingElement(sheetXml, rid);
      entries[sheetPath] = strToU8(sheetXml);
    }
    const chartPath = freeName('xl/charts', 'chart');
    entries[chartPath] = strToU8(buildChartXml(chart, cache));
    addOverride(chartPath, CHART_CONTENT_TYPE);
    const chartRid = addRelationship(entries, relsPathOf(drawingPath), `${REL_NS}/chart`, relativeTarget(drawingPath, chartPath));
    let drawingXml = strFromU8(entries[drawingPath]);
    drawingXml = withRootNamespaces(drawingXml);
    shapeId += 1;
    // 루트가 접두사 없이(기본 네임스페이스) 쓰인 파일(openpyxl 등)도 있다 — 닫는 태그를 접두사와 무관하게 찾는다
    drawingXml = drawingXml.replace(/(<\/(?:[\w.-]+:)?wsDr>)\s*$/, `${buildChartAnchorXml(chart, chartRid, shapeId)}$1`);
    entries[drawingPath] = strToU8(drawingXml);
  }
  if (typesXml) entries['[Content_Types].xml'] = strToU8(typesXml);
  return zipSync(entries);
}
