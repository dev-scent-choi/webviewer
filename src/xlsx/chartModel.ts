/**
 * 시트 차트 모델 — 서버(파일의 차트 읽기·새 차트 XML 만들기)와 편집기(SVG로 그리기)가 함께 쓴다.
 * 이 파일은 브라우저 번들에도 들어가므로 Node 전용 모듈을 import하지 않는다(XML은 정규식으로 읽는다).
 */

export type ChartKind = 'column' | 'bar' | 'line' | 'pie';

/** 0부터 세는 칸 범위(양 끝 포함). sheetName이 없으면 차트가 놓인 시트. */
export interface CellRangeRef {
  readonly sheetName?: string;
  readonly r1: number;
  readonly c1: number;
  readonly r2: number;
  readonly c2: number;
}

export interface ChartAnchorPoint {
  readonly row: number;
  readonly col: number;
  /** 칸 왼쪽 위에서 떨어진 거리(px) */
  readonly dx: number;
  readonly dy: number;
}

export interface ChartSeries {
  readonly name: string;
  readonly nameRef: CellRangeRef | null;
  readonly values: { readonly ref: CellRangeRef | null; readonly cache: readonly number[] };
}

export interface GridChart {
  readonly id: string;
  /** 차트가 놓인 시트 이름 */
  readonly sheetName: string;
  readonly kind: ChartKind;
  readonly title: string;
  readonly from: ChartAnchorPoint;
  /** 너비·높이(px) — 파일의 twoCellAnchor는 끝 칸으로 정해지므로 편집기가 열 폭으로 계산한다 */
  readonly size: { readonly width: number; readonly height: number } | null;
  readonly to: ChartAnchorPoint | null;
  readonly categories: { readonly ref: CellRangeRef | null; readonly cache: readonly string[] };
  readonly series: readonly ChartSeries[];
  /** file: 원본 파일에 있던 차트(저장 시 원본 그대로 보존), new: 편집기에서 넣은 차트 */
  readonly source: 'file' | 'new';
}

export const GRID_CHARTS_MAP = 'gridCharts';
export const EMU_PER_PX = 9525;

export function columnLetters(col: number): string {
  let n = col + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function columnIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** "'실적'!$B$2:$B$4" / "Sheet1!B2" → 범위 */
export function parseRangeFormula(formula: string): CellRangeRef | null {
  const text = formula.trim();
  const bang = text.lastIndexOf('!');
  let sheetName: string | undefined;
  let refs = text;
  if (bang >= 0) {
    sheetName = text.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
    refs = text.slice(bang + 1);
  }
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(refs);
  if (!m) return null;
  const c1 = columnIndex(m[1]);
  const r1 = Number(m[2]) - 1;
  const c2 = m[3] ? columnIndex(m[3]) : c1;
  const r2 = m[4] ? Number(m[4]) - 1 : r1;
  return { ...(sheetName !== undefined ? { sheetName } : {}), r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

export function rangeFormula(sheetName: string, ref: CellRangeRef): string {
  const sheet = `'${sheetName.replace(/'/g, "''")}'`;
  const a = `$${columnLetters(ref.c1)}$${ref.r1 + 1}`;
  const b = `$${columnLetters(ref.c2)}$${ref.r2 + 1}`;
  return `${sheet}!${a}${ref.r1 === ref.r2 && ref.c1 === ref.c2 ? '' : `:${b}`}`;
}

/** 범위의 칸 좌표를 행 우선으로 */
export function cellsOf(ref: CellRangeRef): Array<[number, number]> {
  const out: Array<[number, number]> = [];
  for (let r = ref.r1; r <= ref.r2; r += 1) for (let c = ref.c1; c <= ref.c2; c += 1) out.push([r, c]);
  return out;
}

/**
 * 선택 범위로 새 차트 계열을 만든다(Excel과 같은 규칙): 첫 행에 숫자가 아닌 칸이 있으면 머리글(계열
 * 이름), 첫 열이 숫자가 아니면 항목 이름. 계열은 열 단위.
 */
export function seriesFromRange(
  range: CellRangeRef,
  valueAt: (row: number, col: number) => string,
): { categories: GridChart['categories']; series: ChartSeries[] } {
  const isNum = (text: string) => text.trim() !== '' && Number.isFinite(Number(text));
  let headerRow = false;
  for (let c = range.c1; c <= range.c2; c += 1) if (valueAt(range.r1, c).trim() !== '' && !isNum(valueAt(range.r1, c))) headerRow = true;
  let labelCol = false;
  for (let r = range.r1 + (headerRow ? 1 : 0); r <= range.r2; r += 1) if (valueAt(r, range.c1).trim() !== '' && !isNum(valueAt(r, range.c1))) labelCol = true;
  if (range.c1 === range.c2) labelCol = false;
  const dataR1 = range.r1 + (headerRow ? 1 : 0);
  const dataC1 = range.c1 + (labelCol ? 1 : 0);
  const categories = labelCol
    ? { ref: { r1: dataR1, c1: range.c1, r2: range.r2, c2: range.c1 }, cache: rows(dataR1, range.r2).map((r) => valueAt(r, range.c1)) }
    : { ref: null, cache: rows(dataR1, range.r2).map((r) => String(r - dataR1 + 1)) };
  const series: ChartSeries[] = [];
  for (let c = dataC1; c <= range.c2; c += 1) {
    const name = headerRow ? valueAt(range.r1, c) : `계열${c - dataC1 + 1}`;
    series.push({
      name,
      nameRef: headerRow ? { r1: range.r1, c1: c, r2: range.r1, c2: c } : null,
      values: { ref: { r1: dataR1, c1: c, r2: range.r2, c2: c }, cache: rows(dataR1, range.r2).map((r) => Number(valueAt(r, c)) || 0) },
    });
  }
  return { categories, series };
}

function rows(from: number, to: number): number[] {
  const out: number[] = [];
  for (let r = from; r <= to; r += 1) out.push(r);
  return out;
}

// ---------------------------------------------------------------------------
// 파일의 차트 읽기(drawing + chart XML)
// ---------------------------------------------------------------------------

/** 요소 이름의 네임스페이스 접두사를 떼어 낸다(<c:ser>·<ser> 둘 다 <ser>로) — Excel은 접두사를 쓰고
 * openpyxl 등은 기본 네임스페이스로 쓴다. 속성(r:id 등)은 그대로 둔다. */
function unprefixed(xml: string): string {
  return xml.replace(/<(\/?)[A-Za-z][\w.-]*:/g, '<$1');
}

function textOf(xml: string): string {
  return [...xml.matchAll(/<t(?:\s[^>]*)?>([^<]*)<\/t>/g)].map((m) => decodeXml(m[1])).join('');
}

function decodeXml(value: string): string {
  return value
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec: string) => String.fromCodePoint(Number(dec)))
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function cachePoints(xml: string): string[] {
  const points: string[] = [];
  for (const m of xml.matchAll(/<pt\b[^>]*idx="(\d+)"[^>]*>\s*<v>([^<]*)<\/v>/g)) points[Number(m[1])] = decodeXml(m[2]);
  return [...points].map((p) => p ?? '');
}

function refOf(xml: string | undefined): CellRangeRef | null {
  const f = xml ? /<f>([^<]*)<\/f>/.exec(xml)?.[1] : undefined;
  return f ? parseRangeFormula(decodeXml(f)) : null;
}

/** chart XML 하나 → 종류·제목·계열(지원하지 않는 종류면 null) */
export function parseChartXml(source: string): Pick<GridChart, 'kind' | 'title' | 'categories' | 'series'> | null {
  const xml = unprefixed(source);
  const plot = /<(barChart|bar3DChart|lineChart|line3DChart|areaChart|pieChart|pie3DChart|doughnutChart)>([\s\S]*?)<\/\1>/.exec(xml);
  if (!plot) return null;
  const type = plot[1];
  const body = plot[2];
  const kind: ChartKind = type.startsWith('bar')
    ? /<barDir val="bar"\/>/.test(body)
      ? 'bar'
      : 'column'
    : type.startsWith('pie') || type === 'doughnutChart'
      ? 'pie'
      : 'line';
  const titleXml = /<title>([\s\S]*?)<\/title>/.exec(xml)?.[1] ?? '';
  const series: ChartSeries[] = [];
  let categories: GridChart['categories'] = { ref: null, cache: [] };
  for (const m of body.matchAll(/<ser>([\s\S]*?)<\/ser>/g)) {
    const ser = m[1];
    const tx = /<tx>([\s\S]*?)<\/tx>/.exec(ser)?.[1];
    const cat = /<cat>([\s\S]*?)<\/cat>/.exec(ser)?.[1];
    const val = /<val>([\s\S]*?)<\/val>/.exec(ser)?.[1];
    const name = tx ? (cachePoints(tx)[0] ?? decodeXml(/<v>([^<]*)<\/v>/.exec(tx)?.[1] ?? '')) : `계열${series.length + 1}`;
    if (cat && categories.cache.length === 0 && !categories.ref) categories = { ref: refOf(cat), cache: cachePoints(cat) };
    series.push({ name, nameRef: refOf(tx), values: { ref: refOf(val), cache: val ? cachePoints(val).map((v) => Number(v) || 0) : [] } });
  }
  if (series.length === 0) return null;
  const title = titleXml ? textOf(titleXml) : '';
  return { kind, title, categories, series };
}

function anchorPoint(xml: string, tag: 'from' | 'to'): ChartAnchorPoint | null {
  const m = new RegExp(`<${tag}>([\\s\\S]*?)</${tag}>`).exec(xml);
  if (!m) return null;
  const num = (name: string) => Number(new RegExp(`<${name}>(-?\\d+)</${name}>`).exec(m[1])?.[1] ?? 0);
  return { col: num('col'), row: num('row'), dx: num('colOff') / EMU_PER_PX, dy: num('rowOff') / EMU_PER_PX };
}

/** 그리기 XML의 차트 앵커들 — 차트 관계 id와 위치 */
export function chartAnchorsOf(source: string): Array<{ rid: string; from: ChartAnchorPoint; to: ChartAnchorPoint | null; size: GridChart['size'] }> {
  const xml = unprefixed(source);
  const out: Array<{ rid: string; from: ChartAnchorPoint; to: ChartAnchorPoint | null; size: GridChart['size'] }> = [];
  for (const m of xml.matchAll(/<(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>([\s\S]*?)<\/\1>/g)) {
    const body = m[2];
    const rid = /<chart\b[^>]*\br:id="([^"]+)"/.exec(body)?.[1];
    if (!rid) continue;
    const from = anchorPoint(body, 'from') ?? { row: 0, col: 0, dx: 0, dy: 0 };
    const ext = /<ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(body);
    out.push({
      rid,
      from,
      to: m[1] === 'twoCellAnchor' ? anchorPoint(body, 'to') : null,
      size: m[1] !== 'twoCellAnchor' && ext ? { width: Number(ext[1]) / EMU_PER_PX, height: Number(ext[2]) / EMU_PER_PX } : null,
    });
  }
  return out;
}


// ---------------------------------------------------------------------------
// 새 차트 XML 만들기(저장)
// ---------------------------------------------------------------------------

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const OFFICE_SERIES_COLORS = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

/** 새 차트 하나의 chart XML — 참조 범위와 지금 값(캐시)을 함께 쓴다(Excel·LibreOffice 모두 바로 그린다). */
export function buildChartXml(chart: GridChart, cache: { categories: readonly string[]; series: ReadonlyArray<{ name: string; values: readonly number[] }> }): string {
  const sheet = chart.sheetName;
  const strCache = (items: readonly string[]) =>
    `<c:strCache><c:ptCount val="${items.length}"/>${items.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (items: readonly number[]) =>
    `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${items.length}"/>${items.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`;
  const pie = chart.kind === 'pie';
  const sers = chart.series
    .map((ser, i) => {
      const color = OFFICE_SERIES_COLORS[i % OFFICE_SERIES_COLORS.length];
      const name = cache.series[i]?.name ?? ser.name;
      const values = cache.series[i]?.values ?? ser.values.cache;
      const tx = ser.nameRef
        ? `<c:tx><c:strRef><c:f>${escapeXml(rangeFormula(ser.nameRef.sheetName ?? sheet, ser.nameRef))}</c:f>${strCache([name])}</c:strRef></c:tx>`
        : `<c:tx><c:v>${escapeXml(name)}</c:v></c:tx>`;
      const fill = pie
        ? cache.categories.map((_, p) => `<c:dPt><c:idx val="${p}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${OFFICE_SERIES_COLORS[p % OFFICE_SERIES_COLORS.length]}"/></a:solidFill></c:spPr></c:dPt>`).join('')
        : chart.kind === 'line'
          ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>`
          : `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>`;
      const cat = chart.categories.ref
        ? `<c:cat><c:strRef><c:f>${escapeXml(rangeFormula(chart.categories.ref.sheetName ?? sheet, chart.categories.ref))}</c:f>${strCache(cache.categories)}</c:strRef></c:cat>`
        : `<c:cat><c:strLit>${strCache(cache.categories).replace(/^<c:strCache>|<\/c:strCache>$/g, '')}</c:strLit></c:cat>`;
      const val = ser.values.ref
        ? `<c:val><c:numRef><c:f>${escapeXml(rangeFormula(ser.values.ref.sheetName ?? sheet, ser.values.ref))}</c:f>${numCache(values)}</c:numRef></c:val>`
        : `<c:val><c:numLit>${numCache(values).replace(/^<c:numCache>|<\/c:numCache>$/g, '')}</c:numLit></c:val>`;
      return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${fill}${cat}${val}${chart.kind === 'line' ? '<c:smooth val="0"/>' : ''}</c:ser>`;
    })
    .join('');
  const axes =
    '<c:catAx><c:axId val="50010001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>' +
    `<c:axPos val="${chart.kind === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` +
    '<c:crossAx val="50010002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
    '<c:valAx><c:axId val="50010002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>' +
    `<c:axPos val="${chart.kind === 'bar' ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>` +
    '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
    '<c:crossAx val="50010001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>';
  const plot =
    chart.kind === 'pie'
      ? `<c:pieChart><c:varyColors val="1"/>${sers}<c:firstSliceAng val="0"/></c:pieChart>`
      : chart.kind === 'line'
        ? `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}<c:marker val="1"/><c:axId val="50010001"/><c:axId val="50010002"/></c:lineChart>${axes}`
        : `<c:barChart><c:barDir val="${chart.kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>${sers}<c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="50010001"/><c:axId val="50010002"/></c:barChart>${axes}`;
  const title = chart.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="ko-KR" sz="1400" b="0"/><a:t>${escapeXml(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  const legend = chart.series.length > 1 || pie ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : '';
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    `<c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:chartSpace>'
  );
}

/** 새 차트를 그리기에 붙이는 앵커(oneCellAnchor — 왼쪽 위 칸 + 크기) */
export function buildChartAnchorXml(chart: GridChart, rid: string, shapeId: number): string {
  const width = Math.round((chart.size?.width ?? 480) * EMU_PER_PX);
  const height = Math.round((chart.size?.height ?? 288) * EMU_PER_PX);
  return (
    `<xdr:oneCellAnchor><xdr:from><xdr:col>${chart.from.col}</xdr:col><xdr:colOff>${Math.round(chart.from.dx * EMU_PER_PX)}</xdr:colOff>` +
    `<xdr:row>${chart.from.row}</xdr:row><xdr:rowOff>${Math.round(chart.from.dy * EMU_PER_PX)}</xdr:rowOff></xdr:from>` +
    `<xdr:ext cx="${width}" cy="${height}"/>` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${shapeId}" name="차트 ${shapeId}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`
  );
}
