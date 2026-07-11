import ExcelJS from 'exceljs';

/** 그리드 뷰 데이터 모델 — 원본 파일 대신 표시용 값·스타일만 담는다. */
export interface GridCell {
  readonly v: string;
  readonly s?: number;
}

export interface GridMerge {
  readonly r: number;
  readonly c: number;
  readonly rs: number;
  readonly cs: number;
}

export interface GridSheet {
  readonly name: string;
  readonly rowCount: number;
  readonly colCount: number;
  readonly colWidths: readonly number[];
  readonly merges: readonly GridMerge[];
  readonly rows: ReadonlyArray<ReadonlyArray<GridCell | 0>>;
}

export interface GridStyle {
  readonly b?: 1;
  readonly col?: string;
  readonly bg?: string;
  readonly a?: 'l' | 'c' | 'r';
  readonly w?: 1;
}

export interface GridDocument {
  readonly sheets: readonly GridSheet[];
  readonly styles: readonly GridStyle[];
}

/** JSONL 헤더(1행): 셀 데이터를 제외한 시트 메타 정보 */
export interface GridHeader {
  readonly styles: readonly GridStyle[];
  readonly sheets: ReadonlyArray<Omit<GridSheet, 'rows'>>;
}

/**
 * 그리드를 JSONL로 직렬화한다: 1행 = 헤더, 이후 시트 순서대로 행당 1줄.
 * 수십만 행 문서를 요청 범위만 파싱해 서빙할 수 있도록 하기 위함이다.
 */
export function serializeGrid(grid: GridDocument): Buffer {
  const header: GridHeader = {
    styles: grid.styles,
    sheets: grid.sheets.map(({ rows: _rows, ...meta }) => meta),
  };
  const parts: string[] = [JSON.stringify(header)];
  for (const sheet of grid.sheets) {
    for (const row of sheet.rows) {
      parts.push(JSON.stringify(row));
    }
  }
  return Buffer.from(parts.join('\n'), 'utf8');
}

const MAX_ROWS = 300_000;
const MAX_COLS = 120;
const DEFAULT_COL_WIDTH_PX = 80;
const CHAR_TO_PX = 7.2;

function argbToCss(argb?: string): string | undefined {
  if (!argb || argb.length !== 8) return undefined;
  return `#${argb.slice(2)}`;
}

function formatNumber(value: number, numFmt?: string): string {
  if (!Number.isFinite(value)) return String(value);
  if (numFmt && /[#0],/.test(numFmt)) {
    return value.toLocaleString('ko-KR', { maximumFractionDigits: 6 });
  }
  return String(value);
}

function formatDate(value: Date): string {
  const y = value.getUTCFullYear();
  const m = String(value.getUTCMonth() + 1).padStart(2, '0');
  const d = String(value.getUTCDate()).padStart(2, '0');
  const hh = value.getUTCHours();
  const mm = value.getUTCMinutes();
  return hh === 0 && mm === 0
    ? `${y}-${m}-${d}`
    : `${y}-${m}-${d} ${String(hh).padStart(2, '0')}:${String(mm).padStart(2, '0')}`;
}

function displayValue(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return formatDate(value);
  if (typeof value === 'number') return formatNumber(value, cell.numFmt);
  if (typeof value === 'object') {
    const rich = value as ExcelJS.CellRichTextValue;
    if (Array.isArray(rich.richText)) return rich.richText.map((run) => run.text).join('');
    const link = value as ExcelJS.CellHyperlinkValue;
    if (typeof link.text === 'string') return link.text;
    const formula = value as ExcelJS.CellFormulaValue;
    if ('formula' in formula) {
      const result = formula.result;
      if (result === null || result === undefined) return '';
      if (result instanceof Date) return formatDate(result);
      if (typeof result === 'number') return formatNumber(result, cell.numFmt);
      return String(result);
    }
    return '';
  }
  return String(value);
}

function parseMergeRange(range: string): GridMerge | null {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!match) return null;
  const colNum = (letters: string) =>
    [...letters].reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
  const c1 = colNum(match[1]);
  const r1 = Number(match[2]);
  const c2 = colNum(match[3]);
  const r2 = Number(match[4]);
  return { r: r1, c: c1, rs: r2 - r1 + 1, cs: c2 - c1 + 1 };
}

class StylePalette {
  private readonly styles: GridStyle[] = [];
  private readonly index = new Map<string, number>();

  register(style: GridStyle): number | undefined {
    if (Object.keys(style).length === 0) return undefined;
    const key = JSON.stringify(style);
    const existing = this.index.get(key);
    if (existing !== undefined) return existing;
    this.styles.push(style);
    this.index.set(key, this.styles.length - 1);
    return this.styles.length - 1;
  }

  toArray(): GridStyle[] {
    return this.styles;
  }
}

function extractStyle(cell: ExcelJS.Cell, palette: StylePalette): number | undefined {
  const style: {
    b?: 1;
    col?: string;
    bg?: string;
    a?: 'l' | 'c' | 'r';
    w?: 1;
  } = {};
  if (cell.font?.bold) style.b = 1;
  const fontColor = argbToCss(cell.font?.color?.argb);
  if (fontColor && fontColor !== '#000000') style.col = fontColor;
  const fill = cell.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const bg = argbToCss(fill.fgColor?.argb);
    if (bg && bg !== '#FFFFFF') style.bg = bg;
  }
  const horizontal = cell.alignment?.horizontal;
  if (horizontal === 'center') style.a = 'c';
  else if (horizontal === 'right') style.a = 'r';
  if (cell.alignment?.wrapText) style.w = 1;
  return palette.register(style);
}

export async function parseXlsxToGrid(xlsxBytes: Buffer): Promise<GridDocument> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(xlsxBytes as unknown as ArrayBuffer);
  const palette = new StylePalette();
  const sheets: GridSheet[] = [];

  workbook.eachSheet((sheet) => {
    if (sheet.state && sheet.state !== 'visible') return;
    const rowCount = Math.min(sheet.actualRowCount || 0, MAX_ROWS);
    const colCount = Math.min(sheet.actualColumnCount || 0, MAX_COLS);
    if (rowCount === 0 || colCount === 0) return;

    const colWidths: number[] = [];
    for (let c = 1; c <= colCount; c += 1) {
      const width = sheet.getColumn(c).width;
      colWidths.push(width ? Math.round(width * CHAR_TO_PX) : DEFAULT_COL_WIDTH_PX);
    }

    const rows: Array<Array<GridCell | 0>> = [];
    for (let r = 1; r <= rowCount; r += 1) {
      const row = sheet.getRow(r);
      const cells: Array<GridCell | 0> = new Array(colCount).fill(0);
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        if (colNumber > colCount) return;
        const text = displayValue(cell);
        const styleRef = extractStyle(cell, palette);
        if (text === '' && styleRef === undefined) return;
        cells[colNumber - 1] = styleRef === undefined ? { v: text } : { v: text, s: styleRef };
      });
      rows.push(cells);
    }

    // 서식만 남은 유령 행(값 없는 꼬리 행)을 잘라내 행 수를 실제 데이터 기준으로 맞춘다
    let lastWithValue = rows.length;
    while (lastWithValue > 1) {
      const row = rows[lastWithValue - 1];
      const hasValue = row.some((cell) => cell !== 0 && cell.v !== '');
      if (hasValue) break;
      lastWithValue -= 1;
    }
    rows.length = lastWithValue;

    const merges = (
      (sheet.model as { merges?: readonly string[] }).merges ?? []
    )
      .map(parseMergeRange)
      .filter(
        (merge): merge is GridMerge => merge !== null && merge.r <= rows.length && merge.c <= colCount,
      );

    sheets.push({ name: sheet.name, rowCount: rows.length, colCount, colWidths, merges, rows });
  });

  return { sheets, styles: palette.toArray() };
}

/** 간단한 RFC4180 CSV 파서 (따옴표·이스케이프 지원) */
export function parseCsvToGrid(text: string): GridDocument {
  const rows: string[][] = [];
  let field = '';
  let record: string[] = [];
  let inQuotes = false;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      record.push(field);
      field = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i += 1;
      record.push(field);
      field = '';
      rows.push(record);
      record = [];
    } else {
      field += ch;
    }
  }
  if (field !== '' || record.length > 0) {
    record.push(field);
    rows.push(record);
  }

  const bounded = rows.slice(0, MAX_ROWS);
  const colCount = Math.min(Math.max(...bounded.map((r) => r.length), 0), MAX_COLS);
  const gridRows = bounded.map((r) => {
    const cells: Array<GridCell | 0> = new Array(colCount).fill(0);
    for (let c = 0; c < Math.min(r.length, colCount); c += 1) {
      if (r[c] !== '') cells[c] = { v: r[c] };
    }
    return cells;
  });

  return {
    sheets: [
      {
        name: 'Sheet1',
        rowCount: gridRows.length,
        colCount,
        colWidths: new Array(colCount).fill(120),
        merges: [],
        rows: gridRows,
      },
    ],
    styles: [],
  };
}
