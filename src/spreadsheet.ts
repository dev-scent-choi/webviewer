import ExcelJS from 'exceljs';
import type { DocumentStore } from './storage.js';
import type { DocumentMeta } from './types.js';
import { buildSheetNameIndex, computeGridCellDisplay } from './xlsx/formulas.js';
import { dateToDateText, isDateNumFmt, serialToDateText } from './collab/cellValueShared.js';
import { EXCEL_INDEXED_COLORS, formatNumberCode, hashIfTooWide, isNumberOrDateValue, isTextNumFmt } from './collab/numberFormat.js';
import { WILDCARD_WORK_BUDGET, withWildcardBudget, type WildcardBudget } from './collab/wildcard.js';
import excelBuiltinNumFmts from 'exceljs/lib/xlsx/defaultnumformats.js';

/**
 * 파일에 서식 코드 없이 번호(numFmtId)만 적힌 기본 서식 — 엑셀은 번호의 뜻을 사용자 로캘로 정한다. ExcelJS는 미국식 코드만 알고
 * 한국어 엑셀의 통화(5~8)·회계(41~44)·한글 날짜(27~36, 50~58)는 아예 몰라 "일반"(날짜는 일련번호)이 됐다.
 * 우리 사용자는 한국어 엑셀이므로 ExcelJS의 기본 서식 표를 한국어 엑셀 코드로 채운다(읽기 전에 한 번 — 모든 경로가 이 모듈을 거친다).
 * 주의: 이 표는 ExcelJS 모듈 전역이다 — 이 프로세스에서 ExcelJS로 읽는 모든 파일이 기본 서식 번호를 한국어 엑셀 뜻으로 읽는다
 * (다른 로캘 파일을 그 로캘 뜻으로 읽어야 할 일이 생기면 여기를 바꿔야 한다). 쓰기용 번호 표(defaultFmtHash)는 ExcelJS가
 * 모듈을 불러올 때 영어 코드로 이미 만들어 두므로, 저장할 때 이 한국어 코드가 기본 번호로 바뀌어 들어가지는 않는다.
 */
const KO_BUILTIN_NUMFMTS: Readonly<Record<number, string>> = {
  5: '"₩"#,##0;\\-"₩"#,##0',
  6: '"₩"#,##0;[Red]\\-"₩"#,##0',
  7: '"₩"#,##0.00;\\-"₩"#,##0.00',
  8: '"₩"#,##0.00;[Red]\\-"₩"#,##0.00',
  14: 'yyyy-mm-dd',
  22: 'yyyy-mm-dd h:mm',
  37: '#,##0;-#,##0',
  38: '#,##0;[Red]-#,##0',
  39: '#,##0.00;-#,##0.00',
  40: '#,##0.00;[Red]-#,##0.00',
  41: '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-',
  42: '_-"₩"* #,##0_-;\\-"₩"* #,##0_-;_-"₩"* "-"_-;_-@_-',
  43: '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-',
  44: '_-"₩"* #,##0.00_-;\\-"₩"* #,##0.00_-;_-"₩"* "-"??_-;_-@_-',
};
for (const [id, entry] of Object.entries(excelBuiltinNumFmts as Record<string, { f?: string; 'ko-kr'?: string }>)) {
  const ko = KO_BUILTIN_NUMFMTS[Number(id)] ?? entry['ko-kr'];
  if (ko) entry.f = ko;
}
for (const [id, f] of Object.entries(KO_BUILTIN_NUMFMTS)) (excelBuiltinNumFmts as Record<string, { f?: string }>)[id] ??= { f };
import type { BorderSpec } from './collab/cellBoxCss.js';
import { sheetPathsByName, stripDrawings } from './xlsx/drawingPassthrough.js';
import { tableCellLook, withTableLook } from './xlsx/tableStyles.js';
import { sheetTablesOf } from './xlsx/sheetExtras.js';
import { unzipSync, strFromU8 } from 'fflate';
import {
  EXCEL_DEFAULT_FONT_PX,
  EXCEL_DEFAULT_ROW_HEIGHT_PT,
  autoFitRowHeightPx,
  defaultColWidthPx,
  defaultFontOfStylesXml,
  encodeRowHeights,
  estimateWrappedLines,
  excelColWidthToPx,
  maxDigitWidthPx,
  rowHeightPtToPx,
  type RowHeightRuns,
} from './xlsx/sheetGeometry.js';

/** 대형 시트를 훑는 동기 반복문 중간에 이벤트 루프에 제어권을 돌려준다 — 이게 없으면 큰 xlsx를
 * 열 때(파싱이든 동시편집 시딩이든) 그 몇 초~몇십 초 동안 Node가 단일 스레드라 그 문서를 연
 * 사용자뿐 아니라 그 순간 서버에 접속한 모든 사용자·모든 다른 문서 요청까지 함께 멈춘다. */
export function yieldToEventLoop(): Promise<void> {
  return new Promise((resolve) => setImmediate(resolve));
}

/** 이 개수의 행마다 한 번씩 yieldToEventLoop를 호출한다 — 너무 자주 양보하면(예: 매 행) setImmediate
 * 오버헤드 자체가 커지고, 너무 뜸하면 양보 효과가 없다. 2000행은 실측 없이 정한 보수적인 값이라
 * 필요하면 나중에 조정 가능하다. */
export const YIELD_EVERY_N_ROWS = 2000;

/** parseXlsxToGrid의 수식 계산 패스가 이 개수의 **수식 셀**마다 한 번씩 양보한다 — 행이 아니라
 * 셀 단위인 이유는, 그 패스가 도는 대상이 "모든 행"이 아니라 "수식이 들어 있는 칸만 모은 목록"이라
 * 행 기준 간격으로는 (수식이 한 행에 몰려 있을 때) 아무 효과가 없을 수 있기 때문이다. 수식 한 칸의
 * 계산 비용은 참조 범위 크기에 비례해 최대 MAX_RANGE_CELLS(xlsx/formulas.ts)까지 갈 수 있으므로,
 * 행 단위(2000)보다 훨씬 촘촘하게 잡아 대형 파일 업로드 한 번이 서버 전체를 멈추지 않게 한다
 * (room.ts의 RECOMPUTE_YIELD_EVERY_N_ROWS가 같은 이유로 페이지 크기의 1/4을 쓰는 것과 같은 취지). */
export const YIELD_EVERY_N_FORMULA_CELLS = 500;

/** 그리드 뷰 데이터 모델 — 원본 파일 대신 표시용 값·스타일만 담는다. */
export interface GridCell {
  /** 화면에 보일 글자(칸 서식을 입힌 모양) */
  readonly v: string;
  /** 값 글자 — 서식 때문에 v와 다를 때만(예: v "(4,321)" → n "-4321"). gridCellValue 참고 */
  readonly n?: string;
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
  /** 열별 화면 폭(px) — 숨긴 열은 0 */
  readonly colWidths: readonly number[];
  /** colWidths 밖(파일에 너비가 없는 뒤쪽 열)의 폭(px) */
  readonly defaultColWidth: number;
  /** 기본 행 높이(px)와, 그와 다른 행(숨긴 행은 0)의 높이 묶음 — sheetGeometry.ts 참고 */
  readonly defaultRowHeight: number;
  readonly rowHeights: RowHeightRuns;
  /** 시트 보기에서 격자선을 껐으면 false */
  readonly showGridLines: boolean;
  readonly merges: readonly GridMerge[];
  /** 엑셀 "틀 고정" — 위 N행/왼쪽 N열은 스크롤해도 계속 보인다. 없으면 둘 다 0. */
  readonly freezeRows: number;
  readonly freezeCols: number;
  readonly rows: ReadonlyArray<ReadonlyArray<GridCell | 0>>;
}

/** 테두리 굵기 — 1=가는 선, 2=굵은 선(엑셀 thin/medium을 단순화. 그 외 스타일은 가장 가까운 쪽으로 반올림). */
export type BorderWeight = 1 | 2;

/** 숫자 표시 형식 종류 — gridShared.ts의 CellStyleAttrs와 동일한 단순화(원본 서식 코드를 몇 가지
 * 자주 쓰는 범주로 분류). */
export type NumberFormatKind = 'number' | 'currency' | 'percent';

export interface GridStyle {
  readonly b?: 1;
  readonly i?: 1;
  readonly u?: 1;
  readonly st?: 1;
  /** 글자 크기(px). 기본(11pt≈15px)과 다를 때만 기록 */
  readonly fs?: number;
  /** 글꼴 이름. 기본 글꼴과 같으면 기록하지 않는다. */
  readonly ff?: string;
  readonly col?: string;
  readonly bg?: string;
  readonly a?: 'l' | 'c' | 'r';
  /** 세로 정렬 — 위/가운데/아래. */
  readonly va?: 't' | 'm' | 'b';
  readonly w?: 1;
  /** 숫자 표시 형식(쉼표/통화/백분율) + 소수 자릿수. */
  readonly nf?: NumberFormatKind;
  readonly nd?: number;
  /** 날짜 표시 형식 칸(gridShared.ts CellStyleAttrs.df와 같은 의미). */
  readonly df?: 1;
  /** 원본 엑셀 표시 형식 코드(일반 서식이 아닐 때) — 보기·편집 화면 모두 이 코드로 numberFormat.ts가 그린다 */
  readonly fmt?: string;
  /** 서식 코드가 값에 따라 정한 글자 색([Red] 등) — 보기 화면 전용(편집 화면은 값마다 다시 계산한다) */
  readonly fc?: string;
  /** 보기 화면 전용 — 조건부 서식을 구운 결과(xlsx/viewerCondFormat.ts): 데이터 막대 [색, 길이 %], 아이콘 [기호, 색, 값 숨김] */
  readonly cb?: readonly [string, number];
  readonly ci?: readonly [string, string, 1?];
  /** 시트 보호가 걸려 있고 이 셀이 잠겨 있음(=편집 불가) */
  readonly lk?: 1;
  /** 엑셀 데이터 유효성 검사(정수/소수)가 걸려 있음 */
  readonly nu?: 1;
  readonly bt?: BorderWeight;
  readonly br?: BorderWeight;
  readonly bb?: BorderWeight;
  readonly bl?: BorderWeight;
  /** 굵기(bt…)만으로 못 나타내는 테두리 변 — [엑셀 선 종류, 색] (collab/cellBoxCss.ts) */
  readonly bd?: { readonly t?: BorderSpec; readonly r?: BorderSpec; readonly b?: BorderSpec; readonly l?: BorderSpec };
  /** 무늬 채우기 [무늬, 무늬 색, 바탕색](solid 아닌 것) */
  readonly pat?: readonly [string, string?, string?];
  /** 글자 회전(엑셀 textRotation, 255 = 세로쓰기) */
  readonly rot?: number;
  /** 들여쓰기 단계 */
  readonly ind?: number;
  /** 셀에 맞춤(글자 크기를 줄여 칸에 맞춘다) */
  readonly shr?: 1;
}

export interface GridDocument {
  readonly sheets: readonly GridSheet[];
  readonly styles: readonly GridStyle[];
  /** 보기 화면 전용 — TODAY()·NOW() 조건부 서식을 구운 날(현지 YYYY-MM-DD). 날짜가 바뀌면 그리드를 다시 만든다 */
  readonly cfDay?: string;
  /** 보기 화면 전용 — 조건부 서식 일부를 시간 예산·비용 상한 때문에 굽지 않았다(화면이 알린다) */
  readonly cfSkipped?: 1;
}

/**
 * 캐시(derived kind 'grid') 포맷 버전 — 그리드 "구조"(JSONL 줄 모양)가 아니라, 그 안의 셀 값을
 * "어떻게 계산했는지"가 바뀔 때마다 올린다. 지금 값(2)은 xlsx/formulas.ts의 시트 간 참조(cross-sheet
 * reference) 버그 수정에 대응한다 — 그 수정 전에 만들어진 캐시는(버전 필드가 아예 없다, 옛
 * serializeGrid가 안 써 넣었으므로) 시트를 넘나드는 수식이 조용히 틀린 값(또는 계산 전이라면
 * 원문 그대로)으로 저장돼 있을 수 있다. runGridExtraction(routes/documents.ts)이 이 값을 읽어
 * "다음 조회 시점에" 낡은 캐시를 감지해 원본에서 다시 계산하도록 하는 lazy self-heal의 기준선이다
 * (routes/viewer.ts의 loadGridUncached 참고) — 기존 문서를 한꺼번에 마이그레이션하지 않는다(그 자체가
 * 새로운 리스크이므로). 앞으로 계산 로직이 바뀔 때마다(예: 새로운 함수 지원, 다른 계산 버그 수정)
 * 이 값을 올리면 같은 메커니즘이 그대로 재사용된다. */
export const GRID_CACHE_FORMAT_VERSION = 5;

/** JSONL 헤더(1행): 셀 데이터를 제외한 시트 메타 정보 */
export interface GridHeader {
  /** GRID_CACHE_FORMAT_VERSION 문서 주석 참고. 이 필드가 없거나(구버전 캐시) 현재 버전과 다르면
   * "낡은 캐시"로 취급한다. */
  readonly formatVersion?: number;
  /** GridDocument.cfDay·cfSkipped와 같다 */
  readonly cfDay?: string;
  readonly cfSkipped?: 1;
  readonly styles: readonly GridStyle[];
  readonly sheets: ReadonlyArray<Omit<GridSheet, 'rows'>>;
}

/**
 * 그리드를 JSONL로 직렬화한다: 1행 = 헤더, 이후 시트 순서대로 행당 1줄.
 * 수십만 행 문서를 요청 범위만 파싱해 서빙할 수 있도록 하기 위함이다.
 */
export function serializeGrid(grid: GridDocument): Buffer {
  const header: GridHeader = {
    formatVersion: GRID_CACHE_FORMAT_VERSION,
    ...(grid.cfDay ? { cfDay: grid.cfDay } : {}),
    ...(grid.cfSkipped ? { cfSkipped: grid.cfSkipped } : {}),
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

/**
 * JSONL 버퍼에서 각 줄이 시작하는 바이트 오프셋 배열을 만든다(0번째는 항상 0 = 헤더 시작줄).
 * src/routes/viewer.ts의 그리드 조회(sliceRows/searchGrid)가 필요한 범위만 파싱할 수 있게
 * 해주는 인덱스 — 수십만 행이면 이 스캔 자체가 O(파일 크기)라, 그리드를 만드는 시점에 미리
 * 계산해 saveGridDerived로 함께 저장해 두면 조회 시점에 매번 다시 스캔하지 않아도 된다.
 */
export function computeLineOffsets(buffer: Buffer): readonly number[] {
  const lineOffsets: number[] = [0];
  for (let i = buffer.indexOf(10); i !== -1; i = buffer.indexOf(10, i + 1)) {
    lineOffsets.push(i + 1);
  }
  return lineOffsets;
}

/**
 * 그리드를 JSONL로 직렬화해 저장하고(derived kind 'grid'), 같은 바이트로 만든 줄 오프셋
 * 인덱스도 함께 저장한다(derived kind 'gridindex'). 업로드 시 최초 추출
 * (routes/documents.ts의 runGridExtraction)과 동시편집 저장 후 재생성(collab/room.ts)
 * 양쪽에서 재사용해, 그리드 내용이 바뀔 때마다 인덱스가 항상 최신 상태로 유지되게 한다
 * (그리드가 바뀌지 않는 한 인덱스도 따로 무효화할 필요가 없다).
 *
 * **알려진, 좁은 쓰기 경합(리뷰 지적, 구조적으로 고치지 않기로 함)**: 아래 두 `store.saveDerived`
 * 호출은 서로 독립적이고(문서 하나당 락이 없다), storage.ts의 saveDerived 자체도 saveOriginal과
 * 달리 temp-file+rename의 원자적 쓰기가 아니라 `fs.writeFile`로 그 경로에 직접 쓴다. 그래서
 * 같은 문서에 대해 이 함수가 **동시에 두 번** 호출되면(예: collab/room.ts의 flushToStore가 이
 * 문서를 유휴 자동저장하는 바로 그 순간, routes/viewer.ts의 regenerateStaleGridCache가 같은
 * 문서를 낡은 포맷 캐시로 감지해 재생성을 트리거하는 경우) 두 호출의 grid/gridindex 쓰기가
 * 시간순으로 뒤섞여, "한쪽의 grid + 다른 쪽의 gridindex"처럼 서로 짝이 안 맞는 조합이 잠깐 disk에
 * 남을 수 있다 — 이 조합이 우연히 routes/viewer.ts의 offsetsLookValid 같은 구조적 정합성 검사를
 * (진짜 내용은 안 맞아도 통과할 만큼) 우연히 통과하면, 다음 성공적인 저장 전까지 일시적으로 깨진
 * 파싱 결과를 서빙할 수 있다.
 *
 * 이 경합은 좁다 — "이 문서가 포맷이 낡은 grid 캐시를 갖고 있으면서, 동시에 바로 그 순간 실제로
 * 동시편집 중이어야" 하는 아주 구체적인 타이밍이 겹쳐야 하고, 어느 한쪽이든 다음에 성공적으로
 * (한쪽만) 저장하면 다시 짝이 맞는 grid/gridindex 쌍이 돼 저절로 회복된다(자가치유) — 영구적인
 * 데이터 손실이 아니다. 제대로 고치려면 문서 단위의 실제 프로세스 간/요청 간 쓰기 조정(락)이나
 * 원자적 다중 파일 쓰기 스킴이 필요한데, 이 정도 코드 변경은 이 함수의 범위를 넘어선다 — 지금은
 * 이 주석으로만 남겨 다음에 이 경로를 건드리는 사람이 바로 알 수 있게 해둔다.
 */
export async function saveGridDerived(
  store: Pick<DocumentStore, 'saveDerived'>,
  docId: string,
  grid: GridDocument,
): Promise<void> {
  const serialized = serializeGrid(grid);
  await store.saveDerived(docId, 'grid', serialized);
  await store.saveDerived(
    docId,
    'gridindex',
    Buffer.from(JSON.stringify({ lineOffsets: computeLineOffsets(serialized) })),
  );
}

/**
 * 문서 원본(xlsx/csv)을 읽어 GridDocument로 만든다 — routes/documents.ts의 최초 업로드 추출
 * (runGridExtraction)과 routes/viewer.ts의 낡은 캐시 self-heal(loadGridUncached) 양쪽이 공유하는
 * "확장자별로 어떤 파서를 쓸지"의 단일 소스. store를 인자로 받는 이유는 saveGridDerived와 같다 —
 * 이 모듈이 storage.ts를 값으로(순환 없이) 참조하지 않게 하기 위함. */
export async function buildGridFromOriginal(
  store: Pick<DocumentStore, 'readOriginal'>,
  meta: Pick<DocumentMeta, 'id' | 'ext'>,
): Promise<GridDocument> {
  const original = await store.readOriginal(meta.id, meta.ext);
  if (meta.ext === 'csv') {
    return parseCsvToGrid(original.toString('utf8'));
  }
  return parseXlsxToGrid(original);
}

/**
 * xlsx 파싱이 그리드로 받아들이는 최대 행/열 — 이 코드베이스가 "시트 하나의 현실적인 크기"로
 * 인정하는 유일한 기준선이다. 동시편집 서버(collab/room.ts)의 canonical 그리드 성장 상한도
 * 반드시 이 값을 그대로 써야 한다: 파서가 절대 만들어내지 못하는 크기를 canonical만 허용하면,
 * 조작된 소켓 메시지 하나로 파일에서는 있을 수 없는 크기의 배열을 서버에 할당시킬 수 있다
 * (YIELD_EVERY_N_ROWS를 여기서 export해 서버/클라이언트가 공유하는 것과 같은 이유).
 */
// 엑셀 자체 한도(1,048,576행) — SAP에서 뽑은 87만 행 같은 실무 파일이 있다. 예전 30만 행 한도는 그 뒤 행이 보이지도 저장되지도 않았다
export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 120;

function argbToCss(argb?: string): string | undefined {
  // 16진수만 — 이 값은 보기 화면 <style>에 그대로 들어가 'FF;}*{x:' 같은 값이 규칙을 깰 수 있었다
  if (!argb || !/^[0-9a-f]{6}([0-9a-f]{2})?$/i.test(argb)) return undefined;
  // 대부분의 엑셀 파일(과 ExcelJS 자신)은 8자리 AARRGGBB로 쓰지만, POI는 알파 없이 6자리
  // RRGGBB만 쓰는 경우가 있다 — 두 형태 모두 지원해야 색이 조용히 누락되지 않는다.
  if (argb.length === 8) return `#${argb.slice(2)}`;
  if (argb.length === 6) return `#${argb}`;
  return undefined;
}

/**
 * 테마 색 인덱스 순서 (Excel은 dk1/lt1, dk2/lt2가 스왑된 관례를 사용한다)
 */
const THEME_COLOR_ORDER = [
  'lt1', 'dk1', 'lt2', 'dk2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink',
] as const;

/** 워크북 theme1.xml에서 테마 팔레트(RRGGBB)를 추출한다 */
export function parseThemeColors(themeXml: string | undefined): readonly string[] {
  if (!themeXml) return [];
  const scheme = themeXml.match(/<a:clrScheme[\s\S]*?<\/a:clrScheme>/)?.[0];
  if (!scheme) return [];
  const byName = new Map<string, string>();
  for (const match of scheme.matchAll(
    /<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>\s*<a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr[^>]*lastClr="([0-9A-Fa-f]{6})")/g,
  )) {
    byName.set(match[1], (match[2] ?? match[3]).toUpperCase());
  }
  return THEME_COLOR_ORDER.map((name) => byName.get(name) ?? '000000');
}

/** Excel tint: HSL 명도(L)를 tint>0이면 흰색 쪽, tint<0이면 검정 쪽으로 보간한다 */
function applyTint(hex: string, tint: number): string {
  const r = parseInt(hex.slice(0, 2), 16) / 255;
  const g = parseInt(hex.slice(2, 4), 16) / 255;
  const b = parseInt(hex.slice(4, 6), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  const l = (max + min) / 2;
  const d = max - min;
  let s = 0;
  if (d !== 0) {
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
    else if (max === g) h = ((b - r) / d + 2) / 6;
    else h = ((r - g) / d + 4) / 6;
  }
  const newL = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hueToRgb = (p: number, q: number, t: number) => {
    let tt = t;
    if (tt < 0) tt += 1;
    if (tt > 1) tt -= 1;
    if (tt < 1 / 6) return p + (q - p) * 6 * tt;
    if (tt < 1 / 2) return q;
    if (tt < 2 / 3) return p + (q - p) * (2 / 3 - tt) * 6;
    return p;
  };
  let r2: number;
  let g2: number;
  let b2: number;
  if (s === 0) {
    r2 = g2 = b2 = newL;
  } else {
    const q = newL < 0.5 ? newL * (1 + s) : newL + s - newL * s;
    const p = 2 * newL - q;
    r2 = hueToRgb(p, q, h + 1 / 3);
    g2 = hueToRgb(p, q, h);
    b2 = hueToRgb(p, q, h - 1 / 3);
  }
  const toHex = (v: number) => Math.round(Math.min(1, Math.max(0, v)) * 255).toString(16).padStart(2, '0');
  return `${toHex(r2)}${toHex(g2)}${toHex(b2)}`.toUpperCase();
}

interface ExcelColorLike {
  readonly argb?: string;
  readonly theme?: number;
  readonly tint?: number;
  readonly indexed?: number;
}

/** argb / theme+tint / indexed 세 가지 색 표현을 모두 CSS 색으로 해석한다 */
export function resolveColor(
  color: ExcelColorLike | undefined,
  themeColors: readonly string[],
): string | undefined {
  if (!color) return undefined;
  if (color.argb) return argbToCss(color.argb);
  if (color.theme !== undefined) {
    const base = themeColors[color.theme];
    if (!base) return undefined;
    const withTint = color.tint ? applyTint(base, color.tint) : base;
    return `#${withTint}`;
  }
  if (color.indexed !== undefined) {
    const hex = EXCEL_INDEXED_COLORS[color.indexed];
    return hex ? `#${hex}` : undefined;
  }
  return undefined;
}

const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
/** 숫자·날짜처럼 보이는 글자 — 글자 값인데 이런 모양이면 왼쪽 정렬을 박아 둔다(화면은 모양으로 정렬을 고른다) */
const NUMERIC_LOOKING_TEXT = /^\s*[-+(]?[₩$€£¥]?\s*(\d[\d,]*(\.\d*)?|\.\d+)([eE][-+]?\d+)?\)?%?\s*$|^\d{4}-\d{2}-\d{2}/;

/** 칸 하나의 표시 글자(text)와 값 글자(value — 수식 계산·편집기·저장이 쓰는 원래 값) */
export interface CellTexts {
  readonly text: string;
  readonly value: string;
  /** 서식 코드의 글자 색([Red] 등) */
  readonly color?: string;
  /** 값이 글자(문자열)다 — 엑셀은 숫자처럼 보여도 왼쪽 정렬한다 */
  readonly isText?: true;
  /** 오류값(#DIV/0! 등) — 엑셀은 가운데 정렬한다 */
  readonly isError?: true;
}

/** 계산 결과·캐시 결과·리터럴 값 하나를 칸 서식대로(표시·값 글자). 서식은 numberFormat.ts 하나로만 한다. */
function valueTexts(value: unknown, numFmt: string | undefined): CellTexts {
  if (value === null || value === undefined) return { text: '', value: '' };
  const code = numFmt && numFmt !== 'General' ? numFmt : undefined;
  if (value instanceof Date) {
    const raw = dateToDateText(value);
    return code ? { ...formatNumberCode((value.getTime() - EXCEL_EPOCH_MS) / 86_400_000, code), value: raw } : { text: raw, value: raw };
  }
  // 리터럴 논리값·캐시된 논리값 결과 모두 엑셀처럼 대문자(isSpillEligible의 'TRUE'/'FALSE' 특례가 이 모양을 본다)
  if (typeof value === 'boolean') return value ? { text: 'TRUE', value: 'TRUE' } : { text: 'FALSE', value: 'FALSE' };
  if (typeof value === 'number') {
    // 부동소수 잡음(0.1+0.2) 제거 — 엑셀처럼 유효 숫자 15자리(큰 수에 1e10을 곱했다 나누면 오히려 잡음이 생겼다)
    const n = Number.isFinite(value) ? Number(value.toPrecision(15)) : value;
    // 날짜 형식 칸의 수(=A4+7 같은 계산 결과)는 값도 날짜 글자로 둔다 — 편집기·수식이 날짜로 다룬다
    const raw = Number.isFinite(n) && isDateNumFmt(numFmt) ? serialToDateText(n) : String(n);
    return { ...formatNumberCode(n, code), value: raw };
  }
  if (typeof value === 'object' && typeof (value as { error?: unknown }).error === 'string') {
    // 엑셀 오류값(#DIV/0! 등) — ExcelJS는 { error: '#DIV/0!' } 객체로 준다(예전엔 "[object Object]"로 보였다)
    const err = (value as { error: string }).error;
    return { text: err, value: err, isError: true };
  }
  const text = String(value);
  return { text, value: text, isText: true };
}

/** 수식 칸인지 — 일반 수식과, ExcelJS가 { sharedFormula, result }로 주는 공유 수식 복제 칸 둘 다.
 * 표시 텍스트가 '='로 시작한다는 것만으로는 부족하다: 사용자가 "=abc"라는 **문자열**을 넣어 둔 칸도 있다. */
function isFormulaCellValue(value: ExcelJS.CellValue): boolean {
  return value !== null && typeof value === 'object' && !(value instanceof Date) && ('formula' in value || 'sharedFormula' in value);
}

export function cellTexts(cell: ExcelJS.Cell): CellTexts {
  const value = cell.value;
  if (value === null || value === undefined) return { text: '', value: '' };
  if (typeof value === 'object' && !(value instanceof Date)) {
    const rich = value as ExcelJS.CellRichTextValue;
    if (Array.isArray(rich.richText)) {
      const text = rich.richText.map((run) => run.text).join('');
      return { text, value: text, isText: true };
    }
    const link = value as ExcelJS.CellHyperlinkValue;
    if (typeof link.text === 'string') return { text: link.text, value: link.text, isText: true };
    if (isFormulaCellValue(value)) {
      const result = (value as ExcelJS.CellFormulaValue).result;
      // 캐시된 계산 결과가 없는 경우(우리 export가 저장한 수식은 결과를 캐시하지 않는다) — 수식 원문을 남기고,
      // parseXlsxToGrid의 2차 패스(applyComputedFormulaValues)가 서버에서 계산해 바꿔 넣는다. 계산도 못 하면 원문이
      // 그대로 남아 최소한 "여기 수식이 있다"는 건 알 수 있다.
      if (result === null || result === undefined) {
        const formula = `=${cell.formula}`;
        return { text: formula, value: formula };
      }
      return valueTexts(result, cell.numFmt);
    }
    if (typeof (value as { error?: unknown }).error !== 'string') return { text: '', value: '' };
  }
  return valueTexts(value, cell.numFmt);
}

/** 칸 표시 글자 — cellTexts(cell).text */
export function displayValue(cell: ExcelJS.Cell): string {
  return cellTexts(cell).text;
}

/** 칸의 값 글자 — 표시 글자(v)가 서식 때문에 값과 다르면 n에 원래 값이 있다(수식 계산·편집기 시딩·저장은 이것을 쓴다) */
export function gridCellValue(cell: GridCell): string {
  return cell.n ?? cell.v;
}

/** 표시·값 글자로 그리드 칸을 만든다(값이 표시와 같으면 n은 싣지 않는다 — 파일 크기·메모리) */
function gridCellOf(t: CellTexts, styleRef: number | undefined): GridCell {
  const cell: { v: string; n?: string; s?: number } = { v: t.text };
  if (t.value !== t.text) cell.n = t.value;
  if (styleRef !== undefined) cell.s = styleRef;
  return cell;
}

/** 값에 따라 달라지는 서식을 덧붙인다: 서식 코드의 글자 색(fc), 숫자처럼 보이는 글자 값의 왼쪽 정렬(엑셀 기본) */
function valueStyleRef(palette: StylePalette, styleRef: number | undefined, t: CellTexts): number | undefined {
  const base: GridStyle = styleRef === undefined ? {} : palette.toArray()[styleRef];
  const align = base.a ? undefined : t.isError ? 'c' : t.isText && t.text !== '' && (Boolean(base.fmt || base.nf) || NUMERIC_LOOKING_TEXT.test(t.text)) ? 'l' : undefined;
  if (!t.color && !align) return styleRef;
  return palette.registerVariant(styleRef, t.color, align);
}

/** ExcelJS 칸 → 그리드 칸(표시·값·서식). 값도 서식도 없으면 null — parseXlsxToGrid와 스트리밍 읽기(lazyGridStore)가 함께 쓴다 */
export function gridCellFromExcel(
  cell: ExcelJS.Cell,
  palette: StylePalette,
  themeColors: readonly string[],
  sheetProtected: boolean,
): GridCell | null {
  const t = cellTexts(cell);
  const styleRef = extractStyle(cell, palette, themeColors, sheetProtected);
  if (t.text === '' && styleRef === undefined) return null;
  return gridCellOf(t, valueStyleRef(palette, styleRef, t));
}

/** 수식 계산 패스가 계산해 채워야 하는 칸 하나(= displayValue가 수식 원문을 그대로 남긴 칸). */
interface PendingFormulaCell {
  readonly sheetIdx: number;
  readonly row: number;
  readonly col: number;
  /** 원본 엑셀 서식 코드 — 계산 결과를 일반 숫자 칸과 같은 규칙으로 서식화하기 위해 들고 다닌다. */
  readonly numFmt?: string;
}

/**
 * 워크북을 다 읽은 뒤 도는 2차 패스 — displayValue가 "캐시된 결과가 없어" 수식 원문을 그대로
 * 남겨둔 칸들을 서버에서 직접 계산해 그 결과로 바꿔 넣는다. 읽기전용 뷰어가 동시편집기
 * (collab/room.ts의 recomputeFormulasForPage)와 **같은** 값을 보여주게 하기 위한 것으로, 계산
 * 자체도 같은 함수(xlsx/formulas.ts의 computeGridCellDisplay)를 쓴다.
 *
 * **전체 시트를 다 만든 뒤에 도는 이유**: 수식은 자기 시트의 어느 칸이든(창 밖이든 아래쪽이든)
 * 참조할 수 있으므로, 행을 채우는 도중에 계산하면 아직 읽지 않은 아래쪽 참조가 빈 칸으로 보인다.
 *
 * **memo/tainted를 패스 전체가 공유하는 이유**: 체인형 수식(뒤 칸이 앞 칸을 참조)에서 memo가
 * 없으면 총 계산 횟수가 O(N²)로 폭증한다(computeGridCellDisplay 문서 주석 참고). 그래서 이
 * 함수 호출 하나 = 재계산 패스 하나로 보고 memo/tainted를 **여기서 한 번만** 만든다 — 셀마다·
 * 행마다 새로 만들면 안 된다.
 *
 * **sheetNameIndex도 마찬가지로 여기서 한 번만 만드는 이유**: computeGridCellDisplay가 자신에게
 * sheetNameIndex를 안 넘겨주면 호출마다(=여기 아래 루프의 pending 셀 하나마다) 새로 만든다 —
 * 재귀 안에서는 재사용되지만, 이 루프가 도는 pending.length번의 top-level 호출 각각은 서로 다시
 * 만든다는 뜻이다. memo/tainted와 똑같이 재계산 패스(=이 함수 호출) 하나에 시트 구성은 바뀌지
 * 않으므로 여기서 한 번만 만들어 모든 pending 셀에 걸쳐 재사용한다(buildSheetNameIndex 문서
 * 주석 참고).
 *
 * **계산이 다 끝난 뒤에야 표시용 그리드에 써 넣는 이유(중요)**: 계산 입력은 표시용 그리드 그 자체다. 계산 도중에
 * 결과를 써 넣으면 아직 계산 중인 다른 수식(예: `N49=SUM(B49:M49)`)이 반쯤 바뀐 입력을 보게 된다(memo 캐시 키가 좌표뿐인
 * 것도 패스 동안 입력이 바뀌지 않는다는 불변식 위에서만 안전하다). 결과를 전부 모은 뒤 한 번에 반영한다.
 */
async function applyComputedFormulaValues(
  sheets: readonly GridSheet[],
  sheetRows: ReadonlyArray<Array<Array<GridCell | 0>>>,
  palette: StylePalette,
  pending: readonly PendingFormulaCell[],
): Promise<void> {
  if (pending.length === 0) return;
  // 계산 엔진은 칸의 값 글자(gridCellValue — 서식 전 값)를 읽으므로 표시용 그리드를 그대로 넘긴다
  const inputGrid: GridDocument = { sheets, styles: palette.toArray() };
  const memo = new Map<string, unknown>();
  const tainted = new Set<string>();
  const sheetNameIndex = buildSheetNameIndex(inputGrid);
  const computedTexts: Array<CellTexts | null> = new Array(pending.length).fill(null);
  // 와일드카드(COUNTIF·MATCH·SEARCH …) 작업 예산은 업로드 계산 전체에 하나 — 넘으면 그 칸은 계산 실패(수식 원문)로 남는다(재검토 C1)
  const wildcard: WildcardBudget = { left: WILDCARD_WORK_BUDGET };
  for (let i = 0; i < pending.length; i += 1) {
    const { sheetIdx, row, col, numFmt } = pending[i];
    try {
      computedTexts[i] = valueTexts(
        withWildcardBudget(wildcard, () => computeGridCellDisplay(inputGrid, sheetIdx, row, col, new Set(), memo, tainted, sheetNameIndex)),
        numFmt,
      );
    } catch {
      // 계산이 예외로 끝나면(예: 지원하지 않는 구문으로 인한 깊은 재귀) 빈 문자열로 남겨 아래에서
      // 수식 원문을 그대로 두게 한다 — 업로드 자체가 실패하는 것보다 "여기 수식이 있다"고
      // 보여주는 편이 낫다(이 변경 이전의 동작으로 폴백).
      computedTexts[i] = null;
    }
    if ((i + 1) % YIELD_EVERY_N_FORMULA_CELLS === 0) await yieldToEventLoop();
  }

  for (let i = 0; i < pending.length; i += 1) {
    const t = computedTexts[i];
    // 계산이 실패한 칸(null)은 반영하지 않는다 — 수식 원문이 남아 "여기 수식이 있다"는 건 알 수 있다.
    // 계산 결과가 빈 글자(=IF(D13="","",…))면 엑셀처럼 빈 칸으로 보인다(칸 자체는 서식과 함께 남는다).
    if (!t) continue;
    const { sheetIdx, row, col } = pending[i];
    const cells = sheetRows[sheetIdx]?.[row];
    const cell = cells?.[col];
    // 꼬리 행 트리밍으로 사라진 칸일 수 있으므로 존재 여부를 먼저 본다.
    if (!cells || cell === 0 || cell === undefined) continue;
    cells[col] = gridCellOf(t, valueStyleRef(palette, cell.s, t));
    if ((i + 1) % YIELD_EVERY_N_FORMULA_CELLS === 0) await yieldToEventLoop();
  }
}

/** 칸 폭보다 긴 서식 숫자·날짜는 엑셀처럼 '####'으로 보인다(값 n은 그대로 남긴다). 병합 칸은 폭이 달라 건너뛴다. */
function hashOverflowingNumbers(sheet: GridSheet, styles: readonly GridStyle[]): void {
  const merged = new Set(sheet.merges.map((m) => `${m.r - 1}:${m.c - 1}`));
  sheet.rows.forEach((row, r) => {
    const cells = row as Array<GridCell | 0>;
    for (let c = 0; c < cells.length; c += 1) {
      const cell = cells[c];
      if (cell === 0 || cell.s === undefined || cell.v === '') continue;
      const style = styles[cell.s];
      if (!style.fmt || isTextNumFmt(style.fmt) || style.w) continue;
      const width = sheet.colWidths[c] ?? sheet.defaultColWidth;
      const fontPx = style.fs ?? EXCEL_DEFAULT_FONT_PX;
      // 한글이 없고 모든 글자를 가장 넓게(0.55em) 잡아도 들어가면 볼 것도 없다 — 큰 시트의 거의 모든 칸이 여기서 끝난다
      if (cell.v.length * fontPx * 0.55 + 2 <= width && !/[⺀-￿]/.test(cell.v)) continue;
      if (merged.has(`${r}:${c}`) || !isNumberOrDateValue(gridCellValue(cell))) continue;
      const shown = hashIfTooWide(cell.v, width, fontPx);
      if (shown !== cell.v) cells[c] = { ...cell, v: shown, n: gridCellValue(cell) };
    }
  });
}

export function parseMergeRange(range: string): GridMerge | null {
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

export class StylePalette {
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

  private readonly variants = new Map<string, number | undefined>();
  /** base 서식에 서식 코드 색(fc)·값에 따른 정렬을 더한 서식 — 큰 시트에서 칸마다 JSON을 만들지 않게 조합별로 기억한다 */
  registerVariant(base: number | undefined, fc: string | undefined, align: 'l' | 'c' | undefined): number | undefined {
    const key = `${base}|${fc ?? ''}|${align ?? ''}`;
    if (this.variants.has(key)) return this.variants.get(key);
    const style: { -readonly [K in keyof GridStyle]?: GridStyle[K] } = { ...(base === undefined ? {} : this.styles[base]) };
    if (fc) style.fc = fc;
    if (align) style.a = align;
    const index = this.register(style);
    this.variants.set(key, index);
    return index;
  }
}

const DEFAULT_FONT_SIZE_PT = 11;
const PT_TO_PX = 96 / 72;
const MAX_FONT_PX = 26;

/** 엑셀 테두리 스타일 이름 → 단순화된 굵기(1=가는 선, 2=굵은 선). 스타일이 없으면 undefined. */
function borderWeight(border?: Partial<ExcelJS.Border>): BorderWeight | undefined {
  if (!border?.style) return undefined;
  return border.style === 'medium' || border.style === 'thick' || border.style === 'double' ? 2 : 1;
}

/** 원본 엑셀 서식 코드(예: "#,##0.00", "0.0%", "\"₩\"#,##0")를 그대로 저장하지 않고, 툴바가 다루는
 * 몇 가지 자주 쓰는 범주(쉼표/통화/백분율)+소수 자릿수로 단순화해서 읽는다 — 서식 코드 전체를
 * 파싱하는 건 범위 밖이고, 어차피 우리 툴바도 이 세 가지 버튼만 제공한다.
 *
 * **날짜 서식은 의도적으로 분류하지 않는다**(리뷰에서 지적된 gap — 날짜 칸이 기본 왼쪽 정렬/스필
 * 대상으로 잘못 취급되는 문제 — 는 여기서 NumberFormatKind에 'date'를 추가하는 대신 isSpillEligible
 * 쪽에서 표시 텍스트 모양(YYYY-MM-DD)만으로 처리했다). 이유: NumberFormatKind/style.nf는 이 파일
 * 밖에서도 실제 값 서식(그리드 편집기의 쉼표/통화/백분율 표시 형식 드롭다운 cell-numfmt-select —
 * public/editor-grid.html·src/client/grid-editor.ts, xlsx로 재내보낼 때 서식 코드를 고르는
 * src/xlsx/collab.ts의 numFmtStringFor)를 좌우하는 공유 타입이다. 그 소비처들은 전부 이 세 범주
 * only 기준으로 쓰여 있어(드롭다운엔 애초에 "날짜" 옵션이 없고, numFmtStringFor의 switch도 세
 * 범주만 안다), 'date'를 새 멤버로 넣으면 날짜 칸의 서식 드롭다운이 어떤 옵션도 선택되지 않은 채
 * 빈 채로 뜨는 등 이 정렬/스필 버그와 무관한 곳에서 새로운(더 눈에 띄는) 부작용이 생긴다. 반면 이
 * 앱이 날짜를 실제로 그리는 경로는 formatDate(이 파일) 단 하나뿐이고, 원본 서식 코드가 무엇이든
 * 항상 "YYYY-MM-DD"나 "YYYY-MM-DD HH:MM" 모양으로만 렌더링하므로, 그 모양을 표시 텍스트에서 직접
 * 알아보는 쪽이 공유 타입을 건드리지 않고도 훨씬 좁고 안전하게 같은 문제를 고친다. */
function classifyNumFmt(numFmt?: string): { nf?: NumberFormatKind; nd?: number } {
  if (!numFmt || numFmt === 'General' || numFmt === '@') return {};
  const decimals = numFmt.match(/\.(0+)/)?.[1].length ?? 0;
  if (numFmt.includes('%')) return { nf: 'percent', nd: decimals };
  if (/[$₩¥€]|"[^"]+"/.test(numFmt)) return { nf: 'currency', nd: decimals };
  if (numFmt.includes('#,##0') || /^0(\.0+)?$/.test(numFmt.trim())) return { nf: 'number', nd: decimals };
  return {};
}

export function extractStyle(
  cell: ExcelJS.Cell,
  palette: StylePalette,
  themeColors: readonly string[],
  sheetProtected: boolean,
): number | undefined {
  const style: { -readonly [K in keyof GridStyle]?: GridStyle[K] } = {};
  const font = cell.font;
  if (font?.bold) style.b = 1;
  if (font?.italic) style.i = 1;
  if (font?.underline) style.u = 1;
  if (font?.strike) style.st = 1;
  if (font?.size && font.size !== DEFAULT_FONT_SIZE_PT) {
    style.fs = Math.min(MAX_FONT_PX, Math.round(font.size * PT_TO_PX));
  }
  if (font?.name && font.name !== 'Calibri') style.ff = font.name;
  const fontColor = resolveColor(font?.color, themeColors);
  if (fontColor && fontColor.toUpperCase() !== '#000000') style.col = fontColor;
  const fill = cell.fill;
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const bg = resolveColor(fill.fgColor, themeColors);
    // 흰색이어도 캡처한다 — 잠긴 칸은 기본적으로 회색으로 렌더링되는데, 셀에 명시적으로 흰 배경을
    // 지정해두면(이 커넥션-현황 시트의 1행처럼) 그 지정이 잠금 회색보다 우선해야 한다.
    if (bg) style.bg = bg;
  } else if (fill && fill.type === 'pattern' && fill.pattern && fill.pattern !== 'none') {
    // 무늬 채우기(gray125·darkGrid·lightTrellis …) — 무늬 색(fg)·바탕색(bg). 색이 없으면 엑셀처럼 검정 무늬·흰 바탕
    const fg = resolveColor(fill.fgColor, themeColors);
    const bg = resolveColor(fill.bgColor, themeColors);
    style.pat = bg ? [fill.pattern, fg ?? '#000000', bg] : fg ? [fill.pattern, fg] : [fill.pattern];
  }
  const horizontal = cell.alignment?.horizontal;
  if (horizontal === 'center') style.a = 'c';
  else if (horizontal === 'right') style.a = 'r';
  // 명시적 왼쪽 정렬도 캡처한다 — 안 그러면 이미 소스 xlsx에서 명시적으로 왼쪽 정렬된 칸(예:
  // 숫자 서식이면서 왼쪽 정렬인 칸)이 style.a 미지정으로 저장돼, isSpillEligible(뷰어/편집기
  // 둘 다) 쪽에서 "명시적 style.a==='l'은 항상 넘친다" 분기를 절대 못 타고 style.nf 기반 기본
  // 오른쪽 정렬로 잘못 분류되어 스필이 되레 억제된다(엑셀 실제 동작과 반대인, 옆 빈 칸으로 안
  // 넘치는 방향의 사소한 표시 차이). 'l'이 이제 실제 파일에서도 나올 수 있게 됐지만, 이 값을
  // 읽는 다른 소비처(ALIGN_TO_CSS/ALIGN_TO_EXCEL, 뷰어 injectGridStyles)는 이미 'l'을 다루도록
  // 돼 있어 부작용이 없다(확인 완료).
  else if (horizontal === 'left') style.a = 'l';
  const vertical = cell.alignment?.vertical;
  if (vertical === 'top') style.va = 't';
  else if (vertical === 'middle') style.va = 'm';
  else if (vertical === 'bottom') style.va = 'b';
  if (cell.alignment?.wrapText) style.w = 1;
  const rotation = cell.alignment?.textRotation;
  if (rotation === 'vertical') style.rot = 255;
  else if (typeof rotation === 'number' && rotation !== 0) style.rot = rotation < 0 ? 90 - rotation : rotation; // ExcelJS는 시계 방향을 음수로 준다
  if (cell.alignment?.indent) style.ind = cell.alignment.indent;
  if (cell.alignment?.shrinkToFit) style.shr = 1;
  // 날짜 먼저 — "yyyy\"년\"" 같은 글자 붙은 날짜 서식이 통화로 분류되지 않게
  const numFmt = cell.numFmt;
  const { nf, nd } = isDateNumFmt(numFmt) ? {} : classifyNumFmt(numFmt);
  if (nf) {
    style.nf = nf;
    if (nd) style.nd = nd;
  } else if (isDateNumFmt(numFmt)) {
    style.df = 1;
  }
  if (numFmt && numFmt !== 'General') style.fmt = numFmt;
  // 시트 보호가 걸려 있을 때만 잠금 여부가 실제로 의미를 가진다(엑셀도 그렇게 동작함) — 그렇지
  // 않은 대부분의 일반 문서는 셀마다 protection.locked가 true(엑셀 기본값)라도 편집 가능해야 한다.
  if (sheetProtected && cell.protection?.locked !== false) style.lk = 1;
  // 엑셀의 실제 데이터 유효성 검사(정수/소수 제한)를 그대로 읽는다 — webviewer가 새로 지어낸
  // 개념이 아니라 어떤 xlsx든 이 규칙이 있으면 똑같이 동작한다.
  const validationType = cell.dataValidation?.type;
  if (validationType === 'whole' || validationType === 'decimal') style.nu = 1;
  const border = cell.border;
  const bt = borderWeight(border?.top);
  if (bt) style.bt = bt;
  const br = borderWeight(border?.right);
  if (br) style.br = br;
  const bb = borderWeight(border?.bottom);
  if (bb) style.bb = bb;
  const bl = borderWeight(border?.left);
  if (bl) style.bl = bl;
  // 굵기로 못 나타내는 변(색 있는 선·점선·이중선 …)은 선 종류와 색을 그대로 남긴다
  const bd: { t?: BorderSpec; r?: BorderSpec; b?: BorderSpec; l?: BorderSpec } = {};
  for (const [key, side] of [['t', border?.top], ['r', border?.right], ['b', border?.bottom], ['l', border?.left]] as const) {
    if (!side?.style) continue;
    const color = resolveColor(side.color, themeColors)?.toUpperCase();
    const plainColor = !color || color === '#000000';
    if (side.style !== 'thin' && side.style !== 'medium') bd[key] = plainColor ? [side.style] : [side.style, color];
    else if (!plainColor) bd[key] = [side.style, color];
  }
  if (Object.keys(bd).length > 0) style.bd = bd;
  return palette.register(style);
}

/** 병합 칸 위치 — 'multi'(여러 행에 걸친 병합의 왼쪽 위)·'single'(한 행 안 가로 병합의 왼쪽)·'covered'(덮인 칸) */
export type MergeRole = 'multi' | 'single' | 'covered';

export function mergeRolesByRow(merges: readonly GridMerge[]): Map<number, Map<number, { role: MergeRole; cs: number }>> {
  const byRow = new Map<number, Map<number, { role: MergeRole; cs: number }>>();
  for (const m of merges) {
    for (let r = m.r - 1; r < m.r - 1 + m.rs; r += 1) {
      let row = byRow.get(r);
      if (!row) byRow.set(r, (row = new Map()));
      for (let c = m.c - 1; c < m.c - 1 + m.cs; c += 1) {
        const anchor = r === m.r - 1 && c === m.c - 1;
        row.set(c, { role: anchor ? (m.rs > 1 ? 'multi' : 'single') : 'covered', cs: m.cs });
      }
    }
  }
  return byRow;
}

/**
 * 높이가 파일에 없는 행의 "자동 맞춤" 높이(px) — 엑셀은 이런 행을 줄 바꿈 칸의 줄 수·큰 글자에 맞춰
 * 늘려 보여 준다. 병합 칸은 엑셀도 줄 바꿈 자동 맞춤에서 빼므로(한 행 안 가로 병합은 글자 크기만
 * 반영) 똑같이 한다. 기본보다 높을 때만 값을 돌려준다.
 */
export function autoFitRowPx(
  cells: ReadonlyArray<GridCell | 0>,
  styles: readonly GridStyle[],
  colWidthAt: (col: number) => number,
  mergeRoles: ReadonlyMap<number, { role: MergeRole; cs: number }> | undefined,
  defaultRowPx: number,
): number | undefined {
  let best = defaultRowPx;
  for (let c = 0; c < cells.length; c += 1) {
    const cell = cells[c];
    if (cell === 0 || cell.v === '') continue;
    const merge = mergeRoles?.get(c);
    if (merge && merge.role !== 'single') continue;
    const style = cell.s === undefined ? undefined : styles[cell.s];
    const fontPx = style?.fs ?? EXCEL_DEFAULT_FONT_PX;
    const lines = style?.w && !merge ? estimateWrappedLines(cell.v, colWidthAt(c), fontPx) : 1;
    if (lines < 2 && fontPx <= EXCEL_DEFAULT_FONT_PX + 0.5) continue;
    best = Math.max(best, autoFitRowHeightPx(lines, fontPx, defaultRowPx));
  }
  return best > defaultRowPx ? best : undefined;
}

/** 파일에 적힌 높이(숨김=0) + 나머지 행의 자동 맞춤 높이 → 묶음 */
export function computeRowHeights(
  rows: ReadonlyArray<ReadonlyArray<GridCell | 0>>,
  styles: readonly GridStyle[],
  colWidths: readonly number[],
  merges: readonly GridMerge[],
  explicitRowPx: ReadonlyMap<number, number>,
  defaultRowPx: number,
): RowHeightRuns {
  const roles = mergeRolesByRow(merges);
  const colWidthAt = (c: number) => colWidths[c] ?? 0;
  const entries: Array<[number, number]> = [];
  let lastRow = rows.length;
  for (const r of explicitRowPx.keys()) lastRow = Math.max(lastRow, r + 1);
  for (let r = 0; r < lastRow; r += 1) {
    const explicit = explicitRowPx.get(r);
    if (explicit !== undefined) {
      entries.push([r, explicit]);
      continue;
    }
    const cells = rows[r];
    const fit = cells ? autoFitRowPx(cells, styles, colWidthAt, roles.get(r), defaultRowPx) : undefined;
    if (fit !== undefined) entries.push([r, fit]);
  }
  return encodeRowHeights(entries, defaultRowPx);
}

export async function parseXlsxToGrid(xlsxBytes: Buffer): Promise<GridDocument> {
  const workbook = new ExcelJS.Workbook();
  // 차트가 든 그리기 파트를 ExcelJS가 읽다 죽지 않게 떼고 읽는다(그리드 보기에는 쓰지 않는 파트다)
  await workbook.xlsx.load(stripDrawings(new Uint8Array(xlsxBytes)).bytes as unknown as ArrayBuffer);
  const palette = new StylePalette();
  // exceljs 타입 선언(string[])과 달리 런타임은 { theme1: xml문자열 } 형태다
  const themes = (workbook.model as unknown as { themes?: Record<string, string> }).themes;
  const themeColors = parseThemeColors(themes?.theme1);
  const sheets: GridSheet[] = [];
  // sheets와 1:1로 대응하는 "쓰기 가능한" 행 배열 핸들 — GridSheet.rows는 readonly로 선언돼 있어
  // 수식 계산 패스가 결과를 써 넣을 수 없으므로, 같은 배열 객체를 여기에도 담아 둔다.
  const sheetRows: Array<Array<Array<GridCell | 0>>> = [];
  const pendingFormulas: PendingFormulaCell[] = [];

  // workbook.eachSheet/sheet.eachRow는 콜백형 동기 API라 중간에 await로 이벤트 루프에 양보할 수
  // 없다 — 배열(workbook.worksheets)을 직접 도는 for 루프로 바꿔, 대형 시트를 열 때 그 문서를 연
  // 사용자뿐 아니라 그 순간 서버에 접속한 다른 모든 사용자·다른 문서 요청까지 함께 멈추지 않도록
  // 주기적으로 양보한다.
  // 열 너비 단위의 기준 — 통합 문서 기본 글꼴(표준 스타일)의 숫자 폭
  // 스타일·시트 관계·표 파트(엑셀 표의 표 스타일)만 푼다 — 시트 XML(데이터)은 풀지 않는다
  const parts = unzipSync(new Uint8Array(xlsxBytes), {
    filter: (file) =>
      file.name === 'xl/styles.xml' || file.name === 'xl/workbook.xml' || file.name === 'xl/_rels/workbook.xml.rels' ||
      file.name.startsWith('xl/tables/') || file.name.startsWith('xl/worksheets/_rels/'),
  });
  const stylesPart = parts['xl/styles.xml'];
  const sheetPaths = sheetPathsByName(parts);
  const defaultFont = defaultFontOfStylesXml(stylesPart ? strFromU8(stylesPart) : undefined);
  const mdw = maxDigitWidthPx(defaultFont.name, defaultFont.size);
  for (const sheet of workbook.worksheets) {
    if (sheet.state && sheet.state !== 'visible') continue;
    // sheet.actualRowCount는 ExcelJS 자체 문서상 "실제 값이 있는 행의 '개수'"라, 성긴 시트(예:
    // 1행과 350행에만 값이 있고 그 사이가 전부 비어 있는 시트)에서는 실제 최대 행 번호(350)보다
    // 훨씬 작은 값(예: 2)을 돌려준다 — worksheet.js의 실제 구현(eachRow가 row.hasValues인 행만
    // 세는 것)으로 직접 확인함. 이 값을 rowCount로 그대로 쓰면 뒤쪽의 진짜 데이터가 통째로 잘려
    // 나간다(동시편집기 lazyGridStore.ts의 ingestXlsxLazily에서 이미 한 번 겪고 고친 것과 같은
    // 부류의 버그 — 그쪽은 스트리밍 리더가 실제로 마주친 <row> 엘리먼트의 row.number를 그대로
    // 신뢰하는 방식으로 고쳤다). 여기서는 대신 sheet.rowCount(내부 _lastRowNumber — worksheet.js의
    // model setter가 `this._rows[row.number - 1] = row`로 채워서, 값이 있든 서식만 있든 실제로
    // <row> XML 엘리먼트가 존재했던 가장 큰 행 번호를 그대로 반영한다)를 쓴다. 값 없이 서식만
    // 남은 "유령 행"이 맨 끝에 더 포함되는 과다 계산 걱정은 없다 — 아래 lastWithValue 트리밍이
    // 어차피 그런 꼬리 행을 실제 데이터 기준으로 다시 잘라낸다.
    const rowCount = Math.min(sheet.rowCount || 0, MAX_ROWS);
    if (rowCount === 0) continue;
    // exceljs 타입 선언에는 없지만 런타임에는 시트 보호 시 채워진다(themes와 동일한 이유로 캐스팅)
    const sheetProtected = Boolean((sheet as unknown as { sheetProtection?: unknown }).sheetProtection);

    // actualColumnCount는 맨 오른쪽 열이 수식 셀만 있을 때 그 열을 세지 않는 경우가 있다(ExcelJS
    // 자체 버그로 보임 — 값이 아니라 수식만 있는 열은 "dimensions" 추적에서 빠지는 듯하다). 이 값만
    // 믿으면 예: 맨 오른쪽 합계(=SUM) 열이 통째로 누락되고, 그 결과 동시편집 저장 시
    // exportXlsxFromGrid의 "구조 밖 열 잘라내기" 로직이 그 열을 실제로 지워버린다. 그래서 아래
    // 본문 루프에서 실제로 쓰인 최대 열 번호(maxColUsed)도 함께 구해 더 큰 쪽을 쓴다 — 예전에는
    // 이걸 위해 전체 행을 한 번 더 훑는 별도 사전 스캔이 있었지만, 각 행의 칸 배열을 고정 폭이
    // 아니라 필요할 때마다 늘려가며 채우면(streaming 경로 ingestXlsxLazily와 같은 기법) 본문
    // 루프 한 번으로 maxColUsed까지 같이 구할 수 있어 ExcelJS 행/셀 순회를 두 번 하지 않는다.
    let maxColUsed = 0;
    const rows: Array<Array<GridCell | 0>> = [];
    // 이 시트가 sheets 배열에서 갖게 될 인덱스 — 아래 sheets.push가 이 반복의 마지막이고, 여기까지
    // 왔으면 더 이상 continue로 건너뛰지 않으므로 지금 값이 곧 최종 인덱스다(숨김/빈 시트가
    // 걸러지므로 워크북 상의 시트 번호와는 다를 수 있다).
    const sheetIdx = sheets.length;
    for (let r = 1; r <= rowCount; r += 1) {
      const row = sheet.getRow(r);
      const cells: Array<GridCell | 0> = [];
      // includeEmpty:true여야 한다 — 값은 없지만 스타일(잠금·숫자 전용 유효성 검사·테두리 등)만
      // 있는 칸도 빠짐없이 읽는다. false로 두면 "값이 없는 칸"은 아예 순회 대상에서 빠져, 예를 들어
      // 아직 아무도 입력하지 않은 숫자 전용 칸이 잠금·유효성 검사 정보 없이 렌더링되어 아무 텍스트나
      // 입력·붙여넣기해도 막히지 않는 문제가 있었다(실제로 재현된 버그).
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber > MAX_COLS) return;
        if (colNumber > maxColUsed) maxColUsed = colNumber;
        while (cells.length < colNumber) cells.push(0);
        const gridCell = gridCellFromExcel(cell, palette, themeColors, sheetProtected);
        if (!gridCell) return;
        cells[colNumber - 1] = gridCell;
        // 캐시된 결과가 없어 수식 원문을 그대로 남긴 칸 — 워크북을 다 읽은 뒤
        // applyComputedFormulaValues가 직접 계산해 채운다(cellTexts 문서 주석 참고).
        if (gridCell.v.startsWith('=') && isFormulaCellValue(cell.value)) {
          pendingFormulas.push({ sheetIdx, row: r - 1, col: colNumber - 1, numFmt: cell.numFmt });
        }
      });
      rows.push(cells);
      if (r % YIELD_EVERY_N_ROWS === 0) await yieldToEventLoop();
    }
    const colCount = Math.min(Math.max(sheet.actualColumnCount || 0, maxColUsed), MAX_COLS);
    if (colCount === 0) continue;
    // 위 루프는 각 행을 그 행에서 실제로 마주친 마지막 칸까지만 채우므로(동적 성장), 시트 전체의
    // colCount보다 짧게 끝난 행이 있을 수 있다 — 고정 폭 배열이었던 예전 방식과 동일한 모양으로
    // 맞추기 위해 0으로 패딩한다.
    for (const cells of rows) {
      while (cells.length < colCount) cells.push(0);
    }

    const sheetFormat = sheet.properties as { defaultRowHeight?: number; defaultColWidth?: number };
    const defaultColWidth = defaultColWidthPx(sheetFormat.defaultColWidth, 8, mdw);
    const defaultRowHeight = rowHeightPtToPx(sheetFormat.defaultRowHeight || EXCEL_DEFAULT_ROW_HEIGHT_PT);
    const colWidths: number[] = [];
    for (let c = 1; c <= colCount; c += 1) {
      const column = sheet.getColumn(c);
      colWidths.push(column.hidden ? 0 : column.width ? excelColWidthToPx(column.width, mdw) : defaultColWidth);
    }
    // 값이 없는 오른쪽 열도 너비를 따로 정해 두었으면(양식 문서가 흔히 그렇다) 그 열까지 너비를 싣는다
    const extraWidths: number[] = [];
    for (let c = colCount + 1; c <= MAX_COLS; c += 1) {
      const column = sheet.getColumn(c);
      extraWidths.push(column.hidden ? 0 : column.width ? excelColWidthToPx(column.width, mdw) : defaultColWidth);
    }
    while (extraWidths.length > 0 && extraWidths[extraWidths.length - 1] === defaultColWidth) extraWidths.pop();
    colWidths.push(...extraWidths);
    // 파일에 적힌 행 높이·숨김 — 값이 없는 꼬리 행도(서식만 있는 행) 화면 여백 행으로 그려지므로 다 읽는다
    const explicitRowPx = new Map<number, number>();
    for (let r = 1; r <= rowCount; r += 1) {
      const row = sheet.getRow(r);
      if (row.hidden) explicitRowPx.set(r - 1, 0);
      else if (row.height) explicitRowPx.set(r - 1, rowHeightPtToPx(row.height));
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

    // 엑셀 표의 표 스타일(머리 줄 색·줄무늬·요약 행) — 칸에 직접 준 서식이 이긴다
    const sheetPath = sheetPaths.get(sheet.name);
    for (const table of sheetPath ? sheetTablesOf(parts, sheetPath) : []) {
      for (let r = table.r0; r <= Math.min(table.r1, rows.length - 1); r += 1) {
        const cells = rows[r];
        for (let c = table.c0; c <= Math.min(table.c1, colCount - 1); c += 1) {
          const look = tableCellLook(table, r, c, themeColors);
          if (!look) continue;
          const cell = cells[c];
          const base: GridStyle = cell === 0 || cell.s === undefined ? {} : palette.toArray()[cell.s];
          const s = palette.register(withTableLook(base, look));
          cells[c] = cell === 0 ? (s === undefined ? 0 : { v: '', s }) : s === undefined ? { v: cell.v, ...(cell.n !== undefined ? { n: cell.n } : {}) } : { ...cell, s };
        }
      }
    }

    const merges = (
      (sheet.model as { merges?: readonly string[] }).merges ?? []
    )
      .map(parseMergeRange)
      .filter(
        (merge): merge is GridMerge => merge !== null && merge.r <= rows.length && merge.c <= colCount,
      );

    const frozenView = sheet.views?.find((v) => (v as { state?: string }).state === 'frozen');
    const freezeRows = frozenView ? ((frozenView as { ySplit?: number }).ySplit ?? 0) : 0;
    const freezeCols = frozenView ? ((frozenView as { xSplit?: number }).xSplit ?? 0) : 0;

    const rowHeights = computeRowHeights(rows, palette.toArray(), colWidths, merges, explicitRowPx, defaultRowHeight);
    const showGridLines = (sheet.views?.[0] as { showGridLines?: boolean } | undefined)?.showGridLines !== false;
    sheets.push({
      name: sheet.name,
      rowCount: rows.length,
      colCount,
      colWidths,
      defaultColWidth,
      defaultRowHeight,
      rowHeights,
      showGridLines,
      merges,
      freezeRows,
      freezeCols,
      rows,
    });
    sheetRows.push(rows);
  }

  // 2차 패스 — 시트를 전부 만든 뒤에야 수식이 참조하는 칸이 모두 채워져 있다.
  await applyComputedFormulaValues(sheets, sheetRows, palette, pendingFormulas);
  for (const sheet of sheets) hashOverflowingNumbers(sheet, palette.toArray());
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
        defaultColWidth: defaultColWidthPx(),
        defaultRowHeight: rowHeightPtToPx(EXCEL_DEFAULT_ROW_HEIGHT_PT),
        rowHeights: [],
        showGridLines: true,
        merges: [],
        freezeRows: 0,
        freezeCols: 0,
        rows: gridRows,
      },
    ],
    styles: [],
  };
}
