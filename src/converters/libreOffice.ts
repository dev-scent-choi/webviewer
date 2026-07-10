import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { promisify } from 'node:util';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { trimPdfMargins } from './pdfTrim.js';
import type { Converter } from './types.js';

const execFileAsync = promisify(execFile);

const CONVERT_TIMEOUT_MS = 180_000;

const CANDIDATE_PATHS = [
  'C:\\Program Files\\LibreOffice\\program\\soffice.exe',
  'C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe',
  '/usr/bin/soffice',
  '/usr/local/bin/soffice',
];

function detectSoffice(): string | null {
  const fromEnv = process.env.DOCVIEW_SOFFICE;
  if (fromEnv) {
    return existsSync(fromEnv) ? fromEnv : null;
  }
  return CANDIDATE_PATHS.find((candidate) => existsSync(candidate)) ?? null;
}

/** null이면 LibreOffice 미설치 → Office/DXF 변환기가 등록되지 않는다. */
export const sofficePath = detectSoffice();

/** 머리글 등에 노출되는 표시용 파일명 — 경로 구분자와 예약 문자를 제거한다. */
function toSafeDisplayName(sourceName: string | undefined, ext: string): string {
  const base = (sourceName ?? '').replace(/\.[^.]+$/, '');
  const forbidden = new Set(['\\', '/', ':', '*', '?', '"', '<', '>', '|']);
  const safe = [...base]
    .filter((ch) => !forbidden.has(ch) && ch.charCodeAt(0) >= 0x20)
    .join('')
    .trim()
    .slice(0, 80);
  return `${safe.length > 0 ? safe : 'document'}.${ext}`;
}

/**
 * xlsx 안의 고정 행 높이(ht/customHeight)를 제거한다.
 * Excel과 LibreOffice의 글꼴 폭 계산 차이로, Excel에서 꼭 맞게 저장된 행 높이가
 * LibreOffice에서는 부족해 줄바꿈 텍스트가 잘린다. 높이를 제거하면 Calc가
 * 내용에 맞게 행 높이를 자동 계산하므로 텍스트가 잘리지 않는다.
 */
export function stripCustomRowHeights(xlsxBytes: Uint8Array): Uint8Array {
  const entries = unzipSync(xlsxBytes);
  const updated: Record<string, Uint8Array> = {};
  for (const [name, bytes] of Object.entries(entries)) {
    if (/^xl\/worksheets\/[^/]+\.xml$/.test(name)) {
      const xml = strFromU8(bytes);
      const fixed = xml.replace(/<row\b[^>]*?>/g, (tag) =>
        tag.replace(/ ht="[^"]*"/g, '').replace(/ customHeight="(?:1|true)"/g, ''),
      );
      updated[name] = strToU8(fixed);
    } else {
      updated[name] = bytes;
    }
  }
  return zipSync(updated);
}

type SourcePreprocessor = (bytes: Uint8Array) => Uint8Array;

async function convertWithSoffice(
  sourcePath: string,
  targetPdfPath: string,
  extraArgs: readonly string[],
  sourceName?: string,
  preprocess?: SourcePreprocessor,
): Promise<void> {
  if (!sofficePath) {
    throw new Error('LibreOffice(soffice)를 찾을 수 없습니다.');
  }
  // 호출마다 고유 프로필/작업 디렉터리를 써서 동시 실행 충돌과 사용자 세션 간섭을 피한다.
  const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-lo-'));
  try {
    // LibreOffice가 문서 머리글에 UUID 대신 원래 파일명을 쓰도록 표시용 이름으로 복사해 변환한다.
    const ext = path.extname(sourcePath).slice(1).toLowerCase();
    const displayName = toSafeDisplayName(sourceName, ext);
    const stagedSource = path.join(workDir, displayName);
    if (preprocess) {
      const original = await fs.readFile(sourcePath);
      let staged: Uint8Array;
      try {
        staged = preprocess(original);
      } catch {
        // 전처리에 실패해도 변환 자체는 원본으로 계속한다.
        staged = original;
      }
      await fs.writeFile(stagedSource, staged);
    } else {
      await fs.copyFile(sourcePath, stagedSource);
    }

    await execFileAsync(
      sofficePath,
      [
        '--headless',
        '--norestore',
        `-env:UserInstallation=${pathToFileURL(path.join(workDir, 'profile')).href}`,
        ...extraArgs,
        '--outdir',
        workDir,
        stagedSource,
      ],
      { timeout: CONVERT_TIMEOUT_MS },
    );
    // soffice는 변환 실패 시에도 종료 코드 0을 반환할 수 있어 산출물 존재로 판정한다.
    const produced = path.join(workDir, `${path.parse(displayName).name}.pdf`);
    try {
      await fs.access(produced);
    } catch {
      throw new Error(`LibreOffice가 PDF를 생성하지 못했습니다: ${displayName}`);
    }
    await fs.copyFile(produced, targetPdfPath);
  } finally {
    await fs.rm(workDir, { recursive: true, force: true });
  }
}

export const officeConverter: Converter = {
  name: 'libreoffice',
  extensions: ['doc', 'docx', 'ppt', 'pptx', 'odt', 'odp', 'rtf', 'txt'],
  async convert(sourcePath, targetPdfPath, sourceName) {
    await convertWithSoffice(sourcePath, targetPdfPath, ['--convert-to', 'pdf'], sourceName);
  },
};

/**
 * 스프레드시트는 기본 변환 시 A4 인쇄 폭으로 시트가 조각나 화면이 깨져 보인다.
 * SinglePageSheets 옵션으로 "시트 1장 = PDF 1페이지(내용 크기)"로 내보낸다. (LibreOffice 7.2+)
 */
const SPREADSHEET_PDF_FILTER =
  'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":"true"}}';

export const spreadsheetConverter: Converter = {
  name: 'libreoffice-calc',
  extensions: ['xls', 'xlsx', 'ods', 'csv'],
  async convert(sourcePath, targetPdfPath, sourceName) {
    const isXlsx = path.extname(sourcePath).toLowerCase() === '.xlsx';
    await convertWithSoffice(
      sourcePath,
      targetPdfPath,
      ['--convert-to', SPREADSHEET_PDF_FILTER],
      sourceName,
      isXlsx ? stripCustomRowHeights : undefined,
    );
  },
};

export const dxfConverter: Converter = {
  name: 'libreoffice-dxf',
  extensions: ['dxf'],
  async convert(sourcePath, targetPdfPath, sourceName) {
    await convertWithSoffice(
      sourcePath,
      targetPdfPath,
      ['--infilter=DXF', '--convert-to', 'pdf:draw_pdf_Export'],
      sourceName,
    );
    // 도면이 거대한 빈 페이지에 작게 놓이는 문제를 내용 영역 크롭으로 보정한다
    await trimPdfMargins(targetPdfPath);
  },
};
