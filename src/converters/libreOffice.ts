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

/**
 * 프로필 풀: 매 변환마다 새 프로필을 만들면 LibreOffice가 프로필 초기화에
 * 수 초를 쓰므로, 웜업된 프로필을 슬롯 단위로 재사용한다. 슬롯별로 직렬화되어
 * 같은 프로필을 두 프로세스가 동시에 쓰는 일은 없다.
 */
const PROFILE_POOL_SIZE = 2;
const profileRoot = path.join(os.tmpdir(), 'docview-lo-profiles');
const slotQueues: Array<Promise<unknown>> = Array.from({ length: PROFILE_POOL_SIZE }, () =>
  Promise.resolve(),
);
let nextSlot = 0;

async function withProfileSlot<T>(task: (profileDir: string) => Promise<T>): Promise<T> {
  const slot = nextSlot;
  nextSlot = (nextSlot + 1) % PROFILE_POOL_SIZE;
  const previous = slotQueues[slot];
  let release!: () => void;
  const current = new Promise<void>((resolve) => {
    release = resolve;
  });
  slotQueues[slot] = previous.then(() => current);
  await previous.catch(() => {});
  const profileDir = path.join(profileRoot, `slot-${slot}`);
  try {
    await fs.mkdir(profileDir, { recursive: true });
    return await task(profileDir);
  } catch (error) {
    // 변환 실패가 프로필 손상 때문일 수 있으므로 다음 실행에서 새로 만들도록 비운다.
    await fs.rm(profileDir, { recursive: true, force: true }).catch(() => {});
    throw error;
  } finally {
    release();
  }
}

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

    await withProfileSlot(async (profileDir) => {
      await execFileAsync(
        sofficePath,
        [
          '--headless',
          '--norestore',
          `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
          ...extraArgs,
          '--outdir',
          workDir,
          stagedSource,
        ],
        { timeout: CONVERT_TIMEOUT_MS },
      );
    });
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

/**
 * 서버 시작 시 프로필 풀을 미리 웜업해 첫 변환의 콜드 스타트(수 초)를 없앤다.
 * 각 슬롯에서 초소형 문서를 한 번씩 변환해 LibreOffice 프로필을 생성해 둔다.
 */
export async function warmUpProfiles(): Promise<void> {
  if (!sofficePath) return;
  await Promise.all(
    Array.from({ length: PROFILE_POOL_SIZE }, async (_, index) => {
      const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-warmup-'));
      try {
        const seed = path.join(workDir, `warmup-${index}.csv`);
        await fs.writeFile(seed, 'warmup\n', 'utf8');
        await convertWithSoffice(seed, path.join(workDir, 'warmup.pdf'), ['--convert-to', 'pdf']);
      } catch {
        /* 웜업 실패는 무시 — 실제 변환 시 콜드 스타트로 대체된다 */
      } finally {
        await fs.rm(workDir, { recursive: true, force: true });
      }
    }),
  );
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
    // 도면이 거대한 빈 페이지에 작게 놓이는 문제를 내용 영역 크롭으로 보정한다.
    // 크롭은 표시 품질 개선일 뿐이므로 실패해도 변환 자체를 실패시키지 않는다.
    try {
      await trimPdfMargins(targetPdfPath);
    } catch (error) {
      console.warn('도면 여백 크롭 실패 (원본 페이지 그대로 표시):', error);
    }
  },
};
