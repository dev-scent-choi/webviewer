import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';
import { dxfConverter } from './libreOffice.js';
import type { Converter } from './types.js';

const execFileAsync = promisify(execFile);

const DWG2DXF_TIMEOUT_MS = 120_000;
const STDERR_BUFFER_BYTES = 32 * 1024 * 1024;

function detectDwg2dxf(): string | null {
  const fromEnv = process.env.DOCVIEW_DWG2DXF;
  if (fromEnv) {
    return existsSync(fromEnv) ? fromEnv : null;
  }
  const bundled = path.resolve(
    path.dirname(fileURLToPath(import.meta.url)),
    '..',
    '..',
    'tools',
    process.platform === 'win32' ? 'dwg2dxf.exe' : 'dwg2dxf',
  );
  return existsSync(bundled) ? bundled : null;
}

/** null이면 LibreDWG(dwg2dxf) 미설치 → DWG 변환기가 등록되지 않는다. */
export const dwg2dxfPath = detectDwg2dxf();

/**
 * DWG(오토데스크 독점 포맷) → DXF는 LibreDWG로, DXF → PDF는 기존
 * LibreOffice Draw 경로로 변환한다. 최신 DWG 전용 객체 일부는 LibreDWG가
 * 경고와 함께 건너뛸 수 있다(비시각적 메타데이터 위주).
 */
export const dwgConverter: Converter = {
  name: 'libredwg',
  extensions: ['dwg'],
  async convert(sourcePath, targetPdfPath, sourceName) {
    if (!dwg2dxfPath) {
      throw new Error('DWG 변환기(dwg2dxf)를 찾을 수 없습니다.');
    }
    const workDir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-dwg-'));
    try {
      const dxfPath = path.join(workDir, 'converted.dxf');
      // dwg2dxf는 정상 변환 시에도 경고를 stderr로 출력하므로 산출물 존재로 성공을 판정한다
      await execFileAsync(dwg2dxfPath, ['-y', '-o', dxfPath, sourcePath], {
        timeout: DWG2DXF_TIMEOUT_MS,
        maxBuffer: STDERR_BUFFER_BYTES,
      });
      try {
        await fs.access(dxfPath);
      } catch {
        throw new Error('DWG→DXF 변환 결과가 생성되지 않았습니다.');
      }
      const displayName = (sourceName ?? 'drawing.dwg').replace(/\.dwg$/i, '.dxf');
      await dxfConverter.convert(dxfPath, targetPdfPath, displayName);
    } finally {
      await fs.rm(workDir, { recursive: true, force: true });
    }
  },
};
