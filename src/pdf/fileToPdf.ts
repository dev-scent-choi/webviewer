/**
 * 다른 파일 → PDF(7단계) — 업로드 때 쓰는 변환기(LibreOffice·그림 등, src/converters)를 그대로 다시 쓴다.
 * PDF 편집기의 "파일에서 넣기·파일 합치기"(워드·엑셀·PPT·한글·그림도 넣기)와 "PDF로 바꿔 내려받기"가 쓴다.
 */
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { findConverter } from '../converters/index.js';
import { ConversionError } from '../converters/types.js';

export class FileToPdfError extends Error {
  constructor(
    readonly status: number,
    message: string,
  ) {
    super(message);
  }
}

export function extOf(name: string): string {
  return (/\.([A-Za-z0-9]{1,10})$/.exec(name)?.[1] ?? '').toLowerCase();
}

export async function fileToPdf(bytes: Uint8Array, name: string): Promise<Uint8Array> {
  const ext = extOf(name);
  if (ext === 'pdf') return bytes;
  const converter = ext ? findConverter(ext) : null;
  if (!converter) throw new FileToPdfError(415, `PDF로 바꿀 수 없는 형식입니다${ext ? `(.${ext})` : ''}.`);
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'docview-topdf-'));
  try {
    const source = path.join(dir, `source.${ext}`);
    const target = path.join(dir, 'out.pdf');
    await fs.writeFile(source, bytes);
    try {
      await converter.convert(source, target, name);
    } catch (error) {
      throw new FileToPdfError(422, error instanceof ConversionError ? error.message : `${name}을(를) PDF로 바꾸지 못했습니다.`);
    }
    return new Uint8Array(await fs.readFile(target));
  } finally {
    await fs.rm(dir, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 }).catch(() => undefined);
  }
}
