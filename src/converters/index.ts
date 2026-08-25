import { dwg2dxfPath, dwgConverter } from './dwgConverter.js';
import { imageConverter } from './imageConverter.js';
import {
  dxfConverter,
  gridPrintConverter,
  officeConverter,
  sofficePath,
  spreadsheetConverter,
} from './libreOffice.js';
import { pdfPassthroughConverter } from './pdfPassthrough.js';
import { rasterConverter } from './rasterConverter.js';
import type { Converter } from './types.js';

const converters: readonly Converter[] = [
  pdfPassthroughConverter,
  imageConverter,
  rasterConverter,
  // xlsx/csv 그리드 열람은 LibreOffice 없이 동작하므로 항상 등록한다.
  // (출력용 PDF 지연 생성만 LibreOffice가 필요하며, 그 시점에 없으면 별도로 실패한다)
  gridPrintConverter,
  ...(sofficePath ? [officeConverter, spreadsheetConverter, dxfConverter] : []),
  ...(sofficePath && dwg2dxfPath ? [dwgConverter] : []),
];

const byExtension = new Map<string, Converter>(
  converters.flatMap((converter) => converter.extensions.map((ext) => [ext, converter] as const)),
);

export function findConverter(ext: string): Converter | null {
  return byExtension.get(ext.toLowerCase()) ?? null;
}

export const supportedExtensions: readonly string[] = [...byExtension.keys()];
