import { dwg2dxfPath, dwgConverter } from './dwgConverter.js';
import { imageConverter } from './imageConverter.js';
import { dxfConverter, officeConverter, sofficePath, spreadsheetConverter } from './libreOffice.js';
import { pdfPassthroughConverter } from './pdfPassthrough.js';
import { rasterConverter } from './rasterConverter.js';
import type { Converter } from './types.js';

const converters: readonly Converter[] = [
  pdfPassthroughConverter,
  imageConverter,
  rasterConverter,
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
