import { randomBytes } from 'node:crypto';
import { PDFDocument } from 'pdf-lib';

/** fastify.inject용 multipart 본문을 직접 조립한다. */
export function multipartPayload(
  filename: string,
  contentType: string,
  data: Buffer,
): { payload: Buffer; headers: Record<string, string> } {
  const boundary = `----docviewtest${randomBytes(8).toString('hex')}`;
  const head = Buffer.from(
    `--${boundary}\r\n` +
      `Content-Disposition: form-data; name="file"; filename="${filename}"\r\n` +
      `Content-Type: ${contentType}\r\n\r\n`,
  );
  const tail = Buffer.from(`\r\n--${boundary}--\r\n`);
  return {
    payload: Buffer.concat([head, data, tail]),
    headers: { 'content-type': `multipart/form-data; boundary=${boundary}` },
  };
}

export async function makeSamplePdf(pages = 2): Promise<Buffer> {
  const doc = await PDFDocument.create();
  for (let i = 0; i < pages; i += 1) {
    doc.addPage([595, 842]);
  }
  return Buffer.from(await doc.save());
}

/** 1x1 픽셀 유효 PNG */
export function makeSamplePng(): Buffer {
  return Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64',
  );
}
