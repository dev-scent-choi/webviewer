import { Readable } from 'node:stream';
import zlib from 'node:zlib';

/**
 * 차트·그림·링크·목록 상자 읽기(drawingPassthrough·sheetExtras)용 xlsx 파트 모음 — 큰 시트의 칸 데이터(<sheetData>)는
 * 빼고 읽는다.
 *
 * 예전엔 이 셋이 각자 파일 전체를 unzipSync로 풀고 수백 MB짜리 시트 XML을 문자열로 바꿨다. 56MB 파일에서 요청마다
 * 9초씩 서버를 멈춰 세워 편집기 첫 데이터가 그만큼 늦었고, 메모리도 크게 튀었다. 그런데 필요한 것(그리기 연결,
 * 하이퍼링크, 데이터 유효성)은 전부 <sheetData> 바깥에 있다 — 큰 시트는 압축을 흘려 풀며 데이터 부분을 건너뛰고,
 * 그 안에서는 =HYPERLINK(...) 수식이 든 칸만 남긴다(링크 목록이 그 수식도 링크로 쓴다).
 */

/** 이보다 큰 시트 XML만 데이터 부분을 건너뛴다 — 작은 파일은 예전과 똑같이 통째로 읽는다 */
export const SKELETON_SHEET_BYTES = 8 * 1024 * 1024;
const WORKSHEET_PART = /^xl\/worksheets\/[^/]+\.xml$/;
/** 세 읽기 함수 모두 공유 문자열을 쓰지 않는다 — 크면 아예 풀지 않는다 */
const UNUSED_LARGE_PART = /^xl\/sharedStrings\.xml$/;

interface ZipEntry {
  readonly name: string;
  readonly method: number;
  readonly compressedSize: number;
  readonly size: number;
  readonly localOffset: number;
}

function listZipEntries(data: Uint8Array): ZipEntry[] {
  if (data.length < 22) throw new Error('zip이 아닙니다');
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const u16 = (o: number) => view.getUint16(o, true);
  const u32 = (o: number) => view.getUint32(o, true);
  const u64 = (o: number) => u32(o) + u32(o + 4) * 2 ** 32;
  let eocd = -1;
  for (let i = data.length - 22; i >= Math.max(0, data.length - 22 - 0xffff); i -= 1) {
    if (u32(i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error('zip 끝 레코드가 없습니다');
  let count = u16(eocd + 10);
  let offset = u32(eocd + 16);
  if ((count === 0xffff || offset === 0xffffffff) && eocd >= 20 && u32(eocd - 20) === 0x07064b50) {
    const z64 = u64(eocd - 12);
    if (z64 + 56 > data.length || u32(z64) !== 0x06064b50) throw new Error('zip64 레코드가 깨졌습니다');
    count = u64(z64 + 32);
    offset = u64(z64 + 48);
  }
  const decoder = new TextDecoder();
  const entries: ZipEntry[] = [];
  let p = offset;
  for (let n = 0; n < count; n += 1) {
    if (p + 46 > data.length || u32(p) !== 0x02014b50) throw new Error('zip 목록이 깨졌습니다');
    const method = u16(p + 10);
    let compressedSize = u32(p + 20);
    let size = u32(p + 24);
    const nameLen = u16(p + 28);
    const extraLen = u16(p + 30);
    const commentLen = u16(p + 32);
    let localOffset = u32(p + 42);
    const name = decoder.decode(data.subarray(p + 46, p + 46 + nameLen));
    // zip64: 0xffffffff로 적힌 값만 확장 필드(0x0001)에 크기·압축 크기·위치 순서로 들어 있다
    let e = p + 46 + nameLen;
    const extraEnd = e + extraLen;
    while (e + 4 <= extraEnd) {
      const id = u16(e);
      const len = u16(e + 2);
      if (id === 0x0001) {
        let q = e + 4;
        if (size === 0xffffffff) { size = u64(q); q += 8; }
        if (compressedSize === 0xffffffff) { compressedSize = u64(q); q += 8; }
        if (localOffset === 0xffffffff) localOffset = u64(q);
      }
      e += 4 + len;
    }
    entries.push({ name, method, compressedSize, size, localOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

function compressedBytesOf(data: Uint8Array, entry: ZipEntry): Uint8Array {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const lo = entry.localOffset;
  if (lo + 30 > data.length || view.getUint32(lo, true) !== 0x04034b50) throw new Error(`zip 파트 머리가 깨졌습니다: ${entry.name}`);
  const start = lo + 30 + view.getUint16(lo + 26, true) + view.getUint16(lo + 28, true);
  return data.subarray(start, start + entry.compressedSize);
}

function inflateEntry(data: Uint8Array, entry: ZipEntry): Uint8Array {
  const raw = compressedBytesOf(data, entry);
  if (entry.method === 0) return raw;
  if (entry.method === 8) return zlib.inflateRawSync(raw);
  throw new Error(`지원하지 않는 압축 방식(${entry.method}): ${entry.name}`);
}

const OPEN = Buffer.from('<sheetData');
const CLOSE = Buffer.from('</sheetData>');
const HYPERLINK = Buffer.from('HYPERLINK(');
/** 다음 조각과 이어 붙여 찾을 때 넘겨 두는 꼬리 — 칸 하나(<c …>…</c>)가 이보다 길면 그 칸의 HYPERLINK는 놓친다 */
const CARRY = 4096;

/**
 * 시트 XML(압축 해제 흐름)을 받아 <sheetData> 안을 비운 뼈대를 만든다 — 앞부분(열 너비·틀 고정 등)과 뒷부분
 * (병합·하이퍼링크·유효성·그리기 연결)은 그대로, 데이터 안에서는 HYPERLINK( 를 담은 <c> 원소만 남긴다.
 */
export async function skeletonizeSheetXml(source: AsyncIterable<Buffer | Uint8Array>): Promise<Buffer> {
  const head: Buffer[] = [];
  const kept: Buffer[] = [];
  const tail: Buffer[] = [];
  let mode: 'head' | 'skip' | 'tail' = 'head';
  let skipped = false;
  let pending: Buffer = Buffer.alloc(0);
  for await (const piece of source) {
    const chunk = Buffer.isBuffer(piece) ? piece : Buffer.from(piece);
    if (mode === 'tail') {
      tail.push(chunk);
      continue;
    }
    let buf = pending.length > 0 ? Buffer.concat([pending, chunk]) : chunk;
    pending = Buffer.alloc(0);
    if (mode === 'head') {
      const open = buf.indexOf(OPEN);
      if (open < 0) {
        const keep = Math.max(0, buf.length - OPEN.length);
        head.push(buf.subarray(0, keep));
        pending = Buffer.from(buf.subarray(keep));
        continue;
      }
      const gt = buf.indexOf(0x3e /* > */, open);
      if (gt < 0) {
        head.push(buf.subarray(0, open));
        pending = Buffer.from(buf.subarray(open));
        continue;
      }
      if (buf[gt - 1] === 0x2f /* / */) {
        // <sheetData/> — 빈 시트, 건너뛸 게 없다
        head.push(buf.subarray(0, gt + 1));
        mode = 'tail';
        tail.push(buf.subarray(gt + 1));
        continue;
      }
      head.push(buf.subarray(0, open));
      mode = 'skip';
      skipped = true;
      buf = buf.subarray(gt + 1);
    }
    // mode === 'skip'
    const close = buf.indexOf(CLOSE);
    const searchEnd = close >= 0 ? close : buf.length;
    let pos = 0;
    let carryFrom = -1;
    for (;;) {
      const h = buf.indexOf(HYPERLINK, pos);
      if (h < 0 || h >= searchEnd) break;
      const cellStart = buf.lastIndexOf('<c ', h);
      const cellEnd = buf.indexOf('</c>', h);
      if (cellEnd < 0 || cellEnd > searchEnd) {
        if (close < 0) carryFrom = cellStart >= pos ? cellStart : h; // 칸이 다음 조각까지 이어진다
        break;
      }
      if (cellStart >= pos) kept.push(Buffer.from(buf.subarray(cellStart, cellEnd + 4)));
      pos = cellEnd + 4;
    }
    if (close >= 0) {
      mode = 'tail';
      tail.push(buf.subarray(close + CLOSE.length));
      continue;
    }
    const from = carryFrom >= 0 ? carryFrom : Math.max(pos, buf.length - CARRY);
    pending = Buffer.from(buf.subarray(from));
  }
  if (mode === 'head') head.push(pending);
  const middle = skipped ? [Buffer.from('<sheetData>'), ...kept, Buffer.from('</sheetData>')] : [];
  return Buffer.concat([...head, ...middle, ...tail]);
}

/** 위 설명대로 읽은 파트 모음(unzipSync 결과와 같은 모양) */
export async function readXlsxSkeleton(data: Uint8Array): Promise<Record<string, Uint8Array>> {
  const out: Record<string, Uint8Array> = {};
  for (const entry of listZipEntries(data)) {
    if (entry.name.endsWith('/')) continue;
    const large = entry.size > SKELETON_SHEET_BYTES;
    if (large && UNUSED_LARGE_PART.test(entry.name)) continue;
    if (large && WORKSHEET_PART.test(entry.name) && (entry.method === 8 || entry.method === 0)) {
      const raw = compressedBytesOf(data, entry);
      const stream = entry.method === 8 ? Readable.from(chunksOf(raw)).pipe(zlib.createInflateRaw()) : Readable.from(chunksOf(raw));
      out[entry.name] = await skeletonizeSheetXml(stream);
      continue;
    }
    out[entry.name] = inflateEntry(data, entry);
  }
  return out;
}

function* chunksOf(bytes: Uint8Array, size = 256 * 1024): Generator<Buffer> {
  for (let i = 0; i < bytes.length; i += size) yield Buffer.from(bytes.buffer, bytes.byteOffset + i, Math.min(size, bytes.length - i));
}
