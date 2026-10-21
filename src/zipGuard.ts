import { unzipSync, type UnzipOptions, type Unzipped } from 'fflate';

/**
 * docx/xlsx/pptx/hwp는 전부 zip이고, 업로드된 원본 파일은 열람·편집·재변환될 때마다 이 프로젝트
 * 곳곳(docx/pptx import·export, room.ts의 지연 시딩, viewer.ts의 슬라이드 미디어 서빙 등)에서
 * fflate의 unzipSync로 매번 새로 압축 해제된다 — "zip bomb"(작은 압축 파일이 압축을 풀면 그보다
 * 수백~수천 배 큰 바이트로 부풀어 메모리/CPU를 고갈시키는 공격)에 대한 별도 방어가 없었다.
 *
 * fflate의 동기 API는 압축 해제 도중에 중단할 방법이 없어(스트리밍 API로 바꾸면 이 프로젝트의
 * 십수 개 호출부를 전부 다시 짜야 해 그 자체가 회귀 위험이 크다) "다 풀고 나서 그 결과가 너무
 * 크면 버린다"는 절충을 택한다 — 첫 번째 압축 해제 자체의 CPU 비용은 못 막지만, 그 결과가
 * 쓰이거나(그리드로 만들어지거나 디스크에 다시 쓰이는 등) 문서가 "정상" 상태로 등록되어 열릴
 * 때마다 반복 공격당하는 것은 막는다. 실사용 문서(이미지가 많은 pptx 등)도 넉넉히 통과하도록
 * 상한을 크게 잡는다 — 이 자체가 무결성 검사가 아니라 극단값만 걸러내는 안전망이라는 뜻이다.
 */
export const DEFAULT_MAX_UNCOMPRESSED_BYTES = 1024 * 1024 * 1024; // 1GB

export class ZipTooLargeError extends Error {
  constructor(contextLabel: string, totalBytes: number, maxBytes: number) {
    super(
      `${contextLabel}: 압축 해제 결과가 너무 큽니다(${totalBytes.toLocaleString()} bytes > ` +
        `${maxBytes.toLocaleString()} bytes 상한) — 손상되었거나 zip bomb일 수 있습니다.`,
    );
    this.name = 'ZipTooLargeError';
  }
}

/** unzipSync와 시그니처가 같다(contextLabel·maxBytes만 추가) — 기존 호출부는 import와 호출부에
 * contextLabel만 더하면 그대로 쓸 수 있다. contextLabel은 에러 메시지·로그에서 어느 호출부였는지
 * 구분하기 위한 것뿐이다. maxBytes는 테스트가 기가바이트 단위 버퍼 없이 상한 동작을 검증할 수
 * 있도록 둔 것으로, 실제 호출부는 지정하지 않고 기본값(1GB)을 그대로 쓴다. */
export function guardedUnzipSync(
  data: Uint8Array,
  contextLabel: string,
  opts?: UnzipOptions,
  maxBytes: number = DEFAULT_MAX_UNCOMPRESSED_BYTES,
): Unzipped {
  assertZipNotTooLarge(data, contextLabel, maxBytes);
  const entries = opts ? unzipSync(data, opts) : unzipSync(data);
  let total = 0;
  for (const bytes of Object.values(entries)) total += bytes.length;
  if (total > maxBytes) throw new ZipTooLargeError(contextLabel, total, maxBytes);
  return entries;
}

const EMPTY = new Uint8Array(0);

/**
 * XML·관계 파트만 실제로 풀고, 나머지(그림·동영상·내장 파일)는 이름만 남긴다(빈 바이트) — 슬라이드를 처음 열 때
 * 필요한 건 XML뿐인데 예전엔 매번 파일 전체(실측 21MB pptx의 그림까지)를 풀었다. 그림은 "있는지"만 확인하므로
 * 빈 자리표시로 충분하다. 바이트가 필요한 곳(그림 내려주기·저장)은 이 함수를 쓰지 않는다.
 */
export function unzipXmlParts(data: Uint8Array, contextLabel: string): Unzipped {
  const skipped: string[] = [];
  const entries = guardedUnzipSync(data, contextLabel, {
    filter: (file) => {
      const isXml = /\.(xml|rels)$/i.test(file.name);
      if (!isXml) skipped.push(file.name);
      return isXml;
    },
  });
  for (const name of skipped) entries[name] = EMPTY;
  return entries;
}

/**
 * zip의 중앙 디렉터리에 적힌 "풀린 크기"의 합 — 압축을 풀지 않고 읽는다(ZIP64 포함). zip이 아니거나 구조를 읽지
 * 못하면 null. fflate는 이 선언 크기만큼의 출력 버퍼에 풀기 때문에, 이 합이 상한 안이면 풀기에 드는 메모리도 그 안이다.
 */
export function declaredUnzippedSize(data: Uint8Array): number | null {
  const entries = declaredEntrySizes(data);
  if (!entries) return null;
  let total = 0;
  for (const size of entries.values()) total += size;
  return total;
}

/** 항목 이름 → 풀린 크기(중앙 디렉터리 선언값). zip이 아니거나 못 읽으면 null */
export function declaredEntrySizes(data: Uint8Array): Map<string, number> | null {
  if (data.length < 22 || data[0] !== 0x50 || data[1] !== 0x4b) return null;
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
  if (eocd < 0) return null;
  let count = u16(eocd + 10);
  let offset = u32(eocd + 16);
  if ((count === 0xffff || offset === 0xffffffff) && eocd >= 20 && u32(eocd - 20) === 0x07064b50) {
    const z64 = u64(eocd - 12);
    if (z64 + 56 > data.length || u32(z64) !== 0x06064b50) return null;
    count = u64(z64 + 32);
    offset = u64(z64 + 48);
  }
  const sizes = new Map<string, number>();
  const nameDecoder = new TextDecoder();
  let p = offset;
  for (let n = 0; n < count; n += 1) {
    if (p + 46 > data.length || u32(p) !== 0x02014b50) return null;
    let size = u32(p + 24);
    const nameLen = u16(p + 28);
    const extraLen = u16(p + 30);
    const commentLen = u16(p + 32);
    if (size === 0xffffffff) {
      // ZIP64 추가 필드(0x0001)의 첫 8바이트가 풀린 크기
      let e = p + 46 + nameLen;
      const end = e + extraLen;
      size = -1;
      while (e + 4 <= end && e + 4 <= data.length) {
        const id = u16(e);
        const len = u16(e + 2);
        if (id === 0x0001 && len >= 8 && e + 12 <= data.length) {
          size = u64(e + 4);
          break;
        }
        e += 4 + len;
      }
      if (size < 0) return null;
    }
    if (p + 46 + nameLen > data.length) return null;
    sizes.set(nameDecoder.decode(data.subarray(p + 46, p + 46 + nameLen)), size);
    p += 46 + nameLen + extraLen + commentLen;
  }
  return sizes;
}

/** 선언된 풀린 크기가 상한을 넘으면 오류 — 업로드·저장 때 미리, 그리고 풀기 직전에 부른다 */
export function assertZipNotTooLarge(data: Uint8Array, contextLabel: string, maxBytes: number = DEFAULT_MAX_UNCOMPRESSED_BYTES): void {
  const declared = declaredUnzippedSize(data);
  if (declared !== null && declared > maxBytes) throw new ZipTooLargeError(contextLabel, declared, maxBytes);
}
