/**
 * PDF 쪽 내용 스트림 해석기 — 바이트를 명령(연산자 + 피연산자)으로 나누고 각 명령의 원래 바이트 위치를 기억한다.
 * 고치지 않는 명령은 원래 바이트를 그대로 옮겨 쓰고(textRewrite.ts), 고친 명령만 새로 쓴다 — 해석기가 모르는 문법이 있어도
 * 내용이 바뀌지 않게. 인라인 그림(BI … ID … EI)은 통째로 한 명령으로 둔다.
 */

export type Operand =
  | number
  | boolean
  | null
  | { readonly name: string }
  | { readonly bytes: Uint8Array; readonly hex: boolean }
  | { readonly array: Operand[] }
  | { readonly dict: Map<string, Operand> };

export interface ContentOp {
  readonly op: string;
  readonly operands: Operand[];
  /** 원래 바이트 범위 [start, end) — 피연산자 포함 */
  readonly start: number;
  readonly end: number;
}

const WS = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20]);
const DELIM = new Set([0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25]);
const isRegular = (c: number) => !WS.has(c) && !DELIM.has(c);

class Lexer {
  pos = 0;
  constructor(readonly b: Uint8Array) {}

  skipWs(): void {
    const b = this.b;
    while (this.pos < b.length) {
      const c = b[this.pos];
      if (WS.has(c)) this.pos += 1;
      else if (c === 0x25) {
        // 주석 — 줄 끝까지
        while (this.pos < b.length && b[this.pos] !== 0x0a && b[this.pos] !== 0x0d) this.pos += 1;
      } else break;
    }
  }

  /** 다음 토큰 — 연산자는 { kw }, 끝이면 undefined(리터럴 null과 구분) */
  next(): Operand | { kw: string } | { close: string } | undefined {
    this.skipWs();
    const b = this.b;
    if (this.pos >= b.length) return undefined;
    const c = b[this.pos];
    if (c === 0x2f) return { name: this.readName() };
    if (c === 0x28) return { bytes: this.readLiteral(), hex: false };
    if (c === 0x3c) {
      if (b[this.pos + 1] === 0x3c) {
        this.pos += 2;
        return { dict: this.readDict() };
      }
      return { bytes: this.readHex(), hex: true };
    }
    if (c === 0x5b) {
      this.pos += 1;
      const array: Operand[] = [];
      for (;;) {
        const t = this.next();
        if (t === undefined) break;
        if (t !== null && typeof t === 'object' && 'close' in t) {
          if (t.close === ']') break;
          continue;
        }
        // 배열 안의 키워드(잘못된 내용)는 버린다
        if (t !== null && typeof t === 'object' && 'kw' in t) continue;
        array.push(t as Operand);
      }
      return { array };
    }
    if (c === 0x5d) {
      this.pos += 1;
      return { close: ']' };
    }
    if (c === 0x3e && b[this.pos + 1] === 0x3e) {
      this.pos += 2;
      return { close: '>>' };
    }
    if (c === 0x7b || c === 0x7d || c === 0x29 || c === 0x3e) {
      this.pos += 1;
      return { kw: String.fromCharCode(c) };
    }
    const start = this.pos;
    while (this.pos < b.length && isRegular(b[this.pos])) this.pos += 1;
    if (this.pos === start) {
      this.pos += 1;
      return { kw: String.fromCharCode(c) };
    }
    const word = latin1(b.subarray(start, this.pos));
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(word)) return Number(word);
    if (word === 'true') return true;
    if (word === 'false') return false;
    if (word === 'null') return null;
    return { kw: word };
  }

  private readName(): string {
    const b = this.b;
    this.pos += 1;
    let out = '';
    while (this.pos < b.length && isRegular(b[this.pos])) {
      const c = b[this.pos];
      if (c === 0x23 && this.pos + 2 < b.length) {
        const hex = latin1(b.subarray(this.pos + 1, this.pos + 3));
        if (/^[0-9a-fA-F]{2}$/.test(hex)) {
          out += String.fromCharCode(parseInt(hex, 16));
          this.pos += 3;
          continue;
        }
      }
      out += String.fromCharCode(c);
      this.pos += 1;
    }
    return out;
  }

  private readLiteral(): Uint8Array {
    const b = this.b;
    this.pos += 1;
    const out: number[] = [];
    let depth = 1;
    while (this.pos < b.length) {
      const c = b[this.pos++];
      if (c === 0x5c) {
        const e = b[this.pos++];
        if (e === undefined) break;
        const map: Record<number, number> = { 0x6e: 0x0a, 0x72: 0x0d, 0x74: 0x09, 0x62: 0x08, 0x66: 0x0c, 0x28: 0x28, 0x29: 0x29, 0x5c: 0x5c };
        if (e in map) out.push(map[e]);
        else if (e >= 0x30 && e <= 0x37) {
          let v = e - 0x30;
          for (let k = 0; k < 2 && b[this.pos] >= 0x30 && b[this.pos] <= 0x37; k += 1) v = v * 8 + (b[this.pos++] - 0x30);
          out.push(v & 0xff);
        } else if (e === 0x0d) {
          if (b[this.pos] === 0x0a) this.pos += 1;
        } else if (e !== 0x0a) out.push(e);
        continue;
      }
      if (c === 0x28) depth += 1;
      else if (c === 0x29) {
        depth -= 1;
        if (depth === 0) break;
      }
      out.push(c);
    }
    return Uint8Array.from(out);
  }

  private readHex(): Uint8Array {
    const b = this.b;
    this.pos += 1;
    let digits = '';
    while (this.pos < b.length && b[this.pos] !== 0x3e) {
      const ch = String.fromCharCode(b[this.pos++]);
      if (/[0-9a-fA-F]/.test(ch)) digits += ch;
    }
    this.pos += 1;
    if (digits.length % 2) digits += '0';
    const out = new Uint8Array(digits.length / 2);
    for (let i = 0; i < out.length; i += 1) out[i] = parseInt(digits.slice(i * 2, i * 2 + 2), 16);
    return out;
  }

  private readDict(): Map<string, Operand> {
    const dict = new Map<string, Operand>();
    for (;;) {
      const k = this.next();
      if (k === undefined || (typeof k === 'object' && k !== null && 'close' in k)) break;
      if (typeof k !== 'object' || k === null || !('name' in k)) continue;
      const v = this.next();
      if (v === undefined || (typeof v === 'object' && v !== null && 'close' in v)) break;
      if (typeof v === 'object' && v !== null && 'kw' in v) continue;
      dict.set(k.name, v as Operand);
    }
    return dict;
  }
}

function latin1(b: Uint8Array): string {
  let s = '';
  for (const c of b) s += String.fromCharCode(c);
  return s;
}

/** 내용 스트림 → 명령 목록 */
export function parseContent(bytes: Uint8Array): ContentOp[] {
  const lx = new Lexer(bytes);
  const ops: ContentOp[] = [];
  let operands: Operand[] = [];
  let start = -1;
  for (;;) {
    lx.skipWs();
    const tokenStart = lx.pos;
    const t = lx.next();
    if (t === undefined) break;
    if (start < 0) start = tokenStart;
    if (typeof t === 'object' && t !== null && 'close' in t) continue;
    if (typeof t === 'object' && t !== null && 'kw' in t) {
      if (t.kw === 'BI') {
        // 인라인 그림: 사전(키 값 …) ID 한 칸 공백 뒤 자료 … EI
        const dict = new Map<string, Operand>();
        for (;;) {
          const k = lx.next();
          if (k === undefined) break;
          if (typeof k === 'object' && k !== null && 'kw' in k && k.kw === 'ID') break;
          const v = lx.next();
          if (typeof k === 'object' && k !== null && 'name' in k && v !== undefined && !(typeof v === 'object' && v !== null && ('kw' in v || 'close' in v))) dict.set(k.name, v as Operand);
        }
        lx.pos += 1;
        const b = bytes;
        let i = lx.pos;
        while (i < b.length) {
          if (b[i] === 0x45 && b[i + 1] === 0x49 && (i + 2 >= b.length || WS.has(b[i + 2])) && (i === 0 || WS.has(b[i - 1]))) break;
          i += 1;
        }
        lx.pos = Math.min(b.length, i + 2);
        ops.push({ op: 'BI', operands: [{ dict }], start, end: lx.pos });
      } else {
        ops.push({ op: t.kw, operands, start, end: lx.pos });
      }
      operands = [];
      start = -1;
      continue;
    }
    operands.push(t as Operand);
  }
  return ops;
}

// ---------- 새 명령 쓰기 ----------
export function num(n: number): string {
  if (!Number.isFinite(n)) return '0';
  const r = Math.round(n * 10000) / 10000;
  return Object.is(r, -0) ? '0' : String(r);
}

export function hexOf(bytes: Uint8Array): string {
  let s = '<';
  for (const c of bytes) s += c.toString(16).padStart(2, '0');
  return `${s}>`;
}

export function nameOf(name: string): string {
  let s = '/';
  for (const ch of name) {
    const c = ch.charCodeAt(0);
    s += c < 0x21 || c > 0x7e || DELIM.has(c) || c === 0x23 ? `#${c.toString(16).padStart(2, '0')}` : ch;
  }
  return s;
}

export function operandToString(o: Operand): string {
  if (typeof o === 'number') return num(o);
  if (typeof o === 'boolean') return o ? 'true' : 'false';
  if (o === null) return 'null';
  if ('name' in o) return nameOf(o.name);
  if ('bytes' in o) return hexOf(o.bytes);
  if ('array' in o) return `[${o.array.map(operandToString).join(' ')}]`;
  return `<<${[...o.dict].map(([k, v]) => `${nameOf(k)} ${operandToString(v)}`).join(' ')}>>`;
}

export function opToString(op: string, operands: readonly Operand[]): string {
  return `${operands.map(operandToString).join(' ')}${operands.length ? ' ' : ''}${op}`;
}

// ---------- 행렬 ----------
export type Matrix = [number, number, number, number, number, number];
export const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];
/** m1 × m2 (PDF 순서: 먼저 m1, 그다음 m2) */
export function mul(m1: readonly number[], m2: readonly number[]): Matrix {
  return [
    m1[0] * m2[0] + m1[1] * m2[2],
    m1[0] * m2[1] + m1[1] * m2[3],
    m1[2] * m2[0] + m1[3] * m2[2],
    m1[2] * m2[1] + m1[3] * m2[3],
    m1[4] * m2[0] + m1[5] * m2[2] + m2[4],
    m1[4] * m2[1] + m1[5] * m2[3] + m2[5],
  ];
}
export function apply(m: readonly number[], x: number, y: number): [number, number] {
  return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
}
export function invert(m: readonly number[]): Matrix | null {
  const det = m[0] * m[3] - m[1] * m[2];
  if (Math.abs(det) < 1e-12) return null;
  return [m[3] / det, -m[1] / det, -m[2] / det, m[0] / det, (m[2] * m[5] - m[3] * m[4]) / det, (m[1] * m[4] - m[0] * m[5]) / det];
}
