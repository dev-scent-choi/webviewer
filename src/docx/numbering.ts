/**
 * word/numbering.xml의 번호 매기기 정의와 번호 글자 계산 — 서버(가져오기)가 정의를 읽어 편집기에 내려주고,
 * 편집기(docxLayout.ts)가 문서 순서대로 번호(1. / 1.1.1. / 가) / (ㄱ) / ●)를 매겨 문단 앞에 그린다.
 * 번호는 문서에 글자로 넣지 않는다 — 원본의 w:numPr가 그대로 저장되므로 Word에서 다시 열어도 진짜 목록이다.
 */

export interface NumberingLevel {
  readonly start: number;
  readonly numFmt: string;
  readonly lvlText: string;
  /** 이 수준 번호를 다시 1부터 세게 만드는 상위 수준(w:lvlRestart, 0 = 다시 안 셈). 없으면 바로 위 수준 아무거나. */
  readonly restart: number | null;
  readonly isLgl: boolean;
  /** 번호 뒤 간격 — tab(기본)·space·nothing */
  readonly suff: 'tab' | 'space' | 'nothing';
  /** 수준의 들여쓰기(dxa) — 문단에 직접 들여쓰기가 없을 때 쓴다 */
  readonly left: number;
  readonly hanging: number;
  readonly firstLine: number;
  readonly bold: boolean;
  readonly color: string | null;
}

export interface NumberingDefs {
  /** numId → 수준 9개(빈 자리는 null) */
  readonly nums: Readonly<Record<string, ReadonlyArray<NumberingLevel | null>>>;
}

/** 기호 글꼴(Symbol·Wingdings)의 사설 영역 글자 → 화면 글꼴로 그릴 수 있는 같은 모양 */
const SYMBOL_BULLETS: Readonly<Record<string, string>> = {
  '': '●', '': '■', '': '➢', '': '✓', '': '❖', '': '■', '': '□', '': '➔',
  '': '◆', '': '○', '': '•', '': '◇', '': '☑', '': '●', '': '❑',
};

const GANADA = '가나다라마바사아자차카타파하';
const CHOSUNG = 'ㄱㄴㄷㄹㅁㅂㅅㅇㅈㅊㅋㅌㅍㅎ';
const KOREAN_DIGITS = ['영', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
const KOREAN_COUNTING = ['', '하나', '둘', '셋', '넷', '다섯', '여섯', '일곱', '여덟', '아홉', '열'];
const CJK_DIGITS = ['〇', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];

function roman(n: number): string {
  const table: ReadonlyArray<[number, string]> = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
  let out = '';
  let rest = n;
  for (const [value, letters] of table) while (rest >= value) { out += letters; rest -= value; }
  return out;
}

/** Word의 letter 형식: 1→A, 26→Z, 27→AA(같은 글자를 겹침) */
function letters(n: number): string {
  const ch = String.fromCharCode(65 + ((n - 1) % 26));
  return ch.repeat(Math.floor((n - 1) / 26) + 1);
}

const cycle = (set: string, n: number) => set[(n - 1) % set.length];

/** 숫자 하나를 Word 번호 형식(w:numFmt)으로 */
export function formatNumber(n: number, numFmt: string): string {
  switch (numFmt) {
    case 'upperRoman': return roman(n);
    case 'lowerRoman': return roman(n).toLowerCase();
    case 'upperLetter': return letters(n);
    case 'lowerLetter': return letters(n).toLowerCase();
    case 'decimalZero': return n < 10 ? `0${n}` : String(n);
    case 'ganada': return cycle(GANADA, n);
    case 'chosung': return cycle(CHOSUNG, n);
    case 'decimalEnclosedCircle':
    case 'decimalEnclosedCircleChinese': return n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n);
    case 'decimalEnclosedParen': return n >= 1 && n <= 20 ? String.fromCharCode(0x2474 + n - 1) : `(${n})`;
    case 'decimalEnclosedFullstop': return n >= 1 && n <= 20 ? String.fromCharCode(0x2488 + n - 1) : `${n}.`;
    case 'decimalFullWidth':
    case 'decimalFullWidth2': return String(n).replace(/\d/g, (d) => String.fromCharCode(0xff10 + Number(d)));
    case 'koreanDigital':
    case 'koreanDigital2': return String(n).split('').map((d) => KOREAN_DIGITS[Number(d)]).join('');
    case 'koreanCounting': return n <= 10 ? KOREAN_COUNTING[n] : String(n);
    case 'ideographDigital':
    case 'japaneseCounting':
    case 'chineseCounting':
    case 'chineseCountingThousand':
    case 'koreanLegal': return n <= 10 ? CJK_DIGITS[n] : String(n);
    case 'none': return '';
    default: return String(n); // decimal·ordinal 등
  }
}

/** 문서 순서대로 번호를 매기는 셈 상태 — numId마다 수준별 현재 값을 따로 센다(Word와 같다). */
export class NumberingCounter {
  private readonly counts = new Map<string, number[]>();

  constructor(private readonly defs: NumberingDefs) {}

  /** 이 문단의 번호 글자와 수준 정의. 정의가 없으면 null. */
  next(numId: string, ilvl: number): { label: string; level: NumberingLevel } | null {
    const levels = this.defs.nums[numId];
    const level = levels?.[ilvl];
    if (!levels || !level) return null;
    const counts = this.counts.get(numId) ?? [];
    this.counts.set(numId, counts);
    counts[ilvl] = counts[ilvl] === undefined ? level.start : counts[ilvl] + 1;
    // 이 수준보다 깊은 수준은 다시 처음부터 센다(w:lvlRestart로 막은 수준은 그대로)
    for (let deeper = ilvl + 1; deeper < 9; deeper += 1) {
      const restart = levels[deeper]?.restart;
      if (restart === 0) continue;
      if (restart === null || restart === undefined || ilvl < restart) counts[deeper] = undefined as unknown as number;
    }
    if (level.numFmt === 'bullet') {
      const ch = level.lvlText;
      return { label: [...ch].map((c) => SYMBOL_BULLETS[c] ?? c).join(''), level };
    }
    const label = level.lvlText.replace(/%([1-9])/g, (_, d: string) => {
      const i = Number(d) - 1;
      const lv = levels[i];
      if (!lv) return '';
      const value = counts[i] ?? lv.start;
      return formatNumber(value, level.isLgl && i < ilvl ? 'decimal' : lv.numFmt);
    });
    return { label, level };
  }
}

/** numbering.xml(파싱된 자식 목록을 다루기 싫어 정규식으로 충분한 구조만 읽는다 — 요소가 중첩되지 않는 평평한 형식이다) */
export function parseNumberingXml(xml: string): NumberingDefs {
  const attr = (s: string, tag: string, name = 'w:val'): string | undefined =>
    new RegExp(`<${tag}\\b[^>]*\\b${name}="([^"]*)"`).exec(s)?.[1];
  const parseLevels = (body: string): Array<NumberingLevel | null> => {
    const levels: Array<NumberingLevel | null> = Array.from({ length: 9 }, () => null);
    for (const m of body.matchAll(/<w:lvl\b[^>]*w:ilvl="(\d)"[^>]*>([\s\S]*?)<\/w:lvl>/g)) {
      const s = m[2];
      const restartRaw = attr(s, 'w:lvlRestart');
      const ind = /<w:ind\b[^>]*\/>/.exec(s)?.[0] ?? '';
      const num = (name: string) => Number(new RegExp(`\\bw:${name}="(-?\\d+)"`).exec(ind)?.[1] ?? 0);
      const rPr = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(s)?.[1] ?? '';
      const color = attr(rPr, 'w:color');
      levels[Number(m[1])] = {
        start: Number(attr(s, 'w:start') ?? 1),
        numFmt: attr(s, 'w:numFmt') ?? 'decimal',
        lvlText: decodeXmlText(attr(s, 'w:lvlText') ?? ''),
        restart: restartRaw === undefined ? null : Number(restartRaw),
        isLgl: /<w:isLgl\b(?![^>]*w:val="(0|false)")/.test(s),
        suff: (attr(s, 'w:suff') as NumberingLevel['suff'] | undefined) ?? 'tab',
        left: num('left') || num('start'),
        hanging: num('hanging'),
        firstLine: num('firstLine'),
        bold: /<w:b\/>|<w:b w:val="(1|true|on)"/.test(rPr),
        color: color && /^[0-9A-Fa-f]{6}$/.test(color) ? `#${color.toLowerCase()}` : null,
      };
    }
    return levels;
  };
  const abstract = new Map<string, { levels: Array<NumberingLevel | null>; styleLink?: string }>();
  for (const m of xml.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"[^>]*>([\s\S]*?)<\/w:abstractNum>/g)) {
    abstract.set(m[1], { levels: parseLevels(m[2]), styleLink: attr(m[2], 'w:numStyleLink') });
  }
  const nums: Record<string, Array<NumberingLevel | null>> = {};
  const numByStyleLink = new Map<string, string>();
  for (const m of xml.matchAll(/<w:abstractNum\b[^>]*w:abstractNumId="(\d+)"[^>]*>([\s\S]*?)<\/w:abstractNum>/g)) {
    const link = attr(m[2], 'w:styleLink');
    if (link) numByStyleLink.set(link, m[1]);
  }
  for (const m of xml.matchAll(/<w:num\b[^>]*w:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g)) {
    const absId = attr(m[2], 'w:abstractNumId');
    let base = absId ? abstract.get(absId) : undefined;
    // numStyleLink: 정의가 다른 abstractNum(styleLink)에 있다
    if (base?.styleLink && numByStyleLink.has(base.styleLink)) base = abstract.get(numByStyleLink.get(base.styleLink)!);
    if (!base) continue;
    const levels = [...base.levels];
    for (const o of m[2].matchAll(/<w:lvlOverride\b[^>]*w:ilvl="(\d)"[^>]*>([\s\S]*?)<\/w:lvlOverride>/g)) {
      const i = Number(o[1]);
      const overridden = parseLevels(o[2])[i];
      const startOverride = attr(o[2], 'w:startOverride');
      if (overridden) levels[i] = overridden;
      else if (startOverride !== undefined && levels[i]) levels[i] = { ...levels[i]!, start: Number(startOverride) };
    }
    nums[m[1]] = levels;
  }
  return { nums };
}

function decodeXmlText(s: string): string {
  return s
    .replace(/&#x([0-9a-f]+);/gi, (_, h: string) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d: string) => String.fromCodePoint(Number(d)))
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}
