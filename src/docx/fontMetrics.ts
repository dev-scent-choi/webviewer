import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import fontkit from '@pdf-lib/fontkit';

/**
 * Word의 "1줄"(줄 간격 1.0) 높이 = 1.3 × (OS/2 usWinAscent + usWinDescent) / unitsPerEm × 글자 크기.
 * 근거(Office 정답 PDF): 맑은 고딕 (2229+495)/2048 = 1.330 → 1.729(잰 값 1.73), 바탕 (879+145)/1024 = 1.000 → 1.30
 * (d01: 11pt·1.15줄·뒤 간격 10pt 문단의 줄 간격 26.5pt → (26.5−10)/1.15/11 = 1.30). 두 글꼴 모두 같은 식에 맞는다.
 * 1.3배는 동아시아 글꼴(OS/2 코드 페이지에 한국어·일본어·중국어)에서 확인한 값이라 그런 글꼴만 표에 넣는다 — 라틴 글꼴(Arial 등)은
 * 브라우저 줄 높이를 그대로 따른다(Word 출력과 맞춰 보지 못했다).
 */
export const WORD_LINE_FACTOR = 1.3;
/** OS/2 ulCodePageRange1 비트 17~21: 일본어(JIS)·중국어 간체·한국어(완성형)·중국어 번체·한국어(조합형) */
const CJK_CODE_PAGES = 0x3e0000;

/** 서버에 깔린 글꼴의 1줄 비율(글꼴 이름 — 한글 이름 포함 → 비율). 처음 쓸 때 한 번 훑는다(약 1초, 사이사이 다른 요청을 받는다) */
let ratios: Record<string, number> | null = null;
let scanning: Promise<void> | null = null;

function fontDirs(): string[] {
  const dirs = [path.join(process.env.WINDIR ?? 'C:\\Windows', 'Fonts'), '/usr/share/fonts', '/usr/local/share/fonts', path.join(os.homedir(), '.fonts')];
  if (process.env.LOCALAPPDATA) dirs.push(path.join(process.env.LOCALAPPDATA, 'Microsoft', 'Windows', 'Fonts'));
  return dirs;
}

/** 글꼴 파일 하나의 (이름 → 1줄 비율) */
export function lineRatiosOfFont(bytes: Uint8Array): Record<string, number> {
  const out: Record<string, number> = {};
  const created = fontkit.create(bytes as Buffer) as unknown as { fonts?: unknown[] };
  for (const font of (created.fonts ?? [created]) as Array<{ unitsPerEm: number; 'OS/2'?: { winAscent: number; winDescent: number; codePageRange?: number[] }; name?: { records?: { fontFamily?: Record<string, string> } } }>) {
    const os2 = font['OS/2'];
    if (!os2 || !(font.unitsPerEm > 0) || !(os2.winAscent + os2.winDescent > 0)) continue;
    if (!((os2.codePageRange?.[0] ?? 0) & CJK_CODE_PAGES)) continue;
    const ratio = Number(((WORD_LINE_FACTOR * (os2.winAscent + Math.abs(os2.winDescent))) / font.unitsPerEm).toFixed(3));
    for (const name of Object.values(font.name?.records?.fontFamily ?? {})) if (typeof name === 'string' && name.length <= 64) out[name] = ratio;
  }
  return out;
}

async function scan(): Promise<void> {
  const found: Record<string, number> = {};
  for (const dir of fontDirs()) {
    const files = await fs.readdir(dir, { recursive: true }).catch(() => [] as string[]);
    for (const file of files) {
      if (!/\.(ttf|ttc|otf)$/i.test(file)) continue;
      try {
        Object.assign(found, lineRatiosOfFont(await fs.readFile(path.join(dir, file))));
      } catch {
        // 읽을 수 없는 글꼴은 건너뛴다
      }
      await new Promise((resolve) => setImmediate(resolve));
    }
  }
  ratios = found;
}

/** 지금 아는 글꼴별 1줄 비율 — 아직 훑는 중이면 빈 표(편집기는 schema.ts의 기본 표를 쓴다) */
export function installedLineRatios(): Record<string, number> {
  scanning ??= scan().catch(() => {
    ratios = {};
  });
  return ratios ?? {};
}
