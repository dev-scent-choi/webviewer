/**
 * 변환용 임시 폴더 청소 — 변환은 OS 임시 폴더에 올린 문서의 사본을 두고 끝나면 finally에서 지운다.
 * 그런데 변환 도중 서버가 재시작·강제 종료되면(tsx watch 재시작, 작업 종료 등) finally가 돌지 않아
 * 사본이 그대로 남았다(실측: 2026-10-19 기준 70개, 21MB짜리 pptx 사본 포함). 향기랩은 "올린 문서를
 * 저장하지 않는다"고 약속하므로, 서버를 켤 때와 그 뒤 주기적으로 오래된 것을 지운다.
 *
 * 같은 PC에서 두 인스턴스(8000·8010)가 같은 임시 폴더를 쓰므로, 다른 인스턴스가 지금 변환 중인 폴더를
 * 지우지 않도록 나이(마지막 수정 후 STALE_AFTER_MS)로만 판단한다 — 변환 제한 시간(3분)보다 넉넉히 길다.
 * LibreOffice 프로필(docview-lo-profiles)은 재사용하는 것이라 건드리지 않는다.
 */
import { promises as fs } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { sweepStaleChunks } from './chunkedUpload.js';

export const STALE_AFTER_MS = 30 * 60_000;
export const SWEEP_INTERVAL_MS = 30 * 60_000;

/** 이 앱이 만드는 임시 폴더·파일 이름 — docview-lo-profiles는 일부러 빠져 있다(`docview-lo-`와 달리 뒤가 `profiles`) */
const TEMP_NAME = /^docview-(?:convert|svg|dwg|legacyxlsx|lo|print|dxfsvg|warmup)-[A-Za-z0-9]{6}$|^docview-xlsx-stream-[0-9a-f-]{36}\.xlsx$/;

export function isSweepableTempName(name: string): boolean {
  return TEMP_NAME.test(name);
}

/** 오래된 변환 임시 폴더·파일을 지우고 지운 개수를 돌려준다. 실패한 항목은 다음 청소 때 다시 시도한다. */
export async function sweepStaleTempDirs(dir: string = os.tmpdir(), now: number = Date.now()): Promise<number> {
  let names: string[];
  try {
    names = await fs.readdir(dir);
  } catch {
    return 0;
  }
  let removed = 0;
  await Promise.all(
    names.filter(isSweepableTempName).map(async (name) => {
      const target = path.join(dir, name);
      try {
        const stat = await fs.stat(target);
        if (now - stat.mtimeMs < STALE_AFTER_MS) return; // 지금 변환 중일 수 있다
        await fs.rm(target, { recursive: true, force: true, maxRetries: 3, retryDelay: 200 });
        removed += 1;
      } catch {
        /* 잠겨 있거나 이미 없어졌다 — 다음 청소 때 다시 */
      }
    }),
  );
  return removed;
}

/** 서버 시작 때 한 번, 그 뒤 주기적으로 청소한다. 타이머는 프로세스 종료를 막지 않는다. */
export function startTempSweeper(log: (removed: number) => void): void {
  const run = () => {
    void sweepStaleTempDirs().then((removed) => removed > 0 && log(removed));
    // 올리다 만 조각 업로드(docview-chunks)도 — 새 조각이 안 오면 지워지지 않고 남았다(향기랩은 '저장하지 않음' 약속)
    void sweepStaleChunks(true);
  };
  run();
  setInterval(run, SWEEP_INTERVAL_MS).unref();
}
