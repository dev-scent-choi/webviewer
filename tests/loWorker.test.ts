import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { afterAll, describe, expect, test } from 'vitest';
import { PDFDocument } from 'pdf-lib';
import { LibreOfficeWorker } from '../src/converters/loWorker.js';
import { sofficePath } from '../src/converters/libreOffice.js';

const programDir = sofficePath ? path.dirname(sofficePath) : '';
const python = [path.join(programDir, 'python.exe'), path.join(programDir, 'python')].find((p) => programDir && existsSync(p));

/** 앞 변환의 뒷정리("한가함")가 끝날 때까지는 일부러 맡지 않으므로, 맡을 때까지 잠깐씩 다시 시도한다 */
async function convertWhenIdle(worker: LibreOfficeWorker, src: string, out: string, filterData: Record<string, boolean> | null): Promise<boolean> {
  for (let i = 0; i < 80; i += 1) {
    if (await worker.tryConvert(src, out, filterData)) return true;
    await new Promise((r) => setTimeout(r, 250));
  }
  return false;
}

// LibreOffice가 깔린 PC에서만(CI 등에는 없다)
describe.skipIf(!sofficePath || !python)('LibreOffice 상주 변환기', () => {
  const dir = mkdtempSync(path.join(os.tmpdir(), 'loworker-test-'));
  const tag = `test${process.pid}`;
  const worker = new LibreOfficeWorker(python!, path.resolve('src/converters/loWorker.py'), sofficePath!, tag);
  afterAll(async () => {
    worker.kill();
    // soffice가 꺼지며 파일을 놓을 때까지 잠깐 기다린 뒤 시험 폴더·시험 프로필을 지운다
    await new Promise((r) => setTimeout(r, 6000));
    for (const target of [dir, path.join(os.tmpdir(), 'docview-lo-profiles', `worker-${tag}`)]) {
      try {
        rmSync(target, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 });
      } catch {
        /* 남아도 시험 결과와는 무관 — 다음 임시 폴더 청소 때 정리 */
      }
    }
  }, 30_000);

  test('csv를 PDF로 바꾸고, 두 번째부터는 빠르다', { timeout: 120_000 }, async () => {
    const src = path.join(dir, 'a.csv');
    writeFileSync(src, '이름,값\n가,1\n나,2\n');
    expect(await worker.warmUp()).toBe(true);
    const out1 = path.join(dir, 'a1.pdf');
    expect(await worker.tryConvert(src, out1, { SinglePageSheets: true })).toBe(true);
    expect((await PDFDocument.load(readFileSync(out1))).getPageCount()).toBe(1);
    await new Promise((r) => setTimeout(r, 500)); // 앞 변환의 뒷정리
    const started = Date.now();
    expect(await worker.tryConvert(src, path.join(dir, 'a2.pdf'), null)).toBe(true);
    // 새 프로필의 처음 몇 번은 느리다(실측 2.9초) — 여기선 "멈추지 않고 바뀐다"만 본다. 속도는 운영 측정으로 확인(0.9초)
    expect(Date.now() - started).toBeLessThan(5000);
  });

  test('한글 파일 이름도 바꾼다(표준 입출력 UTF-8 — cp949로 읽어 "type detection failed"가 났었다)', { timeout: 60_000 }, async () => {
    const src = path.join(dir, '한글 보고서(최종).csv');
    writeFileSync(src, '이름,값\n가,1\n');
    const out = path.join(dir, '한글 보고서(최종).pdf');
    expect(await convertWhenIdle(worker, src, out, null)).toBe(true);
    expect(existsSync(out)).toBe(true);
  });

  test('없는 파일·큰 파일·한글(hwp)은 맡지 않는다(false — 호출부가 예전 방식으로)', { timeout: 60_000 }, async () => {
    expect(await worker.tryConvert(path.join(dir, '없음.docx'), path.join(dir, 'x.pdf'), null)).toBe(false);
    const big = path.join(dir, 'big.csv');
    writeFileSync(big, Buffer.alloc(6 * 1024 * 1024, 65));
    expect(await worker.tryConvert(big, path.join(dir, 'big.pdf'), null)).toBe(false);
    const hwp = path.join(dir, 'a.hwp');
    writeFileSync(hwp, 'x');
    expect(await worker.tryConvert(hwp, path.join(dir, 'hwp.pdf'), null)).toBe(false);
  });

  test('열 수 없는 문서 뒤에도 계속 쓸 수 있다', { timeout: 120_000 }, async () => {
    const broken = path.join(dir, 'broken.docx');
    writeFileSync(broken, 'not a zip');
    // LibreOffice는 망가진 docx를 글자 문서로 열기도 한다 — 성공이든 실패든 변환기는 살아 있어야 한다
    expect(typeof (await worker.tryConvert(broken, path.join(dir, 'broken.pdf'), null))).toBe('boolean');
    const src = path.join(dir, 'after.csv');
    writeFileSync(src, 'a\n1\n');
    expect(await convertWhenIdle(worker, src, path.join(dir, 'after.pdf'), null)).toBe(true);
  });
});
