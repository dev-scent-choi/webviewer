/**
 * LibreOffice 상주 변환기(loWorker.py)를 띄우고 변환을 맡긴다.
 *
 * 매 변환마다 soffice를 새로 띄우면 시작에만 1.5초 남짓이 든다(실측: 작은 docx·pptx·xlsx가 1.2~1.4초).
 * 상주 변환기는 같은 파일을 0.2~0.4초에 바꾼다(출력 PDF는 같다 — 쪽 수·쪽 크기·파일 크기 비교).
 *
 * 안전장치 — 이건 빠르게 하는 길일 뿐, 실패하면 호출부가 예전 방식(soffice 명령)으로 바로 다시 한다:
 * - 큰 파일(WORKER_MAX_BYTES 이상)은 맡기지 않는다: 빨라지지도 않고(그림 많은 21MB pptx는 같은 8.6초),
 *   같은 큰 파일을 연달아 바꾸면 soffice가 죽었다.
 * - 이미 다른 변환 중이면 기다리지 않고 바로 거절한다(동시 업로드는 예전 방식의 프로필 슬롯이 나눠 맡는다).
 * - 제한 시간을 넘기면 변환기를 통째로 끄고 다음 요청 때 다시 띄운다. WORKER_JOBS_BEFORE_RESTART번마다 새로 띄워
 *   오래 켜 둔 soffice의 메모리가 쌓이지 않게 한다.
 * - 인스턴스(포트)마다 프로필·파이프 이름이 따로라 8000·8010이 서로 간섭하지 않는다.
 */
import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process';
import { existsSync, statSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { ensureFontSubstitutions } from './loFontSubstitution.js';

const WORKER_MAX_BYTES = 5 * 1024 * 1024;
/** 맡는 형식 — 오피스 문서만. 한글(hwp)은 변환이 느려(56쪽 18초) 이득이 적은데 상주 soffice 메모리를 1GB로 키워
 * 그대로 붙잡고 있었다(실측) — 예전 방식(변환이 끝나면 프로세스가 끝나 메모리를 돌려준다)으로 둔다. */
const WORKER_EXTS = new Set(['docx', 'doc', 'odt', 'rtf', 'txt', 'pptx', 'ppt', 'odp', 'xlsx', 'xls', 'ods', 'csv']);
const WORKER_JOBS_BEFORE_RESTART = 50;
const JOB_TIMEOUT_MS = 120_000;
const START_TIMEOUT_MS = 90_000;
/** 띄우기에 실패하면 이만큼 지난 뒤에야 다시 시도한다(그동안은 예전 방식) */
const RETRY_AFTER_FAILURE_MS = 10 * 60_000;
/** DOCVIEW_LO_WORKER_DEBUG=1이면 변환마다 어느 길로 갔는지 남긴다(운영 기본은 조용히) */
const DEBUG = process.env.DOCVIEW_LO_WORKER_DEBUG === '1';

function debugSkip(src: string, reason: string): false {
  if (DEBUG) console.info(`[상주 변환기] ${path.basename(src)} → 예전 방식(${reason})`);
  return false;
}

type FilterData = Readonly<Record<string, boolean | string | number>>;

interface Job {
  readonly id: number;
  readonly file: string;
  resolve(ok: boolean): void;
  timer: NodeJS.Timeout;
}

export class LibreOfficeWorker {
  private child: ChildProcessWithoutNullStreams | null = null;
  private ready: Promise<boolean> | null = null;
  private job: Job | null = null;
  /** 결과를 돌려준 뒤 문서를 닫고 soffice가 멀쩡한지 확인하는 중 — 그동안은 새 변환을 맡지 않는다(예전 방식으로) */
  private settling = false;
  /** 띄우기를 마치고 연결된 상태 — 아직이면 변환을 기다리게 하지 않고 예전 방식으로 보낸다 */
  private isReady = false;
  private nextId = 1;
  private jobsDone = 0;
  private failedAt = 0;
  private buffer = '';
  private onReady: ((ok: boolean) => void) | null = null;

  constructor(
    private readonly pythonPath: string,
    private readonly scriptPath: string,
    private readonly sofficePath: string,
    private readonly tag: string,
  ) {}

  /** 서버 시작 때 미리 띄워 둔다(첫 변환이 띄우는 시간을 기다리지 않게) */
  warmUp(): Promise<boolean> {
    return this.ensureStarted();
  }

  /** 맡을 수 있으면 변환하고 true, 맡을 수 없거나 실패하면 false(호출부가 예전 방식으로) */
  async tryConvert(src: string, out: string, filterData: FilterData | null): Promise<boolean> {
    if (!WORKER_EXTS.has(path.extname(src).slice(1).toLowerCase())) return false;
    try {
      if (statSync(src).size >= WORKER_MAX_BYTES) return false;
    } catch {
      return false;
    }
    if (this.job || this.settling) return debugSkip(src, this.job ? '다른 변환 중' : '뒷정리 중');
    // 아직 띄우는 중이면(첫 프로필 생성·스스로 다시 켜짐은 10초 넘게 걸린다 — 실측: 그동안 첫 변환이 12초를 기다렸다)
    // 기다리지 않고 예전 방식으로 보내고, 띄우기는 뒤에서 계속한다
    if (!this.isReady || !this.child) {
      void this.ensureStarted();
      return debugSkip(src, '아직 준비 안 됨');
    }
    const began = Date.now();
    const id = this.nextId++;
    const ok = await new Promise<boolean>((resolve) => {
      const timer = setTimeout(() => {
        this.job = null;
        resolve(false);
        this.kill(); // 멈춘 soffice — 다음 요청 때 새로 띄운다
      }, JOB_TIMEOUT_MS);
      this.job = { id, file: path.basename(src), resolve, timer };
      this.settling = true; // "idle" 응답이 올 때까지
      this.child!.stdin.write(`${JSON.stringify({ id, src, out, filterData })}\n`);
    });
    this.jobsDone += 1;
    if (DEBUG) console.info(`[상주 변환기] ${path.basename(src)} ${ok ? '성공' : '실패'} ${Date.now() - began}ms`);
    if (this.jobsDone >= WORKER_JOBS_BEFORE_RESTART) this.kill();
    return ok;
  }

  private ensureStarted(): Promise<boolean> {
    if (this.ready) return this.ready;
    if (Date.now() - this.failedAt < RETRY_AFTER_FAILURE_MS) return Promise.resolve(false);
    this.ready = new Promise<boolean>((resolve) => {
      const profile = path.join(os.tmpdir(), 'docview-lo-profiles', `worker-${this.tag}`);
      ensureFontSubstitutions(profile);
      const child = spawn(this.pythonPath, [this.scriptPath, this.sofficePath, `docview_lo_${this.tag}_${process.pid}`, pathToFileURL(profile).href], {
        stdio: ['pipe', 'pipe', 'pipe'],
        windowsHide: true,
      });
      this.child = child;
      this.jobsDone = 0;
      const startTimer = setTimeout(() => finishStart(false), START_TIMEOUT_MS);
      const finishStart = (ok: boolean) => {
        clearTimeout(startTimer);
        this.onReady = null;
        if (!ok) {
          this.failedAt = Date.now();
          this.kill();
        }
        this.isReady = ok;
        if (DEBUG) console.info(`[상주 변환기] 준비 ${ok ? '됨' : '실패'}`);
        resolve(ok);
      };
      this.onReady = finishStart;
      child.stdout.setEncoding('utf8');
      child.stdout.on('data', (chunk: string) => this.onData(chunk));
      // 변환기의 진단 출력(표준 오류) — 평소엔 버리고, 디버그 때만 보인다(stdout만 약속된 응답)
      child.stderr.setEncoding('utf8');
      child.stderr.on('data', (text: string) => {
        if (DEBUG) process.stderr.write(text);
      });
      child.on('error', () => finishStart(false));
      child.on('exit', () => {
        if (this.child === child) {
          this.child = null;
          this.ready = null;
          this.settling = false;
          this.isReady = false;
        }
        this.onReady?.(false);
        const job = this.job;
        if (job) {
          this.job = null;
          clearTimeout(job.timer);
          job.resolve(false);
        }
      });
    });
    return this.ready;
  }

  private onData(chunk: string): void {
    this.buffer += chunk;
    let newline: number;
    while ((newline = this.buffer.indexOf('\n')) >= 0) {
      const line = this.buffer.slice(0, newline).trim();
      this.buffer = this.buffer.slice(newline + 1);
      if (!line) continue;
      let message: { ready?: boolean; idle?: boolean; restarted?: string | null; fatal?: string; id?: number; ok?: boolean; error?: string };
      try {
        message = JSON.parse(line);
      } catch {
        continue;
      }
      if (message.ready !== undefined) {
        this.onReady?.(message.ready === true);
        continue;
      }
      if (message.idle) {
        this.settling = false;
        if (message.restarted && DEBUG) console.info(`[상주 변환기] ${message.restarted}`);
        if (message.fatal) console.warn(`상주 변환기를 다시 띄우지 못했다(예전 방식으로 계속): ${message.fatal}`);
        continue;
      }
      const job = this.job;
      if (job && message.id === job.id) {
        this.job = null;
        clearTimeout(job.timer);
        if (!message.ok) console.warn(`상주 변환기 실패(예전 방식으로 다시 변환, ${job.file}): ${message.error ?? ''}`);
        job.resolve(message.ok === true);
      }
    }
  }

  /** 변환기와 그 soffice를 끈다 — 변환기는 표준 입력이 닫히면 soffice를 끄고 끝난다 */
  kill(): void {
    const child = this.child;
    this.child = null;
    this.ready = null;
    this.buffer = '';
    this.settling = false;
    this.isReady = false;
    if (!child) return;
    try {
      child.stdin.end();
    } catch {
      /* 이미 닫힘 */
    }
    // 응답이 없으면 프로세스 트리째 끈다(Windows: soffice는 변환기의 자식이라 트리로 끈다)
    setTimeout(() => {
      if (child.exitCode !== null) return;
      if (process.platform === 'win32' && child.pid) spawn('taskkill', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true });
      else child.kill('SIGKILL');
    }, 5000).unref();
  }
}

/** LibreOffice와 함께 깔린 python(uno 모듈 포함)과 변환기 스크립트가 있을 때만 만든다. DOCVIEW_LO_WORKER=0이면 끈다.
 * (한때 기본으로 꺼 두었다 — 발표 문서를 닫으면 soffice가 죽는 문제를 "닫지 않고 모았다가 새로 띄우기"로 풀고 다시 켰다) */
export function createLibreOfficeWorker(sofficePath: string | null): LibreOfficeWorker | null {
  if (!sofficePath || process.env.DOCVIEW_LO_WORKER === '0') return null;
  const programDir = path.dirname(sofficePath);
  const python = [path.join(programDir, 'python.exe'), path.join(programDir, 'python')].find((p) => existsSync(p));
  const script = path.join(path.dirname(fileURLToPath(import.meta.url)), 'loWorker.py');
  if (!python || !existsSync(script)) return null;
  return new LibreOfficeWorker(python, script, sofficePath, process.env.DOCVIEW_PORT ?? '8000');
}
