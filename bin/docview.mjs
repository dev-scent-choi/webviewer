#!/usr/bin/env node
// `npm install`(또는 `npx`)로 설치했을 때의 실행 진입점 — `npm start`(`build:editor && tsx
// src/server.ts`)와 정확히 같은 일을 한다. 유일한 차이는 이 스크립트가 사용자가 어느 디렉터리에서
// 실행하든(전역 설치 후 아무 데서나 `docview`를 치는 경우 포함) 항상 "설치된 패키지 자기 자신의
// 폴더"를 기준으로 동작해야 한다는 점이다 — package.json의 "scripts"는 npm이 자동으로 CWD를
// 패키지 루트로 맞춰 주지만, bin 스크립트는 사용자의 현재 CWD에서 그대로 실행되므로 그 보장이 없다.
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { createRequire } from 'node:module';
import { spawn } from 'node:child_process';

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)));

function run(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: packageRoot, stdio: 'inherit' });
    child.on('exit', (code) => (code === 0 ? resolve() : reject(new Error(`${command} ${args.join(' ')}가 ${code}로 종료됨`))));
    child.on('error', reject);
  });
}

// tsx가 docview 바로 밑 node_modules에 중첩돼 있을지, 설치한 쪽 최상위 node_modules로
// 호이스팅됐을지는 npm이 그때그때 정한다(버전 충돌 여부에 따라 달라짐) — 그래서 경로를
// 하드코딩하지 않고 Node의 실제 모듈 해석 알고리즘(require.resolve)으로 tsx 패키지 자신의
// 위치를 찾은 뒤 그 밑의 dist/cli.mjs를 가리킨다.
const require = createRequire(import.meta.url);
const tsxCli = join(dirname(require.resolve('tsx/package.json')), 'dist', 'cli.mjs');

await run(process.execPath, [join(packageRoot, 'scripts', 'build-editor.mjs')]);
await run(process.execPath, [tsxCli, join(packageRoot, 'src', 'server.ts')]);
