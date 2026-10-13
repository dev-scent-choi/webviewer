#!/usr/bin/env node
// `npm install`(또는 `npx`)로 설치했을 때의 실행 진입점. dist/server.js는 npm pack/publish 전에
// prepack(package.json)이 이미 컴파일해 둔 결과물이라, 여기서는 그걸 그대로 실행하기만 한다 —
// tsx/esbuild 둘 다 필요 없다(consumer 쪽 설치에 그 두 패키지가 없어도 된다).
//
// 사용자가 어느 디렉터리에서 실행하든(전역 설치 후 아무 데서나 `docview`를 치는 경우 포함) 항상
// "설치된 패키지 자기 자신의 폴더"를 기준으로 동작해야 한다 — package.json의 "scripts"는 npm이
// 자동으로 CWD를 패키지 루트로 맞춰 주지만, bin 스크립트는 사용자의 현재 CWD에서 그대로 실행되므로
// 그 보장이 없다.
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { spawn } from 'node:child_process';

const packageRoot = dirname(dirname(fileURLToPath(import.meta.url)));

const child = spawn(process.execPath, [join(packageRoot, 'dist', 'server.js')], {
  cwd: packageRoot,
  stdio: 'inherit',
});
child.on('exit', (code) => process.exit(code ?? 1));
child.on('error', (err) => {
  console.error(err);
  process.exit(1);
});
