// 시험은 워커마다 자기 LibreOffice 프로필 폴더를 쓴다 — 운영 서버(같은 임시 폴더의 slot-0·1)·다른 시험 파일과 프로필 잠금이 부딪치지 않게
import { rmSync } from 'node:fs';
import os from 'node:os';
import path from 'node:path';

const root = path.join(os.tmpdir(), `docview-lo-profiles-test-${process.pid}`);
process.env.DOCVIEW_LO_PROFILE_ROOT = root;
process.on('exit', () => rmSync(root, { recursive: true, force: true }));
