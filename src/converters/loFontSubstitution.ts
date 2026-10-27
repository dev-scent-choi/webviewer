/**
 * LibreOffice 글꼴 대체표 — 서버에 없는 글꼴을 비슷한 모양의 글꼴로 그리게 한다.
 *
 * KoPub돋움체(한국출판인회의)는 "서버에 탑재해 웹서비스에 쓰려면 별도 승인"이 필요해 설치하지 않았다(2026-10-19).
 * 그러면 LibreOffice는 이 글꼴을 바탕(명조)으로 대체해, 미리보기·잘라 붙인 표 등이 원본(고딕)과 전혀 다르게 보였다.
 * 돋움은 고딕 계열이므로 맑은 고딕으로 대체하도록 프로필의 대체표(도구 → 옵션 → 글꼴과 같은 설정)에 넣는다 —
 * 글꼴 파일을 쓰는 게 아니라 이름만 바꿔 부르는 것이다.
 *
 * 변환에 쓰는 모든 프로필(명령 방식의 슬롯, 상주 변환기)에 soffice를 띄우기 전에 넣는다. 이미 있으면 건드리지 않는다.
 *
 * 같은 자리에 보안 설정도 하나 넣는다(BLOCK_LINKS): 올린 문서가 외부 그림을 링크로만 걸어 두면(file:///, UNC 공유,
 * http 내부 주소) 변환할 때 LibreOffice가 서버에서 그걸 읽어 PDF에 넣었다 — 서버 파일 유출·내부망 요청·NTLM 해시 유출.
 * 문서 안에 든 그림은 그대로 그려진다(2026-10-26 시험: 링크 그림 1→0장, 그림 60장 docx 는 결과가 같음).
 */
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import path from 'node:path';

const GOTHIC = 'Malgun Gothic';

/** [원본 글꼴 이름, 대신 쓸 글꼴] — 굵기별로 이름이 따로인 글꼴이라 각각 적는다 */
export const FONT_SUBSTITUTIONS: ReadonlyArray<readonly [string, string]> = [
  ['KoPub돋움체 Light', GOTHIC],
  ['KoPub돋움체 Medium', GOTHIC],
  ['KoPub돋움체 Bold', GOTHIC],
  ['KoPubWorld돋움체 Light', GOTHIC],
  ['KoPubWorld돋움체 Medium', GOTHIC],
  ['KoPubWorld돋움체 Bold', GOTHIC],
  ['KoPubWorld돋움체_Pro Light', GOTHIC],
  ['KoPubWorld돋움체_Pro Medium', GOTHIC],
  ['KoPubWorld돋움체_Pro Bold', GOTHIC],
];

const MARKER = 'docview_font_';
const BLOCK_LINKS =
  '<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse"><value>true</value></prop></item>';

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export function fontSubstitutionItems(): string {
  const enable =
    '<item oor:path="/org.openoffice.Office.Common/Font/Substitution"><prop oor:name="Replacement" oor:op="fuse"><value>true</value></prop></item>';
  const pairs = FONT_SUBSTITUTIONS.map(
    ([from, to], i) =>
      `<item oor:path="/org.openoffice.Office.Common/Font/Substitution/FontPairs"><node oor:name="${MARKER}${i}" oor:op="replace">` +
      '<prop oor:name="Always" oor:op="fuse"><value>true</value></prop>' +
      '<prop oor:name="OnScreenOnly" oor:op="fuse"><value>false</value></prop>' +
      `<prop oor:name="ReplaceFont" oor:op="fuse"><value>${escapeXml(from)}</value></prop>` +
      `<prop oor:name="SubstituteFont" oor:op="fuse"><value>${escapeXml(to)}</value></prop>` +
      '</node></item>',
  );
  return [BLOCK_LINKS, enable, ...pairs].join('\n');
}

/** 프로필 폴더(UserInstallation)의 user/registrymodifications.xcu에 대체표를 넣는다. 실패해도 변환은 계속한다. */
export function ensureFontSubstitutions(profileDir: string): void {
  try {
    const file = path.join(profileDir, 'user', 'registrymodifications.xcu');
    if (existsSync(file)) {
      const xml = readFileSync(file, 'utf8');
      if (xml.includes(`${MARKER}${FONT_SUBSTITUTIONS.length - 1}"`) && xml.includes(BLOCK_LINKS)) return;
      // 예전 판(개수가 다른 표, 링크 차단이 없던 판)이 있으면 지우고 새로 넣는다
      const cleaned = xml
        .replace(new RegExp(`<item oor:path="/org.openoffice.Office.Common/Font/Substitution/FontPairs"><node oor:name="${MARKER}\\d+"[\\s\\S]*?</item>\\n?`, 'g'), '')
        .replace(/<item oor:path="\/org\.openoffice\.Office\.Common\/Security\/Scripting"><prop oor:name="BlockUntrustedRefererLinks"[\s\S]*?<\/item>\n?/g, '');
      writeFileSync(file, cleaned.replace('</oor:items>', `${fontSubstitutionItems()}\n</oor:items>`), 'utf8');
      return;
    }
    // 아직 한 번도 안 띄운 프로필 — soffice가 처음 만들 때 이 파일을 읽어 합친다
    mkdirSync(path.dirname(file), { recursive: true });
    writeFileSync(
      file,
      '<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n' +
        `${fontSubstitutionItems()}\n</oor:items>\n`,
      'utf8',
    );
  } catch (error) {
    // shortcut: 프로필 파일을 못 쓰면 링크 차단도 빠진 채 변환한다 — 경고만 남김, 차단을 꼭 보장해야 하면 여기서 변환을 멈출 것
    console.warn('LibreOffice 프로필 설정(글꼴 대체·링크 차단)을 쓰지 못했습니다:', error);
  }
}
