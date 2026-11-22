// 아트크래프트 리본 꾸밈 규격(설계 2026-11-22-artcraft-ribbon-skin §4, 캡처 실측 2026-11-22)과 순수 판정 — scripts/verify-ribbon-skin-e2e.mjs가 쓴다.
// 숫자는 public/css/style.css 의 --skin-title-h·--skin-tabs-h·--skin-ribbon-h·--skin-status-h 와 같은 값이어야 한다.

export const SKIN_PX = Object.freeze({ title: 24, tabs: 22, ribbon: 62, status: 18 });
export const SKIN_TOLERANCE_PX = 2;

/** 잰 높이 중 규격 ±2px 밖인 칸 — 값이 없는 칸(null·undefined)은 그 화면에 없는 줄이라 건너뛴다 */
export function heightProblems(measured) {
  const out = [];
  for (const [key, want] of Object.entries(SKIN_PX)) {
    const got = measured[key];
    if (got === null || got === undefined) continue;
    if (Math.abs(got - want) > SKIN_TOLERANCE_PX) out.push(`${key} ${got}px (규격 ${want}±${SKIN_TOLERANCE_PX})`);
  }
  return out;
}

/** 영문자는 있는데 한글이 하나도 없는 라벨 — 숫자·기호만 있는 것(100%, 1.15, ▲)은 라벨 문제가 아니다 */
export function latinOnlyLabels(labels) {
  return labels.map((label) => label.trim()).filter((label) => /[A-Za-z]/.test(label) && !/[가-힣]/.test(label));
}

/** 큰 단추: 아이콘 아래 끝이 이름 위 끝보다 위(1px 반올림 허용), 둘의 가로 가운데가 3px 안 */
export function bigButtonLayoutOk(icon, label) {
  const iconMid = icon.left + icon.width / 2;
  const labelMid = label.left + label.width / 2;
  return icon.top + icon.height <= label.top + 1 && Math.abs(iconMid - labelMid) <= 3;
}
