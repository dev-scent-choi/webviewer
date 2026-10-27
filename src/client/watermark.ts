/** 화면 열람용 워터마크(보는 사람 이름 + 시각) — docx·xlsx 편집기가 함께 쓴다. public/js/viewer.js 는 번들을 거치지 않아 따로 있다. */
function escapeXmlText(value: string): string {
  return value.replace(/[<>&'"]/g, (ch) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', "'": '&apos;', '"': '&quot;' })[ch] as string);
}

/** "2026-09-06 23:30:11" — toISOString은 UTC라 보는 사람의 시계와 어긋나 보여 로컬 시각으로 조립한다 */
function formatWatermarkTimestamp(date: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0');
  return (
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ` +
    `${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`
  );
}

export function applyWatermark(viewerName: string): void {
  if (!viewerName) return;
  const text = `${viewerName}  ${formatWatermarkTimestamp(new Date())}`;
  const svg =
    `<svg xmlns="http://www.w3.org/2000/svg" width="300" height="180">` +
    `<text x="150" y="100" transform="rotate(-28 150 100)" text-anchor="middle" ` +
    `font-family="sans-serif" font-size="14" fill="rgba(0,0,0,0.16)">${escapeXmlText(text)}</text>` +
    `</svg>`;
  document.documentElement.style.setProperty('--watermark-bg', `url("data:image/svg+xml,${encodeURIComponent(svg)}")`);
}
