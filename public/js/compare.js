// 같은 문서를 리브레오피스(서버 변환)와 아트크래프트(브라우저 엔진)로 나란히 — 메인 화면의 "아트크래프트와 나란히 비교" 체크박스가 연다.
// 토큰 발급은 메인 화면(main.js issueToken)과 같은 경로·권한(올린 사람 본인 화면이라 전체 권한, 편집은 끔).
const docId = new URLSearchParams(location.search).get('doc');
const errEl = document.getElementById('err');

async function issue(engine) {
  const res = await fetch(`/api/documents/${encodeURIComponent(docId)}/tokens`, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ viewer: '비교', canPrint: true, canDownload: true, canCopy: true, canEdit: false, ...(engine === 'libreoffice' ? {} : { engine }) }),
  });
  const json = await res.json();
  if (!res.ok || !json.success) throw new Error(json.error || `요청 실패 (HTTP ${res.status})`);
  return json.data;
}

(async () => {
  if (!docId) throw new Error('문서가 지정되지 않았습니다.');
  const meta = await fetch(`/api/documents/${encodeURIComponent(docId)}`).then((r) => r.json());
  document.getElementById('name').textContent = meta?.data?.name ?? docId;
  const [lo, ac] = await Promise.all([issue('libreoffice'), issue('artcraft')]);
  document.getElementById('left').src = lo.viewerUrl;
  document.getElementById('right').src = ac.viewerUrl;
  if (ac.engine !== 'artcraft') document.getElementById('right-note').textContent = `서버 변환으로 대신 엶 — ${ac.engineReason ?? ''}`;
})().catch((error) => {
  errEl.textContent = error instanceof Error ? error.message : String(error);
});
