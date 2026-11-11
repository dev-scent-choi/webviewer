/**
 * "쪽을 준비하는 중"(설계 V8) — 첫 변환(editReady)이 끝나지 않은 문서. 쪽 자리에 안내와 진행 막대를 두고 /api/view/status를
 * 1.5초마다 본다(감사 기록 없음). 그동안 편집 단추는 이미 쓸 수 있다. 끝나면 resolve, 실패하면 그 문구로 reject.
 */
const POLL_MS = 1500;
/** 이만큼 지나도 끝나지 않으면 멈추고 안내한다(서버가 처리 중에 멈춘 경우 끝없이 묻지 않게) */
const PREP_TIMEOUT_MS = 10 * 60_000;

interface StatusData {
  readonly status: 'processing' | 'ready' | 'failed';
  readonly progress: number;
  readonly error?: string;
}

export async function waitUntilPrepared(tokenQuery: string): Promise<void> {
  const screen = document.getElementById('prep-screen') as HTMLElement;
  const bar = document.getElementById('prep-progress') as HTMLProgressElement;
  screen.hidden = false;
  const deadline = Date.now() + PREP_TIMEOUT_MS;
  try {
    for (;;) {
      if (Date.now() > deadline) throw new Error('쪽 준비가 너무 오래 걸립니다. 잠시 뒤 다시 열어 주세요.');
      const res = await fetch(`/api/view/status?${tokenQuery}`, { cache: 'no-store' }).catch(() => null);
      const json = res ? ((await res.json().catch(() => null)) as { success?: boolean; data?: StatusData; error?: string } | null) : null;
      if (res && !res.ok) throw new Error(json?.error || '문서를 열 수 없습니다.');
      const data = json?.data;
      if (data?.status === 'ready') return;
      if (data?.status === 'failed') throw new Error(data.error || '문서를 변환하지 못했습니다.');
      if (data) bar.value = data.progress;
      await new Promise((resolve) => setTimeout(resolve, POLL_MS));
    }
  } finally {
    screen.hidden = true;
  }
}
