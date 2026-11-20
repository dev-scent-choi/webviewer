/**
 * 문서별로 한 번에 하나만 도는 작업 + 실패 뒤 쉬기 — 보기 그리드 재생성(routes/viewer.ts)에 쓴다.
 * 같은 문서에 요청이 몰리면(자정 직후 TODAY() 그리드가 한꺼번에 낡는 등) 진행 중인 작업의 결과를 함께 기다리고,
 * 작업이 실패하면 쉬는 시간 동안은 다시 하지 않는다(실패한 재생성을 요청마다 되풀이해 서버가 바빠지지 않게).
 * 쉬는 시간은 부를 때 정한다 — 그리드가 아예 없어 보기 화면이 오류가 되는 경우는 짧게(일시 실패에서 빨리 회복).
 */
export function singleFlightWithBackoff(defaultBackoffMs: number, now: () => number = Date.now) {
  const inFlight = new Map<string, Promise<boolean>>();
  const failedUntil = new Map<string, number>();
  /** 쉬는 시간이 끝난 기록은 지운다(문서 수만큼 쌓이지 않게) */
  const prune = () => {
    const t = now();
    for (const [key, until] of failedUntil) if (until <= t) failedUntil.delete(key);
  };
  return {
    /** fn을 문서(key)마다 하나만 돌린다. 쉬는 중이면 false */
    run(key: string, fn: () => Promise<boolean>, backoffMs = defaultBackoffMs): Promise<boolean> {
      const until = failedUntil.get(key);
      if (until !== undefined && now() < until) return Promise.resolve(false);
      let p = inFlight.get(key);
      if (!p) {
        p = fn()
          .catch(() => false)
          .then((ok) => {
            prune();
            if (ok) failedUntil.delete(key);
            else failedUntil.set(key, now() + backoffMs);
            return ok;
          })
          .finally(() => inFlight.delete(key));
        inFlight.set(key, p);
      }
      return p;
    },
    /** 지금 쉬는 중인지(최근에 실패했는지) */
    backingOff(key: string): boolean {
      const until = failedUntil.get(key);
      return until !== undefined && now() < until;
    },
    /** 남아 있는 실패 기록 수(시험용) */
    size: () => failedUntil.size,
  };
}
