/**
 * 문서별로 한 번에 하나만 도는 작업 + 실패 뒤 쉬기 — 보기 그리드 재생성(routes/viewer.ts)에 쓴다.
 * 같은 문서에 요청이 몰리면(자정 직후 TODAY() 그리드가 한꺼번에 낡는 등) 진행 중인 작업의 결과를 함께 기다리고,
 * 작업이 실패하면 backoffMs 동안은 다시 하지 않는다(실패한 재생성을 요청마다 되풀이해 서버가 바빠지지 않게).
 */
export function singleFlightWithBackoff(backoffMs: number, now: () => number = Date.now) {
  const inFlight = new Map<string, Promise<boolean>>();
  const failedUntil = new Map<string, number>();
  return {
    /** fn을 문서(key)마다 하나만 돌린다. 쉬는 중이면 false */
    run(key: string, fn: () => Promise<boolean>): Promise<boolean> {
      const until = failedUntil.get(key);
      if (until !== undefined && now() < until) return Promise.resolve(false);
      let p = inFlight.get(key);
      if (!p) {
        p = fn()
          .catch(() => false)
          .then((ok) => {
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
  };
}
