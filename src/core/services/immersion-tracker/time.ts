declare global {
  var __subminerTestNowMs: number | string | undefined;
}

export function nowMs(): number {
  const testNowMs = globalThis.__subminerTestNowMs;
  if (typeof testNowMs === 'number' && Number.isFinite(testNowMs)) {
    return Math.floor(testNowMs);
  }

  const perf = globalThis.performance;
  if (perf && Number.isFinite(perf.timeOrigin)) {
    return Math.floor(perf.timeOrigin + perf.now());
  }

  return Date.now();
}
