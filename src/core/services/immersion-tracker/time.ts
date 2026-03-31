declare global {
  var __subminerTestNowMs: number | string | undefined;
}

function getMockNowMs(testNowMs: number | string | undefined): number | null {
  if (typeof testNowMs === 'number' && Number.isFinite(testNowMs)) {
    return Math.floor(testNowMs);
  }
  if (typeof testNowMs === 'string') {
    const parsed = Number(testNowMs.trim());
    if (Number.isFinite(parsed)) {
      return Math.trunc(parsed);
    }
  }
  return null;
}

export function nowMs(): number {
  const mockedNowMs = getMockNowMs(globalThis.__subminerTestNowMs);
  if (mockedNowMs !== null) {
    return mockedNowMs;
  }

  const perf = globalThis.performance;
  if (perf && Number.isFinite(perf.timeOrigin)) {
    return Math.floor(perf.timeOrigin + perf.now());
  }

  return Date.now();
}
