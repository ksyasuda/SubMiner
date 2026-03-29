export function nowMs(): number {
  const perf = globalThis.performance;
  if (perf && Number.isFinite(perf.timeOrigin)) {
    return Math.floor(perf.timeOrigin + perf.now());
  }

  return Date.now();
}
