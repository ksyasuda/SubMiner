export function nowMs(): number {
  const perf = globalThis.performance;
  if (perf) {
    return Math.floor(perf.timeOrigin + perf.now());
  }

  return Number(process.hrtime.bigint() / 1000000n);
}
