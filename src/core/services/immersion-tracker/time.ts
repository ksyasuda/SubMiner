const SQLITE_SAFE_EPOCH_BASE_MS = 2_000_000_000;

export function nowMs(): number {
  const perf = globalThis.performance;
  if (perf) {
    return SQLITE_SAFE_EPOCH_BASE_MS + Math.floor(perf.now());
  }

  return SQLITE_SAFE_EPOCH_BASE_MS;
}
