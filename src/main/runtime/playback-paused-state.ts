type PlaybackPausedMpvClient = {
  connected?: boolean;
  requestProperty?: (name: string) => Promise<unknown>;
};

function coercePlaybackPaused(value: unknown): boolean | null {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value !== 0;
  if (typeof value === 'string') {
    const normalized = value.trim().toLowerCase();
    if (normalized === 'no' || normalized === 'false' || normalized === '0') return false;
    if (normalized === 'yes' || normalized === 'true' || normalized === '1') return true;
  }
  return null;
}

export async function resolveFreshPlaybackPaused(deps: {
  getCachedPlaybackPaused: () => boolean | null;
  getMpvClient: () => PlaybackPausedMpvClient | null;
}): Promise<boolean | null> {
  const cachedPaused = deps.getCachedPlaybackPaused();
  if (cachedPaused === true) {
    return true;
  }

  const client = deps.getMpvClient();
  if (client?.connected === true && typeof client.requestProperty === 'function') {
    try {
      const livePaused = coercePlaybackPaused(await client.requestProperty('pause'));
      if (livePaused !== null) {
        return livePaused;
      }
    } catch {
      // Avoid trusting a stale cached "playing" state for hover auto-pause.
    }
  }

  return cachedPaused === false ? null : cachedPaused;
}
