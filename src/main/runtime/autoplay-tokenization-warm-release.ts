function normalizeMediaPath(mediaPath: string | null | undefined): string | null {
  if (typeof mediaPath !== 'string') {
    return null;
  }
  const trimmed = mediaPath.trim();
  return trimmed.length > 0 ? trimmed : null;
}

export function createAutoplayTokenizationWarmRelease(deps: {
  isTokenizationWarmupReady: () => boolean;
  startTokenizationWarmups: () => Promise<void>;
  getCurrentMediaPath: () => string | null | undefined;
  primeCurrentSubtitle?: (mediaPath: string) => void | Promise<void>;
  signalAutoplayReady: () => void;
  warn: (message: string, error: unknown) => void;
}): (mediaPath: string | null | undefined) => void {
  const signalIfCurrent = (mediaPath: string): void => {
    const currentMediaPath = normalizeMediaPath(deps.getCurrentMediaPath());
    if (!currentMediaPath || currentMediaPath !== mediaPath) {
      return;
    }
    deps.signalAutoplayReady();
  };

  return (mediaPath) => {
    const normalizedPath = normalizeMediaPath(mediaPath);
    if (!normalizedPath) {
      return;
    }
    try {
      void Promise.resolve(deps.primeCurrentSubtitle?.(normalizedPath)).catch((error) => {
        deps.warn('Startup subtitle priming failed before autoplay readiness release:', error);
      });
    } catch (error) {
      deps.warn('Startup subtitle priming failed before autoplay readiness release:', error);
    }
    if (deps.isTokenizationWarmupReady()) {
      signalIfCurrent(normalizedPath);
      return;
    }
    void deps
      .startTokenizationWarmups()
      .then(() => {
        signalIfCurrent(normalizedPath);
      })
      .catch((error) => {
        deps.warn('Startup tokenization warmup failed before autoplay readiness release:', error);
      });
  };
}
