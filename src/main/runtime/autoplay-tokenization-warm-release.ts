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

  const primeSubtitleForRelease = (mediaPath: string): Promise<void> | null => {
    if (!deps.primeCurrentSubtitle) {
      return null;
    }
    try {
      return Promise.resolve(deps.primeCurrentSubtitle(mediaPath)).catch((error) => {
        deps.warn('Startup subtitle priming failed before autoplay readiness release:', error);
      });
    } catch (error) {
      deps.warn('Startup subtitle priming failed before autoplay readiness release:', error);
      return null;
    }
  };

  return (mediaPath) => {
    const normalizedPath = normalizeMediaPath(mediaPath);
    if (!normalizedPath) {
      return;
    }
    const primePromise = primeSubtitleForRelease(normalizedPath);
    if (deps.isTokenizationWarmupReady()) {
      if (!primePromise) {
        signalIfCurrent(normalizedPath);
        return;
      }
      void primePromise.then(() => {
        signalIfCurrent(normalizedPath);
      });
      return;
    }
    const warmupPromise = deps.startTokenizationWarmups();
    const readinessPromise = primePromise
      ? Promise.all([primePromise, warmupPromise]).then(() => {})
      : warmupPromise;
    void readinessPromise
      .then(() => {
        signalIfCurrent(normalizedPath);
      })
      .catch((error) => {
        deps.warn('Startup tokenization warmup failed before autoplay readiness release:', error);
      });
  };
}
