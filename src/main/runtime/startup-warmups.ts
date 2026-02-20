export function createLaunchBackgroundWarmupTaskHandler(deps: {
  now: () => number;
  logDebug: (message: string) => void;
  logWarn: (message: string) => void;
}) {
  return (label: string, task: () => Promise<void>): void => {
    const startedAtMs = deps.now();
    void task()
      .then(() => {
        deps.logDebug(`[startup-warmup] ${label} completed in ${deps.now() - startedAtMs}ms`);
      })
      .catch((error) => {
        const message = error instanceof Error ? error.message : String(error);
        deps.logWarn(`[startup-warmup] ${label} failed: ${message}`);
      });
  };
}

export function createStartBackgroundWarmupsHandler(deps: {
  getStarted: () => boolean;
  setStarted: (started: boolean) => void;
  isTexthookerOnlyMode: () => boolean;
  launchTask: (label: string, task: () => Promise<void>) => void;
  createMecabTokenizerAndCheck: () => Promise<void>;
  ensureYomitanExtensionLoaded: () => Promise<void>;
  prewarmSubtitleDictionaries: () => Promise<void>;
  shouldAutoConnectJellyfinRemote: () => boolean;
  startJellyfinRemoteSession: () => Promise<void>;
}) {
  return (): void => {
    if (deps.getStarted()) return;
    if (deps.isTexthookerOnlyMode()) return;

    deps.setStarted(true);
    deps.launchTask('mecab', async () => {
      await deps.createMecabTokenizerAndCheck();
    });
    deps.launchTask('yomitan-extension', async () => {
      await deps.ensureYomitanExtensionLoaded();
    });
    deps.launchTask('subtitle-dictionaries', async () => {
      await deps.prewarmSubtitleDictionaries();
    });
    if (deps.shouldAutoConnectJellyfinRemote()) {
      deps.launchTask('jellyfin-remote-session', async () => {
        await deps.startJellyfinRemoteSession();
      });
    }
  };
}
