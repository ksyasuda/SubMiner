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
  shouldWarmupMecab: () => boolean;
  shouldWarmupYomitanExtension: () => boolean;
  shouldWarmupSubtitleDictionaries: () => boolean;
  shouldWarmupJellyfinRemoteSession: () => boolean;
  shouldAutoConnectJellyfinRemote: () => boolean;
  startJellyfinRemoteSession: () => Promise<void>;
}) {
  return (): void => {
    if (deps.getStarted()) return;
    if (deps.isTexthookerOnlyMode()) return;

    deps.setStarted(true);
    if (deps.shouldWarmupMecab()) {
      deps.launchTask('mecab', async () => {
        await deps.createMecabTokenizerAndCheck();
      });
    }
    if (deps.shouldWarmupYomitanExtension()) {
      deps.launchTask('yomitan-extension', async () => {
        await deps.ensureYomitanExtensionLoaded();
      });
    }
    if (deps.shouldWarmupSubtitleDictionaries()) {
      deps.launchTask('subtitle-dictionaries', async () => {
        await deps.prewarmSubtitleDictionaries();
      });
    }
    if (deps.shouldWarmupJellyfinRemoteSession() && deps.shouldAutoConnectJellyfinRemote()) {
      deps.launchTask('jellyfin-remote-session', async () => {
        await deps.startJellyfinRemoteSession();
      });
    }
  };
}
