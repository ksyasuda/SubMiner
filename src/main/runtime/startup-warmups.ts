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
  logDebug?: (message: string) => void;
}) {
  return (): void => {
    if (deps.getStarted()) return;
    if (deps.isTexthookerOnlyMode()) return;

    const warmupMecab = deps.shouldWarmupMecab();
    const warmupYomitanExtension = deps.shouldWarmupYomitanExtension();
    const warmupSubtitleDictionaries = deps.shouldWarmupSubtitleDictionaries();
    const warmupJellyfinRemoteSession = deps.shouldWarmupJellyfinRemoteSession();
    const autoConnectJellyfinRemote = deps.shouldAutoConnectJellyfinRemote();

    deps.setStarted(true);
    const shouldWarmupTokenization =
      warmupMecab || warmupYomitanExtension || warmupSubtitleDictionaries;
    if (shouldWarmupTokenization) {
      deps.launchTask('subtitle-tokenization', async () => {
        await Promise.all([
          warmupYomitanExtension
            ? (async () => {
                deps.logDebug?.('[startup-warmup] stage start: yomitan-extension');
                await deps.ensureYomitanExtensionLoaded();
                deps.logDebug?.('[startup-warmup] stage ready: yomitan-extension');
              })()
            : Promise.resolve().then(() => {
                deps.logDebug?.('[startup-warmup] stage skipped: yomitan-extension');
              }),
          warmupMecab
            ? (async () => {
                deps.logDebug?.('[startup-warmup] stage start: mecab');
                await deps.createMecabTokenizerAndCheck();
                deps.logDebug?.('[startup-warmup] stage ready: mecab');
              })()
            : Promise.resolve().then(() => {
                deps.logDebug?.('[startup-warmup] stage skipped: mecab');
              }),
          warmupSubtitleDictionaries
            ? (async () => {
                deps.logDebug?.('[startup-warmup] stage start: subtitle-dictionaries');
                await deps.prewarmSubtitleDictionaries();
                deps.logDebug?.('[startup-warmup] stage ready: subtitle-dictionaries');
              })()
            : Promise.resolve().then(() => {
                deps.logDebug?.('[startup-warmup] stage skipped: subtitle-dictionaries');
              }),
        ]);
      });
    }
    if (warmupJellyfinRemoteSession && autoConnectJellyfinRemote) {
      deps.launchTask('jellyfin-remote-session', async () => {
        deps.logDebug?.('[startup-warmup] stage start: jellyfin-remote-session');
        await deps.startJellyfinRemoteSession();
        deps.logDebug?.('[startup-warmup] stage ready: jellyfin-remote-session');
      });
    } else if (!warmupJellyfinRemoteSession) {
      deps.logDebug?.('[startup-warmup] stage skipped: jellyfin-remote-session (disabled)');
    } else {
      deps.logDebug?.('[startup-warmup] stage skipped: jellyfin-remote-session (auto-connect off)');
    }
  };
}
