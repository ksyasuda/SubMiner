type YomitanExtensionLike = unknown;
type BrowserWindowLike = unknown;
type SessionLike = unknown;

export function createOpenYomitanSettingsHandler(deps: {
  ensureYomitanExtensionLoaded: () => Promise<YomitanExtensionLike | null>;
  getYomitanExtension?: () => YomitanExtensionLike | null;
  getYomitanExtensionLoadInFlight?: () => Promise<unknown> | null;
  openYomitanSettingsWindow: (params: {
    yomitanExt: YomitanExtensionLike;
    getExistingWindow: () => BrowserWindowLike | null;
    setWindow: (window: BrowserWindowLike | null) => void;
    yomitanSession?: SessionLike | null;
    onWindowClosed?: () => void;
  }) => void;
  getExistingWindow: () => BrowserWindowLike | null;
  setWindow: (window: BrowserWindowLike | null) => void;
  getYomitanSession?: () => SessionLike | null;
  logWarn: (message: string) => void;
  logError: (message: string, error: unknown) => void;
}) {
  return (): void => {
    void (async () => {
      if (deps.getYomitanExtension) {
        let loadedExtension = deps.getYomitanExtension();
        if (!loadedExtension) {
          if (deps.getYomitanExtensionLoadInFlight?.()) {
            deps.logWarn(
              'Yomitan settings requested while Yomitan is still loading. Try again in a few seconds.',
            );
            return;
          }
          loadedExtension = await deps.ensureYomitanExtensionLoaded();
          if (!loadedExtension) {
            deps.logWarn('Unable to open Yomitan settings: extension failed to load.');
            return;
          }
        }

        const yomitanSession = deps.getYomitanSession?.() ?? null;
        deps.openYomitanSettingsWindow({
          yomitanExt: loadedExtension,
          getExistingWindow: deps.getExistingWindow,
          setWindow: deps.setWindow,
          yomitanSession,
        });
        return;
      }

      const extension = await deps.ensureYomitanExtensionLoaded();
      if (!extension) {
        deps.logWarn('Unable to open Yomitan settings: extension failed to load.');
        return;
      }
      const yomitanSession = deps.getYomitanSession?.() ?? null;
      deps.openYomitanSettingsWindow({
        yomitanExt: extension,
        getExistingWindow: deps.getExistingWindow,
        setWindow: deps.setWindow,
        yomitanSession,
      });
    })().catch((error) => {
      deps.logError('Failed to open Yomitan settings window.', error);
    });
  };
}
