type YomitanExtensionLike = unknown;
type BrowserWindowLike = unknown;

export function createOpenYomitanSettingsHandler(deps: {
  ensureYomitanExtensionLoaded: () => Promise<YomitanExtensionLike | null>;
  openYomitanSettingsWindow: (params: {
    yomitanExt: YomitanExtensionLike;
    getExistingWindow: () => BrowserWindowLike | null;
    setWindow: (window: BrowserWindowLike | null) => void;
  }) => void;
  getExistingWindow: () => BrowserWindowLike | null;
  setWindow: (window: BrowserWindowLike | null) => void;
  logWarn: (message: string) => void;
  logError: (message: string, error: unknown) => void;
}) {
  return (): void => {
    void (async () => {
      const extension = await deps.ensureYomitanExtensionLoaded();
      if (!extension) {
        deps.logWarn('Unable to open Yomitan settings: extension failed to load.');
        return;
      }
      deps.openYomitanSettingsWindow({
        yomitanExt: extension,
        getExistingWindow: deps.getExistingWindow,
        setWindow: deps.setWindow,
      });
    })().catch((error) => {
      deps.logError('Failed to open Yomitan settings window.', error);
    });
  };
}
