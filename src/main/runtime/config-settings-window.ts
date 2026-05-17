export interface ConfigSettingsWindowLike {
  isDestroyed(): boolean;
  focus(): void;
  loadFile(path: string): unknown;
  on(event: 'closed', handler: () => void): unknown;
  destroy?(): unknown;
}

export interface OpenConfigSettingsWindowDeps<TWindow extends ConfigSettingsWindowLike> {
  getSettingsWindow(): TWindow | null;
  setSettingsWindow(window: TWindow | null): void;
  createSettingsWindow(): TWindow;
  settingsHtmlPath: string;
  log?: (message: string) => void;
}

export function createOpenConfigSettingsWindowHandler<TWindow extends ConfigSettingsWindowLike>(
  deps: OpenConfigSettingsWindowDeps<TWindow>,
): () => boolean {
  return () => {
    const existing = deps.getSettingsWindow();
    if (existing && !existing.isDestroyed()) {
      existing.focus();
      return true;
    }

    const window = deps.createSettingsWindow();
    void Promise.resolve(window.loadFile(deps.settingsHtmlPath)).catch((error) => {
      const message = error instanceof Error ? error.message : String(error);
      deps.log?.(`Failed to load configuration settings window: ${message}`);
      deps.setSettingsWindow(null);
      window.destroy?.();
    });
    deps.setSettingsWindow(window);
    window.on('closed', () => {
      deps.setSettingsWindow(null);
    });
    window.focus();
    return true;
  };
}
