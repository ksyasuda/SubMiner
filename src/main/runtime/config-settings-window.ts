export interface ConfigSettingsWindowLike {
  isDestroyed(): boolean;
  focus(): void;
  loadFile(path: string): unknown;
  on(event: 'closed', handler: () => void): unknown;
}

export interface OpenConfigSettingsWindowDeps<TWindow extends ConfigSettingsWindowLike> {
  getSettingsWindow(): TWindow | null;
  setSettingsWindow(window: TWindow | null): void;
  createSettingsWindow(): TWindow;
  settingsHtmlPath: string;
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
    window.loadFile(deps.settingsHtmlPath);
    deps.setSettingsWindow(window);
    window.on('closed', () => {
      deps.setSettingsWindow(null);
    });
    window.focus();
    return true;
  };
}
