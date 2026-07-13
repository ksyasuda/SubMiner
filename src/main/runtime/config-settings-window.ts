export interface ConfigSettingsWindowLike {
  isDestroyed(): boolean;
  show(): void;
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
  promoteSettingsWindowAboveOverlay?: (window: TWindow) => void;
  onClosed?: () => void;
  log?: (message: string) => void;
}

export function createOpenConfigSettingsWindowHandler<TWindow extends ConfigSettingsWindowLike>(
  deps: OpenConfigSettingsWindowDeps<TWindow>,
): () => boolean {
  return () => {
    const showAndFocus = (window: TWindow): void => {
      window.show();
      window.focus();
      deps.promoteSettingsWindowAboveOverlay?.(window);
    };

    const existing = deps.getSettingsWindow();
    if (existing && !existing.isDestroyed()) {
      showAndFocus(existing);
      return true;
    }

    const window = deps.createSettingsWindow();
    void Promise.resolve(window.loadFile(deps.settingsHtmlPath)).catch((error) => {
      const message = error instanceof Error ? error.message : String(error);
      deps.log?.(`Failed to load settings window: ${message}`);
      deps.setSettingsWindow(null);
      window.destroy?.();
    });
    deps.setSettingsWindow(window);
    window.on('closed', () => {
      deps.setSettingsWindow(null);
      deps.onClosed?.();
    });
    showAndFocus(window);
    return true;
  };
}
