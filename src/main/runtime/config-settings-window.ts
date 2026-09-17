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
  // macOS only: showing/focusing a window does not activate a background app, so the window
  // opens behind whatever is frontmost. See activateMacOSApp.
  activateApp?: () => void;
  onClosed?: () => void;
  log?: (message: string) => void;
}

export function createOpenConfigSettingsWindowHandler<TWindow extends ConfigSettingsWindowLike>(
  deps: OpenConfigSettingsWindowDeps<TWindow>,
): () => boolean {
  return () => {
    const showAndFocus = (window: TWindow): void => {
      window.show();
      // Activate the app before focusing: the window can only become key once the app itself
      // is frontmost.
      deps.activateApp?.();
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
