import type { BrowserWindowConstructorOptions } from 'electron';

import type { PluginWindowRequest } from '../../core/services/plugins/types';

interface PluginWindowLike {
  isDestroyed(): boolean;
  show(): void;
  focus(): void;
  loadFile(filePath: string): Promise<void>;
  destroy(): void;
  once(event: 'closed', listener: () => void): unknown;
}

/**
 * Creates plugin windows and keeps one per request key, so a plugin that reopens the same page
 * focuses the existing window. Windows are raised above the overlay like the settings windows.
 */
export function createPluginWindowRuntime<TWindow extends PluginWindowLike>(deps: {
  createBrowserWindow: (options: BrowserWindowConstructorOptions) => TWindow;
  promoteAboveOverlay: (window: TWindow) => void;
  logError: (message: string) => void;
}) {
  const windows = new Map<string, TWindow>();

  const showAndFocus = (window: TWindow): void => {
    window.show();
    window.focus();
    deps.promoteAboveOverlay(window);
  };

  return {
    open(request: PluginWindowRequest): { window: TWindow; created: boolean } {
      const existing = windows.get(request.key);
      if (existing && !existing.isDestroyed()) {
        showAndFocus(existing);
        return { window: existing, created: false };
      }

      const window = deps.createBrowserWindow({
        width: request.width,
        height: request.height,
        title: request.title,
        show: false,
        autoHideMenuBar: true,
        webPreferences: {
          nodeIntegration: false,
          contextIsolation: true,
          ...(request.preloadPath ? { preload: request.preloadPath } : {}),
        },
      });
      windows.set(request.key, window);
      window.once('closed', () => {
        if (windows.get(request.key) === window) windows.delete(request.key);
      });
      window.loadFile(request.htmlPath).catch((error: unknown) => {
        const message = error instanceof Error ? error.message : String(error);
        deps.logError(`Failed to load plugin window ${request.htmlPath}: ${message}`);
        if (!window.isDestroyed()) window.destroy();
      });
      showAndFocus(window);
      return { window, created: true };
    },
  };
}
