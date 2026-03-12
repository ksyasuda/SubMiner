import type { BrowserWindowConstructorOptions, Session } from 'electron';
import * as path from 'path';
import type { OverlayWindowKind } from './overlay-window-input';

export function buildOverlayWindowOptions(
  kind: OverlayWindowKind,
  options: {
    isDev: boolean;
    yomitanSession?: Session | null;
  },
): BrowserWindowConstructorOptions {
  const showNativeDebugFrame = process.platform === 'win32' && options.isDev;

  return {
    show: false,
    width: 800,
    height: 600,
    x: 0,
    y: 0,
    transparent: true,
    frame: false,
    alwaysOnTop: true,
    skipTaskbar: true,
    resizable: false,
    hasShadow: false,
    focusable: true,
    acceptFirstMouse: true,
    ...(process.platform === 'win32' ? { thickFrame: showNativeDebugFrame } : {}),
    webPreferences: {
      preload: path.join(__dirname, '..', '..', 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      webSecurity: true,
      session: options.yomitanSession ?? undefined,
      additionalArguments: [`--overlay-layer=${kind}`],
    },
  };
}
