import type { BrowserWindowConstructorOptions, Session } from 'electron';
import * as path from 'path';
import type { OverlayWindowKind } from './overlay-window-input';

export const OVERLAY_WINDOW_TITLES: Record<OverlayWindowKind, string> = {
  visible: 'SubMiner Overlay',
  modal: 'SubMiner Overlay Modal',
};

export function buildOverlayWindowOptions(
  kind: OverlayWindowKind,
  options: {
    isDev: boolean;
    yomitanSession?: Session | null;
  },
): BrowserWindowConstructorOptions {
  const showNativeDebugFrame = process.platform === 'win32' && options.isDev;
  const shouldStartAlwaysOnTop = !(process.platform === 'win32' && kind === 'visible');

  return {
    show: false,
    title: OVERLAY_WINDOW_TITLES[kind],
    width: 800,
    height: 600,
    x: 0,
    y: 0,
    transparent: true,
    backgroundColor: '#00000000',
    frame: false,
    alwaysOnTop: shouldStartAlwaysOnTop,
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
      backgroundThrottling: false,
      webSecurity: true,
      session: options.yomitanSession ?? undefined,
      additionalArguments: [`--overlay-layer=${kind}`],
    },
  };
}
