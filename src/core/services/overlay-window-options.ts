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
    linuxX11FullscreenOverlay?: boolean;
    platform?: NodeJS.Platform;
    yomitanSession?: Session | null;
  },
): BrowserWindowConstructorOptions {
  const platform = options.platform ?? process.platform;
  const showNativeDebugFrame = platform === 'win32' && options.isDev;
  const isLinuxVisibleOverlay = platform === 'linux' && kind === 'visible';
  const isLinuxFullscreenOverlay =
    isLinuxVisibleOverlay && options.linuxX11FullscreenOverlay === true;
  const shouldStartAlwaysOnTop =
    !(platform === 'win32' && kind === 'visible') &&
    (!isLinuxVisibleOverlay || isLinuxFullscreenOverlay);
  const shouldAllowCompositorResize = isLinuxVisibleOverlay && !isLinuxFullscreenOverlay;

  return {
    show: false,
    title: OVERLAY_WINDOW_TITLES[kind],
    width: 800,
    height: 600,
    x: 0,
    y: 0,
    transparent: true,
    paintWhenInitiallyHidden: true,
    backgroundColor: '#00000000',
    frame: false,
    alwaysOnTop: shouldStartAlwaysOnTop,
    skipTaskbar: true,
    resizable: shouldAllowCompositorResize,
    hasShadow: false,
    focusable: !isLinuxFullscreenOverlay,
    acceptFirstMouse: true,
    // A macOS panel is a fullscreen auxiliary window, so modal surfaces stay on the
    // active mpv Space instead of opening on SubMiner's last regular desktop.
    ...(platform === 'darwin' && kind === 'modal' ? { type: 'panel' as const } : {}),
    ...(platform === 'win32' ? { thickFrame: showNativeDebugFrame } : {}),
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
