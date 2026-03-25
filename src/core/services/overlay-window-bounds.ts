import type { WindowGeometry } from '../../types';

type ScreenDipConverter = {
  screenToDipRect: (
    window: Electron.BrowserWindow | null,
    rect: Electron.Rectangle,
  ) => Electron.Rectangle;
};

export function normalizeOverlayWindowBoundsForPlatform(
  geometry: WindowGeometry,
  platform: NodeJS.Platform,
  screen: ScreenDipConverter | null,
): WindowGeometry {
  if (platform !== 'win32' || !screen) {
    return geometry;
  }

  return screen.screenToDipRect(null, {
    x: geometry.x,
    y: geometry.y,
    width: geometry.width,
    height: geometry.height,
  });
}
