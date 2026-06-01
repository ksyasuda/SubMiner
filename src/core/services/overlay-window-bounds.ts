import type { WindowGeometry } from '../../types';

type ScreenDipConverter = {
  screenToDipRect: (
    window: Electron.BrowserWindow | null,
    rect: Electron.Rectangle,
  ) => Electron.Rectangle;
};

type ContentBoundsWindow = {
  isDestroyed: () => boolean;
  getBounds: () => Electron.Rectangle;
  getContentBounds: () => Electron.Rectangle;
};

function resolveContentAlignedBounds(
  geometry: WindowGeometry,
  window?: ContentBoundsWindow | null,
): WindowGeometry {
  if (!window || window.isDestroyed()) {
    return geometry;
  }

  let outer: Electron.Rectangle;
  let content: Electron.Rectangle;
  try {
    outer = window.getBounds();
    content = window.getContentBounds();
  } catch {
    return geometry;
  }

  const leftInset = content.x - outer.x;
  const topInset = content.y - outer.y;
  const rightInset = outer.x + outer.width - (content.x + content.width);
  const bottomInset = outer.y + outer.height - (content.y + content.height);
  const insets = [leftInset, topInset, rightInset, bottomInset];
  if (insets.some((inset) => !Number.isFinite(inset) || inset < 0)) {
    return geometry;
  }

  return {
    x: geometry.x - leftInset,
    y: geometry.y - topInset,
    width: geometry.width + leftInset + rightInset,
    height: geometry.height + topInset + bottomInset,
  };
}

export function normalizeOverlayWindowBoundsForPlatform(
  geometry: WindowGeometry,
  platform: NodeJS.Platform,
  screen: ScreenDipConverter | null,
  window?: ContentBoundsWindow | null,
): WindowGeometry {
  if (platform === 'linux') {
    return resolveContentAlignedBounds(geometry, window);
  }

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
