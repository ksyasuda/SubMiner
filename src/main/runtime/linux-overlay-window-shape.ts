export type LinuxOverlayShapeRect = {
  x: number;
  y: number;
  width: number;
  height: number;
};

export type LinuxOverlayShapeWindow = {
  isDestroyed: () => boolean;
  getBounds?: () => LinuxOverlayShapeRect;
  setShape?: (rects: LinuxOverlayShapeRect[]) => void;
};

function toPositivePixel(value: number): number | null {
  if (!Number.isFinite(value) || value <= 0) {
    return null;
  }
  return Math.max(1, Math.round(value));
}

export function buildFullWindowShapeRect(
  bounds: LinuxOverlayShapeRect,
): LinuxOverlayShapeRect | null {
  const width = toPositivePixel(bounds.width);
  const height = toPositivePixel(bounds.height);
  if (width === null || height === null) {
    return null;
  }

  return {
    x: 0,
    y: 0,
    width,
    height,
  };
}

export function restoreLinuxOverlayWindowShape(window: LinuxOverlayShapeWindow | null): boolean {
  if (!window || window.isDestroyed()) {
    return false;
  }
  if (typeof window.setShape !== 'function' || typeof window.getBounds !== 'function') {
    return false;
  }

  const rect = buildFullWindowShapeRect(window.getBounds());
  if (!rect) {
    return false;
  }

  window.setShape([rect]);
  return true;
}
