import type { WindowGeometry } from '../../types';

export const SECONDARY_OVERLAY_MAX_HEIGHT_RATIO = 0.2;

function toInteger(value: number): number {
  return Number.isFinite(value) ? Math.round(value) : 0;
}

function clampPositive(value: number): number {
  return Math.max(1, toInteger(value));
}

export function splitOverlayGeometryForSecondaryBar(geometry: WindowGeometry): {
  secondary: WindowGeometry;
  primary: WindowGeometry;
} {
  const x = toInteger(geometry.x);
  const y = toInteger(geometry.y);
  const width = clampPositive(geometry.width);
  const totalHeight = clampPositive(geometry.height);

  const secondaryHeight = clampPositive(
    Math.min(totalHeight, Math.round(totalHeight * SECONDARY_OVERLAY_MAX_HEIGHT_RATIO)),
  );
  const primaryHeight = clampPositive(totalHeight - secondaryHeight);

  return {
    secondary: {
      x,
      y,
      width,
      height: secondaryHeight,
    },
    primary: {
      x,
      y: y + secondaryHeight,
      width,
      height: primaryHeight,
    },
  };
}
