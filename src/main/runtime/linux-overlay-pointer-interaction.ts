/*
  Linux overlay pointer-interaction loop.

  Electron cannot forward mouse-move events through a click-through window on Linux/X11
  (the `forward` option of setIgnoreMouseEvents is unsupported there — electron/electron#16777).
  The overlay's hover/lookup interaction relied on those forwarded events, so under XWayland
  the click-through overlay never sees the cursor and stays inert.

  This restores the Windows/macOS behavior with either a Linux input shape (preferred) or a
  main-process cursor poll fallback. Input shapes keep only reported subtitle/sidebar rects
  mouse-active so entering a subtitle does not have to flip BrowserWindow mouse-ignore state.
  The cursor poll remains for runtimes where BrowserWindow.setShape is unavailable.
*/

import type { OverlayContentMeasurement } from '../../types';

export type PointerPoint = { x: number; y: number };
export type PointerRect = { x: number; y: number; width: number; height: number };
export type PointerViewport = { width: number; height: number };

export type OverlayContentMeasurementLike = {
  viewport: PointerViewport;
  contentRect: PointerRect | null;
  interactiveRects?: PointerRect[] | null;
} | null;

type PointerInteractionWindow = {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  getBounds: () => PointerRect;
};

type PointerInteractionMousePassthroughWindow = {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
};

type PointerInteractionShapeWindow = PointerInteractionMousePassthroughWindow & {
  getBounds: () => PointerRect;
  setShape?: (rects: PointerRect[]) => void;
};

export type LinuxOverlayPointerInteractionDeps = {
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => PointerInteractionWindow | null;
  getCursorScreenPoint: () => PointerPoint;
  getSubtitleMeasurement: () => OverlayContentMeasurementLike;
  getRendererInteractiveHint: () => boolean;
  /** True when a modal/stats overlay owns input — leave interaction state to that logic. */
  shouldSuspend: () => boolean;
  /** True when a separate app window should stay above the overlay. */
  shouldSuppressInteraction?: () => boolean;
  shouldUseInputShape?: () => boolean;
  getInteractionActive: () => boolean;
  isInteractionStateApplied?: () => boolean;
  setInteractionActive: (active: boolean) => void;
};

export const LINUX_OVERLAY_POINTER_POLL_INTERVAL_MS = 60;
// Padding (in window px) so the cursor doesn't have to land pixel-perfectly on the text.
const SUBTITLE_HIT_PADDING_PX = 6;

let pointerInteractionInterval: ReturnType<typeof setInterval> | null = null;

export function mapOverlayMeasurementForPointerInteraction(
  measurement: OverlayContentMeasurement | null,
): OverlayContentMeasurementLike {
  if (!measurement) return null;
  return {
    viewport: measurement.viewport,
    contentRect: measurement.contentRect,
    ...(measurement.interactiveRects ? { interactiveRects: measurement.interactiveRects } : {}),
  };
}

export function shouldSuppressPointerInteractionForForegroundWindow(options: {
  hasForegroundSeparateWindow: boolean;
  isTrackingMpvWindow: boolean;
  isMpvWindowFocused: boolean;
  isOverlayWindowFocused: boolean;
}): boolean {
  if (options.hasForegroundSeparateWindow) return true;
  if (!options.isTrackingMpvWindow) return false;
  return !options.isMpvWindowFocused && !options.isOverlayWindowFocused;
}

/** Mutable timer state for {@link resolveForegroundSuppressionWithGrace}. */
export type ForegroundSuppressionGraceState = { lossSinceMs: number | null };

/**
 * Suppress subtitle pointer interaction for a foreground window, but only once the foreground
 * loss has been *stable* for `graceMs`. A separate SubMiner window defers immediately; a plain
 * focus blip (e.g. the overlay briefly becoming the X11 active window at playback start) is
 * ignored so subtitles don't go inert for a poll cycle while focus settles back onto mpv.
 */
export function resolveForegroundSuppressionWithGrace(options: {
  hasForegroundSeparateWindow: boolean;
  isTrackingMpvWindow: boolean;
  isMpvWindowFocused: boolean;
  isOverlayWindowFocused: boolean;
  nowMs: number;
  graceMs: number;
  state: ForegroundSuppressionGraceState;
}): boolean {
  if (options.hasForegroundSeparateWindow) {
    options.state.lossSinceMs = null;
    return true;
  }

  const rawSuppress = shouldSuppressPointerInteractionForForegroundWindow(options);
  if (!rawSuppress) {
    options.state.lossSinceMs = null;
    return false;
  }

  if (options.state.lossSinceMs === null) {
    options.state.lossSinceMs = options.nowMs;
  }
  return options.nowMs - options.state.lossSinceMs >= options.graceMs;
}

function isCursorOverRect(
  cursor: PointerPoint,
  bounds: PointerRect,
  viewport: PointerViewport,
  rect: PointerRect,
): boolean {
  if (!(bounds.width > 0) || !(bounds.height > 0)) return false;

  const scaleX = bounds.width / viewport.width;
  const scaleY = bounds.height / viewport.height;
  const left = bounds.x + rect.x * scaleX - SUBTITLE_HIT_PADDING_PX;
  const top = bounds.y + rect.y * scaleY - SUBTITLE_HIT_PADDING_PX;
  const right = left + rect.width * scaleX + SUBTITLE_HIT_PADDING_PX * 2;
  const bottom = top + rect.height * scaleY + SUBTITLE_HIT_PADDING_PX * 2;

  return cursor.x >= left && cursor.x <= right && cursor.y >= top && cursor.y <= bottom;
}

function measuredRectsForInput(measurement: OverlayContentMeasurementLike): PointerRect[] {
  if (!measurement) return [];
  return Array.isArray(measurement.interactiveRects) && measurement.interactiveRects.length > 0
    ? measurement.interactiveRects
    : measurement.contentRect
      ? [measurement.contentRect]
      : [];
}

function hasMeasuredInputRects(measurement: OverlayContentMeasurementLike): boolean {
  return measuredRectsForInput(measurement).some((rect) => rect.width > 0 && rect.height > 0);
}

export function shouldPrimeLinuxOverlayInteractionFromMeasurement(deps: {
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => PointerInteractionWindow | null;
  getSubtitleMeasurement: () => OverlayContentMeasurementLike;
  shouldSuspend: () => boolean;
  shouldSuppressInteraction?: () => boolean;
}): boolean {
  if (!deps.getVisibleOverlayVisible()) return false;
  if (deps.shouldSuspend()) return false;

  const mainWindow = deps.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible()) {
    return false;
  }

  if (deps.shouldSuppressInteraction?.()) return false;
  return hasMeasuredInputRects(deps.getSubtitleMeasurement());
}

function clampRectToWindow(rect: PointerRect, bounds: PointerRect): PointerRect | null {
  const left = Math.max(0, Math.floor(rect.x));
  const top = Math.max(0, Math.floor(rect.y));
  const right = Math.min(Math.ceil(bounds.width), Math.ceil(rect.x + rect.width));
  const bottom = Math.min(Math.ceil(bounds.height), Math.ceil(rect.y + rect.height));
  if (right <= left || bottom <= top) return null;
  return {
    x: left,
    y: top,
    width: right - left,
    height: bottom - top,
  };
}

function mapMeasuredRectToWindowShape(
  bounds: PointerRect,
  viewport: PointerViewport,
  rect: PointerRect,
): PointerRect | null {
  if (!(bounds.width > 0) || !(bounds.height > 0)) return null;
  if (!(viewport.width > 0) || !(viewport.height > 0)) return null;

  const scaleX = bounds.width / viewport.width;
  const scaleY = bounds.height / viewport.height;
  return clampRectToWindow(
    {
      x: rect.x * scaleX - SUBTITLE_HIT_PADDING_PX,
      y: rect.y * scaleY - SUBTITLE_HIT_PADDING_PX,
      width: rect.width * scaleX + SUBTITLE_HIT_PADDING_PX * 2,
      height: rect.height * scaleY + SUBTITLE_HIT_PADDING_PX * 2,
    },
    bounds,
  );
}

function resolveInputShapeRects(options: {
  bounds: PointerRect;
  measurement: OverlayContentMeasurementLike;
  rendererInteractiveHint: boolean;
}): PointerRect[] {
  const { bounds } = options;
  if (!(bounds.width > 0) || !(bounds.height > 0)) return [];

  if (options.rendererInteractiveHint) {
    return [
      {
        x: 0,
        y: 0,
        width: Math.ceil(bounds.width),
        height: Math.ceil(bounds.height),
      },
    ];
  }

  const measurement = options.measurement;
  if (!measurement) return [];
  return measuredRectsForInput(measurement)
    .map((rect) => mapMeasuredRectToWindowShape(bounds, measurement.viewport, rect))
    .filter((rect): rect is PointerRect => rect !== null);
}

/** Hit-test the global cursor against subtitle bar rects, mapping viewport px → screen px. */
export function isCursorOverSubtitle(
  cursor: PointerPoint,
  bounds: PointerRect,
  measurement: OverlayContentMeasurementLike,
): boolean {
  if (!measurement) return false;
  const { viewport } = measurement;
  if (!(viewport.width > 0) || !(viewport.height > 0)) return false;

  const rects = measuredRectsForInput(measurement);

  return rects.some((rect) => isCursorOverRect(cursor, bounds, viewport, rect));
}

/**
 * Returns the desired interactive state, or null when the loop should not touch it
 * (overlay hidden/destroyed or another surface owns input).
 */
export function resolveDesiredOverlayInteractive(
  deps: LinuxOverlayPointerInteractionDeps,
): boolean | null {
  if (!deps.getVisibleOverlayVisible()) return false;
  if (deps.shouldSuspend()) return null;

  const mainWindow = deps.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible()) {
    return null;
  }

  if (deps.shouldSuppressInteraction?.()) return false;
  if (deps.getRendererInteractiveHint()) return true;
  return isCursorOverSubtitle(
    deps.getCursorScreenPoint(),
    mainWindow.getBounds(),
    deps.getSubtitleMeasurement(),
  );
}

export function tickLinuxOverlayPointerInteraction(deps: LinuxOverlayPointerInteractionDeps): void {
  if (deps.shouldUseInputShape?.()) return;
  const desired = resolveDesiredOverlayInteractive(deps);
  if (desired === null) return;
  if (deps.getInteractionActive() === desired && deps.isInteractionStateApplied?.() !== false) {
    return;
  }
  deps.setInteractionActive(desired);
}

export function applyLinuxOverlayInputShape(deps: {
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => PointerInteractionShapeWindow | null;
  getSubtitleMeasurement: () => OverlayContentMeasurementLike;
  getRendererInteractiveHint: () => boolean;
  shouldSuspend: () => boolean;
  shouldSuppressInteraction?: () => boolean;
}): { handled: boolean; active: boolean } {
  const mainWindow = deps.getMainWindow();
  if (!mainWindow || typeof mainWindow.setShape !== 'function') {
    return { handled: false, active: false };
  }

  if (
    !deps.getVisibleOverlayVisible() ||
    deps.shouldSuspend() ||
    mainWindow.isDestroyed() ||
    deps.shouldSuppressInteraction?.()
  ) {
    mainWindow.setIgnoreMouseEvents(true, { forward: true });
    mainWindow.setShape([]);
    return { handled: true, active: false };
  }

  const rects = resolveInputShapeRects({
    bounds: mainWindow.getBounds(),
    measurement: deps.getSubtitleMeasurement(),
    rendererInteractiveHint: deps.getRendererInteractiveHint(),
  });

  if (rects.length === 0) {
    mainWindow.setIgnoreMouseEvents(true, { forward: true });
    mainWindow.setShape([]);
    return { handled: true, active: false };
  }

  mainWindow.setShape(rects);
  mainWindow.setIgnoreMouseEvents(false);
  return { handled: true, active: true };
}

export function applyLinuxOverlayPointerInteractionMousePassthrough(deps: {
  active: boolean;
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => PointerInteractionMousePassthroughWindow | null;
  shouldSuspend: () => boolean;
  shouldSuppressInteraction?: () => boolean;
  updateVisibleOverlayVisibility: () => void;
}): boolean {
  if (!deps.getVisibleOverlayVisible() || deps.shouldSuspend()) {
    deps.updateVisibleOverlayVisibility();
    return false;
  }

  const mainWindow = deps.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible()) {
    deps.updateVisibleOverlayVisibility();
    return false;
  }

  if (deps.shouldSuppressInteraction?.()) {
    deps.updateVisibleOverlayVisibility();
    return false;
  }

  if (deps.active) {
    mainWindow.setIgnoreMouseEvents(false);
  } else {
    mainWindow.setIgnoreMouseEvents(true, { forward: true });
  }
  return true;
}

export function ensureLinuxOverlayPointerInteractionLoop(
  deps: LinuxOverlayPointerInteractionDeps,
  platform: NodeJS.Platform = process.platform,
): void {
  if (pointerInteractionInterval !== null) return;
  if (platform !== 'linux') return;

  pointerInteractionInterval = setInterval(() => {
    tickLinuxOverlayPointerInteraction(deps);
  }, LINUX_OVERLAY_POINTER_POLL_INTERVAL_MS);
  pointerInteractionInterval.unref?.();
}

export function stopLinuxOverlayPointerInteractionLoop(): void {
  if (pointerInteractionInterval === null) return;
  clearInterval(pointerInteractionInterval);
  pointerInteractionInterval = null;
}
