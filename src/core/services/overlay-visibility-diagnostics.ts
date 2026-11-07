import type { BrowserWindow } from 'electron';
import type { BaseWindowTracker } from '../../window-trackers';
import { createLogger } from '../../logger';
import {
  OVERLAY_WINDOW_CONTENT_READY_FLAG,
  OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG,
} from './overlay-window-flags';

const log = createLogger('overlay:visibility');

export type OverlayVisibilitySnapshot = {
  expectedVisible: boolean;
  visible: boolean;
  opacity: number;
  contentReady: boolean;
  tracking: boolean;
  [key: string]: unknown;
};

// One watchdog per window, independent of further visibility events: a renderer
// that never becomes ready must still leave evidence at the default warn level.
export function createOverlayVisibilityDiagnostics(deps: {
  warn: (message: string, details: Record<string, unknown>) => void;
  now: () => number;
  schedule: (callback: () => void, delayMs: number) => ReturnType<typeof setTimeout>;
  cancel: (timer: ReturnType<typeof setTimeout>) => void;
}) {
  let timer: ReturnType<typeof setTimeout> | null = null;
  let readSnapshot: () => OverlayVisibilitySnapshot | null = () => null;
  let stalledAtMs: number | null = null;
  let warned = false;
  let lastWarningAtMs = 0;

  const check = () => {
    if (timer !== null) deps.cancel(timer);
    timer = null;
    const snapshot = readSnapshot();
    if (!snapshot?.expectedVisible) {
      stalledAtMs = null;
      warned = false;
      return;
    }
    if (snapshot.visible && snapshot.opacity > 0 && snapshot.contentReady) {
      if (warned) {
        deps.warn('Windows overlay visibility recovered', {
          ...snapshot,
          elapsedMs: deps.now() - stalledAtMs!,
        });
      }
      stalledAtMs = null;
      warned = false;
      return;
    }
    stalledAtMs ??= deps.now();
    const elapsedMs = deps.now() - stalledAtMs;
    if (elapsedMs >= 5_000 && (!warned || deps.now() - lastWarningAtMs >= 30_000)) {
      const reason = !snapshot.tracking
        ? 'mpv-window-not-tracked'
        : !snapshot.contentReady
          ? 'renderer-not-ready'
          : !snapshot.visible
            ? 'window-hidden'
            : 'window-transparent';
      deps.warn('Windows overlay is not visible after startup or tracking loss', {
        ...snapshot,
        reason,
        elapsedMs,
      });
      warned = true;
      lastWarningAtMs = deps.now();
    }
    timer = deps.schedule(
      check,
      warned ? 30_000 - (deps.now() - lastWarningAtMs) : 5_000 - elapsedMs,
    );
    timer.unref?.();
  };

  return {
    update(reader: () => OverlayVisibilitySnapshot | null) {
      readSnapshot = reader;
      // Frequent polls must not postpone a warning or bypass its rate limit.
      check();
    },
  };
}

const diagnosticsByWindow = new WeakMap<
  BrowserWindow,
  ReturnType<typeof createOverlayVisibilityDiagnostics>
>();

export function watchWindowsOverlayVisibility(args: {
  mainWindow: BrowserWindow;
  windowTracker: BaseWindowTracker | null;
  visibleOverlayVisible: boolean;
  modalActive?: boolean;
  suspendVisibleOverlay?: boolean;
}): void {
  const window = args.mainWindow;
  let diagnostics = diagnosticsByWindow.get(window);
  if (!diagnostics) {
    diagnostics = createOverlayVisibilityDiagnostics({
      warn: (message, details) => log.warn(message, details),
      now: Date.now,
      schedule: setTimeout,
      cancel: clearTimeout,
    });
    diagnosticsByWindow.set(window, diagnostics);
  }
  diagnostics.update(() => {
    if (window.isDestroyed()) return null;
    const tracker = args.windowTracker as
      | (BaseWindowTracker & { getDiagnosticState?: () => Record<string, unknown> })
      | null;
    const flags = window as unknown as Record<string, unknown>;
    const minimized = tracker?.isTargetWindowMinimized?.() ?? false;
    return {
      ...tracker?.getDiagnosticState?.(),
      expectedVisible:
        args.visibleOverlayVisible &&
        !args.modalActive &&
        !args.suspendVisibleOverlay &&
        !minimized,
      visible: window.isVisible(),
      opacity: window.getOpacity?.() ?? 1,
      contentReady: flags[OVERLAY_WINDOW_CONTENT_READY_FLAG] === true,
      documentLoaded: flags[OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG] === true,
      rendererLoading: window.webContents?.isLoading?.() ?? null,
      tracking: tracker?.isTracking() ?? false,
      trackerInitialized: tracker !== null,
      geometry: tracker?.getGeometry() ?? null,
      focused: tracker?.isTargetWindowFocused?.() ?? false,
      minimized,
    };
  });
}
