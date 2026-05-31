import { isSupportedWaylandCompositor } from '../../shared/mpv-x11-backend';

/*
  Linux overlay z-order keep-alive loop.

  The visible overlay re-asserts its always-on-top level only when mpv's geometry changes
  (the bounds-update path) or on a fullscreen toggle (the fullscreen refresh burst). When mpv
  is raised above the overlay WITHOUT a geometry change — click-to-raise, focus change, or a
  compositor restack on KDE/GNOME/other X11/XWayland window managers — nothing re-raises the
  overlay and it stays buried. Windows guards against this with a foreground poll loop; this is
  the Linux equivalent: a lightweight periodic re-assert while the overlay is shown and mpv
  remains the foreground window. If another app is active, the overlay releases its global
  keep-above level so that app can cover it.

  Gated to X11/XWayland sessions (not Hyprland/Sway, which place the overlay natively and would
  otherwise be spammed with hyprctl dispatches).
*/

type KeepAliveOverlayWindow = {
  isDestroyed: () => boolean;
  isVisible: () => boolean;
  focus?: () => void;
};

export type LinuxOverlayZOrderKeepAliveDeps = {
  getVisibleOverlayVisible: () => boolean;
  getMainWindow: () => KeepAliveOverlayWindow | null;
  isTrackingMpvWindow: () => boolean;
  isMpvWindowFocused: () => boolean;
  isOverlayWindowFocused: () => boolean;
  /** True when a modal/stats overlay or active interaction owns the top — skip re-asserting. */
  shouldSuppressReassert: () => boolean;
  raiseMpvWindow: () => Promise<boolean>;
  releaseOverlayLayerOrder: () => void;
  enforceOverlayLayerOrder: () => void;
  focusOverlayWindow?: () => void;
};

export const LINUX_OVERLAY_ZORDER_KEEPALIVE_INTERVAL_MS = 700;

let keepAliveInterval: ReturnType<typeof setInterval> | null = null;
let keepAliveTickInFlight = false;

export function shouldRunLinuxOverlayZOrderKeepAlive(
  env: NodeJS.ProcessEnv = process.env,
): boolean {
  return process.platform === 'linux' && !isSupportedWaylandCompositor(env);
}

export async function tickLinuxOverlayZOrderKeepAlive(
  deps: LinuxOverlayZOrderKeepAliveDeps,
): Promise<void> {
  if (!deps.getVisibleOverlayVisible()) return;
  if (!deps.isTrackingMpvWindow()) return;

  const mainWindow = deps.getMainWindow();
  if (!mainWindow || mainWindow.isDestroyed() || !mainWindow.isVisible()) {
    return;
  }

  const overlayFocused = deps.isOverlayWindowFocused();
  const mpvFocused = deps.isMpvWindowFocused();
  if (!mpvFocused && !overlayFocused) {
    deps.releaseOverlayLayerOrder();
    return;
  }
  if (deps.shouldSuppressReassert()) return;

  if (overlayFocused && !mpvFocused) {
    await deps.raiseMpvWindow();
  }
  deps.enforceOverlayLayerOrder();
  if (overlayFocused && !mpvFocused) {
    deps.focusOverlayWindow?.();
  }
}

export function ensureLinuxOverlayZOrderKeepAliveLoop(
  deps: LinuxOverlayZOrderKeepAliveDeps,
  env: NodeJS.ProcessEnv = process.env,
): void {
  if (keepAliveInterval !== null) return;
  if (!shouldRunLinuxOverlayZOrderKeepAlive(env)) return;

  keepAliveInterval = setInterval(() => {
    if (keepAliveTickInFlight) return;
    keepAliveTickInFlight = true;
    void tickLinuxOverlayZOrderKeepAlive(deps)
      .catch(() => {})
      .finally(() => {
        keepAliveTickInFlight = false;
      });
  }, LINUX_OVERLAY_ZORDER_KEEPALIVE_INTERVAL_MS);
  keepAliveInterval.unref?.();
}

export function stopLinuxOverlayZOrderKeepAliveLoop(): void {
  if (keepAliveInterval === null) return;
  clearInterval(keepAliveInterval);
  keepAliveInterval = null;
  keepAliveTickInFlight = false;
}
