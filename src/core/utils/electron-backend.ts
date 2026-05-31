import { CliArgs, shouldStartApp } from '../../cli/args';
import { createLogger } from '../../logger';
import { isSupportedWaylandCompositor } from '../../shared/mpv-x11-backend';

const logger = createLogger('core:electron-backend');

function getElectronOzonePlatformHint(env: NodeJS.ProcessEnv = process.env): string | null {
  const hint = env.ELECTRON_OZONE_PLATFORM_HINT?.trim().toLowerCase();
  if (hint) return hint;
  const ozone = env.OZONE_PLATFORM?.trim().toLowerCase();
  if (ozone) return ozone;
  return null;
}

/**
 * Should the Electron app be pinned to the X11/XWayland ozone backend? True on Linux
 * unless we're on a natively-supported Wayland compositor (Hyprland/Sway) or the user
 * explicitly opted into the (unsupported) Wayland backend — which is reported by
 * {@link enforceUnsupportedWaylandMode} instead.
 *
 * The overlay relies on `setAlwaysOnTop`/`moveTop` to stay above mpv; those are no-ops
 * under a native Wayland surface, so XWayland is required for parity with Win/macOS. An
 * explicit `ELECTRON_OZONE_PLATFORM_HINT=wayland` is still overridden to x11 here (the
 * Electron Wayland backend is unsupported); the Hyprland/Sway case is left untouched so
 * {@link enforceUnsupportedWaylandMode} can report it.
 */
export function shouldForceX11ElectronBackend(env: NodeJS.ProcessEnv = process.env): boolean {
  if (process.platform !== 'linux') return false;
  return !isSupportedWaylandCompositor(env);
}

export function forceX11Backend(args: CliArgs): void {
  if (!shouldStartApp(args)) return;
  if (!shouldForceX11ElectronBackend()) return;
  if (getElectronOzonePlatformHint() === 'x11') return;

  process.env.ELECTRON_OZONE_PLATFORM_HINT = 'x11';
  process.env.OZONE_PLATFORM = 'x11';
}

export function enforceUnsupportedWaylandMode(args: CliArgs): void {
  if (process.platform !== 'linux') return;
  if (!shouldStartApp(args)) return;
  const hint = getElectronOzonePlatformHint();
  if (hint !== 'wayland') return;

  const message =
    'Unsupported Electron backend: Wayland. Set ELECTRON_OZONE_PLATFORM_HINT=x11 and restart SubMiner.';
  logger.error(message);
  throw new Error(message);
}
