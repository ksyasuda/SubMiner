import { CliArgs, shouldStartApp } from '../../cli/args';
import { createLogger } from '../../logger';
import { isSupportedWaylandCompositor } from '../../shared/mpv-x11-backend';

const logger = createLogger('core:electron-backend');

export const X11_ELECTRON_BOOTSTRAP_ENV = 'SUBMINER_X11_BOOTSTRAPPED';
const X11_ELECTRON_OZONE_ARG = '--ozone-platform=x11';

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
export function shouldForceX11ElectronBackend(
  env: NodeJS.ProcessEnv = process.env,
  platform: NodeJS.Platform = process.platform,
): boolean {
  if (platform !== 'linux') return false;
  return !isSupportedWaylandCompositor(env);
}

export function resolveX11ElectronRelaunchArgs(
  args: string[],
  env: NodeJS.ProcessEnv = process.env,
  platform: NodeJS.Platform = process.platform,
): string[] | null {
  if (!shouldForceX11ElectronBackend(env, platform)) return null;
  if (env[X11_ELECTRON_BOOTSTRAP_ENV] === '1') return null;

  const retainedArgs: string[] = [];
  let alreadyForced = false;
  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index];
    if (arg === '--ozone-platform') {
      const value = args[index + 1];
      alreadyForced = value?.trim().toLowerCase() === 'x11';
      if (value && !value.startsWith('--')) index += 1;
      continue;
    }
    if (arg?.startsWith('--ozone-platform=')) {
      alreadyForced = arg.slice('--ozone-platform='.length).trim().toLowerCase() === 'x11';
      continue;
    }
    if (arg) retainedArgs.push(arg);
  }

  return alreadyForced ? null : [...retainedArgs, X11_ELECTRON_OZONE_ARG];
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
