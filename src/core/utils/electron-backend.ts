import { CliArgs, shouldStartApp } from "../../cli/args";
import { createLogger } from "../../logger";

const logger = createLogger("core:electron-backend");

function getElectronOzonePlatformHint(): string | null {
  const hint = process.env.ELECTRON_OZONE_PLATFORM_HINT?.trim().toLowerCase();
  if (hint) return hint;
  const ozone = process.env.OZONE_PLATFORM?.trim().toLowerCase();
  if (ozone) return ozone;
  return null;
}

function shouldPreferWaylandBackend(): boolean {
  return Boolean(
    process.env.HYPRLAND_INSTANCE_SIGNATURE || process.env.SWAYSOCK,
  );
}

export function forceX11Backend(args: CliArgs): void {
  if (process.platform !== "linux") return;
  if (!shouldStartApp(args)) return;
  if (shouldPreferWaylandBackend()) return;

  const hint = getElectronOzonePlatformHint();
  if (hint === "x11") return;

  process.env.ELECTRON_OZONE_PLATFORM_HINT = "x11";
  process.env.OZONE_PLATFORM = "x11";
}

export function enforceUnsupportedWaylandMode(args: CliArgs): void {
  if (process.platform !== "linux") return;
  if (!shouldStartApp(args)) return;
  const hint = getElectronOzonePlatformHint();
  if (hint !== "wayland") return;

  const message =
    "Unsupported Electron backend: Wayland. Set ELECTRON_OZONE_PLATFORM_HINT=x11 and restart SubMiner.";
  logger.error(message);
  throw new Error(message);
}
