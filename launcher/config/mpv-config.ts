import { parseMpvLaunchMode } from '../../src/shared/mpv-launch-mode.js';
import type { Backend } from '../types.js';
import type { LauncherMpvConfig } from '../types.js';

function parseBackend(value: unknown): Backend | undefined {
  if (typeof value !== 'string') return undefined;
  const normalized = value.trim().toLowerCase();
  if (
    normalized === 'auto' ||
    normalized === 'hyprland' ||
    normalized === 'sway' ||
    normalized === 'x11' ||
    normalized === 'macos' ||
    normalized === 'windows'
  ) {
    return normalized;
  }
  return undefined;
}

function parseNonEmptyString(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : undefined;
}

export function parseLauncherMpvConfig(root: Record<string, unknown>): LauncherMpvConfig {
  const mpvRaw = root.mpv;
  if (!mpvRaw || typeof mpvRaw !== 'object') return {};
  const mpv = mpvRaw as Record<string, unknown>;

  return {
    launchMode: parseMpvLaunchMode(mpv.launchMode),
    socketPath: parseNonEmptyString(mpv.socketPath),
    backend: parseBackend(mpv.backend),
    autoStartSubMiner:
      typeof mpv.autoStartSubMiner === 'boolean' ? mpv.autoStartSubMiner : undefined,
    pauseUntilOverlayReady:
      typeof mpv.pauseUntilOverlayReady === 'boolean' ? mpv.pauseUntilOverlayReady : undefined,
    subminerBinaryPath: parseNonEmptyString(mpv.subminerBinaryPath),
    aniskipEnabled: typeof mpv.aniskipEnabled === 'boolean' ? mpv.aniskipEnabled : undefined,
    aniskipButtonKey: parseNonEmptyString(mpv.aniskipButtonKey),
  };
}
