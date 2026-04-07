import { parseMpvLaunchMode } from '../../src/shared/mpv-launch-mode.js';
import type { LauncherMpvConfig } from '../types.js';

export function parseLauncherMpvConfig(root: Record<string, unknown>): LauncherMpvConfig {
  const mpvRaw = root.mpv;
  if (!mpvRaw || typeof mpvRaw !== 'object') return {};
  const mpv = mpvRaw as Record<string, unknown>;

  return {
    launchMode: parseMpvLaunchMode(mpv.launchMode),
  };
}
