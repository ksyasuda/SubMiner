import type { MpvLaunchMode } from '../types/config';

export const MPV_LAUNCH_MODE_VALUES = ['normal', 'maximized', 'fullscreen'] as const satisfies
  readonly MpvLaunchMode[];

export function parseMpvLaunchMode(value: unknown): MpvLaunchMode | undefined {
  if (typeof value !== 'string') {
    return undefined;
  }

  const normalized = value.trim().toLowerCase();
  return MPV_LAUNCH_MODE_VALUES.find((candidate) => candidate === normalized);
}

export function resolveRawMpvLaunchMode(
  mpv?: {
    launchMode?: unknown;
    startFullscreen?: unknown;
  } | null,
): MpvLaunchMode | undefined {
  if (!mpv) {
    return undefined;
  }

  if (mpv.launchMode !== undefined) {
    return parseMpvLaunchMode(mpv.launchMode);
  }

  if (typeof mpv.startFullscreen === 'boolean') {
    return mpv.startFullscreen ? 'fullscreen' : 'normal';
  }

  return undefined;
}

export function buildMpvLaunchModeArgs(launchMode: MpvLaunchMode): string[] {
  switch (launchMode) {
    case 'fullscreen':
      return ['--fullscreen'];
    case 'maximized':
      return ['--window-maximized=yes'];
    default:
      return [];
  }
}
