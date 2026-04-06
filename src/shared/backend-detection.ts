export type SessionBackend = 'hyprland' | 'kwin' | 'sway' | 'x11' | 'macos' | 'windows' | null;

export function detectSessionBackend(
  platform: NodeJS.Platform = process.platform,
  env: NodeJS.ProcessEnv = process.env,
): SessionBackend {
  if (platform === 'win32') return 'windows';
  if (platform === 'darwin') return 'macos';
  if (env.HYPRLAND_INSTANCE_SIGNATURE) return 'hyprland';
  if (env.SWAYSOCK) return 'sway';

  const xdgCurrentDesktop = (env.XDG_CURRENT_DESKTOP || '').toLowerCase();
  const xdgSessionDesktop = (env.XDG_SESSION_DESKTOP || '').toLowerCase();
  const xdgSessionType = (env.XDG_SESSION_TYPE || '').toLowerCase();
  const hasWayland = Boolean(env.WAYLAND_DISPLAY) || xdgSessionType === 'wayland';
  const isKWinDesktop =
    xdgCurrentDesktop.includes('kde') ||
    xdgCurrentDesktop.includes('plasma') ||
    xdgSessionDesktop.includes('kde') ||
    xdgSessionDesktop.includes('plasma');

  if (hasWayland && isKWinDesktop) return 'kwin';
  if (env.DISPLAY || platform === 'linux') return 'x11';
  return null;
}
