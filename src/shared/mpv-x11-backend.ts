/*
  Shared XWayland/X11 backend forcing for mpv and the Electron app.

  On Wayland sessions the SubMiner overlay can only be reliably kept above mpv when
  BOTH processes run under XWayland: the Wayland protocol forbids clients from
  controlling window stacking, so Electron's `setAlwaysOnTop`/`moveTop` become
  no-ops under a native Wayland surface. Hyprland and Sway are the exception — they
  are supported natively via compositor-specific window placement — so all forcing
  here is gated to "Linux + Wayland session + NOT Hyprland/Sway".

  This module is shared between the `launcher/` bundle and the Electron `src/` build
  so the gate and the mpv backend args stay in one place.
*/

/**
 * mpv args that pin the *windowing* stack to X11/XWayland, in Vulkan-then-OpenGL order.
 * mpv walks the list and skips contexts that do not match the configured `--gpu-api`,
 * so this works for both a Vulkan and an OpenGL config.
 *
 * Deliberately does NOT set `--vo`/`--gpu-api`: forcing `--vo=gpu --gpu-api=opengl` here
 * used to drop configs off `vo=gpu-next` onto the legacy renderer, where user shaders
 * written for gpu-next (e.g. ArtCNN, `//!COMPONENTS 4` LUMA hooks) abort mpv with
 * `copy_image: Assertion '*offset + count < sizeof(dst)' failed` as soon as their
 * upscale-only `//!WHEN` condition turns on, i.e. on the first fullscreen toggle.
 */
export const MPV_X11_BACKEND_ARGS = ['--gpu-context=x11vk,x11egl,x11'] as const;

export type LinuxDesktopEnv = {
  xdgCurrentDesktop: string;
  xdgSessionDesktop: string;
  hasWayland: boolean;
};

export function getLinuxDesktopEnv(env: NodeJS.ProcessEnv = process.env): LinuxDesktopEnv {
  const xdgCurrentDesktop = (env.XDG_CURRENT_DESKTOP || '').toLowerCase();
  const xdgSessionDesktop = (env.XDG_SESSION_DESKTOP || '').toLowerCase();
  const xdgSessionType = (env.XDG_SESSION_TYPE || '').toLowerCase();
  return {
    xdgCurrentDesktop,
    xdgSessionDesktop,
    hasWayland: Boolean(env.WAYLAND_DISPLAY) || xdgSessionType === 'wayland',
  };
}

/**
 * Compositors that SubMiner supports natively on Wayland (no XWayland forcing).
 * Detected via their socket env vars or the XDG desktop identifiers.
 */
export function isSupportedWaylandCompositor(env: NodeJS.ProcessEnv = process.env): boolean {
  const desktop = getLinuxDesktopEnv(env);
  return (
    Boolean(env.HYPRLAND_INSTANCE_SIGNATURE || env.SWAYSOCK) ||
    desktop.xdgCurrentDesktop.includes('hyprland') ||
    desktop.xdgCurrentDesktop.includes('sway') ||
    desktop.xdgSessionDesktop.includes('hyprland') ||
    desktop.xdgSessionDesktop.includes('sway')
  );
}

/**
 * Should this Linux session be pushed onto XWayland/X11? True for a Wayland session
 * that is not one of the natively-supported compositors and has an X11 display
 * available for the fallback. This is the "auto" decision shared by the Electron app
 * and SubMiner-managed mpv launches.
 */
export function shouldForceX11WaylandSession(env: NodeJS.ProcessEnv = process.env): boolean {
  if (process.platform !== 'linux') return false;
  if (!env.DISPLAY?.trim()) return false;
  if (!getLinuxDesktopEnv(env).hasWayland) return false;
  return !isSupportedWaylandCompositor(env);
}

/**
 * Launcher-facing decision that also honors an explicit `--backend` choice:
 * - `x11` forces the X11 stack whenever an X11 display exists,
 * - `auto` defers to {@link shouldForceX11WaylandSession}.
 */
export function shouldForceX11MpvBackend(
  backend: string,
  env: NodeJS.ProcessEnv = process.env,
): boolean {
  if (process.platform !== 'linux' || !env.DISPLAY?.trim()) {
    return false;
  }
  if (backend === 'x11') return true;
  return backend === 'auto' && shouldForceX11WaylandSession(env);
}

/**
 * Strip Wayland/compositor hints and pin the session type to X11 on the given env
 * object (mutates in place and returns it) so a child mpv process picks XWayland.
 */
export function applyX11EnvOverrides(env: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  delete env.WAYLAND_DISPLAY;
  delete env.HYPRLAND_INSTANCE_SIGNATURE;
  delete env.SWAYSOCK;
  env.XDG_SESSION_TYPE = 'x11';
  return env;
}
