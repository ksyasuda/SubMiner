import type { BrowserWindow } from 'electron';
import {
  ensureHyprlandWindowFloatingByTitle,
  shouldAttemptHyprlandWindowPlacement,
} from './hyprland-window-placement';

type SettingsWindowLevelController = Pick<
  BrowserWindow,
  'getTitle' | 'isDestroyed' | 'moveTop' | 'setAlwaysOnTop'
>;

type PromoteSettingsWindowOptions = {
  platform?: NodeJS.Platform;
  env?: NodeJS.ProcessEnv;
  ensureHyprlandWindowFloatingByTitle?: typeof ensureHyprlandWindowFloatingByTitle;
};

export function shouldPromoteSettingsWindowAboveOverlay(
  platform: NodeJS.Platform = process.platform,
  env: NodeJS.ProcessEnv = process.env,
): boolean {
  return shouldAttemptHyprlandWindowPlacement(platform, env);
}

export function promoteSettingsWindowAboveOverlay(
  window: SettingsWindowLevelController,
  options: PromoteSettingsWindowOptions = {},
): boolean {
  const platform = options.platform ?? process.platform;
  const env = options.env ?? process.env;
  if (window.isDestroyed() || !shouldPromoteSettingsWindowAboveOverlay(platform, env)) {
    return false;
  }

  window.setAlwaysOnTop(true);
  window.moveTop();

  const title = window.getTitle().trim();
  if (title) {
    (options.ensureHyprlandWindowFloatingByTitle ?? ensureHyprlandWindowFloatingByTitle)({
      title,
      platform,
      env,
    });
  }

  return true;
}
