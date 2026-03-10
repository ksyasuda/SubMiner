import fs from 'node:fs';
import path from 'node:path';

export const WINDOWS_MPV_SHORTCUT_NAME = 'SubMiner mpv.lnk';

export interface WindowsMpvShortcutPaths {
  startMenuPath: string;
  desktopPath: string;
}

export interface WindowsShortcutLinkDetails {
  target: string;
  args?: string;
  cwd?: string;
  description?: string;
  icon?: string;
  iconIndex?: number;
}

export interface WindowsMpvShortcutInstallResult {
  ok: boolean;
  status: 'installed' | 'skipped' | 'failed';
  message: string;
}

export function resolveWindowsStartMenuProgramsDir(appDataDir: string): string {
  return path.join(appDataDir, 'Microsoft', 'Windows', 'Start Menu', 'Programs');
}

export function resolveWindowsMpvShortcutPaths(options: {
  appDataDir: string;
  desktopDir: string;
}): WindowsMpvShortcutPaths {
  return {
    startMenuPath: path.join(
      resolveWindowsStartMenuProgramsDir(options.appDataDir),
      WINDOWS_MPV_SHORTCUT_NAME,
    ),
    desktopPath: path.join(options.desktopDir, WINDOWS_MPV_SHORTCUT_NAME),
  };
}

export function detectWindowsMpvShortcuts(
  paths: WindowsMpvShortcutPaths,
  existsSync: (candidate: string) => boolean = fs.existsSync,
): { startMenuInstalled: boolean; desktopInstalled: boolean } {
  return {
    startMenuInstalled: existsSync(paths.startMenuPath),
    desktopInstalled: existsSync(paths.desktopPath),
  };
}

export function buildWindowsMpvShortcutDetails(exePath: string): WindowsShortcutLinkDetails {
  return {
    target: exePath,
    args: '--launch-mpv',
    cwd: path.dirname(exePath),
    description: 'Launch mpv with the SubMiner profile',
    icon: exePath,
    iconIndex: 0,
  };
}

export function applyWindowsMpvShortcuts(options: {
  preferences: { startMenuEnabled: boolean; desktopEnabled: boolean };
  paths: WindowsMpvShortcutPaths;
  exePath: string;
  writeShortcutLink: (
    shortcutPath: string,
    operation: 'create' | 'update' | 'replace',
    details: WindowsShortcutLinkDetails,
  ) => boolean;
  rmSync?: (candidate: string, options: { force: true }) => void;
  mkdirSync?: (candidate: string, options: { recursive: true }) => void;
}): WindowsMpvShortcutInstallResult {
  const rmSync = options.rmSync ?? fs.rmSync;
  const mkdirSync = options.mkdirSync ?? fs.mkdirSync;
  const details = buildWindowsMpvShortcutDetails(options.exePath);
  const failures: string[] = [];

  const ensureShortcut = (shortcutPath: string): void => {
    mkdirSync(path.dirname(shortcutPath), { recursive: true });
    const ok = options.writeShortcutLink(shortcutPath, 'replace', details);
    if (!ok) {
      failures.push(shortcutPath);
    }
  };

  const removeShortcut = (shortcutPath: string): void => {
    rmSync(shortcutPath, { force: true });
  };

  if (options.preferences.startMenuEnabled) ensureShortcut(options.paths.startMenuPath);
  else removeShortcut(options.paths.startMenuPath);

  if (options.preferences.desktopEnabled) ensureShortcut(options.paths.desktopPath);
  else removeShortcut(options.paths.desktopPath);

  if (failures.length > 0) {
    return {
      ok: false,
      status: 'failed',
      message: `Failed to create Windows mpv shortcuts: ${failures.join(', ')}`,
    };
  }

  if (!options.preferences.startMenuEnabled && !options.preferences.desktopEnabled) {
    return {
      ok: true,
      status: 'skipped',
      message: 'Disabled Windows mpv shortcuts.',
    };
  }

  return {
    ok: true,
    status: 'installed',
    message: 'Updated Windows mpv shortcuts.',
  };
}
