import fs from 'node:fs';
import { log } from '../log.js';
import { runAppCommandWithInherit } from '../mpv.js';
import { commandExists } from '../util.js';
import { resolveMainConfigPath } from '../config-path.js';
import type { LauncherCommandContext } from './context.js';

interface DoctorCommandDeps {
  commandExists(command: string): boolean;
  configExists(path: string): boolean;
  resolveMainConfigPath(): string;
  runAppCommandWithInherit(appPath: string, appArgs: string[]): void;
}

const defaultDeps: DoctorCommandDeps = {
  commandExists,
  configExists: fs.existsSync,
  resolveMainConfigPath,
  runAppCommandWithInherit,
};

export function runDoctorCommand(
  context: LauncherCommandContext,
  deps: DoctorCommandDeps = defaultDeps,
): boolean {
  const { args, appPath, mpvSocketPath, processAdapter } = context;
  if (!args.doctor) {
    return false;
  }

  const configPath = deps.resolveMainConfigPath();
  const mpvFound = deps.commandExists('mpv');
  const checks: Array<{ label: string; ok: boolean; detail: string }> = [
    {
      label: 'app binary',
      ok: Boolean(appPath),
      detail: appPath || 'not found (set SUBMINER_APPIMAGE_PATH)',
    },
    {
      label: 'mpv',
      ok: mpvFound,
      detail: mpvFound ? 'found' : 'missing',
    },
    {
      label: 'yt-dlp',
      ok: deps.commandExists('yt-dlp'),
      detail: deps.commandExists('yt-dlp') ? 'found' : 'missing (optional unless YouTube URLs)',
    },
    {
      label: 'ffmpeg',
      ok: deps.commandExists('ffmpeg'),
      detail: deps.commandExists('ffmpeg')
        ? 'found'
        : 'missing (optional unless legacy subtitle fallback is enabled)',
    },
    {
      label: 'fzf',
      ok: deps.commandExists('fzf'),
      detail: deps.commandExists('fzf') ? 'found' : 'missing (optional if using rofi)',
    },
    {
      label: 'rofi',
      ok: deps.commandExists('rofi'),
      detail: deps.commandExists('rofi') ? 'found' : 'missing (optional if using fzf)',
    },
    {
      label: 'config',
      ok: deps.configExists(configPath),
      detail: configPath,
    },
    {
      label: 'mpv socket path',
      ok: true,
      detail: mpvSocketPath,
    },
  ];

  for (const check of checks) {
    log(check.ok ? 'info' : 'warn', args.logLevel, `[doctor] ${check.label}: ${check.detail}`);
  }

  if (args.doctorRefreshKnownWords) {
    if (!appPath) {
      processAdapter.exit(1);
      return true;
    }
    deps.runAppCommandWithInherit(appPath, ['--refresh-known-words']);
    return true;
  }

  const hasHardFailure = checks.some((entry) =>
    entry.label === 'app binary' || entry.label === 'mpv' ? !entry.ok : false,
  );
  processAdapter.exit(hasHardFailure ? 1 : 0);
  return true;
}
