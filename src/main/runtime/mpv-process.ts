import fs from 'node:fs';
import path from 'node:path';
import { spawn, spawnSync } from 'node:child_process';
import {
  MPV_X11_BACKEND_ARGS,
  applyX11EnvOverrides,
  shouldForceX11WaylandSession,
} from '../../shared/mpv-x11-backend';

export interface WindowsMpvPathDeps {
  getEnv: (name: string) => string | undefined;
  runWhere: () => { status: number | null; stdout: string; error?: Error };
  fileExists: (candidate: string) => boolean;
}

export type ConfiguredWindowsMpvPathStatus = 'blank' | 'configured' | 'invalid';

function fileExists(candidate: string): boolean {
  try {
    return fs.statSync(candidate).isFile();
  } catch {
    return false;
  }
}

export function getConfiguredWindowsMpvPathStatus(
  configuredMpvPath = '',
  exists: (candidate: string) => boolean = fileExists,
): ConfiguredWindowsMpvPathStatus {
  const configPath = configuredMpvPath.trim();
  if (!configPath) {
    return 'blank';
  }
  return exists(configPath) ? 'configured' : 'invalid';
}

export function createWindowsMpvPathDeps(
  overrides: Partial<WindowsMpvPathDeps> = {},
): WindowsMpvPathDeps {
  return {
    getEnv: overrides.getEnv ?? ((name) => process.env[name]),
    fileExists: overrides.fileExists ?? fileExists,
    runWhere:
      overrides.runWhere ??
      (() => {
        const result = spawnSync('where.exe', ['mpv.exe'], {
          encoding: 'utf8',
          windowsHide: true,
        });
        return {
          status: result.status,
          stdout: result.stdout ?? '',
          error: result.error ?? undefined,
        };
      }),
  };
}

export function resolveWindowsMpvPath(deps: WindowsMpvPathDeps, configuredMpvPath = ''): string {
  const configPath = configuredMpvPath.trim();
  const configuredPathStatus = getConfiguredWindowsMpvPathStatus(configPath, deps.fileExists);
  if (configuredPathStatus === 'configured') {
    return configPath;
  }
  if (configuredPathStatus === 'invalid') {
    return '';
  }

  const envPath = deps.getEnv('SUBMINER_MPV_PATH')?.trim();
  if (envPath && deps.fileExists(envPath)) {
    return envPath;
  }

  const whereResult = deps.runWhere();
  if (whereResult.status === 0) {
    const firstPath = whereResult.stdout
      .split(/\r?\n/)
      .map((line) => line.trim())
      .find((line) => line.length > 0 && deps.fileExists(line));
    if (firstPath) {
      return firstPath;
    }
  }

  return '';
}

export function spawnMpvProcess(
  executablePath: string,
  args: string[],
  env: NodeJS.ProcessEnv = process.env,
): ReturnType<typeof spawn> {
  const forceX11 = shouldForceX11WaylandSession(env);
  return spawn(executablePath, forceX11 ? [...args, ...MPV_X11_BACKEND_ARGS] : args, {
    detached: true,
    stdio: 'ignore',
    windowsHide: true,
    env: forceX11 ? applyX11EnvOverrides({ ...env }) : env,
  });
}

export interface PosixMpvPathDeps {
  platform: NodeJS.Platform;
  getEnv: (name: string) => string | undefined;
  isExecutable: (candidate: string) => boolean;
}

// Apps opened from Finder, the Dock, or Spotlight inherit launchd's minimal PATH
// (/usr/bin:/bin:/usr/sbin:/sbin), so Homebrew and app-bundle installs are probed directly.
const MACOS_MPV_FALLBACK_PATHS = [
  '/opt/homebrew/bin/mpv',
  '/usr/local/bin/mpv',
  '/Applications/mpv.app/Contents/MacOS/mpv',
];

function isExecutableFile(candidate: string): boolean {
  try {
    fs.accessSync(candidate, fs.constants.X_OK);
    return fs.statSync(candidate).isFile();
  } catch {
    return false;
  }
}

export function createPosixMpvPathDeps(
  overrides: Partial<PosixMpvPathDeps> = {},
): PosixMpvPathDeps {
  return {
    platform: overrides.platform ?? process.platform,
    getEnv: overrides.getEnv ?? ((name) => process.env[name]),
    isExecutable: overrides.isExecutable ?? isExecutableFile,
  };
}

// Resolution order: mpv.executablePath, SUBMINER_MPV_PATH, PATH, then macOS install
// locations. A configured path that is not executable resolves to '' instead of falling back.
export function resolvePosixMpvPath(deps: PosixMpvPathDeps, configuredMpvPath = ''): string {
  const configPath = configuredMpvPath.trim();
  if (configPath) {
    return deps.isExecutable(configPath) ? configPath : '';
  }

  const envPath = deps.getEnv('SUBMINER_MPV_PATH')?.trim();
  if (envPath && deps.isExecutable(envPath)) {
    return envPath;
  }

  const pathCandidates = (deps.getEnv('PATH') ?? '')
    .split(path.posix.delimiter)
    .filter((dir) => dir.length > 0)
    .map((dir) => path.posix.join(dir, 'mpv'));
  const candidates =
    deps.platform === 'darwin' ? [...pathCandidates, ...MACOS_MPV_FALLBACK_PATHS] : pathCandidates;
  return candidates.find((candidate) => deps.isExecutable(candidate)) ?? '';
}

// Returns the mpv binary SubMiner should spawn itself, or '' when none is found.
export function findMpvExecutablePath(configuredMpvPath = ''): string {
  return process.platform === 'win32'
    ? resolveWindowsMpvPath(createWindowsMpvPathDeps(), configuredMpvPath)
    : resolvePosixMpvPath(createPosixMpvPathDeps(), configuredMpvPath);
}

export function resolveMpvExecutablePath(configuredMpvPath = ''): string {
  const executablePath = findMpvExecutablePath(configuredMpvPath);
  if (!executablePath) {
    const binary = process.platform === 'win32' ? 'mpv.exe' : 'mpv';
    throw new Error(
      `Could not find ${binary}. Check mpv.executablePath, SUBMINER_MPV_PATH, or PATH.`,
    );
  }
  return executablePath;
}
