import fs from 'node:fs';
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

export function resolveMpvExecutablePath(configuredMpvPath = ''): string {
  const executablePath =
    process.platform === 'win32'
      ? resolveWindowsMpvPath(createWindowsMpvPathDeps(), configuredMpvPath)
      : 'mpv';
  if (!executablePath) {
    throw new Error(
      'Could not find mpv.exe. Check mpv.executablePath, SUBMINER_MPV_PATH, or PATH.',
    );
  }
  return executablePath;
}
