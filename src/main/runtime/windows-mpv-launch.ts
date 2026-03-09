import fs from 'node:fs';
import { spawn, spawnSync } from 'node:child_process';

export interface WindowsMpvLaunchDeps {
  getEnv: (name: string) => string | undefined;
  runWhere: () => { status: number | null; stdout: string; error?: Error };
  fileExists: (candidate: string) => boolean;
  spawnDetached: (command: string, args: string[]) => void;
  showError: (title: string, content: string) => void;
}

function normalizeCandidate(candidate: string | undefined): string {
  return typeof candidate === 'string' ? candidate.trim() : '';
}

export function resolveWindowsMpvPath(deps: WindowsMpvLaunchDeps): string {
  const envPath = normalizeCandidate(deps.getEnv('SUBMINER_MPV_PATH'));
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

export function buildWindowsMpvLaunchArgs(targets: string[]): string[] {
  return ['--player-operation-mode=pseudo-gui', '--profile=subminer', ...targets];
}

export function launchWindowsMpv(
  targets: string[],
  deps: WindowsMpvLaunchDeps,
): { ok: boolean; mpvPath: string } {
  const mpvPath = resolveWindowsMpvPath(deps);
  if (!mpvPath) {
    deps.showError(
      'SubMiner mpv launcher',
      'Could not find mpv.exe. Install mpv and add it to PATH, or set SUBMINER_MPV_PATH.',
    );
    return { ok: false, mpvPath: '' };
  }

  try {
    deps.spawnDetached(mpvPath, buildWindowsMpvLaunchArgs(targets));
    return { ok: true, mpvPath };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    deps.showError('SubMiner mpv launcher', `Failed to launch mpv.\nPath: ${mpvPath}\n${message}`);
    return { ok: false, mpvPath };
  }
}

export function createWindowsMpvLaunchDeps(options: {
  getEnv?: (name: string) => string | undefined;
  fileExists?: (candidate: string) => boolean;
  showError: (title: string, content: string) => void;
}): WindowsMpvLaunchDeps {
  return {
    getEnv: options.getEnv ?? ((name) => process.env[name]),
    runWhere: () => {
      const result = spawnSync('where.exe', ['mpv.exe'], {
        encoding: 'utf8',
        windowsHide: true,
      });
      return {
        status: result.status,
        stdout: result.stdout ?? '',
        error: result.error ?? undefined,
      };
    },
    fileExists:
      options.fileExists ??
      ((candidate) => {
        try {
          return fs.statSync(candidate).isFile();
        } catch {
          return false;
        }
      }),
    spawnDetached: (command, args) => {
      const child = spawn(command, args, {
        detached: true,
        stdio: 'ignore',
        windowsHide: true,
      });
      child.unref();
    },
    showError: options.showError,
  };
}
