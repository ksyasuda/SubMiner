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

export function resolveWindowsMpvPath(
  deps: WindowsMpvLaunchDeps,
  configuredMpvPath = '',
): string {
  const configPath = normalizeCandidate(configuredMpvPath);
  if (configPath && deps.fileExists(configPath)) {
    return configPath;
  }

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

const DEFAULT_WINDOWS_MPV_SOCKET = '\\\\.\\pipe\\subminer-socket';

function readExtraArgValue(extraArgs: string[], flag: string): string | undefined {
  let value: string | undefined;
  for (let i = 0; i < extraArgs.length; i += 1) {
    const arg = extraArgs[i];
    if (arg === flag) {
      const next = extraArgs[i + 1];
      if (next && !next.startsWith('-')) {
        value = next;
        i += 1;
      }
      continue;
    }
    if (arg?.startsWith(`${flag}=`)) {
      value = arg.slice(flag.length + 1);
    }
  }
  return value;
}

export function buildWindowsMpvLaunchArgs(
  targets: string[],
  extraArgs: string[] = [],
  binaryPath?: string,
  pluginEntrypointPath?: string,
): string[] {
  const launchIdle = targets.length === 0;
  const inputIpcServer =
    readExtraArgValue(extraArgs, '--input-ipc-server') ?? DEFAULT_WINDOWS_MPV_SOCKET;
  const scriptOpts =
    typeof binaryPath === 'string' && binaryPath.trim().length > 0
      ? `--script-opts=subminer-binary_path=${binaryPath.trim().replace(/,/g, '\\,')},subminer-socket_path=${inputIpcServer.replace(/,/g, '\\,')}`
      : null;
  const scriptEntrypoint =
    typeof pluginEntrypointPath === 'string' && pluginEntrypointPath.trim().length > 0
      ? `--script=${pluginEntrypointPath.trim()}`
      : null;

  return [
    '--player-operation-mode=pseudo-gui',
    '--force-window=immediate',
    ...(launchIdle ? ['--idle=yes'] : []),
    ...(scriptEntrypoint ? [scriptEntrypoint] : []),
    `--input-ipc-server=${inputIpcServer}`,
    '--alang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
    '--slang=ja,jp,jpn,japanese,en,eng,english,enus,en-us',
    '--sub-auto=fuzzy',
    '--sub-file-paths=subs;subtitles',
    '--sid=auto',
    '--secondary-sid=auto',
    '--secondary-sub-visibility=no',
    ...(scriptOpts ? [scriptOpts] : []),
    ...extraArgs,
    ...targets,
  ];
}

export function launchWindowsMpv(
  targets: string[],
  deps: WindowsMpvLaunchDeps,
  extraArgs: string[] = [],
  binaryPath?: string,
  pluginEntrypointPath?: string,
  configuredMpvPath?: string,
): { ok: boolean; mpvPath: string } {
  const mpvPath = resolveWindowsMpvPath(deps, configuredMpvPath);
  if (!mpvPath) {
    deps.showError(
      'SubMiner mpv launcher',
      'Could not find mpv.exe. Set mpv.executablePath, set SUBMINER_MPV_PATH, or add mpv.exe to PATH.',
    );
    return { ok: false, mpvPath: '' };
  }

  try {
    deps.spawnDetached(
      mpvPath,
      buildWindowsMpvLaunchArgs(targets, extraArgs, binaryPath, pluginEntrypointPath),
    );
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
