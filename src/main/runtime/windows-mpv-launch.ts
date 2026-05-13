import fs from 'node:fs';
import { spawn, spawnSync } from 'node:child_process';
import { buildMpvLaunchModeArgs } from '../../shared/mpv-launch-mode';
import type { MpvLaunchMode } from '../../types/config';
import type { InstalledMpvPluginDetection } from './first-run-setup-plugin';

export interface WindowsMpvLaunchDeps {
  getEnv: (name: string) => string | undefined;
  runWhere: () => { status: number | null; stdout: string; error?: Error };
  fileExists: (candidate: string) => boolean;
  spawnDetached: (command: string, args: string[]) => Promise<void>;
  showError: (title: string, content: string) => void;
}

export type ConfiguredWindowsMpvPathStatus = 'blank' | 'configured' | 'invalid';

export interface WindowsMpvRuntimePluginPolicy {
  detectInstalledMpvPlugin?: (mpvPath: string) => InstalledMpvPluginDetection;
  notifyInstalledPluginDetected?: (detection: InstalledMpvPluginDetection) => void;
  resolveInstalledPluginBeforeLaunch?: (
    detection: InstalledMpvPluginDetection,
    mpvPath: string,
  ) => Promise<'removed' | 'continue' | 'cancel'> | 'removed' | 'continue' | 'cancel';
}

function normalizeCandidate(candidate: string | undefined): string {
  return typeof candidate === 'string' ? candidate.trim() : '';
}

function defaultWindowsMpvFileExists(candidate: string): boolean {
  try {
    return fs.statSync(candidate).isFile();
  } catch {
    return false;
  }
}

export function getConfiguredWindowsMpvPathStatus(
  configuredMpvPath = '',
  fileExists: (candidate: string) => boolean = defaultWindowsMpvFileExists,
): ConfiguredWindowsMpvPathStatus {
  const configPath = normalizeCandidate(configuredMpvPath);
  if (!configPath) {
    return 'blank';
  }
  return fileExists(configPath) ? 'configured' : 'invalid';
}

export function resolveWindowsMpvPath(deps: WindowsMpvLaunchDeps, configuredMpvPath = ''): string {
  const configPath = normalizeCandidate(configuredMpvPath);
  const configuredPathStatus = getConfiguredWindowsMpvPathStatus(configPath, deps.fileExists);
  if (configuredPathStatus === 'configured') {
    return configPath;
  }
  if (configuredPathStatus === 'invalid') {
    return '';
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
  launchMode: MpvLaunchMode = 'normal',
): string[] {
  const launchIdle = targets.length === 0;
  const inputIpcServer =
    readExtraArgValue(extraArgs, '--input-ipc-server') ?? DEFAULT_WINDOWS_MPV_SOCKET;
  const scriptEntrypoint =
    typeof pluginEntrypointPath === 'string' && pluginEntrypointPath.trim().length > 0
      ? `--script=${pluginEntrypointPath.trim()}`
      : null;
  const hasBinaryPath = typeof binaryPath === 'string' && binaryPath.trim().length > 0;
  const shouldPassSubminerScriptOpts = scriptEntrypoint || hasBinaryPath;
  const scriptOptPairs = shouldPassSubminerScriptOpts
    ? [`subminer-socket_path=${inputIpcServer.replace(/,/g, '\\,')}`]
    : [];
  if (hasBinaryPath) {
    scriptOptPairs.unshift(`subminer-binary_path=${binaryPath.trim().replace(/,/g, '\\,')}`);
  }
  const scriptOpts = scriptOptPairs.length > 0 ? `--script-opts=${scriptOptPairs.join(',')}` : null;

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
    ...buildMpvLaunchModeArgs(launchMode),
    ...extraArgs,
    ...targets,
  ];
}

export async function launchWindowsMpv(
  targets: string[],
  deps: WindowsMpvLaunchDeps,
  extraArgs: string[] = [],
  binaryPath?: string,
  pluginEntrypointPath?: string,
  configuredMpvPath?: string,
  launchMode: MpvLaunchMode = 'normal',
  runtimePluginPolicy?: WindowsMpvRuntimePluginPolicy,
): Promise<{ ok: boolean; mpvPath: string }> {
  const normalizedConfiguredPath = normalizeCandidate(configuredMpvPath);
  const mpvPath = resolveWindowsMpvPath(deps, normalizedConfiguredPath);
  if (!mpvPath) {
    deps.showError(
      'SubMiner mpv launcher',
      normalizedConfiguredPath
        ? `Configured mpv.executablePath was not found: ${normalizedConfiguredPath}`
        : 'Could not find mpv.exe. Set mpv.executablePath, set SUBMINER_MPV_PATH, or add mpv.exe to PATH.',
    );
    return { ok: false, mpvPath: '' };
  }

  try {
    let installedPlugin = runtimePluginPolicy?.detectInstalledMpvPlugin?.(mpvPath);
    let installedPluginPrompted = false;
    if (installedPlugin?.installed) {
      const resolution = await runtimePluginPolicy?.resolveInstalledPluginBeforeLaunch?.(
        installedPlugin,
        mpvPath,
      );
      installedPluginPrompted = resolution != null;
      if (resolution === 'cancel') {
        return { ok: false, mpvPath };
      }
      if (resolution === 'removed') {
        installedPlugin = runtimePluginPolicy?.detectInstalledMpvPlugin?.(mpvPath);
      }
    }
    const runtimePluginEntrypointPath = installedPlugin?.installed
      ? undefined
      : pluginEntrypointPath;
    if (installedPlugin?.installed && !installedPluginPrompted) {
      runtimePluginPolicy?.notifyInstalledPluginDetected?.(installedPlugin);
    }
    await deps.spawnDetached(
      mpvPath,
      buildWindowsMpvLaunchArgs(
        targets,
        extraArgs,
        binaryPath,
        runtimePluginEntrypointPath,
        launchMode,
      ),
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
    fileExists: options.fileExists ?? defaultWindowsMpvFileExists,
    spawnDetached: (command, args) =>
      new Promise((resolve, reject) => {
        try {
          const child = spawn(command, args, {
            detached: true,
            stdio: 'ignore',
            windowsHide: true,
          });
          let settled = false;
          child.once('error', (error) => {
            if (settled) return;
            settled = true;
            reject(error);
          });
          child.once('spawn', () => {
            if (settled) return;
            settled = true;
            child.unref();
            resolve();
          });
        } catch (error) {
          reject(error);
        }
      }),
    showError: options.showError,
  };
}
