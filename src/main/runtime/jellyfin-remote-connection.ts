import { buildMpvLaunchModeArgs } from '../../shared/mpv-launch-mode';
import {
  buildSubminerPluginRuntimeScriptOptParts,
  type SubminerPluginRuntimeScriptOptConfig,
} from '../../shared/subminer-plugin-script-opts';
import type { MpvLaunchMode } from '../../types/config';
import type { InstalledMpvPluginDetection } from './first-run-setup-plugin';

type MpvClientLike = {
  connected: boolean;
  connect: () => void;
};

type SpawnedProcessLike = {
  on: (event: 'error', listener: (error: unknown) => void) => void;
  unref: () => void;
};

export type WaitForMpvConnectedDeps = {
  getMpvClient: () => MpvClientLike | null;
  now: () => number;
  sleep: (delayMs: number) => Promise<void>;
};

export function createWaitForMpvConnectedHandler(deps: WaitForMpvConnectedDeps) {
  return async (timeoutMs = 7000): Promise<boolean> => {
    const client = deps.getMpvClient();
    if (!client) return false;
    if (client.connected) return true;
    try {
      client.connect();
    } catch {}

    const startedAt = deps.now();
    while (deps.now() - startedAt < timeoutMs) {
      if (deps.getMpvClient()?.connected) return true;
      await deps.sleep(100);
    }
    return Boolean(deps.getMpvClient()?.connected);
  };
}

export type LaunchMpvForJellyfinDeps = {
  getSocketPath: () => string;
  getLaunchMode: () => MpvLaunchMode;
  platform: NodeJS.Platform;
  execPath: string;
  getRuntimePluginEntrypoint?: () => string | null | undefined;
  getInstalledPluginDetection?: () => InstalledMpvPluginDetection;
  getPluginRuntimeConfig?: () => SubminerPluginRuntimeScriptOptConfig;
  getDefaultMpvLogPath: () => string;
  defaultMpvArgs: readonly string[];
  removeSocketPath: (socketPath: string) => void;
  spawnMpv: (args: string[]) => SpawnedProcessLike;
  logWarn: (message: string, error: unknown) => void;
  logInfo: (message: string) => void;
};

export function createLaunchMpvIdleForJellyfinPlaybackHandler(deps: LaunchMpvForJellyfinDeps) {
  return (): void => {
    const socketPath = deps.getSocketPath();
    if (deps.platform !== 'win32') {
      try {
        deps.removeSocketPath(socketPath);
      } catch {
        // ignore stale socket cleanup errors
      }
    }

    const pluginRuntimeConfig = deps.getPluginRuntimeConfig?.();
    const scriptOptParts = pluginRuntimeConfig
      ? buildSubminerPluginRuntimeScriptOptParts(
          {
            ...pluginRuntimeConfig,
            socketPath,
          },
          deps.execPath,
        )
      : [`subminer-binary_path=${deps.execPath}`, `subminer-socket_path=${socketPath}`];
    const scriptOpts = `--script-opts=${scriptOptParts.join(',')}`;
    const installedPlugin = deps.getInstalledPluginDetection?.();
    const runtimePluginEntrypoint = installedPlugin?.installed
      ? ''
      : (deps.getRuntimePluginEntrypoint?.()?.trim() ?? '');
    if (installedPlugin?.installed && installedPlugin.path) {
      deps.logInfo(`Using installed mpv plugin for Jellyfin playback: ${installedPlugin.path}`);
    }
    const defaultMpvLogPath = deps.getDefaultMpvLogPath();
    const mpvArgs = [
      ...deps.defaultMpvArgs,
      ...buildMpvLaunchModeArgs(deps.getLaunchMode()),
      ...(runtimePluginEntrypoint ? [`--script=${runtimePluginEntrypoint}`] : []),
      '--idle=yes',
      scriptOpts,
      ...(defaultMpvLogPath.trim() ? [`--log-file=${defaultMpvLogPath}`] : []),
      `--input-ipc-server=${socketPath}`,
    ];
    const proc = deps.spawnMpv(mpvArgs);
    proc.on('error', (error) => {
      deps.logWarn('Failed to launch mpv for Jellyfin remote playback', error);
    });
    proc.unref();
    deps.logInfo(`Launched mpv for Jellyfin playback on socket: ${socketPath}`);
  };
}

export type EnsureMpvConnectedDeps = {
  getMpvClient: () => MpvClientLike | null;
  setMpvClient: (client: MpvClientLike | null) => void;
  createMpvClient: () => MpvClientLike;
  waitForMpvConnected: (timeoutMs: number) => Promise<boolean>;
  launchMpvIdleForJellyfinPlayback: () => void;
  getAutoLaunchInFlight: () => Promise<boolean> | null;
  setAutoLaunchInFlight: (promise: Promise<boolean> | null) => void;
  connectTimeoutMs: number;
  autoLaunchTimeoutMs: number;
};

export function createEnsureMpvConnectedForJellyfinPlaybackHandler(deps: EnsureMpvConnectedDeps) {
  return async (): Promise<boolean> => {
    if (!deps.getMpvClient()) {
      deps.setMpvClient(deps.createMpvClient());
    }

    const connected = await deps.waitForMpvConnected(deps.connectTimeoutMs);
    if (connected) return true;

    if (!deps.getAutoLaunchInFlight()) {
      const inFlight = (async () => {
        deps.launchMpvIdleForJellyfinPlayback();
        return deps.waitForMpvConnected(deps.autoLaunchTimeoutMs);
      })().finally(() => {
        deps.setAutoLaunchInFlight(null);
      });
      deps.setAutoLaunchInFlight(inFlight);
    }

    return deps.getAutoLaunchInFlight() as Promise<boolean>;
  };
}
