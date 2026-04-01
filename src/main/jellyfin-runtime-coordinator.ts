import fs from 'node:fs';
import path from 'node:path';
import { spawn } from 'node:child_process';

import { BrowserWindow } from 'electron';

import { DEFAULT_CONFIG } from '../config';
import {
  JellyfinRemoteSessionService,
  authenticateWithPasswordRuntime,
  jellyfinTicksToSecondsRuntime,
  listJellyfinItemsRuntime,
  listJellyfinLibrariesRuntime,
  listJellyfinSubtitleTracksRuntime,
  resolveJellyfinPlaybackPlanRuntime,
  sendMpvCommandRuntime,
} from '../core/services';
import type { MpvIpcClient } from '../core/services/mpv';
import type { JellyfinSetupWindowLike } from './jellyfin-runtime';
import { createJellyfinRuntime } from './jellyfin-runtime';

export interface JellyfinRuntimeCoordinatorInput {
  getResolvedConfig: Parameters<typeof createJellyfinRuntime>[0]['getResolvedConfig'];
  configService: {
    patchRawConfig: Parameters<typeof createJellyfinRuntime>[0]['patchRawConfig'];
  };
  tokenStore: Parameters<typeof createJellyfinRuntime>[0]['tokenStore'];
  platform: NodeJS.Platform;
  execPath: string;
  defaultMpvLogPath: string;
  defaultMpvArgs: readonly string[];
  connectTimeoutMs: number;
  autoLaunchTimeoutMs: number;
  langPref: string;
  progressIntervalMs: number;
  ticksPerSecond: number;
  appState: {
    mpvSocketPath: string;
    mpvClient: MpvIpcClient | null;
    jellyfinSetupWindow: BrowserWindow | null;
  };
  actions: {
    createMpvClient: () => MpvIpcClient;
    applyJellyfinMpvDefaults: (client: MpvIpcClient) => void;
    showMpvOsd: (message: string) => void;
  };
  logger: {
    info: (message: string) => void;
    warn: (message: string, details?: unknown) => void;
    debug: (message: string, details?: unknown) => void;
    error: (message: string, error?: unknown) => void;
  };
}

export function createJellyfinRuntimeCoordinator(input: JellyfinRuntimeCoordinatorInput) {
  return createJellyfinRuntime<JellyfinSetupWindowLike>({
    getResolvedConfig: () => input.getResolvedConfig(),
    getEnv: (name) => process.env[name],
    patchRawConfig: (patch) => {
      input.configService.patchRawConfig(patch);
    },
    defaultJellyfinConfig: DEFAULT_CONFIG.jellyfin,
    tokenStore: input.tokenStore,
    platform: input.platform,
    execPath: input.execPath,
    defaultMpvLogPath: input.defaultMpvLogPath,
    defaultMpvArgs: [...input.defaultMpvArgs],
    connectTimeoutMs: input.connectTimeoutMs,
    autoLaunchTimeoutMs: input.autoLaunchTimeoutMs,
    langPref: input.langPref,
    progressIntervalMs: input.progressIntervalMs,
    ticksPerSecond: input.ticksPerSecond,
    getMpvSocketPath: () => input.appState.mpvSocketPath,
    getMpvClient: () => input.appState.mpvClient,
    setMpvClient: (client) => {
      input.appState.mpvClient = client as MpvIpcClient | null;
    },
    createMpvClient: () => input.actions.createMpvClient(),
    sendMpvCommand: (client, command) => sendMpvCommandRuntime(client as MpvIpcClient, command),
    applyJellyfinMpvDefaults: (client) =>
      input.actions.applyJellyfinMpvDefaults(client as MpvIpcClient),
    showMpvOsd: (message) => input.actions.showMpvOsd(message),
    removeSocketPath: (socketPath) => {
      fs.rmSync(socketPath, { force: true });
    },
    spawnMpv: (args) =>
      spawn('mpv', args, {
        detached: true,
        stdio: 'ignore',
      }),
    wait: (delayMs) => new Promise((resolve) => setTimeout(resolve, delayMs)),
    authenticateWithPassword: (serverUrl, username, password, clientInfo) =>
      authenticateWithPasswordRuntime(
        serverUrl,
        username,
        password,
        clientInfo as Parameters<typeof authenticateWithPasswordRuntime>[3],
      ),
    listJellyfinLibraries: (session, clientInfo) =>
      listJellyfinLibrariesRuntime(
        session as Parameters<typeof listJellyfinLibrariesRuntime>[0],
        clientInfo as Parameters<typeof listJellyfinLibrariesRuntime>[1],
      ),
    listJellyfinItems: (session, clientInfo, params) =>
      listJellyfinItemsRuntime(
        session as Parameters<typeof listJellyfinItemsRuntime>[0],
        clientInfo as Parameters<typeof listJellyfinItemsRuntime>[1],
        params as Parameters<typeof listJellyfinItemsRuntime>[2],
      ),
    listJellyfinSubtitleTracks: (session, clientInfo, itemId) =>
      listJellyfinSubtitleTracksRuntime(
        session as Parameters<typeof listJellyfinSubtitleTracksRuntime>[0],
        clientInfo as Parameters<typeof listJellyfinSubtitleTracksRuntime>[1],
        itemId,
      ),
    writeJellyfinPreviewAuth: (responsePath, payload) => {
      fs.mkdirSync(path.dirname(responsePath), { recursive: true });
      fs.writeFileSync(responsePath, JSON.stringify(payload, null, 2), 'utf-8');
    },
    resolvePlaybackPlan: (params) =>
      resolveJellyfinPlaybackPlanRuntime(
        (params as { session: Parameters<typeof resolveJellyfinPlaybackPlanRuntime>[0] }).session,
        (params as { clientInfo: Parameters<typeof resolveJellyfinPlaybackPlanRuntime>[1] })
          .clientInfo,
        (
          params as {
            jellyfinConfig: ReturnType<
              JellyfinRuntimeCoordinatorInput['getResolvedConfig']
            >['jellyfin'];
          }
        ).jellyfinConfig,
        {
          itemId: (params as { itemId: string }).itemId,
          audioStreamIndex:
            (params as { audioStreamIndex?: number | null }).audioStreamIndex ?? undefined,
          subtitleStreamIndex:
            (params as { subtitleStreamIndex?: number | null }).subtitleStreamIndex ?? undefined,
        },
      ),
    convertTicksToSeconds: (ticks) => jellyfinTicksToSecondsRuntime(ticks),
    createRemoteSessionService: (options) => new JellyfinRemoteSessionService(options as never),
    defaultDeviceId: DEFAULT_CONFIG.jellyfin.deviceId,
    defaultClientName: DEFAULT_CONFIG.jellyfin.clientName,
    defaultClientVersion: DEFAULT_CONFIG.jellyfin.clientVersion,
    createBrowserWindow: (options) => {
      const window = new BrowserWindow(options);
      input.appState.jellyfinSetupWindow = window;
      window.on('closed', () => {
        input.appState.jellyfinSetupWindow = null;
      });
      return window as unknown as JellyfinSetupWindowLike;
    },
    encodeURIComponent: (value) => encodeURIComponent(value),
    logInfo: (message) => input.logger.info(message),
    logWarn: (message, details) => input.logger.warn(message, details),
    logDebug: (message, details) => input.logger.debug(message, details),
    logError: (message, error) => input.logger.error(message, error),
    now: () => Date.now(),
  });
}
