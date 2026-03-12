import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { resolveConfigDir } from '../config/path-resolution';

export type SetupStateStatus = 'incomplete' | 'in_progress' | 'completed' | 'cancelled';
export type SetupCompletionSource = 'user' | 'legacy_auto_detected' | null;
export type SetupPluginInstallStatus = 'unknown' | 'installed' | 'skipped' | 'failed';
export type SetupWindowsMpvShortcutInstallStatus = 'unknown' | 'installed' | 'skipped' | 'failed';

export interface SetupWindowsMpvShortcutPreferences {
  startMenuEnabled: boolean;
  desktopEnabled: boolean;
}

export interface SetupState {
  version: 2;
  status: SetupStateStatus;
  completedAt: string | null;
  completionSource: SetupCompletionSource;
  lastSeenYomitanDictionaryCount: number;
  pluginInstallStatus: SetupPluginInstallStatus;
  pluginInstallPathSummary: string | null;
  windowsMpvShortcutPreferences: SetupWindowsMpvShortcutPreferences;
  windowsMpvShortcutLastStatus: SetupWindowsMpvShortcutInstallStatus;
}

export interface ConfigFilePaths {
  jsoncPath: string;
  jsonPath: string;
}

export interface MpvInstallPaths {
  supported: boolean;
  mpvConfigDir: string;
  scriptsDir: string;
  scriptOptsDir: string;
  pluginEntrypointPath: string;
  pluginDir: string;
  pluginConfigPath: string;
}

function getPlatformPath(platform: NodeJS.Platform): typeof path.posix | typeof path.win32 {
  return platform === 'win32' ? path.win32 : path.posix;
}

function asObject(value: unknown): Record<string, unknown> | null {
  return value && typeof value === 'object' && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

export function createDefaultSetupState(): SetupState {
  return {
    version: 2,
    status: 'incomplete',
    completedAt: null,
    completionSource: null,
    lastSeenYomitanDictionaryCount: 0,
    pluginInstallStatus: 'unknown',
    pluginInstallPathSummary: null,
    windowsMpvShortcutPreferences: {
      startMenuEnabled: true,
      desktopEnabled: true,
    },
    windowsMpvShortcutLastStatus: 'unknown',
  };
}

export function normalizeSetupState(value: unknown): SetupState | null {
  const record = asObject(value);
  if (!record) return null;
  const version = record.version;
  const status = record.status;
  const pluginInstallStatus = record.pluginInstallStatus;
  const completionSource = record.completionSource;
  const windowsPrefs = asObject(record.windowsMpvShortcutPreferences);
  const windowsMpvShortcutLastStatus = record.windowsMpvShortcutLastStatus;

  if (
    (version !== 1 && version !== 2) ||
    (status !== 'incomplete' &&
      status !== 'in_progress' &&
      status !== 'completed' &&
      status !== 'cancelled') ||
    (pluginInstallStatus !== 'unknown' &&
      pluginInstallStatus !== 'installed' &&
      pluginInstallStatus !== 'skipped' &&
      pluginInstallStatus !== 'failed') ||
    (version === 2 &&
      windowsMpvShortcutLastStatus !== 'unknown' &&
      windowsMpvShortcutLastStatus !== 'installed' &&
      windowsMpvShortcutLastStatus !== 'skipped' &&
      windowsMpvShortcutLastStatus !== 'failed') ||
    (completionSource !== null &&
      completionSource !== 'user' &&
      completionSource !== 'legacy_auto_detected')
  ) {
    return null;
  }

  return {
    version: 2,
    status,
    completedAt: typeof record.completedAt === 'string' ? record.completedAt : null,
    completionSource,
    lastSeenYomitanDictionaryCount:
      typeof record.lastSeenYomitanDictionaryCount === 'number' &&
      Number.isFinite(record.lastSeenYomitanDictionaryCount) &&
      record.lastSeenYomitanDictionaryCount >= 0
        ? Math.floor(record.lastSeenYomitanDictionaryCount)
        : 0,
    pluginInstallStatus,
    pluginInstallPathSummary:
      typeof record.pluginInstallPathSummary === 'string' ? record.pluginInstallPathSummary : null,
    windowsMpvShortcutPreferences: {
      startMenuEnabled:
        version === 2 && typeof windowsPrefs?.startMenuEnabled === 'boolean'
          ? windowsPrefs.startMenuEnabled
          : true,
      desktopEnabled:
        version === 2 && typeof windowsPrefs?.desktopEnabled === 'boolean'
          ? windowsPrefs.desktopEnabled
          : true,
    },
    windowsMpvShortcutLastStatus:
      version === 2 &&
      (windowsMpvShortcutLastStatus === 'unknown' ||
        windowsMpvShortcutLastStatus === 'installed' ||
        windowsMpvShortcutLastStatus === 'skipped' ||
        windowsMpvShortcutLastStatus === 'failed')
        ? windowsMpvShortcutLastStatus
        : 'unknown',
  };
}

export function isSetupCompleted(state: SetupState | null | undefined): boolean {
  return state?.status === 'completed';
}

export function getDefaultConfigDir(options?: {
  platform?: NodeJS.Platform;
  appDataDir?: string;
  xdgConfigHome?: string;
  homeDir?: string;
  existsSync?: (candidate: string) => boolean;
}): string {
  return resolveConfigDir({
    platform: options?.platform ?? process.platform,
    appDataDir: options?.appDataDir ?? process.env.APPDATA,
    xdgConfigHome: options?.xdgConfigHome ?? process.env.XDG_CONFIG_HOME,
    homeDir: options?.homeDir ?? os.homedir(),
    existsSync: options?.existsSync ?? fs.existsSync,
  });
}

export function getDefaultConfigFilePaths(configDir: string): ConfigFilePaths {
  return {
    jsoncPath: path.join(configDir, 'config.jsonc'),
    jsonPath: path.join(configDir, 'config.json'),
  };
}

export function getSetupStatePath(configDir: string): string {
  return path.join(configDir, 'setup-state.json');
}

export function readSetupState(
  statePath: string,
  deps?: {
    existsSync?: (candidate: string) => boolean;
    readFileSync?: (candidate: string, encoding: BufferEncoding) => string;
  },
): SetupState | null {
  const existsSync = deps?.existsSync ?? fs.existsSync;
  const readFileSync = deps?.readFileSync ?? fs.readFileSync;
  if (!existsSync(statePath)) return null;

  try {
    return normalizeSetupState(JSON.parse(readFileSync(statePath, 'utf8')));
  } catch {
    return null;
  }
}

export function writeSetupState(
  statePath: string,
  state: SetupState,
  deps?: {
    mkdirSync?: (candidate: string, options: { recursive: true }) => void;
    writeFileSync?: (candidate: string, content: string, encoding: BufferEncoding) => void;
  },
): void {
  const mkdirSync = deps?.mkdirSync ?? fs.mkdirSync;
  const writeFileSync = deps?.writeFileSync ?? fs.writeFileSync;
  mkdirSync(path.dirname(statePath), { recursive: true });
  writeFileSync(statePath, `${JSON.stringify(state, null, 2)}\n`, 'utf8');
}

export function ensureDefaultConfigBootstrap(options: {
  configDir: string;
  configFilePaths: ConfigFilePaths;
  generateTemplate: () => string;
  existsSync?: (candidate: string) => boolean;
  mkdirSync?: (candidate: string, options: { recursive: true }) => void;
  writeFileSync?: (candidate: string, content: string, encoding: BufferEncoding) => void;
}): void {
  const existsSync = options.existsSync ?? fs.existsSync;
  const mkdirSync = options.mkdirSync ?? fs.mkdirSync;
  const writeFileSync = options.writeFileSync ?? fs.writeFileSync;

  if (existsSync(options.configFilePaths.jsoncPath) || existsSync(options.configFilePaths.jsonPath)) {
    return;
  }

  mkdirSync(options.configDir, { recursive: true });
  writeFileSync(options.configFilePaths.jsoncPath, options.generateTemplate(), 'utf8');
}

export function resolveDefaultMpvInstallPaths(
  platform: NodeJS.Platform,
  homeDir: string,
  xdgConfigHome?: string,
): MpvInstallPaths {
  const platformPath = getPlatformPath(platform);
  const mpvConfigDir =
    platform === 'darwin'
      ? platformPath.join(homeDir, 'Library', 'Application Support', 'mpv')
      : platform === 'linux'
        ? platformPath.join(xdgConfigHome?.trim() || platformPath.join(homeDir, '.config'), 'mpv')
        : platformPath.join(homeDir, 'AppData', 'Roaming', 'mpv');

  return {
    supported: platform === 'linux' || platform === 'darwin' || platform === 'win32',
    mpvConfigDir,
    scriptsDir: platformPath.join(mpvConfigDir, 'scripts'),
    scriptOptsDir: platformPath.join(mpvConfigDir, 'script-opts'),
    pluginEntrypointPath: platformPath.join(mpvConfigDir, 'scripts', 'subminer', 'main.lua'),
    pluginDir: platformPath.join(mpvConfigDir, 'scripts', 'subminer'),
    pluginConfigPath: platformPath.join(mpvConfigDir, 'script-opts', 'subminer.conf'),
  };
}
