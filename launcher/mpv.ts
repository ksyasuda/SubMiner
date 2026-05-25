import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import net from 'node:net';
import { spawn, spawnSync } from 'node:child_process';
import { buildMpvLaunchModeArgs } from '../src/shared/mpv-launch-mode.js';
import {
  isAppControlServerAvailable as checkAppControlServerAvailable,
  sendAppControlCommand,
} from '../src/shared/app-control-client.js';
import { getDefaultConfigDir } from '../src/shared/setup-state.js';
import {
  detectInstalledMpvPlugin,
  type InstalledMpvPluginDetection,
} from '../src/main/runtime/first-run-setup-plugin.js';
import type { LogLevel, Backend, Args, MpvTrack, PluginRuntimeConfig } from './types.js';
import { DEFAULT_MPV_SUBMINER_ARGS, DEFAULT_YOUTUBE_YTDL_FORMAT } from './types.js';
import { appendToAppLog, getAppLogPath, log, fail, getMpvLogPath } from './log.js';
import { buildSubminerScriptOpts, resolveAniSkipMetadataForFile } from './aniskip-metadata.js';
import { buildPluginRuntimeScriptOptParts } from './config/plugin-runtime-config.js';
import { nowMs } from './time.js';
import {
  commandExists,
  getPathEnv,
  isExecutable,
  resolveBinaryPathCandidate,
  resolveCommandInvocation,
  isYoutubeTarget,
  uniqueNormalizedLangCodes,
  sleep,
  normalizeLangCode,
  realpathMaybe,
} from './util.js';

export const state = {
  overlayProc: null as ReturnType<typeof spawn> | null,
  mpvProc: null as ReturnType<typeof spawn> | null,
  youtubeSubgenChildren: new Set<ReturnType<typeof spawn>>(),
  appPath: '' as string,
  overlayManagedByLauncher: false,
  stopRequested: false,
};

type SpawnTarget = {
  command: string;
  args: string[];
  env?: NodeJS.ProcessEnv;
};

type PathModule = Pick<
  typeof path,
  'dirname' | 'join' | 'extname' | 'delimiter' | 'sep' | 'resolve' | 'isAbsolute' | 'normalize'
>;

const DETACHED_IDLE_MPV_PID_FILE = path.join(os.tmpdir(), 'subminer-idle-mpv.pid');
const OVERLAY_START_SOCKET_READY_TIMEOUT_MS = 900;
const OVERLAY_START_COMMAND_SETTLE_TIMEOUT_MS = 700;
const TRANSPORTED_APP_ARGC_ENV = 'SUBMINER_APP_ARGC';
const TRANSPORTED_APP_ARG_PREFIX = 'SUBMINER_APP_ARG_';

export interface LauncherRuntimePluginPlan {
  scriptPath: string | null;
  installedPlugin: InstalledMpvPluginDetection;
  warningMessage: string | null;
  errorMessage: string | null;
}

function resolvePluginCandidatePath(candidate: string, pathModule: PathModule): string {
  return pathModule.isAbsolute(candidate)
    ? pathModule.normalize(candidate)
    : pathModule.resolve(candidate);
}

export function parseMpvArgString(input: string): string[] {
  const chars = input;
  const args: string[] = [];
  let current = '';
  let tokenStarted = false;
  let inSingleQuote = false;
  let inDoubleQuote = false;
  let escaping = false;
  const canEscape = (nextChar: string | undefined): boolean =>
    nextChar === undefined ||
    nextChar === '"' ||
    nextChar === "'" ||
    nextChar === '\\' ||
    /\s/.test(nextChar);

  for (let i = 0; i < chars.length; i += 1) {
    const ch = chars[i] || '';
    if (escaping) {
      current += ch;
      tokenStarted = true;
      escaping = false;
      continue;
    }

    if (inSingleQuote) {
      if (ch === "'") {
        inSingleQuote = false;
      } else {
        current += ch;
        tokenStarted = true;
      }
      continue;
    }

    if (inDoubleQuote) {
      if (ch === '\\') {
        if (canEscape(chars[i + 1])) {
          escaping = true;
        } else {
          current += ch;
          tokenStarted = true;
        }
        continue;
      }
      if (ch === '"') {
        inDoubleQuote = false;
        continue;
      }
      current += ch;
      tokenStarted = true;
      continue;
    }

    if (ch === '\\') {
      if (canEscape(chars[i + 1])) {
        escaping = true;
        tokenStarted = true;
      } else {
        current += ch;
        tokenStarted = true;
      }
      continue;
    }
    if (ch === "'") {
      tokenStarted = true;
      inSingleQuote = true;
      continue;
    }
    if (ch === '"') {
      tokenStarted = true;
      inDoubleQuote = true;
      continue;
    }
    if (/\s/.test(ch)) {
      if (tokenStarted) {
        args.push(current);
        current = '';
        tokenStarted = false;
      }
      continue;
    }
    current += ch;
    tokenStarted = true;
  }

  if (escaping) {
    fail('Could not parse mpv args: trailing backslash');
  }
  if (inSingleQuote || inDoubleQuote) {
    fail('Could not parse mpv args: unmatched quote');
  }
  if (tokenStarted) {
    args.push(current);
  }

  return args;
}

function readTrackedDetachedMpvPid(): number | null {
  try {
    const raw = fs.readFileSync(DETACHED_IDLE_MPV_PID_FILE, 'utf8').trim();
    const pid = Number.parseInt(raw, 10);
    return Number.isInteger(pid) && pid > 0 ? pid : null;
  } catch {
    return null;
  }
}

function clearTrackedDetachedMpvPid(): void {
  try {
    fs.rmSync(DETACHED_IDLE_MPV_PID_FILE, { force: true });
  } catch {
    // ignore
  }
}

function trackDetachedMpvPid(pid: number): void {
  try {
    fs.writeFileSync(DETACHED_IDLE_MPV_PID_FILE, String(pid), 'utf8');
  } catch {
    // ignore
  }
}

function isProcessAlive(pid: number): boolean {
  try {
    process.kill(pid, 0);
    return true;
  } catch {
    return false;
  }
}

function processLooksLikeMpv(pid: number): boolean {
  if (process.platform !== 'linux') return true;
  try {
    const cmdline = fs.readFileSync(`/proc/${pid}/cmdline`, 'utf8');
    return cmdline.includes('mpv');
  } catch {
    return false;
  }
}

async function terminateTrackedDetachedMpv(logLevel: LogLevel): Promise<void> {
  const pid = readTrackedDetachedMpvPid();
  if (!pid) return;
  if (!isProcessAlive(pid)) {
    clearTrackedDetachedMpvPid();
    return;
  }
  if (!processLooksLikeMpv(pid)) {
    clearTrackedDetachedMpvPid();
    return;
  }

  try {
    process.kill(pid, 'SIGTERM');
  } catch {
    clearTrackedDetachedMpvPid();
    return;
  }

  const deadline = nowMs() + 1500;
  while (nowMs() < deadline) {
    if (!isProcessAlive(pid)) {
      clearTrackedDetachedMpvPid();
      return;
    }
    await sleep(100);
  }

  try {
    process.kill(pid, 'SIGKILL');
  } catch {
    // ignore
  }
  clearTrackedDetachedMpvPid();
  log('debug', logLevel, `Terminated stale detached mpv pid=${pid}`);
}

export function makeTempDir(prefix: string): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), prefix));
}

function normalizeRuntimePluginEntrypoint(
  candidate: string,
  deps: {
    pathModule: typeof path;
    existsSync: (candidate: string) => boolean;
  },
): string | null {
  const trimmed = candidate.trim();
  if (!trimmed) return null;
  if (trimmed.endsWith('.lua')) {
    return deps.existsSync(trimmed) ? trimmed : null;
  }
  const entrypoint = deps.pathModule.join(trimmed, 'main.lua');
  return deps.existsSync(entrypoint) ? entrypoint : null;
}

function pushMacAppRuntimePluginCandidate(
  candidates: string[],
  appPath: string,
  pathModule: typeof path,
): void {
  const appIndex = appPath.indexOf('.app');
  if (appIndex < 0) return;
  candidates.push(
    pathModule.join(
      appPath.slice(0, appIndex + '.app'.length),
      'Contents',
      'Resources',
      'plugin',
      'subminer',
    ),
  );
}

export function resolveLauncherRuntimePluginPath(options: {
  appPath: string;
  scriptPath?: string;
  platform?: NodeJS.Platform;
  homeDir?: string;
  cwd?: string;
  env?: NodeJS.ProcessEnv;
  dirname?: string;
  pathModule?: typeof path;
  existsSync?: (candidate: string) => boolean;
}): string | null {
  const platform = options.platform ?? process.platform;
  const pathModule = options.pathModule ?? path;
  const existsSync = options.existsSync ?? fs.existsSync;
  const env = options.env ?? process.env;
  const dirname = options.dirname ?? __dirname;
  const cwd = options.cwd ?? process.cwd();
  const homeDir = options.homeDir ?? os.homedir();
  const candidates: string[] = [];

  if (env.SUBMINER_MPV_PLUGIN_PATH) {
    candidates.push(env.SUBMINER_MPV_PLUGIN_PATH);
  }

  pushMacAppRuntimePluginCandidate(candidates, options.appPath, pathModule);

  const appDir = pathModule.dirname(options.appPath);
  candidates.push(
    pathModule.join(appDir, 'resources', 'plugin', 'subminer'),
    pathModule.join(appDir, 'plugin', 'subminer'),
  );

  if (options.scriptPath) {
    const scriptDir = pathModule.dirname(realpathMaybe(options.scriptPath));
    candidates.push(
      pathModule.join(scriptDir, '..', 'share', 'SubMiner', 'plugin', 'subminer'),
      pathModule.join(scriptDir, '..', 'lib', 'SubMiner', 'plugin', 'subminer'),
      pathModule.join(scriptDir, 'plugin', 'subminer'),
    );
  }

  if (platform === 'darwin') {
    candidates.push(
      pathModule.join(homeDir, 'Library', 'Application Support', 'SubMiner', 'plugin', 'subminer'),
    );
  } else if (platform !== 'win32') {
    candidates.push(
      pathModule.join(
        env.XDG_DATA_HOME?.trim() || pathModule.join(homeDir, '.local', 'share'),
        'SubMiner',
        'plugin',
        'subminer',
      ),
    );
  }

  candidates.push(
    pathModule.join(cwd, 'plugin', 'subminer'),
    pathModule.join(dirname, '..', 'plugin', 'subminer'),
    pathModule.join(dirname, '..', '..', 'plugin', 'subminer'),
  );

  const seen = new Set<string>();
  for (const candidate of candidates) {
    const resolved = resolvePluginCandidatePath(candidate, pathModule);
    if (seen.has(resolved)) continue;
    seen.add(resolved);
    const entrypoint = normalizeRuntimePluginEntrypoint(resolved, { pathModule, existsSync });
    if (entrypoint) {
      return entrypoint;
    }
  }

  return null;
}

export function resolveLauncherRuntimePluginPlan(options: {
  runtimePluginPath?: string | null;
  platform?: NodeJS.Platform;
  homeDir?: string;
  xdgConfigHome?: string;
  appDataDir?: string;
  mpvExecutablePath?: string;
  existsSync?: (candidate: string) => boolean;
  readFileSync?: (candidate: string, encoding: BufferEncoding) => string;
}): LauncherRuntimePluginPlan {
  const installedPlugin = detectInstalledMpvPlugin({
    platform: options.platform ?? process.platform,
    homeDir: options.homeDir ?? os.homedir(),
    xdgConfigHome: options.xdgConfigHome ?? process.env.XDG_CONFIG_HOME,
    appDataDir: options.appDataDir ?? process.env.APPDATA,
    mpvExecutablePath: options.mpvExecutablePath,
    existsSync: options.existsSync,
    readFileSync: options.readFileSync,
  });

  if (installedPlugin.installed) {
    const versionText = installedPlugin.version
      ? `Detected plugin version: ${installedPlugin.version}.`
      : 'Detected plugin version: unknown or legacy.';
    return {
      scriptPath: null,
      installedPlugin,
      warningMessage: `SubMiner detected an installed mpv plugin at: ${installedPlugin.path}. This mpv session will use the installed plugin. Remove it to use SubMiner's bundled runtime plugin automatically. ${versionText}`,
      errorMessage: null,
    };
  }

  if (options.runtimePluginPath) {
    return {
      scriptPath: options.runtimePluginPath,
      installedPlugin,
      warningMessage: null,
      errorMessage: null,
    };
  }

  return {
    scriptPath: null,
    installedPlugin,
    warningMessage: null,
    errorMessage:
      'Packaged mpv plugin assets were not found. Install the SubMiner assets bundle or set SUBMINER_MPV_PLUGIN_PATH to plugin/subminer/main.lua.',
  };
}

function appendRuntimePluginLaunchArgs(
  mpvArgs: string[],
  plan: LauncherRuntimePluginPlan,
  logLevel: LogLevel,
): void {
  if (plan.warningMessage) {
    log('warn', logLevel, plan.warningMessage);
  }
  if (plan.errorMessage) {
    fail(plan.errorMessage);
  }
  if (plan.scriptPath) {
    mpvArgs.push(`--script=${plan.scriptPath}`);
  }
}

export function detectBackend(
  backend: Backend,
  env: NodeJS.ProcessEnv = process.env,
): Exclude<Backend, 'auto'> {
  if (backend !== 'auto') return backend;
  if (process.platform === 'win32') return 'windows';
  if (process.platform === 'darwin') return 'macos';
  const linuxDesktopEnv = getLinuxDesktopEnv(env);

  if (
    env.HYPRLAND_INSTANCE_SIGNATURE ||
    linuxDesktopEnv.xdgCurrentDesktop.includes('hyprland') ||
    linuxDesktopEnv.xdgSessionDesktop.includes('hyprland')
  ) {
    return 'hyprland';
  }
  if (linuxDesktopEnv.hasWayland && commandExists('hyprctl')) return 'hyprland';
  if (env.DISPLAY) return 'x11';
  fail('Could not detect display backend');
}

type LinuxDesktopEnv = {
  xdgCurrentDesktop: string;
  xdgSessionDesktop: string;
  hasWayland: boolean;
};

function getLinuxDesktopEnv(env: NodeJS.ProcessEnv): LinuxDesktopEnv {
  const xdgCurrentDesktop = (env.XDG_CURRENT_DESKTOP || '').toLowerCase();
  const xdgSessionDesktop = (env.XDG_SESSION_DESKTOP || '').toLowerCase();
  const xdgSessionType = (env.XDG_SESSION_TYPE || '').toLowerCase();
  return {
    xdgCurrentDesktop,
    xdgSessionDesktop,
    hasWayland: Boolean(env.WAYLAND_DISPLAY) || xdgSessionType === 'wayland',
  };
}

function shouldForceX11MpvBackend(args: Pick<Args, 'backend'>, env: NodeJS.ProcessEnv): boolean {
  if (process.platform !== 'linux' || !env.DISPLAY?.trim()) {
    return false;
  }

  const linuxDesktopEnv = getLinuxDesktopEnv(env);
  const supportedWaylandBackend =
    Boolean(env.HYPRLAND_INSTANCE_SIGNATURE || env.SWAYSOCK) ||
    linuxDesktopEnv.xdgCurrentDesktop.includes('hyprland') ||
    linuxDesktopEnv.xdgCurrentDesktop.includes('sway') ||
    linuxDesktopEnv.xdgSessionDesktop.includes('hyprland') ||
    linuxDesktopEnv.xdgSessionDesktop.includes('sway');
  return (
    args.backend === 'x11' ||
    (args.backend === 'auto' && linuxDesktopEnv.hasWayland && !supportedWaylandBackend)
  );
}

function resolveAppBinaryCandidate(candidate: string, pathModule: PathModule = path): string {
  const direct = resolveBinaryPathCandidate(candidate);
  if (!direct) return '';

  if (process.platform === 'win32') {
    try {
      if (fs.existsSync(direct) && fs.statSync(direct).isDirectory()) {
        for (const candidateBinary of ['SubMiner.exe', 'subminer.exe']) {
          const nestedCandidate = pathModule.join(direct, candidateBinary);
          if (isExecutable(nestedCandidate)) {
            return nestedCandidate;
          }
        }
        return '';
      }
    } catch {
      // ignore
    }

    if (isExecutable(direct)) {
      return direct;
    }

    if (!pathModule.extname(direct)) {
      for (const extension of ['.exe', '.cmd', '.bat']) {
        const withExtension = `${direct}${extension}`;
        if (isExecutable(withExtension)) {
          return withExtension;
        }
      }
    }

    return '';
  }

  if (isExecutable(direct)) {
    return direct;
  }

  if (process.platform !== 'darwin') {
    return '';
  }

  const appIndex = direct.indexOf('.app/');
  const appPath =
    direct.endsWith('.app') && direct.includes('.app')
      ? direct
      : appIndex >= 0
        ? direct.slice(0, appIndex + '.app'.length)
        : '';
  if (!appPath) return '';

  const candidates = [
    pathModule.join(appPath, 'Contents', 'MacOS', 'SubMiner'),
    pathModule.join(appPath, 'Contents', 'MacOS', 'subminer'),
  ];

  for (const candidateBinary of candidates) {
    if (isExecutable(candidateBinary)) {
      return candidateBinary;
    }
  }

  return '';
}

function findCommandOnPath(candidates: string[], pathModule: PathModule = path): string {
  const pathDirs = getPathEnv().split(pathModule.delimiter);
  for (const candidateName of candidates) {
    for (const dir of pathDirs) {
      if (!dir) continue;

      const directCandidate = pathModule.join(dir, candidateName);
      if (isExecutable(directCandidate)) {
        return directCandidate;
      }

      if (process.platform === 'win32' && !pathModule.extname(candidateName)) {
        for (const extension of ['.exe', '.cmd', '.bat']) {
          const extendedCandidate = pathModule.join(dir, `${candidateName}${extension}`);
          if (isExecutable(extendedCandidate)) {
            return extendedCandidate;
          }
        }
      }
    }
  }

  return '';
}

export function findAppBinary(selfPath: string, pathModule: PathModule = path): string | null {
  const envPaths = [process.env.SUBMINER_APPIMAGE_PATH, process.env.SUBMINER_BINARY_PATH].filter(
    (candidate): candidate is string => Boolean(candidate),
  );

  for (const envPath of envPaths) {
    const resolved = resolveAppBinaryCandidate(envPath, pathModule);
    if (resolved) {
      return resolved;
    }
  }

  const candidates: string[] = [];
  if (process.platform === 'win32') {
    const localAppData =
      process.env.LOCALAPPDATA?.trim() ||
      (process.env.APPDATA?.trim() || '').replace(/[\\/]Roaming$/i, `${pathModule.sep}Local`) ||
      pathModule.join(os.homedir(), 'AppData', 'Local');
    const programFiles = process.env.ProgramFiles?.trim() || 'C:\\Program Files';
    const programFilesX86 = process.env['ProgramFiles(x86)']?.trim() || 'C:\\Program Files (x86)';
    candidates.push(pathModule.join(localAppData, 'Programs', 'SubMiner', 'SubMiner.exe'));
    candidates.push(pathModule.join(programFiles, 'SubMiner', 'SubMiner.exe'));
    candidates.push(pathModule.join(programFilesX86, 'SubMiner', 'SubMiner.exe'));
    candidates.push('C:\\SubMiner\\SubMiner.exe');
  } else if (process.platform === 'darwin') {
    candidates.push('/Applications/SubMiner.app/Contents/MacOS/SubMiner');
    candidates.push('/Applications/SubMiner.app/Contents/MacOS/subminer');
    candidates.push(
      pathModule.join(os.homedir(), 'Applications/SubMiner.app/Contents/MacOS/SubMiner'),
    );
    candidates.push(
      pathModule.join(os.homedir(), 'Applications/SubMiner.app/Contents/MacOS/subminer'),
    );
  } else {
    candidates.push(pathModule.join(os.homedir(), '.local/bin/SubMiner.AppImage'));
    candidates.push('/opt/SubMiner/SubMiner.AppImage');
  }

  for (const candidate of candidates) {
    const resolved = resolveAppBinaryCandidate(candidate, pathModule);
    if (resolved) return resolved;
  }

  const fromPath = findCommandOnPath(
    process.platform === 'win32' ? ['SubMiner', 'subminer'] : ['subminer'],
    pathModule,
  );

  if (fromPath) {
    const resolvedSelf = realpathMaybe(selfPath);
    const resolvedCandidate = realpathMaybe(fromPath);
    if (resolvedSelf !== resolvedCandidate) return fromPath;
  }

  return null;
}

export function sendMpvCommand(socketPath: string, command: unknown[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const socket = net.createConnection(socketPath);
    socket.once('connect', () => {
      socket.write(`${JSON.stringify({ command })}\n`);
      socket.end();
      resolve();
    });
    socket.once('error', (error) => {
      reject(error);
    });
  });
}

interface MpvResponseEnvelope {
  request_id?: number;
  error?: string;
  data?: unknown;
}

export function sendMpvCommandWithResponse(
  socketPath: string,
  command: unknown[],
  timeoutMs = 5000,
): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const requestId = nowMs() + Math.floor(Math.random() * 1000);
    const socket = net.createConnection(socketPath);
    let buffer = '';

    const cleanup = (): void => {
      try {
        socket.destroy();
      } catch {
        // ignore
      }
    };

    const timer = setTimeout(() => {
      cleanup();
      reject(new Error(`MPV command timed out after ${timeoutMs}ms`));
    }, timeoutMs);

    const finish = (value: unknown): void => {
      clearTimeout(timer);
      cleanup();
      resolve(value);
    };

    socket.once('connect', () => {
      const message = JSON.stringify({ command, request_id: requestId });
      socket.write(`${message}\n`);
    });

    socket.on('data', (chunk: Buffer) => {
      buffer += chunk.toString();
      const lines = buffer.split(/\r?\n/);
      buffer = lines.pop() ?? '';
      for (const line of lines) {
        if (!line.trim()) continue;
        let parsed: MpvResponseEnvelope;
        try {
          parsed = JSON.parse(line);
        } catch {
          continue;
        }
        if (parsed.request_id !== requestId) continue;
        if (parsed.error && parsed.error !== 'success') {
          reject(new Error(`MPV error: ${parsed.error}`));
          cleanup();
          clearTimeout(timer);
          return;
        }
        finish(parsed.data);
        return;
      }
    });

    socket.once('error', (error) => {
      clearTimeout(timer);
      cleanup();
      reject(error);
    });
  });
}

export async function getMpvTracks(socketPath: string): Promise<MpvTrack[]> {
  const response = await sendMpvCommandWithResponse(
    socketPath,
    ['get_property', 'track-list'],
    8000,
  );
  if (!Array.isArray(response)) return [];

  return response
    .filter((track): track is MpvTrack => {
      if (!track || typeof track !== 'object') return false;
      const candidate = track as Record<string, unknown>;
      return candidate.type === 'sub';
    })
    .map((track) => {
      const candidate = track as Record<string, unknown>;
      return {
        type: typeof candidate.type === 'string' ? candidate.type : undefined,
        id:
          typeof candidate.id === 'number'
            ? candidate.id
            : typeof candidate.id === 'string'
              ? Number.parseInt(candidate.id, 10)
              : undefined,
        lang: typeof candidate.lang === 'string' ? candidate.lang : undefined,
        title: typeof candidate.title === 'string' ? candidate.title : undefined,
      };
    });
}

function isPreferredStreamLang(candidate: string, preferred: string[]): boolean {
  const normalized = normalizeLangCode(candidate);
  if (!normalized) return false;
  if (preferred.includes(normalized)) return true;
  if (normalized === 'ja' && preferred.includes('jpn')) return true;
  if (normalized === 'jpn' && preferred.includes('ja')) return true;
  if (normalized === 'en' && preferred.includes('eng')) return true;
  if (normalized === 'eng' && preferred.includes('en')) return true;
  return false;
}

export function findPreferredSubtitleTrack(
  tracks: MpvTrack[],
  preferredLanguages: string[],
): MpvTrack | null {
  const normalizedPreferred = uniqueNormalizedLangCodes(preferredLanguages);
  const subtitleTracks = tracks.filter((track) => track.type === 'sub');
  if (normalizedPreferred.length === 0) return subtitleTracks[0] ?? null;

  for (const lang of normalizedPreferred) {
    const matched = subtitleTracks.find(
      (track) => track.lang && isPreferredStreamLang(track.lang, [lang]),
    );
    if (matched) return matched;
  }

  return null;
}

export async function waitForSubtitleTrackList(
  socketPath: string,
  logLevel: LogLevel,
): Promise<MpvTrack[]> {
  const maxAttempts = 40;
  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    const tracks = await getMpvTracks(socketPath).catch(() => [] as MpvTrack[]);
    if (tracks.length > 0) return tracks;
    if (attempt % 10 === 0) {
      log('debug', logLevel, `Waiting for mpv tracks (${attempt}/${maxAttempts})`);
    }
    await sleep(250);
  }
  return [];
}

export async function loadSubtitleIntoMpv(
  socketPath: string,
  subtitlePath: string,
  select: boolean,
  logLevel: LogLevel,
): Promise<void> {
  for (let attempt = 1; ; attempt += 1) {
    const mpvExited =
      state.mpvProc !== null &&
      state.mpvProc.exitCode !== null &&
      state.mpvProc.exitCode !== undefined;
    if (mpvExited) {
      throw new Error(`mpv exited before subtitle could be loaded: ${subtitlePath}`);
    }

    if (!fs.existsSync(socketPath)) {
      if (attempt % 20 === 0) {
        log(
          'debug',
          logLevel,
          `Waiting for mpv socket before loading subtitle (${attempt} attempts): ${path.basename(subtitlePath)}`,
        );
      }
      await sleep(250);
      continue;
    }
    try {
      await sendMpvCommand(
        socketPath,
        select ? ['sub-add', subtitlePath, 'select'] : ['sub-add', subtitlePath],
      );
      log('info', logLevel, `Loaded generated subtitle into mpv: ${path.basename(subtitlePath)}`);
      return;
    } catch {
      if (attempt % 20 === 0) {
        log(
          'debug',
          logLevel,
          `Retrying subtitle load into mpv (${attempt} attempts): ${path.basename(subtitlePath)}`,
        );
      }
      await sleep(250);
    }
  }
}

export function shouldResolveAniSkipMetadata(
  target: string,
  targetKind: 'file' | 'url',
  preloadedSubtitles?: { primaryPath?: string; secondaryPath?: string },
): boolean {
  if (targetKind !== 'file') {
    return false;
  }
  if (preloadedSubtitles?.primaryPath || preloadedSubtitles?.secondaryPath) {
    return false;
  }
  return !isYoutubeTarget(target);
}

export async function startMpv(
  target: string,
  targetKind: 'file' | 'url',
  args: Args,
  socketPath: string,
  appPath: string,
  preloadedSubtitles?: { primaryPath?: string; secondaryPath?: string },
  options?: {
    startPaused?: boolean;
    disableYoutubeSubtitleAutoLoad?: boolean;
    runtimePluginPath?: string | null;
    runtimePluginConfig?: PluginRuntimeConfig;
  },
): Promise<void> {
  if (targetKind === 'file' && (!fs.existsSync(target) || !fs.statSync(target).isFile())) {
    fail(`Video file not found: ${target}`);
  }

  if (targetKind === 'url') {
    log('info', args.logLevel, `Playing URL: ${target}`);
  } else {
    log('info', args.logLevel, `Playing: ${path.basename(target)}`);
  }

  const mpvArgs: string[] = [];
  mpvArgs.push(...buildConfiguredMpvDefaultArgs(args));
  appendRuntimePluginLaunchArgs(
    mpvArgs,
    resolveLauncherRuntimePluginPlan({
      runtimePluginPath:
        options?.runtimePluginPath ?? resolveLauncherRuntimePluginPath({ appPath }),
    }),
    args.logLevel,
  );
  if (targetKind === 'url' && isYoutubeTarget(target)) {
    log('info', args.logLevel, 'Applying URL playback options');
    mpvArgs.push('--ytdl=yes');
    const subtitleLangs = uniqueNormalizedLangCodes([
      ...args.youtubePrimarySubLangs,
      ...args.youtubeSecondarySubLangs,
    ]).join(',');
    const audioLangs = uniqueNormalizedLangCodes(args.youtubeAudioLangs).join(',');
    log('info', args.logLevel, 'Applying YouTube playback options');
    log('debug', args.logLevel, `YouTube subtitle langs: ${subtitleLangs}`);
    log('debug', args.logLevel, `YouTube audio langs: ${audioLangs}`);
    mpvArgs.push(`--ytdl-format=${DEFAULT_YOUTUBE_YTDL_FORMAT}`, `--alang=${audioLangs}`);
    if (options?.disableYoutubeSubtitleAutoLoad !== true) {
      mpvArgs.push(
        '--sub-auto=fuzzy',
        `--slang=${subtitleLangs}`,
        '--ytdl-raw-options-append=write-subs=',
        '--ytdl-raw-options-append=sub-format=vtt/best',
        `--ytdl-raw-options-append=sub-langs=${subtitleLangs}`,
      );
    } else {
      mpvArgs.push('--sub-auto=no');
    }
  }
  if (args.mpvArgs) {
    mpvArgs.push(...parseMpvArgString(args.mpvArgs));
  }
  if (preloadedSubtitles?.primaryPath) {
    mpvArgs.push(`--sub-file=${preloadedSubtitles.primaryPath}`);
  }
  if (preloadedSubtitles?.secondaryPath) {
    mpvArgs.push(`--sub-file=${preloadedSubtitles.secondaryPath}`);
  }
  if (options?.startPaused) {
    mpvArgs.push('--pause=yes');
  }
  const aniSkipMetadata = shouldResolveAniSkipMetadata(target, targetKind, preloadedSubtitles)
    ? await resolveAniSkipMetadataForFile(target)
    : null;
  const extraScriptOpts =
    targetKind === 'url' &&
    isYoutubeTarget(target) &&
    options?.disableYoutubeSubtitleAutoLoad === true
      ? ['subminer-auto_start_pause_until_ready=no']
      : [];
  const runtimeScriptOpts = options?.runtimePluginConfig
    ? buildPluginRuntimeScriptOptParts(options.runtimePluginConfig, appPath)
    : [`subminer-binary_path=${appPath}`, `subminer-socket_path=${socketPath}`];
  const scriptOpts = buildSubminerScriptOpts(appPath, socketPath, aniSkipMetadata, args.logLevel, [
    ...runtimeScriptOpts,
    ...extraScriptOpts,
  ]);
  if (aniSkipMetadata) {
    log(
      'debug',
      args.logLevel,
      `AniSkip metadata (${aniSkipMetadata.source}): title="${aniSkipMetadata.title}" season=${aniSkipMetadata.season ?? '-'} episode=${aniSkipMetadata.episode ?? '-'}`,
    );
  }
  mpvArgs.push(`--script-opts=${scriptOpts}`);
  mpvArgs.push(`--log-file=${getMpvLogPath()}`);

  try {
    fs.rmSync(socketPath, { force: true });
  } catch {
    // ignore
  }

  mpvArgs.push(`--input-ipc-server=${socketPath}`);
  mpvArgs.push(target);

  const mpvTarget = resolveCommandInvocation('mpv', mpvArgs, {
    normalizeWindowsShellArgs: false,
  });
  state.mpvProc = spawn(mpvTarget.command, mpvTarget.args, {
    stdio: 'inherit',
    env: buildMpvEnv(args),
  });
}

async function waitForOverlayStartCommandSettled(
  proc: ReturnType<typeof spawn>,
  logLevel: LogLevel,
  timeoutMs: number,
): Promise<void> {
  await new Promise<void>((resolve) => {
    let settled = false;
    let timer: ReturnType<typeof setTimeout> | null = null;

    const finish = () => {
      if (settled) return;
      settled = true;
      if (timer) clearTimeout(timer);
      proc.off('exit', onExit);
      proc.off('error', onError);
      resolve();
    };

    const onExit = (code: number | null) => {
      if (typeof code === 'number' && code !== 0) {
        log('warn', logLevel, `Overlay start command exited with status ${code}`);
      }
      finish();
    };

    const onError = (error: Error) => {
      log('warn', logLevel, `Overlay start command failed: ${error.message}`);
      finish();
    };

    proc.once('exit', onExit);
    proc.once('error', onError);
    timer = setTimeout(finish, timeoutMs);

    if (proc.exitCode !== null && proc.exitCode !== undefined) {
      onExit(proc.exitCode);
    }
  });
}

export async function startOverlay(
  appPath: string,
  args: Args,
  socketPath: string,
  extraAppArgs: string[] = [],
  configDir: string = getLauncherConfigDir(),
): Promise<void> {
  const backend = detectBackend(args.backend);
  log('info', args.logLevel, `Starting SubMiner overlay (backend: ${backend})...`);
  const alreadyManagedByLauncher = state.overlayManagedByLauncher && state.appPath === appPath;

  const overlayArgs = [
    '--start',
    '--managed-playback',
    '--backend',
    backend,
    '--socket',
    socketPath,
    ...extraAppArgs,
  ];
  if (args.logLevel !== 'info') overlayArgs.push('--log-level', args.logLevel);
  if (args.useTexthooker) overlayArgs.push('--texthooker');

  const controlResult = await sendAppControlCommand(overlayArgs, {
    configDir,
  });
  if (controlResult.ok) {
    log('debug', args.logLevel, 'Attached to running SubMiner app via control socket');
    if (alreadyManagedByLauncher) {
      markOverlayManagedByLauncher(appPath);
    } else {
      clearOverlayManagedByLauncher();
      state.overlayProc = null;
    }

    const socketReady = await waitForUnixSocketReady(
      socketPath,
      OVERLAY_START_SOCKET_READY_TIMEOUT_MS,
    );
    if (!socketReady) {
      log(
        'debug',
        args.logLevel,
        'Overlay start continuing before mpv socket readiness was confirmed',
      );
    }
    return;
  }
  if (controlResult.unavailable !== true) {
    log(
      'warn',
      args.logLevel,
      `Running SubMiner app control command failed: ${controlResult.error ?? 'unknown error'}`,
    );
    if (!alreadyManagedByLauncher) {
      clearOverlayManagedByLauncher();
      state.overlayProc = null;
    }
  }

  const appAlreadyRunning = isAppAlreadyRunning(appPath, args.logLevel);
  const borrowingExistingApp = appAlreadyRunning && !alreadyManagedByLauncher;
  const spawnOverlayArgs = [...overlayArgs];
  if (!borrowingExistingApp) spawnOverlayArgs.unshift('--background');

  const target = resolveAppSpawnTarget(appPath, spawnOverlayArgs);
  state.overlayProc = spawn(target.command, target.args, {
    stdio: ['ignore', 'pipe', 'pipe'],
    env: buildAppEnv(process.env, target.env),
  });
  attachAppProcessLogging(state.overlayProc);
  if (borrowingExistingApp) {
    log(
      'debug',
      args.logLevel,
      'SubMiner app is already running; launcher will not stop it after playback',
    );
    clearOverlayManagedByLauncher();
  } else {
    markOverlayManagedByLauncher(appPath);
  }

  const [socketReady] = await Promise.all([
    waitForUnixSocketReady(socketPath, OVERLAY_START_SOCKET_READY_TIMEOUT_MS),
    waitForOverlayStartCommandSettled(
      state.overlayProc,
      args.logLevel,
      OVERLAY_START_COMMAND_SETTLE_TIMEOUT_MS,
    ),
  ]);

  if (!socketReady) {
    log(
      'debug',
      args.logLevel,
      'Overlay start continuing before mpv socket readiness was confirmed',
    );
  }
}

function getLauncherConfigDir(): string {
  return getDefaultConfigDir({
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
  });
}

export async function isRunningAppControlServerAvailable(
  logLevel: LogLevel,
  configDir: string = getLauncherConfigDir(),
): Promise<boolean> {
  const available = await checkAppControlServerAvailable({
    configDir,
  });
  if (available) {
    log('debug', logLevel, 'Running SubMiner app control socket detected');
  }
  return available;
}

export function markOverlayManagedByLauncher(appPath?: string): void {
  if (appPath) {
    state.appPath = appPath;
  }
  state.overlayManagedByLauncher = true;
}

function clearOverlayManagedByLauncher(): void {
  state.appPath = '';
  state.overlayManagedByLauncher = false;
}

function isAppAlreadyRunning(appPath: string, logLevel: LogLevel): boolean {
  const result = runSyncAppCommand(appPath, ['--app-ping'], false);
  if (result.error) {
    log('debug', logLevel, `App ping failed before overlay start: ${result.error.message}`);
    return false;
  }
  return result.status === 0;
}

export function openUrlInDefaultBrowser(url: string, logLevel: LogLevel): void {
  const target =
    process.platform === 'darwin'
      ? { command: 'open', args: [url] }
      : process.platform === 'win32'
        ? { command: 'cmd', args: ['/c', 'start', '', url] }
        : { command: 'xdg-open', args: [url] };
  const result = spawnSync(target.command, target.args, {
    stdio: 'ignore',
    env: process.env,
    windowsHide: true,
  });
  if (result.error) {
    log('warn', logLevel, `Failed to open browser for ${url}: ${result.error.message}`);
  }
}

export function launchTexthookerOnly(
  appPath: string,
  args: Args,
  deps: { openBrowser?: (url: string) => void } = {},
): never {
  const overlayArgs = ['--texthooker'];
  if (args.texthookerOpenBrowser) overlayArgs.push('--open-browser');
  if (args.logLevel !== 'info') overlayArgs.push('--log-level', args.logLevel);

  log('info', args.logLevel, 'Launching texthooker mode...');
  const result = runSyncAppCommand(appPath, overlayArgs, true);
  if (result.error) {
    fail(`Failed to launch texthooker mode: ${result.error.message}`);
  }
  if (args.texthookerOpenBrowser && (result.status ?? 0) === 0) {
    const url = 'http://127.0.0.1:5174';
    const openBrowser =
      deps.openBrowser ??
      ((browserUrl: string) => openUrlInDefaultBrowser(browserUrl, args.logLevel));
    openBrowser(url);
  }
  process.exit(result.status ?? 0);
}

export function stopOverlay(args: Args): void {
  if (state.stopRequested) return;
  state.stopRequested = true;

  stopManagedOverlayApp(args);

  if (state.mpvProc && !state.mpvProc.killed) {
    try {
      state.mpvProc.kill('SIGTERM');
    } catch {
      // ignore
    }
  }

  for (const child of state.youtubeSubgenChildren) {
    if (!child.killed) {
      try {
        child.kill('SIGTERM');
      } catch {
        // ignore
      }
    }
  }
  state.youtubeSubgenChildren.clear();

  void terminateTrackedDetachedMpv(args.logLevel);
}

export async function cleanupPlaybackSession(args: Args): Promise<void> {
  stopManagedOverlayApp(args);

  if (state.mpvProc && !state.mpvProc.killed) {
    try {
      state.mpvProc.kill('SIGTERM');
    } catch {
      // ignore
    }
  }

  for (const child of state.youtubeSubgenChildren) {
    if (!child.killed) {
      try {
        child.kill('SIGTERM');
      } catch {
        // ignore
      }
    }
  }
  state.youtubeSubgenChildren.clear();

  await terminateTrackedDetachedMpv(args.logLevel);
}

function stopManagedOverlayApp(args: Args): void {
  if (!(state.overlayManagedByLauncher && state.appPath)) {
    return;
  }

  log('info', args.logLevel, 'Stopping SubMiner overlay...');

  const stopArgs = ['--stop'];
  if (args.logLevel !== 'info') stopArgs.push('--log-level', args.logLevel);

  const target = resolveAppSpawnTarget(state.appPath, stopArgs);
  const result = spawnSync(target.command, target.args, {
    stdio: 'ignore',
    env: buildAppEnv(process.env, target.env),
  });
  if (result.error) {
    log('warn', args.logLevel, `Failed to stop SubMiner overlay: ${result.error.message}`);
  } else if (typeof result.status === 'number' && result.status !== 0) {
    log('warn', args.logLevel, `SubMiner overlay stop command exited with status ${result.status}`);
  }

  if (state.overlayProc && !state.overlayProc.killed) {
    try {
      state.overlayProc.kill('SIGTERM');
    } catch {
      // ignore
    }
  }
}

function clearTransportedAppArgs(env: Record<string, string | undefined>): void {
  for (const key of Object.keys(env)) {
    if (key === TRANSPORTED_APP_ARGC_ENV || /^SUBMINER_APP_ARG_\d+$/.test(key)) {
      delete env[key];
    }
  }
}

function buildTransportedAppArgsEnv(appArgs: string[]): NodeJS.ProcessEnv {
  const env: NodeJS.ProcessEnv = {
    [TRANSPORTED_APP_ARGC_ENV]: String(appArgs.length),
  };
  appArgs.forEach((arg, index) => {
    env[`${TRANSPORTED_APP_ARG_PREFIX}${index}`] = arg;
  });
  return env;
}

function shouldTransportAppArgsForAppImage(appPath: string): boolean {
  return process.platform === 'linux' && /\.AppImage$/i.test(appPath);
}

function buildAppEnv(
  baseEnv: NodeJS.ProcessEnv = process.env,
  extraEnv: NodeJS.ProcessEnv = {},
): NodeJS.ProcessEnv {
  const env: Record<string, string | undefined> = {
    ...baseEnv,
    SUBMINER_APP_LOG: getAppLogPath(),
    SUBMINER_MPV_LOG: getMpvLogPath(),
  };
  delete env.ELECTRON_RUN_AS_NODE;
  clearTransportedAppArgs(env);
  Object.assign(env, extraEnv);
  const layers = env.VK_INSTANCE_LAYERS;
  if (typeof layers === 'string' && layers.trim().length > 0) {
    const filtered = layers
      .split(':')
      .map((part) => part.trim())
      .filter((part) => part.length > 0 && !/lsfg/i.test(part));
    if (filtered.length > 0) {
      env.VK_INSTANCE_LAYERS = filtered.join(':');
    } else {
      delete env.VK_INSTANCE_LAYERS;
    }
  }
  return env;
}

export function buildMpvEnv(
  args: Pick<Args, 'backend'>,
  baseEnv: NodeJS.ProcessEnv = process.env,
): NodeJS.ProcessEnv {
  const env = buildAppEnv(baseEnv);
  if (!shouldForceX11MpvBackend(args, env)) {
    return env;
  }

  delete env.WAYLAND_DISPLAY;
  delete env.HYPRLAND_INSTANCE_SIGNATURE;
  delete env.SWAYSOCK;
  env.XDG_SESSION_TYPE = 'x11';
  return env;
}

export function buildMpvBackendArgs(
  args: Pick<Args, 'backend'>,
  baseEnv: NodeJS.ProcessEnv = process.env,
): string[] {
  if (!shouldForceX11MpvBackend(args, baseEnv)) {
    return [];
  }
  return ['--vo=gpu', '--gpu-api=opengl', '--gpu-context=x11egl,x11'];
}

export function buildConfiguredMpvDefaultArgs(
  args: Pick<Args, 'profile' | 'backend' | 'launchMode'>,
  baseEnv: NodeJS.ProcessEnv = process.env,
): string[] {
  const mpvArgs: string[] = [];
  if (args.profile) mpvArgs.push(`--profile=${args.profile}`);
  mpvArgs.push(...DEFAULT_MPV_SUBMINER_ARGS);
  if (process.platform === 'darwin') {
    // macOS menu accelerators do not reach mpv script bindings unless disabled.
    mpvArgs.push('--macos-menu-shortcuts=no');
  }
  mpvArgs.push(...buildMpvBackendArgs(args, baseEnv));
  mpvArgs.push(...buildMpvLaunchModeArgs(args.launchMode));
  return mpvArgs;
}

function appendCapturedAppOutput(kind: 'STDOUT' | 'STDERR', chunk: string): void {
  const normalized = chunk.replace(/\r\n/g, '\n');
  for (const line of normalized.split('\n')) {
    if (!line) continue;
    appendToAppLog(`[${kind}] ${line}`);
  }
}

const KNOWN_ELECTRON_MENU_DIAGNOSTIC =
  'representedObject is not a WeakPtrToElectronMenuModelAsNSObject';

function filterKnownElectronDiagnostics(chunk: string): string {
  const lines = chunk.match(/[^\n]*\n|[^\n]+/g) ?? [];
  return lines.filter((line) => !line.includes(KNOWN_ELECTRON_MENU_DIAGNOSTIC)).join('');
}

function attachAppProcessLogging(
  proc: ReturnType<typeof spawn>,
  options?: {
    mirrorStdout?: boolean;
    mirrorStderr?: boolean;
  },
): void {
  proc.stdout?.setEncoding('utf8');
  proc.stderr?.setEncoding('utf8');
  proc.stdout?.on('data', (chunk: string) => {
    appendCapturedAppOutput('STDOUT', chunk);
    if (options?.mirrorStdout) process.stdout.write(chunk);
  });
  proc.stderr?.on('data', (chunk: string) => {
    const filteredChunk = filterKnownElectronDiagnostics(chunk);
    if (!filteredChunk) {
      return;
    }
    appendCapturedAppOutput('STDERR', filteredChunk);
    if (options?.mirrorStderr) process.stderr.write(filteredChunk);
  });
}

function runSyncAppCommand(
  appPath: string,
  appArgs: string[],
  mirrorOutput: boolean,
): {
  status: number;
  stdout: string;
  stderr: string;
  error?: Error;
} {
  const target = resolveAppSpawnTarget(appPath, appArgs);
  const result = spawnSync(target.command, target.args, {
    env: buildAppEnv(process.env, target.env),
    encoding: 'utf8',
  });
  if (result.stdout) {
    appendCapturedAppOutput('STDOUT', result.stdout);
    if (mirrorOutput) process.stdout.write(result.stdout);
  }
  if (result.stderr) {
    const filteredStderr = filterKnownElectronDiagnostics(result.stderr);
    if (filteredStderr) {
      appendCapturedAppOutput('STDERR', filteredStderr);
      if (mirrorOutput) process.stderr.write(filteredStderr);
    }
  }
  return {
    status: result.status ?? 1,
    stdout: result.stdout ?? '',
    stderr: result.stderr ? filterKnownElectronDiagnostics(result.stderr) : '',
    error: result.error ?? undefined,
  };
}

function maybeCaptureAppArgs(appArgs: string[]): boolean {
  const capturePath = process.env.SUBMINER_TEST_CAPTURE?.trim();
  if (!capturePath) {
    return false;
  }

  fs.writeFileSync(capturePath, `${appArgs.join('\n')}${appArgs.length > 0 ? '\n' : ''}`, 'utf8');
  return true;
}

function resolveAppSpawnTarget(appPath: string, appArgs: string[]): SpawnTarget {
  if (shouldTransportAppArgsForAppImage(appPath)) {
    return {
      command: appPath,
      args: [],
      env: buildTransportedAppArgsEnv(appArgs),
    };
  }
  if (process.platform !== 'win32') {
    return { command: appPath, args: appArgs };
  }
  return resolveCommandInvocation(appPath, appArgs);
}

export function runAppCommandWithInherit(appPath: string, appArgs: string[]): void {
  if (maybeCaptureAppArgs(appArgs)) {
    process.exit(0);
  }

  const target = resolveAppSpawnTarget(appPath, appArgs);
  const proc = spawn(target.command, target.args, {
    stdio: ['ignore', 'pipe', 'pipe'],
    env: buildAppEnv(process.env, target.env),
  });
  attachAppProcessLogging(proc, { mirrorStdout: true, mirrorStderr: true });
  proc.once('error', (error) => {
    fail(`Failed to run app command: ${error.message}`);
  });
  proc.once('close', (code) => {
    process.exit(code ?? 0);
  });
}

export function runAppCommandSilently(appPath: string, appArgs: string[]): void {
  if (maybeCaptureAppArgs(appArgs)) {
    process.exit(0);
  }

  const target = resolveAppSpawnTarget(appPath, appArgs);
  const proc = spawn(target.command, target.args, {
    stdio: ['ignore', 'pipe', 'pipe'],
    env: buildAppEnv(process.env, target.env),
  });
  attachAppProcessLogging(proc);
  proc.once('error', (error) => {
    fail(`Failed to run app command: ${error.message}`);
  });
  proc.once('close', (code) => {
    process.exit(code ?? 0);
  });
}

export function runAppCommandCaptureOutput(
  appPath: string,
  appArgs: string[],
): {
  status: number;
  stdout: string;
  stderr: string;
  error?: Error;
} {
  if (maybeCaptureAppArgs(appArgs)) {
    return {
      status: 0,
      stdout: '',
      stderr: '',
    };
  }

  return runSyncAppCommand(appPath, appArgs, false);
}

export function runAppCommandAttached(
  appPath: string,
  appArgs: string[],
  logLevel: LogLevel,
  label: string,
): Promise<number> {
  if (maybeCaptureAppArgs(appArgs)) {
    return Promise.resolve(0);
  }

  const target = resolveAppSpawnTarget(appPath, appArgs);
  log(
    'debug',
    logLevel,
    `${label}: launching attached app with args: ${[target.command, ...target.args].join(' ')}`,
  );

  return new Promise((resolve, reject) => {
    const proc = spawn(target.command, target.args, {
      stdio: ['ignore', 'pipe', 'pipe'],
      env: buildAppEnv(process.env, target.env),
    });
    attachAppProcessLogging(proc, { mirrorStdout: true, mirrorStderr: true });
    proc.once('error', (error) => {
      reject(error);
    });
    proc.once('close', (code, signal) => {
      if (code !== null) {
        resolve(code);
      } else if (signal) {
        resolve(128);
      } else {
        resolve(0);
      }
    });
  });
}

export function runAppCommandWithInheritLogged(
  appPath: string,
  appArgs: string[],
  logLevel: LogLevel,
  label: string,
): never {
  if (maybeCaptureAppArgs(appArgs)) {
    process.exit(0);
  }

  const target = resolveAppSpawnTarget(appPath, appArgs);
  log(
    'debug',
    logLevel,
    `${label}: launching app with args: ${[target.command, ...target.args].join(' ')}`,
  );
  const result = runSyncAppCommand(appPath, appArgs, true);
  if (result.error) {
    fail(`Failed to run app command: ${result.error.message}`);
  }
  log('debug', logLevel, `${label}: app command exited with status ${result.status ?? 0}`);
  process.exit(result.status ?? 0);
}

export function launchAppStartDetached(appPath: string, logLevel: LogLevel): void {
  const startArgs = ['--start'];
  if (logLevel !== 'info') startArgs.push('--log-level', logLevel);
  launchAppCommandDetached(appPath, startArgs, logLevel, 'start');
}

export function launchAppCommandDetached(
  appPath: string,
  appArgs: string[],
  logLevel: LogLevel,
  label: string,
): void {
  if (maybeCaptureAppArgs(appArgs)) {
    return;
  }
  const target = resolveAppSpawnTarget(appPath, appArgs);
  log(
    'debug',
    logLevel,
    `${label}: launching detached app with args: ${[target.command, ...target.args].join(' ')}`,
  );
  const appLogPath = getAppLogPath();
  fs.mkdirSync(path.dirname(appLogPath), { recursive: true });
  const stdoutFd = fs.openSync(appLogPath, 'a');
  const stderrFd = fs.openSync(appLogPath, 'a');
  try {
    const proc = spawn(target.command, target.args, {
      stdio: ['ignore', stdoutFd, stderrFd],
      detached: true,
      env: buildAppEnv(process.env, target.env),
    });
    proc.once('error', (error) => {
      log('warn', logLevel, `${label}: failed to launch detached app: ${error.message}`);
    });
    proc.unref();
  } finally {
    fs.closeSync(stdoutFd);
    fs.closeSync(stderrFd);
  }
}

export function launchMpvIdleDetached(
  socketPath: string,
  appPath: string,
  args: Args,
  runtimePluginPath?: string | null,
  runtimePluginConfig?: PluginRuntimeConfig,
): Promise<void> {
  return (async () => {
    await terminateTrackedDetachedMpv(args.logLevel);
    try {
      fs.rmSync(socketPath, { force: true });
    } catch {
      // ignore
    }

    const mpvArgs: string[] = buildConfiguredMpvDefaultArgs(args);
    appendRuntimePluginLaunchArgs(
      mpvArgs,
      resolveLauncherRuntimePluginPlan({
        runtimePluginPath: runtimePluginPath ?? resolveLauncherRuntimePluginPath({ appPath }),
      }),
      args.logLevel,
    );
    if (args.mpvArgs) {
      mpvArgs.push(...parseMpvArgString(args.mpvArgs));
    }
    mpvArgs.push('--idle=yes');
    const runtimeScriptOpts = runtimePluginConfig
      ? buildPluginRuntimeScriptOptParts(runtimePluginConfig, appPath)
      : [`subminer-binary_path=${appPath}`, `subminer-socket_path=${socketPath}`];
    mpvArgs.push(
      `--script-opts=${buildSubminerScriptOpts(
        appPath,
        socketPath,
        null,
        args.logLevel,
        runtimeScriptOpts,
      )}`,
    );
    mpvArgs.push(`--log-file=${getMpvLogPath()}`);
    mpvArgs.push(`--input-ipc-server=${socketPath}`);
    const mpvTarget = resolveCommandInvocation('mpv', mpvArgs, {
      normalizeWindowsShellArgs: false,
    });
    const proc = spawn(mpvTarget.command, mpvTarget.args, {
      stdio: 'ignore',
      detached: true,
      env: buildMpvEnv(args),
    });
    if (typeof proc.pid === 'number' && proc.pid > 0) {
      trackDetachedMpvPid(proc.pid);
    }
    proc.unref();
  })();
}

async function sleepMs(ms: number): Promise<void> {
  await new Promise<void>((resolve) => setTimeout(resolve, ms));
}

async function canConnectUnixSocket(socketPath: string): Promise<boolean> {
  return await new Promise<boolean>((resolve) => {
    const socket = net.createConnection(socketPath);
    let settled = false;

    const finish = (value: boolean) => {
      if (settled) return;
      settled = true;
      try {
        socket.destroy();
      } catch {
        // ignore
      }
      resolve(value);
    };

    socket.once('connect', () => finish(true));
    socket.once('error', () => finish(false));
    socket.setTimeout(400, () => finish(false));
  });
}

export async function waitForUnixSocketReady(
  socketPath: string,
  timeoutMs: number,
): Promise<boolean> {
  const deadline = nowMs() + timeoutMs;
  while (nowMs() < deadline) {
    try {
      if (process.platform === 'win32' || fs.existsSync(socketPath)) {
        const ready = await canConnectUnixSocket(socketPath);
        if (ready) return true;
      }
    } catch {
      // ignore transient fs errors
    }
    await sleepMs(150);
  }
  return false;
}
