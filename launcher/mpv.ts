import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import net from 'node:net';
import { spawn, spawnSync } from 'node:child_process';
import type { LogLevel, Backend, Args, MpvTrack } from './types.js';
import { DEFAULT_MPV_SUBMINER_ARGS, DEFAULT_YOUTUBE_YTDL_FORMAT } from './types.js';
import { log, fail, getMpvLogPath } from './log.js';
import { buildSubminerScriptOpts, resolveAniSkipMetadataForFile } from './aniskip-metadata.js';
import {
  commandExists,
  getPathEnv,
  isExecutable,
  resolveBinaryPathCandidate,
  resolveCommandInvocation,
  realpathMaybe,
  isYoutubeTarget,
  uniqueNormalizedLangCodes,
  sleep,
  normalizeLangCode,
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
};

const DETACHED_IDLE_MPV_PID_FILE = path.join(os.tmpdir(), 'subminer-idle-mpv.pid');
const OVERLAY_START_SOCKET_READY_TIMEOUT_MS = 900;
const OVERLAY_START_COMMAND_SETTLE_TIMEOUT_MS = 700;

export function parseMpvArgString(input: string): string[] {
  const chars = input;
  const args: string[] = [];
  let current = '';
  let tokenStarted = false;
  let inSingleQuote = false;
  let inDoubleQuote = false;
  let escaping = false;
  const canEscape = (nextChar: string | undefined): boolean =>
    nextChar === undefined || nextChar === '"' || nextChar === "'" || nextChar === '\\' || /\s/.test(nextChar);

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

  const deadline = Date.now() + 1500;
  while (Date.now() < deadline) {
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

export function detectBackend(backend: Backend): Exclude<Backend, 'auto'> {
  if (backend !== 'auto') return backend;
  if (process.platform === 'darwin') return 'macos';
  const xdgCurrentDesktop = (process.env.XDG_CURRENT_DESKTOP || '').toLowerCase();
  const xdgSessionDesktop = (process.env.XDG_SESSION_DESKTOP || '').toLowerCase();
  const xdgSessionType = (process.env.XDG_SESSION_TYPE || '').toLowerCase();
  const hasWayland = Boolean(process.env.WAYLAND_DISPLAY) || xdgSessionType === 'wayland';

  if (
    process.env.HYPRLAND_INSTANCE_SIGNATURE ||
    xdgCurrentDesktop.includes('hyprland') ||
    xdgSessionDesktop.includes('hyprland')
  ) {
    return 'hyprland';
  }
  if (hasWayland && commandExists('hyprctl')) return 'hyprland';
  if (process.env.DISPLAY) return 'x11';
  fail('Could not detect display backend');
}

function resolveMacAppBinaryCandidate(candidate: string): string {
  const direct = resolveBinaryPathCandidate(candidate);
  if (!direct) return '';

  if (process.platform !== 'darwin') {
    return isExecutable(direct) ? direct : '';
  }

  if (isExecutable(direct)) {
    return direct;
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
    path.join(appPath, 'Contents', 'MacOS', 'SubMiner'),
    path.join(appPath, 'Contents', 'MacOS', 'subminer'),
  ];

  for (const candidateBinary of candidates) {
    if (isExecutable(candidateBinary)) {
      return candidateBinary;
    }
  }

  return '';
}

export function findAppBinary(selfPath: string): string | null {
  const envPaths = [process.env.SUBMINER_APPIMAGE_PATH, process.env.SUBMINER_BINARY_PATH].filter(
    (candidate): candidate is string => Boolean(candidate),
  );

  for (const envPath of envPaths) {
    const resolved = resolveMacAppBinaryCandidate(envPath);
    if (resolved) {
      return resolved;
    }
  }

  const candidates: string[] = [];
  if (process.platform === 'darwin') {
    candidates.push('/Applications/SubMiner.app/Contents/MacOS/SubMiner');
    candidates.push('/Applications/SubMiner.app/Contents/MacOS/subminer');
    candidates.push(path.join(os.homedir(), 'Applications/SubMiner.app/Contents/MacOS/SubMiner'));
    candidates.push(path.join(os.homedir(), 'Applications/SubMiner.app/Contents/MacOS/subminer'));
  }

  candidates.push(path.join(os.homedir(), '.local/bin/SubMiner.AppImage'));
  candidates.push('/opt/SubMiner/SubMiner.AppImage');

  for (const candidate of candidates) {
    if (isExecutable(candidate)) return candidate;
  }

  const fromPath = getPathEnv()
    .split(path.delimiter)
    .map((dir) => path.join(dir, 'subminer'))
    .find((candidate) => isExecutable(candidate));

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
    const requestId = Date.now() + Math.floor(Math.random() * 1000);
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
  options?: { startPaused?: boolean },
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
  if (args.profile) mpvArgs.push(`--profile=${args.profile}`);
  mpvArgs.push(...DEFAULT_MPV_SUBMINER_ARGS);
  if (args.mpvArgs) {
    mpvArgs.push(...parseMpvArgString(args.mpvArgs));
  }

  if (targetKind === 'url' && isYoutubeTarget(target)) {
    log('info', args.logLevel, 'Applying URL playback options');
    mpvArgs.push('--ytdl=yes', '--ytdl-raw-options=');

    if (isYoutubeTarget(target)) {
      const subtitleLangs = uniqueNormalizedLangCodes([
        ...args.youtubePrimarySubLangs,
        ...args.youtubeSecondarySubLangs,
      ]).join(',');
      const audioLangs = uniqueNormalizedLangCodes(args.youtubeAudioLangs).join(',');
      log('info', args.logLevel, 'Applying YouTube playback options');
      log('debug', args.logLevel, `YouTube subtitle langs: ${subtitleLangs}`);
      log('debug', args.logLevel, `YouTube audio langs: ${audioLangs}`);
      mpvArgs.push(`--ytdl-format=${DEFAULT_YOUTUBE_YTDL_FORMAT}`, `--alang=${audioLangs}`);
      mpvArgs.push(
        '--sub-auto=fuzzy',
        `--slang=${subtitleLangs}`,
        '--ytdl-raw-options-append=write-subs=',
        '--ytdl-raw-options-append=sub-format=vtt/best',
        `--ytdl-raw-options-append=sub-langs=${subtitleLangs}`,
      );
    }
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
  const scriptOpts = buildSubminerScriptOpts(appPath, socketPath, aniSkipMetadata, args.logLevel);
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

  const mpvTarget = resolveCommandInvocation('mpv', mpvArgs);
  state.mpvProc = spawn(mpvTarget.command, mpvTarget.args, { stdio: 'inherit' });
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

export async function startOverlay(appPath: string, args: Args, socketPath: string): Promise<void> {
  const backend = detectBackend(args.backend);
  log('info', args.logLevel, `Starting SubMiner overlay (backend: ${backend})...`);

  const overlayArgs = ['--start', '--backend', backend, '--socket', socketPath];
  if (args.logLevel !== 'info') overlayArgs.push('--log-level', args.logLevel);
  if (args.useTexthooker) overlayArgs.push('--texthooker');

  const target = resolveAppSpawnTarget(appPath, overlayArgs);
  state.overlayProc = spawn(target.command, target.args, {
    stdio: 'inherit',
    env: buildAppEnv(),
  });
  state.overlayManagedByLauncher = true;

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

export function launchTexthookerOnly(appPath: string, args: Args): never {
  const overlayArgs = ['--texthooker'];
  if (args.logLevel !== 'info') overlayArgs.push('--log-level', args.logLevel);

  log('info', args.logLevel, 'Launching texthooker mode...');
  const result = spawnSync(appPath, overlayArgs, {
    stdio: 'inherit',
    env: buildAppEnv(),
  });
  if (result.error) {
    fail(`Failed to launch texthooker mode: ${result.error.message}`);
  }
  process.exit(result.status ?? 0);
}

export function stopOverlay(args: Args): void {
  if (state.stopRequested) return;
  state.stopRequested = true;

  if (state.overlayManagedByLauncher && state.appPath) {
    log('info', args.logLevel, 'Stopping SubMiner overlay...');

    const stopArgs = ['--stop'];
    if (args.logLevel !== 'info') stopArgs.push('--log-level', args.logLevel);

    const result = spawnSync(state.appPath, stopArgs, {
      stdio: 'ignore',
      env: buildAppEnv(),
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

function buildAppEnv(): NodeJS.ProcessEnv {
  const env: Record<string, string | undefined> = {
    ...process.env,
    SUBMINER_MPV_LOG: getMpvLogPath(),
  };
  delete env.ELECTRON_RUN_AS_NODE;
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

function maybeCaptureAppArgs(appArgs: string[]): boolean {
  const capturePath = process.env.SUBMINER_TEST_CAPTURE?.trim();
  if (!capturePath) {
    return false;
  }

  fs.writeFileSync(capturePath, `${appArgs.join('\n')}${appArgs.length > 0 ? '\n' : ''}`, 'utf8');
  return true;
}

function resolveAppSpawnTarget(appPath: string, appArgs: string[]): SpawnTarget {
  if (process.platform !== 'win32') {
    return { command: appPath, args: appArgs };
  }
  return resolveCommandInvocation(appPath, appArgs);
}

export function runAppCommandWithInherit(appPath: string, appArgs: string[]): never {
  if (maybeCaptureAppArgs(appArgs)) {
    process.exit(0);
  }

  const target = resolveAppSpawnTarget(appPath, appArgs);
  const result = spawnSync(target.command, target.args, {
    stdio: 'inherit',
    env: buildAppEnv(),
  });
  if (result.error) {
    fail(`Failed to run app command: ${result.error.message}`);
  }
  process.exit(result.status ?? 0);
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

  const target = resolveAppSpawnTarget(appPath, appArgs);
  const result = spawnSync(target.command, target.args, {
    env: buildAppEnv(),
    encoding: 'utf8',
  });

  return {
    status: result.status ?? 1,
    stdout: result.stdout ?? '',
    stderr: result.stderr ?? '',
    error: result.error ?? undefined,
  };
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
      stdio: 'inherit',
      env: buildAppEnv(),
    });
    proc.once('error', (error) => {
      reject(error);
    });
    proc.once('exit', (code, signal) => {
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
  const result = spawnSync(target.command, target.args, {
    stdio: 'inherit',
    env: buildAppEnv(),
  });
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
  const proc = spawn(target.command, target.args, {
    stdio: 'ignore',
    detached: true,
    env: buildAppEnv(),
  });
  proc.once('error', (error) => {
    log('warn', logLevel, `${label}: failed to launch detached app: ${error.message}`);
  });
  proc.unref();
}

export function launchMpvIdleDetached(
  socketPath: string,
  appPath: string,
  args: Args,
): Promise<void> {
  return (async () => {
    await terminateTrackedDetachedMpv(args.logLevel);
    try {
      fs.rmSync(socketPath, { force: true });
    } catch {
      // ignore
    }

    const mpvArgs: string[] = [];
    if (args.profile) mpvArgs.push(`--profile=${args.profile}`);
    mpvArgs.push(...DEFAULT_MPV_SUBMINER_ARGS);
    if (args.mpvArgs) {
      mpvArgs.push(...parseMpvArgString(args.mpvArgs));
    }
    mpvArgs.push('--idle=yes');
    mpvArgs.push(`--script-opts=${buildSubminerScriptOpts(appPath, socketPath, null, args.logLevel)}`);
    mpvArgs.push(`--log-file=${getMpvLogPath()}`);
    mpvArgs.push(`--input-ipc-server=${socketPath}`);
    const mpvTarget = resolveCommandInvocation('mpv', mpvArgs);
    const proc = spawn(mpvTarget.command, mpvTarget.args, {
      stdio: 'ignore',
      detached: true,
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
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      if (fs.existsSync(socketPath)) {
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
