import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { resolveConfigFilePath } from '../src/config/path-resolution.js';
import type { Args } from './types.js';
import { log, fail } from './log.js';
import { commandExists, isYoutubeTarget, resolvePathMaybe, realpathMaybe } from './util.js';
import {
  parseArgs,
  loadLauncherYoutubeSubgenConfig,
  loadLauncherJellyfinConfig,
  readPluginRuntimeConfig,
} from './config.js';
import { showRofiMenu, showFzfMenu, collectVideos } from './picker.js';
import {
  state,
  startMpv,
  startOverlay,
  stopOverlay,
  launchTexthookerOnly,
  findAppBinary,
  loadSubtitleIntoMpv,
  runAppCommandWithInherit,
  launchMpvIdleDetached,
  waitForUnixSocketReady,
} from './mpv.js';
import { generateYoutubeSubtitles } from './youtube.js';
import { runJellyfinPlayMenu } from './jellyfin.js';

function checkDependencies(args: Args): void {
  const missing: string[] = [];

  if (!commandExists('mpv')) missing.push('mpv');

  if (args.targetKind === 'url' && isYoutubeTarget(args.target) && !commandExists('yt-dlp')) {
    missing.push('yt-dlp');
  }

  if (
    args.targetKind === 'url' &&
    isYoutubeTarget(args.target) &&
    args.youtubeSubgenMode !== 'off' &&
    !commandExists('ffmpeg')
  ) {
    missing.push('ffmpeg');
  }

  if (missing.length > 0) fail(`Missing dependencies: ${missing.join(' ')}`);
}

function checkPickerDependencies(args: Args): void {
  if (args.useRofi) {
    if (!commandExists('rofi')) fail('Missing dependency: rofi');
    return;
  }

  if (!commandExists('fzf')) fail('Missing dependency: fzf');
}

async function chooseTarget(
  args: Args,
  scriptPath: string,
): Promise<{ target: string; kind: 'file' | 'url' } | null> {
  if (args.target) {
    return { target: args.target, kind: args.targetKind as 'file' | 'url' };
  }

  const searchDir = realpathMaybe(resolvePathMaybe(args.directory));
  if (!fs.existsSync(searchDir) || !fs.statSync(searchDir).isDirectory()) {
    fail(`Directory not found: ${searchDir}`);
  }

  const videos = collectVideos(searchDir, args.recursive);
  if (videos.length === 0) {
    fail(`No video files found in: ${searchDir}`);
  }

  log('info', args.logLevel, `Browsing: ${searchDir} (${videos.length} videos found)`);

  const selected = args.useRofi
    ? showRofiMenu(videos, searchDir, args.recursive, scriptPath, args.logLevel)
    : showFzfMenu(videos);

  if (!selected) return null;
  return { target: selected, kind: 'file' };
}

function registerCleanup(args: Args): void {
  process.on('SIGINT', () => {
    stopOverlay(args);
    process.exit(130);
  });
  process.on('SIGTERM', () => {
    stopOverlay(args);
    process.exit(143);
  });
}

function resolveMainConfigPath(): string {
  return resolveConfigFilePath({
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
    existsSync: fs.existsSync,
  });
}

function runDoctor(args: Args, appPath: string | null, mpvSocketPath: string): never {
  const checks: Array<{ label: string; ok: boolean; detail: string }> = [
    {
      label: 'app binary',
      ok: Boolean(appPath),
      detail: appPath || 'not found (set SUBMINER_APPIMAGE_PATH)',
    },
    {
      label: 'mpv',
      ok: commandExists('mpv'),
      detail: commandExists('mpv') ? 'found' : 'missing',
    },
    {
      label: 'yt-dlp',
      ok: commandExists('yt-dlp'),
      detail: commandExists('yt-dlp') ? 'found' : 'missing (optional unless YouTube URLs)',
    },
    {
      label: 'ffmpeg',
      ok: commandExists('ffmpeg'),
      detail: commandExists('ffmpeg') ? 'found' : 'missing (optional unless subtitle generation)',
    },
    {
      label: 'fzf',
      ok: commandExists('fzf'),
      detail: commandExists('fzf') ? 'found' : 'missing (optional if using rofi)',
    },
    {
      label: 'rofi',
      ok: commandExists('rofi'),
      detail: commandExists('rofi') ? 'found' : 'missing (optional if using fzf)',
    },
    {
      label: 'config',
      ok: fs.existsSync(resolveMainConfigPath()),
      detail: resolveMainConfigPath(),
    },
    {
      label: 'mpv socket path',
      ok: true,
      detail: mpvSocketPath,
    },
  ];

  const hasHardFailure = checks.some((entry) =>
    entry.label === 'app binary' || entry.label === 'mpv' ? !entry.ok : false,
  );

  for (const check of checks) {
    log(check.ok ? 'info' : 'warn', args.logLevel, `[doctor] ${check.label}: ${check.detail}`);
  }
  process.exit(hasHardFailure ? 1 : 0);
}

async function main(): Promise<void> {
  const scriptPath = process.argv[1] || 'subminer';
  const scriptName = path.basename(scriptPath);
  const launcherConfig = loadLauncherYoutubeSubgenConfig();
  const launcherJellyfinConfig = loadLauncherJellyfinConfig();
  const args = parseArgs(process.argv.slice(2), scriptName, launcherConfig);
  const pluginRuntimeConfig = readPluginRuntimeConfig(args.logLevel);
  const mpvSocketPath = pluginRuntimeConfig.socketPath;

  log('debug', args.logLevel, `Wrapper log level set to: ${args.logLevel}`);

  const appPath = findAppBinary(process.argv[1] || 'subminer');
  if (args.doctor) {
    runDoctor(args, appPath, mpvSocketPath);
  }

  if (args.configPath) {
    process.stdout.write(`${resolveMainConfigPath()}\n`);
    return;
  }

  if (args.configShow) {
    const configPath = resolveMainConfigPath();
    if (!fs.existsSync(configPath)) {
      fail(`Config file not found: ${configPath}`);
    }
    const contents = fs.readFileSync(configPath, 'utf8');
    process.stdout.write(contents);
    if (!contents.endsWith('\n')) {
      process.stdout.write('\n');
    }
    return;
  }

  if (args.mpvSocket) {
    process.stdout.write(`${mpvSocketPath}\n`);
    return;
  }

  if (args.mpvStatus) {
    const ready = await waitForUnixSocketReady(mpvSocketPath, 500);
    log(
      ready ? 'info' : 'warn',
      args.logLevel,
      `[mpv] socket ${ready ? 'ready' : 'not ready'}: ${mpvSocketPath}`,
    );
    process.exit(ready ? 0 : 1);
  }

  if (!appPath) {
    if (process.platform === 'darwin') {
      fail(
        'SubMiner app binary not found. Install SubMiner.app to /Applications or ~/Applications, or set SUBMINER_APPIMAGE_PATH.',
      );
    }
    fail('SubMiner AppImage not found. Install to ~/.local/bin/ or set SUBMINER_APPIMAGE_PATH.');
  }
  state.appPath = appPath;

  if (args.appPassthrough) {
    runAppCommandWithInherit(appPath, args.appArgs);
  }

  if (args.mpvIdle) {
    await launchMpvIdleDetached(mpvSocketPath, appPath, args);
    const ready = await waitForUnixSocketReady(mpvSocketPath, 8000);
    if (!ready) {
      fail(`MPV IPC socket not ready after idle launch: ${mpvSocketPath}`);
    }
    log('info', args.logLevel, `[mpv] idle instance ready on ${mpvSocketPath}`);
    return;
  }

  if (args.texthookerOnly) {
    launchTexthookerOnly(appPath, args);
  }

  if (args.jellyfin) {
    const forwarded = ['--jellyfin'];
    if (args.logLevel !== 'info') forwarded.push('--log-level', args.logLevel);
    runAppCommandWithInherit(appPath, forwarded);
  }

  if (args.jellyfinLogin) {
    const serverUrl = args.jellyfinServer || launcherJellyfinConfig.serverUrl || '';
    const username = args.jellyfinUsername || launcherJellyfinConfig.username || '';
    const password = args.jellyfinPassword || '';
    if (!serverUrl || !username || !password) {
      fail(
        '--jellyfin-login requires server, username, and password. Pass flags or run `subminer --jellyfin`.',
      );
    }
    const forwarded = [
      '--jellyfin-login',
      '--jellyfin-server',
      serverUrl,
      '--jellyfin-username',
      username,
      '--jellyfin-password',
      password,
    ];
    if (args.logLevel !== 'info') forwarded.push('--log-level', args.logLevel);
    runAppCommandWithInherit(appPath, forwarded);
  }

  if (args.jellyfinLogout) {
    const forwarded = ['--jellyfin-logout'];
    if (args.logLevel !== 'info') forwarded.push('--log-level', args.logLevel);
    runAppCommandWithInherit(appPath, forwarded);
  }

  if (args.jellyfinPlay) {
    if (!args.useRofi && !commandExists('fzf')) {
      fail('fzf not found. Install fzf or use -R for rofi.');
    }
    if (args.useRofi && !commandExists('rofi')) {
      fail('rofi not found. Install rofi or omit -R for fzf.');
    }
    await runJellyfinPlayMenu(appPath, args, scriptPath, mpvSocketPath);
  }

  if (args.jellyfinDiscovery) {
    const forwarded = ['--start'];
    if (args.logLevel !== 'info') forwarded.push('--log-level', args.logLevel);
    runAppCommandWithInherit(appPath, forwarded);
  }

  if (!args.target) {
    checkPickerDependencies(args);
  }

  const targetChoice = await chooseTarget(args, process.argv[1] || 'subminer');
  if (!targetChoice) {
    log('info', args.logLevel, 'No video selected, exiting');
    process.exit(0);
  }

  checkDependencies({
    ...args,
    target: targetChoice ? targetChoice.target : args.target,
    targetKind: targetChoice ? targetChoice.kind : 'url',
  });

  registerCleanup(args);

  let selectedTarget = targetChoice
    ? {
        target: targetChoice.target,
        kind: targetChoice.kind as 'file' | 'url',
      }
    : { target: args.target, kind: 'url' as const };

  const isYoutubeUrl = selectedTarget.kind === 'url' && isYoutubeTarget(selectedTarget.target);
  let preloadedSubtitles: { primaryPath?: string; secondaryPath?: string } | undefined;

  if (isYoutubeUrl && args.youtubeSubgenMode === 'preprocess') {
    log('info', args.logLevel, 'YouTube subtitle mode: preprocess');
    const generated = await generateYoutubeSubtitles(selectedTarget.target, args);
    preloadedSubtitles = {
      primaryPath: generated.primaryPath,
      secondaryPath: generated.secondaryPath,
    };
    log(
      'info',
      args.logLevel,
      `YouTube preprocess result: primary=${generated.primaryPath ? 'ready' : 'missing'}, secondary=${generated.secondaryPath ? 'ready' : 'missing'}`,
    );
  } else if (isYoutubeUrl && args.youtubeSubgenMode === 'automatic') {
    log('info', args.logLevel, 'YouTube subtitle mode: automatic (background)');
  } else if (isYoutubeUrl) {
    log('info', args.logLevel, 'YouTube subtitle mode: off');
  }

  startMpv(
    selectedTarget.target,
    selectedTarget.kind,
    args,
    mpvSocketPath,
    appPath,
    preloadedSubtitles,
  );

  if (isYoutubeUrl && args.youtubeSubgenMode === 'automatic') {
    void generateYoutubeSubtitles(selectedTarget.target, args, async (lang, subtitlePath) => {
      try {
        await loadSubtitleIntoMpv(mpvSocketPath, subtitlePath, lang === 'primary', args.logLevel);
      } catch (error) {
        log(
          'warn',
          args.logLevel,
          `Generated subtitle ready but failed to load in mpv: ${(error as Error).message}`,
        );
      }
    }).catch((error) => {
      log(
        'warn',
        args.logLevel,
        `Background subtitle generation failed: ${(error as Error).message}`,
      );
    });
  }

  const ready = await waitForUnixSocketReady(mpvSocketPath, 10000);
  const shouldStartOverlay = args.startOverlay || args.autoStartOverlay;
  if (shouldStartOverlay) {
    if (ready) {
      log('info', args.logLevel, 'MPV IPC socket ready, starting SubMiner overlay');
    } else {
      log(
        'info',
        args.logLevel,
        'MPV IPC socket not ready after timeout, starting SubMiner overlay anyway',
      );
    }
    await startOverlay(appPath, args, mpvSocketPath);
  } else if (ready) {
    log(
      'info',
      args.logLevel,
      'MPV IPC socket ready, overlay auto-start disabled (use y-s to start)',
    );
  } else {
    log(
      'info',
      args.logLevel,
      'MPV IPC socket not ready yet, overlay auto-start disabled (use y-s to start)',
    );
  }

  await new Promise<void>((resolve) => {
    if (!state.mpvProc) {
      stopOverlay(args);
      resolve();
      return;
    }
    state.mpvProc.on('exit', (code) => {
      stopOverlay(args);
      process.exitCode = code ?? 0;
      resolve();
    });
  });
}

main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  fail(message);
});
