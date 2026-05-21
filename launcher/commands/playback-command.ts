import fs from 'node:fs';
import os from 'node:os';
import { spawn } from 'node:child_process';
import { fail, log } from '../log.js';
import { commandExists, isYoutubeTarget, realpathMaybe, resolvePathMaybe } from '../util.js';
import { collectVideos, showFzfMenu, showRofiMenu } from '../picker.js';
import {
  cleanupPlaybackSession,
  launchAppCommandDetached,
  resolveLauncherRuntimePluginPath,
  isRunningAppControlServerAvailable,
  startMpv,
  startOverlay,
  state,
  stopOverlay,
  waitForUnixSocketReady,
} from '../mpv.js';
import type { Args } from '../types.js';
import { nowMs } from '../time.js';
import type { LauncherCommandContext } from './context.js';
import { ensureLauncherSetupReady } from '../setup-gate.js';
import {
  getDefaultConfigDir,
  getSetupStatePath,
  readSetupState,
} from '../../src/shared/setup-state.js';
import { detectInstalledFirstRunPluginCandidates } from '../../src/main/runtime/first-run-setup-plugin.js';
import { hasLauncherExternalYomitanProfileConfig } from '../config.js';

const SETUP_WAIT_TIMEOUT_MS = 10 * 60 * 1000;
const SETUP_POLL_INTERVAL_MS = 500;

function checkDependencies(args: Args): void {
  const missing: string[] = [];

  if (!commandExists('mpv')) missing.push('mpv');

  const isYoutubeUrl = args.targetKind === 'url' && isYoutubeTarget(args.target);
  if (args.targetKind === 'url' && !isYoutubeUrl && !commandExists('yt-dlp')) {
    missing.push('yt-dlp');
  }

  if (args.targetKind === 'url' && !isYoutubeUrl && !commandExists('ffmpeg')) {
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

function registerCleanup(context: LauncherCommandContext): void {
  const { args, processAdapter } = context;
  processAdapter.onSignal('SIGINT', () => {
    stopOverlay(args);
    processAdapter.exit(130);
  });
  processAdapter.onSignal('SIGTERM', () => {
    stopOverlay(args);
    processAdapter.exit(143);
  });
}

async function ensurePlaybackSetupReady(context: LauncherCommandContext): Promise<void> {
  const { args, appPath } = context;
  if (!appPath) return;

  const configDir = getDefaultConfigDir({
    xdgConfigHome: process.env.XDG_CONFIG_HOME,
    homeDir: os.homedir(),
  });
  const statePath = getSetupStatePath(configDir);
  const ready = await ensureLauncherSetupReady({
    readSetupState: () => readSetupState(statePath),
    isExternalYomitanConfigured: () => hasLauncherExternalYomitanProfileConfig(),
    hasLegacyMpvPlugin: () =>
      detectInstalledFirstRunPluginCandidates({
        platform: process.platform,
        homeDir: os.homedir(),
        xdgConfigHome: process.env.XDG_CONFIG_HOME,
        appDataDir: process.env.APPDATA,
      }).length > 0,
    launchSetupApp: () => {
      const setupArgs = ['--background', '--setup'];
      if (args.logLevel) {
        setupArgs.push('--log-level', args.logLevel);
      }
      const child = spawn(appPath, setupArgs, {
        detached: true,
        stdio: 'ignore',
      });
      child.unref();
    },
    sleep: (ms) => new Promise((resolve) => setTimeout(resolve, ms)),
    now: () => nowMs(),
    timeoutMs: SETUP_WAIT_TIMEOUT_MS,
    pollIntervalMs: SETUP_POLL_INTERVAL_MS,
  });

  if (!ready) {
    fail('SubMiner setup is incomplete. Complete setup in the app, then retry playback.');
  }
}

export async function runPlaybackCommand(context: LauncherCommandContext): Promise<void> {
  return runPlaybackCommandWithDeps(context, {
    ensurePlaybackSetupReady,
    chooseTarget,
    checkDependencies,
    registerCleanup,
    startMpv,
    waitForUnixSocketReady,
    startOverlay,
    launchAppCommandDetached,
    isAppControlServerAvailable: isRunningAppControlServerAvailable,
    log,
    cleanupPlaybackSession,
    getMpvProc: () => state.mpvProc,
  });
}

type PlaybackCommandDeps = {
  ensurePlaybackSetupReady: (context: LauncherCommandContext) => Promise<void>;
  chooseTarget: (
    args: Args,
    scriptPath: string,
  ) => Promise<{ target: string; kind: 'file' | 'url' } | null>;
  checkDependencies: (args: Args) => void;
  registerCleanup: (context: LauncherCommandContext) => void;
  startMpv: typeof startMpv;
  waitForUnixSocketReady: typeof waitForUnixSocketReady;
  startOverlay: typeof startOverlay;
  launchAppCommandDetached: typeof launchAppCommandDetached;
  isAppControlServerAvailable?: (logLevel: Args['logLevel']) => Promise<boolean>;
  log: typeof log;
  cleanupPlaybackSession: typeof cleanupPlaybackSession;
  getMpvProc: () => typeof state.mpvProc;
};

export async function runPlaybackCommandWithDeps(
  context: LauncherCommandContext,
  deps: PlaybackCommandDeps,
): Promise<void> {
  const { args, appPath, scriptPath, mpvSocketPath, pluginRuntimeConfig, processAdapter } = context;
  if (!appPath) {
    fail('SubMiner AppImage not found. Install to ~/.local/bin/ or set SUBMINER_APPIMAGE_PATH.');
  }

  await deps.ensurePlaybackSetupReady(context);

  if (!args.target) {
    checkPickerDependencies(args);
  }

  const targetChoice = await deps.chooseTarget(args, scriptPath);
  if (!targetChoice) {
    deps.log('info', args.logLevel, 'No video selected, exiting');
    processAdapter.exit(0);
  }

  deps.checkDependencies({
    ...args,
    target: targetChoice ? targetChoice.target : args.target,
    targetKind: targetChoice ? targetChoice.kind : 'url',
  });

  deps.registerCleanup(context);

  const selectedTarget = targetChoice
    ? {
        target: targetChoice.target,
        kind: targetChoice.kind as 'file' | 'url',
      }
    : { target: args.target, kind: 'url' as const };

  const isYoutubeUrl = selectedTarget.kind === 'url' && isYoutubeTarget(selectedTarget.target);
  const isAppOwnedYoutubeFlow = isYoutubeUrl;
  const youtubeMode = args.youtubeMode ?? 'download';

  if (isYoutubeUrl) {
    deps.log('info', args.logLevel, 'YouTube subtitle flow: app-owned picker after mpv bootstrap');
  }

  const pluginAutoStartEnabled = pluginRuntimeConfig.autoStart;
  const shouldLauncherAttachRunningApp =
    pluginAutoStartEnabled &&
    !args.startOverlay &&
    !args.autoStartOverlay &&
    !isAppOwnedYoutubeFlow &&
    ((await deps.isAppControlServerAvailable?.(args.logLevel)) ?? false);
  const effectivePluginRuntimeConfig = shouldLauncherAttachRunningApp
    ? { ...pluginRuntimeConfig, autoStart: false }
    : pluginRuntimeConfig;

  const shouldPauseUntilOverlayReady =
    pluginRuntimeConfig.autoStart &&
    pluginRuntimeConfig.autoStartVisibleOverlay &&
    pluginRuntimeConfig.autoStartPauseUntilReady;

  if (shouldPauseUntilOverlayReady) {
    deps.log(
      'info',
      args.logLevel,
      'Configured to pause mpv until overlay and tokenization are ready',
    );
  }

  await deps.startMpv(
    selectedTarget.target,
    selectedTarget.kind,
    args,
    mpvSocketPath,
    appPath,
    undefined,
    {
      startPaused: shouldPauseUntilOverlayReady || isAppOwnedYoutubeFlow,
      disableYoutubeSubtitleAutoLoad: isAppOwnedYoutubeFlow,
      runtimePluginPath: resolveLauncherRuntimePluginPath({ appPath, scriptPath }),
      runtimePluginConfig: {
        ...effectivePluginRuntimeConfig,
        backend: args.backend,
        texthookerEnabled: args.useTexthooker && effectivePluginRuntimeConfig.texthookerEnabled,
      },
    },
  );

  const ready = await deps.waitForUnixSocketReady(mpvSocketPath, 10000);
  const shouldStartOverlay =
    args.startOverlay ||
    args.autoStartOverlay ||
    isAppOwnedYoutubeFlow ||
    shouldLauncherAttachRunningApp;
  if (shouldStartOverlay) {
    if (ready) {
      deps.log('info', args.logLevel, 'MPV IPC socket ready, starting SubMiner overlay');
    } else {
      deps.log(
        'info',
        args.logLevel,
        'MPV IPC socket not ready after timeout, starting SubMiner overlay anyway',
      );
    }
    const extraAppArgs = isAppOwnedYoutubeFlow
      ? ['--youtube-play', selectedTarget.target, '--youtube-mode', youtubeMode]
      : shouldLauncherAttachRunningApp
        ? [
            pluginRuntimeConfig.autoStartVisibleOverlay
              ? '--show-visible-overlay'
              : '--hide-visible-overlay',
            ...(pluginRuntimeConfig.texthookerEnabled ? ['--texthooker'] : []),
          ]
        : [];
    await deps.startOverlay(appPath, args, mpvSocketPath, extraAppArgs);
  } else if (pluginAutoStartEnabled) {
    if (ready) {
      deps.log('info', args.logLevel, 'MPV IPC socket ready, relying on mpv plugin auto-start');
    } else {
      deps.log(
        'info',
        args.logLevel,
        'MPV IPC socket not ready yet, relying on mpv plugin auto-start',
      );
    }
  } else if (ready) {
    deps.log(
      'info',
      args.logLevel,
      'MPV IPC socket ready, overlay auto-start disabled (use y-s to start)',
    );
  } else {
    deps.log(
      'info',
      args.logLevel,
      'MPV IPC socket not ready yet, overlay auto-start disabled (use y-s to start)',
    );
  }

  await new Promise<void>((resolve) => {
    const mpvProc = deps.getMpvProc();
    if (!mpvProc) {
      stopOverlay(args);
      resolve();
      return;
    }

    const finalize = (code: number | null | undefined) => {
      void deps.cleanupPlaybackSession(args).finally(() => {
        processAdapter.setExitCode(code ?? 0);
        resolve();
      });
    };

    if (mpvProc.exitCode !== null && mpvProc.exitCode !== undefined) {
      finalize(mpvProc.exitCode);
      return;
    }

    mpvProc.once('exit', (code) => {
      finalize(code);
    });
  });
}
