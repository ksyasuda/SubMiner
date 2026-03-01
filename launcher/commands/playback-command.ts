import fs from 'node:fs';
import path from 'node:path';
import { fail, log } from '../log.js';
import { commandExists, isYoutubeTarget, realpathMaybe, resolvePathMaybe } from '../util.js';
import { collectVideos, showFzfMenu, showRofiMenu } from '../picker.js';
import {
  loadSubtitleIntoMpv,
  startMpv,
  startOverlay,
  state,
  stopOverlay,
  waitForUnixSocketReady,
} from '../mpv.js';
import { generateYoutubeSubtitles } from '../youtube.js';
import type { Args } from '../types.js';
import type { LauncherCommandContext } from './context.js';

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

export async function runPlaybackCommand(context: LauncherCommandContext): Promise<void> {
  const { args, appPath, scriptPath, mpvSocketPath, pluginRuntimeConfig, processAdapter } = context;
  if (!appPath) {
    fail('SubMiner AppImage not found. Install to ~/.local/bin/ or set SUBMINER_APPIMAGE_PATH.');
  }

  if (!args.target) {
    checkPickerDependencies(args);
  }

  const targetChoice = await chooseTarget(args, scriptPath);
  if (!targetChoice) {
    log('info', args.logLevel, 'No video selected, exiting');
    processAdapter.exit(0);
  }

  checkDependencies({
    ...args,
    target: targetChoice ? targetChoice.target : args.target,
    targetKind: targetChoice ? targetChoice.kind : 'url',
  });

  registerCleanup(context);

  const selectedTarget = targetChoice
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

  const shouldPauseUntilOverlayReady =
    pluginRuntimeConfig.autoStart &&
    pluginRuntimeConfig.autoStartVisibleOverlay &&
    pluginRuntimeConfig.autoStartPauseUntilReady;

  if (shouldPauseUntilOverlayReady) {
    log(
      'info',
      args.logLevel,
      'Configured to pause mpv until overlay and tokenization are ready',
    );
  }

  startMpv(
    selectedTarget.target,
    selectedTarget.kind,
    args,
    mpvSocketPath,
    appPath,
    preloadedSubtitles,
    { startPaused: shouldPauseUntilOverlayReady },
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
  const pluginAutoStartEnabled = pluginRuntimeConfig.autoStart;
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
  } else if (pluginAutoStartEnabled) {
    if (ready) {
      log('info', args.logLevel, 'MPV IPC socket ready, relying on mpv plugin auto-start');
    } else {
      log(
        'info',
        args.logLevel,
        'MPV IPC socket not ready yet, relying on mpv plugin auto-start',
      );
    }
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
    const mpvProc = state.mpvProc;
    if (!mpvProc) {
      stopOverlay(args);
      resolve();
      return;
    }

    const finalize = (code: number | null | undefined) => {
      stopOverlay(args);
      processAdapter.setExitCode(code ?? 0);
      resolve();
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
