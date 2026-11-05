import path from 'node:path';
import type { IpcMain } from 'electron';
import { toJellyfinOriginalStreamUrl, type createOverlayManager } from '../../core/services';
import type { ResolvedConfig } from '../../types';
import { IPC_CHANNELS } from '../../shared/ipc/contracts';
import { registerSubtitleGenerationIpc } from '../runtime/subtitle-generation-ipc';
import { openSubtitleGenerationModal } from '../runtime/subtitle-generation-open';
import { createSubtitleGenerationRuntime } from '../runtime/subtitle-generation-runtime';
import type { AppState } from '../state';
import type { JellyfinAssembly } from './jellyfin';
import type { OverlayModalsAssembly } from './overlay-modals';
import type { YoutubeAssembly } from './youtube';

export interface SubtitleGenerationAssemblyDeps {
  appState: Pick<AppState, 'mpvClient'>;
  getConfig: () => ResolvedConfig;
  /** Managed Whisper models live next to the config file. */
  getModelDirectory: () => string;
  overlayManager: Pick<ReturnType<typeof createOverlayManager>, 'getMainWindow' | 'getModalWindow'>;
  overlayModals: Pick<OverlayModalsAssembly, 'createOverlayHostedModalOpenDeps'>;
  youtube: {
    whisperSubtitles: Pick<YoutubeAssembly['whisperSubtitles'], 'generate'>;
    flow: Pick<YoutubeAssembly['flow'], 'createSubtitleTempDir'>;
    playback: Pick<YoutubeAssembly['playback'], 'getYoutubeSourceUrlForStream'>;
  };
  jellyfin: Pick<JellyfinAssembly, 'getActivePlayback'>;
  ipcMain: Pick<IpcMain, 'handle'>;
}

export type SubtitleGenerationAssembly = ReturnType<typeof createSubtitleGenerationAssembly>;

/**
 * Owns Whisper subtitle generation for the playing media (local files, YouTube and Jellyfin
 * streams): the generation runtime, its overlay modal and its renderer IPC.
 */
export function createSubtitleGenerationAssembly(deps: SubtitleGenerationAssemblyDeps) {
  const { appState, overlayManager, youtube } = deps;

  function getOverlayWindows() {
    return [overlayManager.getMainWindow(), overlayManager.getModalWindow()];
  }

  const runtime = createSubtitleGenerationRuntime({
    getConfig: () => deps.getConfig().subtitleGeneration,
    getModelDirectory: () => deps.getModelDirectory(),
    generateYoutube: async ({ url, config, signal, onProgress }) =>
      youtube.whisperSubtitles.generate({
        url,
        config,
        signal,
        onProgress,
        outputPath: path.join(await youtube.flow.createSubtitleTempDir(), 'youtube-whisper.ja.srt'),
      }),
    resolveYoutubeSourceUrl: (mediaPath) =>
      youtube.playback.getYoutubeSourceUrlForStream(mediaPath),
    resolveJellyfinMedia: (mediaPath) => {
      const playback = deps.jellyfin.getActivePlayback();
      const sourceUrl = playback && toJellyfinOriginalStreamUrl(mediaPath, playback.itemId);
      return sourceUrl
        ? {
            sourceUrl,
            transcoded: playback.playMethod === 'Transcode',
            audioStreamIndex: playback.audioStreamIndex ?? null,
          }
        : null;
    },
    getMpvClient: () => appState.mpvClient,
    onProgress: (progress) => {
      for (const window of getOverlayWindows()) {
        if (window && !window.isDestroyed())
          window.webContents.send(IPC_CHANNELS.event.subtitleGenerationProgress, progress);
      }
    },
  });

  function openModal(): Promise<boolean> {
    return openSubtitleGenerationModal(deps.overlayModals.createOverlayHostedModalOpenDeps());
  }

  /** Registers the generation IPC; only the overlay windows may call it. */
  function registerIpc(): void {
    registerSubtitleGenerationIpc({
      ipc: deps.ipcMain,
      isAllowedSender: (sender) =>
        getOverlayWindows().some(
          (window) => window && !window.isDestroyed() && window.webContents === sender,
        ),
      runtime,
      openModal,
    });
  }

  return { runtime, openModal, registerIpc };
}
