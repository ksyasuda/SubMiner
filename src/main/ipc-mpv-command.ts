import type { RuntimeOptionApplyResult, RuntimeOptionId } from '../types';
import { handleMpvCommandFromIpc } from '../core/services';
import { createMpvCommandRuntimeServiceDeps } from './dependencies';
import { SPECIAL_COMMANDS } from '../config';
import { resolveProxyCommandOsdRuntime } from './runtime/mpv-proxy-osd';

type MpvPropertyClientLike = {
  connected: boolean;
  requestProperty: (name: string) => Promise<unknown>;
};

export interface MpvCommandFromIpcRuntimeDeps {
  triggerSubsyncFromConfig: () => void;
  openRuntimeOptionsPalette: () => void;
  openJimaku: () => void;
  openYoutubeTrackPicker: () => void | Promise<void>;
  openPlaylistBrowser: () => void | Promise<void>;
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1) => RuntimeOptionApplyResult;
  showMpvOsd: (text: string) => void;
  showRawMpvOsd?: (text: string) => void;
  showPlaybackFeedback?: (text: string) => void;
  replayCurrentSubtitle: () => void;
  playNextSubtitle: () => void;
  sendMpvCommand: (command: (string | number)[]) => void;
  getMpvClient: () => MpvPropertyClientLike | null;
  isMpvConnected: () => boolean;
  hasRuntimeOptionsManager: () => boolean;
}

export function handleMpvCommandFromIpcRuntime(
  command: (string | number)[],
  deps: MpvCommandFromIpcRuntimeDeps,
): void {
  handleMpvCommandFromIpc(
    command,
    createMpvCommandRuntimeServiceDeps({
      specialCommands: SPECIAL_COMMANDS,
      triggerSubsyncFromConfig: deps.triggerSubsyncFromConfig,
      openRuntimeOptionsPalette: deps.openRuntimeOptionsPalette,
      openJimaku: deps.openJimaku,
      openYoutubeTrackPicker: deps.openYoutubeTrackPicker,
      openPlaylistBrowser: deps.openPlaylistBrowser,
      runtimeOptionsCycle: deps.cycleRuntimeOption,
      showMpvOsd: deps.showMpvOsd,
      showRawMpvOsd: deps.showRawMpvOsd,
      showPlaybackFeedback: deps.showPlaybackFeedback,
      mpvReplaySubtitle: deps.replayCurrentSubtitle,
      mpvPlayNextSubtitle: deps.playNextSubtitle,
      mpvSendCommand: deps.sendMpvCommand,
      resolveProxyCommandOsd: (nextCommand) =>
        resolveProxyCommandOsdRuntime(nextCommand, deps.getMpvClient),
      isMpvConnected: deps.isMpvConnected,
      hasRuntimeOptionsManager: deps.hasRuntimeOptionsManager,
    }),
  );
}
