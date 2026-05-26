import type { RuntimeOptionApplyResult, RuntimeOptionId } from '../../types';
import type { SessionActionId } from '../../types/session-bindings';
import type { SessionActionDispatchRequest } from '../../types/runtime';

export interface SessionActionExecutorDeps {
  toggleStatsOverlay: () => void;
  toggleVisibleOverlay: () => void;
  copyCurrentSubtitle: () => void;
  copySubtitleCount: (count: number) => void;
  updateLastCardFromClipboard: () => Promise<void>;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsyncFromConfig: () => Promise<void>;
  mineSentenceCard: () => Promise<void>;
  mineSentenceCount: (count: number) => void;
  toggleSecondarySub: () => void;
  toggleSubtitleSidebar: () => void;
  markLastCardAsAudioCard: () => Promise<void>;
  markActiveVideoWatched: () => Promise<boolean>;
  openRuntimeOptionsPalette: () => void;
  openSessionHelp: () => void;
  openCharacterDictionaryManager: () => void;
  openControllerSelect: () => void;
  openControllerDebug: () => void;
  openJimaku: () => void;
  openYoutubeTrackPicker: () => void | Promise<void>;
  openPlaylistBrowser: () => boolean | void | Promise<boolean | void>;
  replayCurrentSubtitle: () => void;
  playNextSubtitle: () => void;
  shiftSubDelayToAdjacentSubtitle: (direction: 'next' | 'previous') => Promise<void>;
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1) => RuntimeOptionApplyResult;
  playNextPlaylistItem: () => void;
  showMpvOsd: (text: string) => void;
}

function resolveCount(count: number | undefined): number {
  const normalized = typeof count === 'number' && Number.isInteger(count) ? count : 1;
  return Math.min(9, Math.max(1, normalized));
}

function assertUnreachableSessionAction(actionId: never): never {
  throw new Error(`Unhandled session action: ${String(actionId)}`);
}

export async function dispatchSessionAction(
  request: SessionActionDispatchRequest,
  deps: SessionActionExecutorDeps,
): Promise<void> {
  switch (request.actionId) {
    case 'toggleStatsOverlay':
      deps.toggleStatsOverlay();
      return;
    case 'toggleVisibleOverlay':
      deps.toggleVisibleOverlay();
      return;
    case 'copySubtitle':
      deps.copyCurrentSubtitle();
      return;
    case 'copySubtitleMultiple':
      deps.copySubtitleCount(resolveCount(request.payload?.count));
      return;
    case 'updateLastCardFromClipboard':
      await deps.updateLastCardFromClipboard();
      return;
    case 'triggerFieldGrouping':
      await deps.triggerFieldGrouping();
      return;
    case 'triggerSubsync':
      await deps.triggerSubsyncFromConfig();
      return;
    case 'mineSentence':
      await deps.mineSentenceCard();
      return;
    case 'mineSentenceMultiple':
      deps.mineSentenceCount(resolveCount(request.payload?.count));
      return;
    case 'toggleSecondarySub':
      deps.toggleSecondarySub();
      return;
    case 'toggleSubtitleSidebar':
      deps.toggleSubtitleSidebar();
      return;
    case 'markAudioCard':
      await deps.markLastCardAsAudioCard();
      return;
    case 'markWatched': {
      const marked = await deps.markActiveVideoWatched();
      if (marked) {
        deps.showMpvOsd('Marked as watched');
        deps.playNextPlaylistItem();
      }
      return;
    }
    case 'openRuntimeOptions':
      deps.openRuntimeOptionsPalette();
      return;
    case 'openSessionHelp':
      deps.openSessionHelp();
      return;
    case 'openCharacterDictionaryManager':
      deps.openCharacterDictionaryManager();
      return;
    case 'openControllerSelect':
      deps.openControllerSelect();
      return;
    case 'openControllerDebug':
      deps.openControllerDebug();
      return;
    case 'openJimaku':
      deps.openJimaku();
      return;
    case 'openYoutubePicker':
      await deps.openYoutubeTrackPicker();
      return;
    case 'openPlaylistBrowser':
      await deps.openPlaylistBrowser();
      return;
    case 'replayCurrentSubtitle':
      deps.replayCurrentSubtitle();
      return;
    case 'playNextSubtitle':
      deps.playNextSubtitle();
      return;
    case 'shiftSubDelayPrevLine':
      await deps.shiftSubDelayToAdjacentSubtitle('previous');
      return;
    case 'shiftSubDelayNextLine':
      await deps.shiftSubDelayToAdjacentSubtitle('next');
      return;
    case 'cycleRuntimeOption': {
      const runtimeOptionId = request.payload?.runtimeOptionId as RuntimeOptionId | undefined;
      if (!runtimeOptionId) {
        deps.showMpvOsd('Runtime option id is required.');
        return;
      }
      const direction = request.payload?.direction === -1 ? -1 : 1;
      const result = deps.cycleRuntimeOption(runtimeOptionId, direction);
      if (!result.ok && result.error) {
        deps.showMpvOsd(result.error);
      }
      return;
    }
    default:
      return assertUnreachableSessionAction(request.actionId);
  }
}
