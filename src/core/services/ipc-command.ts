import {
  RuntimeOptionApplyResult,
  RuntimeOptionId,
  SubsyncManualRunRequest,
  SubsyncResult,
} from '../../types';

export interface HandleMpvCommandFromIpcOptions {
  specialCommands: {
    SUBSYNC_TRIGGER: string;
    RUNTIME_OPTIONS_OPEN: string;
    JIMAKU_OPEN: string;
    RUNTIME_OPTION_CYCLE_PREFIX: string;
    REPLAY_SUBTITLE: string;
    PLAY_NEXT_SUBTITLE: string;
    SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START: string;
    SHIFT_SUB_DELAY_TO_PREVIOUS_SUBTITLE_START: string;
    YOUTUBE_PICKER_OPEN: string;
    PLAYLIST_BROWSER_OPEN: string;
  };
  triggerSubsyncFromConfig: () => void;
  openRuntimeOptionsPalette: () => void;
  openJimaku: () => void;
  openYoutubeTrackPicker: () => void | Promise<void>;
  openPlaylistBrowser: () => void | Promise<void>;
  runtimeOptionsCycle: (id: RuntimeOptionId, direction: 1 | -1) => RuntimeOptionApplyResult;
  showMpvOsd: (text: string) => void;
  mpvReplaySubtitle: () => void;
  mpvPlayNextSubtitle: () => void;
  shiftSubDelayToAdjacentSubtitle: (direction: 'next' | 'previous') => Promise<void>;
  mpvSendCommand: (command: (string | number)[]) => void;
  resolveProxyCommandOsd?: (command: (string | number)[]) => Promise<string | null>;
  isMpvConnected: () => boolean;
  hasRuntimeOptionsManager: () => boolean;
}

const MPV_PROPERTY_COMMANDS = new Set([
  'add',
  'set',
  'set_property',
  'cycle',
  'cycle-values',
  'multiply',
]);

function resolveProxyCommandOsdTemplate(command: (string | number)[]): string | null {
  const operation = typeof command[0] === 'string' ? command[0] : '';
  const property = typeof command[1] === 'string' ? command[1] : '';
  if (!MPV_PROPERTY_COMMANDS.has(operation)) return null;
  if (property === 'sub-pos') {
    return 'Subtitle position: ${sub-pos}';
  }
  if (property === 'sid') {
    return 'Subtitle track: ${sid}';
  }
  if (property === 'secondary-sid') {
    return 'Secondary subtitle track: ${secondary-sid}';
  }
  if (property === 'sub-delay') {
    return 'Subtitle delay: ${sub-delay}';
  }
  return null;
}

function showResolvedProxyCommandOsd(
  command: (string | number)[],
  options: HandleMpvCommandFromIpcOptions,
): void {
  const template = resolveProxyCommandOsdTemplate(command);
  if (!template) return;

  const emit = async () => {
    try {
      const resolved = await options.resolveProxyCommandOsd?.(command);
      options.showMpvOsd(resolved || template);
    } catch {
      options.showMpvOsd(template);
    }
  };

  void emit();
}

export function handleMpvCommandFromIpc(
  command: (string | number)[],
  options: HandleMpvCommandFromIpcOptions,
): void {
  const first = typeof command[0] === 'string' ? command[0] : '';
  if (first === options.specialCommands.SUBSYNC_TRIGGER) {
    options.triggerSubsyncFromConfig();
    return;
  }

  if (first === options.specialCommands.RUNTIME_OPTIONS_OPEN) {
    options.openRuntimeOptionsPalette();
    return;
  }

  if (first === options.specialCommands.JIMAKU_OPEN) {
    options.openJimaku();
    return;
  }

  if (first === options.specialCommands.YOUTUBE_PICKER_OPEN) {
    void options.openYoutubeTrackPicker();
    return;
  }

  if (first === options.specialCommands.PLAYLIST_BROWSER_OPEN) {
    Promise.resolve()
      .then(() => options.openPlaylistBrowser())
      .catch((error) => {
        const message = error instanceof Error ? error.message : String(error);
        options.showMpvOsd(`Playlist browser failed: ${message}`);
      });
    return;
  }

  if (
    first === options.specialCommands.SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START ||
    first === options.specialCommands.SHIFT_SUB_DELAY_TO_PREVIOUS_SUBTITLE_START
  ) {
    const direction =
      first === options.specialCommands.SHIFT_SUB_DELAY_TO_NEXT_SUBTITLE_START
        ? 'next'
        : 'previous';
    options.shiftSubDelayToAdjacentSubtitle(direction).catch((error) => {
      options.showMpvOsd(`Subtitle delay shift failed: ${(error as Error).message}`);
    });
    return;
  }

  if (first.startsWith(options.specialCommands.RUNTIME_OPTION_CYCLE_PREFIX)) {
    if (!options.hasRuntimeOptionsManager()) return;
    const [, idToken, directionToken] = first.split(':');
    const id = idToken as RuntimeOptionId;
    const direction: 1 | -1 = directionToken === 'prev' ? -1 : 1;
    const result = options.runtimeOptionsCycle(id, direction);
    if (!result.ok && result.error) {
      options.showMpvOsd(result.error);
    }
    return;
  }

  if (options.isMpvConnected()) {
    if (first === options.specialCommands.REPLAY_SUBTITLE) {
      options.mpvReplaySubtitle();
    } else if (first === options.specialCommands.PLAY_NEXT_SUBTITLE) {
      options.mpvPlayNextSubtitle();
    } else {
      options.mpvSendCommand(command);
      showResolvedProxyCommandOsd(command, options);
    }
  }
}

export async function runSubsyncManualFromIpc(
  request: SubsyncManualRunRequest,
  options: {
    isSubsyncInProgress: () => boolean;
    setSubsyncInProgress: (inProgress: boolean) => void;
    showMpvOsd: (text: string) => void;
    runWithSpinner: (task: () => Promise<SubsyncResult>) => Promise<SubsyncResult>;
    runSubsyncManual: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
  },
): Promise<SubsyncResult> {
  if (options.isSubsyncInProgress()) {
    const busy = 'Subsync already running';
    options.showMpvOsd(busy);
    return { ok: false, message: busy };
  }
  try {
    options.setSubsyncInProgress(true);
    const result = await options.runWithSpinner(() => options.runSubsyncManual(request));
    options.showMpvOsd(result.message);
    return result;
  } catch (error) {
    const message = `Subsync failed: ${(error as Error).message}`;
    options.showMpvOsd(message);
    return { ok: false, message };
  } finally {
    options.setSubsyncInProgress(false);
  }
}
