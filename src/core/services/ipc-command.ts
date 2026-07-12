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
    ANIMETOSHO_OPEN: string;
    RUNTIME_OPTION_CYCLE_PREFIX: string;
    REPLAY_SUBTITLE: string;
    PLAY_NEXT_SUBTITLE: string;
    YOUTUBE_PICKER_OPEN: string;
    PLAYLIST_BROWSER_OPEN: string;
  };
  triggerSubsyncFromConfig: () => void;
  openRuntimeOptionsPalette: () => void;
  openJimaku: () => void;
  openAnimetosho: () => void;
  openYoutubeTrackPicker: () => void | Promise<void>;
  openPlaylistBrowser: () => void | Promise<void>;
  runtimeOptionsCycle: (id: RuntimeOptionId, direction: 1 | -1) => RuntimeOptionApplyResult;
  showMpvOsd: (text: string) => void;
  showRawMpvOsd?: (text: string) => void;
  showPlaybackFeedback?: (text: string) => void;
  mpvReplaySubtitle: () => void;
  mpvPlayNextSubtitle: () => void;
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

interface ProxyCommandFeedback {
  template: string;
  rawMpvOsd: boolean;
}

function resolveProxyCommandOsdTemplate(command: (string | number)[]): ProxyCommandFeedback | null {
  const operation = typeof command[0] === 'string' ? command[0] : '';
  if (operation === 'sub-step') {
    return { template: 'Subtitle delay: ${sub-delay}', rawMpvOsd: true };
  }

  const property = typeof command[1] === 'string' ? command[1] : '';
  if (!MPV_PROPERTY_COMMANDS.has(operation)) return null;
  if (property === 'sub-pos') {
    return { template: 'Subtitle position: ${sub-pos}', rawMpvOsd: false };
  }
  if (property === 'sid') {
    return { template: 'Subtitle track: ${sid}', rawMpvOsd: false };
  }
  if (property === 'secondary-sid') {
    return { template: 'Secondary subtitle track: ${secondary-sid}', rawMpvOsd: false };
  }
  if (property === 'sub-delay') {
    return { template: 'Subtitle delay: ${sub-delay}', rawMpvOsd: true };
  }
  return null;
}

function showResolvedProxyCommandOsd(
  command: (string | number)[],
  options: HandleMpvCommandFromIpcOptions,
): void {
  const feedback = resolveProxyCommandOsdTemplate(command);
  if (!feedback) return;
  const showFeedback = feedback.rawMpvOsd
    ? (options.showRawMpvOsd ?? options.showMpvOsd)
    : (options.showPlaybackFeedback ?? options.showMpvOsd);

  const emit = async () => {
    try {
      const resolved = await options.resolveProxyCommandOsd?.(command);
      showFeedback(resolved || feedback.template);
    } catch {
      showFeedback(feedback.template);
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

  if (first === options.specialCommands.ANIMETOSHO_OPEN) {
    options.openAnimetosho();
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

  if (first === 'show-text') {
    const message = (typeof command[1] === 'string' ? command[1] : String(command[1] ?? '')).trim();
    if (message) {
      const showFeedback = options.showPlaybackFeedback ?? options.showMpvOsd;
      showFeedback(message);
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
