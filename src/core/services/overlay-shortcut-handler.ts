import { ConfiguredShortcuts } from '../utils/shortcut-config';
import { OverlayShortcutHandlers } from './overlay-shortcut';
import { createLogger } from '../../logger';

const logger = createLogger('main:overlay-shortcut-handler');

export interface OverlayShortcutFallbackHandlers {
  openRuntimeOptions: () => void;
  openCharacterDictionary: () => void;
  openJimaku: () => void;
  markAudioCard: () => void;
  copySubtitleMultiple: (timeoutMs: number) => void;
  copySubtitle: () => void;
  toggleSecondarySub: () => void;
  updateLastCardFromClipboard: () => void;
  triggerFieldGrouping: () => void;
  triggerSubsync: () => void;
  mineSentence: () => void;
  mineSentenceMultiple: (timeoutMs: number) => void;
}

export interface OverlayShortcutRuntimeDeps {
  showMpvOsd: (text: string) => void;
  openRuntimeOptions: () => void;
  openCharacterDictionary: () => void;
  openJimaku: () => void;
  markAudioCard: () => Promise<void>;
  copySubtitleMultiple: (timeoutMs: number) => void;
  copySubtitle: () => void;
  toggleSecondarySub: () => void;
  updateLastCardFromClipboard: () => Promise<void>;
  triggerFieldGrouping: () => Promise<void>;
  triggerSubsync: () => Promise<void>;
  mineSentence: () => Promise<void>;
  mineSentenceMultiple: (timeoutMs: number) => void;
}

function wrapAsync(
  task: () => Promise<void>,
  deps: OverlayShortcutRuntimeDeps,
  logLabel: string,
  osdLabel: string,
): () => void {
  return () => {
    task().catch((err) => {
      logger.error(`${logLabel} failed:`, err);
      deps.showMpvOsd(`${osdLabel}: ${(err as Error).message}`);
    });
  };
}

export function createOverlayShortcutRuntimeHandlers(deps: OverlayShortcutRuntimeDeps): {
  overlayHandlers: OverlayShortcutHandlers;
  fallbackHandlers: OverlayShortcutFallbackHandlers;
} {
  const overlayHandlers: OverlayShortcutHandlers = {
    copySubtitle: () => {
      deps.copySubtitle();
    },
    copySubtitleMultiple: (timeoutMs) => {
      deps.copySubtitleMultiple(timeoutMs);
    },
    updateLastCardFromClipboard: wrapAsync(
      () => deps.updateLastCardFromClipboard(),
      deps,
      'updateLastCardFromClipboard',
      'Update failed',
    ),
    triggerFieldGrouping: wrapAsync(
      () => deps.triggerFieldGrouping(),
      deps,
      'triggerFieldGrouping',
      'Field grouping failed',
    ),
    triggerSubsync: wrapAsync(
      () => deps.triggerSubsync(),
      deps,
      'triggerSubsyncFromConfig',
      'Subsync failed',
    ),
    mineSentence: wrapAsync(
      () => deps.mineSentence(),
      deps,
      'mineSentenceCard',
      'Mine sentence failed',
    ),
    mineSentenceMultiple: (timeoutMs) => {
      deps.mineSentenceMultiple(timeoutMs);
    },
    toggleSecondarySub: () => deps.toggleSecondarySub(),
    markAudioCard: wrapAsync(
      () => deps.markAudioCard(),
      deps,
      'markLastCardAsAudioCard',
      'Audio card failed',
    ),
    openRuntimeOptions: () => {
      deps.openRuntimeOptions();
    },
    openCharacterDictionary: () => {
      deps.openCharacterDictionary();
    },
    openJimaku: () => {
      deps.openJimaku();
    },
  };

  const fallbackHandlers: OverlayShortcutFallbackHandlers = {
    openRuntimeOptions: overlayHandlers.openRuntimeOptions,
    openCharacterDictionary: overlayHandlers.openCharacterDictionary,
    openJimaku: overlayHandlers.openJimaku,
    markAudioCard: overlayHandlers.markAudioCard,
    copySubtitleMultiple: overlayHandlers.copySubtitleMultiple,
    copySubtitle: overlayHandlers.copySubtitle,
    toggleSecondarySub: overlayHandlers.toggleSecondarySub,
    updateLastCardFromClipboard: overlayHandlers.updateLastCardFromClipboard,
    triggerFieldGrouping: overlayHandlers.triggerFieldGrouping,
    triggerSubsync: overlayHandlers.triggerSubsync,
    mineSentence: overlayHandlers.mineSentence,
    mineSentenceMultiple: overlayHandlers.mineSentenceMultiple,
  };

  return { overlayHandlers, fallbackHandlers };
}

export function runOverlayShortcutLocalFallback(
  input: Electron.Input,
  shortcuts: ConfiguredShortcuts,
  matcher: (input: Electron.Input, accelerator: string, allowWhenRegistered?: boolean) => boolean,
  handlers: OverlayShortcutFallbackHandlers,
): boolean {
  const actions: Array<{
    accelerator: string | null | undefined;
    run: () => void;
    allowWhenRegistered?: boolean;
  }> = [
    {
      accelerator: shortcuts.openRuntimeOptions,
      run: () => {
        handlers.openRuntimeOptions();
      },
    },
    {
      accelerator: shortcuts.openCharacterDictionary,
      run: () => {
        handlers.openCharacterDictionary();
      },
    },
    {
      accelerator: shortcuts.openJimaku,
      run: () => {
        handlers.openJimaku();
      },
      allowWhenRegistered: true,
    },
    {
      accelerator: shortcuts.markAudioCard,
      run: () => {
        handlers.markAudioCard();
      },
    },
    {
      accelerator: shortcuts.copySubtitle,
      run: () => {
        handlers.copySubtitle();
      },
    },
    {
      accelerator: shortcuts.toggleSecondarySub,
      run: () => handlers.toggleSecondarySub(),
      allowWhenRegistered: true,
    },
    {
      accelerator: shortcuts.updateLastCardFromClipboard,
      run: () => {
        handlers.updateLastCardFromClipboard();
      },
    },
    {
      accelerator: shortcuts.triggerFieldGrouping,
      run: () => {
        handlers.triggerFieldGrouping();
      },
    },
    {
      accelerator: shortcuts.triggerSubsync,
      run: () => {
        handlers.triggerSubsync();
      },
    },
    {
      accelerator: shortcuts.mineSentence,
      run: () => {
        handlers.mineSentence();
      },
    },
  ];

  for (const action of actions) {
    if (!action.accelerator) continue;
    if (matcher(input, action.accelerator, action.allowWhenRegistered === true)) {
      action.run();
      return true;
    }
  }

  return false;
}
