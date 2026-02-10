import { ConfiguredShortcuts } from "../utils/shortcut-config";

export interface OverlayShortcutFallbackHandlers {
  openRuntimeOptions: () => void;
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

export function runOverlayShortcutLocalFallback(
  input: Electron.Input,
  shortcuts: ConfiguredShortcuts,
  matcher: (
    input: Electron.Input,
    accelerator: string,
    allowWhenRegistered?: boolean,
  ) => boolean,
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
      accelerator: shortcuts.markAudioCard,
      run: () => {
        handlers.markAudioCard();
      },
    },
    {
      accelerator: shortcuts.copySubtitleMultiple,
      run: () => {
        handlers.copySubtitleMultiple(shortcuts.multiCopyTimeoutMs);
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
    {
      accelerator: shortcuts.mineSentenceMultiple,
      run: () => {
        handlers.mineSentenceMultiple(shortcuts.multiCopyTimeoutMs);
      },
    },
  ];

  for (const action of actions) {
    if (!action.accelerator) continue;
    if (
      matcher(
        input,
        action.accelerator,
        action.allowWhenRegistered === true,
      )
    ) {
      action.run();
      return true;
    }
  }

  return false;
}
