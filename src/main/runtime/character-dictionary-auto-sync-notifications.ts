import type { CharacterDictionaryAutoSyncStatusEvent } from './character-dictionary-auto-sync';
import type { StartupOsdSequencerCharacterDictionaryEvent } from './startup-osd-sequencer';

export type CharacterDictionaryAutoSyncNotificationEvent = CharacterDictionaryAutoSyncStatusEvent;

export interface CharacterDictionaryAutoSyncNotificationDeps {
  getNotificationType: () => 'osd' | 'system' | 'both' | 'none' | undefined;
  showOsd: (message: string) => void;
  showDesktopNotification: (title: string, options: { body?: string }) => void;
  startupOsdSequencer?: {
    notifyCharacterDictionaryStatus: (event: StartupOsdSequencerCharacterDictionaryEvent) => void;
  };
}

function shouldShowOsd(type: 'osd' | 'system' | 'both' | 'none' | undefined): boolean {
  return type !== 'none';
}

export function notifyCharacterDictionaryAutoSyncStatus(
  event: CharacterDictionaryAutoSyncNotificationEvent,
  deps: CharacterDictionaryAutoSyncNotificationDeps,
): void {
  const type = deps.getNotificationType();
  if (shouldShowOsd(type)) {
    if (deps.startupOsdSequencer) {
      deps.startupOsdSequencer.notifyCharacterDictionaryStatus({
        phase: event.phase,
        message: event.message,
      });
      return;
    }
    deps.showOsd(event.message);
  }
}
