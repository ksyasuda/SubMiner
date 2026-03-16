import type { CharacterDictionaryAutoSyncStatusEvent } from './character-dictionary-auto-sync';
import type { StartupOsdSequencerCharacterDictionaryEvent } from './startup-osd-sequencer';
import { shouldShowDesktopNotification, shouldShowLogicalOsd } from './overlay-notifications';

export type CharacterDictionaryAutoSyncNotificationEvent = CharacterDictionaryAutoSyncStatusEvent;

export interface CharacterDictionaryAutoSyncNotificationDeps {
  getNotificationType: () => 'osd' | 'system' | 'both' | 'none' | undefined;
  showOsd: (message: string) => void;
  showDesktopNotification: (title: string, options: { body?: string }) => void;
  startupOsdSequencer?: {
    notifyCharacterDictionaryStatus: (event: StartupOsdSequencerCharacterDictionaryEvent) => void;
  };
}

export function notifyCharacterDictionaryAutoSyncStatus(
  event: CharacterDictionaryAutoSyncNotificationEvent,
  deps: CharacterDictionaryAutoSyncNotificationDeps,
): void {
  const type = deps.getNotificationType();
  if (shouldShowLogicalOsd(type)) {
    if (deps.startupOsdSequencer) {
      deps.startupOsdSequencer.notifyCharacterDictionaryStatus({
        phase: event.phase,
        message: event.message,
      });
    } else {
      deps.showOsd(event.message);
    }
  }

  if (shouldShowDesktopNotification(type)) {
    deps.showDesktopNotification('SubMiner', { body: event.message });
  }
}
