import type { CharacterDictionaryAutoSyncStatusEvent } from './character-dictionary-auto-sync';
import type { StartupOsdSequencerCharacterDictionaryEvent } from './startup-osd-sequencer';

export type CharacterDictionaryAutoSyncNotificationEvent = CharacterDictionaryAutoSyncStatusEvent;

export interface CharacterDictionaryAutoSyncNotificationDeps {
  getNotificationType: () => 'osd' | 'system' | 'both' | 'none' | undefined;
  showOsd: (message: string) => boolean | void;
  showDesktopNotification: (title: string, options: { body?: string }) => void;
  startupOsdSequencer?: {
    notifyCharacterDictionaryStatus: (
      event: StartupOsdSequencerCharacterDictionaryEvent,
    ) => boolean;
  };
}

function shouldShowOsd(type: 'osd' | 'system' | 'both' | 'none' | undefined): boolean {
  return type !== 'none';
}

function shouldFallbackToDesktop(
  type: 'osd' | 'system' | 'both' | 'none' | undefined,
  phase: CharacterDictionaryAutoSyncNotificationEvent['phase'],
): boolean {
  return (
    (type === 'system' || type === 'both') &&
    (phase === 'generating' || phase === 'building' || phase === 'importing')
  );
}

export function notifyCharacterDictionaryAutoSyncStatus(
  event: CharacterDictionaryAutoSyncNotificationEvent,
  deps: CharacterDictionaryAutoSyncNotificationDeps,
): void {
  const type = deps.getNotificationType();
  if (shouldShowOsd(type)) {
    if (deps.startupOsdSequencer) {
      const shown = deps.startupOsdSequencer.notifyCharacterDictionaryStatus({
        phase: event.phase,
        message: event.message,
      });
      if (!shown && shouldFallbackToDesktop(type, event.phase)) {
        deps.showDesktopNotification('SubMiner', { body: event.message });
      }
      return;
    }
    const shown = deps.showOsd(event.message) !== false;
    if (!shown && shouldFallbackToDesktop(type, event.phase)) {
      deps.showDesktopNotification('SubMiner', { body: event.message });
    }
  }
}
