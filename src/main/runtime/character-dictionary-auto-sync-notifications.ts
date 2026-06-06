import type { CharacterDictionaryAutoSyncStatusEvent } from './character-dictionary-auto-sync';
import type { StartupOsdSequencerCharacterDictionaryEvent } from './startup-osd-sequencer';
import type { NotificationType, OverlayNotificationPayload } from '../../types/notification';
import { shouldShowDesktop, shouldShowOverlay, shouldShowOsd } from './notification-routing';

export type CharacterDictionaryAutoSyncNotificationEvent = CharacterDictionaryAutoSyncStatusEvent;

export interface CharacterDictionaryAutoSyncNotificationDeps {
  getNotificationType: () => NotificationType | undefined;
  showOsd: (message: string) => boolean | void;
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
  showDesktopNotification: (title: string, options: { body?: string }) => void;
  startupOsdSequencer?: {
    notifyCharacterDictionaryStatus: (
      event: StartupOsdSequencerCharacterDictionaryEvent,
    ) => boolean;
  };
}

function isTerminalPhase(phase: CharacterDictionaryAutoSyncNotificationEvent['phase']): boolean {
  return phase === 'ready' || phase === 'failed';
}

function overlayVariantForPhase(
  phase: CharacterDictionaryAutoSyncNotificationEvent['phase'],
): OverlayNotificationPayload['variant'] {
  if (phase === 'ready') return 'success';
  if (phase === 'failed') return 'error';
  return 'progress';
}

export function notifyCharacterDictionaryAutoSyncStatus(
  event: CharacterDictionaryAutoSyncNotificationEvent,
  deps: CharacterDictionaryAutoSyncNotificationDeps,
): void {
  const type = deps.getNotificationType() ?? 'overlay';
  if (type === 'none') return;
  let overlayShown = false;
  let startupSequencerShown = false;

  if (shouldShowOverlay(type)) {
    if (deps.showOverlayNotification) {
      deps.showOverlayNotification({
        id: 'character-dictionary-auto-sync',
        title: 'Character dictionary',
        body: event.message,
        variant: overlayVariantForPhase(event.phase),
        persistent: !isTerminalPhase(event.phase),
      });
      overlayShown = true;
    } else if (!shouldShowDesktop(type)) {
      deps.showDesktopNotification('SubMiner', { body: event.message });
    }
  }

  if (shouldShowOsd(type)) {
    if (deps.startupOsdSequencer) {
      startupSequencerShown = deps.startupOsdSequencer.notifyCharacterDictionaryStatus({
        phase: event.phase,
        message: event.message,
      });
    } else {
      deps.showOsd(event.message);
    }
  }

  if (shouldShowDesktop(type) && !overlayShown && !startupSequencerShown) {
    deps.showDesktopNotification('SubMiner', { body: event.message });
  }
}
