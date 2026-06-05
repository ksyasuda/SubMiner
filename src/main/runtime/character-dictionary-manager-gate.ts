import type { NotificationType, OverlayNotificationPayload } from '../../types/notification';

export type CharacterDictionaryManagerNotificationType = NotificationType;

export const CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE =
  'Enable Name Match in Settings to use the character dictionary manager.';

export interface CharacterDictionaryManagerGateDeps {
  isCharacterDictionaryEnabled: () => boolean;
  getNotificationType: () => CharacterDictionaryManagerNotificationType;
  openManager: () => void;
  showOsd: (message: string) => void;
  showOverlayNotification?: (payload: OverlayNotificationPayload) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
  logWarn?: (message: string, error?: unknown) => void;
}

function notifyManagerDisabled(deps: CharacterDictionaryManagerGateDeps): void {
  const type = deps.getNotificationType();
  if (type === 'none') {
    return;
  }
  if (type === 'overlay' || type === 'both') {
    deps.showOverlayNotification?.({
      title: 'SubMiner',
      body: CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE,
      variant: 'warning',
    });
  }
  if (type === 'osd' || type === 'osd-system') {
    deps.showOsd(CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE);
  }
  if (type === 'system' || type === 'both' || type === 'osd-system') {
    try {
      deps.showDesktopNotification('SubMiner', {
        body: CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE,
      });
    } catch (error) {
      deps.logWarn?.('Unable to show character dictionary manager notification.', error);
    }
  }
}

export function openCharacterDictionaryManagerWithConfigGate(
  deps: CharacterDictionaryManagerGateDeps,
): void {
  if (deps.isCharacterDictionaryEnabled()) {
    deps.openManager();
    return;
  }
  notifyManagerDisabled(deps);
}
