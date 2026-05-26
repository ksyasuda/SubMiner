export type CharacterDictionaryManagerNotificationType = 'osd' | 'system' | 'both' | 'none';

export const CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE =
  'Enable character dictionary annotations in Settings to use the character dictionary manager.';

export interface CharacterDictionaryManagerGateDeps {
  isCharacterDictionaryEnabled: () => boolean;
  getNotificationType: () => CharacterDictionaryManagerNotificationType;
  openManager: () => void;
  showOsd: (message: string) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
  logWarn?: (message: string, error?: unknown) => void;
}

function notifyManagerDisabled(deps: CharacterDictionaryManagerGateDeps): void {
  const type = deps.getNotificationType();
  if (type === 'osd' || type === 'both') {
    deps.showOsd(CHARACTER_DICTIONARY_MANAGER_DISABLED_MESSAGE);
  }
  if (type === 'system' || type === 'both') {
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
