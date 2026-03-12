import {
  getCharacterDictionaryDisabledReason,
  isCharacterDictionaryRuntimeEnabled,
} from './character-dictionary-availability';

export function createYomitanProfilePolicy(options: {
  externalProfilePath: string;
  logInfo: (message: string) => void;
}) {
  const externalProfilePath = options.externalProfilePath.trim();

  return {
    externalProfilePath,
    isExternalReadOnlyMode: (): boolean => externalProfilePath.length > 0,
    isCharacterDictionaryEnabled: (): boolean =>
      isCharacterDictionaryRuntimeEnabled(externalProfilePath),
    getCharacterDictionaryDisabledReason: (): string | null =>
      getCharacterDictionaryDisabledReason(externalProfilePath),
    logSkippedWrite: (action: string): void => {
      options.logInfo(
        `[yomitan] skipping ${action}: yomitan.externalProfilePath is configured; external profile mode is read-only`,
      );
    },
  };
}
