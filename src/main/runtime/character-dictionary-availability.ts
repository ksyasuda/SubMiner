export function isCharacterDictionaryRuntimeEnabled(externalProfilePath: string): boolean {
  return externalProfilePath.trim().length === 0;
}

export function getCharacterDictionaryDisabledReason(externalProfilePath: string): string | null {
  if (isCharacterDictionaryRuntimeEnabled(externalProfilePath)) {
    return null;
  }
  return 'Character dictionary is disabled while yomitan.externalProfilePath is configured.';
}
