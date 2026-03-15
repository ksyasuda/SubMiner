import * as path from 'path';

function redactSkippedYomitanWriteValue(
  actionName:
    | 'importYomitanDictionary'
    | 'deleteYomitanDictionary'
    | 'upsertYomitanDictionarySettings',
  rawValue: string,
): string {
  const trimmed = rawValue.trim();
  if (!trimmed) {
    return '<redacted>';
  }

  if (actionName === 'importYomitanDictionary') {
    const basename = path.basename(trimmed);
    return basename || '<redacted>';
  }

  return '<redacted>';
}

export function formatSkippedYomitanWriteAction(
  actionName:
    | 'importYomitanDictionary'
    | 'deleteYomitanDictionary'
    | 'upsertYomitanDictionarySettings',
  rawValue: string,
): string {
  return `${actionName}(${redactSkippedYomitanWriteValue(actionName, rawValue)})`;
}
