import { createOpenYomitanSettingsHandler } from './yomitan-settings-opener';

export function createYomitanSettingsRuntime(
  deps: Parameters<typeof createOpenYomitanSettingsHandler>[0],
) {
  return {
    openYomitanSettings: createOpenYomitanSettingsHandler(deps),
  };
}
