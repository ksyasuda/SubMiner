import { createBuildOpenYomitanSettingsMainDepsHandler } from './app-runtime-main-deps';
import { createOpenYomitanSettingsHandler } from './yomitan-settings-opener';

type OpenYomitanSettingsMainDeps = Parameters<
  typeof createBuildOpenYomitanSettingsMainDepsHandler
>[0];

export function createYomitanSettingsRuntime(deps: OpenYomitanSettingsMainDeps) {
  const openYomitanSettingsMainDeps = createBuildOpenYomitanSettingsMainDepsHandler(deps)();
  const openYomitanSettings = createOpenYomitanSettingsHandler(openYomitanSettingsMainDeps);

  return {
    openYomitanSettings,
  };
}
