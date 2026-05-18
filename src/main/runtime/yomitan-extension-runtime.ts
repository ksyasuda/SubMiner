import {
  createEnsureYomitanExtensionLoadedHandler,
  createLoadYomitanExtensionHandler,
} from './yomitan-extension-loader';
import type { Extension } from 'electron';
import {
  createBuildEnsureYomitanExtensionLoadedMainDepsHandler,
  createBuildLoadYomitanExtensionMainDepsHandler,
} from './yomitan-extension-loader-main-deps';

type LoadYomitanExtensionMainDeps = Parameters<
  typeof createBuildLoadYomitanExtensionMainDepsHandler
>[0];

type EnsureYomitanExtensionLoadedMainDeps = Omit<
  Parameters<typeof createBuildEnsureYomitanExtensionLoadedMainDepsHandler>[0],
  'loadYomitanExtension'
>;

export type YomitanExtensionRuntimeDeps = LoadYomitanExtensionMainDeps &
  EnsureYomitanExtensionLoadedMainDeps & {
    onYomitanExtensionLoaded?: (extension: Extension) => void | Promise<void>;
  };

export function createYomitanExtensionRuntime(deps: YomitanExtensionRuntimeDeps) {
  const buildLoadYomitanExtensionMainDepsHandler = createBuildLoadYomitanExtensionMainDepsHandler({
    loadYomitanExtensionCore: deps.loadYomitanExtensionCore,
    userDataPath: deps.userDataPath,
    externalProfilePath: deps.externalProfilePath,
    getYomitanParserWindow: deps.getYomitanParserWindow,
    setYomitanParserWindow: deps.setYomitanParserWindow,
    setYomitanParserReadyPromise: deps.setYomitanParserReadyPromise,
    setYomitanParserInitPromise: deps.setYomitanParserInitPromise,
    setYomitanExtension: deps.setYomitanExtension,
    setYomitanSession: deps.setYomitanSession,
  });
  const loadYomitanExtensionHandler = createLoadYomitanExtensionHandler(
    buildLoadYomitanExtensionMainDepsHandler(),
  );

  const buildEnsureYomitanExtensionLoadedMainDepsHandler =
    createBuildEnsureYomitanExtensionLoadedMainDepsHandler({
      getYomitanExtension: deps.getYomitanExtension,
      getLoadInFlight: deps.getLoadInFlight,
      setLoadInFlight: deps.setLoadInFlight,
      loadYomitanExtension: () => loadYomitanExtensionHandler(),
    });
  const ensureYomitanExtensionLoadedHandler = createEnsureYomitanExtensionLoadedHandler(
    buildEnsureYomitanExtensionLoadedMainDepsHandler(),
  );

  let lastNotifiedExtension: Extension | null = null;
  async function notifyYomitanExtensionLoaded(extension: Extension | null): Promise<void> {
    if (!extension || extension === lastNotifiedExtension) {
      return;
    }
    lastNotifiedExtension = extension;
    await deps.onYomitanExtensionLoaded?.(extension);
  }

  return {
    loadYomitanExtension: async (): Promise<ReturnType<typeof deps.getYomitanExtension>> => {
      const extension = await loadYomitanExtensionHandler();
      await notifyYomitanExtensionLoaded(extension);
      return extension;
    },
    ensureYomitanExtensionLoaded: async (): Promise<
      ReturnType<typeof deps.getYomitanExtension>
    > => {
      const extension = await ensureYomitanExtensionLoadedHandler();
      await notifyYomitanExtensionLoaded(extension);
      return extension;
    },
  };
}
