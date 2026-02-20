import { createEnsureYomitanExtensionLoadedHandler, createLoadYomitanExtensionHandler } from './yomitan-extension-loader';
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
  EnsureYomitanExtensionLoadedMainDeps;

export function createYomitanExtensionRuntime(deps: YomitanExtensionRuntimeDeps) {
  const buildLoadYomitanExtensionMainDepsHandler = createBuildLoadYomitanExtensionMainDepsHandler({
    loadYomitanExtensionCore: deps.loadYomitanExtensionCore,
    userDataPath: deps.userDataPath,
    getYomitanParserWindow: deps.getYomitanParserWindow,
    setYomitanParserWindow: deps.setYomitanParserWindow,
    setYomitanParserReadyPromise: deps.setYomitanParserReadyPromise,
    setYomitanParserInitPromise: deps.setYomitanParserInitPromise,
    setYomitanExtension: deps.setYomitanExtension,
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

  return {
    loadYomitanExtension: (): Promise<ReturnType<typeof deps.getYomitanExtension>> =>
      loadYomitanExtensionHandler(),
    ensureYomitanExtensionLoaded: (): Promise<ReturnType<typeof deps.getYomitanExtension>> =>
      ensureYomitanExtensionLoadedHandler(),
  };
}
