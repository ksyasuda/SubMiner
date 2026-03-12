import type {
  createEnsureYomitanExtensionLoadedHandler,
  createLoadYomitanExtensionHandler,
} from './yomitan-extension-loader';

type LoadYomitanExtensionMainDeps = Parameters<typeof createLoadYomitanExtensionHandler>[0];
type EnsureYomitanExtensionLoadedMainDeps = Parameters<
  typeof createEnsureYomitanExtensionLoadedHandler
>[0];

export function createBuildLoadYomitanExtensionMainDepsHandler(deps: LoadYomitanExtensionMainDeps) {
  return (): LoadYomitanExtensionMainDeps => ({
    loadYomitanExtensionCore: (options) => deps.loadYomitanExtensionCore(options),
    userDataPath: deps.userDataPath,
    externalProfilePath: deps.externalProfilePath,
    getYomitanParserWindow: () => deps.getYomitanParserWindow(),
    setYomitanParserWindow: (window) => deps.setYomitanParserWindow(window),
    setYomitanParserReadyPromise: (promise) => deps.setYomitanParserReadyPromise(promise),
    setYomitanParserInitPromise: (promise) => deps.setYomitanParserInitPromise(promise),
    setYomitanExtension: (extension) => deps.setYomitanExtension(extension),
    setYomitanSession: (session) => deps.setYomitanSession(session),
  });
}

export function createBuildEnsureYomitanExtensionLoadedMainDepsHandler(
  deps: EnsureYomitanExtensionLoadedMainDeps,
) {
  return (): EnsureYomitanExtensionLoadedMainDeps => ({
    getYomitanExtension: () => deps.getYomitanExtension(),
    getLoadInFlight: () => deps.getLoadInFlight(),
    setLoadInFlight: (promise) => deps.setLoadInFlight(promise),
    loadYomitanExtension: () => deps.loadYomitanExtension(),
  });
}
