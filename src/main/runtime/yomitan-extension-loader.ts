import type { Extension } from 'electron';
import type { YomitanExtensionLoaderDeps } from '../../core/services/yomitan-extension-loader';

export function createLoadYomitanExtensionHandler(deps: {
  loadYomitanExtensionCore: (options: YomitanExtensionLoaderDeps) => Promise<Extension | null>;
  userDataPath: YomitanExtensionLoaderDeps['userDataPath'];
  externalProfilePath?: YomitanExtensionLoaderDeps['externalProfilePath'];
  getYomitanParserWindow: YomitanExtensionLoaderDeps['getYomitanParserWindow'];
  setYomitanParserWindow: YomitanExtensionLoaderDeps['setYomitanParserWindow'];
  setYomitanParserReadyPromise: YomitanExtensionLoaderDeps['setYomitanParserReadyPromise'];
  setYomitanParserInitPromise: YomitanExtensionLoaderDeps['setYomitanParserInitPromise'];
  setYomitanExtension: YomitanExtensionLoaderDeps['setYomitanExtension'];
  setYomitanSession: YomitanExtensionLoaderDeps['setYomitanSession'];
}) {
  return async (): Promise<Extension | null> => {
    return deps.loadYomitanExtensionCore({
      userDataPath: deps.userDataPath,
      externalProfilePath: deps.externalProfilePath,
      getYomitanParserWindow: deps.getYomitanParserWindow,
      setYomitanParserWindow: deps.setYomitanParserWindow,
      setYomitanParserReadyPromise: deps.setYomitanParserReadyPromise,
      setYomitanParserInitPromise: deps.setYomitanParserInitPromise,
      setYomitanExtension: deps.setYomitanExtension,
      setYomitanSession: deps.setYomitanSession,
    });
  };
}

export function createEnsureYomitanExtensionLoadedHandler(deps: {
  getYomitanExtension: () => Extension | null;
  getLoadInFlight: () => Promise<Extension | null> | null;
  setLoadInFlight: (promise: Promise<Extension | null> | null) => void;
  loadYomitanExtension: () => Promise<Extension | null>;
}) {
  return async (): Promise<Extension | null> => {
    const existing = deps.getYomitanExtension();
    if (existing) {
      return existing;
    }

    const inFlight = deps.getLoadInFlight();
    if (inFlight) {
      return inFlight;
    }

    const promise = deps.loadYomitanExtension().finally(() => {
      deps.setLoadInFlight(null);
    });
    deps.setLoadInFlight(promise);
    return promise;
  };
}
