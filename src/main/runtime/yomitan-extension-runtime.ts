import {
  createEnsureYomitanExtensionLoadedHandler,
  createLoadYomitanExtensionHandler,
} from './yomitan-extension-loader';
import type { Extension } from 'electron';

type LoadYomitanExtensionMainDeps = Parameters<typeof createLoadYomitanExtensionHandler>[0];

type EnsureYomitanExtensionLoadedMainDeps = Omit<
  Parameters<typeof createEnsureYomitanExtensionLoadedHandler>[0],
  'loadYomitanExtension'
>;

export type YomitanExtensionRuntimeDeps = LoadYomitanExtensionMainDeps &
  EnsureYomitanExtensionLoadedMainDeps & {
    onYomitanExtensionLoaded?: (extension: Extension) => void | Promise<void>;
  };

export function createYomitanExtensionRuntime(deps: YomitanExtensionRuntimeDeps) {
  const loadYomitanExtensionHandler = createLoadYomitanExtensionHandler({
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

  const ensureYomitanExtensionLoadedHandler = createEnsureYomitanExtensionLoadedHandler({
    getYomitanExtension: deps.getYomitanExtension,
    getLoadInFlight: deps.getLoadInFlight,
    setLoadInFlight: deps.setLoadInFlight,
    loadYomitanExtension: () => loadYomitanExtensionHandler(),
  });

  let lastNotifiedExtension: Extension | null = null;
  let notifyingExtension: Extension | null = null;
  let notificationPromise: Promise<void> | null = null;
  async function notifyYomitanExtensionLoaded(extension: Extension | null): Promise<void> {
    if (!extension || extension === lastNotifiedExtension) {
      return;
    }
    if (extension === notifyingExtension && notificationPromise) {
      await notificationPromise;
      return;
    }
    notifyingExtension = extension;
    notificationPromise = (async () => {
      await deps.onYomitanExtensionLoaded?.(extension);
      lastNotifiedExtension = extension;
    })();
    try {
      await notificationPromise;
    } finally {
      if (notifyingExtension === extension) {
        notifyingExtension = null;
        notificationPromise = null;
      }
    }
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
