import type { RuntimeOptionsManager } from '../../runtime-options';
import type { JimakuApiResponse, JimakuLanguagePreference, ResolvedConfig } from '../../types';
import {
  getJimakuLanguagePreference as getJimakuLanguagePreferenceCore,
  getJimakuMaxEntryResults as getJimakuMaxEntryResultsCore,
  isAutoUpdateEnabledRuntime as isAutoUpdateEnabledRuntimeCore,
  jimakuFetchJson as jimakuFetchJsonCore,
  resolveJimakuApiKey as resolveJimakuApiKeyCore,
  shouldAutoInitializeOverlayRuntimeFromConfig as shouldAutoInitializeOverlayRuntimeFromConfigCore,
} from '../../core/services';

export type ConfigDerivedRuntimeDeps = {
  getResolvedConfig: () => ResolvedConfig;
  getRuntimeOptionsManager: () => RuntimeOptionsManager | null;
  defaultJimakuLanguagePreference: JimakuLanguagePreference;
  defaultJimakuMaxEntryResults: number;
  defaultJimakuApiBaseUrl: string;
};

export function createConfigDerivedRuntime(deps: ConfigDerivedRuntimeDeps): {
  shouldAutoInitializeOverlayRuntimeFromConfig: () => boolean;
  isAutoUpdateEnabledRuntime: () => boolean;
  getJimakuLanguagePreference: () => JimakuLanguagePreference;
  getJimakuMaxEntryResults: () => number;
  resolveJimakuApiKey: () => Promise<string | null>;
  jimakuFetchJson: <T>(
    endpoint: string,
    query?: Record<string, string | number | boolean | null | undefined>,
  ) => Promise<JimakuApiResponse<T>>;
} {
  return {
    shouldAutoInitializeOverlayRuntimeFromConfig: () =>
      shouldAutoInitializeOverlayRuntimeFromConfigCore(deps.getResolvedConfig()),
    isAutoUpdateEnabledRuntime: () =>
      isAutoUpdateEnabledRuntimeCore(deps.getResolvedConfig(), deps.getRuntimeOptionsManager()),
    getJimakuLanguagePreference: () =>
      getJimakuLanguagePreferenceCore(
        () => deps.getResolvedConfig(),
        deps.defaultJimakuLanguagePreference,
      ),
    getJimakuMaxEntryResults: () =>
      getJimakuMaxEntryResultsCore(
        () => deps.getResolvedConfig(),
        deps.defaultJimakuMaxEntryResults,
      ),
    resolveJimakuApiKey: () => resolveJimakuApiKeyCore(() => deps.getResolvedConfig()),
    jimakuFetchJson: <T>(
      endpoint: string,
      query: Record<string, string | number | boolean | null | undefined> = {},
    ): Promise<JimakuApiResponse<T>> =>
      jimakuFetchJsonCore<T>(endpoint, query, {
        getResolvedConfig: () => deps.getResolvedConfig(),
        defaultBaseUrl: deps.defaultJimakuApiBaseUrl,
        defaultMaxEntryResults: deps.defaultJimakuMaxEntryResults,
        defaultLanguagePreference: deps.defaultJimakuLanguagePreference,
      }),
  };
}
