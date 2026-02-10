import {
  JimakuApiResponse,
  JimakuConfig,
  JimakuLanguagePreference,
} from "../../types";
import {
  jimakuFetchJson as jimakuFetchJsonRequest,
  resolveJimakuApiKey as resolveJimakuApiKeyFromConfig,
} from "../../jimaku/utils";

export function getJimakuConfigService(
  getResolvedConfig: () => { jimaku?: JimakuConfig },
): JimakuConfig {
  const config = getResolvedConfig();
  return config.jimaku ?? {};
}

export function getJimakuBaseUrlService(
  getResolvedConfig: () => { jimaku?: JimakuConfig },
  defaultBaseUrl: string,
): string {
  const config = getJimakuConfigService(getResolvedConfig);
  return config.apiBaseUrl || defaultBaseUrl;
}

export function getJimakuLanguagePreferenceService(
  getResolvedConfig: () => { jimaku?: JimakuConfig },
  defaultPreference: JimakuLanguagePreference,
): JimakuLanguagePreference {
  const config = getJimakuConfigService(getResolvedConfig);
  return config.languagePreference || defaultPreference;
}

export function getJimakuMaxEntryResultsService(
  getResolvedConfig: () => { jimaku?: JimakuConfig },
  defaultValue: number,
): number {
  const config = getJimakuConfigService(getResolvedConfig);
  const value = config.maxEntryResults;
  if (typeof value === "number" && Number.isFinite(value) && value > 0) {
    return Math.floor(value);
  }
  return defaultValue;
}

export async function resolveJimakuApiKeyService(
  getResolvedConfig: () => { jimaku?: JimakuConfig },
): Promise<string | null> {
  return resolveJimakuApiKeyFromConfig(getJimakuConfigService(getResolvedConfig));
}

export async function jimakuFetchJsonService<T>(
  endpoint: string,
  query: Record<string, string | number | boolean | null | undefined> = {},
  options: {
    getResolvedConfig: () => { jimaku?: JimakuConfig };
    defaultBaseUrl: string;
    defaultMaxEntryResults: number;
    defaultLanguagePreference: JimakuLanguagePreference;
  },
): Promise<JimakuApiResponse<T>> {
  const apiKey = await resolveJimakuApiKeyService(options.getResolvedConfig);
  if (!apiKey) {
    return {
      ok: false,
      error: {
        error:
          "Jimaku API key not set. Configure jimaku.apiKey or jimaku.apiKeyCommand.",
        code: 401,
      },
    };
  }

  return jimakuFetchJsonRequest<T>(endpoint, query, {
    baseUrl: getJimakuBaseUrlService(
      options.getResolvedConfig,
      options.defaultBaseUrl,
    ),
    apiKey,
  });
}
