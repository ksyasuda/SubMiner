import {
  JimakuApiResponse,
  JimakuConfig,
  JimakuLanguagePreference,
} from "../../types";
import {
  jimakuFetchJson as jimakuFetchJsonRequest,
  resolveJimakuApiKey as resolveJimakuApiKeyFromConfig,
} from "../../jimaku/utils";

export function getJimakuConfig(
  getResolvedConfig: () => { jimaku?: JimakuConfig },
): JimakuConfig {
  const config = getResolvedConfig();
  return config.jimaku ?? {};
}

export function getJimakuBaseUrl(
  getResolvedConfig: () => { jimaku?: JimakuConfig },
  defaultBaseUrl: string,
): string {
  const config = getJimakuConfig(getResolvedConfig);
  return config.apiBaseUrl || defaultBaseUrl;
}

export function getJimakuLanguagePreference(
  getResolvedConfig: () => { jimaku?: JimakuConfig },
  defaultPreference: JimakuLanguagePreference,
): JimakuLanguagePreference {
  const config = getJimakuConfig(getResolvedConfig);
  return config.languagePreference || defaultPreference;
}

export function getJimakuMaxEntryResults(
  getResolvedConfig: () => { jimaku?: JimakuConfig },
  defaultValue: number,
): number {
  const config = getJimakuConfig(getResolvedConfig);
  const value = config.maxEntryResults;
  if (typeof value === "number" && Number.isFinite(value) && value > 0) {
    return Math.floor(value);
  }
  return defaultValue;
}

export async function resolveJimakuApiKey(
  getResolvedConfig: () => { jimaku?: JimakuConfig },
): Promise<string | null> {
  return resolveJimakuApiKeyFromConfig(getJimakuConfig(getResolvedConfig));
}

export async function jimakuFetchJson<T>(
  endpoint: string,
  query: Record<string, string | number | boolean | null | undefined> = {},
  options: {
    getResolvedConfig: () => { jimaku?: JimakuConfig };
    defaultBaseUrl: string;
    defaultMaxEntryResults: number;
    defaultLanguagePreference: JimakuLanguagePreference;
  },
): Promise<JimakuApiResponse<T>> {
  const apiKey = await resolveJimakuApiKey(options.getResolvedConfig);
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
    baseUrl: getJimakuBaseUrl(
      options.getResolvedConfig,
      options.defaultBaseUrl,
    ),
    apiKey,
  });
}
