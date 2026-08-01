import { BRIDGE_CONTEXT_KEY } from './types';
import type {
  BridgeAnime,
  BridgeAnimePage,
  BridgeCapabilities,
  BridgeEpisode,
  BridgePreference,
  BridgeSourceDescriptor,
  BridgeVideo,
} from './types';

const EXTENSION_ID_HEADER = 'x-mangatan-extension-id';
const EXTENSION_ID_PATTERN = /^[0-9a-f]{64}$/;

export interface BridgeSource {
  /** Base64-encoded extension APK. */
  apkBase64: string;
  /** Selects one source inside a multi-source (SourceFactory) APK. */
  sourceId?: string;
  preferences?: BridgePreference[];
}

export interface BridgeClientOptions {
  /** Loopback base URL of the running bridge, e.g. `http://127.0.0.1:53112`. */
  baseUrl: string;
  fetchImpl?: typeof fetch;
}

/** The bridge reports extension failures as HTTP 200 with an error body. */
export class BridgeExtensionError extends Error {
  readonly code?: number;
  constructor(message: string, code?: number) {
    super(message);
    this.name = 'BridgeExtensionError';
    this.code = code;
  }
}

/**
 * Client for the M-Extension-Server `/dalvik` RPC endpoint.
 *
 * The server caches uploaded APKs and returns a content hash, letting
 * subsequent calls send that id instead of re-uploading megabytes of base64.
 * A 409 means the cache was evicted, so the APK is resent once.
 */
export class AnimeBridgeClient {
  private readonly baseUrl: string;
  private readonly fetchImpl: typeof fetch;
  private readonly extensionIds = new Map<string, string>();

  constructor(options: BridgeClientOptions) {
    this.baseUrl = options.baseUrl.replace(/\/+$/, '');
    this.fetchImpl = options.fetchImpl ?? fetch;
  }

  async getCapabilities(): Promise<BridgeCapabilities> {
    const response = await this.fetchImpl(`${this.baseUrl}/capabilities`);
    if (!response.ok) {
      throw new Error(`Anime bridge capabilities check failed (${response.status}).`);
    }
    return (await response.json()) as BridgeCapabilities;
  }

  /** True once the bridge is up and reports the features this client needs. */
  async isReady(): Promise<boolean> {
    try {
      const capabilities = await this.getCapabilities();
      return (
        capabilities.mangatanMihonBridge === 1 &&
        capabilities.sourceFactory === true &&
        capabilities.preferenceCallbacks === true
      );
    } catch {
      return false;
    }
  }

  async searchAnime(
    source: BridgeSource,
    query: string,
    page = 1,
    filterList: unknown[] = [],
  ): Promise<BridgeAnimePage> {
    return this.call<BridgeAnimePage>(source, 'getSearchAnime', {
      page,
      search: query,
      filterList,
    });
  }

  /**
   * List the sources an extension APK provides. A single APK may expose many
   * (a SourceFactory), so this is how a package becomes selectable entries.
   */
  async listAnimeSources(source: BridgeSource): Promise<BridgeSourceDescriptor[]> {
    return this.call<BridgeSourceDescriptor[]>(source, 'sourcesAnime', {});
  }

  /** The extension's own settings schema, with current values. */
  async getSourcePreferences(source: BridgeSource): Promise<BridgePreference[]> {
    return this.call<BridgePreference[]>(source, 'preferencesAnime', {});
  }

  /**
   * Commit a preference change. The whole array is sent back with the edited
   * entry, and `changedPreferenceKey` tells the extension which one moved so it
   * can react (the Jellyfin source logs in when the address or password lands).
   * Returns the extension's refreshed schema.
   */
  async setSourcePreference(
    source: BridgeSource,
    changedPreferenceKey: string,
  ): Promise<BridgePreference[]> {
    return this.call<BridgePreference[]>(source, 'setPreferenceAnime', {}, changedPreferenceKey);
  }

  /** Full metadata for one anime: description, cover art, genres, status. */
  async getAnimeDetails(source: BridgeSource, animeUrl: string): Promise<BridgeAnime> {
    return this.call<BridgeAnime>(source, 'getDetailsAnime', {
      animeData: { url: animeUrl },
    });
  }

  async getPopularAnime(source: BridgeSource, page = 1): Promise<BridgeAnimePage> {
    return this.call<BridgeAnimePage>(source, 'getPopularAnime', { page });
  }

  async getEpisodeList(source: BridgeSource, animeUrl: string): Promise<BridgeEpisode[]> {
    return this.call<BridgeEpisode[]>(source, 'getEpisodeList', {
      animeData: { url: animeUrl },
    });
  }

  async getVideoList(source: BridgeSource, episodeUrl: string): Promise<BridgeVideo[]> {
    return this.call<BridgeVideo[]>(source, 'getVideoList', {
      episodeData: { url: episodeUrl },
    });
  }

  private buildPreferences(
    source: BridgeSource,
    changedPreferenceKey?: string,
  ): BridgePreference[] {
    const context: BridgePreference = { key: BRIDGE_CONTEXT_KEY };
    if (source.sourceId !== undefined) context.sourceId = source.sourceId;
    if (changedPreferenceKey !== undefined) context.changedPreferenceKey = changedPreferenceKey;
    return [...(source.preferences ?? []), context];
  }

  private async call<T>(
    source: BridgeSource,
    method: string,
    extras: Record<string, unknown>,
    changedPreferenceKey?: string,
  ): Promise<T> {
    const cacheKey = source.sourceId ?? source.apkBase64;
    const cachedId = this.extensionIds.get(cacheKey);

    let response = await this.post(method, extras, source, cachedId, changedPreferenceKey);
    if (response.status === 409 && cachedId !== undefined) {
      // Server evicted the cached APK; upload it again.
      this.extensionIds.delete(cacheKey);
      response = await this.post(method, extras, source, undefined, changedPreferenceKey);
    }

    if (!response.ok) {
      throw new Error(`Anime bridge ${method} failed (${response.status}).`);
    }

    const returnedId = response.headers.get(EXTENSION_ID_HEADER)?.trim();
    if (returnedId && EXTENSION_ID_PATTERN.test(returnedId)) {
      this.extensionIds.set(cacheKey, returnedId);
    }

    const body = (await response.json()) as T;
    assertNoExtensionError(body, method);
    return body;
  }

  private post(
    method: string,
    extras: Record<string, unknown>,
    source: BridgeSource,
    extensionId: string | undefined,
    changedPreferenceKey?: string,
  ): Promise<Response> {
    const payload: Record<string, unknown> = {
      method,
      ...extras,
      preferences: this.buildPreferences(source, changedPreferenceKey),
      ...(extensionId === undefined ? { data: source.apkBase64 } : { extensionId }),
    };

    return this.fetchImpl(`${this.baseUrl}/dalvik`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        Accept: 'application/json',
      },
      body: JSON.stringify(payload),
    });
  }
}

function assertNoExtensionError(body: unknown, method: string): void {
  if (body === null || typeof body !== 'object' || Array.isArray(body)) return;
  const error = (body as { error?: unknown }).error;
  if (typeof error !== 'string') return;
  const code = (body as { code?: unknown }).code;
  throw new BridgeExtensionError(
    `Anime bridge ${method} failed: ${error}`,
    typeof code === 'number' ? code : undefined,
  );
}
