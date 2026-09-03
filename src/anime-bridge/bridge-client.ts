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
  /**
   * Identity of the APK's contents. Keys the extension-id cache, so an upgraded
   * APK is re-uploaded instead of reusing the previous build's id.
   */
  fingerprint: string;
  /**
   * Reads and base64-encodes the APK. Called only when the bridge actually
   * needs the bytes, so multi-megabyte payloads are not held on the heap.
   */
  loadApkBase64: () => Promise<string>;
  /** Selects one source inside a multi-source (SourceFactory) APK. */
  sourceId?: string;
  preferences?: BridgePreference[];
}

export interface BridgeClientOptions {
  /** Loopback base URL of the running bridge, e.g. `http://127.0.0.1:53112`. */
  baseUrl: string;
  fetchImpl?: typeof fetch;
  /**
   * Per-request deadline. Node's `fetch` has none, so a sidecar that accepts
   * the socket and then stalls would leave every call pending forever.
   */
  requestTimeoutMs?: number;
}

/** Extension calls can be slow (a source may scrape several pages). */
const DEFAULT_REQUEST_TIMEOUT_MS = 60_000;

/** The readiness probe is a local health check; it should answer at once. */
const CAPABILITIES_TIMEOUT_MS = 5_000;

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
  private readonly requestTimeoutMs: number;
  private readonly extensionIds = new Map<string, string>();

  constructor(options: BridgeClientOptions) {
    this.baseUrl = options.baseUrl.replace(/\/+$/, '');
    this.fetchImpl = options.fetchImpl ?? fetch;
    this.requestTimeoutMs = options.requestTimeoutMs ?? DEFAULT_REQUEST_TIMEOUT_MS;
  }

  /**
   * `timeoutMs` lets a caller with its own deadline (the readiness loop) cap the
   * probe below the default, so a short readiness budget is actually honored.
   */
  async getCapabilities(timeoutMs = CAPABILITIES_TIMEOUT_MS): Promise<BridgeCapabilities> {
    const response = await this.fetchImpl(`${this.baseUrl}/capabilities`, {
      signal: AbortSignal.timeout(Math.max(0, Math.min(timeoutMs, this.requestTimeoutMs))),
    });
    if (!response.ok) {
      throw new Error(`Anime bridge capabilities check failed (${response.status}).`);
    }
    return (await response.json()) as BridgeCapabilities;
  }

  /** True once the bridge is up and reports the features this client needs. */
  async isReady(timeoutMs?: number): Promise<boolean> {
    try {
      const capabilities = await this.getCapabilities(timeoutMs);
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
    // Keyed by APK contents, not by source id: an in-place upgrade keeps the
    // same source id, and reusing its cached extension id would silently run
    // the previous build (the bridge has no reason to answer 409).
    const cacheKey = `${source.fingerprint}:${source.sourceId ?? ''}`;
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

  private async post(
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
      ...(extensionId === undefined ? { data: await source.loadApkBase64() } : { extensionId }),
    };

    return this.fetchImpl(`${this.baseUrl}/dalvik`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8',
        Accept: 'application/json',
      },
      body: JSON.stringify(payload),
      signal: AbortSignal.timeout(this.requestTimeoutMs),
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
