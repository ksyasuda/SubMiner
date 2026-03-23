import { DEFAULT_ANKI_CONNECT_CONFIG } from '../config';
import type { AnkiConnectConfig } from '../types';
import {
  getKnownWordCacheLifecycleConfig,
  getKnownWordCacheRefreshIntervalMinutes,
  getKnownWordCacheScopeForConfig,
} from './known-word-cache';

export interface AnkiIntegrationRuntimeProxyServer {
  start(options: { host: string; port: number; upstreamUrl: string }): void;
  stop(): void;
  waitUntilReady(): Promise<void>;
}

interface AnkiIntegrationRuntimeDeps {
  initialConfig: AnkiConnectConfig;
  pollingRunner: {
    start(): void;
    stop(): void;
  };
  knownWordCache: {
    startLifecycle(): void;
    stopLifecycle(): void;
    clearKnownWordCacheState(): void;
  };
  proxyServerFactory: () => AnkiIntegrationRuntimeProxyServer;
  logInfo: (message: string, ...args: unknown[]) => void;
  logWarn: (message: string, ...args: unknown[]) => void;
  logError: (message: string, ...args: unknown[]) => void;
  onConfigChanged?: (config: AnkiConnectConfig) => void;
}

function trimToNonEmptyString(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

function normalizeAnkiAiConfig(
  config: AnkiConnectConfig['ai'],
): NonNullable<AnkiConnectConfig['ai']> {
  if (config && typeof config === 'object') {
    return {
      enabled: config.enabled === true,
      model: trimToNonEmptyString(config.model) ?? '',
      systemPrompt: trimToNonEmptyString(config.systemPrompt) ?? '',
    };
  }

  return {
    enabled: config === true,
    model: '',
    systemPrompt: '',
  };
}

export function normalizeAnkiIntegrationConfig(config: AnkiConnectConfig): AnkiConnectConfig {
  const resolvedUrl = trimToNonEmptyString(config.url) ?? DEFAULT_ANKI_CONNECT_CONFIG.url;
  const proxySource =
    config.proxy && typeof config.proxy === 'object'
      ? (config.proxy as NonNullable<AnkiConnectConfig['proxy']>)
      : {};
  const normalizedProxyPort =
    typeof proxySource.port === 'number' &&
    Number.isInteger(proxySource.port) &&
    proxySource.port >= 1 &&
    proxySource.port <= 65535
      ? proxySource.port
      : DEFAULT_ANKI_CONNECT_CONFIG.proxy?.port;
  const normalizedProxyHost =
    trimToNonEmptyString(proxySource.host) ?? DEFAULT_ANKI_CONNECT_CONFIG.proxy?.host;
  const normalizedProxyUpstreamUrl = trimToNonEmptyString(proxySource.upstreamUrl) ?? resolvedUrl;

  return {
    ...DEFAULT_ANKI_CONNECT_CONFIG,
    ...config,
    url: resolvedUrl,
    fields: {
      ...DEFAULT_ANKI_CONNECT_CONFIG.fields,
      ...(config.fields ?? {}),
    },
    proxy: {
      ...DEFAULT_ANKI_CONNECT_CONFIG.proxy,
      ...(config.proxy ?? {}),
      enabled: proxySource.enabled === true,
      host: normalizedProxyHost,
      port: normalizedProxyPort,
      upstreamUrl: normalizedProxyUpstreamUrl,
    },
    ai: normalizeAnkiAiConfig(config.ai),
    media: {
      ...DEFAULT_ANKI_CONNECT_CONFIG.media,
      ...(config.media ?? {}),
    },
    knownWords: {
      ...DEFAULT_ANKI_CONNECT_CONFIG.knownWords,
      ...(config.knownWords ?? {}),
    },
    nPlusOne: {
      ...DEFAULT_ANKI_CONNECT_CONFIG.nPlusOne,
      ...(config.nPlusOne ?? {}),
    },
    behavior: {
      ...DEFAULT_ANKI_CONNECT_CONFIG.behavior,
      ...(config.behavior ?? {}),
    },
    metadata: {
      ...DEFAULT_ANKI_CONNECT_CONFIG.metadata,
      ...(config.metadata ?? {}),
    },
    isLapis: {
      ...DEFAULT_ANKI_CONNECT_CONFIG.isLapis,
      ...(config.isLapis ?? {}),
    },
    isKiku: {
      ...DEFAULT_ANKI_CONNECT_CONFIG.isKiku,
      ...(config.isKiku ?? {}),
    },
  } as AnkiConnectConfig;
}

export class AnkiIntegrationRuntime {
  private config: AnkiConnectConfig;
  private proxyServer: AnkiIntegrationRuntimeProxyServer | null = null;
  private started = false;

  constructor(private readonly deps: AnkiIntegrationRuntimeDeps) {
    this.config = normalizeAnkiIntegrationConfig(deps.initialConfig);
  }

  getConfig(): AnkiConnectConfig {
    return this.config;
  }

  waitUntilReady(): Promise<void> {
    if (!this.started || !this.isProxyTransportEnabled()) {
      return Promise.resolve();
    }
    return this.getOrCreateProxyServer().waitUntilReady();
  }

  start(): void {
    if (this.started) {
      this.stop();
    }

    this.deps.knownWordCache.startLifecycle();
    this.startTransport();
    this.started = true;
  }

  stop(): void {
    this.stopTransport();
    this.deps.knownWordCache.stopLifecycle();
    this.started = false;
    this.deps.logInfo('Stopped AnkiConnect integration');
  }

  applyRuntimeConfigPatch(patch: Partial<AnkiConnectConfig>): void {
    const wasKnownWordCacheEnabled = this.config.knownWords?.highlightEnabled === true;
    const previousKnownWordCacheConfig = wasKnownWordCacheEnabled
      ? this.getKnownWordCacheLifecycleConfig(this.config)
      : null;
    const previousTransportKey = this.getTransportConfigKey(this.config);

    const mergedConfig: AnkiConnectConfig = {
      ...this.config,
      ...patch,
      knownWords:
        patch.knownWords !== undefined
          ? {
              ...(this.config.knownWords ?? DEFAULT_ANKI_CONNECT_CONFIG.knownWords),
              ...patch.knownWords,
            }
          : this.config.knownWords,
      nPlusOne:
        patch.nPlusOne !== undefined
          ? {
              ...(this.config.nPlusOne ?? DEFAULT_ANKI_CONNECT_CONFIG.nPlusOne),
              ...patch.nPlusOne,
            }
          : this.config.nPlusOne,
      fields:
        patch.fields !== undefined
          ? { ...this.config.fields, ...patch.fields }
          : this.config.fields,
      media:
        patch.media !== undefined ? { ...this.config.media, ...patch.media } : this.config.media,
      behavior:
        patch.behavior !== undefined
          ? { ...this.config.behavior, ...patch.behavior }
          : this.config.behavior,
      proxy:
        patch.proxy !== undefined ? { ...this.config.proxy, ...patch.proxy } : this.config.proxy,
      metadata:
        patch.metadata !== undefined
          ? { ...this.config.metadata, ...patch.metadata }
          : this.config.metadata,
      isLapis:
        patch.isLapis !== undefined
          ? { ...this.config.isLapis, ...patch.isLapis }
          : this.config.isLapis,
      isKiku:
        patch.isKiku !== undefined
          ? { ...this.config.isKiku, ...patch.isKiku }
          : this.config.isKiku,
    };
    this.config = normalizeAnkiIntegrationConfig(mergedConfig);
    this.deps.onConfigChanged?.(this.config);
    const nextKnownWordCacheEnabled = this.config.knownWords?.highlightEnabled === true;

    if (wasKnownWordCacheEnabled && !nextKnownWordCacheEnabled) {
      if (this.started) {
        this.deps.knownWordCache.stopLifecycle();
      }
      this.deps.knownWordCache.clearKnownWordCacheState();
    } else if (this.started && !wasKnownWordCacheEnabled && nextKnownWordCacheEnabled) {
      this.deps.knownWordCache.startLifecycle();
    } else if (
      this.started &&
      wasKnownWordCacheEnabled &&
      nextKnownWordCacheEnabled &&
      previousKnownWordCacheConfig !== null &&
      previousKnownWordCacheConfig !== this.getKnownWordCacheLifecycleConfig(this.config)
    ) {
      this.deps.knownWordCache.startLifecycle();
    }

    const nextTransportKey = this.getTransportConfigKey(this.config);
    if (this.started && previousTransportKey !== nextTransportKey) {
      this.stopTransport();
      this.startTransport();
    }
  }

  private getKnownWordCacheLifecycleConfig(config: AnkiConnectConfig): string {
    return getKnownWordCacheLifecycleConfig(config);
  }

  private getKnownWordRefreshIntervalMinutes(config: AnkiConnectConfig): number {
    return getKnownWordCacheRefreshIntervalMinutes(config);
  }

  private getKnownWordCacheScopeForConfig(config: AnkiConnectConfig): string {
    return getKnownWordCacheScopeForConfig(config);
  }

  getOrCreateProxyServer(): AnkiIntegrationRuntimeProxyServer {
    if (!this.proxyServer) {
      this.proxyServer = this.deps.proxyServerFactory();
    }
    return this.proxyServer;
  }

  private isProxyTransportEnabled(config: AnkiConnectConfig = this.config): boolean {
    return config.proxy?.enabled === true;
  }

  private getTransportConfigKey(config: AnkiConnectConfig = this.config): string {
    if (this.isProxyTransportEnabled(config)) {
      return [
        'proxy',
        config.proxy?.host ?? '',
        String(config.proxy?.port ?? ''),
        config.proxy?.upstreamUrl ?? '',
      ].join(':');
    }
    return ['polling', String(config.pollingRate ?? DEFAULT_ANKI_CONNECT_CONFIG.pollingRate)].join(
      ':',
    );
  }

  private startTransport(): void {
    if (this.isProxyTransportEnabled()) {
      const proxyHost = this.config.proxy?.host ?? '127.0.0.1';
      const proxyPort = this.config.proxy?.port ?? 8766;
      const upstreamUrl = this.config.proxy?.upstreamUrl ?? this.config.url ?? '';
      this.getOrCreateProxyServer().start({
        host: proxyHost,
        port: proxyPort,
        upstreamUrl,
      });
      this.deps.logInfo(
        `Starting AnkiConnect integration with local proxy: http://${proxyHost}:${proxyPort} -> ${upstreamUrl}`,
      );
      return;
    }

    this.deps.logInfo(
      'Starting AnkiConnect integration with polling rate:',
      this.config.pollingRate,
    );
    this.deps.pollingRunner.start();
  }

  private stopTransport(): void {
    this.deps.pollingRunner.stop();
    this.proxyServer?.stop();
  }
}
