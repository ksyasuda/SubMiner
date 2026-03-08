import { ResolvedConfig } from '../../types';
import { ConfigOptionRegistryEntry } from './shared';

export function buildCoreConfigOptionRegistry(
  defaultConfig: ResolvedConfig,
): ConfigOptionRegistryEntry[] {
  return [
    {
      path: 'logging.level',
      kind: 'enum',
      enumValues: ['debug', 'info', 'warn', 'error'],
      defaultValue: defaultConfig.logging.level,
      description: 'Minimum log level for runtime logging.',
    },
    {
      path: 'texthooker.launchAtStartup',
      kind: 'boolean',
      defaultValue: defaultConfig.texthooker.launchAtStartup,
      description: 'Launch texthooker server automatically when SubMiner starts.',
    },
    {
      path: 'websocket.enabled',
      kind: 'enum',
      enumValues: ['auto', 'true', 'false'],
      defaultValue: defaultConfig.websocket.enabled,
      description: 'Built-in subtitle websocket server mode.',
    },
    {
      path: 'websocket.port',
      kind: 'number',
      defaultValue: defaultConfig.websocket.port,
      description: 'Built-in subtitle websocket server port.',
    },
    {
      path: 'annotationWebsocket.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.annotationWebsocket.enabled,
      description: 'Annotated subtitle websocket server enabled state.',
    },
    {
      path: 'annotationWebsocket.port',
      kind: 'number',
      defaultValue: defaultConfig.annotationWebsocket.port,
      description: 'Annotated subtitle websocket server port.',
    },
    {
      path: 'subsync.defaultMode',
      kind: 'enum',
      enumValues: ['auto', 'manual'],
      defaultValue: defaultConfig.subsync.defaultMode,
      description: 'Subsync default mode.',
    },
    {
      path: 'subsync.replace',
      kind: 'boolean',
      defaultValue: defaultConfig.subsync.replace,
      description: 'Replace the active subtitle file when sync completes.',
    },
    {
      path: 'startupWarmups.lowPowerMode',
      kind: 'boolean',
      defaultValue: defaultConfig.startupWarmups.lowPowerMode,
      description: 'Defer startup warmups except Yomitan extension.',
    },
    {
      path: 'startupWarmups.mecab',
      kind: 'boolean',
      defaultValue: defaultConfig.startupWarmups.mecab,
      description: 'Warm up MeCab tokenizer at startup.',
    },
    {
      path: 'startupWarmups.yomitanExtension',
      kind: 'boolean',
      defaultValue: defaultConfig.startupWarmups.yomitanExtension,
      description: 'Warm up Yomitan extension at startup.',
    },
    {
      path: 'startupWarmups.subtitleDictionaries',
      kind: 'boolean',
      defaultValue: defaultConfig.startupWarmups.subtitleDictionaries,
      description: 'Warm up subtitle dictionaries at startup.',
    },
    {
      path: 'startupWarmups.jellyfinRemoteSession',
      kind: 'boolean',
      defaultValue: defaultConfig.startupWarmups.jellyfinRemoteSession,
      description: 'Warm up Jellyfin remote session at startup.',
    },
    {
      path: 'shortcuts.multiCopyTimeoutMs',
      kind: 'number',
      defaultValue: defaultConfig.shortcuts.multiCopyTimeoutMs,
      description: 'Timeout for multi-copy/mine modes.',
    },
  ];
}
