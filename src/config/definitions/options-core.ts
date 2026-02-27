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
      path: 'subsync.defaultMode',
      kind: 'enum',
      enumValues: ['auto', 'manual'],
      defaultValue: defaultConfig.subsync.defaultMode,
      description: 'Subsync default mode.',
    },
    {
      path: 'shortcuts.multiCopyTimeoutMs',
      kind: 'number',
      defaultValue: defaultConfig.shortcuts.multiCopyTimeoutMs,
      description: 'Timeout for multi-copy/mine modes.',
    },
    {
      path: 'bind_visible_overlay_to_mpv_sub_visibility',
      kind: 'boolean',
      defaultValue: defaultConfig.bind_visible_overlay_to_mpv_sub_visibility,
      description:
        'Link visible overlay toggles to MPV primary subtitle visibility.',
    },
  ];
}
