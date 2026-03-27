import { ResolvedConfig } from '../../types/config.js';
import { ConfigOptionRegistryEntry } from './shared.js';

export function buildStatsConfigOptionRegistry(
  defaultConfig: ResolvedConfig,
): ConfigOptionRegistryEntry[] {
  return [
    {
      path: 'stats.toggleKey',
      kind: 'string',
      defaultValue: defaultConfig.stats.toggleKey,
      description: 'Key code to toggle the stats overlay.',
    },
    {
      path: 'stats.markWatchedKey',
      kind: 'string',
      defaultValue: defaultConfig.stats.markWatchedKey,
      description:
        'Key code to mark the current video as watched and advance to the next playlist entry.',
    },
    {
      path: 'stats.serverPort',
      kind: 'number',
      defaultValue: defaultConfig.stats.serverPort,
      description: 'Port for the stats HTTP server.',
    },
    {
      path: 'stats.autoStartServer',
      kind: 'boolean',
      defaultValue: defaultConfig.stats.autoStartServer,
      description: 'Automatically start the stats server on launch.',
    },
    {
      path: 'stats.autoOpenBrowser',
      kind: 'boolean',
      defaultValue: defaultConfig.stats.autoOpenBrowser,
      description: 'Automatically open the stats dashboard in a browser when the server starts.',
    },
  ];
}
