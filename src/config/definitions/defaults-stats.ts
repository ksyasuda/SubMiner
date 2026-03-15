import { ResolvedConfig } from '../../types.js';

export const STATS_DEFAULT_CONFIG: Pick<ResolvedConfig, 'stats'> = {
  stats: {
    toggleKey: 'Backquote',
    serverPort: 5175,
    autoStartServer: true,
    autoOpenBrowser: true,
  },
};
