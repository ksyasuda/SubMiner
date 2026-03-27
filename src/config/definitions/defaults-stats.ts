import { ResolvedConfig } from '../../types/config.js';

export const STATS_DEFAULT_CONFIG: Pick<ResolvedConfig, 'stats'> = {
  stats: {
    toggleKey: 'Backquote',
    markWatchedKey: 'KeyW',
    serverPort: 6969,
    autoStartServer: true,
    autoOpenBrowser: true,
  },
};
