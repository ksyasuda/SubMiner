import { ResolveContext } from './context';
import { asBoolean, asNumber, asString, isObject } from './shared';

export function applyStatsConfig(context: ResolveContext): void {
  const { src, resolved, warn } = context;

  if (!isObject(src.stats)) return;

  const toggleKey = asString(src.stats.toggleKey);
  if (toggleKey !== undefined) {
    resolved.stats.toggleKey = toggleKey;
  } else if (src.stats.toggleKey !== undefined) {
    warn('stats.toggleKey', src.stats.toggleKey, resolved.stats.toggleKey, 'Expected string.');
  }

  const serverPort = asNumber(src.stats.serverPort);
  if (serverPort !== undefined) {
    resolved.stats.serverPort = serverPort;
  } else if (src.stats.serverPort !== undefined) {
    warn('stats.serverPort', src.stats.serverPort, resolved.stats.serverPort, 'Expected number.');
  }

  const autoStartServer = asBoolean(src.stats.autoStartServer);
  if (autoStartServer !== undefined) {
    resolved.stats.autoStartServer = autoStartServer;
  } else if (src.stats.autoStartServer !== undefined) {
    warn(
      'stats.autoStartServer',
      src.stats.autoStartServer,
      resolved.stats.autoStartServer,
      'Expected boolean.',
    );
  }

  const autoOpenBrowser = asBoolean(src.stats.autoOpenBrowser);
  if (autoOpenBrowser !== undefined) {
    resolved.stats.autoOpenBrowser = autoOpenBrowser;
  } else if (src.stats.autoOpenBrowser !== undefined) {
    warn(
      'stats.autoOpenBrowser',
      src.stats.autoOpenBrowser,
      resolved.stats.autoOpenBrowser,
      'Expected boolean.',
    );
  }
}
