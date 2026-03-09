import type { AnkiConnectConfig } from '../../types';

export function getPreferredYomitanAnkiServerUrl(config: AnkiConnectConfig): string {
  if (config.enabled === true && config.proxy?.enabled === true) {
    const host = config.proxy.host || '127.0.0.1';
    const port = config.proxy.port || 8766;
    return `http://${host}:${port}`;
  }

  return config.url || 'http://127.0.0.1:8765';
}

export function shouldForceOverrideYomitanAnkiServer(config: AnkiConnectConfig): boolean {
  return config.enabled === true && config.proxy?.enabled === true;
}
