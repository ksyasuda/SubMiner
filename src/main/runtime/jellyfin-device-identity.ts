export const DEFAULT_JELLYFIN_CLIENT_VERSION = '0.1.0';
export const DEFAULT_JELLYFIN_CLIENT_NAME = 'SubMiner';

export function normalizeJellyfinHostName(value: string): string {
  return value.trim();
}

export function createHostDerivedJellyfinDeviceId(hostName: string): string {
  return normalizeJellyfinHostName(hostName) || 'device';
}

export function resolveJellyfinDeviceId(params: { hostName: string }): string {
  return createHostDerivedJellyfinDeviceId(params.hostName);
}

export function resolveJellyfinRemoteDeviceName(params: { hostName: string }): string {
  return normalizeJellyfinHostName(params.hostName) || 'device';
}
