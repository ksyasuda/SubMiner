export interface MpvVolumePropertySource {
  requestProperty?: (name: string) => Promise<unknown>;
}

export async function resolveMpvVolumeScale(
  mpvClient: MpvVolumePropertySource,
  enabled: boolean,
): Promise<number | undefined> {
  if (!enabled) {
    return undefined;
  }

  if (!mpvClient.requestProperty) {
    return 1;
  }

  try {
    const volume = await mpvClient.requestProperty('volume');
    if (typeof volume !== 'number' || !Number.isFinite(volume) || volume < 0) {
      return 1;
    }
    return (volume / 100) ** 3;
  } catch {
    return 1;
  }
}
