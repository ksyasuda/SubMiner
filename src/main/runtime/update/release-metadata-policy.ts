type AppUpdateMetadata = {
  available: boolean;
  version: string;
  canUpdate?: boolean;
};

type UpdateMetadataRequest = {
  source?: 'manual' | 'automatic' | 'launcher';
};

export function shouldFetchReleaseMetadataForPlatform(
  platform: NodeJS.Platform,
  appUpdate: AppUpdateMetadata,
  request: UpdateMetadataRequest = {},
): boolean {
  if (platform !== 'darwin') {
    return true;
  }
  if (appUpdate.canUpdate !== false) {
    return true;
  }
  return request.source === 'manual' || request.source === 'launcher';
}
