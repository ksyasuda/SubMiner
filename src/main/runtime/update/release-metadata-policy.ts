type AppUpdateMetadata = {
  available: boolean;
  version: string;
  canUpdate?: boolean;
};

export function shouldFetchReleaseMetadataForPlatform(
  platform: NodeJS.Platform,
  appUpdate: AppUpdateMetadata,
): boolean {
  if (platform !== 'darwin') {
    return true;
  }
  return appUpdate.canUpdate !== false;
}
