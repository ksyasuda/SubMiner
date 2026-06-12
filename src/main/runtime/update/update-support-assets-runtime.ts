export async function runSupportAssetUpdatesForLauncherResult<
  TLauncherResult,
  TSupportResult extends { status: string; command?: string; message?: string },
>(options: {
  launcherResult: TLauncherResult;
  assetDescription?: string;
  updateSupportAssets: () => Promise<TSupportResult[]>;
  logWarn: (message: string, details?: unknown) => void;
}): Promise<TLauncherResult> {
  const assetDescription = options.assetDescription ?? 'Support asset update';
  try {
    const supportResults = await options.updateSupportAssets();
    for (const result of supportResults) {
      if (result.status === 'protected' && result.command) {
        options.logWarn(`${assetDescription} requires manual command: ${result.command}`);
      } else if (result.status === 'hash-mismatch' || result.status === 'missing-asset') {
        options.logWarn(`${assetDescription} skipped: ${result.message ?? result.status}`);
      }
    }
  } catch (error) {
    options.logWarn('Support asset update failed after launcher update', error);
  }
  return options.launcherResult;
}
