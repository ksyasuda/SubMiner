export async function runSupportAssetUpdatesForLauncherResult<
  TLauncherResult,
  TSupportResult extends { status: string; command?: string; message?: string },
>(options: {
  launcherResult: TLauncherResult;
  updateSupportAssets: () => Promise<TSupportResult[]>;
  logWarn: (message: string, details?: unknown) => void;
}): Promise<TLauncherResult> {
  try {
    const supportResults = await options.updateSupportAssets();
    for (const result of supportResults) {
      if (result.status === 'protected' && result.command) {
        options.logWarn(`Rofi theme update requires manual command: ${result.command}`);
      } else if (result.status === 'hash-mismatch' || result.status === 'missing-asset') {
        options.logWarn(`Rofi theme update skipped: ${result.message ?? result.status}`);
      }
    }
  } catch (error) {
    options.logWarn('Support asset update failed after launcher update', error);
  }
  return options.launcherResult;
}
