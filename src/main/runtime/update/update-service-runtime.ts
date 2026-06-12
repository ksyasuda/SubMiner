import { app, dialog } from 'electron';
import { execFile } from 'node:child_process';
import path from 'node:path';
import type { UpdateChannel, UpdatesConfig } from '../../../types/config';
import type { OverlayNotificationPayload } from '../../../types/notification';
import { createElectronAppUpdater, isNativeUpdaterSupported } from './app-updater';
import { createCurlFetch, createGlobalFetch } from './fetch-adapter';
import { createCurlHttpExecutor } from './curl-http-executor';
import { createFetchHttpExecutor } from './fetch-http-executor';
import {
  fetchLatestStableRelease,
  fetchReleaseAssetBuffer,
  fetchReleaseAssetText,
  findReleaseAsset,
  parseSha256Sums,
  type GitHubRelease,
} from './release-assets';
import { shouldFetchReleaseMetadataForPlatform } from './release-metadata-policy';
import { updateLauncherFromRelease } from './launcher-updater';
import { notifyUpdateAvailable } from './update-notifications';
import { createUpdateDialogPresenter } from './update-dialogs';
import { createFileUpdateStateStore, createUpdateService } from './update-service';
import { updateSupportAssetsFromRelease } from './support-assets';
import { runSupportAssetUpdatesForLauncherResult } from './update-support-assets-runtime';

const SUBMINER_BUNDLE_ID = 'com.sudacode.SubMiner';

export interface UpdateServiceRuntimeDeps {
  userDataPath: string;
  getUpdatesConfig: () => Required<UpdatesConfig>;
  logInfo: (message: string) => void;
  logWarn: (message: string, details?: unknown) => void;
  showOverlayNotification: (payload: OverlayNotificationPayload) => void;
  showDesktopNotification: (title: string, options: { body: string }) => void;
  showMpvOsd: (message: string) => void;
  withStatsWindowLayerSuspendedForNativeDialog: <T>(showDialog: () => Promise<T>) => Promise<T>;
}

export function createUpdateServiceRuntime(deps: UpdateServiceRuntimeDeps): {
  getUpdateService: () => ReturnType<typeof createUpdateService>;
} {
  const updateStateStore = createFileUpdateStateStore(
    path.join(deps.userDataPath, 'update-state.json'),
  );
  let updateService: ReturnType<typeof createUpdateService> | null = null;
  const globalFetchForUpdater = createGlobalFetch();
  const curlFetch = createCurlFetch();

  function createNativeUpdaterHttpExecutor() {
    if (process.platform === 'win32') {
      return createFetchHttpExecutor();
    }
    return createCurlHttpExecutor();
  }

  function getFetchForUpdater() {
    if (process.platform === 'win32') return globalFetchForUpdater;
    return curlFetch;
  }

  async function updateLauncherFromSelectedRelease(
    launcherPath?: string,
    channel: UpdateChannel = deps.getUpdatesConfig().channel,
    release: GitHubRelease | null = null,
  ) {
    const fetchForUpdater = getFetchForUpdater();
    if (!release) {
      return { status: 'missing-asset', message: `No ${channel} GitHub release found.` };
    }
    const sumsAsset = findReleaseAsset(release, 'SHA256SUMS.txt');
    if (!sumsAsset) {
      return { status: 'missing-asset', message: 'Release has no SHA256SUMS.txt asset.' };
    }
    const sums = parseSha256Sums(
      await fetchReleaseAssetText(fetchForUpdater, sumsAsset.browser_download_url),
    );
    const launcherResult = await updateLauncherFromRelease({
      release,
      sha256Sums: sums,
      launcherPath,
      downloadAsset: (url) => fetchReleaseAssetBuffer(fetchForUpdater, url),
    });
    return runSupportAssetUpdatesForLauncherResult({
      launcherResult,
      updateSupportAssets: () =>
        updateSupportAssetsFromRelease({
          release,
          sha256Sums: sums,
          downloadAsset: (url) => fetchReleaseAssetBuffer(fetchForUpdater, url),
        }),
      logWarn: (message, details) => deps.logWarn(message, details),
    });
  }

  function getUpdateService() {
    if (updateService) return updateService;
    const appUpdater = createElectronAppUpdater({
      currentVersion: app.getVersion(),
      isPackaged: app.isPackaged,
      log: (message) => deps.logInfo(message),
      getChannel: () => deps.getUpdatesConfig().channel,
      configureHttpExecutor: createNativeUpdaterHttpExecutor,
      disableDifferentialDownload: true,
      isNativeUpdaterSupported: () =>
        isNativeUpdaterSupported({
          platform: process.platform,
          isPackaged: app.isPackaged,
          execPath: process.execPath,
          env: process.env,
          log: (message) => deps.logWarn(message),
        }),
    });
    const updateDialogPresenter = createUpdateDialogPresenter({
      platform: process.platform,
      focusApp: async () => {
        if (process.platform !== 'darwin') {
          app.focus({ steal: true });
          return;
        }
        try {
          await app.dock?.show();
        } catch (error) {
          deps.logWarn('Failed to show macOS dock before update dialog', error);
        }
        // app.focus({ steal: true }) alone does not reliably activate the process
        // when SubMiner was reached via `subminer -u` (single-instance forwarding
        // from a CLI-spawned child). osascript's `activate` uses LaunchServices,
        // which is the only path that reliably brings the running app forward.
        await new Promise<void>((resolve) => {
          execFile(
            '/usr/bin/osascript',
            ['-e', `tell application id "${SUBMINER_BUNDLE_ID}" to activate`],
            { timeout: 2000 },
            (error) => {
              if (error) {
                deps.logWarn(
                  `Failed to activate SubMiner via osascript: ${error instanceof Error ? error.message : String(error)}`,
                );
              }
              resolve();
            },
          );
        });
        app.focus({ steal: true });
      },
      withStatsWindowLayerSuspended: (showDialog) =>
        deps.withStatsWindowLayerSuspendedForNativeDialog(showDialog),
      showMessageBox: (options) => dialog.showMessageBox(options),
    });
    updateService = createUpdateService({
      getConfig: () => deps.getUpdatesConfig(),
      getCurrentVersion: () => app.getVersion(),
      now: () => Date.now(),
      readState: () => updateStateStore.readState(),
      writeState: (state) => updateStateStore.writeState(state),
      checkAppUpdate: (channel) => appUpdater.checkForUpdates(channel),
      shouldFetchReleaseMetadata: ({ request, appUpdate }) =>
        shouldFetchReleaseMetadataForPlatform(process.platform, appUpdate, request),
      fetchLatestStableRelease: (channel) =>
        fetchLatestStableRelease({ fetch: getFetchForUpdater(), channel }),
      updateLauncher: (launcherPath, channel, release) =>
        updateLauncherFromSelectedRelease(launcherPath, channel, release),
      showNoUpdateDialog: (version) => updateDialogPresenter.showNoUpdateDialog(version),
      showUpdateAvailableDialog: (version) =>
        updateDialogPresenter.showUpdateAvailableDialog(version),
      showUpdateFailedDialog: (message) => updateDialogPresenter.showUpdateFailedDialog(message),
      showManualUpdateRequiredDialog: (version) =>
        updateDialogPresenter.showManualUpdateRequiredDialog(version),
      downloadAppUpdate: () => appUpdater.downloadUpdate(),
      showRestartDialog: () => updateDialogPresenter.showRestartDialog(),
      quitAndInstall: () => appUpdater.quitAndInstall(),
      notifyUpdateAvailable: (version) =>
        notifyUpdateAvailable(
          { notificationType: deps.getUpdatesConfig().notificationType, version },
          {
            showSystemNotification: (title, body) => deps.showDesktopNotification(title, { body }),
            showOverlayNotification: (payload) => deps.showOverlayNotification(payload),
            showOsdNotification: (message) => {
              deps.showMpvOsd(message);
            },
            log: (message) => deps.logWarn(message),
          },
        ),
      log: (message) => deps.logWarn(message),
    });
    return updateService;
  }

  return { getUpdateService };
}
