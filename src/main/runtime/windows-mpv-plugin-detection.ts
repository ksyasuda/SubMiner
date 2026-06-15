import { app, dialog, shell } from 'electron';
import * as os from 'os';
import { i18n } from '../../i18n/index.js';
import {
  detectInstalledMpvPlugin,
  detectWindowsMpvPluginRemovalCandidates,
  removeLegacyMpvPluginCandidates,
  resolvePackagedRuntimePluginPath,
} from './first-run-setup-plugin';

export interface WindowsMpvPluginDetectionRuntimeDeps {
  mainDirname: string;
  logWarn: (message: string) => void;
}

export function createWindowsMpvPluginDetectionRuntime(
  deps: WindowsMpvPluginDetectionRuntimeDeps,
): {
  resolveBundledMpvRuntimePluginEntrypoint: () => string | undefined;
  detectWindowsInstalledMpvPlugin: (
    mpvExecutablePath: string,
  ) => ReturnType<typeof detectInstalledMpvPlugin>;
  logInstalledMpvPluginDetected: (detection: {
    path: string | null;
    version: string | null;
  }) => void;
  promptForLegacyMpvPluginRemovalBeforeWindowsLaunch: (
    mpvPath: string,
    detection: { path: string | null; version: string | null },
  ) => Promise<'removed' | 'continue' | 'cancel'>;
} {
  function resolveBundledMpvRuntimePluginEntrypoint(): string | undefined {
    return (
      resolvePackagedRuntimePluginPath({
        dirname: deps.mainDirname,
        appPath: app.getAppPath(),
        resourcesPath: process.resourcesPath,
      }) ?? undefined
    );
  }

  function detectWindowsInstalledMpvPlugin(mpvExecutablePath: string) {
    return detectInstalledMpvPlugin({
      platform: 'win32',
      homeDir: os.homedir(),
      appDataDir: app.getPath('appData'),
      mpvExecutablePath,
    });
  }

  function logInstalledMpvPluginDetected(detection: {
    path: string | null;
    version: string | null;
  }) {
    if (!detection.path) return;
    deps.logWarn(
      i18n.t('legacyPlugin.log.detected', {
        path: detection.path,
        version: detection.version ?? i18n.t('legacyPlugin.log.unknownOrLegacy'),
      }),
    );
  }

  async function promptForLegacyMpvPluginRemovalBeforeWindowsLaunch(
    mpvPath: string,
    detection: { path: string | null; version: string | null },
  ): Promise<'removed' | 'continue' | 'cancel'> {
    const response = await dialog.showMessageBox({
      type: 'warning',
      title: i18n.t('legacyPlugin.dialog.detected.title'),
      message: [
        i18n.t('legacyPlugin.dialog.detected.line1'),
        detection.path ?? i18n.t('legacyPlugin.dialog.detected.unknownPath'),
        '',
        i18n.t('legacyPlugin.dialog.detected.line2'),
        i18n.t('legacyPlugin.dialog.detected.versionLabel', {
          version: detection.version ?? i18n.t('legacyPlugin.log.unknownOrLegacy'),
        }),
      ].join('\n'),
      detail: i18n.t('legacyPlugin.dialog.detected.detail'),
      buttons: [
        i18n.t('legacyPlugin.dialog.detected.removeBtn'),
        i18n.t('legacyPlugin.dialog.detected.continueBtn'),
        i18n.t('legacyPlugin.dialog.detected.cancelBtn'),
      ],
      defaultId: 0,
      cancelId: 2,
    });

    if (response.response === 2) {
      return 'cancel';
    }
    if (response.response === 1) {
      return 'continue';
    }

    const result = await removeLegacyMpvPluginCandidates({
      candidates: detectWindowsMpvPluginRemovalCandidates({
        homeDir: os.homedir(),
        appDataDir: app.getPath('appData'),
        mpvExecutablePath: mpvPath,
      }),
      trashItem: (candidatePath) => shell.trashItem(candidatePath),
    });
    if (result.ok) {
      await dialog.showMessageBox({
        type: 'info',
        title: i18n.t('legacyPlugin.dialog.removed.title'),
        message: i18n.t('legacyPlugin.dialog.removed.message'),
      });
      return 'removed';
    }

    await dialog.showMessageBox({
      type: 'error',
      title: i18n.t('legacyPlugin.dialog.failed.title'),
      message: i18n.t('legacyPlugin.dialog.failed.message'),
      detail: result.failedPaths.map((failure) => `${failure.path}: ${failure.message}`).join('\n'),
    });
    return 'cancel';
  }

  return {
    resolveBundledMpvRuntimePluginEntrypoint,
    detectWindowsInstalledMpvPlugin,
    logInstalledMpvPluginDetected,
    promptForLegacyMpvPluginRemovalBeforeWindowsLaunch,
  };
}
