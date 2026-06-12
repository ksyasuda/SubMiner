import { app, dialog, shell } from 'electron';
import * as os from 'os';
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
      `SubMiner detected an installed mpv plugin at ${detection.path}. This mpv session will use the installed plugin. Remove it to use the bundled runtime plugin automatically. Detected plugin version: ${detection.version ?? 'unknown or legacy'}.`,
    );
  }

  async function promptForLegacyMpvPluginRemovalBeforeWindowsLaunch(
    mpvPath: string,
    detection: { path: string | null; version: string | null },
  ): Promise<'removed' | 'continue' | 'cancel'> {
    const response = await dialog.showMessageBox({
      type: 'warning',
      title: 'SubMiner mpv plugin detected',
      message: [
        'SubMiner detected an installed mpv plugin at:',
        detection.path ?? 'unknown path',
        '',
        "This mpv session will use the installed plugin unless it is removed. Remove it now to use SubMiner's bundled runtime plugin automatically.",
        `Detected plugin version: ${detection.version ?? 'unknown or legacy'}`,
      ].join('\n'),
      detail:
        'Remove the legacy SubMiner mpv plugin files from mpv before launching this video? This moves the files to the OS trash.',
      buttons: ['Remove legacy plugin', 'Continue with installed plugin', 'Cancel'],
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
        title: 'Legacy mpv plugin removed',
        message:
          'Legacy mpv plugin removed. SubMiner-managed playback will use the bundled runtime plugin.',
      });
      return 'removed';
    }

    await dialog.showMessageBox({
      type: 'error',
      title: 'Could not remove legacy mpv plugin',
      message: 'Some legacy SubMiner mpv plugin files could not be moved to the trash.',
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
