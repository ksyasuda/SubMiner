import os from 'node:os';
import { spawn } from 'node:child_process';
import { app, dialog, shell } from 'electron';
import { printHelp } from './cli/help';
import {
  configureEarlyAppPaths,
  normalizeLaunchMpvExtraArgs,
  normalizeLaunchMpvTargets,
  normalizeStartupArgv,
  applyEarlyLinuxCommandLineSwitches,
  sanitizeStartupEnv,
  sanitizeBackgroundEnv,
  sanitizeHelpEnv,
  sanitizeLaunchMpvEnv,
  hasTransportedStartupArgs,
  shouldForwardStartupArgvViaAppControl,
  shouldDetachBackgroundLaunch,
  shouldHandleHelpOnlyAtEntry,
  shouldHandleLaunchMpvAtEntry,
  shouldHandleStatsDaemonCommandAtEntry,
} from './main-entry-runtime';
import { requestSingleInstanceLockEarly } from './main/early-single-instance';
import { readConfiguredWindowsMpvLaunch } from './main-entry-launch-config';
import { sendAppControlCommand } from './shared/app-control-client';
import {
  detectInstalledFirstRunPluginCandidates,
  detectInstalledMpvPlugin,
  removeLegacyMpvPluginCandidates,
  resolvePackagedRuntimePluginPath,
} from './main/runtime/first-run-setup-plugin';
import { createWindowsMpvLaunchDeps, launchWindowsMpv } from './main/runtime/windows-mpv-launch';
import { runStatsDaemonControlFromProcess } from './stats-daemon-entry';
import { createFatalErrorReporter, registerFatalErrorHandlers } from './main/fatal-error';

const DEFAULT_TEXTHOOKER_PORT = 5174;

function applySanitizedEnv(sanitizedEnv: NodeJS.ProcessEnv): void {
  if (sanitizedEnv.NODE_NO_WARNINGS) {
    process.env.NODE_NO_WARNINGS = sanitizedEnv.NODE_NO_WARNINGS;
  }

  if (sanitizedEnv.VK_INSTANCE_LAYERS) {
    process.env.VK_INSTANCE_LAYERS = sanitizedEnv.VK_INSTANCE_LAYERS;
  } else {
    delete process.env.VK_INSTANCE_LAYERS;
  }
}

function resolveBundledWindowsMpvPluginEntrypoint(): string | undefined {
  return (
    resolvePackagedRuntimePluginPath({
      dirname: __dirname,
      appPath: app.getAppPath(),
      resourcesPath: process.resourcesPath,
    }) ?? undefined
  );
}

function buildInstalledWindowsMpvPluginMessage(pathValue: string, version: string | null): string {
  return [
    'SubMiner detected an installed mpv plugin at:',
    pathValue,
    '',
    "This mpv session will use the installed plugin. Remove it to use SubMiner's bundled runtime plugin automatically.",
    `Detected plugin version: ${version ?? 'unknown or legacy'}`,
  ].join('\n');
}

async function promptForWindowsLegacyMpvPluginRemoval(
  mpvPath: string,
  detection: { path: string | null; version: string | null },
): Promise<'removed' | 'continue' | 'cancel'> {
  const response = await dialog.showMessageBox({
    type: 'warning',
    title: 'SubMiner mpv plugin detected',
    message: buildInstalledWindowsMpvPluginMessage(
      detection.path ?? 'unknown path',
      detection.version,
    ),
    detail:
      'Remove the legacy SubMiner mpv plugin files from mpv before launching this video? This moves the files to the OS trash. SubMiner-managed playback will then use the bundled runtime plugin.',
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

  const candidates = detectInstalledFirstRunPluginCandidates({
    platform: 'win32',
    homeDir: os.homedir(),
    appDataDir: app.getPath('appData'),
    mpvExecutablePath: mpvPath,
  });
  const result = await removeLegacyMpvPluginCandidates({
    candidates,
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

function createWindowsRuntimePluginPolicy() {
  return {
    detectInstalledMpvPlugin: (mpvPath: string) =>
      detectInstalledMpvPlugin({
        platform: 'win32',
        homeDir: os.homedir(),
        appDataDir: app.getPath('appData'),
        mpvExecutablePath: mpvPath,
      }),
    notifyInstalledPluginDetected: (detection: {
      installed: boolean;
      path: string | null;
      version: string | null;
    }) => {
      if (!detection.installed || !detection.path) return;
      dialog.showMessageBoxSync({
        type: 'warning',
        title: 'SubMiner mpv plugin detected',
        message: buildInstalledWindowsMpvPluginMessage(detection.path, detection.version),
      });
    },
    resolveInstalledPluginBeforeLaunch: (
      detection: { path: string | null; version: string | null },
      mpvPath: string,
    ) => promptForWindowsLegacyMpvPluginRemoval(mpvPath, detection),
  };
}

process.argv = normalizeStartupArgv(process.argv, process.env);
applyEarlyLinuxCommandLineSwitches(app.commandLine, process.argv);
applySanitizedEnv(sanitizeStartupEnv(process.env));
const userDataPath = configureEarlyAppPaths(app);
const reportFatalError = createFatalErrorReporter({
  showErrorBox: (title, details) => dialog.showErrorBox(title, details),
  consoleError: (message, error) => console.error(message, error),
});
registerFatalErrorHandlers({
  reportFatalError,
  exit: (code) => app.exit(code),
});

function startMainProcess(): void {
  const gotSingleInstanceLock = requestSingleInstanceLockEarly(app);
  if (!gotSingleInstanceLock) {
    app.exit(0);
    return;
  }
  try {
    require('./main.js');
  } catch (error) {
    reportFatalError(error, {
      title: 'SubMiner startup failed',
      context: 'SubMiner failed while loading the main process.',
    });
    app.exit(1);
  }
}

async function forwardStartupArgvViaAppControlIfAvailable(): Promise<boolean> {
  if (!shouldForwardStartupArgvViaAppControl(process.argv, process.env)) {
    return false;
  }

  const result = await sendAppControlCommand(process.argv, {
    configDir: userDataPath,
    timeoutMs: 500,
  });
  if (result.ok) {
    app.exit(0);
    return true;
  }
  if (!result.unavailable) {
    console.error(`SubMiner app-control handoff failed: ${result.error ?? 'unknown error'}`);
    app.exit(1);
    return true;
  }
  return false;
}

async function runEntryProcess(): Promise<void> {
  if (shouldHandleHelpOnlyAtEntry(process.argv, process.env)) {
    const sanitizedEnv = sanitizeHelpEnv(process.env);
    process.env.NODE_NO_WARNINGS = sanitizedEnv.NODE_NO_WARNINGS;
    if (!sanitizedEnv.VK_INSTANCE_LAYERS) {
      delete process.env.VK_INSTANCE_LAYERS;
    }
    printHelp(DEFAULT_TEXTHOOKER_PORT);
    process.exit(0);
    return;
  }

  if (shouldHandleLaunchMpvAtEntry(process.argv, process.env)) {
    const sanitizedEnv = sanitizeLaunchMpvEnv(process.env);
    applySanitizedEnv(sanitizedEnv);
    void app.whenReady().then(async () => {
      const configuredMpvLaunch = readConfiguredWindowsMpvLaunch(userDataPath);
      const result = await launchWindowsMpv(
        normalizeLaunchMpvTargets(process.argv),
        createWindowsMpvLaunchDeps({
          getEnv: (name) => process.env[name],
          showError: (title, content) => {
            dialog.showErrorBox(title, content);
          },
        }),
        normalizeLaunchMpvExtraArgs(process.argv),
        process.execPath,
        resolveBundledWindowsMpvPluginEntrypoint(),
        configuredMpvLaunch.executablePath,
        configuredMpvLaunch.launchMode,
        createWindowsRuntimePluginPolicy(),
        configuredMpvLaunch.pluginRuntimeConfig,
      );
      app.exit(result.ok ? 0 : 1);
    });
    return;
  }

  if (shouldHandleStatsDaemonCommandAtEntry(process.argv, process.env)) {
    void app.whenReady().then(async () => {
      const exitCode = await runStatsDaemonControlFromProcess(app.getPath('userData'));
      app.exit(exitCode);
    });
    return;
  }

  if (await forwardStartupArgvViaAppControlIfAvailable()) {
    return;
  }

  if (shouldDetachBackgroundLaunch(process.argv, process.env)) {
    const childArgs = hasTransportedStartupArgs(process.env) ? [] : process.argv.slice(1);
    const child = spawn(process.execPath, childArgs, {
      detached: true,
      stdio: 'ignore',
      env: sanitizeBackgroundEnv(process.env),
    });
    child.unref();
    process.exit(0);
    return;
  }

  startMainProcess();
}

void runEntryProcess().catch((error) => {
  console.error('SubMiner app-control handoff failed:', error);
  startMainProcess();
});
