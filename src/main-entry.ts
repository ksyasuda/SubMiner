import os from 'node:os';
import { app, dialog, shell } from 'electron';
import { printHelp } from './cli/help';
import {
  applyBackgroundBootstrapCommandLineSwitches,
  configureEarlyAppPaths,
  exitBackgroundBootstrap,
  normalizeLaunchMpvExtraArgs,
  normalizeLaunchMpvTargets,
  normalizeStartupArgv,
  applyEarlyLinuxCommandLineSwitches,
  resolveAppControlHandoffTimeoutMs,
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
  spawnDetachedApp,
} from './main-entry-runtime';
import {
  requestSingleInstanceLockEarly,
  shouldBypassSingleInstanceLockForArgv,
} from './main/early-single-instance';
import { readConfiguredWindowsMpvLaunch } from './main-entry-launch-config';
import { isAppControlServerAvailable, sendAppControlCommand } from './shared/app-control-client';
import {
  detectInstalledFirstRunPluginCandidates,
  detectInstalledMpvPlugin,
  removeLegacyMpvPluginCandidates,
  resolvePackagedRuntimePluginPath,
} from './main/runtime/first-run-setup-plugin';
import { createWindowsMpvLaunchDeps, launchWindowsMpv } from './main/runtime/windows-mpv-launch';
import { runStatsDaemonControlFromProcess } from './stats-daemon-entry';
import { handleSyncCliAtEntry } from './main/sync-cli';
import { createFatalErrorReporter, registerFatalErrorHandlers } from './main/fatal-error';
import { enforceElectronRuntimeGuard } from './main/electron-runtime-guard';
import { buildMpvLoggingArgs } from './shared/mpv-logging-args';
import {
  applyLogFileTogglesToEnv,
  isLogFileEnabled,
  appendLogLine,
  pruneLogDirectoryForPath,
  resolveDefaultLogFilePath,
  type LogRotation,
} from './shared/log-files';
import {
  resolveX11ElectronRelaunchArgs,
  X11_ELECTRON_BOOTSTRAP_ENV,
} from './core/utils/electron-backend';

const DEFAULT_TEXTHOOKER_PORT = 5174;

function appendWindowsMpvLaunchLog(message: string, logRotation?: LogRotation): void {
  if (!isLogFileEnabled('app')) {
    return;
  }
  const timestamp = new Date().toISOString().replace('T', ' ').slice(0, 19);
  appendLogLine(
    process.env.SUBMINER_APP_LOG?.trim() || resolveDefaultLogFilePath('app'),
    `[subminer] - ${timestamp} - INFO - [main:windows-mpv-launch] ${message}`,
    { rotation: logRotation },
  );
}

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
applyBackgroundBootstrapCommandLineSwitches(app.commandLine, process.argv, process.env);
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
  // Normal launches serialize the runtime guard with the profile-scoped lock. Stats daemon
  // commands keep their existing lock bypass when Electron runs in Node mode.
  const gotSingleInstanceLock =
    shouldBypassSingleInstanceLockForArgv(process.argv) || requestSingleInstanceLockEarly(app);
  if (!gotSingleInstanceLock) {
    app.exit(0);
    return;
  }

  const runtimeGuard = enforceElectronRuntimeGuard({
    electronVersion: process.versions.electron ?? '',
    userDataPath,
  });
  if (!runtimeGuard.ok) {
    console.error(runtimeGuard.details);
    dialog.showErrorBox(runtimeGuard.title, runtimeGuard.details);
    app.exit(1);
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
    timeoutMs: resolveAppControlHandoffTimeoutMs(),
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
  // Headless sync CLI: must run first (its own --help/--version handling) and
  // exit before app.whenReady() so it works over SSH with no display server.
  if (await handleSyncCliAtEntry(process.argv, process.env, app.getVersion())) return;

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
    await app.whenReady();
    const configuredMpvLaunch = readConfiguredWindowsMpvLaunch(userDataPath);
    const extraArgs = normalizeLaunchMpvExtraArgs(process.argv);
    applyLogFileTogglesToEnv(configuredMpvLaunch.logFiles);
    const mpvLogPath = isLogFileEnabled('mpv')
      ? process.env.SUBMINER_MPV_LOG?.trim() || resolveDefaultLogFilePath('mpv')
      : '';
    if (mpvLogPath) {
      pruneLogDirectoryForPath(mpvLogPath, configuredMpvLaunch.logRotation);
    }
    const result = await launchWindowsMpv(
      normalizeLaunchMpvTargets(process.argv),
      createWindowsMpvLaunchDeps({
        getEnv: (name) => process.env[name],
        isAppControlServerAvailable: () =>
          isAppControlServerAvailable({
            configDir: userDataPath,
            timeoutMs: 350,
          }),
        sendAppControlCommand: (argv) =>
          sendAppControlCommand(argv, {
            configDir: userDataPath,
            timeoutMs: 1000,
          }),
        showError: (title, content) => {
          dialog.showErrorBox(title, content);
        },
        logInfo: (message) => appendWindowsMpvLaunchLog(message, configuredMpvLaunch.logRotation),
      }),
      [...extraArgs, ...buildMpvLoggingArgs(configuredMpvLaunch.logLevel, mpvLogPath, extraArgs)],
      process.execPath,
      resolveBundledWindowsMpvPluginEntrypoint(),
      configuredMpvLaunch.executablePath,
      configuredMpvLaunch.launchMode,
      createWindowsRuntimePluginPolicy(),
      configuredMpvLaunch.pluginRuntimeConfig,
    );
    app.exit(result.ok ? 0 : 1);
    return;
  }

  if (shouldHandleStatsDaemonCommandAtEntry(process.argv, process.env)) {
    await app.whenReady();
    const exitCode = await runStatsDaemonControlFromProcess(app.getPath('userData'));
    app.exit(exitCode);
    return;
  }

  if (await forwardStartupArgvViaAppControlIfAvailable()) {
    return;
  }

  const childArgs = hasTransportedStartupArgs(process.env) ? [] : process.argv.slice(1);
  const x11ChildArgs = resolveX11ElectronRelaunchArgs(childArgs, process.env);

  if (shouldDetachBackgroundLaunch(process.argv, process.env)) {
    const childEnv = sanitizeBackgroundEnv(process.env);
    if (x11ChildArgs) childEnv[X11_ELECTRON_BOOTSTRAP_ENV] = '1';
    spawnDetachedApp(x11ChildArgs ?? childArgs, childEnv);
    // Let Electron stop bootstrap Chromium children before its AppImage mount is released.
    exitBackgroundBootstrap(app);
    return;
  }

  if (x11ChildArgs) {
    const childEnv = sanitizeStartupEnv(process.env);
    childEnv[X11_ELECTRON_BOOTSTRAP_ENV] = '1';
    spawnDetachedApp(x11ChildArgs, childEnv);
    exitBackgroundBootstrap(app);
    return;
  }

  startMainProcess();
}

void runEntryProcess().catch((error) => {
  console.error('SubMiner app-control handoff failed:', error);
  startMainProcess();
});
