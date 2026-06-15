"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const node_os_1 = __importDefault(require("node:os"));
const node_child_process_1 = require("node:child_process");
const electron_1 = require("electron");
const help_1 = require("./cli/help");
const main_entry_runtime_1 = require("./main-entry-runtime");
const early_single_instance_1 = require("./main/early-single-instance");
const main_entry_launch_config_1 = require("./main-entry-launch-config");
const app_control_client_1 = require("./shared/app-control-client");
const first_run_setup_plugin_1 = require("./main/runtime/first-run-setup-plugin");
const windows_mpv_launch_1 = require("./main/runtime/windows-mpv-launch");
const stats_daemon_entry_1 = require("./stats-daemon-entry");
const fatal_error_1 = require("./main/fatal-error");
const mpv_logging_args_1 = require("./shared/mpv-logging-args");
const log_files_1 = require("./shared/log-files");
const DEFAULT_TEXTHOOKER_PORT = 5174;
function appendWindowsMpvLaunchLog(message, logRotation) {
    if (!(0, log_files_1.isLogFileEnabled)('app')) {
        return;
    }
    const timestamp = new Date().toISOString().replace('T', ' ').slice(0, 19);
    (0, log_files_1.appendLogLine)(process.env.SUBMINER_APP_LOG?.trim() || (0, log_files_1.resolveDefaultLogFilePath)('app'), `[subminer] - ${timestamp} - INFO - [main:windows-mpv-launch] ${message}`, { rotation: logRotation });
}
function applySanitizedEnv(sanitizedEnv) {
    if (sanitizedEnv.NODE_NO_WARNINGS) {
        process.env.NODE_NO_WARNINGS = sanitizedEnv.NODE_NO_WARNINGS;
    }
    if (sanitizedEnv.VK_INSTANCE_LAYERS) {
        process.env.VK_INSTANCE_LAYERS = sanitizedEnv.VK_INSTANCE_LAYERS;
    }
    else {
        delete process.env.VK_INSTANCE_LAYERS;
    }
}
function resolveBundledWindowsMpvPluginEntrypoint() {
    return ((0, first_run_setup_plugin_1.resolvePackagedRuntimePluginPath)({
        dirname: __dirname,
        appPath: electron_1.app.getAppPath(),
        resourcesPath: process.resourcesPath,
    }) ?? undefined);
}
function buildInstalledWindowsMpvPluginMessage(pathValue, version) {
    return [
        'SubMiner detected an installed mpv plugin at:',
        pathValue,
        '',
        "This mpv session will use the installed plugin. Remove it to use SubMiner's bundled runtime plugin automatically.",
        `Detected plugin version: ${version ?? 'unknown or legacy'}`,
    ].join('\n');
}
async function promptForWindowsLegacyMpvPluginRemoval(mpvPath, detection) {
    const response = await electron_1.dialog.showMessageBox({
        type: 'warning',
        title: 'SubMiner mpv plugin detected',
        message: buildInstalledWindowsMpvPluginMessage(detection.path ?? 'unknown path', detection.version),
        detail: 'Remove the legacy SubMiner mpv plugin files from mpv before launching this video? This moves the files to the OS trash. SubMiner-managed playback will then use the bundled runtime plugin.',
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
    const candidates = (0, first_run_setup_plugin_1.detectInstalledFirstRunPluginCandidates)({
        platform: 'win32',
        homeDir: node_os_1.default.homedir(),
        appDataDir: electron_1.app.getPath('appData'),
        mpvExecutablePath: mpvPath,
    });
    const result = await (0, first_run_setup_plugin_1.removeLegacyMpvPluginCandidates)({
        candidates,
        trashItem: (candidatePath) => electron_1.shell.trashItem(candidatePath),
    });
    if (result.ok) {
        await electron_1.dialog.showMessageBox({
            type: 'info',
            title: 'Legacy mpv plugin removed',
            message: 'Legacy mpv plugin removed. SubMiner-managed playback will use the bundled runtime plugin.',
        });
        return 'removed';
    }
    await electron_1.dialog.showMessageBox({
        type: 'error',
        title: 'Could not remove legacy mpv plugin',
        message: 'Some legacy SubMiner mpv plugin files could not be moved to the trash.',
        detail: result.failedPaths.map((failure) => `${failure.path}: ${failure.message}`).join('\n'),
    });
    return 'cancel';
}
function createWindowsRuntimePluginPolicy() {
    return {
        detectInstalledMpvPlugin: (mpvPath) => (0, first_run_setup_plugin_1.detectInstalledMpvPlugin)({
            platform: 'win32',
            homeDir: node_os_1.default.homedir(),
            appDataDir: electron_1.app.getPath('appData'),
            mpvExecutablePath: mpvPath,
        }),
        notifyInstalledPluginDetected: (detection) => {
            if (!detection.installed || !detection.path)
                return;
            electron_1.dialog.showMessageBoxSync({
                type: 'warning',
                title: 'SubMiner mpv plugin detected',
                message: buildInstalledWindowsMpvPluginMessage(detection.path, detection.version),
            });
        },
        resolveInstalledPluginBeforeLaunch: (detection, mpvPath) => promptForWindowsLegacyMpvPluginRemoval(mpvPath, detection),
    };
}
process.argv = (0, main_entry_runtime_1.normalizeStartupArgv)(process.argv, process.env);
(0, main_entry_runtime_1.applyEarlyLinuxCommandLineSwitches)(electron_1.app.commandLine, process.argv);
applySanitizedEnv((0, main_entry_runtime_1.sanitizeStartupEnv)(process.env));
const userDataPath = (0, main_entry_runtime_1.configureEarlyAppPaths)(electron_1.app);
const reportFatalError = (0, fatal_error_1.createFatalErrorReporter)({
    showErrorBox: (title, details) => electron_1.dialog.showErrorBox(title, details),
    consoleError: (message, error) => console.error(message, error),
});
(0, fatal_error_1.registerFatalErrorHandlers)({
    reportFatalError,
    exit: (code) => electron_1.app.exit(code),
});
function startMainProcess() {
    const gotSingleInstanceLock = (0, early_single_instance_1.requestSingleInstanceLockEarly)(electron_1.app);
    if (!gotSingleInstanceLock) {
        electron_1.app.exit(0);
        return;
    }
    try {
        require('./main.js');
    }
    catch (error) {
        reportFatalError(error, {
            title: 'SubMiner startup failed',
            context: 'SubMiner failed while loading the main process.',
        });
        electron_1.app.exit(1);
    }
}
async function forwardStartupArgvViaAppControlIfAvailable() {
    if (!(0, main_entry_runtime_1.shouldForwardStartupArgvViaAppControl)(process.argv, process.env)) {
        return false;
    }
    const result = await (0, app_control_client_1.sendAppControlCommand)(process.argv, {
        configDir: userDataPath,
        timeoutMs: 500,
    });
    if (result.ok) {
        electron_1.app.exit(0);
        return true;
    }
    if (!result.unavailable) {
        console.error(`SubMiner app-control handoff failed: ${result.error ?? 'unknown error'}`);
        electron_1.app.exit(1);
        return true;
    }
    return false;
}
async function runEntryProcess() {
    if ((0, main_entry_runtime_1.shouldHandleHelpOnlyAtEntry)(process.argv, process.env)) {
        const sanitizedEnv = (0, main_entry_runtime_1.sanitizeHelpEnv)(process.env);
        process.env.NODE_NO_WARNINGS = sanitizedEnv.NODE_NO_WARNINGS;
        if (!sanitizedEnv.VK_INSTANCE_LAYERS) {
            delete process.env.VK_INSTANCE_LAYERS;
        }
        (0, help_1.printHelp)(DEFAULT_TEXTHOOKER_PORT);
        process.exit(0);
        return;
    }
    if ((0, main_entry_runtime_1.shouldHandleLaunchMpvAtEntry)(process.argv, process.env)) {
        const sanitizedEnv = (0, main_entry_runtime_1.sanitizeLaunchMpvEnv)(process.env);
        applySanitizedEnv(sanitizedEnv);
        await electron_1.app.whenReady();
        const configuredMpvLaunch = (0, main_entry_launch_config_1.readConfiguredWindowsMpvLaunch)(userDataPath);
        const extraArgs = (0, main_entry_runtime_1.normalizeLaunchMpvExtraArgs)(process.argv);
        (0, log_files_1.applyLogFileTogglesToEnv)(configuredMpvLaunch.logFiles);
        const mpvLogPath = (0, log_files_1.isLogFileEnabled)('mpv')
            ? process.env.SUBMINER_MPV_LOG?.trim() || (0, log_files_1.resolveDefaultLogFilePath)('mpv')
            : '';
        if (mpvLogPath) {
            (0, log_files_1.pruneLogDirectoryForPath)(mpvLogPath, configuredMpvLaunch.logRotation);
        }
        const result = await (0, windows_mpv_launch_1.launchWindowsMpv)((0, main_entry_runtime_1.normalizeLaunchMpvTargets)(process.argv), (0, windows_mpv_launch_1.createWindowsMpvLaunchDeps)({
            getEnv: (name) => process.env[name],
            isAppControlServerAvailable: () => (0, app_control_client_1.isAppControlServerAvailable)({
                configDir: userDataPath,
                timeoutMs: 350,
            }),
            sendAppControlCommand: (argv) => (0, app_control_client_1.sendAppControlCommand)(argv, {
                configDir: userDataPath,
                timeoutMs: 1000,
            }),
            showError: (title, content) => {
                electron_1.dialog.showErrorBox(title, content);
            },
            logInfo: (message) => appendWindowsMpvLaunchLog(message, configuredMpvLaunch.logRotation),
        }), [...extraArgs, ...(0, mpv_logging_args_1.buildMpvLoggingArgs)(configuredMpvLaunch.logLevel, mpvLogPath, extraArgs)], process.execPath, resolveBundledWindowsMpvPluginEntrypoint(), configuredMpvLaunch.executablePath, configuredMpvLaunch.launchMode, createWindowsRuntimePluginPolicy(), configuredMpvLaunch.pluginRuntimeConfig);
        electron_1.app.exit(result.ok ? 0 : 1);
        return;
    }
    if ((0, main_entry_runtime_1.shouldHandleStatsDaemonCommandAtEntry)(process.argv, process.env)) {
        await electron_1.app.whenReady();
        const exitCode = await (0, stats_daemon_entry_1.runStatsDaemonControlFromProcess)(electron_1.app.getPath('userData'));
        electron_1.app.exit(exitCode);
        return;
    }
    if (await forwardStartupArgvViaAppControlIfAvailable()) {
        return;
    }
    if ((0, main_entry_runtime_1.shouldDetachBackgroundLaunch)(process.argv, process.env)) {
        const childArgs = (0, main_entry_runtime_1.hasTransportedStartupArgs)(process.env) ? [] : process.argv.slice(1);
        const child = (0, node_child_process_1.spawn)(process.execPath, childArgs, {
            detached: true,
            stdio: 'ignore',
            env: (0, main_entry_runtime_1.sanitizeBackgroundEnv)(process.env),
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
//# sourceMappingURL=main-entry.js.map