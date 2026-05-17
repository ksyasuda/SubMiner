"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.resolveMacAppBundlePath = resolveMacAppBundlePath;
exports.isMacApplicationsFolderBundle = isMacApplicationsFolderBundle;
exports.isKnownLinuxPackageManagedAppImage = isKnownLinuxPackageManagedAppImage;
exports.isNativeUpdaterSupported = isNativeUpdaterSupported;
exports.configureAutoUpdater = configureAutoUpdater;
exports.createElectronAppUpdater = createElectronAppUpdater;
const node_fs_1 = require("node:fs");
const node_child_process_1 = require("node:child_process");
const node_os_1 = __importDefault(require("node:os"));
const node_path_1 = __importDefault(require("node:path"));
const node_util_1 = require("node:util");
const electron_updater_1 = require("electron-updater");
const release_assets_1 = require("./release-assets");
const updaterErrorListeners = new WeakMap();
const execFileAsync = (0, node_util_1.promisify)(node_child_process_1.execFile);
function resolveMacAppBundlePath(execPath) {
    const marker = '.app/Contents/MacOS/';
    const markerIndex = execPath.indexOf(marker);
    if (markerIndex < 0)
        return null;
    return execPath.slice(0, markerIndex + '.app'.length);
}
async function readMacCodeSignature(appBundlePath) {
    try {
        const result = await execFileAsync('/usr/bin/codesign', ['-dv', '--verbose=4', appBundlePath], {
            encoding: 'utf8',
        });
        return `${result.stdout ?? ''}\n${result.stderr ?? ''}`;
    }
    catch {
        return null;
    }
}
function realpathOrOriginal(filePath) {
    try {
        return (0, node_fs_1.realpathSync)(filePath);
    }
    catch {
        return filePath;
    }
}
function isSameOrInsideDirectory(parentPath, candidatePath) {
    const relative = node_path_1.default.relative(parentPath, candidatePath);
    return (relative === '' ||
        (relative.length > 0 && !relative.startsWith('..') && !node_path_1.default.isAbsolute(relative)));
}
function isMacApplicationsFolderBundle(appBundlePath, homeDir = node_os_1.default.homedir()) {
    const resolvedBundlePath = node_path_1.default.resolve(appBundlePath);
    return (isSameOrInsideDirectory('/Applications', resolvedBundlePath) ||
        isSameOrInsideDirectory(node_path_1.default.join(homeDir, 'Applications'), resolvedBundlePath));
}
function isKnownLinuxPackageManagedAppImage(appImagePath) {
    return realpathOrOriginal(appImagePath) === '/opt/SubMiner/SubMiner.AppImage';
}
async function isNativeUpdaterSupported(options) {
    if (!options.isPackaged) {
        options.log?.('Skipping native updater because this build is not packaged.');
        return false;
    }
    if (options.platform === 'linux') {
        options.log?.('Skipping native Linux updater because Linux tray checks use GitHub release assets.');
        return false;
    }
    if (options.platform !== 'darwin') {
        options.log?.('Skipping native updater because this platform uses GitHub metadata checks.');
        return false;
    }
    const appBundlePath = resolveMacAppBundlePath(options.execPath);
    if (!appBundlePath) {
        options.log?.('Skipping native macOS updater because the app bundle path could not be resolved.');
        return false;
    }
    if (!isMacApplicationsFolderBundle(appBundlePath, options.homeDir)) {
        options.log?.('Skipping native macOS updater because the app is not installed in an Applications folder.');
        return false;
    }
    const signature = await (options.readCodeSignature ?? readMacCodeSignature)(appBundlePath);
    if (!signature) {
        options.log?.('Skipping native macOS updater because the app code signature could not be read.');
        return false;
    }
    if (/Signature=adhoc\b/.test(signature) || /TeamIdentifier=not set\b/.test(signature)) {
        options.log?.('Skipping native macOS updater because this build is ad-hoc signed.');
        return false;
    }
    return true;
}
function configureAutoUpdater(updater, log = () => { }, channel = 'stable') {
    updater.autoDownload = false;
    // On macOS this avoids invoking Squirrel until the explicit restart/install step.
    updater.autoInstallOnAppQuit = false;
    updater.allowPrerelease = channel === 'prerelease';
    updater.allowDowngrade = false;
    updater.logger = {
        info: () => { },
        debug: () => { },
        warn: (message) => log(message),
        error: (message) => log(message),
    };
    const previousErrorListener = updaterErrorListeners.get(updater);
    if (previousErrorListener) {
        if (updater.off) {
            updater.off('error', previousErrorListener);
        }
        else {
            updater.removeListener?.('error', previousErrorListener);
        }
    }
    if (updater.on) {
        const errorListener = (error) => {
            const message = error instanceof Error ? error.message : String(error);
            log(`Updater error event: ${message}`);
        };
        updater.on('error', errorListener);
        updaterErrorListeners.set(updater, errorListener);
    }
    return updater;
}
function createElectronAppUpdater(options) {
    const getChannel = options.getChannel ?? (() => 'stable');
    const updater = configureAutoUpdater(options.updater ?? electron_updater_1.autoUpdater, options.log, getChannel());
    let nativeUpdaterSupported = null;
    async function getNativeUpdaterSupported() {
        if (!options.isNativeUpdaterSupported)
            return true;
        if (nativeUpdaterSupported === null) {
            nativeUpdaterSupported = Promise.resolve(options.isNativeUpdaterSupported());
        }
        return nativeUpdaterSupported;
    }
    return {
        async checkForUpdates(channel) {
            if (!options.isPackaged) {
                return {
                    available: false,
                    version: options.currentVersion,
                    canUpdate: false,
                };
            }
            if (!(await getNativeUpdaterSupported())) {
                options.log('Skipping native app update check because native updater is unsupported.');
                return {
                    available: false,
                    version: options.currentVersion,
                    canUpdate: false,
                };
            }
            configureAutoUpdater(updater, options.log, channel ?? getChannel());
            const result = await updater.checkForUpdates();
            const version = result?.updateInfo?.version ?? options.currentVersion;
            return {
                available: (0, release_assets_1.compareSemverLike)(version, options.currentVersion) > 0,
                version,
                canUpdate: true,
            };
        },
        async downloadUpdate() {
            if (!options.isPackaged) {
                options.log('Skipping app update download because this build is not packaged.');
                return;
            }
            if (!(await getNativeUpdaterSupported())) {
                options.log('Skipping app update download because native updater is unsupported.');
                return;
            }
            await updater.downloadUpdate();
        },
        async quitAndInstall() {
            if (!options.isPackaged) {
                options.log('Skipping app update install because this build is not packaged.');
                return;
            }
            if (!(await getNativeUpdaterSupported())) {
                options.log('Skipping app update install because native updater is unsupported.');
                return;
            }
            updater.quitAndInstall(false, true);
        },
    };
}
//# sourceMappingURL=app-updater.js.map