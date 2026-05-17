'use strict';
var __importDefault =
  (this && this.__importDefault) ||
  function (mod) {
    return mod && mod.__esModule ? mod : { default: mod };
  };
Object.defineProperty(exports, '__esModule', { value: true });
exports.createUpdateService = createUpdateService;
exports.createFileUpdateStateStore = createFileUpdateStateStore;
const node_fs_1 = __importDefault(require('node:fs'));
const node_path_1 = __importDefault(require('node:path'));
const release_assets_1 = require('./release-assets');
function getBestLatestVersion(currentVersion, appUpdate, release) {
  const releaseVersion = (0, release_assets_1.parseReleaseVersion)(release);
  const candidates = [appUpdate.version, releaseVersion].filter(
    (value) => typeof value === 'string' && value.length > 0,
  );
  const latest = candidates.reduce(
    (best, candidate) =>
      (0, release_assets_1.compareSemverLike)(candidate, best) > 0 ? candidate : best,
    currentVersion,
  );
  return {
    available:
      appUpdate.available || (0, release_assets_1.compareSemverLike)(latest, currentVersion) > 0,
    version: latest,
  };
}
function shouldSkipAutomaticCheck(config, state, now) {
  if (!config.enabled) return true;
  if (!state.lastAutomaticCheckAt) return false;
  const intervalMs = Math.max(1, config.checkIntervalHours) * 60 * 60 * 1000;
  return now - state.lastAutomaticCheckAt < intervalMs;
}
function summarizeError(error) {
  const raw = error instanceof Error ? error.message : String(error);
  const firstLine = raw
    .split('\n')
    .map((line) => line.trim())
    .find((line) => line.length > 0);
  return firstLine ?? 'unknown error';
}
function createUpdateService(deps) {
  const inFlightBySource = new Map();
  async function runCheck(request) {
    const now = deps.now();
    const config = deps.getConfig();
    const channel = config.channel;
    const state = await deps.readState();
    const isAutomatic = request.source === 'automatic';
    if (isAutomatic && !request.force && shouldSkipAutomaticCheck(config, state, now)) {
      return { status: 'skipped' };
    }
    try {
      const appUpdate = await deps.checkAppUpdate(channel).catch((error) => {
        if (isAutomatic) {
          deps.log(`App update metadata check failed: ${summarizeError(error)}`);
        }
        return {
          available: false,
          version: deps.getCurrentVersion(),
        };
      });
      const shouldFetchReleaseMetadata =
        deps.shouldFetchReleaseMetadata?.({ request, channel, appUpdate }) ?? true;
      const release = shouldFetchReleaseMetadata
        ? await deps.fetchLatestStableRelease(channel).catch((error) => {
            deps.log(`GitHub release update check failed: ${summarizeError(error)}`);
            return null;
          })
        : null;
      const currentVersion = deps.getCurrentVersion();
      const latest = getBestLatestVersion(currentVersion, appUpdate, release);
      if (isAutomatic) {
        const nextState = {
          ...state,
          lastAutomaticCheckAt: now,
        };
        if (latest.available && state.lastNotifiedVersion !== latest.version) {
          await deps.notifyUpdateAvailable(latest.version);
          nextState.lastNotifiedVersion = latest.version;
        }
        await deps.writeState(nextState);
      }
      if (!latest.available) {
        if (!isAutomatic) {
          await deps.showNoUpdateDialog(currentVersion);
        }
        return { status: 'up-to-date', version: currentVersion };
      }
      if (isAutomatic) {
        return { status: 'update-available', version: latest.version };
      }
      const choice = await deps.showUpdateAvailableDialog(latest.version);
      if (choice === 'close') {
        return { status: 'update-available', version: latest.version };
      }
      const canInstallAppUpdate = appUpdate.available && appUpdate.canUpdate !== false;
      let appUpdateApplied = false;
      if (canInstallAppUpdate) {
        await deps.downloadAppUpdate();
        appUpdateApplied = true;
      }
      const launcherResult = await deps.updateLauncher(request.launcherPath, channel, release);
      if (launcherResult.status === 'protected' && launcherResult.command) {
        deps.log(`Launcher update requires manual command: ${launcherResult.command}`);
      }
      if (!appUpdateApplied) {
        await deps.showManualUpdateRequiredDialog(latest.version);
        return { status: 'update-available', version: latest.version };
      }
      const restartChoice = await deps.showRestartDialog();
      if (restartChoice === 'restart') {
        await deps.quitAndInstall();
      }
      return { status: 'updated', version: latest.version };
    } catch (error) {
      const message = summarizeError(error);
      if (isAutomatic) {
        deps.log(`Automatic update check failed: ${message}`);
      } else {
        await deps.showUpdateFailedDialog(message);
      }
      return { status: 'failed', error: message };
    }
  }
  return {
    checkForUpdates(request) {
      const inFlight = inFlightBySource.get(request.source);
      if (inFlight) return inFlight;
      const nextInFlight = runCheck(request).finally(() => {
        inFlightBySource.delete(request.source);
      });
      inFlightBySource.set(request.source, nextInFlight);
      return nextInFlight;
    },
    startAutomaticChecks(options = {}) {
      const setTimeoutFn = deps.setTimeout ?? setTimeout;
      const setIntervalFn = deps.setInterval ?? setInterval;
      const startupDelayMs = options.startupDelayMs ?? 15_000;
      const pollIntervalMs = options.pollIntervalMs ?? 60 * 60 * 1000;
      setTimeoutFn(() => {
        void this.checkForUpdates({ source: 'automatic' });
      }, startupDelayMs);
      setIntervalFn(() => {
        void this.checkForUpdates({ source: 'automatic' });
      }, pollIntervalMs);
    },
  };
}
function createFileUpdateStateStore(statePath) {
  return {
    async readState() {
      try {
        return JSON.parse(await node_fs_1.default.promises.readFile(statePath, 'utf8'));
      } catch {
        return {};
      }
    },
    async writeState(state) {
      await node_fs_1.default.promises.mkdir(node_path_1.default.dirname(statePath), {
        recursive: true,
      });
      await node_fs_1.default.promises.writeFile(
        statePath,
        `${JSON.stringify(state, null, 2)}\n`,
        'utf8',
      );
    },
  };
}
//# sourceMappingURL=update-service.js.map
