import fs from 'node:fs';
import path from 'node:path';
import type { UpdateChannel, UpdatesConfig } from '../../../types/config';
import type { GitHubRelease } from './release-assets';
import { compareSemverLike, parseReleaseVersion } from './release-assets';

export interface UpdateState {
  lastAutomaticCheckAt?: number;
  lastNotifiedVersion?: string;
}

export type UpdateCheckSource = 'manual' | 'automatic' | 'launcher';

export interface UpdateCheckRequest {
  source: UpdateCheckSource;
  force?: boolean;
  launcherPath?: string;
  installWhenAvailable?: boolean;
}

export type UpdateCheckStatus =
  | 'up-to-date'
  | 'update-available'
  | 'updated'
  | 'skipped'
  | 'failed';

export interface UpdateCheckResult {
  status: UpdateCheckStatus;
  version?: string;
  error?: string;
}

type AppUpdateMetadata = {
  available: boolean;
  version: string;
  canUpdate?: boolean;
};

export interface UpdateServiceDeps {
  getConfig: () => Required<UpdatesConfig>;
  getCurrentVersion: () => string;
  now: () => number;
  readState: () => Promise<UpdateState>;
  writeState: (state: UpdateState) => Promise<void>;
  checkAppUpdate: (channel: UpdateChannel) => Promise<AppUpdateMetadata>;
  shouldFetchReleaseMetadata?: (input: {
    request: UpdateCheckRequest;
    channel: UpdateChannel;
    appUpdate: AppUpdateMetadata;
  }) => boolean;
  fetchLatestStableRelease: (channel: UpdateChannel) => Promise<GitHubRelease | null>;
  updateLauncher: (
    launcherPath?: string,
    channel?: UpdateChannel,
    release?: GitHubRelease | null,
  ) => Promise<{ status: string; command?: string }>;
  showNoUpdateDialog: (version: string) => Promise<void>;
  showUpdateAvailableDialog: (version: string) => Promise<'update' | 'close'>;
  showUpdateFailedDialog: (message: string) => Promise<void>;
  showManualUpdateRequiredDialog: (version: string) => Promise<void>;
  downloadAppUpdate: () => Promise<void>;
  showRestartDialog: () => Promise<'restart' | 'later'>;
  quitAndInstall: () => void | Promise<void>;
  notifyUpdateAvailable: (version: string) => Promise<void>;
  log: (message: string) => void;
  setTimeout?: (callback: () => void, delayMs: number) => unknown;
  setInterval?: (callback: () => void, delayMs: number) => unknown;
}

function getBestLatestVersion(
  currentVersion: string,
  appUpdate: { available: boolean; version: string },
  release: GitHubRelease | null,
): { available: boolean; version: string } {
  const releaseVersion = parseReleaseVersion(release);
  const candidates = [appUpdate.version, releaseVersion].filter(
    (value): value is string => typeof value === 'string' && value.length > 0,
  );
  const latest = candidates.reduce(
    (best, candidate) => (compareSemverLike(candidate, best) > 0 ? candidate : best),
    currentVersion,
  );
  return {
    available: appUpdate.available || compareSemverLike(latest, currentVersion) > 0,
    version: latest,
  };
}

function shouldSkipAutomaticCheck(
  config: Required<UpdatesConfig>,
  state: UpdateState,
  now: number,
) {
  if (!config.enabled) return true;
  if (!state.lastAutomaticCheckAt) return false;
  const intervalMs = Math.max(1, config.checkIntervalHours) * 60 * 60 * 1000;
  return now - state.lastAutomaticCheckAt < intervalMs;
}

function summarizeError(error: unknown): string {
  const raw = error instanceof Error ? error.message : String(error);
  const firstLine = raw
    .split('\n')
    .map((line) => line.trim())
    .find((line) => line.length > 0);
  return firstLine ?? 'unknown error';
}

export function createUpdateService(deps: UpdateServiceDeps) {
  const inFlightBySource = new Map<UpdateCheckSource, Promise<UpdateCheckResult>>();

  async function runCheck(request: UpdateCheckRequest): Promise<UpdateCheckResult> {
    const now = deps.now();
    const config = deps.getConfig();
    const channel = config.channel;
    const state = await deps.readState();
    const isAutomatic = request.source === 'automatic';

    if (isAutomatic && !request.force && shouldSkipAutomaticCheck(config, state, now)) {
      return { status: 'skipped' };
    }

    try {
      const appUpdate: AppUpdateMetadata = await deps.checkAppUpdate(channel).catch((error) => {
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
        const nextState: UpdateState = {
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

      if (!request.installWhenAvailable) {
        const choice = await deps.showUpdateAvailableDialog(latest.version);
        if (choice === 'close') {
          return { status: 'update-available', version: latest.version };
        }
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
    checkForUpdates(request: UpdateCheckRequest): Promise<UpdateCheckResult> {
      const inFlight = inFlightBySource.get(request.source);
      if (inFlight) return inFlight;
      const nextInFlight = runCheck(request).finally(() => {
        inFlightBySource.delete(request.source);
      });
      inFlightBySource.set(request.source, nextInFlight);
      return nextInFlight;
    },
    startAutomaticChecks(options: { startupDelayMs?: number; pollIntervalMs?: number } = {}): void {
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

export function createFileUpdateStateStore(statePath: string): {
  readState: () => Promise<UpdateState>;
  writeState: (state: UpdateState) => Promise<void>;
} {
  return {
    async readState(): Promise<UpdateState> {
      try {
        return JSON.parse(await fs.promises.readFile(statePath, 'utf8')) as UpdateState;
      } catch {
        return {};
      }
    },
    async writeState(state: UpdateState): Promise<void> {
      await fs.promises.mkdir(path.dirname(statePath), { recursive: true });
      await fs.promises.writeFile(statePath, `${JSON.stringify(state, null, 2)}\n`, 'utf8');
    },
  };
}
