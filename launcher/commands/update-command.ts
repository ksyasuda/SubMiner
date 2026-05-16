import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import packageJson from '../../package.json';
import { runAppCommandCaptureOutput } from '../mpv.js';
import { log as launcherLog } from '../log.js';
import { nowMs } from '../time.js';
import { sleep } from '../util.js';
import type { LauncherCommandContext } from './context.js';
import { readLauncherMainConfigObject } from '../config/shared-config-reader.js';
import type { UpdateChannel } from '../../src/types/config.js';
import { updateAppImageFromRelease } from '../../src/main/runtime/update/appimage-updater.js';
import { updateLauncherFromRelease } from '../../src/main/runtime/update/launcher-updater.js';
import {
  compareSemverLike,
  fetchLatestStableRelease,
  fetchReleaseAssetBuffer,
  fetchReleaseAssetText,
  findReleaseAsset,
  parseReleaseVersion,
  parseSha256Sums,
  type FetchLike,
} from '../../src/main/runtime/update/release-assets.js';
import { updateSupportAssetsFromRelease } from '../../src/main/runtime/update/support-assets.js';

type UpdateCommandResponse = {
  ok: boolean;
  status?: string;
  version?: string;
  error?: string;
};

type DirectReleaseUpdateRequest = {
  appPath: string;
  launcherPath: string;
  channel: UpdateChannel;
};

type DirectReleaseUpdateResult = {
  appImage: { status: string; command?: string; message?: string };
  launcher: { status: string; command?: string; message?: string };
  supportAssets: Array<{ status: string; command?: string; message?: string }>;
};

type UpdateCommandDeps = {
  createTempDir: (prefix: string) => string;
  joinPath: (...parts: string[]) => string;
  runAppCommandCaptureOutput: (
    appPath: string,
    appArgs: string[],
  ) => { status: number; stdout: string; stderr: string; error?: Error };
  waitForUpdateResponse: (responsePath: string) => Promise<UpdateCommandResponse>;
  removeDir: (targetPath: string) => void;
  runDirectReleaseUpdate: (
    request: DirectReleaseUpdateRequest,
  ) => Promise<DirectReleaseUpdateResult>;
  readMainConfig: () => Record<string, unknown> | null;
  log: typeof launcherLog;
};

const UPDATE_RESPONSE_TIMEOUT_MS = 10 * 60 * 1000;
const CURRENT_VERSION = packageJson.version;

function getFetchForLauncherUpdater(): FetchLike {
  return globalThis.fetch.bind(globalThis) as FetchLike;
}

async function runDirectReleaseUpdate(
  request: DirectReleaseUpdateRequest,
): Promise<DirectReleaseUpdateResult> {
  const fetchForUpdater = getFetchForLauncherUpdater();
  const release = await fetchLatestStableRelease({
    fetch: fetchForUpdater,
    channel: request.channel,
  });
  const releaseVersion = parseReleaseVersion(release);
  if (releaseVersion && compareSemverLike(releaseVersion, CURRENT_VERSION) <= 0) {
    return {
      appImage: { status: 'up-to-date' },
      launcher: { status: 'up-to-date' },
      supportAssets: [{ status: 'up-to-date' }],
    };
  }

  const sumsAsset = release ? findReleaseAsset(release, 'SHA256SUMS.txt') : null;
  const sha256Sums =
    sumsAsset && release
      ? parseSha256Sums(
          await fetchReleaseAssetText(fetchForUpdater, sumsAsset.browser_download_url),
        )
      : new Map<string, string>();
  const downloadAsset = (url: string) => fetchReleaseAssetBuffer(fetchForUpdater, url);

  const [appImage, launcher, supportAssets] = await Promise.all([
    updateAppImageFromRelease({
      release,
      sha256Sums,
      appImagePath: request.appPath,
      downloadAsset,
    }),
    updateLauncherFromRelease({
      release,
      sha256Sums,
      launcherPath: request.launcherPath,
      downloadAsset,
    }),
    updateSupportAssetsFromRelease({
      release,
      sha256Sums,
      downloadAsset,
    }),
  ]);

  return { appImage, launcher, supportAssets };
}

function readUpdateChannel(root: Record<string, unknown> | null): UpdateChannel {
  const updates =
    root?.updates && typeof root.updates === 'object' && !Array.isArray(root.updates)
      ? (root.updates as Record<string, unknown>)
      : null;
  return updates?.channel === 'prerelease' ? 'prerelease' : 'stable';
}

function logUpdateResult(
  label: string,
  result: { status: string; command?: string; message?: string },
  configuredLogLevel: NonNullable<LauncherCommandContext['args']['logLevel']>,
  deps: Pick<UpdateCommandDeps, 'log'>,
): void {
  const displayStatus = result.status === 'up-to-date' ? 'up to date' : result.status;
  deps.log('info', configuredLogLevel, `${label} update: ${displayStatus}`);
  if (result.command) {
    deps.log(
      'warn',
      configuredLogLevel,
      `${label} update requires manual command: ${result.command}`,
    );
  } else if (result.message) {
    deps.log('warn', configuredLogLevel, `${label} update note: ${result.message}`);
  }
}

const defaultDeps: UpdateCommandDeps = {
  createTempDir: (prefix) => fs.mkdtempSync(path.join(os.tmpdir(), prefix)),
  joinPath: (...parts) => path.join(...parts),
  runAppCommandCaptureOutput: (appPath, appArgs) => runAppCommandCaptureOutput(appPath, appArgs),
  waitForUpdateResponse: async (responsePath) => {
    const deadline = nowMs() + UPDATE_RESPONSE_TIMEOUT_MS;
    while (nowMs() < deadline) {
      try {
        if (fs.existsSync(responsePath)) {
          return JSON.parse(fs.readFileSync(responsePath, 'utf8')) as UpdateCommandResponse;
        }
      } catch {
        // retry until timeout
      }
      await sleep(100);
    }
    return { ok: false, error: 'Timed out waiting for SubMiner update response.' };
  },
  removeDir: (targetPath) => {
    fs.rmSync(targetPath, { recursive: true, force: true });
  },
  runDirectReleaseUpdate,
  readMainConfig: readLauncherMainConfigObject,
  log: launcherLog,
};

export async function runUpdateCommand(
  context: LauncherCommandContext,
  deps: Partial<UpdateCommandDeps> = {},
): Promise<boolean> {
  const resolvedDeps: UpdateCommandDeps = { ...defaultDeps, ...deps };
  const { args, appPath, scriptPath } = context;
  if (!args.update || !appPath) {
    return false;
  }

  if (context.processAdapter.platform() === 'linux') {
    const result = await resolvedDeps.runDirectReleaseUpdate({
      appPath,
      launcherPath: scriptPath,
      channel: readUpdateChannel(resolvedDeps.readMainConfig()),
    });
    const logLevel = args.logLevel ?? 'warn';
    logUpdateResult('AppImage', result.appImage, logLevel, resolvedDeps);
    logUpdateResult('Launcher', result.launcher, logLevel, resolvedDeps);
    for (const supportResult of result.supportAssets) {
      logUpdateResult('Rofi theme', supportResult, logLevel, resolvedDeps);
    }
    return true;
  }

  const tempDir = resolvedDeps.createTempDir('subminer-update-');
  const responsePath = resolvedDeps.joinPath(tempDir, 'response.json');

  try {
    const result = resolvedDeps.runAppCommandCaptureOutput(appPath, [
      '--update',
      '--update-launcher-path',
      scriptPath,
      '--update-response-path',
      responsePath,
    ]);
    if (result.error) {
      throw result.error;
    }
    if (result.status !== 0) {
      throw new Error(`SubMiner update command exited with status ${result.status}.`);
    }
    const response = await resolvedDeps.waitForUpdateResponse(responsePath);
    if (!response.ok) {
      throw new Error(response.error || 'SubMiner update check failed.');
    }
    return true;
  } finally {
    resolvedDeps.removeDir(tempDir);
  }
}
