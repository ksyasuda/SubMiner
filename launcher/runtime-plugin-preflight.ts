import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { log as launcherLog } from './log.js';
import { runAppCommandCaptureOutput, resolveLauncherRuntimePluginPath } from './mpv.js';
import { nowMs } from './time.js';
import { sleep } from './util.js';
import { detectInstalledMpvPlugin } from '../src/main/runtime/first-run-setup-plugin.js';
import {
  resolveManagedLinuxRuntimePluginPaths,
  type EnsureLinuxRuntimePluginAssetsResult,
} from '../src/main/runtime/linux-runtime-plugin-assets.js';

const RESPONSE_TIMEOUT_MS = 30_000;

type PreflightLog = (
  level: 'debug' | 'info' | 'warn' | 'error',
  configured: 'debug' | 'info' | 'warn' | 'error',
  message: string,
) => void;

type EnsureLinuxRuntimePluginAvailableOptions = {
  appPath?: string;
  scriptPath?: string;
  logLevel?: 'debug' | 'info' | 'warn' | 'error';
  platform?: NodeJS.Platform;
  homeDir?: string;
  xdgConfigHome?: string;
  xdgDataHome?: string;
  appDataDir?: string;
  detectInstalledPlugin?: () => boolean;
  resolveRuntimePluginPath?: () => string | null;
  installManagedPluginAssets?: () => Promise<EnsureLinuxRuntimePluginAssetsResult>;
  log?: PreflightLog;
};

type RuntimePluginPreflightResponse = {
  ok: boolean;
  status: 'installed' | 'already-present' | 'failed';
  path?: string;
  error?: string;
};

function resolveConfiguredLogLevel(
  logLevel: EnsureLinuxRuntimePluginAvailableOptions['logLevel'],
): 'debug' | 'info' | 'warn' | 'error' {
  return logLevel ?? 'warn';
}

async function waitForInstallResponse(
  responsePath: string,
): Promise<RuntimePluginPreflightResponse | null> {
  const deadline = nowMs() + RESPONSE_TIMEOUT_MS;
  while (nowMs() < deadline) {
    try {
      if (fs.existsSync(responsePath)) {
        return JSON.parse(fs.readFileSync(responsePath, 'utf8')) as RuntimePluginPreflightResponse;
      }
    } catch {
      // retry until timeout
    }
    await sleep(100);
  }
  return null;
}

type InstallManagedPluginAssetsViaAppDeps = {
  runAppCommandCaptureOutput?: typeof runAppCommandCaptureOutput;
  waitForInstallResponse?: typeof waitForInstallResponse;
};

export async function installManagedPluginAssetsViaApp(
  options: {
    appPath: string;
    logLevel?: 'debug' | 'info' | 'warn' | 'error';
  },
  deps: InstallManagedPluginAssetsViaAppDeps = {},
): Promise<EnsureLinuxRuntimePluginAssetsResult> {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'subminer-runtime-plugin-'));
  const responsePath = path.join(tempDir, 'response.json');
  const runAppCommand = deps.runAppCommandCaptureOutput ?? runAppCommandCaptureOutput;
  const waitForResponse = deps.waitForInstallResponse ?? waitForInstallResponse;
  try {
    const appArgs = [
      '--ensure-linux-runtime-plugin-assets',
      '--ensure-linux-runtime-plugin-assets-response-path',
      responsePath,
    ];
    const result = runAppCommand(options.appPath, appArgs);
    if (result.error) {
      return {
        ok: false,
        status: 'failed',
        error: result.error.message,
      };
    }
    if (result.status !== 0) {
      const stderr = result.stderr.trim();
      const stdout = result.stdout.trim();
      return {
        ok: false,
        status: 'failed',
        error:
          stderr ||
          stdout ||
          `Linux runtime plugin asset install command exited with status ${result.status}.`,
      };
    }
    const response = await waitForResponse(responsePath);
    if (response) {
      return response;
    }
    const stderr = result.stderr.trim();
    const stdout = result.stdout.trim();
    return {
      ok: false,
      status: 'failed',
      error:
        stderr ||
        stdout ||
        `Timed out waiting for Linux runtime plugin asset response after app exit status ${result.status}.`,
    };
  } finally {
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
}

export async function ensureLinuxRuntimePluginAvailable(
  options: EnsureLinuxRuntimePluginAvailableOptions,
): Promise<void> {
  const platform = options.platform ?? process.platform;
  if (platform !== 'linux') {
    return;
  }

  const configuredLogLevel = resolveConfiguredLogLevel(options.logLevel);
  const log = options.log ?? launcherLog;
  const homeDir = options.homeDir ?? os.homedir();
  const detectInstalledPlugin =
    options.detectInstalledPlugin ??
    (() =>
      detectInstalledMpvPlugin({
        platform,
        homeDir,
        xdgConfigHome: options.xdgConfigHome ?? process.env.XDG_CONFIG_HOME,
        appDataDir: options.appDataDir ?? process.env.APPDATA,
      }).installed);
  if (detectInstalledPlugin()) {
    return;
  }

  const resolveRuntimePluginPath =
    options.resolveRuntimePluginPath ??
    (() => {
      if (!options.appPath) return null;
      return resolveLauncherRuntimePluginPath({
        appPath: options.appPath,
        scriptPath: options.scriptPath,
        platform,
        homeDir,
        env: process.env,
      });
    });
  if (resolveRuntimePluginPath()) {
    return;
  }

  log(
    'info',
    configuredLogLevel,
    'Linux runtime plugin assets missing; installing managed plugin assets.',
  );
  const installManagedPluginAssets =
    options.installManagedPluginAssets ??
    (() => {
      if (!options.appPath) {
        throw new Error(
          'Linux managed runtime plugin assets could not be installed. Launch aborted before starting mpv.',
        );
      }
      return installManagedPluginAssetsViaApp({
        appPath: options.appPath,
        logLevel: options.logLevel,
      });
    });
  const installResult = await installManagedPluginAssets();
  if (!installResult.ok) {
    const message = installResult.error || 'Unknown Linux runtime plugin asset install failure.';
    log('warn', configuredLogLevel, `Managed Linux runtime plugin install failed: ${message}`);
    throw new Error(message);
  }

  log(
    'info',
    configuredLogLevel,
    `Managed Linux runtime plugin installed: ${installResult.path ?? 'unknown path'}`,
  );
  const runtimePluginPath = resolveRuntimePluginPath();
  if (runtimePluginPath) {
    return;
  }

  const managedPaths = resolveManagedLinuxRuntimePluginPaths({
    homeDir,
    xdgDataHome: options.xdgDataHome ?? process.env.XDG_DATA_HOME,
  });
  const message =
    `Linux managed runtime plugin assets could not be installed. ` +
    `Checked path: ${managedPaths.pluginEntrypointPath}. ` +
    'Launch aborted before starting mpv.';
  log('warn', configuredLogLevel, message);
  throw new Error(message);
}
