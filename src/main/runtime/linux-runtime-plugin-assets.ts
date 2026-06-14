import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { resolvePackagedFirstRunPluginAssets } from './first-run-setup-plugin';

export interface ManagedLinuxRuntimePluginPaths {
  rootDir: string;
  pluginDir: string;
  pluginEntrypointPath: string;
  pluginConfigPath: string;
}

export interface EnsureLinuxRuntimePluginAssetsResult {
  ok: boolean;
  status: 'installed' | 'already-present' | 'failed';
  path?: string;
  error?: string;
}

interface RuntimePluginAssetSources {
  pluginDirSource: string;
  pluginConfigSource: string;
}

interface RuntimePluginDirentLike {
  name: string;
  isDirectory(): boolean;
}

interface EnsureLinuxRuntimePluginAssetsOptions {
  platform?: NodeJS.Platform;
  homeDir?: string;
  xdgDataHome?: string;
  pathModule?: typeof path;
  existsSync?: (candidate: string) => boolean;
  resolveBundledAssets?: () => RuntimePluginAssetSources | null;
  mkdir?: (targetPath: string, options: { recursive: true }) => Promise<void>;
  readdir?: (
    targetPath: string,
    options: { withFileTypes: true },
  ) => Promise<RuntimePluginDirentLike[]>;
  copyFile?: (sourcePath: string, targetPath: string) => Promise<void>;
  rename?: (fromPath: string, toPath: string) => Promise<void>;
  rm?: (targetPath: string, options: { recursive?: boolean; force?: boolean }) => Promise<void>;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function resolveManagedLinuxRuntimePluginPaths(options: {
  homeDir?: string;
  xdgDataHome?: string;
  pathModule?: typeof path;
}): ManagedLinuxRuntimePluginPaths {
  const pathModule = options.pathModule ?? path;
  const homeDir = options.homeDir ?? os.homedir();
  const xdgDataHome =
    options.xdgDataHome ?? process.env.XDG_DATA_HOME ?? pathModule.join(homeDir, '.local', 'share');
  const rootDir = pathModule.join(xdgDataHome, 'SubMiner', 'plugin');
  const pluginDir = pathModule.join(rootDir, 'subminer');
  return {
    rootDir,
    pluginDir,
    pluginEntrypointPath: pathModule.join(pluginDir, 'main.lua'),
    pluginConfigPath: pathModule.join(rootDir, 'subminer.conf'),
  };
}

async function copyDirectoryRecursive(
  sourceDir: string,
  targetDir: string,
  options: Required<
    Pick<EnsureLinuxRuntimePluginAssetsOptions, 'mkdir' | 'readdir' | 'copyFile' | 'pathModule'>
  >,
): Promise<void> {
  await options.mkdir(targetDir, { recursive: true });
  const entries = await options.readdir(sourceDir, { withFileTypes: true });
  for (const entry of entries) {
    const sourcePath = options.pathModule.join(sourceDir, entry.name);
    const targetPath = options.pathModule.join(targetDir, entry.name);
    if (entry.isDirectory()) {
      await copyDirectoryRecursive(sourcePath, targetPath, options);
      continue;
    }
    await options.copyFile(sourcePath, targetPath);
  }
}

function resolveBundledAssetsDefault(
  existsSync: (candidate: string) => boolean,
): RuntimePluginAssetSources | null {
  return resolvePackagedFirstRunPluginAssets({
    dirname: __dirname,
    appPath: process.execPath,
    resourcesPath: process.resourcesPath,
    existsSync,
  });
}

export async function ensureLinuxRuntimePluginAssets(
  options: EnsureLinuxRuntimePluginAssetsOptions = {},
): Promise<EnsureLinuxRuntimePluginAssetsResult> {
  const platform = options.platform ?? process.platform;
  if (platform !== 'linux') {
    return {
      ok: false,
      status: 'failed',
      error: 'Linux runtime plugin asset install is only supported on Linux.',
    };
  }

  const pathModule = options.pathModule ?? path;
  const existsSync = options.existsSync ?? fs.existsSync;
  const mkdir =
    options.mkdir ??
    (async (targetPath, mkdirOptions) => {
      await fs.promises.mkdir(targetPath, mkdirOptions);
    });
  const readdir =
    options.readdir ??
    ((targetPath, readdirOptions) =>
      fs.promises.readdir(targetPath, readdirOptions) as Promise<RuntimePluginDirentLike[]>);
  const copyFile =
    options.copyFile ?? ((sourcePath, targetPath) => fs.promises.copyFile(sourcePath, targetPath));
  const rename = options.rename ?? ((fromPath, toPath) => fs.promises.rename(fromPath, toPath));
  const rm = options.rm ?? ((targetPath, rmOptions) => fs.promises.rm(targetPath, rmOptions));

  const managedPaths = resolveManagedLinuxRuntimePluginPaths({
    homeDir: options.homeDir,
    xdgDataHome: options.xdgDataHome,
    pathModule,
  });
  if (existsSync(managedPaths.pluginEntrypointPath) && existsSync(managedPaths.pluginConfigPath)) {
    return {
      ok: true,
      status: 'already-present',
      path: managedPaths.pluginEntrypointPath,
    };
  }

  const bundledAssets = options.resolveBundledAssets
    ? options.resolveBundledAssets()
    : resolveBundledAssetsDefault(existsSync);
  if (!bundledAssets) {
    return {
      ok: false,
      status: 'failed',
      error: 'Bundled Linux runtime plugin assets were not found.',
    };
  }

  const stagingSuffix = `${process.pid}-${Date.now()}`;
  const stagedPluginDir = pathModule.join(managedPaths.rootDir, `.subminer-stage-${stagingSuffix}`);
  const stagedPluginConfigPath = pathModule.join(
    managedPaths.rootDir,
    `.subminer.conf-stage-${stagingSuffix}`,
  );
  let pluginDirInstalled = false;
  let pluginConfigInstalled = false;

  try {
    await mkdir(managedPaths.rootDir, { recursive: true });
    await copyDirectoryRecursive(bundledAssets.pluginDirSource, stagedPluginDir, {
      mkdir,
      readdir,
      copyFile,
      pathModule,
    });
    await copyFile(bundledAssets.pluginConfigSource, stagedPluginConfigPath);
    await rm(managedPaths.pluginDir, { recursive: true, force: true });
    await rm(managedPaths.pluginConfigPath, { force: true });
    await rename(stagedPluginDir, managedPaths.pluginDir);
    pluginDirInstalled = true;
    await rename(stagedPluginConfigPath, managedPaths.pluginConfigPath);
    pluginConfigInstalled = true;

    return {
      ok: true,
      status: 'installed',
      path: managedPaths.pluginEntrypointPath,
    };
  } catch (error) {
    if (pluginDirInstalled && !pluginConfigInstalled) {
      await rm(managedPaths.pluginDir, { recursive: true, force: true }).catch(() => {});
    }
    if (pluginConfigInstalled && !pluginDirInstalled) {
      await rm(managedPaths.pluginConfigPath, { force: true }).catch(() => {});
    }
    await rm(stagedPluginDir, { recursive: true, force: true }).catch(() => {});
    await rm(stagedPluginConfigPath, { force: true }).catch(() => {});
    return {
      ok: false,
      status: 'failed',
      error: errorMessage(error),
    };
  }
}
