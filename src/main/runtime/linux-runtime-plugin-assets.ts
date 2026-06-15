import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { resolvePackagedFirstRunPluginAssets } from './first-run-setup-plugin';

export interface ManagedLinuxRuntimePluginPaths {
  dataDir: string;
  rootDir: string;
  pluginDir: string;
  pluginEntrypointPath: string;
  pluginConfigPath: string;
  themePath: string;
}

export interface EnsureLinuxRuntimePluginAssetsResult {
  ok: boolean;
  status: 'installed' | 'already-present' | 'failed';
  path?: string;
  error?: string;
}

interface RuntimePluginAssetSources {
  pluginDirSource?: string;
  pluginConfigSource?: string;
  themeSourcePath?: string;
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
  const explicitXdgDataHome = options.xdgDataHome?.trim();
  const envXdgDataHome = process.env.XDG_DATA_HOME?.trim();
  const xdgDataHome =
    explicitXdgDataHome || envXdgDataHome || pathModule.join(homeDir, '.local', 'share');
  const dataDir = pathModule.join(xdgDataHome, 'SubMiner');
  const rootDir = pathModule.join(dataDir, 'plugin');
  const pluginDir = pathModule.join(rootDir, 'subminer');
  return {
    dataDir,
    rootDir,
    pluginDir,
    pluginEntrypointPath: pathModule.join(pluginDir, 'main.lua'),
    pluginConfigPath: pathModule.join(rootDir, 'subminer.conf'),
    themePath: pathModule.join(dataDir, 'themes', 'subminer.rasi'),
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

function resolveBundledThemePath(options: {
  dirname: string;
  appPath: string;
  resourcesPath: string;
  existsSync: (candidate: string) => boolean;
}): string | null {
  const roots = [
    path.join(options.resourcesPath, 'assets'),
    path.join(options.resourcesPath, 'app.asar', 'assets'),
    path.join(options.appPath, 'assets'),
    path.join(options.dirname, '..', 'assets'),
    path.join(options.dirname, '..', '..', 'assets'),
    path.join(options.dirname, '..', '..', '..', 'assets'),
  ];

  for (const root of roots) {
    const candidate = path.join(root, 'themes', 'subminer.rasi');
    if (options.existsSync(candidate)) return candidate;
  }

  return null;
}

function resolveBundledAssetsDefault(
  existsSync: (candidate: string) => boolean,
): RuntimePluginAssetSources {
  const resourcesPath = process.resourcesPath ?? path.dirname(process.execPath);
  const pluginAssets = resolvePackagedFirstRunPluginAssets({
    dirname: __dirname,
    appPath: process.execPath,
    resourcesPath,
    existsSync,
  });

  const themeSourcePath = resolveBundledThemePath({
    dirname: __dirname,
    appPath: process.execPath,
    resourcesPath,
    existsSync,
  });

  return {
    ...(pluginAssets ?? {}),
    ...(themeSourcePath ? { themeSourcePath } : {}),
  };
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
  const pluginAssetsExist =
    existsSync(managedPaths.pluginEntrypointPath) && existsSync(managedPaths.pluginConfigPath);
  const themeExists = existsSync(managedPaths.themePath);
  if (pluginAssetsExist && themeExists) {
    return {
      ok: true,
      status: 'already-present',
      path: managedPaths.pluginEntrypointPath,
    };
  }

  const bundledAssets =
    (options.resolveBundledAssets
      ? options.resolveBundledAssets()
      : resolveBundledAssetsDefault(existsSync)) ?? {};

  const shouldInstallPluginAssets = !pluginAssetsExist;
  const shouldInstallTheme = !themeExists;
  if (
    shouldInstallPluginAssets &&
    (!bundledAssets.pluginDirSource || !bundledAssets.pluginConfigSource)
  ) {
    return {
      ok: false,
      status: 'failed',
      error: 'Bundled Linux runtime plugin assets were not found.',
    };
  }
  if (shouldInstallTheme && !bundledAssets.themeSourcePath) {
    return {
      ok: false,
      status: 'failed',
      error: 'Bundled Linux runtime theme asset was not found.',
    };
  }

  const stagingSuffix = `${process.pid}-${Date.now()}`;
  const stagedPluginDir = pathModule.join(managedPaths.rootDir, `.subminer-stage-${stagingSuffix}`);
  const stagedPluginConfigPath = pathModule.join(
    managedPaths.rootDir,
    `.subminer.conf-stage-${stagingSuffix}`,
  );
  const stagedThemePath = pathModule.join(
    pathModule.dirname(managedPaths.themePath),
    `.subminer.rasi-stage-${stagingSuffix}`,
  );
  let pluginDirInstalled = false;
  let pluginConfigInstalled = false;
  let themeInstalled = false;

  try {
    if (shouldInstallPluginAssets) {
      const pluginDirSource = bundledAssets.pluginDirSource;
      const pluginConfigSource = bundledAssets.pluginConfigSource;
      if (!pluginDirSource || !pluginConfigSource) {
        throw new Error('Bundled Linux runtime plugin assets were not found.');
      }
      await mkdir(managedPaths.rootDir, { recursive: true });
      await copyDirectoryRecursive(pluginDirSource, stagedPluginDir, {
        mkdir,
        readdir,
        copyFile,
        pathModule,
      });
      await copyFile(pluginConfigSource, stagedPluginConfigPath);
    }
    if (shouldInstallTheme) {
      const themeSourcePath = bundledAssets.themeSourcePath;
      if (!themeSourcePath) {
        throw new Error('Bundled Linux runtime theme asset was not found.');
      }
      await mkdir(pathModule.dirname(managedPaths.themePath), { recursive: true });
      await copyFile(themeSourcePath, stagedThemePath);
    }
    if (shouldInstallPluginAssets) {
      await rm(managedPaths.pluginDir, { recursive: true, force: true });
      await rm(managedPaths.pluginConfigPath, { force: true });
      await rename(stagedPluginDir, managedPaths.pluginDir);
      pluginDirInstalled = true;
      await rename(stagedPluginConfigPath, managedPaths.pluginConfigPath);
      pluginConfigInstalled = true;
    }
    if (shouldInstallTheme) {
      await rm(managedPaths.themePath, { force: true });
      await rename(stagedThemePath, managedPaths.themePath);
      themeInstalled = true;
    }

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
    if (themeInstalled) {
      await rm(managedPaths.themePath, { force: true }).catch(() => {});
    }
    await rm(stagedPluginDir, { recursive: true, force: true }).catch(() => {});
    await rm(stagedPluginConfigPath, { force: true }).catch(() => {});
    await rm(stagedThemePath, { force: true }).catch(() => {});
    return {
      ok: false,
      status: 'failed',
      error: errorMessage(error),
    };
  }
}
