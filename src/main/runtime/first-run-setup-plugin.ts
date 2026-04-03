import fs from 'node:fs';
import path from 'node:path';
import { resolveDefaultMpvInstallPaths, type MpvInstallPaths } from '../../shared/setup-state';
import type { PluginInstallResult } from './first-run-setup-service';

function timestamp(): string {
  return new Date().toISOString().replaceAll(':', '-');
}

function backupExistingPath(targetPath: string): void {
  if (!fs.existsSync(targetPath)) return;
  fs.renameSync(targetPath, `${targetPath}.bak.${timestamp()}`);
}

function resolveLegacyPluginLoaderPath(installPaths: MpvInstallPaths): string {
  return path.join(installPaths.scriptsDir, 'subminer.lua');
}

function resolveLegacyPluginDebugLoaderPath(installPaths: MpvInstallPaths): string {
  return path.join(installPaths.scriptsDir, 'subminer-loader.lua');
}

function rewriteInstalledWindowsPluginConfig(configPath: string): void {
  const content = fs.readFileSync(configPath, 'utf8');
  const updated = content.replace(/^socket_path=.*$/m, 'socket_path=\\\\.\\pipe\\subminer-socket');
  if (updated !== content) {
    fs.writeFileSync(configPath, updated, 'utf8');
  }
}

function sanitizePluginConfigValue(value: string): string {
  return value.replace(/[\r\n]/g, '').trim();
}

function upsertPluginConfigLine(content: string, key: string, value: string): string {
  const normalizedValue = sanitizePluginConfigValue(value);
  const line = `${key}=${normalizedValue}`;
  const pattern = new RegExp(`^${key}=.*$`, 'm');
  if (pattern.test(content)) {
    return content.replace(pattern, line);
  }

  const suffix = content.endsWith('\n') || content.length === 0 ? '' : '\n';
  return `${content}${suffix}${line}\n`;
}

function rewriteInstalledPluginBinaryPath(configPath: string, binaryPath: string): boolean {
  const content = fs.readFileSync(configPath, 'utf8');
  const updated = upsertPluginConfigLine(content, 'binary_path', binaryPath);
  if (updated === content) {
    return false;
  }
  fs.writeFileSync(configPath, updated, 'utf8');
  return true;
}

function readInstalledPluginBinaryPath(configPath: string): string | null {
  const content = fs.readFileSync(configPath, 'utf8');
  const match = content.match(/^binary_path=(.*)$/m);
  if (!match) {
    return null;
  }
  const rawValue = match[1] ?? '';
  const value = sanitizePluginConfigValue(rawValue);
  return value.length > 0 ? value : null;
}

export function resolvePackagedFirstRunPluginAssets(deps: {
  dirname: string;
  appPath: string;
  resourcesPath: string;
  joinPath?: (...parts: string[]) => string;
  existsSync?: (candidate: string) => boolean;
}): { pluginDirSource: string; pluginConfigSource: string } | null {
  const joinPath = deps.joinPath ?? path.join;
  const existsSync = deps.existsSync ?? fs.existsSync;
  const roots = [
    joinPath(deps.resourcesPath, 'plugin'),
    joinPath(deps.resourcesPath, 'app.asar', 'plugin'),
    joinPath(deps.appPath, 'plugin'),
    joinPath(deps.dirname, '..', 'plugin'),
    joinPath(deps.dirname, '..', '..', 'plugin'),
  ];

  for (const root of roots) {
    const pluginDirSource = joinPath(root, 'subminer');
    const pluginConfigSource = joinPath(root, 'subminer.conf');
    if (
      existsSync(pluginDirSource) &&
      existsSync(pluginConfigSource) &&
      existsSync(joinPath(pluginDirSource, 'main.lua'))
    ) {
      return { pluginDirSource, pluginConfigSource };
    }
  }

  return null;
}

export function detectInstalledFirstRunPlugin(
  installPaths: MpvInstallPaths,
  deps?: { existsSync?: (candidate: string) => boolean },
): boolean {
  const existsSync = deps?.existsSync ?? fs.existsSync;
  return (
    existsSync(installPaths.pluginEntrypointPath) &&
    existsSync(installPaths.pluginDir) &&
    existsSync(installPaths.pluginConfigPath)
  );
}

export function installFirstRunPluginToDefaultLocation(options: {
  platform: NodeJS.Platform;
  homeDir: string;
  xdgConfigHome?: string;
  dirname: string;
  appPath: string;
  resourcesPath: string;
  binaryPath: string;
}): PluginInstallResult {
  const installPaths = resolveDefaultMpvInstallPaths(
    options.platform,
    options.homeDir,
    options.xdgConfigHome,
  );
  if (!installPaths.supported) {
    return {
      ok: false,
      pluginInstallStatus: 'failed',
      pluginInstallPathSummary: installPaths.mpvConfigDir,
      message: 'Automatic mpv plugin install is not supported on this platform yet.',
    };
  }

  const assets = resolvePackagedFirstRunPluginAssets({
    dirname: options.dirname,
    appPath: options.appPath,
    resourcesPath: options.resourcesPath,
  });
  if (!assets) {
    return {
      ok: false,
      pluginInstallStatus: 'failed',
      pluginInstallPathSummary: installPaths.mpvConfigDir,
      message: 'Packaged mpv plugin assets were not found.',
    };
  }

  fs.mkdirSync(installPaths.scriptsDir, { recursive: true });
  fs.mkdirSync(installPaths.scriptOptsDir, { recursive: true });
  backupExistingPath(resolveLegacyPluginLoaderPath(installPaths));
  backupExistingPath(resolveLegacyPluginDebugLoaderPath(installPaths));
  backupExistingPath(installPaths.pluginDir);
  backupExistingPath(installPaths.pluginConfigPath);
  fs.cpSync(assets.pluginDirSource, installPaths.pluginDir, { recursive: true });
  fs.copyFileSync(assets.pluginConfigSource, installPaths.pluginConfigPath);
  rewriteInstalledPluginBinaryPath(installPaths.pluginConfigPath, options.binaryPath);
  if (options.platform === 'win32') {
    rewriteInstalledWindowsPluginConfig(installPaths.pluginConfigPath);
  }

  return {
    ok: true,
    pluginInstallStatus: 'installed',
    pluginInstallPathSummary: installPaths.mpvConfigDir,
    message: `Installed mpv plugin to ${installPaths.mpvConfigDir}.`,
  };
}

export function syncInstalledFirstRunPluginBinaryPath(options: {
  platform: NodeJS.Platform;
  homeDir: string;
  xdgConfigHome?: string;
  binaryPath: string;
}): { updated: boolean; configPath: string | null } {
  const installPaths = resolveDefaultMpvInstallPaths(
    options.platform,
    options.homeDir,
    options.xdgConfigHome,
  );
  if (!installPaths.supported || !fs.existsSync(installPaths.pluginConfigPath)) {
    return { updated: false, configPath: null };
  }

  const configuredBinaryPath = readInstalledPluginBinaryPath(installPaths.pluginConfigPath);
  if (configuredBinaryPath) {
    return { updated: false, configPath: installPaths.pluginConfigPath };
  }

  const updated = rewriteInstalledPluginBinaryPath(
    installPaths.pluginConfigPath,
    options.binaryPath,
  );
  if (options.platform === 'win32') {
    rewriteInstalledWindowsPluginConfig(installPaths.pluginConfigPath);
  }
  return {
    updated,
    configPath: installPaths.pluginConfigPath,
  };
}
