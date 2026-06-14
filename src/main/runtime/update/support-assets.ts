import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import type { GitHubRelease } from './release-assets';
import { compareSemverLike, findReleaseAsset } from './release-assets';

const execFileAsync = promisify(execFile);
const THEME_RELATIVE_PATH = path.join('themes', 'subminer.rasi');
const PLUGIN_ENTRYPOINT_RELATIVE_PATH = path.join('plugin', 'subminer', 'main.lua');
const PLUGIN_VERSION_RELATIVE_PATH = path.join('plugin', 'subminer', 'version.lua');
const PLUGIN_DIR_RELATIVE_PATH = path.join('plugin', 'subminer');

export interface SupportAssetsUpdateResult {
  status: 'updated' | 'skipped' | 'protected' | 'hash-mismatch' | 'missing-asset';
  component?: 'theme' | 'plugin';
  path?: string;
  command?: string;
  message?: string;
}

function sha256(data: Buffer): string {
  return createHash('sha256').update(data).digest('hex');
}

function shellQuote(value: string): string {
  return `'${value.replace(/'/g, `'\\''`)}'`;
}

function parsePluginVersion(content: string): string | null {
  const match = content.match(/\bversion\s*=\s*["']([^"']+)["']/);
  return match?.[1] ?? null;
}

async function pathExists(targetPath: string): Promise<boolean> {
  return await fs.promises
    .access(targetPath)
    .then(() => true)
    .catch(() => false);
}

async function canWrite(targetPath: string): Promise<boolean> {
  return await fs.promises
    .access(targetPath, fs.constants.W_OK)
    .then(() => true)
    .catch(() => false);
}

async function readFileIfExists(targetPath: string): Promise<Buffer | null> {
  try {
    return await fs.promises.readFile(targetPath);
  } catch {
    return null;
  }
}

async function readInstalledPluginVersion(pluginDir: string): Promise<string | null> {
  try {
    return parsePluginVersion(
      await fs.promises.readFile(path.join(pluginDir, 'version.lua'), 'utf8'),
    );
  } catch {
    return null;
  }
}

async function detectManagedSupportAssetDataDirs(dataDirs: string[]): Promise<string[]> {
  const managedDataDirs: string[] = [];
  for (const dataDir of dataDirs) {
    const [hasTheme, hasPlugin] = await Promise.all([
      pathExists(path.join(dataDir, THEME_RELATIVE_PATH)),
      pathExists(path.join(dataDir, PLUGIN_ENTRYPOINT_RELATIVE_PATH)),
    ]);
    if (hasTheme || hasPlugin) {
      managedDataDirs.push(dataDir);
    }
  }
  return managedDataDirs;
}

async function replacePluginDir(sourcePluginDir: string, targetPluginDir: string): Promise<void> {
  const parentDir = path.dirname(targetPluginDir);
  const stagedDir = `${targetPluginDir}.next`;
  const backupDir = `${targetPluginDir}.bak`;
  const targetExists = await pathExists(targetPluginDir);

  await fs.promises.rm(stagedDir, { recursive: true, force: true });
  await fs.promises.rm(backupDir, { recursive: true, force: true });
  await fs.promises.mkdir(parentDir, { recursive: true });
  await fs.promises.cp(sourcePluginDir, stagedDir, { recursive: true });

  if (targetExists) {
    await fs.promises.rename(targetPluginDir, backupDir);
  }
  try {
    await fs.promises.rename(stagedDir, targetPluginDir);
  } catch (err) {
    if (targetExists) {
      await fs.promises.rename(backupDir, targetPluginDir).catch(() => {});
    }
    throw err;
  }
  await fs.promises.rm(backupDir, { recursive: true, force: true });
}

function makeSupportAssetResult(
  status: SupportAssetsUpdateResult['status'],
  component: SupportAssetsUpdateResult['component'],
  dataDir: string,
  message: string,
  command?: string,
): SupportAssetsUpdateResult {
  const result: SupportAssetsUpdateResult = {
    status,
    component,
    path: dataDir,
    message,
  };
  if (command) {
    result.command = command;
  }
  return result;
}

export function detectSupportAssetDataDirs(options: {
  platform: NodeJS.Platform;
  homeDir: string;
  xdgDataHome?: string;
}): string[] {
  if (options.platform === 'linux') {
    const xdgDataHome = options.xdgDataHome || path.posix.join(options.homeDir, '.local/share');
    return [
      path.posix.join(xdgDataHome, 'SubMiner'),
      '/usr/local/share/SubMiner',
      '/usr/share/SubMiner',
    ];
  }
  return [];
}

export function buildProtectedSupportAssetsCommand(assetUrl: string, dataDir: string): string {
  const quotedDir = shellQuote(dataDir);
  const quotedPluginDir = shellQuote(path.posix.join(dataDir, 'plugin/subminer'));
  const quotedStagedPluginDir = shellQuote(path.posix.join(dataDir, 'plugin/subminer.next'));
  const quotedBackupPluginDir = shellQuote(path.posix.join(dataDir, 'plugin/subminer.bak'));
  return [
    'tmp=$(mktemp -d)',
    'trap \'rm -rf "$tmp"\' EXIT',
    `curl -fSL ${shellQuote(assetUrl)} -o "$tmp/subminer-assets.tar.gz"`,
    'tar -xzf "$tmp/subminer-assets.tar.gz" -C "$tmp"',
    `sudo mkdir -p ${quotedDir}/themes`,
    `sudo cp "$tmp/assets/themes/subminer.rasi" ${quotedDir}/themes/subminer.rasi`,
    `sudo mkdir -p ${quotedDir}/plugin`,
    `sudo rm -rf ${quotedStagedPluginDir} ${quotedBackupPluginDir}`,
    `sudo cp -R "$tmp/plugin/subminer" ${quotedStagedPluginDir}`,
    `[ ! -e ${quotedPluginDir} ] || sudo mv ${quotedPluginDir} ${quotedBackupPluginDir}`,
    `sudo mv ${quotedStagedPluginDir} ${quotedPluginDir}`,
    `sudo rm -rf ${quotedBackupPluginDir}`,
  ].join(' && ');
}

export async function updateSupportAssetsFromRelease(options: {
  release: GitHubRelease | null;
  sha256Sums: Map<string, string>;
  downloadAsset: (url: string) => Promise<Buffer>;
  platform?: NodeJS.Platform;
  homeDir?: string;
  xdgDataHome?: string;
}): Promise<SupportAssetsUpdateResult[]> {
  if ((options.platform ?? process.platform) !== 'linux') {
    return [{ status: 'skipped', message: 'Support assets are only installed on Linux.' }];
  }
  if (!options.release) return [{ status: 'missing-asset', message: 'No release found.' }];
  const asset = findReleaseAsset(options.release, 'subminer-assets.tar.gz');
  if (!asset) {
    return [{ status: 'missing-asset', message: 'Release has no support asset archive.' }];
  }
  const expectedSha256 = options.sha256Sums.get('subminer-assets.tar.gz');
  if (!expectedSha256) {
    return [{ status: 'missing-asset', message: 'SHA256SUMS.txt has no support asset entry.' }];
  }

  const dataDirs = detectSupportAssetDataDirs({
    platform: options.platform ?? process.platform,
    homeDir: options.homeDir ?? os.homedir(),
    xdgDataHome: options.xdgDataHome ?? process.env.XDG_DATA_HOME,
  });
  const managedDataDirs = await detectManagedSupportAssetDataDirs(dataDirs);
  if (managedDataDirs.length === 0) {
    return [{ status: 'skipped', message: 'No managed SubMiner support-asset install detected.' }];
  }

  const results: SupportAssetsUpdateResult[] = [];
  const writableDataDirs: string[] = [];
  for (const dataDir of managedDataDirs) {
    if (!fs.existsSync(dataDir) || !fs.statSync(dataDir).isDirectory()) {
      results.push(
        makeSupportAssetResult(
          'skipped',
          'theme',
          dataDir,
          'Support asset path is not a directory.',
        ),
        makeSupportAssetResult(
          'skipped',
          'plugin',
          dataDir,
          'Support asset path is not a directory.',
        ),
      );
      continue;
    }

    if (await canWrite(dataDir)) {
      writableDataDirs.push(dataDir);
      continue;
    }

    const command = buildProtectedSupportAssetsCommand(asset.browser_download_url, dataDir);
    results.push(
      makeSupportAssetResult(
        'protected',
        'theme',
        dataDir,
        'Theme install requires a manual command.',
        command,
      ),
      makeSupportAssetResult(
        'protected',
        'plugin',
        dataDir,
        'Plugin install requires a manual command.',
        command,
      ),
    );
  }

  if (writableDataDirs.length === 0) {
    return results;
  }

  const archive = await options.downloadAsset(asset.browser_download_url);
  const actualSha256 = sha256(archive);
  if (actualSha256 !== expectedSha256.toLowerCase()) {
    return [
      ...results,
      {
        status: 'hash-mismatch',
        message: `Expected ${expectedSha256}, got ${actualSha256}.`,
      },
    ];
  }

  const tempDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'subminer-assets-'));
  try {
    const archivePath = path.join(tempDir, 'subminer-assets.tar.gz');
    await fs.promises.writeFile(archivePath, archive);
    await execFileAsync('tar', ['-xzf', archivePath, '-C', tempDir]);

    const themeSourcePath = path.join(tempDir, 'assets/themes/subminer.rasi');
    if (!(await pathExists(themeSourcePath))) {
      return [
        { status: 'missing-asset', message: 'Support asset archive is missing the rofi theme.' },
      ];
    }
    const themeBytes = await fs.promises.readFile(themeSourcePath);

    const sourcePluginDir = path.join(tempDir, PLUGIN_DIR_RELATIVE_PATH);
    const sourcePluginEntrypoint = path.join(tempDir, PLUGIN_ENTRYPOINT_RELATIVE_PATH);
    if (!(await pathExists(sourcePluginEntrypoint))) {
      return [
        {
          status: 'missing-asset',
          message: 'Support asset archive is missing the runtime plugin.',
        },
      ];
    }
    const sourcePluginVersion = parsePluginVersion(
      await fs.promises
        .readFile(path.join(tempDir, PLUGIN_VERSION_RELATIVE_PATH), 'utf8')
        .catch(() => ''),
    );
    if (!sourcePluginVersion) {
      return [
        {
          status: 'missing-asset',
          message: 'Support asset archive has no readable plugin version metadata.',
        },
      ];
    }

    for (const dataDir of writableDataDirs) {
      const targetThemePath = path.join(dataDir, THEME_RELATIVE_PATH);
      const existingThemeBytes = await readFileIfExists(targetThemePath);
      if (existingThemeBytes && Buffer.compare(existingThemeBytes, themeBytes) === 0) {
        results.push(
          makeSupportAssetResult('skipped', 'theme', dataDir, 'Theme already up to date.'),
        );
      } else {
        await fs.promises.mkdir(path.dirname(targetThemePath), { recursive: true });
        await fs.promises.writeFile(targetThemePath, themeBytes);
        results.push(
          makeSupportAssetResult(
            'updated',
            'theme',
            dataDir,
            existingThemeBytes ? 'Updated theme.' : 'Installed theme.',
          ),
        );
      }

      const targetPluginDir = path.join(dataDir, PLUGIN_DIR_RELATIVE_PATH);
      const targetPluginEntrypoint = path.join(dataDir, PLUGIN_ENTRYPOINT_RELATIVE_PATH);
      const installedPluginVersion = await readInstalledPluginVersion(targetPluginDir);
      const installedEntrypointExists = await pathExists(targetPluginEntrypoint);
      const shouldInstallPlugin =
        !installedEntrypointExists ||
        !installedPluginVersion ||
        compareSemverLike(sourcePluginVersion, installedPluginVersion) > 0;
      if (!shouldInstallPlugin) {
        results.push(
          makeSupportAssetResult('skipped', 'plugin', dataDir, 'Plugin already up to date.'),
        );
        continue;
      }

      await replacePluginDir(sourcePluginDir, targetPluginDir);
      results.push(
        makeSupportAssetResult(
          'updated',
          'plugin',
          dataDir,
          installedEntrypointExists ? 'Updated plugin.' : 'Installed plugin.',
        ),
      );
    }

    return results;
  } finally {
    await fs.promises.rm(tempDir, { recursive: true, force: true });
  }
}
