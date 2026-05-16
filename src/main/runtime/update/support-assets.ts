import { createHash } from 'node:crypto';
import { execFile } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { promisify } from 'node:util';
import type { GitHubRelease } from './release-assets';
import { findReleaseAsset } from './release-assets';

const execFileAsync = promisify(execFile);

export interface SupportAssetsUpdateResult {
  status: 'updated' | 'skipped' | 'protected' | 'hash-mismatch' | 'missing-asset';
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

export function detectSupportAssetDataDirs(options: {
  platform: NodeJS.Platform;
  homeDir: string;
  xdgDataHome?: string;
}): string[] {
  if (options.platform === 'linux') {
    const xdgDataHome = options.xdgDataHome || path.join(options.homeDir, '.local/share');
    return [path.join(xdgDataHome, 'SubMiner'), '/usr/local/share/SubMiner', '/usr/share/SubMiner'];
  }
  return [];
}

export function buildProtectedSupportAssetsCommand(assetUrl: string, dataDir: string): string {
  const quotedDir = shellQuote(dataDir);
  return [
    'tmp=$(mktemp -d)',
    'trap \'rm -rf "$tmp"\' EXIT',
    `curl -fSL ${shellQuote(assetUrl)} -o "$tmp/subminer-assets.tar.gz"`,
    'tar -xzf "$tmp/subminer-assets.tar.gz" -C "$tmp"',
    `sudo mkdir -p ${quotedDir}/themes`,
    `sudo cp "$tmp/assets/themes/subminer.rasi" ${quotedDir}/themes/subminer.rasi`,
  ].join(' && ');
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
  if (!asset) return [{ status: 'missing-asset', message: 'Release has no rofi theme asset.' }];
  const expectedSha256 = options.sha256Sums.get('subminer-assets.tar.gz');
  if (!expectedSha256) {
    return [{ status: 'missing-asset', message: 'SHA256SUMS.txt has no rofi theme entry.' }];
  }

  const dataDirs = detectSupportAssetDataDirs({
    platform: options.platform ?? process.platform,
    homeDir: options.homeDir ?? os.homedir(),
    xdgDataHome: options.xdgDataHome ?? process.env.XDG_DATA_HOME,
  });
  const existingDataDirs: string[] = [];
  for (const dataDir of dataDirs) {
    const hasTheme = await pathExists(path.join(dataDir, 'themes/subminer.rasi'));
    if (hasTheme) existingDataDirs.push(dataDir);
  }
  if (existingDataDirs.length === 0) {
    return [{ status: 'skipped', message: 'No existing rofi theme install detected.' }];
  }

  const protectedResults: SupportAssetsUpdateResult[] = existingDataDirs
    .filter((dataDir) => !fs.existsSync(dataDir) || !fs.statSync(dataDir).isDirectory())
    .map((dataDir) => ({
      status: 'skipped' as const,
      path: dataDir,
      message: 'Support asset path is not a directory.',
    }));
  const writableDataDirs: string[] = [];
  for (const dataDir of existingDataDirs) {
    if (await canWrite(dataDir)) {
      writableDataDirs.push(dataDir);
    } else {
      protectedResults.push({
        status: 'protected',
        path: dataDir,
        command: buildProtectedSupportAssetsCommand(asset.browser_download_url, dataDir),
      });
    }
  }
  if (writableDataDirs.length === 0) return protectedResults;

  const archive = await options.downloadAsset(asset.browser_download_url);
  const actualSha256 = sha256(archive);
  if (actualSha256 !== expectedSha256.toLowerCase()) {
    return [
      ...protectedResults,
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
    const results: SupportAssetsUpdateResult[] = [...protectedResults];
    for (const dataDir of writableDataDirs) {
      const targetThemePath = path.join(dataDir, 'themes/subminer.rasi');
      if (await pathExists(targetThemePath)) {
        await fs.promises.copyFile(
          path.join(tempDir, 'assets/themes/subminer.rasi'),
          targetThemePath,
        );
      }
      results.push({ status: 'updated', path: dataDir });
    }
    return results;
  } finally {
    await fs.promises.rm(tempDir, { recursive: true, force: true });
  }
}
