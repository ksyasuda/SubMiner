import { createHash } from 'node:crypto';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import type { GitHubRelease } from './release-assets';
import { findReleaseAsset } from './release-assets';

type StatLike = {
  isFile: () => boolean;
  mode?: number;
};

export type LauncherUpdateStatus =
  | 'updated'
  | 'skipped'
  | 'protected'
  | 'hash-mismatch'
  | 'not-found'
  | 'missing-asset';

export interface LauncherUpdateResult {
  status: LauncherUpdateStatus;
  path?: string;
  command?: string;
  message?: string;
}

export interface LauncherUpdateFileSystem {
  readFile: (targetPath: string) => Promise<Buffer | string>;
  stat: (targetPath: string) => Promise<StatLike>;
  access: (targetPath: string) => Promise<void>;
  writeFile: (targetPath: string, data: Buffer) => Promise<void>;
  chmod: (targetPath: string, mode: number) => Promise<void>;
  rename: (fromPath: string, toPath: string) => Promise<void>;
  unlink: (targetPath: string) => Promise<void>;
}

export function looksLikeSubminerLauncher(content: Buffer | string): boolean {
  const text = Buffer.isBuffer(content) ? content.toString('utf8') : content;
  return (
    text.includes('SubMiner launcher') ||
    text.includes('Launch MPV with SubMiner') ||
    text.includes('SUBMINER_APPIMAGE_PATH') ||
    text.includes('SubMiner.app') ||
    text.includes('SubMiner.AppImage')
  );
}

export function buildProtectedLauncherUpdateCommand(
  assetUrl: string,
  launcherPath: string,
): string {
  return `sudo curl -fSL ${assetUrl} -o ${launcherPath} && sudo chmod +x ${launcherPath}`;
}

function sha256(data: Buffer): string {
  return createHash('sha256').update(data).digest('hex');
}

function defaultFs(): LauncherUpdateFileSystem {
  return {
    readFile: (targetPath) => fs.promises.readFile(targetPath),
    stat: (targetPath) => fs.promises.stat(targetPath),
    access: async (targetPath) => {
      await fs.promises.access(targetPath, fs.constants.W_OK);
    },
    writeFile: (targetPath, data) => fs.promises.writeFile(targetPath, data),
    chmod: (targetPath, mode) => fs.promises.chmod(targetPath, mode),
    rename: (fromPath, toPath) => fs.promises.rename(fromPath, toPath),
    unlink: async (targetPath) => {
      await fs.promises.unlink(targetPath).catch(() => undefined);
    },
  };
}

export async function updateLauncherAtPath(options: {
  launcherPath: string;
  assetUrl: string;
  expectedSha256: string;
  download: () => Promise<Buffer>;
  fs?: LauncherUpdateFileSystem;
}): Promise<LauncherUpdateResult> {
  const fsDeps = options.fs ?? defaultFs();
  let stat: StatLike;
  try {
    stat = await fsDeps.stat(options.launcherPath);
  } catch {
    return { status: 'not-found', path: options.launcherPath };
  }
  if (!stat.isFile()) {
    return { status: 'skipped', path: options.launcherPath, message: 'Launcher is not a file.' };
  }

  const existing = await fsDeps.readFile(options.launcherPath);
  if (!looksLikeSubminerLauncher(existing)) {
    return {
      status: 'skipped',
      path: options.launcherPath,
      message: 'Existing executable does not look like a SubMiner launcher.',
    };
  }

  try {
    await fsDeps.access(options.launcherPath);
  } catch {
    return {
      status: 'protected',
      path: options.launcherPath,
      command: buildProtectedLauncherUpdateCommand(options.assetUrl, options.launcherPath),
    };
  }

  const data = await options.download();
  const actualSha256 = sha256(data);
  if (actualSha256 !== options.expectedSha256.toLowerCase()) {
    return {
      status: 'hash-mismatch',
      path: options.launcherPath,
      message: `Expected ${options.expectedSha256}, got ${actualSha256}.`,
    };
  }

  const tempPath = path.join(path.dirname(options.launcherPath), '.subminer.update');
  try {
    await fsDeps.writeFile(tempPath, data);
    await fsDeps.chmod(tempPath, stat.mode ? stat.mode & 0o777 : 0o755);
    await fsDeps.rename(tempPath, options.launcherPath);
    return { status: 'updated', path: options.launcherPath };
  } catch (error) {
    await fsDeps.unlink(tempPath);
    throw error;
  }
}

export function detectLauncherCandidates(options: {
  platform: NodeJS.Platform;
  homeDir: string;
  explicitPath?: string;
}): string[] {
  const candidates: string[] = [];
  if (options.explicitPath) candidates.push(options.explicitPath);
  if (options.platform === 'darwin') {
    candidates.push('/usr/local/bin/subminer', '/opt/homebrew/bin/subminer');
    candidates.push(path.join(options.homeDir, '.local/bin/subminer'));
  } else if (options.platform === 'linux') {
    candidates.push(path.join(options.homeDir, '.local/bin/subminer'));
    candidates.push('/usr/local/bin/subminer', '/usr/bin/subminer');
  }
  return [...new Set(candidates)];
}

export async function updateLauncherFromRelease(options: {
  release: GitHubRelease | null;
  sha256Sums: Map<string, string>;
  launcherPath?: string;
  platform?: NodeJS.Platform;
  homeDir?: string;
  downloadAsset: (url: string) => Promise<Buffer>;
  exists?: (targetPath: string) => boolean;
}): Promise<LauncherUpdateResult> {
  if (!options.release) return { status: 'missing-asset', message: 'No release found.' };
  const asset = findReleaseAsset(options.release, 'subminer');
  if (!asset) return { status: 'missing-asset', message: 'Release has no subminer asset.' };
  const expectedSha256 = options.sha256Sums.get('subminer');
  if (!expectedSha256) {
    return { status: 'missing-asset', message: 'SHA256SUMS.txt has no subminer entry.' };
  }

  const exists = options.exists ?? fs.existsSync;
  const candidates = detectLauncherCandidates({
    platform: options.platform ?? process.platform,
    homeDir: options.homeDir ?? os.homedir(),
    explicitPath: options.launcherPath,
  });
  const targetPath = candidates.find((candidate) => exists(candidate));
  if (!targetPath) return { status: 'not-found', message: 'No installed launcher detected.' };

  return await updateLauncherAtPath({
    launcherPath: targetPath,
    assetUrl: asset.browser_download_url,
    expectedSha256,
    download: () => options.downloadAsset(asset.browser_download_url),
  });
}
