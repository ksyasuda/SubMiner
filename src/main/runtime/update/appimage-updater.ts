import { createHash } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import type { GitHubRelease } from './release-assets';
import { findReleaseAsset } from './release-assets';

type StatLike = {
  isFile: () => boolean;
  mode?: number;
};

export type AppImageUpdateStatus =
  | 'updated'
  | 'skipped'
  | 'protected'
  | 'hash-mismatch'
  | 'not-found'
  | 'missing-asset';

export interface AppImageUpdateResult {
  status: AppImageUpdateStatus;
  path?: string;
  command?: string;
  message?: string;
}

export interface AppImageUpdateFileSystem {
  stat: (targetPath: string) => Promise<StatLike>;
  access: (targetPath: string) => Promise<void>;
  writeFile: (targetPath: string, data: Buffer) => Promise<void>;
  chmod: (targetPath: string, mode: number) => Promise<void>;
  rename: (fromPath: string, toPath: string) => Promise<void>;
  unlink: (targetPath: string) => Promise<void>;
}

function sha256(data: Buffer): string {
  return createHash('sha256').update(data).digest('hex');
}

function defaultFs(): AppImageUpdateFileSystem {
  return {
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

function shellQuote(value: string): string {
  return `'${value.replace(/'/g, `'\\''`)}'`;
}

export function buildProtectedAppImageUpdateCommand(
  assetUrl: string,
  appImagePath: string,
  expectedSha256: string,
): string {
  const quotedUrl = shellQuote(assetUrl);
  const quotedPath = shellQuote(appImagePath);
  const quotedSha256 = shellQuote(expectedSha256.toLowerCase());
  return [
    'tmp=$(mktemp)',
    'trap \'rm -f "$tmp"\' EXIT',
    `curl -fSL ${quotedUrl} -o "$tmp"`,
    `printf '%s  %s\\n' ${quotedSha256} "$tmp" | sha256sum -c -`,
    `sudo mv "$tmp" ${quotedPath}`,
    `sudo chmod +x ${quotedPath}`,
  ].join(' && ');
}

function selectAppImageAsset(release: GitHubRelease, appImagePath: string) {
  const basename = path.basename(appImagePath);
  return (
    findReleaseAsset(release, basename) ??
    findReleaseAsset(release, 'SubMiner.AppImage') ??
    release.assets.find((asset) => asset.name.endsWith('.AppImage')) ??
    null
  );
}

export async function updateAppImageFromRelease(options: {
  release: GitHubRelease | null;
  sha256Sums: Map<string, string>;
  appImagePath?: string;
  downloadAsset: (url: string) => Promise<Buffer>;
  fs?: AppImageUpdateFileSystem;
}): Promise<AppImageUpdateResult> {
  if (!options.appImagePath) {
    return { status: 'not-found', message: 'No AppImage path detected.' };
  }
  if (!options.release) return { status: 'missing-asset', message: 'No release found.' };

  const asset = selectAppImageAsset(options.release, options.appImagePath);
  if (!asset) return { status: 'missing-asset', message: 'Release has no AppImage asset.' };

  const expectedSha256 = options.sha256Sums.get(asset.name);
  if (!expectedSha256) {
    return { status: 'missing-asset', message: `SHA256SUMS.txt has no ${asset.name} entry.` };
  }

  const fsDeps = options.fs ?? defaultFs();
  let stat: StatLike;
  try {
    stat = await fsDeps.stat(options.appImagePath);
  } catch {
    return { status: 'not-found', path: options.appImagePath };
  }
  if (!stat.isFile()) {
    return { status: 'skipped', path: options.appImagePath, message: 'AppImage is not a file.' };
  }

  try {
    await fsDeps.access(options.appImagePath);
  } catch {
    return {
      status: 'protected',
      path: options.appImagePath,
      command: buildProtectedAppImageUpdateCommand(
        asset.browser_download_url,
        options.appImagePath,
        expectedSha256,
      ),
    };
  }

  const data = await options.downloadAsset(asset.browser_download_url);
  const actualSha256 = sha256(data);
  if (actualSha256 !== expectedSha256.toLowerCase()) {
    return {
      status: 'hash-mismatch',
      path: options.appImagePath,
      message: `Expected ${expectedSha256}, got ${actualSha256}.`,
    };
  }

  const appImagePathApi = options.appImagePath.startsWith('/') ? path.posix : path;
  const tempPath = appImagePathApi.join(
    appImagePathApi.dirname(options.appImagePath),
    `.${appImagePathApi.basename(options.appImagePath)}.update`,
  );
  try {
    await fsDeps.writeFile(tempPath, data);
    await fsDeps.chmod(tempPath, stat.mode ? stat.mode & 0o777 : 0o755);
    await fsDeps.rename(tempPath, options.appImagePath);
    return { status: 'updated', path: options.appImagePath };
  } catch (error) {
    await fsDeps.unlink(tempPath);
    throw error;
  }
}
