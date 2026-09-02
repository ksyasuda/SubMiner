import { execFile } from 'node:child_process';
import path from 'node:path';
import process from 'node:process';
import { resolveSubtitleSourcePath } from './subtitle-prefetch-source';

const DEFAULT_MOUNT_CACHE_TTL_MS = 5_000;
const NETWORK_FILESYSTEM_TYPES = new Set([
  '9p',
  'afpfs',
  'cifs',
  'davfs',
  'davfs2',
  'fuse.sshfs',
  'nfs',
  'nfs4',
  'smbfs',
  'sshfs',
  'webdav',
]);

function isRemoteUrl(value: string): boolean {
  try {
    const url = new URL(value);
    return url.protocol === 'http:' || url.protocol === 'https:';
  } catch {
    return false;
  }
}

function decodeMountPath(value: string): string {
  return value.replace(/\\([0-7]{3})/g, (_match, digits: string) =>
    String.fromCharCode(Number.parseInt(digits, 8)),
  );
}

function parseNetworkMountPaths(output: string): string[] {
  const networkMountPaths: string[] = [];
  for (const line of output.split('\n')) {
    const optionsStart = line.lastIndexOf(' (');
    if (optionsStart < 0) continue;

    let mountDescription = line.slice(0, optionsStart);
    const options = line.slice(optionsStart + 2, line.indexOf(')', optionsStart));
    const linuxTypeSeparator = mountDescription.lastIndexOf(' type ');
    const filesystemType = (
      linuxTypeSeparator >= 0
        ? mountDescription.slice(linuxTypeSeparator + ' type '.length)
        : (options.split(',').at(0) ?? '')
    )
      .trim()
      .toLowerCase();
    if (!NETWORK_FILESYSTEM_TYPES.has(filesystemType)) continue;

    if (linuxTypeSeparator >= 0) {
      mountDescription = mountDescription.slice(0, linuxTypeSeparator);
    }
    const mountSeparator = mountDescription.indexOf(' on ');
    if (mountSeparator < 0) continue;
    networkMountPaths.push(
      path.posix.normalize(decodeMountPath(mountDescription.slice(mountSeparator + 4).trim())),
    );
  }
  return networkMountPaths;
}

function readMountOutput(platform: NodeJS.Platform): Promise<string> {
  if (platform === 'win32') return Promise.resolve('');
  const command = platform === 'darwin' ? '/sbin/mount' : 'mount';
  return new Promise((resolve, reject) => {
    execFile(
      command,
      [],
      { encoding: 'utf8', timeout: 1_000, maxBuffer: 1024 * 1024 },
      (error, stdout) => {
        if (error) {
          reject(error);
          return;
        }
        resolve(stdout);
      },
    );
  });
}

function isPathWithinMount(filePath: string, mountPath: string): boolean {
  const relativePath = path.posix.relative(mountPath, filePath);
  return (
    relativePath === '' ||
    (relativePath !== '..' &&
      !relativePath.startsWith(`..${path.posix.sep}`) &&
      !path.posix.isAbsolute(relativePath))
  );
}

export type RemoteMediaPathDetector = (mediaPath: string) => Promise<boolean>;

export function createRemoteMediaPathDetector(
  deps: {
    platform?: NodeJS.Platform;
    readMountOutput?: () => Promise<string>;
    now?: () => number;
    mountCacheTtlMs?: number;
  } = {},
): RemoteMediaPathDetector {
  const platform = deps.platform ?? process.platform;
  const getMountOutput = deps.readMountOutput ?? (() => readMountOutput(platform));
  const now = deps.now ?? Date.now;
  const mountCacheTtlMs = deps.mountCacheTtlMs ?? DEFAULT_MOUNT_CACHE_TTL_MS;
  let mountCache: { expiresAt: number; networkMountPaths: Promise<readonly string[]> } | undefined;

  const getNetworkMountPaths = (): Promise<readonly string[]> => {
    const currentTime = now();
    if (mountCache && currentTime < mountCache.expiresAt) {
      return mountCache.networkMountPaths;
    }

    const networkMountPaths = getMountOutput()
      .then(parseNetworkMountPaths)
      .catch(() => []);
    mountCache = {
      expiresAt: currentTime + mountCacheTtlMs,
      networkMountPaths,
    };
    return networkMountPaths;
  };

  return async (mediaPath): Promise<boolean> => {
    const source = mediaPath.trim();
    if (!source) return false;
    if (isRemoteUrl(source)) return true;

    const filePath = resolveSubtitleSourcePath(source);
    if (platform === 'win32') {
      return filePath.startsWith('\\\\');
    }
    if (!path.posix.isAbsolute(filePath)) return false;

    const networkMountPaths = await getNetworkMountPaths();
    const normalizedPath = path.posix.normalize(filePath);
    return networkMountPaths.some((mountPath) => isPathWithinMount(normalizedPath, mountPath));
  };
}
