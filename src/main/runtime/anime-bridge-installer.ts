import { spawn } from 'node:child_process';
import { chmod, mkdir, mkdtemp, rename, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import {
  bundleReleaseUrl,
  bundleVersionFromJar,
  compareBundleVersions,
  findBundleBinaries,
  readBundleMarker,
  resolveBundleAssetName,
  selectBundleAsset,
  systemBundleDirs,
  writeBundleMarker,
  type BundleAsset,
  type BundleBinaries,
} from '../../anime-bridge/sidecar-bundle';
import type { AnimeBrowserBridgeInstall } from '../../types/anime-browser';

/**
 * Locates the M-Extension-Server bundle that runs Aniyomi extension APKs, or
 * downloads and unpacks it. The bundle ships its own JRE, so no system Java is
 * needed.
 *
 * Resolution order: an explicit `anime.bridgeDir`, then a package-manager
 * install (the AUR `mangatan-extension-server` package on Arch), then the copy
 * SubMiner manages under userData. Only the managed copy is ever downloaded or
 * updated; the others belong to whoever put them there. Downloads take the
 * newest upstream release that ships a bundle for this platform.
 */

export type InstallStage = 'locating' | 'downloading' | 'extracting';

/** Neither call has a default deadline, so a hung network would stall install. */
const RELEASES_TIMEOUT_MS = 30_000;
const DOWNLOAD_TIMEOUT_MS = 300_000;
/**
 * Hard ceiling on the buffered bundle. The real asset is a JRE plus a jar, well
 * under this; the cap only stops a wrong or hostile URL from filling memory.
 */
const MAX_BUNDLE_BYTES = 512 * 1024 * 1024;

function oversizedBundle(): Error {
  return new Error(
    `The anime bridge bundle exceeded ${Math.round(MAX_BUNDLE_BYTES / (1024 * 1024))} MB; ` +
      'the download was stopped.',
  );
}

export interface InstallProgress {
  stage: InstallStage;
  /** 0-1 during download, otherwise null. */
  progress: number | null;
}

/** The binaries to launch plus where they came from, for the browser to show. */
export type BridgeInstall = BundleBinaries & AnimeBrowserBridgeInstall;

export interface BridgeReleaseOptions {
  platform?: string;
  arch?: string;
  fetchImpl?: typeof fetch;
}

export interface EnsureBridgeOptions extends BridgeReleaseOptions {
  /** Directory the managed bundle is unpacked into, e.g. `<userData>/anime-bridge`. */
  installDir: string;
  /** `anime.bridgeDir`: a bundle the user pointed at. Must hold a usable bundle. */
  configuredDir?: string;
  /** Package-manager locations to check before downloading. Defaults per platform. */
  systemDirs?: string[];
  /** Replaces the unzip/tar extraction. Tests only. */
  extractImpl?: (zipPath: string, targetDir: string) => Promise<void>;
  /** Replaces atomic directory moves. Tests only. */
  renameImpl?: typeof rename;
  onProgress?: (progress: InstallProgress) => void;
}

/** Extract a zip without adding a dependency, mirroring scripts/build-yomitan.mjs. */
async function extractZip(zipPath: string, targetDir: string): Promise<void> {
  const attempts: Array<[string, string[]]> = [
    ['unzip', ['-qo', zipPath, '-d', targetDir]],
    // bsdtar ships with macOS and Windows 10+ and reads zip archives.
    ['tar', ['-xf', zipPath, '-C', targetDir]],
  ];

  let lastError = '';
  for (const [command, args] of attempts) {
    const result = await runCommand(command, args);
    if (result.ok) return;
    lastError = result.error;
  }
  throw new Error(`Could not extract the anime bridge bundle. ${lastError}`);
}

function runCommand(
  command: string,
  args: string[],
): Promise<{ ok: true } | { ok: false; error: string }> {
  return new Promise((resolve) => {
    const child = spawn(command, args, { stdio: ['ignore', 'ignore', 'pipe'] });
    let stderr = '';
    child.stderr?.on('data', (chunk: Buffer) => {
      stderr += chunk.toString();
    });
    child.once('error', (error) => resolve({ ok: false, error: `${command}: ${error.message}` }));
    child.once('exit', (code) => {
      if (code === 0) resolve({ ok: true });
      else resolve({ ok: false, error: `${command} exited ${code}. ${stderr.trim()}` });
    });
  });
}

async function downloadWithProgress(
  response: Response,
  onProgress: ((fraction: number) => void) | undefined,
): Promise<Uint8Array> {
  const declared = Number(response.headers.get('content-length') ?? '0');
  const reader = response.body?.getReader();
  if (!reader) {
    const buffer = await response.arrayBuffer();
    if (buffer.byteLength > MAX_BUNDLE_BYTES) throw oversizedBundle();
    return new Uint8Array(buffer);
  }

  const chunks: Uint8Array[] = [];
  let received = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (value) {
      // Checked against the bytes actually read, not content-length: a lying
      // (or absent) header must not let the bundle allocate without bound.
      if (received + value.length > MAX_BUNDLE_BYTES) {
        await reader.cancel().catch(() => {});
        throw oversizedBundle();
      }
      chunks.push(value);
      received += value.length;
      if (declared > 0) onProgress?.(Math.min(1, received / declared));
    }
  }

  const merged = new Uint8Array(received);
  let offset = 0;
  for (const chunk of chunks) {
    merged.set(chunk, offset);
    offset += chunk.length;
  }
  return merged;
}

/** The newest upstream release that ships this platform's bundle. */
async function locateLatestBundle(options: BridgeReleaseOptions): Promise<BundleAsset> {
  const platform = options.platform ?? process.platform;
  const arch = options.arch ?? process.arch;
  const fetchImpl = options.fetchImpl ?? fetch;

  const assetName = resolveBundleAssetName(platform, arch);
  if (assetName === null) {
    throw new Error(
      `No anime bridge build is published for ${platform}/${arch}. ` +
        'Supported: macOS (arm64, x64), Linux (x64), Windows (x64).',
    );
  }

  const releasesResponse = await fetchImpl(bundleReleaseUrl(), {
    headers: { Accept: 'application/vnd.github+json' },
    signal: AbortSignal.timeout(RELEASES_TIMEOUT_MS),
  });
  if (!releasesResponse.ok) {
    throw new Error(`Could not list anime bridge releases (${releasesResponse.status}).`);
  }
  const asset = selectBundleAsset(await releasesResponse.json(), assetName);
  if (asset === null) {
    throw new Error(`No anime bridge release ships ${assetName} at a supported version.`);
  }
  return asset;
}

/**
 * The newest release a managed install could move to, or null when it is
 * current or is not SubMiner's to update. An install whose version cannot be
 * read is offered the newest release: re-downloading is the way back to a
 * known state. Network errors propagate; the caller decides how loudly.
 */
export async function findBridgeUpdate(
  install: Pick<AnimeBrowserBridgeInstall, 'origin' | 'version'>,
  options: BridgeReleaseOptions = {},
): Promise<string | null> {
  if (install.origin !== 'managed') return null;
  const latest = await locateLatestBundle(options);
  if (install.version === null) return latest.tagName;
  return compareBundleVersions(latest.tagName, install.version) > 0 ? latest.tagName : null;
}

async function describeInstall(
  binaries: BundleBinaries,
  dir: string,
  origin: AnimeBrowserBridgeInstall['origin'],
): Promise<BridgeInstall> {
  const version =
    (origin === 'managed' ? await readBundleMarker(dir) : null) ??
    bundleVersionFromJar(binaries.jarPath);
  // The update check needs the network and runs once the bridge is up.
  return { ...binaries, dir, origin, version, updateAvailable: null };
}

/**
 * Download the newest release into `targetDir` and unpack it. The directory
 * is created if needed and is not cleared first.
 */
async function downloadLatestBundle(
  targetDir: string,
  options: EnsureBridgeOptions,
): Promise<{ binaries: BundleBinaries; tagName: string }> {
  options.onProgress?.({ stage: 'locating', progress: null });
  const asset = await locateLatestBundle(options);
  const fetchImpl = options.fetchImpl ?? fetch;

  options.onProgress?.({ stage: 'downloading', progress: 0 });
  const downloadResponse = await fetchImpl(asset.downloadUrl, {
    signal: AbortSignal.timeout(DOWNLOAD_TIMEOUT_MS),
  });
  if (!downloadResponse.ok) {
    throw new Error(`Downloading the anime bridge failed (${downloadResponse.status}).`);
  }
  const bytes = await downloadWithProgress(downloadResponse, (fraction) =>
    options.onProgress?.({ stage: 'downloading', progress: fraction }),
  );

  options.onProgress?.({ stage: 'extracting', progress: null });
  await mkdir(targetDir, { recursive: true });
  const zipPath = path.join(targetDir, asset.assetName);
  await writeFile(zipPath, bytes);
  try {
    await (options.extractImpl ?? extractZip)(zipPath, targetDir);
  } finally {
    await rm(zipPath, { force: true });
  }

  const binaries = await findBundleBinaries(targetDir);
  if (!binaries) {
    throw new Error('The anime bridge bundle unpacked without a java runtime or server jar.');
  }
  // Some extractors drop the executable bit; restore it rather than failing at spawn.
  if ((options.platform ?? process.platform) !== 'win32') {
    await chmod(binaries.javaPath, 0o755).catch(() => undefined);
  }
  await writeBundleMarker(targetDir, asset.tagName);
  return { binaries, tagName: asset.tagName };
}

/**
 * Return the bridge's java + jar paths, downloading the newest release into
 * the managed directory only when no other install is usable.
 */
export async function ensureBridgeBinaries(options: EnsureBridgeOptions): Promise<BridgeInstall> {
  const configuredDir = options.configuredDir?.trim();
  if (configuredDir) {
    const configured = await findBundleBinaries(configuredDir);
    if (!configured) {
      throw new Error(
        `anime.bridgeDir (${configuredDir}) holds no java runtime or MExtensionServer jar.`,
      );
    }
    return describeInstall(configured, configuredDir, 'system');
  }

  for (const dir of options.systemDirs ?? systemBundleDirs(options.platform ?? process.platform)) {
    const system = await findBundleBinaries(dir);
    if (system) return describeInstall(system, dir, 'system');
  }

  const existing = await findBundleBinaries(options.installDir);
  if (existing) return describeInstall(existing, options.installDir, 'managed');

  const downloaded = await downloadLatestBundle(options.installDir, options);
  return describeInstall(downloaded.binaries, options.installDir, 'managed');
}

export interface StagedBridgeUpdate {
  version: string;
  /**
   * Replace the managed install with the staged one. Call only once the old
   * bridge has stopped: on Windows its open files cannot be removed, and on
   * every platform a JVM still running out of the old tree would outlive it.
   */
  commit: () => Promise<BridgeInstall>;
}

/**
 * Download the newest release next to the managed install without touching
 * it, so a failed download leaves the running bridge as it was.
 */
export async function stageBridgeUpdate(options: EnsureBridgeOptions): Promise<StagedBridgeUpdate> {
  const stagingDir = `${options.installDir}.next`;
  await rm(stagingDir, { recursive: true, force: true });
  const downloaded = await downloadLatestBundle(stagingDir, options);
  return {
    version: downloaded.tagName,
    commit: async () => {
      const backupRoot = await mkdtemp(`${options.installDir}.backup-`);
      const backupDir = path.join(backupRoot, 'previous');
      const renameImpl = options.renameImpl ?? rename;
      let backupHoldsInstall = false;

      try {
        await renameImpl(options.installDir, backupDir);
        backupHoldsInstall = true;
        try {
          await renameImpl(stagingDir, options.installDir);
        } catch (replaceError) {
          try {
            await renameImpl(backupDir, options.installDir);
            backupHoldsInstall = false;
          } catch (restoreError) {
            throw new AggregateError(
              [replaceError, restoreError],
              `Failed to activate the staged anime bridge and restore the previous install. ` +
                `The previous install remains at ${backupDir}.`,
            );
          }
          throw replaceError;
        }

        const binaries = await findBundleBinaries(options.installDir);
        if (!binaries) {
          const validationError = new Error(
            'The updated anime bridge is missing its java runtime or jar.',
          );
          try {
            await renameImpl(options.installDir, stagingDir);
            await renameImpl(backupDir, options.installDir);
            backupHoldsInstall = false;
          } catch (restoreError) {
            throw new AggregateError(
              [validationError, restoreError],
              `The updated anime bridge failed validation and the previous install could not be ` +
                `restored. The previous install remains at ${backupDir}.`,
            );
          }
          throw validationError;
        }
        await rm(backupRoot, { recursive: true, force: true });
        backupHoldsInstall = false;
        return describeInstall(binaries, options.installDir, 'managed');
      } finally {
        if (!backupHoldsInstall) {
          await rm(backupRoot, { recursive: true, force: true });
        }
      }
    },
  };
}
