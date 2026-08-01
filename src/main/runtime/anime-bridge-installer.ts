import { spawn } from 'node:child_process';
import { chmod, mkdir, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import {
  bundleReleaseUrl,
  findBundleBinaries,
  PINNED_BUNDLE_TAG,
  resolveBundleAssetName,
  selectBundleAsset,
  verifyPinnedBundle,
  type BundleBinaries,
} from '../../anime-bridge/sidecar-bundle';

/**
 * Downloads and unpacks the M-Extension-Server bundle that runs Aniyomi
 * extension APKs. The bundle ships its own JRE, so no system Java is needed.
 */

export type InstallStage = 'locating' | 'downloading' | 'verifying' | 'extracting';

/** Neither call has a default deadline, so a hung network would stall install. */
const RELEASES_TIMEOUT_MS = 30_000;
const DOWNLOAD_TIMEOUT_MS = 300_000;

export interface InstallProgress {
  stage: InstallStage;
  /** 0-1 during download, otherwise null. */
  progress: number | null;
}

export interface EnsureBridgeOptions {
  /** Directory the bundle is unpacked into, e.g. `<userData>/anime-bridge`. */
  installDir: string;
  platform?: string;
  arch?: string;
  fetchImpl?: typeof fetch;
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
  if (!reader) return new Uint8Array(await response.arrayBuffer());

  const chunks: Uint8Array[] = [];
  let received = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    if (value) {
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

/**
 * Return the bundle's java + jar paths, downloading the release first if the
 * install directory does not already hold a usable copy.
 */
export async function ensureBridgeBinaries(options: EnsureBridgeOptions): Promise<BundleBinaries> {
  const existing = await findBundleBinaries(options.installDir);
  if (existing) return existing;

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

  options.onProgress?.({ stage: 'locating', progress: null });
  const releasesResponse = await fetchImpl(bundleReleaseUrl(), {
    headers: { Accept: 'application/vnd.github+json' },
    signal: AbortSignal.timeout(RELEASES_TIMEOUT_MS),
  });
  if (!releasesResponse.ok) {
    throw new Error(
      `Could not read anime bridge release ${PINNED_BUNDLE_TAG} (${releasesResponse.status}).`,
    );
  }
  const asset = selectBundleAsset(await releasesResponse.json(), assetName);
  if (asset === null) {
    throw new Error(`Anime bridge release ${PINNED_BUNDLE_TAG} has no ${assetName}.`);
  }

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

  options.onProgress?.({ stage: 'verifying', progress: null });
  const verification = verifyPinnedBundle(assetName, bytes);
  if (!verification.ok) throw new Error(verification.reason);

  options.onProgress?.({ stage: 'extracting', progress: null });
  await mkdir(options.installDir, { recursive: true });
  const zipPath = path.join(options.installDir, assetName);
  await writeFile(zipPath, bytes);
  try {
    await extractZip(zipPath, options.installDir);
  } finally {
    await rm(zipPath, { force: true });
  }

  const binaries = await findBundleBinaries(options.installDir);
  if (!binaries) {
    throw new Error('The anime bridge bundle unpacked without a java runtime or server jar.');
  }
  // Some extractors drop the executable bit; restore it rather than failing at spawn.
  if (platform !== 'win32') {
    await chmod(binaries.javaPath, 0o755).catch(() => undefined);
  }
  return binaries;
}
