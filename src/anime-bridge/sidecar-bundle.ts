import { createHash } from 'node:crypto';
import { readdir, stat } from 'node:fs/promises';
import path from 'node:path';

/**
 * Locates the M-Extension-Server release bundle for the host platform. Each
 * bundle ships a matching JRE alongside the server JAR, so no system JDK is
 * required.
 */

const BUNDLE_REPO_API = 'https://api.github.com/repos/1Selxo/M-Extension-Server';

/**
 * Fetch the pinned release by tag rather than listing releases: upstream ships
 * several a week, so a paged list would scroll the pinned tag off page one.
 */
export function bundleReleaseUrl(tagName: string = PINNED_BUNDLE_TAG): string {
  return `${BUNDLE_REPO_API}/releases/tags/${encodeURIComponent(tagName)}`;
}

/**
 * The bridge release this integration was verified against.
 *
 * Upstream publishes no checksums for the desktop bundles, so we pin a tag and
 * a hash we computed ourselves rather than tracking "latest". Bumping this
 * means downloading the new asset, verifying it starts and reports the
 * capabilities in `AnimeBridgeClient.isReady`, then updating both fields.
 */
export const PINNED_BUNDLE_TAG = 'v1.0.6.0';

/** SHA-256 of each pinned asset, keyed by release asset name. */
export const PINNED_BUNDLE_SHA256: Readonly<Record<string, string>> = {
  'macOS-arm64-bundle.zip': '5f4fb03abfe88bc46ddf5f4d8221156ee2d66b9cbad7c4bc3ade4baf3a4266e6',
  'linux-x64-bundle.zip': 'c2b869d3905b06a308517fec0b44f70ff76f7212230c60710bba39a7025c3a69',
};

/**
 * Check a downloaded asset against its pin. Assets we have not verified
 * ourselves are rejected rather than trusted, so an unpinned platform fails
 * loudly instead of silently running an unchecked binary.
 */
export function verifyPinnedBundle(
  assetName: string,
  bytes: Uint8Array,
): { ok: true } | { ok: false; reason: string } {
  const expected = PINNED_BUNDLE_SHA256[assetName];
  if (expected === undefined) {
    return { ok: false, reason: `No pinned checksum for ${assetName}; refusing to run it.` };
  }
  const actual = sha256(bytes);
  if (actual !== expected) {
    return {
      ok: false,
      reason: `Checksum mismatch for ${assetName}: expected ${expected}, got ${actual}.`,
    };
  }
  return { ok: true };
}

/** Release asset name for a platform/arch pair, or null when unsupported. */
export function resolveBundleAssetName(platform: string, arch: string): string | null {
  if (platform === 'linux') return arch === 'x64' ? 'linux-x64-bundle.zip' : null;
  if (platform === 'win32') return arch === 'x64' ? 'windows-x64-bundle.zip' : null;
  if (platform === 'darwin') {
    if (arch === 'arm64') return 'macOS-arm64-bundle.zip';
    if (arch === 'x64') return 'macOS-x64-bundle.zip';
  }
  return null;
}

export interface BundleBinaries {
  javaPath: string;
  jarPath: string;
}

async function walk(dir: string, depth: number, onFile: (file: string) => void): Promise<void> {
  if (depth < 0) return;
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) await walk(full, depth - 1, onFile);
    else onFile(full);
  }
}

/**
 * Find the bundled `java` executable and `MExtensionServer-*.jar` inside an
 * extracted bundle. The archive nests them a few levels deep and the exact
 * layout differs per platform, so this searches rather than assuming a path.
 */
export async function findBundleBinaries(rootDir: string): Promise<BundleBinaries | null> {
  const javaCandidates: string[] = [];
  const jarCandidates: string[] = [];

  await walk(rootDir, 6, (file) => {
    const base = path.basename(file);
    if (base === 'java' || base === 'java.exe') javaCandidates.push(file);
    else if (/^MExtensionServer.*\.jar$/.test(base)) jarCandidates.push(file);
  });

  // Prefer the shallowest match so a nested duplicate never shadows the real one.
  const byDepth = (a: string, b: string) => a.split(path.sep).length - b.split(path.sep).length;
  const javaPath = javaCandidates.sort(byDepth)[0];
  const jarPath = jarCandidates.sort(byDepth)[0];
  if (!javaPath || !jarPath) return null;
  return { javaPath, jarPath };
}

/** Verify a downloaded archive against a pinned SHA-256, as Mangatan does. */
export function sha256(bytes: Uint8Array): string {
  return createHash('sha256').update(bytes).digest('hex');
}

export async function isExecutableFile(file: string): Promise<boolean> {
  try {
    const info = await stat(file);
    return info.isFile();
  } catch {
    return false;
  }
}

export interface BundleAsset {
  tagName: string;
  assetName: string;
  downloadUrl: string;
  sizeBytes: number;
}

interface GithubRelease {
  tag_name?: string;
  assets?: Array<{ name?: string; browser_download_url?: string; size?: number }>;
}

/**
 * Pick the asset for this platform from the pinned release. Selecting "newest"
 * instead would download a release whose checksum we never computed, so every
 * upstream publish would break the install with a mismatch.
 */
export function selectBundleAsset(
  releases: unknown,
  assetName: string,
  tagName: string = PINNED_BUNDLE_TAG,
): BundleAsset | null {
  // Accepts either a single release (the by-tag endpoint) or a list.
  const candidates = Array.isArray(releases)
    ? releases
    : releases && typeof releases === 'object'
      ? [releases]
      : [];
  for (const release of candidates as GithubRelease[]) {
    if (release.tag_name !== tagName) continue;
    const asset = release.assets?.find((candidate) => candidate.name === assetName);
    if (asset?.browser_download_url && release.tag_name) {
      return {
        tagName: release.tag_name,
        assetName,
        downloadUrl: asset.browser_download_url,
        sizeBytes: asset.size ?? 0,
      };
    }
  }
  return null;
}
